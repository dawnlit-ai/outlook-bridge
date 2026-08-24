// The factory, capability discovery, and per-instance config isolation.
//
// Nothing here touches Outlook: the calls exercised are the ones that fail
// before any script is generated, which is what makes them safe to run on a
// machine with no Outlook at all.
const test = require('node:test');
const assert = require('node:assert/strict');

const {createOutlookBridge, capabilities, supports} = require('../dist/OutlookService.js');
const {configure, getConfig, withConfig} = require('../dist/runtime.js');
const windows = require('../dist/windows/index.js');
const mac = require('../dist/mac/index.js');

test('both platform capability maps describe exactly the same operations', () => {
    // The compile-time Record<BridgeCapability, boolean> already forces this, but
    // it is the property everything else here depends on, so pin it at runtime
    // too — a hand-edited dist would otherwise slip through.
    assert.deepEqual(
        Object.keys(windows.capabilities).sort(),
        Object.keys(mac.capabilities).sort(),
    );
});

test('Windows is the reference implementation — everything supported', () => {
    assert.ok(Object.values(windows.capabilities).every(Boolean));
});

test('macOS now answers the whole contract too', () => {
    // This map used to announce a half-ported surface. The gaps are closed, so
    // what it pins now is that they stay closed: a capability flipped back to
    // false is a regression, not a documentation change.
    for (const [name, supported] of Object.entries(mac.capabilities)) {
        assert.equal(supported, true, `${name} should be implemented on macOS`);
    }
});

test('every operation in the contract is a function on both platforms', () => {
    for (const name of Object.keys(windows.capabilities)) {
        assert.equal(typeof windows[name], 'function', `windows.${name}`);
        assert.equal(typeof mac[name], 'function', `mac.${name}`);
    }
});

test('capabilities() and supports() agree with each other', () => {
    const map = capabilities();
    for (const [name, value] of Object.entries(map)) {
        assert.equal(supports(name), value, name);
    }
});

test('capabilities() reflects the real platform, not the fallback module', () => {
    // Off Windows and macOS the dispatcher still holds the PowerShell service
    // (whose calls reject) — reporting its all-true map there would be a lie.
    const expected = process.platform === 'win32'
        ? windows.capabilities
        : process.platform === 'darwin' ? mac.capabilities : null;
    if (expected) {
        assert.deepEqual(capabilities(), expected);
    } else {
        assert.ok(Object.values(capabilities()).every(v => v === false));
    }
});

test('a bridge exposes every operation as a function', () => {
    const bridge = createOutlookBridge();
    for (const name of Object.keys(capabilities())) {
        assert.equal(typeof bridge[name], 'function', name);
    }
});

test('createOutlookBridge takes settings without touching the global', () => {
    const before = getConfig().timeoutMs;
    const bridge = createOutlookBridge({timeoutMs: 4321});
    assert.equal(bridge.options.timeoutMs, 4321);
    assert.equal(getConfig().timeoutMs, before, 'global config untouched');
});

test('withOptions derives a new bridge and leaves the original alone', () => {
    const base = createOutlookBridge({timeoutMs: 1000, maxBufferBytes: 2048});
    const derived = base.withOptions({timeoutMs: 5000});
    assert.equal(derived.options.timeoutMs, 5000);
    assert.equal(derived.options.maxBufferBytes, 2048, 'unnamed options carry over');
    assert.equal(base.options.timeoutMs, 1000, 'original untouched');
});

test('two bridges hold independent settings at the same time', () => {
    const a = createOutlookBridge({timeoutMs: 1111});
    const b = createOutlookBridge({timeoutMs: 2222});
    assert.equal(a.options.timeoutMs, 1111);
    assert.equal(b.options.timeoutMs, 2222);
});

test('a scoped config is what getConfig sees inside it, and only inside', () => {
    const outer = getConfig().timeoutMs;
    const scopedValue = withConfig({...getConfig(), timeoutMs: 999}, () => getConfig().timeoutMs);
    assert.equal(scopedValue, 999);
    assert.equal(getConfig().timeoutMs, outer);
});

test('a scoped config survives awaits inside the scope', async () => {
    // This is the property the whole AsyncLocalStorage approach rests on: the
    // services read getConfig() after several awaits, long after run() returned.
    const seen = await withConfig({...getConfig(), timeoutMs: 777}, async () => {
        await new Promise(resolve => setImmediate(resolve));
        await new Promise(resolve => setTimeout(resolve, 1));
        return getConfig().timeoutMs;
    });
    assert.equal(seen, 777);
});

test('concurrent scopes do not leak into each other', async () => {
    const run = (ms) => withConfig({...getConfig(), timeoutMs: ms}, async () => {
        await new Promise(resolve => setTimeout(resolve, 5));
        return getConfig().timeoutMs;
    });
    const [a, b, c] = await Promise.all([run(10), run(20), run(30)]);
    assert.deepEqual([a, b, c], [10, 20, 30]);
});

test('configure() still drives the process-wide default', () => {
    const before = getConfig().timeoutMs;
    configure({timeoutMs: 31_000});
    assert.equal(getConfig().timeoutMs, 31_000);
    configure({timeoutMs: before});
});

test('an unsupported operation rejects rather than throwing synchronously', async () => {
    // Every failure has to arrive on one channel; a caller should never need both
    // a try/catch and a .catch() around the same call.
    const bridge = createOutlookBridge();
    const unsupported = Object.entries(capabilities()).find(([, ok]) => !ok);
    if (!unsupported) return; // Windows and macOS: nothing to assert.
    const [name] = unsupported;
    const result = bridge[name]('someone@example.com', 'x');
    assert.ok(result instanceof Promise, `${name} returns a promise`);
    await assert.rejects(result, (error) => {
        assert.ok(error instanceof Error);
        return true;
    });
});

test('macOS rejects a Windows EntryID before it reaches Outlook', async (t) => {
    if (process.platform !== 'darwin') return t.skip('macOS only');
    // The two id spaces are not interchangeable, and the reads that would fail
    // on a stale id say so up front rather than as a confusing "not found".
    const bridge = createOutlookBridge();
    const windowsEntryId = '00000000AABBCCDD1122334455667788';
    for (const call of [
        () => bridge.readEmailBody(windowsEntryId),
        () => bridge.openOutlookEmail(windowsEntryId),
        () => bridge.replyOutlookEmail({
            emailAccount: 'a@b.com',
            entryId: windowsEntryId,
            htmlBody: '<p>x</p>',
        }),
    ]) {
        await assert.rejects(call(), (error) => {
            assert.equal(error.code, 'INVALID_REQUEST');
            assert.match(error.message, /Outlook for Mac message id/);
            return true;
        });
    }
});

test('per-id operations report a bad id rather than discarding the batch', async (t) => {
    if (process.platform !== 'darwin') return t.skip('macOS only');
    // A batch is not all-or-nothing: an unusable id belongs in `failed` beside a
    // genuine lookup miss, not as an exception that throws away the good ids.
    const bridge = createOutlookBridge();
    const result = await bridge.moveOutlookEmails('a@b.com', ['not-a-mac-id'], 'Archive');
    assert.equal(result.moved, 0);
    assert.equal(result.failed.length, 1);
    assert.equal(result.failed[0].entryId, 'not-a-mac-id');
});

test('editEmailTemplate is part of the contract on both platforms', async (t) => {
    const bridge = createOutlookBridge();
    assert.equal(typeof bridge.editEmailTemplate, 'function');
    assert.equal(capabilities().editEmailTemplate, process.platform === 'win32' || process.platform === 'darwin');
    if (process.platform === 'win32' || process.platform === 'darwin') {
        return t.skip('would open a real compose window');
    }
    await assert.rejects(bridge.editEmailTemplate('label', '<p>x</p>'), (error) => {
        assert.equal(error.code, 'UNSUPPORTED_PLATFORM');
        return true;
    });
});
