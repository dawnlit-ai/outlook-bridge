// Bridge instances, capability discovery, and per-instance settings.
//
// Nothing here touches Outlook: instances are built over a recording fake
// backend, or over none at all for the unsupported-platform case.
const test = require('node:test');
const assert = require('node:assert/strict');

const index = require('../dist/index.js');
const {createBridge, OPERATIONS} = require('../dist/bridge.js');
const {configure, getConfig, getGlobalConfig, withConfig} = require('../dist/runtime.js');
const {fakeBackend} = require('./support.js');

const PUBLIC_OPERATIONS = [
    'getOutlookAccounts', 'sendOutlookEmail', 'replyOutlookEmail', 'readInboxEmails', 'searchInboxByFilter',
    'readSelectedEmail', 'readEmailBody', 'openOutlookEmail', 'listInboxFolders', 'moveOutlookEmails',
    'listOutlookDrafts', 'sendDrafts', 'sendAllDrafts', 'deleteOutlookDrafts', 'deleteOutlookEmails',
    'purgeDeletedItems', 'saveEmailAttachment', 'saveEmailAttachments', 'cleanUndeliverableEmails',
    'collectBouncedRecipients', 'readSentRecipientGroups', 'listOutlookSignatures', 'readOutlookSignatureHtml',
    'readTemplateEmails', 'saveTemplateEmail', 'editEmailTemplate',
];

test('the package exports every operation as a function', () => {
    assert.deepEqual([...OPERATIONS].sort(), [...PUBLIC_OPERATIONS].sort());
    for (const name of PUBLIC_OPERATIONS) assert.equal(typeof index[name], 'function', name);
});

test('both platform backends implement the same methods', () => {
    const {windowsBackend} = require('../dist/windows/index.js');
    const {macBackend} = require('../dist/mac/index.js');
    assert.deepEqual(Object.keys(windowsBackend).sort(), Object.keys(macBackend).sort());
    for (const fn of [...Object.values(windowsBackend), ...Object.values(macBackend)]) {
        assert.equal(typeof fn, 'function');
    }
});

test('capabilities() reflects the real platform', () => {
    const supported = process.platform === 'win32' || process.platform === 'darwin';
    const map = index.capabilities();
    assert.deepEqual(Object.keys(map).sort(), [...PUBLIC_OPERATIONS].sort());
    for (const [name, value] of Object.entries(map)) {
        assert.equal(value, supported, name);
        assert.equal(index.supports(name), supported, name);
    }
});

test('a bridge over no backend reports nothing supported and rejects every call', async () => {
    const bridge = createBridge(null, getGlobalConfig());
    assert.ok(Object.values(bridge.capabilities()).every(value => value === false));
    for (const name of PUBLIC_OPERATIONS) {
        const result = bridge[name]('someone@example.com', {});
        assert.ok(result instanceof Promise, `${name} returns a promise`);
        await assert.rejects(result, error => error.code === 'UNSUPPORTED_PLATFORM', name);
    }
});

test('a bad argument rejects rather than throwing synchronously', async () => {
    const {backend} = fakeBackend();
    const bridge = createBridge(backend, getGlobalConfig());
    const result = bridge.readInboxEmails(42);
    assert.ok(result instanceof Promise);
    await assert.rejects(result, error => error.code === 'INVALID_REQUEST');
});

test('createOutlookBridge takes settings without touching the process-wide ones', () => {
    const before = getConfig().timeoutMs;
    const bridge = index.createOutlookBridge({timeoutMs: 4321});
    assert.equal(bridge.options.timeoutMs, 4321);
    assert.equal(getConfig().timeoutMs, before);
});

test('withOptions derives a new bridge and leaves the original alone', () => {
    const base = index.createOutlookBridge({timeoutMs: 1000, maxBufferBytes: 2048});
    const derived = base.withOptions({timeoutMs: 5000});
    assert.equal(derived.options.timeoutMs, 5000);
    assert.equal(derived.options.maxBufferBytes, 2048, 'unnamed options carry over');
    assert.equal(base.options.timeoutMs, 1000, 'original untouched');
    assert.equal(base.withOptions({timeoutMs: null}).options.timeoutMs, undefined, 'null goes back to per-operation budgets');
});

test('a bridge runs each call with its own settings in force', async () => {
    let seen;
    const {backend} = fakeBackend({
        getOutlookAccounts: () => {
            seen = getConfig().timeoutMs;
            return [];
        }
    });
    await createBridge(backend, {...getGlobalConfig(), timeoutMs: 777}).getOutlookAccounts();
    assert.equal(seen, 777);
});

test('a scoped config survives awaits inside the scope, and never leaks between scopes', async () => {
    const run = ms => withConfig({...getConfig(), timeoutMs: ms}, async () => {
        await new Promise(resolve => setTimeout(resolve, 5));
        return getConfig().timeoutMs;
    });
    assert.deepEqual(await Promise.all([run(10), run(20), run(30)]), [10, 20, 30]);
    assert.equal(getConfig().timeoutMs, getGlobalConfig().timeoutMs);
});

test('configure() drives the process-wide settings', () => {
    configure({timeoutMs: 31_000});
    assert.equal(getConfig().timeoutMs, 31_000);
    configure({timeoutMs: null});
    assert.equal(getConfig().timeoutMs, undefined);
});

test('settings of the wrong type are refused as INVALID_REQUEST', () => {
    assert.throws(() => index.createOutlookBridge({timeoutMs: 'soon'}), error => error.code === 'INVALID_REQUEST');
});
