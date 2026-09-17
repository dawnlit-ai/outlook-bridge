// Settings, time budgets and scratch-file handling.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const {
    attachmentDestination,
    configure,
    getConfig,
    RUN_BUDGETS,
    tempFile,
    timeoutFor,
    withConfig,
} = require('../dist/runtime.js');

test('defaults are set before anything calls configure', () => {
    const config = getConfig();
    assert.equal(config.timeoutMs, undefined, 'each operation uses its own budget until one is set');
    assert.equal(config.maxBufferBytes, 8 * 1024 * 1024);
    assert.equal(config.tempDir, os.tmpdir());
});

test('without a timeoutMs, each run gets its budget; with one, every run gets it', () => {
    assert.equal(timeoutFor('quick'), RUN_BUDGETS.quick);
    assert.equal(timeoutFor('scan'), RUN_BUDGETS.scan);
    withConfig({...getConfig(), timeoutMs: 1234}, () => {
        assert.equal(timeoutFor('quick'), 1234);
        assert.equal(timeoutFor('purge'), 1234);
    });
});

test('an interactive run is never timed, whatever timeoutMs says', () => {
    withConfig({...getConfig(), timeoutMs: 1234}, () => assert.equal(timeoutFor('interactive'), 0));
});

test('configure merges, accepts 0 as "no timeout", clamps a negative and resets on null', () => {
    const before = getConfig().maxBufferBytes;
    configure({timeoutMs: 4242});
    assert.equal(getConfig().timeoutMs, 4242);
    assert.equal(getConfig().maxBufferBytes, before, 'an option left out keeps its value');
    configure({timeoutMs: 0});
    assert.equal(getConfig().timeoutMs, 0);
    configure({timeoutMs: -5});
    assert.equal(getConfig().timeoutMs, 0);
    configure({timeoutMs: null});
    assert.equal(getConfig().timeoutMs, undefined);
});

test('tempFile names are unique within the same millisecond', () => {
    const names = new Set(Array.from({length: 500}, () => tempFile('probe', 'ps1')));
    assert.equal(names.size, 500);
});

test('tempFile honours a configured tempDir and carries the package prefix', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-test-'));
    try {
        withConfig({...getConfig(), tempDir: dir}, () => {
            const file = tempFile('email-body', 'html');
            assert.equal(path.dirname(file), dir);
            assert.match(path.basename(file), /^outlook-bridge-email-body-.*\.html$/);
        });
    } finally {
        fs.rmSync(dir, {recursive: true, force: true});
    }
});

test('with no destination, each call gets a fresh private directory', () => {
    const a = attachmentDestination();
    const b = attachmentDestination();
    try {
        assert.notEqual(a.dir, b.dir);
        assert.ok(fs.statSync(a.dir).isDirectory());
    } finally {
        fs.rmSync(a.dir, {recursive: true, force: true});
        fs.rmSync(b.dir, {recursive: true, force: true});
    }
});

test('a named destination is created, and discarded only if this call created it', () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-test-'));
    try {
        const wanted = path.join(base, 'deep', 'nested');
        const created = attachmentDestination(wanted);
        assert.equal(created.dir, wanted);
        assert.ok(fs.statSync(wanted).isDirectory());
        created.discardIfUnused();
        assert.equal(fs.existsSync(wanted), false, 'an empty directory this call made is removed');

        fs.mkdirSync(wanted, {recursive: true});
        attachmentDestination(wanted).discardIfUnused();
        assert.ok(fs.existsSync(wanted), 'a directory that already existed is left alone');
    } finally {
        fs.rmSync(base, {recursive: true, force: true});
    }
});
