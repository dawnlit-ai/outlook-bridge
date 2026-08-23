// Scratch-file and destination handling. Worth pinning: the shared-directory
// behavior these replace is what let two emails carrying "invoice.pdf"
// overwrite each other.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const {
    configure,
    getConfig,
    tempFile,
    makeAttachmentDir,
    resolveDestDir,
} = require('../dist/runtime.js');

test('defaults are set before anything calls configure', () => {
    const c = getConfig();
    assert.equal(typeof c.timeoutMs, 'number');
    assert.ok(c.timeoutMs > 0);
    assert.equal(c.maxBufferBytes, 8 * 1024 * 1024);
    assert.equal(c.tempDir, os.tmpdir());
});

test('configure merges rather than replacing', () => {
    const before = getConfig().maxBufferBytes;
    configure({timeoutMs: 4242});
    assert.equal(getConfig().timeoutMs, 4242);
    assert.equal(getConfig().maxBufferBytes, before, 'untouched option kept its value');
    configure({timeoutMs: 120_000});
});

test('configure accepts 0 as "no timeout" but clamps a negative', () => {
    configure({timeoutMs: 0});
    assert.equal(getConfig().timeoutMs, 0);
    configure({timeoutMs: -5});
    assert.equal(getConfig().timeoutMs, 0);
    configure({timeoutMs: 120_000});
});

test('tempFile names are unique within the same millisecond', () => {
    // Two concurrent calls used to be able to pick the same path and overwrite
    // each other's script mid-run, since the name was Date.now() alone.
    const names = new Set(Array.from({length: 500}, () => tempFile('probe', 'ps1')));
    assert.equal(names.size, 500);
});

test('tempFile honours a configured tempDir and carries the package prefix', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-test-'));
    try {
        configure({tempDir: dir});
        const f = tempFile('email-body', 'html');
        assert.equal(path.dirname(f), dir);
        assert.match(path.basename(f), /^outlook-bridge-email-body-.*\.html$/);
    } finally {
        configure({tempDir: os.tmpdir()});
        fs.rmSync(dir, {recursive: true, force: true});
    }
});

test('makeAttachmentDir returns a fresh directory each call', () => {
    const a = makeAttachmentDir();
    const b = makeAttachmentDir();
    try {
        assert.notEqual(a, b);
        assert.ok(fs.statSync(a).isDirectory());
        assert.ok(fs.statSync(b).isDirectory());
    } finally {
        fs.rmSync(a, {recursive: true, force: true});
        fs.rmSync(b, {recursive: true, force: true});
    }
});

test('resolveDestDir creates a named destination, nested parents included', () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-test-'));
    try {
        const want = path.join(base, 'deep', 'nested');
        assert.equal(resolveDestDir(want), want);
        assert.ok(fs.statSync(want).isDirectory());
        // Idempotent: an existing directory is returned, not rejected.
        assert.equal(resolveDestDir(want), want);
    } finally {
        fs.rmSync(base, {recursive: true, force: true});
    }
});

test('resolveDestDir with no argument falls back to a private directory', () => {
    const a = resolveDestDir();
    const b = resolveDestDir(undefined);
    try {
        assert.notEqual(a, b);
    } finally {
        fs.rmSync(a, {recursive: true, force: true});
        fs.rmSync(b, {recursive: true, force: true});
    }
});
