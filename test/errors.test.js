// The error taxonomy, and the classifier that decides which one a failed run is.
// These codes are the package's API — the messages beside them are not — so the
// point of pinning them here is that a reworded message can't silently change
// what a consumer's `switch (err.code)` sees.
const test = require('node:test');
const assert = require('node:assert/strict');

const {
    OutlookError,
    UnsupportedPlatformError,
    NotImplementedError,
    AccountNotFoundError,
    NotFoundError,
    InvalidRequestError,
    ScriptError,
    TimeoutError,
    AbortedError,
    classifyRunFailure,
} = require('../dist/errors.js');

test('every error is an OutlookError carrying a stable code', () => {
    const cases = [
        [new UnsupportedPlatformError('linux'), 'UNSUPPORTED_PLATFORM'],
        [new NotImplementedError('replyOutlookEmail', 'macOS'), 'NOT_IMPLEMENTED'],
        [new AccountNotFoundError('a@b.com'), 'ACCOUNT_NOT_FOUND'],
        [new NotFoundError('folder', 'nope'), 'NOT_FOUND'],
        [new InvalidRequestError('nope'), 'INVALID_REQUEST'],
        [new ScriptError({runner: 'powershell', script: 's', stderr: 'boom', durationMs: 1}), 'SCRIPT_FAILED'],
        [new TimeoutError({runner: 'osascript', timeoutMs: 5, script: 's'}), 'TIMEOUT'],
        [new AbortedError('powershell'), 'ABORTED'],
    ];
    for (const [error, code] of cases) {
        assert.ok(error instanceof OutlookError, `${error.name} is an OutlookError`);
        assert.ok(error instanceof Error, `${error.name} is an Error`);
        assert.equal(error.code, code);
        // Compiled to ES2020, a subclass loses its prototype chain without the
        // explicit setPrototypeOf — which would make every instanceof above pass
        // and every specific one below fail.
        assert.equal(error.name, error.constructor.name);
    }
});

test('subclasses survive instanceof against their own type', () => {
    assert.ok(new AccountNotFoundError('a@b.com') instanceof AccountNotFoundError);
    assert.ok(!(new NotFoundError('folder', 'x') instanceof AccountNotFoundError));
});

test('errors carry the detail a caller would otherwise have to parse out', () => {
    assert.equal(new AccountNotFoundError('ops@x.com').account, 'ops@x.com');
    assert.equal(new NotImplementedError('sendAllDrafts', 'macOS').operation, 'sendAllDrafts');
    assert.equal(new NotFoundError('template', 'x').kind, 'template');
    assert.equal(new UnsupportedPlatformError('linux').platform, 'linux');

    const script = new ScriptError({runner: 'powershell', script: '$x = 1', stderr: 'boom', durationMs: 12});
    assert.equal(script.script, '$x = 1');
    assert.equal(script.stderr, 'boom');
    assert.equal(script.durationMs, 12);
    assert.equal(script.message, 'boom');
});

test('classifyRunFailure recognises our own account-not-found sentence', () => {
    const error = classifyRunFailure({
        runner: 'powershell',
        script: 's',
        stderr: "Account 'ops@x.com' not found",
        durationMs: 5,
    });
    assert.equal(error.code, 'ACCOUNT_NOT_FOUND');
    assert.equal(error.account, 'ops@x.com');
});

test('classifyRunFailure falls back to SCRIPT_FAILED for anything else', () => {
    const error = classifyRunFailure({
        runner: 'osascript',
        script: 'tell app',
        stderr: 'Microsoft Outlook got an error: -1728',
        durationMs: 5,
    });
    assert.equal(error.code, 'SCRIPT_FAILED');
    assert.equal(error.runner, 'osascript');
    assert.equal(error.script, 'tell app');
});

test('an aborted run is ABORTED even though it was killed like a timeout', () => {
    // The whole reason the classifier checks the signal first: abort and timeout
    // both arrive as a SIGTERM kill, so `killed` alone cannot tell them apart.
    const controller = new AbortController();
    controller.abort();
    const error = classifyRunFailure({
        runner: 'powershell',
        script: 's',
        stderr: '',
        durationMs: 5,
        nodeError: Object.assign(new Error('aborted'), {killed: true}),
        timeoutMs: 1000,
        signal: controller.signal,
    });
    assert.equal(error.code, 'ABORTED');
});

test('a killed run with no abort is a TIMEOUT carrying its budget', () => {
    const error = classifyRunFailure({
        runner: 'powershell',
        script: 's',
        stderr: '',
        durationMs: 1000,
        nodeError: Object.assign(new Error('killed'), {killed: true}),
        timeoutMs: 1000,
    });
    assert.equal(error.code, 'TIMEOUT');
    assert.equal(error.timeoutMs, 1000);
});

test("a kill outranks stderr, which may hold a partial message", () => {
    // A process killed mid-write can leave misleading output behind; the kill is
    // the more reliable signal, so it must win.
    const error = classifyRunFailure({
        runner: 'powershell',
        script: 's',
        stderr: "Account 'ops@x.com' not found",
        durationMs: 1000,
        nodeError: Object.assign(new Error('killed'), {killed: true}),
        timeoutMs: 1000,
    });
    assert.equal(error.code, 'TIMEOUT');
});

test('an AbortError from node is recognised without the signal', () => {
    const error = classifyRunFailure({
        runner: 'osascript',
        script: 's',
        stderr: '',
        durationMs: 5,
        nodeError: Object.assign(new Error('The operation was aborted'), {name: 'AbortError'}),
    });
    assert.equal(error.code, 'ABORTED');
});
