// The error taxonomy, and the classifier that decides which error a failed run
// is. The codes are the package's API — the messages beside them are not — so
// the point of pinning them is that a reworded message can't change what a
// caller's `switch (error.code)` sees.
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
    OutputTooLargeError,
    TimeoutError,
    AbortedError,
    classifyRunFailure,
    errorFromTaggedMessage,
    failureTag,
    stripFailureTags,
} = require('../dist/errors.js');

test('every error is an OutlookError carrying a stable code', () => {
    const cases = [
        [new UnsupportedPlatformError('linux'), 'UNSUPPORTED_PLATFORM'],
        [new NotImplementedError('Reading Journal', 'macOS'), 'NOT_IMPLEMENTED'],
        [new AccountNotFoundError('a@b.com'), 'ACCOUNT_NOT_FOUND'],
        [new NotFoundError('folder', 'nope'), 'NOT_FOUND'],
        [new InvalidRequestError('nope'), 'INVALID_REQUEST'],
        [new ScriptError({runner: 'powershell', script: 's', stderr: 'boom', durationMs: 1}), 'SCRIPT_FAILED'],
        [new OutputTooLargeError({runner: 'powershell', maxBufferBytes: 10, script: 's'}), 'OUTPUT_TOO_LARGE'],
        [new TimeoutError({runner: 'osascript', timeoutMs: 5, script: 's'}), 'TIMEOUT'],
        [new AbortedError('powershell'), 'ABORTED'],
    ];
    for (const [error, code] of cases) {
        assert.ok(error instanceof OutlookError, `${error.name} is an OutlookError`);
        assert.ok(error instanceof Error, `${error.name} is an Error`);
        assert.equal(error.code, code);
        assert.equal(error.name, error.constructor.name);
    }
});

test('subclasses survive instanceof against their own type', () => {
    assert.ok(new AccountNotFoundError('a@b.com') instanceof AccountNotFoundError);
    assert.ok(!(new NotFoundError('folder', 'x') instanceof AccountNotFoundError));
});

test('errors carry the detail a caller would otherwise have to parse out', () => {
    assert.equal(new AccountNotFoundError('ops@x.com').account, 'ops@x.com');
    assert.equal(new NotFoundError('template', 'x').kind, 'template');
    assert.equal(new UnsupportedPlatformError('linux').platform, 'linux');
    const script = new ScriptError({
        runner: 'powershell',
        script: '$x = 1',
        stderr: 'raw',
        durationMs: 12,
        message: 'boom',
        line: 7
    });
    assert.equal(script.script, '$x = 1');
    assert.equal(script.stderr, 'raw');
    assert.equal(script.message, 'boom');
    assert.equal(script.line, 7);
});

test('a tagged script message becomes the typed error, with the tag removed', () => {
    const notFound = errorFromTaggedMessage(`${failureTag('NOT_FOUND', 'folder')}Folder 'Invoices' not found.`);
    assert.ok(notFound instanceof NotFoundError);
    assert.equal(notFound.kind, 'folder');
    assert.equal(notFound.message, "Folder 'Invoices' not found.");

    const account = errorFromTaggedMessage(`${failureTag('ACCOUNT_NOT_FOUND')}Account 'o'brien@x.com' not found in this Outlook profile.`);
    assert.ok(account instanceof AccountNotFoundError);
    assert.equal(account.account, "o'brien@x.com");

    const invalid = errorFromTaggedMessage(`${failureTag('INVALID_REQUEST')}Mailbox can't send.`);
    assert.equal(invalid.code, 'INVALID_REQUEST');

    assert.equal(errorFromTaggedMessage('Some COM error'), null);
});

test('stripFailureTags leaves a readable message', () => {
    assert.equal(stripFailureTags(`${failureTag('NOT_FOUND', 'email')}No email.`), 'No email.');
});

test('classifyRunFailure turns a tagged failure into its code', () => {
    const error = classifyRunFailure({
        runner: 'powershell',
        script: 's',
        stderr: 'noise',
        message: `${failureTag('NOT_FOUND', 'email')}No email found for entry id 'x'.`,
        durationMs: 5,
        nodeError: Object.assign(new Error('Command failed'), {code: 1}),
    });
    assert.equal(error.code, 'NOT_FOUND');
    assert.equal(error.kind, 'email');
});

test('classifyRunFailure falls back to SCRIPT_FAILED, keeping the script and line', () => {
    const error = classifyRunFailure({
        runner: 'osascript',
        script: 'tell app',
        stderr: 'Microsoft Outlook got an error: -1728',
        durationMs: 5,
        line: 3,
    });
    assert.equal(error.code, 'SCRIPT_FAILED');
    assert.equal(error.runner, 'osascript');
    assert.equal(error.script, 'tell app');
    assert.equal(error.line, 3);
});

test('an aborted run is ABORTED even though it was killed like a timeout', () => {
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

test('an output overflow is OUTPUT_TOO_LARGE, not a timeout or a script failure', () => {
    const error = classifyRunFailure({
        runner: 'powershell',
        script: 's',
        stderr: '',
        durationMs: 10,
        nodeError: Object.assign(new RangeError('stdout maxBuffer length exceeded'), {
            code: 'ERR_CHILD_PROCESS_STDIO_MAXBUFFER',
            killed: true
        }),
        timeoutMs: 1000,
        maxBufferBytes: 1024,
    });
    assert.equal(error.code, 'OUTPUT_TOO_LARGE');
    assert.equal(error.maxBufferBytes, 1024);
});

test('a kill outranks a tagged message, which may be partial', () => {
    const error = classifyRunFailure({
        runner: 'powershell',
        script: 's',
        stderr: '',
        message: `${failureTag('ACCOUNT_NOT_FOUND')}Account 'x' not found`,
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
