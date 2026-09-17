// The package's error taxonomy.
//
// Every failure the package raises on purpose is an `OutlookError` carrying a
// stable `code`. The codes are the API; the messages beside them are not — a
// caller branches on `error.code`, never on message text, so a message can be
// reworded without breaking anyone. `err instanceof OutlookError` alone
// separates an expected failure from a bug.

/**
 * Stable discriminants.
 *
 * - `UNSUPPORTED_PLATFORM` — no Outlook automation exists on this OS.
 * - `NOT_IMPLEMENTED` — the platform could do this, but this corner isn't written.
 * - `ACCOUNT_NOT_FOUND` — no mailbox in the Outlook profile matches the address.
 * - `NOT_FOUND` — a folder, email, attachment, template, signature or file didn't resolve.
 * - `INVALID_REQUEST` — the arguments can't produce a call.
 * - `SCRIPT_FAILED` — the automation script ran and failed.
 * - `OUTPUT_TOO_LARGE` — the script printed more than `maxBufferBytes` and was killed.
 * - `TIMEOUT` — the run exceeded its time budget and was killed.
 * - `ABORTED` — the caller's AbortSignal fired.
 */
export type OutlookErrorCode =
    | 'UNSUPPORTED_PLATFORM'
    | 'NOT_IMPLEMENTED'
    | 'ACCOUNT_NOT_FOUND'
    | 'NOT_FOUND'
    | 'INVALID_REQUEST'
    | 'SCRIPT_FAILED'
    | 'OUTPUT_TOO_LARGE'
    | 'TIMEOUT'
    | 'ABORTED';

/** The interpreter a script ran under. */
export type ScriptRunner = 'powershell' | 'osascript';

/** What a NOT_FOUND error failed to find. */
export type NotFoundKind = 'folder' | 'email' | 'attachment' | 'template' | 'signature' | 'file';

const NOT_FOUND_KINDS: readonly NotFoundKind[] = ['folder', 'email', 'attachment', 'template', 'signature', 'file'];

/** Base class for every error this package raises deliberately. */
export class OutlookError extends Error {
    readonly code: OutlookErrorCode;

    /** The underlying failure, when there was one. */
    readonly cause?: unknown;

    constructor(code: OutlookErrorCode, message: string, options?: { cause?: unknown }) {
        // `cause` is an own enumerable field rather than `super`'s option: that
        // is what puts it in a JSON dump of the error, which is where someone
        // logging a failure looks for it.
        super(message);
        if (options && 'cause' in options) this.cause = options.cause;
        this.code = code;
        // `name` is what an unhandled-rejection dump shows, so make it the subclass.
        this.name = new.target.name;
    }
}

/** This OS has no Outlook automation. */
export class UnsupportedPlatformError extends OutlookError {
    /** The platform that was asked, i.e. `process.platform` at the time. */
    readonly platform: string;

    constructor(platform: string) {
        super('UNSUPPORTED_PLATFORM', `Outlook automation is only supported on Windows and macOS (running on '${platform}').`);
        this.platform = platform;
    }
}

/** The platform could do this, but this corner isn't written. */
export class NotImplementedError extends OutlookError {
    /** What isn't implemented. */
    readonly operation: string;
    readonly platform: string;

    constructor(operation: string, platform: string) {
        super('NOT_IMPLEMENTED', `${operation} is not implemented for Outlook on ${platform}.`);
        this.operation = operation;
        this.platform = platform;
    }
}

/** No mailbox in the Outlook profile matches the requested address. */
export class AccountNotFoundError extends OutlookError {
    readonly account: string;

    constructor(account: string, message = `Outlook account '${account}' not found.`) {
        super('ACCOUNT_NOT_FOUND', message);
        this.account = account;
    }
}

/** A named folder, email, attachment, template, signature or file didn't resolve. */
export class NotFoundError extends OutlookError {
    readonly kind: NotFoundKind;

    constructor(kind: NotFoundKind, message: string) {
        super('NOT_FOUND', message);
        this.kind = kind;
    }
}

/** The arguments can't produce a call. */
export class InvalidRequestError extends OutlookError {
    constructor(message: string) {
        super('INVALID_REQUEST', message);
    }
}

/**
 * The script ran and failed.
 *
 * Carries the generated script verbatim — the most useful thing to have when
 * one of these fires, and otherwise reachable only through a debug hook wired
 * up in advance.
 */
export class ScriptError extends OutlookError {
    readonly runner: ScriptRunner;
    readonly script: string;
    readonly stderr: string;
    readonly durationMs: number;
    /** The line of `script` the failure was raised on, when the runner reports one. */
    readonly line?: number;

    constructor(init: {
        runner: ScriptRunner;
        script: string;
        stderr: string;
        durationMs: number;
        message?: string;
        line?: number;
        cause?: unknown;
    }) {
        super('SCRIPT_FAILED', init.message || init.stderr || `${init.runner} failed with no output.`, {cause: init.cause});
        this.runner = init.runner;
        this.script = init.script;
        this.stderr = init.stderr;
        this.durationMs = init.durationMs;
        if (init.line !== undefined) this.line = init.line;
    }
}

/** The script printed more than `maxBufferBytes` and was killed. */
export class OutputTooLargeError extends OutlookError {
    readonly runner: ScriptRunner;
    readonly maxBufferBytes: number;
    readonly script: string;

    constructor(init: { runner: ScriptRunner; maxBufferBytes: number; script: string }) {
        super(
            'OUTPUT_TOO_LARGE',
            `${init.runner} output exceeded maxBufferBytes (${init.maxBufferBytes}). `
            + 'Raise maxBufferBytes, or narrow the request.',
        );
        this.runner = init.runner;
        this.maxBufferBytes = init.maxBufferBytes;
        this.script = init.script;
    }
}

/** The run exceeded its time budget and was killed. */
export class TimeoutError extends OutlookError {
    readonly runner: ScriptRunner;
    readonly timeoutMs: number;
    readonly script: string;

    constructor(init: { runner: ScriptRunner; timeoutMs: number; script: string }) {
        super('TIMEOUT', `${init.runner} run exceeded its ${init.timeoutMs}ms budget and was killed.`);
        this.runner = init.runner;
        this.timeoutMs = init.timeoutMs;
        this.script = init.script;
    }
}

/** The caller's AbortSignal fired. */
export class AbortedError extends OutlookError {
    readonly runner: ScriptRunner;

    constructor(runner: ScriptRunner, reason?: unknown) {
        super('ABORTED', `${runner} run was aborted by the caller.`, {cause: reason});
        this.runner = runner;
    }
}

// ── Typed failures from inside a script ─────────────────────────────────
//
// A generated script can only fail by printing a message, so a script that
// knows WHY it failed — the folder isn't there, the id resolves to nothing —
// says so by starting the message with a tag: `[outlook-bridge:NOT_FOUND:folder]`.
// The runner turns a tagged message into the matching error class and strips
// the tag, so a caller can branch on NOT_FOUND without reading the sentence.

/** The codes a script may raise by tagging its failure message. */
export type ScriptFailureCode = 'ACCOUNT_NOT_FOUND' | 'NOT_FOUND' | 'INVALID_REQUEST';

/**
 * The tag that opens a typed script failure, ready to be followed by the
 * human-readable message. Plain ASCII with no quote or `$`, so it can sit
 * inside a PowerShell or AppleScript string literal unescaped.
 */
export function failureTag(code: 'ACCOUNT_NOT_FOUND' | 'INVALID_REQUEST'): string;
export function failureTag(code: 'NOT_FOUND', kind: NotFoundKind): string;
export function failureTag(code: ScriptFailureCode, kind?: NotFoundKind): string {
    return `[outlook-bridge:${code}${kind ? `:${kind}` : ''}] `;
}

const FAILURE_TAG = /\[outlook-bridge:(ACCOUNT_NOT_FOUND|NOT_FOUND|INVALID_REQUEST)(?::([a-z]+))?\]\s*/;

/** A script's account-not-found message names the address in single quotes. */
const QUOTED_ACCOUNT = /'(.*)'/;

/**
 * The typed error a tagged script message stands for, or null when the
 * message carries no tag.
 */
export function errorFromTaggedMessage(message: string): OutlookError | null {
    const match = FAILURE_TAG.exec(message);
    if (!match) return null;
    const text = (message.slice(0, match.index) + message.slice(match.index + match[0].length)).trim();
    switch (match[1]) {
        case 'ACCOUNT_NOT_FOUND':
            return new AccountNotFoundError(QUOTED_ACCOUNT.exec(text)?.[1] ?? '', text);
        case 'NOT_FOUND': {
            const kind = NOT_FOUND_KINDS.find(k => k === match[2]) ?? 'email';
            return new NotFoundError(kind, text);
        }
        default:
            return new InvalidRequestError(text);
    }
}

/** Remove any failure tags from a message meant for display — e.g. a per-item error. */
export function stripFailureTags(message: string): string {
    return message.replace(new RegExp(FAILURE_TAG.source, 'g'), '').trim();
}

/**
 * Turn a failed run into the most specific error that can be proven.
 *
 * Order matters. An abort and a timeout both kill the child, so the signal's
 * own state is checked before the kill; an overflow kills it too, and is told
 * apart by its error code. Only then is the script's message worth reading — a
 * process killed mid-write may have printed a partial, misleading one.
 */
export function classifyRunFailure(init: {
    runner: ScriptRunner;
    script: string;
    /** Everything the run wrote to stderr. */
    stderr: string;
    /** The script's own failure message, when the runner could isolate one. */
    message?: string;
    /** The script line the failure was raised on, when known. */
    line?: number;
    durationMs: number;
    /** The error `execFile` handed back. */
    nodeError?: (Error & { killed?: boolean; code?: string | number | null }) | null;
    /** The time budget in force, for the TIMEOUT message. */
    timeoutMs?: number;
    /** The stdout cap in force, for the OUTPUT_TOO_LARGE message. */
    maxBufferBytes?: number;
    signal?: AbortSignal;
}): OutlookError {
    const {runner, script, stderr, durationMs, nodeError, timeoutMs, signal} = init;

    if (signal?.aborted || nodeError?.name === 'AbortError' || nodeError?.code === 'ABORT_ERR') {
        return new AbortedError(runner, signal?.reason);
    }
    if (nodeError?.code === 'ERR_CHILD_PROCESS_STDIO_MAXBUFFER') {
        return new OutputTooLargeError({runner, maxBufferBytes: init.maxBufferBytes ?? 0, script});
    }
    if (nodeError?.killed && timeoutMs) {
        return new TimeoutError({runner, timeoutMs, script});
    }
    const message = (init.message ?? '').trim() || stderr.trim() || nodeError?.message || '';
    const typed = errorFromTaggedMessage(message);
    if (typed) return typed;
    return new ScriptError({
        runner,
        script,
        stderr,
        durationMs,
        message,
        line: init.line,
        cause: nodeError ?? undefined
    });
}
