// The package's error taxonomy.
//
// Every failure used to be a bare `new Error(string)`, which left a consumer
// only one way to tell "this machine can't do that" from "that account doesn't
// exist" from "the script blew up": matching on the message text. Message text
// is not an API — rewording one sentence here would silently break a caller's
// branch — so the distinctions a caller actually acts on carry a stable `code`.
//
// Every error this package throws on purpose is an `OutlookError`. A caller that
// wants one branch for "expected failure" and another for "bug" can switch on
// `err instanceof OutlookError` alone.

/**
 * Stable discriminants. These are the API; the messages beside them are not.
 *
 * - `UNSUPPORTED_PLATFORM` — no Outlook automation exists on this OS at all.
 * - `NOT_IMPLEMENTED` — the platform could do this, but the port isn't written.
 *   Distinct from UNSUPPORTED_PLATFORM on purpose: this one is a gap that may
 *   close in a later release, so a caller may reasonably feature-detect and
 *   degrade rather than refuse outright.
 * - `ACCOUNT_NOT_FOUND` — no configured account matches that SMTP address.
 * - `NOT_FOUND` — a named folder, template, signature or item didn't resolve.
 * - `INVALID_REQUEST` — the arguments can't produce a call; nothing was attempted.
 * - `SCRIPT_FAILED` — the interpreter ran and reported failure.
 * - `TIMEOUT` — the run exceeded its budget and was killed.
 * - `ABORTED` — the caller's AbortSignal fired.
 */
export type OutlookErrorCode =
    | 'UNSUPPORTED_PLATFORM'
    | 'NOT_IMPLEMENTED'
    | 'ACCOUNT_NOT_FOUND'
    | 'NOT_FOUND'
    | 'INVALID_REQUEST'
    | 'SCRIPT_FAILED'
    | 'TIMEOUT'
    | 'ABORTED';

/** Base class for every error this package raises deliberately. */
export class OutlookError extends Error {
    readonly code: OutlookErrorCode;

    /** The underlying failure, when there was one. */
    readonly cause?: unknown;

    constructor(code: OutlookErrorCode, message: string, options?: { cause?: unknown }) {
        // Declared as a field rather than passed to `super`'s `cause` option: an
        // own enumerable property is what puts it in a JSON dump of the error,
        // which is where a consumer logging a failure actually looks for it.
        super(message);
        if (options && 'cause' in options) this.cause = options.cause;
        this.code = code;
        // `name` is what shows up in an unhandled-rejection dump, so make it the
        // subclass rather than a uniform "Error".
        this.name = new.target.name;
    }
}

/** This OS has no Outlook automation backend. */
export class UnsupportedPlatformError extends OutlookError {
    /** The platform that was asked, i.e. `process.platform` at the time. */
    readonly platform: string;

    constructor(platform: string, detail = 'Outlook automation is only supported on Windows and macOS.') {
        super('UNSUPPORTED_PLATFORM', `${detail} (running on '${platform}')`);
        this.platform = platform;
    }
}

/** The platform supports this, but the port isn't written yet. */
export class NotImplementedError extends OutlookError {
    /** The bridge function that isn't ported. */
    readonly operation: string;
    readonly platform: string;

    constructor(operation: string, platform: string) {
        super('NOT_IMPLEMENTED', `'${operation}' is not implemented for Outlook on ${platform} yet.`);
        this.operation = operation;
        this.platform = platform;
    }
}

/** No configured Outlook account matches the requested address. */
export class AccountNotFoundError extends OutlookError {
    readonly account: string;

    constructor(account: string) {
        super('ACCOUNT_NOT_FOUND', `Outlook account '${account}' not found.`);
        this.account = account;
    }
}

/** A named folder, template, signature or item didn't resolve. */
export class NotFoundError extends OutlookError {
    /** What kind of thing was looked up — 'folder', 'template', 'signature', 'email'. */
    readonly kind: string;

    constructor(kind: string, message: string) {
        super('NOT_FOUND', message);
        this.kind = kind;
    }
}

/** The arguments can't produce a call; nothing was attempted. */
export class InvalidRequestError extends OutlookError {
    constructor(message: string) {
        super('INVALID_REQUEST', message);
    }
}

/**
 * The interpreter ran and failed.
 *
 * Carries the generated script verbatim. That is the single most useful thing
 * when one of these fires and it is otherwise unreachable — the `debug` hook
 * only fires if the consumer wired one up ahead of time, whereas this arrives
 * attached to the failure itself.
 */
export class ScriptError extends OutlookError {
    readonly runner: 'powershell' | 'osascript';
    readonly script: string;
    readonly stderr: string;
    readonly durationMs: number;

    constructor(init: {
        runner: 'powershell' | 'osascript';
        script: string;
        stderr: string;
        durationMs: number;
        cause?: unknown;
    }) {
        super('SCRIPT_FAILED', init.stderr || `${init.runner} failed with no output.`, { cause: init.cause });
        this.runner = init.runner;
        this.script = init.script;
        this.stderr = init.stderr;
        this.durationMs = init.durationMs;
    }
}

/** The run exceeded its budget and was killed. */
export class TimeoutError extends OutlookError {
    readonly runner: 'powershell' | 'osascript';
    readonly timeoutMs: number;
    readonly script: string;

    constructor(init: { runner: 'powershell' | 'osascript'; timeoutMs: number; script: string }) {
        super('TIMEOUT', `${init.runner} run exceeded its ${init.timeoutMs}ms budget and was killed.`);
        this.runner = init.runner;
        this.timeoutMs = init.timeoutMs;
        this.script = init.script;
    }
}

/** The caller's AbortSignal fired. */
export class AbortedError extends OutlookError {
    readonly runner: 'powershell' | 'osascript';

    constructor(runner: 'powershell' | 'osascript', reason?: unknown) {
        super('ABORTED', `${runner} run was aborted by the caller.`, { cause: reason });
        this.runner = runner;
    }
}

/**
 * Both platform scripts signal a bad sending account by `throw`ing a string the
 * interpreter then prints. Recognising it here is what turns the most common
 * operational mistake — a typo'd or freshly-removed account — into a code a
 * caller can branch on, instead of a SCRIPT_FAILED they'd have to grep.
 *
 * Deliberately narrow: it only matches the sentence our own scripts emit.
 */
const ACCOUNT_NOT_FOUND_RE = /Account '([^']*)' not found/i;

/**
 * Upgrade a raw interpreter failure to the most specific error we can prove.
 *
 * Order matters. Abort and timeout both kill the child with SIGTERM, so they are
 * indistinguishable from the exit status alone — the signal's own state is what
 * separates them, and it is checked first. Only then is stderr worth reading:
 * a killed process may have printed a partial, misleading message before dying.
 */
export function classifyRunFailure(init: {
    runner: 'powershell' | 'osascript';
    script: string;
    stderr: string;
    durationMs: number;
    /** The error `execFile` handed back, when the failure came from the child. */
    nodeError?: (Error & { killed?: boolean; code?: string | number | null }) | null;
    /** The budget in force, for the TIMEOUT message. */
    timeoutMs?: number;
    signal?: AbortSignal;
}): OutlookError {
    const { runner, script, stderr, durationMs, nodeError, timeoutMs, signal } = init;

    if (signal?.aborted || nodeError?.name === 'AbortError' || nodeError?.code === 'ABORT_ERR') {
        return new AbortedError(runner, signal?.reason);
    }
    if (nodeError?.killed && timeoutMs) {
        return new TimeoutError({ runner, timeoutMs, script });
    }
    const match = ACCOUNT_NOT_FOUND_RE.exec(stderr);
    if (match) return new AccountNotFoundError(match[1]);
    return new ScriptError({ runner, script, stderr, durationMs, cause: nodeError ?? undefined });
}
