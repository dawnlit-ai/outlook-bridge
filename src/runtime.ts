// Settings, and the scratch files every run needs.
//
// Each operation generates a script and runs it under an interpreter, which
// leaves a caller four things they cannot otherwise reach: how long a run may
// take, how much it may print, how to call it off, and what the script actually
// said. Those are the settings here — process-wide through `configure()`, or per
// bridge through `createOutlookBridge()`.
import { AsyncLocalStorage } from 'async_hooks';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { InvalidRequestError, type ScriptRunner } from './errors';

/** One completed script run, handed to the `debug` hook. */
export interface BridgeDebugEvent {
    runner: ScriptRunner;
    /** The script exactly as it ran. */
    script: string;
    durationMs: number;
    /** Present only when the run failed. */
    error?: string;
}

export interface BridgeOptions {
    /**
     * Ceiling on every script run, in milliseconds, replacing each operation's
     * own budget; 0 disables timeouts. Leave unset (or pass null to go back) to
     * let each operation use a budget sized for its work — a minute for a
     * lookup, up to ten for purging Deleted Items. `editEmailTemplate` waits on
     * a person and is never timed; cancel it with `signal`.
     */
    timeoutMs?: number | null;
    /**
     * Max bytes a script may print before the run is killed with
     * OUTPUT_TOO_LARGE. Searches that return whole bodies push against this.
     * Default 8 MiB.
     */
    maxBufferBytes?: number;
    /** Directory for generated scripts, and for saved attachments when a call
     *  names no destination. Default: the OS temp directory. */
    tempDir?: string;
    /**
     * Cancels in-flight runs: the interpreter is killed and the call rejects
     * with ABORTED. Most useful scoped to the calls it should cancel:
     * `bridge.withOptions({ signal }).searchInboxByFilter(...)`.
     */
    signal?: AbortSignal;
    /** `true` logs every script to stderr; a function receives each run
     *  instead. Default: on when OUTLOOK_BRIDGE_DEBUG is set to anything but '' or '0'. */
    debug?: boolean | ((event: BridgeDebugEvent) => void);
}

/** The settings a call runs with, every default applied. */
export interface ResolvedConfig {
    /** undefined = each operation's own budget. */
    readonly timeoutMs: number | undefined;
    readonly maxBufferBytes: number;
    readonly tempDir: string;
    readonly signal: AbortSignal | undefined;
    readonly debug: boolean | ((event: BridgeDebugEvent) => void);
}

const envDebug = process.env.OUTLOOK_BRIDGE_DEBUG;

/** The process-wide settings, as `configure()` leaves them. */
let globalConfig: ResolvedConfig = {
    timeoutMs: undefined,
    maxBufferBytes: 8 * 1024 * 1024,
    tempDir: os.tmpdir(),
    signal: undefined,
    debug: !!envDebug && envDebug !== '0',
};

/**
 * Per-bridge settings, scoped to an async call tree.
 *
 * The platform code reads its settings through `getConfig()` wherever it needs
 * them rather than having a config threaded through every function. A bridge
 * instance runs each call inside `withConfig`, and the store follows every
 * `await` beneath it — so two bridges with different timeouts can run at once
 * in one process without seeing each other's.
 */
const scoped = new AsyncLocalStorage<ResolvedConfig>();

/** Caller options merged over a base; keys left out keep the base's value. */
export function mergeOptions(base: ResolvedConfig, options: BridgeOptions = {}): ResolvedConfig {
    const next = {...base};
    if (options.timeoutMs === null) next.timeoutMs = undefined;
    else if (options.timeoutMs !== undefined) next.timeoutMs = nonNegative(options.timeoutMs, 'timeoutMs');
    if (options.maxBufferBytes !== undefined) next.maxBufferBytes = Math.max(1, nonNegative(options.maxBufferBytes, 'maxBufferBytes'));
    if (options.tempDir !== undefined) next.tempDir = options.tempDir;
    if (options.signal !== undefined) next.signal = options.signal;
    if (options.debug !== undefined) next.debug = options.debug;
    return next;
}

function nonNegative(value: number, name: string): number {
    if (typeof value !== 'number' || Number.isNaN(value)) {
        throw new InvalidRequestError(`${name} must be a number (got ${String(value)}).`);
    }
    return Math.max(0, Math.floor(value));
}

/**
 * Change the process-wide settings. Merges, so naming one option leaves the
 * rest alone.
 *
 * Affects every caller in the process that uses the exported functions. Prefer
 * `createOutlookBridge(options)` in a process shared with code you don't own.
 */
export function configure(options: BridgeOptions): void {
    globalConfig = mergeOptions(globalConfig, options);
}

/** The settings in force for the current call: the enclosing bridge's, else the process-wide ones. */
export function getConfig(): ResolvedConfig {
    return scoped.getStore() ?? globalConfig;
}

/** The process-wide settings, ignoring any bridge scope. */
export function getGlobalConfig(): ResolvedConfig {
    return globalConfig;
}

/** Run `fn` with `config` in force for it and everything it awaits. */
export function withConfig<T>(config: ResolvedConfig, fn: () => T): T {
    return scoped.run(config, fn);
}

/**
 * How long each kind of run may take when no `timeoutMs` is set.
 *
 * Generous on purpose: a budget is a guard against a hung interpreter, not a
 * performance target, and the first call against an Outlook that isn't running
 * yet spends much of its budget starting it.
 */
export const RUN_BUDGETS = {
    /** One lookup: accounts, signatures, opening or reading a single email. */
    quick: 60_000,
    /** Composing, filing, saving attachments, or reading one folder. */
    standard: 120_000,
    /** Walking a mailbox tree, or acting on a batch of items. */
    scan: 300_000,
    /** Emptying Deleted Items. */
    purge: 600_000,
} as const;

/** A budget name, or 'interactive' for a run that waits on a person and is never timed. */
export type RunBudget = keyof typeof RUN_BUDGETS | 'interactive';

/** The timeout one run gets: the configured ceiling if there is one, else its budget. */
export function timeoutFor(budget: RunBudget): number {
    if (budget === 'interactive') return 0;
    return getConfig().timeoutMs ?? RUN_BUDGETS[budget];
}

/** Report one finished run to the configured debug hook. */
export function reportRun(event: BridgeDebugEvent): void {
    const {debug} = getConfig();
    if (!debug) return;
    if (typeof debug === 'function') {
        debug(event);
        return;
    }
    const outcome = event.error ? `FAILED: ${event.error}` : 'ok';
    process.stderr.write(`[outlook-bridge] ${event.runner} ${event.durationMs}ms ${outcome}\n${event.script}\n`);
}

/**
 * A scratch-file path in the configured temp directory. The random part
 * matters: two concurrent calls in the same millisecond would otherwise pick the
 * same name and overwrite each other's script mid-run.
 */
export function tempFile(kind: string, extension: string): string {
    const unique = `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
    return path.join(getConfig().tempDir, `outlook-bridge-${kind}-${unique}.${extension}`);
}

/** Delete a scratch file, ignoring one that is already gone. */
export function removeTempFile(file: string): void {
    try {
        fs.unlinkSync(file);
    } catch {
        // The run is over either way; a leftover scratch file isn't worth failing it.
    }
}

/** A destination for saved attachments, and how to undo creating it. */
export interface AttachmentDestination {
    dir: string;

    /** Remove the directory if this call created it and left it empty. */
    discardIfUnused(): void;
}

/**
 * Where one call saves its attachments: the caller's directory (created if
 * absent), or a fresh private directory of its own.
 *
 * Private by default because attachments keep the names senders gave them, so
 * one shared folder means two emails carrying 'invoice.pdf' overwrite each
 * other — and so do two programs using this package on one machine.
 */
export function attachmentDestination(destDir?: string): AttachmentDestination {
    if (destDir) {
        const existed = fs.existsSync(destDir);
        fs.mkdirSync(destDir, {recursive: true});
        return {
            dir: destDir,
            discardIfUnused() {
                if (!existed) removeIfEmpty(destDir);
            },
        };
    }
    const dir = fs.mkdtempSync(path.join(getConfig().tempDir, 'outlook-bridge-attachments-'));
    return {dir, discardIfUnused: () => removeIfEmpty(dir)};
}

function removeIfEmpty(dir: string): void {
    try {
        if (fs.readdirSync(dir).length === 0) fs.rmdirSync(dir);
    } catch {
        // Nothing to undo if it can't be read.
    }
}
