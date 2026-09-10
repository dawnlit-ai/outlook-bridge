// Process-wide knobs and shared scratch-file handling for the generated
// PowerShell / AppleScript runs.
//
// Every call in this package works by generating a script and shelling out to an
// interpreter, which leaves a consumer four things they cannot otherwise reach:
// how long a run may take, how much it may print, how to call it off, and what
// the script actually said when it failed. Those are the settings here.
import { AsyncLocalStorage } from 'async_hooks';
import fs from 'fs';
import os from 'os';
import path from 'path';

/** One completed script run, handed to the `debug` hook. */
export interface BridgeDebugEvent {
    /** Which interpreter ran it. */
    runner: 'powershell' | 'osascript';
    /** The generated script, verbatim — the thing you actually want when a run fails. */
    script: string;
    durationMs: number;
    /** Present only when the run failed. */
    error?: string;
}

export interface BridgeOptions {
    /**
     * Default ceiling (ms) on a single script run; 0 disables it. Calls that
     * already carry a specific budget (a full-mailbox scan, a purge) keep their
     * own and ignore this.
     */
    timeoutMs?: number;
    /**
     * Max bytes a script may write to stdout before the run is killed. A read
     * that returns message bodies is what pushes against this — raise it before
     * scanning a mailbox that returns thousands of matches.
     */
    maxBufferBytes?: number;
    /** Directory for generated scripts and for saved attachments when the caller
     *  names no destination. Defaults to the OS temp directory. */
    tempDir?: string;
    /**
     * Cancels in-flight runs: the interpreter is killed and the call rejects with
     * an `AbortedError`. A timeout only caps a run's length — this is how a UI
     * with a Cancel button, or a shutting-down process, calls one off early.
     *
     * Most useful scoped to the calls you want to cancel rather than set process
     * wide: `bridge.withOptions({ signal }).searchInboxByFilter(...)`.
     */
    signal?: AbortSignal;
    /** `true` logs every generated script to stderr; a function receives each run
     *  instead. Defaults to on when OUTLOOK_BRIDGE_DEBUG is set to a non-empty,
     *  non-'0' value. */
    debug?: boolean | ((event: BridgeDebugEvent) => void);
}

export type ResolvedConfig = Required<Pick<BridgeOptions, 'timeoutMs' | 'maxBufferBytes' | 'tempDir'>>
    & Pick<BridgeOptions, 'debug' | 'signal'>;

const envDebug = process.env.OUTLOOK_BRIDGE_DEBUG;

/** The process-wide defaults, as mutated by `configure()`. */
const globalConfig: ResolvedConfig = {
    timeoutMs: 120_000,
    maxBufferBytes: 8 * 1024 * 1024,
    tempDir: os.tmpdir(),
    debug: !!envDebug && envDebug !== '0',
};

/**
 * Per-instance config, scoped to an async call tree.
 *
 * `createOutlookBridge()` hands each instance its own settings, but the platform
 * services are thousands of lines of module-level functions that read
 * `getConfig()` directly. Threading a config parameter through all of them would
 * touch every line for no behavioural gain; an AsyncLocalStorage lets an instance
 * wrap its calls instead, and the store follows across every `await` inside one. Two
 * bridges with different timeouts can then run concurrently in one process,
 * which is the thing a second consumer actually needs and the global could never
 * give them.
 */
const scoped = new AsyncLocalStorage<ResolvedConfig>();

/** Merge caller options over a base, ignoring the keys they left out. */
export function mergeOptions(base: ResolvedConfig, options: BridgeOptions = {}): ResolvedConfig {
    const next: ResolvedConfig = { ...base };
    if (options.timeoutMs !== undefined) next.timeoutMs = Math.max(0, options.timeoutMs);
    if (options.maxBufferBytes !== undefined) next.maxBufferBytes = Math.max(1, options.maxBufferBytes);
    if (options.tempDir !== undefined) next.tempDir = options.tempDir;
    if (options.signal !== undefined) next.signal = options.signal;
    if (options.debug !== undefined) next.debug = options.debug;
    return next;
}

/**
 * Override the process-wide defaults. Merges, so naming one option leaves the
 * rest alone. Call before the first automation call.
 *
 * Affects every caller in the process. Prefer `createOutlookBridge(options)` in
 * anything that shares a process with code you don't own.
 */
export function configure(options: BridgeOptions): void {
    Object.assign(globalConfig, mergeOptions(globalConfig, options));
}

/** The settings in force for the current call — the enclosing scope's, else global. */
export function getConfig(): Readonly<ResolvedConfig> {
    return scoped.getStore() ?? globalConfig;
}

/** The process-wide defaults, ignoring any active scope. */
export function getGlobalConfig(): Readonly<ResolvedConfig> {
    return globalConfig;
}

/** Run `fn` with `config` in force for it and everything it awaits. */
export function withConfig<T>(config: ResolvedConfig, fn: () => T): T {
    return scoped.run(config, fn);
}

/** Report one finished script run to the configured debug hook. */
export function reportRun(event: BridgeDebugEvent): void {
    const { debug } = getConfig();
    if (!debug) return;
    if (typeof debug === 'function') {
        debug(event);
        return;
    }
    const outcome = event.error ? `FAILED: ${event.error}` : 'ok';
    process.stderr.write(
        `[outlook-bridge] ${event.runner} ${event.durationMs}ms ${outcome}\n${event.script}\n`,
    );
}

/**
 * Path for a scratch file in the configured temp directory. The random suffix
 * matters: two concurrent calls landing in the same millisecond would otherwise
 * pick the same name and overwrite each other's script mid-run.
 */
export function tempFile(kind: string, extension: string): string {
    const unique = `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
    return path.join(getConfig().tempDir, `outlook-bridge-${kind}-${unique}.${extension}`);
}

/**
 * A fresh directory for one call's saved attachments.
 *
 * Deliberately per-call rather than one shared folder: attachments are saved
 * under the name the sender gave them, so two emails that both carry
 * "invoice.pdf" would otherwise silently overwrite each other — and so would two
 * unrelated programs using this package on the same machine.
 */
export function makeAttachmentDir(): string {
    return fs.mkdtempSync(path.join(getConfig().tempDir, 'outlook-bridge-attachments-'));
}

/** Resolve a caller-supplied destination, creating it when absent. */
export function resolveDestDir(destDir?: string): string {
    if (!destDir) return makeAttachmentDir();
    fs.mkdirSync(destDir, { recursive: true });
    return destDir;
}
