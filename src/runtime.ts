// Process-wide knobs and shared scratch-file handling for the generated
// PowerShell / AppleScript runs.
//
// Every call in this package works by generating a script and shelling out to an
// interpreter, which leaves a consumer three things they cannot otherwise reach:
// how long a run may take, how much it may print, and what the script actually
// said when it failed. Those are the three settings here.
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
    /** `true` logs every generated script to stderr; a function receives each run
     *  instead. Defaults to on when OUTLOOK_BRIDGE_DEBUG is set to a non-empty,
     *  non-'0' value. */
    debug?: boolean | ((event: BridgeDebugEvent) => void);
}

type ResolvedConfig = Required<Pick<BridgeOptions, 'timeoutMs' | 'maxBufferBytes' | 'tempDir'>>
    & Pick<BridgeOptions, 'debug'>;

const envDebug = process.env.OUTLOOK_BRIDGE_DEBUG;

const config: ResolvedConfig = {
    timeoutMs: 120_000,
    maxBufferBytes: 8 * 1024 * 1024,
    tempDir: os.tmpdir(),
    debug: !!envDebug && envDebug !== '0',
};

/**
 * Override the defaults above for this process. Merges, so naming one option
 * leaves the rest alone. Call before the first automation call.
 */
export function configure(options: BridgeOptions): void {
    if (options.timeoutMs !== undefined) config.timeoutMs = Math.max(0, options.timeoutMs);
    if (options.maxBufferBytes !== undefined) config.maxBufferBytes = Math.max(1, options.maxBufferBytes);
    if (options.tempDir !== undefined) config.tempDir = options.tempDir;
    if (options.debug !== undefined) config.debug = options.debug;
}

/** The settings currently in force. */
export function getConfig(): Readonly<ResolvedConfig> {
    return config;
}

/** Report one finished script run to the configured debug hook. */
export function reportRun(event: BridgeDebugEvent): void {
    const {debug} = config;
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
    return path.join(config.tempDir, `outlook-bridge-${kind}-${unique}.${extension}`);
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
    return fs.mkdtempSync(path.join(config.tempDir, 'outlook-bridge-attachments-'));
}

/** Resolve a caller-supplied destination, creating it when absent. */
export function resolveDestDir(destDir?: string): string {
    if (!destDir) return makeAttachmentDir();
    fs.mkdirSync(destDir, {recursive: true});
    return destDir;
}
