// Getting a generated script to Windows PowerShell and its output back.
//
// Every Windows function here is "build a script, run it, coerce the JSON", so
// this module owns the two halves nobody should reimplement: the run itself
// (encoding, timeout, cancellation, failure classification, the debug hook) and
// the escaping rules for putting caller text into a script safely.
import { execFile } from 'child_process';
import fs from 'fs';
import { getConfig, reportRun, tempFile } from '../runtime';
import { classifyRunFailure, UnsupportedPlatformError } from '../errors';

const RUNNER = 'powershell' as const;

/**
 * Refuse before generating anything when this isn't Windows.
 *
 * Called at the top of each operation rather than only inside the runner, so a
 * call that would have written temp files or made preparatory reads doesn't do
 * any of that first.
 */
export function requireWindows(): void {
    if (process.platform !== 'win32') {
        throw new UnsupportedPlatformError(
            process.platform,
            'Outlook COM automation is only supported on Windows.',
        );
    }
}

/**
 * Escape a string for a PowerShell SINGLE-quoted literal — the only context
 * caller-supplied text may go in.
 *
 * A double-quoted PowerShell string expands `$(...)` subexpressions, so text
 * placed there would execute. Where a generated script needs user text inside a
 * double-quoted string (a Restrict query, which must expand `$cutoff`), assign
 * it to a variable with this first and concatenate — see read.ts.
 */
export function psEscape(s: string): string {
    return s.replace(/'/g, "''");
}

/** A PowerShell array literal of single-quoted strings. */
export function psList(values: readonly string[]): string {
    return values.map(v => `'${psEscape(v)}'`).join(',');
}

/** `$true` / `$false` — PowerShell has no bare boolean literal. */
export function psBool(value: boolean): string {
    return value ? '$true' : '$false';
}

// Force UTF-8 on stdout. Windows PowerShell 5.1 otherwise encodes output in the
// OEM/ANSI console code page, whose "best-fit" mapping silently rewrites
// non-ASCII characters — e.g. a curly quote (U+201C) in an email body becomes a
// plain " — which corrupts the JSON these scripts emit (an unescaped quote) and
// breaks JSON.parse. Setting the output encoding first makes non-ASCII survive
// as real UTF-8 bytes, which Node then decodes correctly.
const UTF8_PRELUDE = '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n';

/**
 * Run a generated script and return its stdout.
 *
 * `timeout` overrides the configured default for this one call — pass it where a
 * call has a genuinely different budget (a full-mailbox walk, a purge) rather
 * than relying on the global. See `configure()` for both knobs.
 */
export function runPowerShell(script: string, timeout?: number): Promise<string> {
    if (process.platform !== 'win32') {
        return Promise.reject(new UnsupportedPlatformError(
            process.platform,
            'PowerShell and COM automation are only supported on Windows.',
        ));
    }
    const utf8Script = UTF8_PRELUDE + script;
    return spawnPowerShell(
        ['-NoProfile', '-NonInteractive', '-Command', utf8Script],
        utf8Script,
        timeout,
    );
}

/**
 * Run a generated script from a temp `.ps1` file rather than `-Command`.
 *
 * For the two calls that drive a compose window: a script passed on the command
 * line competes with the command-line length limit, and these carry enough
 * inline text to reach it. Failures classify exactly as `runPowerShell`'s do —
 * the file is an implementation detail, not a second error contract.
 */
export async function runPowerShellFile(script: string, timeout?: number): Promise<string> {
    if (process.platform !== 'win32') {
        throw new UnsupportedPlatformError(
            process.platform,
            'PowerShell and COM automation are only supported on Windows.',
        );
    }
    const scriptFile = tempFile('script', 'ps1');
    fs.writeFileSync(scriptFile, UTF8_PRELUDE + script, 'utf-8');
    try {
        return await spawnPowerShell(
            ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptFile],
            script,
            timeout,
        );
    } finally {
        try {
            fs.unlinkSync(scriptFile);
        } catch { /* the run is done; a leftover scratch file is not worth failing on */
        }
    }
}

/** The one place a PowerShell child is spawned, timed, classified and reported. */
function spawnPowerShell(args: string[], script: string, timeout?: number): Promise<string> {
    const {timeoutMs, maxBufferBytes, signal} = getConfig();
    const effectiveTimeout = timeout ?? timeoutMs;
    const startedAt = Date.now();
    return new Promise((resolve, reject) => {
        execFile(
            'powershell.exe',
            args,
            {maxBuffer: maxBufferBytes, timeout: effectiveTimeout, signal},
            (error, stdout, stderr) => {
                const durationMs = Date.now() - startedAt;
                if (error) {
                    const failure = classifyRunFailure({
                        runner: RUNNER,
                        script,
                        stderr: (stderr || error.message).trim(),
                        durationMs,
                        nodeError: error,
                        timeoutMs: effectiveTimeout,
                        signal,
                    });
                    reportRun({runner: RUNNER, script, durationMs, error: failure.message});
                    reject(failure);
                } else {
                    reportRun({runner: RUNNER, script, durationMs});
                    resolve(stdout.trim());
                }
            },
        );
    });
}

/**
 * Write text a script will read back, and hand back a cleanup callback.
 *
 * Large HTML — an email body, a Word-sized template — goes to a file rather than
 * into the script text: inline, it is what pushes a generated script past the
 * command-line length limit.
 */
export function scriptInput(kind: string, contents: string): { path: string; cleanup(): void } {
    const path = tempFile(kind, 'html');
    fs.writeFileSync(path, contents, 'utf-8');
    return {
        path,
        cleanup() {
            try {
                fs.unlinkSync(path);
            } catch { /* ignore */
            }
        },
    };
}

/** Read a file written by a script, as UTF-8, or '' when it never appeared. */
export function readScriptOutput(path: string): string {
    try {
        const contents = fs.readFileSync(path, 'utf-8');
        fs.unlinkSync(path);
        return contents;
    } catch {
        return '';
    }
}
