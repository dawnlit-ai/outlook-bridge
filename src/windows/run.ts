// Getting a generated script to Windows PowerShell and its output back.
//
// Every Windows operation is "build a script, run it, read what it printed",
// so this module owns the parts nobody should reimplement: how a script is
// wrapped and run, how a failure is read back out of it, and the escaping rules
// for putting caller text into a script safely.
import fs from 'fs';
import { InvalidRequestError, ScriptError, UnsupportedPlatformError } from '../errors';
import { removeTempFile, type RunBudget, tempFile } from '../runtime';
import { runScriptProcess } from '../shared/exec';

const RUNNER = 'powershell' as const;

/** Opens the line a script's catch block writes its failure on. */
const FAILURE_MARKER = '@@outlook-bridge-failure@@';

/** Refuse before generating anything when this isn't Windows. */
export function requireWindows(): void {
    if (process.platform !== 'win32') throw new UnsupportedPlatformError(process.platform);
}

// ── Escaping ─────────────────────────────────────────────────────────────
//
// Caller text goes into a script ONLY as a single-quoted literal, built by
// `psString`. Two things make that the rule:
//  - A double-quoted PowerShell string expands `$(...)`, so caller text placed
//    in one runs as code. Where a message needs caller text, it is assigned to a
//    variable first and the double-quoted string names the variable — a
//    variable's value is never evaluated again.
//  - PowerShell accepts the typographic quotes ‘ ’ ‚ ‛ as single quotes too, so
//    doubling only the ASCII one leaves a way out of the literal. Every quote
//    character is doubled, which PowerShell reads back as that same character.

/** Escape text for the inside of a PowerShell single-quoted literal. */
export function psEscape(text: string): string {
    return text.replace(/['\u2018\u2019\u201A\u201B]/g, quote => quote + quote);
}

/** Caller text as a PowerShell single-quoted literal. */
export function psString(text: string): string {
    return `'${psEscape(text)}'`;
}

/** Strings as a PowerShell array expression: `@('a', 'b')`. */
export function psArray(values: readonly string[]): string {
    return `@(${values.map(psString).join(', ')})`;
}

/** `$true` / `$false` — PowerShell has no bare boolean literal. */
export function psBool(value: boolean): string {
    return value ? '$true' : '$false';
}

/**
 * A whole number for a script. The operations layer has already validated every
 * count, so this refusing anything else is the second line of defence: a
 * non-number never reaches a script as code.
 */
export function psInt(value: number): string {
    if (!Number.isSafeInteger(value)) {
        throw new InvalidRequestError(`Expected a whole number, got ${String(value)}.`);
    }
    return String(value);
}

// ── Running ──────────────────────────────────────────────────────────────

/**
 * The complete script for a body: UTF-8 output, stop-on-error, and a catch that
 * reports any failure as one JSON line on stderr.
 *
 * UTF-8 output matters because Windows PowerShell 5.1 otherwise writes in the
 * console's legacy code page, whose "best fit" mapping rewrites non-ASCII — a
 * curly quote in a subject becomes a straight one — which corrupts the JSON
 * the scripts print. The catch matters because an uncaught PowerShell error
 * prints a multi-line error record ("At line:12 char:48 …") that would
 * otherwise become the error's message.
 */
export function buildScript(body: string): string {
    return `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
try {
${body}
} catch {
    $failure = $_
    $message = [string]$failure.Exception.Message
    if (-not $message) { $message = [string]$failure }
    $line = 0
    try { $line = [int]$failure.InvocationInfo.ScriptLineNumber } catch {}
    [Console]::Error.WriteLine('${FAILURE_MARKER}' + (ConvertTo-Json -Compress @{ message = $message; line = $line }))
    exit 1
}
`;
}

/** The failure a script's catch block reported, or its raw stderr when it never got that far. */
function parseFailure(stderr: string): { message: string; line?: number } {
    const at = stderr.lastIndexOf(FAILURE_MARKER);
    if (at !== -1) {
        const json = stderr.slice(at + FAILURE_MARKER.length).split(/\r?\n/)[0];
        try {
            const parsed = JSON.parse(json) as { message?: unknown; line?: unknown };
            const line = typeof parsed.line === 'number' && parsed.line > 0 ? parsed.line : undefined;
            return {message: String(parsed.message ?? ''), line};
        } catch {
            // Fall through to the raw stderr.
        }
    }
    // A script that fails to PARSE never enters its try block, so PowerShell's
    // own report is all there is.
    return {message: stderr.trim()};
}

async function execute(body: string, budget: RunBudget): Promise<{
    stdout: string;
    script: string;
    durationMs: number
}> {
    requireWindows();
    const script = buildScript(body);
    // Always from a file, never `-Command`: a file has no command-line length
    // limit and no second layer of quoting. The byte-order mark is what makes
    // PowerShell 5.1 read it as UTF-8 rather than the ANSI code page, which
    // would mangle every non-ASCII character in a subject, name or path.
    const file = tempFile('script', 'ps1');
    fs.writeFileSync(file, '\uFEFF' + script, 'utf8');
    try {
        const {stdout, durationMs} = await runScriptProcess({
            runner: RUNNER,
            command: 'powershell.exe',
            args: ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', file],
            script,
            budget,
            parseFailure,
        });
        return {stdout, script, durationMs};
    } finally {
        removeTempFile(file);
    }
}

/** Run a script body and return what it printed, trimmed. */
export async function runPowerShell(body: string, budget: RunBudget): Promise<string> {
    return (await execute(body, budget)).stdout.trim();
}

/**
 * Run a script body that prints JSON, and return it parsed — null when the
 * script printed nothing. Output that isn't JSON is a SCRIPT_FAILED carrying
 * the script, never a bare SyntaxError.
 */
export async function runPowerShellJson(body: string, budget: RunBudget): Promise<unknown> {
    const {stdout, script, durationMs} = await execute(body, budget);
    const output = stdout.trim();
    if (output === '' || output === 'null') return null;
    try {
        return JSON.parse(output);
    } catch (error) {
        throw new ScriptError({
            runner: RUNNER,
            script,
            stderr: '',
            durationMs,
            message: `PowerShell printed output that is not valid JSON: ${output.slice(0, 200)}`,
            cause: error,
        });
    }
}

/**
 * Write text a script will read back, and hand back its cleanup.
 *
 * Large HTML — an email body, a Word-sized template — goes through a file rather
 * than the script itself, where it would have to be escaped and would bloat
 * every error and debug event that carries the script.
 */
export function scriptInput(kind: string, contents: string): { path: string; cleanup(): void } {
    const file = tempFile(kind, 'html');
    fs.writeFileSync(file, contents, 'utf8');
    return {path: file, cleanup: () => removeTempFile(file)};
}

/**
 * Read a file a script wrote, as UTF-8, and delete it; '' when it never
 * appeared. .NET's UTF-8 writer opens the file with a byte-order mark, which is
 * not part of the text.
 */
export function readScriptOutput(file: string): string {
    try {
        return fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '');
    } catch {
        return '';
    } finally {
        removeTempFile(file);
    }
}
