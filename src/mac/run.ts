// Getting a generated AppleScript to osascript and its output back.
//
// The macOS counterpart of windows/run.ts: the run itself, the escaping rules,
// and the record framing every reader here shares.
import { execFile } from 'child_process';
import fs from 'fs';
import { getConfig, reportRun, tempFile } from '../runtime';
import { classifyRunFailure } from '../errors';
import { WORD_SECTION_ANCHOR } from '../shared/replyBody';

const RUNNER = 'osascript' as const;

/**
 * Run an AppleScript via a temp file (avoids arg-length and quoting limits).
 *
 * `AS_HANDLERS` is prepended here rather than by each caller — the Windows runner
 * already owns its UTF-8 prelude the same way. It matters more on this side: a
 * script missing a handler fails at COMPILE time for the WHOLE script, with a
 * message pointing nowhere near the omission, so "remember to paste the handlers"
 * was a footgun every new script had to survive.
 *
 * `timeout` overrides the configured default for this one call; see `configure()`
 * for that and for the stdout cap.
 */
export function runOsaScript(source: string, timeout?: number): Promise<string> {
    // The full text is what osascript compiled, so it is also what the error's
    // line numbers refer to and what the debug hook and ScriptError must carry.
    const script = AS_HANDLERS + source;
    const scriptFile = tempFile('osa', 'applescript');
    fs.writeFileSync(scriptFile, script, 'utf-8');
    const { timeoutMs, maxBufferBytes, signal } = getConfig();
    const effectiveTimeout = timeout ?? timeoutMs;
    const startedAt = Date.now();
    return new Promise((resolve, reject) => {
        execFile(
            'osascript',
            [scriptFile],
            { maxBuffer: maxBufferBytes, timeout: effectiveTimeout, signal },
            (error, stdout, stderr) => {
                const durationMs = Date.now() - startedAt;
                try {
                    fs.unlinkSync(scriptFile);
                } catch { /* ignore */
                }
                if (error) {
                    // Running from a file, osascript prefixes the script path:
                    // "/tmp/outlook-bridge-osa-1.applescript:12:34: execution error:
                    // Microsoft Outlook got an error: … (-1728)". Keep just the human
                    // part — these strings reach the user in per-email failure lists.
                    const msg = (stderr || error.message)
                        .replace(/^(?:.*?:)?\d+:\d+:\s*execution error:\s*/m, '')
                        .trim();
                    const failure = classifyRunFailure({
                        runner: RUNNER,
                        script,
                        stderr: msg,
                        durationMs,
                        nodeError: error,
                        timeoutMs: effectiveTimeout,
                        signal,
                    });
                    reportRun({ runner: RUNNER, script, durationMs, error: failure.message });
                    reject(failure);
                } else {
                    reportRun({ runner: RUNNER, script, durationMs });
                    resolve(stdout.replace(/\n$/, ''));
                }
            },
        );
    });
}

/** Escape a string for embedding inside an AppleScript double-quoted literal. */
export function asEscape(s: string): string {
    return s
        .replace(/\\/g, '\\\\')
        .replace(/"/g, '\\"')
        .replace(/\r\n|\r|\n/g, '\\n');
}

/** AppleScript's boolean literals. */
export function asBool(value: boolean): string {
    return value ? 'true' : 'false';
}

// Control characters as separators: subjects and bodies contain tabs, newlines
// and commas, but never these. Every emitted field also passes through the
// `sanitize` handler below, which is what makes the framing safe rather than
// merely unlikely to break.
export const FIELD_SEP = '\u001f';
export const RECORD_SEP = '\u001e';
export const LIST_SEP = '\u001d';

/** The same three, as AppleScript expressions. */
export const AS_FIELD_SEP = '(character id 31)';
export const AS_RECORD_SEP = '(character id 30)';
export const AS_LIST_SEP = '(character id 29)';

/**
 * Emit one record: the AppleScript expressions in `fields`, separated and
 * terminated so `splitRecords`/`splitFields` can take them apart again.
 *
 * Every field goes through `sanitize`, which is what makes the framing safe
 * rather than merely unlikely to break. A field holding a LIST must be built
 * with `my sanitizeList(...)`, which cleans the elements instead — `sanitize`
 * deliberately leaves the list separator alone.
 */
export function asRow(fields: readonly string[]): string {
    return fields
        .map(f => `my sanitize(${f})`)
        .join(` & ${AS_FIELD_SEP} & `) + ` & ${AS_RECORD_SEP}`;
}

/**
 * The AppleScript handlers every script here relies on. Called with `my` from
 * inside a `tell` block.
 *
 * `isoDate` produces the zero-padded 'yyyy-MM-dd HH:mm' both platforms' readers
 * return; `sanitize` strips the framing characters; `joinList` and `replaceText`
 * are the text-item-delimiter dance nobody should write twice.
 */
export const AS_HANDLERS = `
on pad2(n)
    set s to (n as integer) as string
    if (length of s) < 2 then set s to "0" & s
    return s
end pad2

on isoDate(d)
    if d is missing value then return ""
    return (year of d as string) & "-" & pad2(month of d as integer) & "-" & pad2(day of d) & " " & pad2(hours of d) & ":" & pad2(minutes of d)
end isoDate

on joinList(lst, sep)
    set saved to AppleScript's text item delimiters
    set AppleScript's text item delimiters to sep
    set out to lst as string
    set AppleScript's text item delimiters to saved
    return out
end joinList

on replaceText(t, findWhat, replaceWith)
    set saved to AppleScript's text item delimiters
    set AppleScript's text item delimiters to findWhat
    set parts to text items of t
    set AppleScript's text item delimiters to replaceWith
    set out to parts as string
    set AppleScript's text item delimiters to saved
    return out
end replaceText

-- Strips the FIELD and RECORD separators only. The LIST separator is left
-- alone on purpose: a field may legitimately BE a list, and stripping 29 here
-- collapsed every such field into one run-on string.
on sanitize(v)
    if v is missing value then return ""
    set t to v as string
    repeat with codePoint in {30, 31}
        set marker to (character id (codePoint as integer))
        if t contains marker then set t to my replaceText(t, marker, " ")
    end repeat
    return t
end sanitize

-- A list field: each element loses all three separators, then they are joined
-- with the LIST separator that splitList() takes them apart on.
on sanitizeList(lst)
    set cleaned to {}
    repeat with item_ in lst
        set t to my sanitize(item_)
        repeat with codePoint in {29}
            set marker to (character id (codePoint as integer))
            if t contains marker then set t to my replaceText(t, marker, " ")
        end repeat
        set end of cleaned to t
    end repeat
    return my joinList(cleaned, (character id 29))
end sanitizeList

on insertAboveQuoted(c, ins)
    if c is missing value then return ins
    set c to c as string
    if (length of c) is 0 then return ins
    -- Outlook builds a reply as a document whose own empty paragraph sits inside
    -- WordSection1; inserting straight after that tag puts the new text where the
    -- user's cursor would have been. Anything else falls back to just inside <body>.
    -- 'at' is a reserved AppleScript parameter name, so the cut point cannot be
    -- called that however naturally it reads.
    set cutPoint to my tagEnd(c, "${WORD_SECTION_ANCHOR}")
    if cutPoint is 0 then set cutPoint to my tagEnd(c, "<body")
    if cutPoint is 0 then return ins & c
    if cutPoint is (length of c) then return c & ins
    return (text 1 thru cutPoint of c) & ins & (text (cutPoint + 1) thru -1 of c)
end insertAboveQuoted

-- 'rest' is an AppleScript term (rest of list), so the remainder cannot be
-- called that: assigning to it compiles and then fails at run time.
on tagEnd(c, marker)
    set ix to offset of marker in c
    if ix is 0 then return 0
    if marker ends with ">" then return ix + (length of marker) - 1
    set tailText to text ix thru -1 of c
    set gt to offset of ">" in tailText
    if gt is 0 then return 0
    return ix + gt - 1
end tagEnd
`;

/** Split a script's output into records, dropping the trailing empty one. */
export function splitRecords(raw: string): string[] {
    return raw.split(RECORD_SEP).filter(record => record.trim() !== '');
}

/** Split one record into its fields. */
export function splitFields(record: string): string[] {
    return record.split(FIELD_SEP);
}

/** Split one field that carries a list. */
export function splitList(value: string | undefined): string[] {
    return (value || '').split(LIST_SEP).filter(Boolean);
}

/** A field by position — a missing field reads as ''. */
export function field(fields: readonly string[], index: number): string {
    return fields[index] ?? '';
}

/**
 * The fields of a script's ONE summary record — the shape every mutation here
 * returns (counts first, then any per-item failures as further records).
 */
export function summaryFields(raw: string): string[] {
    return splitFields(splitRecords(raw)[0] ?? '');
}

/** A numeric field, defaulting to `fallback` when absent or unparsable. */
export function intField(fields: readonly string[], index: number, fallback = 0): number {
    const parsed = Number.parseInt(field(fields, index), 10);
    return Number.isNaN(parsed) ? fallback : parsed;
}

/**
 * A boolean field. AppleScript renders booleans as the bare words `true`/`false`,
 * so this is the counterpart of `shared/json.ts`'s coercions for the framed
 * output — the Windows side gets those free from ConvertFrom-Json.
 */
export function boolField(fields: readonly string[], index: number): boolean {
    return field(fields, index).trim() === 'true';
}
