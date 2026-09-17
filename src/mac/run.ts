// Getting a generated AppleScript to osascript and its output back.
//
// The macOS counterpart of windows/run.ts: running a script, the escaping
// rules, and the record framing every reader here shares.
import fs from 'fs';
import { removeTempFile, type RunBudget, tempFile } from '../runtime';
import { runScriptProcess } from '../shared/exec';
import { InvalidRequestError } from '../errors';
import { WORD_SECTION_ANCHOR } from '../shared/replyBody';

const RUNNER = 'osascript' as const;

/**
 * osascript's report of a script error: "<file>:<start>:<end>: execution error:
 * <message> (<number>)". Only the message is worth showing — these reach people
 * in per-item failure lists — and -2700 is the number AppleScript gives every
 * `error "…"` a script raises itself.
 */
function parseFailure(stderr: string): { message: string } {
    const message = stderr
        .replace(/^(?:.*?:)?\d+:\d+:\s*execution error:\s*/m, '')
        .replace(/\s*\(-2700\)\s*$/, '')
        .trim();
    return {message};
}

/**
 * Run a script from a temp file (no argument-length or quoting limits) and
 * return what it printed, minus the newline osascript appends.
 *
 * `AS_HANDLERS` is prepended here rather than by each caller: a script missing
 * a handler fails to compile as a WHOLE, with an error pointing nowhere near the
 * omission, so it can't be left to each script to remember.
 */
export async function runOsaScript(source: string, budget: RunBudget): Promise<string> {
    // The full text is what osascript compiles, so it is also what error offsets
    // refer to and what the debug hook and errors carry.
    const script = AS_HANDLERS + source;
    const file = tempFile('osa', 'applescript');
    fs.writeFileSync(file, script, 'utf8');
    try {
        const {stdout} = await runScriptProcess({
            runner: RUNNER,
            command: 'osascript',
            args: [file],
            script,
            budget,
            parseFailure,
        });
        return stdout.replace(/\n$/, '');
    } finally {
        removeTempFile(file);
    }
}

/**
 * Escape text for the inside of an AppleScript double-quoted literal.
 * AppleScript strings never interpolate, so escaping the quote, the backslash
 * and line breaks is the whole job.
 */
export function asEscape(text: string): string {
    return text
        .replace(/\\/g, '\\\\')
        .replace(/"/g, '\\"')
        .replace(/\r\n|\r|\n/g, '\\n');
}

/** Text as an AppleScript string literal. */
export function asString(text: string): string {
    return `"${asEscape(text)}"`;
}

/** AppleScript's boolean literals. */
export function asBool(value: boolean): string {
    return value ? 'true' : 'false';
}

/** A whole number for a script; anything else is refused before it becomes code. */
export function asInt(value: number): string {
    if (!Number.isSafeInteger(value)) {
        throw new InvalidRequestError(`Expected a whole number, got ${String(value)}.`);
    }
    return String(value);
}

/** Message ids (already validated as integers) as an AppleScript list literal. */
export function asIdList(ids: readonly string[]): string {
    return `{${ids.map(id => asInt(Number(id))).join(', ')}}`;
}

// ── Record framing ───────────────────────────────────────────────────────
//
// Control characters separate fields, records and list items: subjects and
// bodies contain tabs, newlines and commas, but never these — and every emitted
// field passes through the `sanitize` handler, which makes the framing safe
// rather than merely unlikely to break.
export const FIELD_SEP = '';
export const RECORD_SEP = '';
export const LIST_SEP = '';

/** The same three, as AppleScript expressions. */
export const AS_FIELD_SEP = '(character id 31)';
export const AS_RECORD_SEP = '(character id 30)';
export const AS_LIST_SEP = '(character id 29)';

/**
 * One record: the AppleScript expressions in `fields`, separated and terminated
 * so `splitRecords`/`splitFields` can take them apart again. A field holding a
 * LIST must be built with `my sanitizeList(...)`, which cleans the elements —
 * `sanitize` deliberately leaves the list separator alone.
 */
export function asRow(fields: readonly string[]): string {
    return fields
        .map(f => `my sanitize(${f})`)
        .join(` & ${AS_FIELD_SEP} & `) + ` & ${AS_RECORD_SEP}`;
}

/**
 * The handlers every script relies on, called with `my` from inside a `tell`.
 *
 * `isoDate` produces the zero-padded 'yyyy-MM-dd HH:mm' both platforms return;
 * `sanitize` strips the framing characters; `joinList` and `replaceText` are the
 * text-item-delimiter dance nobody should write twice.
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
-- alone on purpose: a field may legitimately BE a list.
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
        set marker to (character id 29)
        if t contains marker then set t to my replaceText(t, marker, " ")
        set end of cleaned to t
    end repeat
    return my joinList(cleaned, (character id 29))
end sanitizeList

on insertAboveQuoted(c, ins)
    if c is missing value then return ins
    set c to c as string
    if (length of c) is 0 then return ins
    -- A reply's own empty paragraph sits inside WordSection1; inserting straight
    -- after that tag puts the new text where the cursor would be. Anything else
    -- falls back to just inside <body>. ('at' is a reserved parameter name, so
    -- the cut point can't be called that.)
    set cutPoint to my tagEnd(c, "${WORD_SECTION_ANCHOR}")
    if cutPoint is 0 then set cutPoint to my tagEnd(c, "<body")
    if cutPoint is 0 then return ins & c
    if cutPoint is (length of c) then return c & ins
    return (text 1 thru cutPoint of c) & ins & (text (cutPoint + 1) thru -1 of c)
end insertAboveQuoted

-- 'rest' is an AppleScript term, so the remainder can't be called that:
-- assigning to it compiles and then fails at run time.
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

/** A script's output as records, without the trailing empty one. */
export function splitRecords(raw: string): string[] {
    return raw.split(RECORD_SEP).filter(record => record.trim() !== '');
}

/** One record's fields. */
export function splitFields(record: string): string[] {
    return record.split(FIELD_SEP);
}

/** A field that carries a list. */
export function splitList(value: string | undefined): string[] {
    return (value || '').split(LIST_SEP).filter(Boolean);
}

/** A field by position; a missing field reads as ''. */
export function field(fields: readonly string[], index: number): string {
    return fields[index] ?? '';
}

/** The fields of a script's first record — the summary a mutation returns ahead of any per-item rows. */
export function summaryFields(raw: string): string[] {
    return splitFields(splitRecords(raw)[0] ?? '');
}

/** A numeric field, `fallback` when absent or unparsable. */
export function intField(fields: readonly string[], index: number, fallback = 0): number {
    const parsed = Number.parseInt(field(fields, index), 10);
    return Number.isNaN(parsed) ? fallback : parsed;
}

/** A boolean field; AppleScript renders booleans as the bare words true/false. */
export function boolField(fields: readonly string[], index: number): boolean {
    return field(fields, index).trim() === 'true';
}
