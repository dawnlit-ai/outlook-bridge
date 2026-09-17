/**
 * Sections and placeholders in HTML email templates.
 *
 * A template email is an ordinary saved mail its owner edits in Outlook, so the
 * greeting, closing and signature stay theirs. Where one template serves several
 * kinds of reply, the variants live inside it between markers:
 *
 *     Hello,
 *     [[ACCEPT]]    Thank you — we accept...            [[/ACCEPT]]
 *     [[DECLINE]]   Thank you for letting us know...    [[/DECLINE]]
 *     [[CLARIFY]]   {{QUESTIONS}}                       [[/CLARIFY]]
 *     Best regards, {{SIGNATURE}}
 *
 * Composing keeps the requested section, drops the others, and substitutes the
 * `{{PLACEHOLDER}}`s — so shared wording is written once and a caller names a
 * section instead of carrying the whole body.
 *
 * Everything works on the raw HTML string rather than a parsed tree: Word's
 * markup carries conditional comments and namespaced tags a parser round-trip
 * would rewrite, and the composed body must stay byte-identical to what was
 * written.
 */

import { InvalidRequestError } from './errors';

/** Marker and placeholder names: what a person can reasonably type in Outlook. */
const NAME = '[A-Za-z0-9_-]{1,40}';

/**
 * What Word may inject *between two characters of a marker*: tags (it splits
 * runs across <span>s freely), comments, non-breaking spaces, and the newlines
 * it wraps long lines with — sometimes mid-word. So `[[ACCEPT]]` is matched
 * character by character with this between, never as a literal string.
 */
const GAP = '(?:<!--[\\s\\S]*?-->|<[^>]*>|&nbsp;|&#160;|\\s)*';

/** Block elements a marker can sit alone inside; the whole block goes when it does. */
const BLOCK_TAGS = ['p', 'div', 'li'];

/**
 * Block elements that make up one *line* of a template, for a token whose
 * empty value takes its whole line out (see `removeTokenLine`). Outermost
 * first: Word puts a <p> inside every table cell, so a token in a details table
 * has to drop the <tr> — dropping the <p> would leave the row's label beside a
 * blank cell.
 *
 * <div> is deliberately absent: it is rarely a line and often the wrapper around
 * the entire body, so a token directly inside one would take the whole email
 * with it. Deleting just the token is the safer miss.
 */
const LINE_TAGS = ['tr', 'li', 'p'];

function escapeRe(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Tag-tolerant matcher for one literal token, e.g. `[[/ACCEPT]]` or `{{QUESTIONS}}`.
 *
 * Memoized: the pattern interleaves `GAP` between every character, so a short
 * token compiles to hundreds of characters, and the replace loops ask for the
 * same token repeatedly while walking a template.
 */
const markerRegexCache = new Map<string, RegExp>();

function markerRegex(token: string): RegExp {
    let regex = markerRegexCache.get(token);
    if (!regex) {
        regex = new RegExp(token.split('').map(escapeRe).join(GAP), 'i');
        markerRegexCache.set(token, regex);
    }
    return regex;
}

/** HTML → rough plain text, enough to spot markers and test a block for emptiness. */
function stripToText(html: string): string {
    return html
        .replace(/<!--[\s\S]*?-->/g, '')
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;|&#160;/gi, ' ');
}

export interface TemplateMarkers {
    /** Section names with both an opening and a closing marker, in document order. */
    sections: string[];
    /** `{{NAME}}` placeholders found anywhere in the template. */
    placeholders: string[];
}

/** Find the markers in a template body (HTML, or the plain text Outlook derives from it). */
export function findTemplateMarkers(body: string): TemplateMarkers {
    const text = stripToText(body || '');
    const opened: string[] = [];
    const closed = new Set<string>();
    const seen = new Set<string>();
    const sectionRe = new RegExp(`\\[\\[\\s*(/?)\\s*(${NAME})\\s*\\]\\]`, 'g');
    let m: RegExpExecArray | null;
    while ((m = sectionRe.exec(text)) !== null) {
        const name = m[2].toUpperCase();
        if (m[1]) {
            closed.add(name);
        } else if (!seen.has(name)) {
            seen.add(name);
            opened.push(name);
        }
    }
    const placeholders: string[] = [];
    const placeholderRe = new RegExp(`\\{\\{\\s*(${NAME})\\s*\\}\\}`, 'g');
    while ((m = placeholderRe.exec(text)) !== null) {
        const name = m[1].toUpperCase();
        if (!placeholders.includes(name)) placeholders.push(name);
    }
    return {sections: opened.filter(n => closed.has(n)), placeholders};
}

/**
 * Every `[[...]]` / `{{...}}` token still in a body, paired or not — the last
 * check before a composed body goes out, so a half-deleted marker stops the
 * send instead of reaching a recipient.
 */
function markerTokens(body: string): string[] {
    const text = stripToText(body || '');
    const re = new RegExp(`\\[\\[\\s*/?\\s*${NAME}\\s*\\]\\]|\\{\\{\\s*${NAME}\\s*\\}\\}`, 'g');
    const found: string[] = [];
    let m: RegExpExecArray | null;
    while ((m = re.exec(text)) !== null) {
        const token = m[0].replace(/\s+/g, '');
        if (!found.includes(token)) found.push(token);
    }
    return found;
}

interface Span {
    start: number;
    end: number;
}

/**
 * The nearest enclosing block element of [start, end), or null when the marker
 * isn't inside one. The closing tag is the first of the same name after the
 * marker — Word doesn't nest <p>, and where that guess is wrong the emptiness
 * test simply fails and the caller falls back to the marker's own span.
 */
function enclosingBlock(html: string, start: number, end: number, tags: string[] = BLOCK_TAGS): {
    block: Span;
    contentStart: number;
    contentEnd: number;
} | null {
    let best: { idx: number; contentStart: number; tag: string } | null = null;
    for (const tag of tags) {
        const openRe = new RegExp(`<${tag}(?:\\s[^>]*)?>`, 'gi');
        let m: RegExpExecArray | null;
        while ((m = openRe.exec(html)) !== null) {
            if (m.index >= start) break;
            if (!best || m.index > best.idx) {
                best = {idx: m.index, contentStart: m.index + m[0].length, tag};
            }
        }
    }
    if (!best) return null;
    const closeRe = new RegExp(`</${best.tag}\\s*>`, 'gi');
    closeRe.lastIndex = end;
    const close = closeRe.exec(html);
    if (!close) return null;
    return {
        block: {start: best.idx, end: close.index + close[0].length},
        contentStart: best.contentStart,
        contentEnd: close.index,
    };
}

/**
 * What to delete for a marker: its whole paragraph when the marker is alone in
 * it (removing only the text would leave a blank line), else just the marker.
 */
function markerSpan(html: string, start: number, end: number): Span {
    const found = enclosingBlock(html, start, end);
    if (!found || found.contentStart > start || found.contentEnd < end) return {start, end};
    const rest = html.slice(found.contentStart, start) + html.slice(end, found.contentEnd);
    return stripToText(rest).trim() === '' ? found.block : {start, end};
}

function locate(html: string, token: string): Span | null {
    const m = markerRegex(token).exec(html);
    return m ? {start: m.index, end: m.index + m[0].length} : null;
}

// ── Single-token fill ───────────────────────────────────────────────────
//
// The section machinery above is for templates holding several variants of
// one email. These are for the simpler case: fixed wording with details typed
// elsewhere and dropped into tokens of the caller's own choosing (`{Name}`,
// `%NAME%`, …). Same tag-tolerant matching, since these templates are edited in
// Outlook too.

/** Replace every occurrence of `token` with `value`, which is never rescanned. */
export function replaceToken(html: string, token: string, value: string): string {
    let done = '';
    let rest = html;
    for (; ;) {
        const at = locate(rest, token);
        if (!at) return done + rest;
        done += rest.slice(0, at.start) + value;
        rest = rest.slice(at.end);
    }
}

/**
 * Remove `token` together with the line it sits on — the table row or
 * paragraph, label included. For an optional detail this is the difference
 * between the email not mentioning it and a heading with nothing under it.
 * Falls back to deleting just the token when it isn't inside a line element.
 */
export function removeTokenLine(html: string, token: string): string {
    let out = html;
    for (; ;) {
        const at = locate(out, token);
        if (!at) return out;
        const cut = lineSpan(out, at);
        out = out.slice(0, cut.start) + out.slice(cut.end);
    }
}

/**
 * The line to delete for a token at `at`: its enclosing row or paragraph, or
 * the token's own span when it isn't in one.
 *
 * `enclosingBlock` returns the *nearest* opening tag before the token, which is
 * only the enclosing one if it hasn't been closed already — a token in a
 * paragraph after a table has a `<tr>` before it and a `</tr>` after it without
 * being inside a row, and taking that span would delete the table.
 */
function lineSpan(html: string, at: Span): Span {
    for (const tag of LINE_TAGS) {
        const found = enclosingBlock(html, at.start, at.end, [tag]);
        if (!found) continue;
        const closedBefore = new RegExp(`</${tag}\\s*>`, 'i').test(html.slice(found.contentStart, at.start));
        if (!closedBefore) return found.block;
    }
    return at;
}

/** Which of `tokens` the template actually uses — what the caller has to supply. */
export function findTokens(html: string, tokens: readonly string[]): string[] {
    const text = stripToText(html || '');
    return tokens.filter(token => markerRegex(token).test(text));
}

/**
 * `{Token}`-shaped text still in a filled-in body: a placeholder misspelled
 * while editing, or one nothing supplied. Reads visible text only, so the
 * braces of a <style> block are not mistaken for placeholders. Catches a
 * leftover `{{TOKEN}}` too, reported by its inner `{TOKEN}`.
 */
export function findUnfilledTokens(html: string): string[] {
    const text = stripToText((html || '').replace(/<(style|script)\b[^>]*>[\s\S]*?<\/\1>/gi, ''));
    const found: string[] = [];
    const re = /\{\s*([A-Za-z][A-Za-z0-9 _-]{0,30})\s*\}/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text)) !== null) {
        const token = `{${m[1].trim()}}`;
        if (!found.includes(token)) found.push(token);
    }
    return found;
}

export interface ComposeOptions {
    /** Section to keep; the other sections are dropped with their markers. */
    section?: string;
    /** `{{NAME}}` → HTML to substitute. Names match case-insensitively. */
    placeholders?: Readonly<Record<string, string>>;
    /** The template's name, for error messages. */
    label?: string;
}

/**
 * Reduce a template body to the one email that goes out: keep `section`, drop
 * the rest, fill the placeholders. Throws INVALID_REQUEST rather than guessing —
 * a template edited into an unusable state has to stop the send, not deliver a
 * body with `[[ACCEPT]]` or an unfilled placeholder in it.
 */
export function composeTemplateBody(html: string, options: ComposeOptions = {}): string {
    const label = options.label ? `Template '${options.label}'` : 'The template';
    const wanted = (options.section || '').trim().toUpperCase();
    const {sections, placeholders: available} = findTemplateMarkers(html);

    if (wanted && !sections.includes(wanted)) {
        throw new InvalidRequestError(
            sections.length
                ? `${label} has no [[${wanted}]] section. Sections found: ${sections.join(', ')}.`
                : `${label} has no [[${wanted}]]...[[/${wanted}]] markers — it holds a single body, so name no section (or add the markers to the template).`,
        );
    }
    if (!wanted && sections.length) {
        throw new InvalidRequestError(
            `${label} is split into sections (${sections.join(', ')}); name the one to send, `
            + 'otherwise the email would go out with every variant and the markers in it.',
        );
    }

    // Deletions are collected first and applied right to left, so each span's
    // offsets still refer to the string they were measured against.
    const cuts: Span[] = [];
    for (const name of sections) {
        const open = locate(html, `[[${name}]]`);
        const close = locate(html, `[[/${name}]]`);
        if (!open || !close || close.start < open.end) {
            throw new InvalidRequestError(`${label} has a malformed [[${name}]] section — expected [[${name}]] ... [[/${name}]] in that order.`);
        }
        const openSpan = markerSpan(html, open.start, open.end);
        const closeSpan = markerSpan(html, close.start, close.end);
        if (name === wanted) {
            cuts.push(openSpan, closeSpan); // keep the content, shed the markers
        } else {
            cuts.push({start: openSpan.start, end: closeSpan.end});
        }
    }
    let out = html;
    for (const cut of cuts.sort((a, b) => b.start - a.start)) {
        out = out.slice(0, cut.start) + out.slice(cut.end);
    }

    for (const [rawKey, value] of Object.entries(options.placeholders || {})) {
        const key = rawKey.trim().toUpperCase();
        if (!available.includes(key)) {
            throw new InvalidRequestError(
                `${label} has no {{${key}}} placeholder`
                + (available.length ? ` (found: ${available.map(p => `{{${p}}}`).join(', ')}).` : '.')
                + ' It may have been removed when the template was last edited.',
            );
        }
        const m = markerRegex(`{{${key}}}`).exec(out);
        // In the template but gone from `out` means it lived in a dropped
        // section: the caller supplying more than one variant needs, not an error.
        if (m) out = out.slice(0, m.index) + (value ?? '') + out.slice(m.index + m[0].length);
    }

    const leftover = markerTokens(out);
    if (leftover.length) {
        const hint = leftover.every(t => t.startsWith('{{'))
            ? 'Supply a value for each placeholder.'
            : 'Check the template — every [[SECTION]] needs its [[/SECTION]].';
        throw new InvalidRequestError(
            `${label} still contains unresolved markers after composing: ${leftover.join(', ')}. ${hint} Nothing was created.`,
        );
    }
    return out;
}
