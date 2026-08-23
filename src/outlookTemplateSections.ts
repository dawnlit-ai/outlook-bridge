/**
 * Section + placeholder composition for the operator's Outlook template emails.
 *
 * A template email is a normal saved mail the operator edits in Outlook, so the
 * greeting, closing and signature are theirs. Where one template serves several
 * reply types, the variants live inside it between markers:
 *
 *     Hello,
 *     [[QUOTE]]     Thank you for your quote...        [[/QUOTE]]
 *     [[NOQUOTE]]   Thank you for letting us know...   [[/NOQUOTE]]
 *     [[CLARIFY]]   {{QUESTIONS}}                      [[/CLARIFY]]
 *     Best regards, <signature>
 *
 * Composing keeps the requested section, drops the others, and substitutes any
 * {{PLACEHOLDER}} — so the shared wording is written once and the caller only
 * names a section instead of carrying (and editing) 40-50k of Word HTML.
 *
 * Everything here works on the raw HTML string rather than a parsed tree: Word's
 * markup carries conditional comments and namespaced tags that a round-trip
 * through a parser would rewrite, and the reply body must stay byte-identical to
 * what the operator wrote.
 */

/** Marker/placeholder names: what an operator can reasonably type in Outlook. */
const NAME = '[A-Za-z0-9_-]{1,40}';

/**
 * What Word may inject *between two characters of a marker*: tags (it splits runs
 * across <span>s freely), comments, non-breaking spaces, and the newlines it wraps
 * long lines with — sometimes mid-word. So `[[QUOTE]]` is matched character by
 * character with this between, never as a literal string.
 */
const GAP = '(?:<!--[\\s\\S]*?-->|<[^>]*>|&nbsp;|&#160;|\\s)*';

/** Block elements a marker can sit alone inside; the whole block goes when it does. */
const BLOCK_TAGS = ['p', 'div', 'li'];

/**
 * Block elements that make up one *line* of a template, for a placeholder whose
 * empty value takes its whole line out (see `removeTokenLine`). Ordered
 * outermost-first: Word puts a <p> inside every table cell, so a token in a
 * details table has to drop the <tr> — dropping the <p> would leave the row's
 * label cell sitting beside a blank one.
 *
 * <div> is deliberately absent: it is rarely a line and often the wrapper around
 * the entire body, so a token sitting directly in one would take the whole email
 * with it. Falling back to deleting just the token is the safer miss.
 */
const LINE_TAGS = ['tr', 'li', 'p'];

function escapeRe(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Tag-tolerant matcher for one literal marker, e.g. `[[/QUOTE]]` or `{{QUESTIONS}}`. */
function markerRegex(token: string): RegExp {
    return new RegExp(token.split('').map(escapeRe).join(GAP), 'i');
}

/** HTML → rough plain text, enough to spot markers and to test a block for emptiness. */
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

/** Find the markers in a template body (HTML or the plain-text body Outlook derives). */
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
 * Every `[[...]]` / `{{...}}` token still in a body, matched pairs or not — the
 * last check before a composed reply is handed to Outlook, so a marker the
 * operator half-deleted stops the run instead of going out to a carrier.
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
 * marker — Word doesn't nest <p>, and where the guess is wrong the emptiness
 * test below simply fails and the caller falls back to the marker's own span.
 */
function enclosingBlock(html: string, start: number, end: number, tags: string[] = BLOCK_TAGS): {
    block: Span;
    contentStart: number;
    contentEnd: number
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
 * What to delete for a marker: the whole paragraph when the marker is alone in it
 * (removing only the text would leave a blank line in the reply), otherwise just
 * the marker's own characters.
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

// ── Single-token fill ───────────────────────────────────────────────
//
// The section machinery below is for templates that hold several variants of one
// reply. These three are for the simpler case: a template whose wording is fixed
// and whose details — a lane, a cargo weight — are typed somewhere else and dropped
// into `{Token}` placeholders. Same tag-tolerant matching, since the operator
// edits these templates in Outlook too.

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
 * Remove `token` along with the line it sits on — the table row or paragraph,
 * label included. For an optional detail this is the difference between the email
 * not mentioning it at all and it carrying a heading with nothing under it.
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
 * The line to delete for a token at `at`: its enclosing row or paragraph, or the
 * token's own span when it isn't in one.
 *
 * `enclosingBlock` returns the *nearest* opening tag before the token, which is
 * only the enclosing one if it hasn't already been closed — a token in a
 * paragraph after a table has a `<tr>` before it and a `</tr>` after it without
 * ever being inside a row, and taking that span would delete the table.
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
export function findTokens(html: string, tokens: string[]): string[] {
    const text = stripToText(html || '');
    return tokens.filter(token => markerRegex(token).test(text));
}

/**
 * `{Token}`-shaped text still in a filled-in body: a placeholder misspelled in
 * Outlook, or one the app has no field for. Reads the visible text only, so the
 * braces in a Word <style> block are not mistaken for placeholders.
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
    /** Section to keep; the other sections are dropped along with their markers. */
    section?: string;
    /** `{{NAME}}` → HTML to substitute. Names are matched case-insensitively. */
    placeholders?: Record<string, string>;
    /** Template name, for error messages. */
    label?: string;
}

/**
 * Resolve a template body down to the one reply that goes out: keep `section`,
 * drop the rest, fill the placeholders. Throws rather than guessing — a template
 * the operator edited into an unusable state must stop the run, not quietly send
 * a reply with `[[QUOTE]]` or an empty question list in it.
 */
export function composeTemplateBody(html: string, options: ComposeOptions = {}): string {
    const label = options.label ? `Template '${options.label}'` : 'The template';
    const wanted = (options.section || '').trim().toUpperCase();
    const {sections, placeholders: available} = findTemplateMarkers(html);

    if (wanted && !sections.includes(wanted)) {
        throw new Error(
            sections.length
                ? `${label} has no [[${wanted}]] section. Sections found: ${sections.join(', ')}.`
                : `${label} has no [[${wanted}]]...[[/${wanted}]] markers — it holds a single body, so drop template_section (or add the markers in Outlook).`,
        );
    }
    if (!wanted && sections.length) {
        throw new Error(
            `${label} is split into sections (${sections.join(', ')}) — name one with template_section, `
            + `otherwise the reply would go out with the markers and every variant in it.`,
        );
    }

    // Deletions are collected first and applied right-to-left, so each span's
    // indices still refer to the string it was measured against.
    const cuts: Span[] = [];
    for (const name of sections) {
        const open = locate(html, `[[${name}]]`);
        const close = locate(html, `[[/${name}]]`);
        if (!open || !close || close.start < open.end) {
            throw new Error(`${label} has a malformed [[${name}]] section — expected [[${name}]] ... [[/${name}]] in that order.`);
        }
        const openSpan = markerSpan(html, open.start, open.end);
        const closeSpan = markerSpan(html, close.start, close.end);
        if (name === wanted) {
            cuts.push(openSpan, closeSpan);   // keep the content, shed the markers
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
            throw new Error(
                `${label} has no {{${key}}} placeholder`
                + (available.length ? ` (found: ${available.map(p => `{{${p}}}`).join(', ')}).` : '.')
                + ` It may have been edited out in Outlook — ask the user rather than guessing where the text goes.`,
            );
        }
        const m = markerRegex(`{{${key}}}`).exec(out);
        // Absent from `out` but present in the template = it lived in a dropped
        // section, which is the caller over-supplying, not an error.
        if (m) out = out.slice(0, m.index) + (value ?? '') + out.slice(m.index + m[0].length);
    }

    const leftover = markerTokens(out);
    if (leftover.length) {
        const hint = leftover.every(t => t.startsWith('{{'))
            ? `Pass a value for each in template_placeholders.`
            : `Check the template in Outlook — every [[SECTION]] needs its [[/SECTION]].`;
        throw new Error(
            `${label} still contains unresolved markers after composing: ${leftover.join(', ')}. ${hint} The reply was not created.`,
        );
    }
    return out;
}
