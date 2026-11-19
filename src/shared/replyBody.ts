// Everything a reply needs before it touches Outlook: resolving a named
// template, keeping its section, filling placeholders and a named signature,
// and reducing the result to HTML that can be inserted into another document.
// A new email written from a template goes the same way but keeps the whole
// document; one written as HTML uses only the signature half. String work plus
// three reads each platform already provides, so it is done once here for both.
import { composeTemplateBody, findTokens, replaceToken } from '../templateBody';
import { InvalidRequestError, NotFoundError } from '../errors';
import type { ReadTemplatesRequest } from '../backend';
import type { TemplateFolderResult } from '../types';

/** The reads a template- or signature-backed reply needs. */
export interface ReplyBodyDeps {
    readTemplateEmails(request: ReadTemplatesRequest): Promise<TemplateFolderResult>;

    readOutlookSignatureHtml(name: string): Promise<string>;

    listOutlookSignatures(): Promise<string[]>;
}

/** A reply's or a new email's content, as the caller described it — already validated. */
export interface ReplyContent {
    readonly account: string;
    readonly htmlBody?: string;
    readonly templateSubject?: string;
    readonly templateFolder: string;
    readonly templateSection?: string;
    readonly templatePlaceholders?: Readonly<Record<string, string>>;
    readonly signatureName?: string;
}

/**
 * The content of an HTML document's `<body>`, or the input when it isn't a
 * whole document. Signature files and saved templates are both whole documents,
 * and neither may nest its `<html>`/`<body>` inside another email's.
 */
function innerBodyHtml(html: string): string {
    const match = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
    return (match ? match[1] : html).trim();
}

/**
 * A saved template's HTML, found by subject.
 *
 * The read is filtered by subject in the platform script: unfiltered, EVERY
 * template's full body would cross the interpreter's stdout on every reply, and
 * a folder of Word-sized templates overruns the output limit on a call that
 * only ever needed one of them.
 */
async function resolveTemplateHtml(
    deps: ReplyBodyDeps,
    account: string,
    folder: string,
    subject: string,
): Promise<string> {
    const result = await deps.readTemplateEmails({ account, folder, limit: 50, includeBody: true, subject });
    if (!result.folderFound) {
        const folders = result.availableFolders.join(', ') || '(none)';
        throw new NotFoundError('folder', `Template folder '${folder}' not found in ${account}. Top-level folders: ${folders}.`);
    }
    const wanted = subject.trim().toLowerCase();
    const match = result.templates.find(t => t.subject.trim().toLowerCase() === wanted);
    if (!match) {
        // The read above was filtered, so it can't name the alternatives. A
        // second read without bodies can, cheaply.
        const all = await deps.readTemplateEmails({ account, folder, limit: 200, includeBody: false });
        const names = all.templates.map(t => t.subject).filter(Boolean).join(', ') || '(none)';
        throw new NotFoundError('template', `Template '${subject}' not found in '${folder}'. Templates there: ${names}.`);
    }
    if (match.htmlBody.trim() === '') {
        throw new NotFoundError('template', `Template '${subject}' has an empty HTML body.`);
    }
    return match.htmlBody;
}

/** Where a template asks for the signature. */
const SIGNATURE_TOKEN = '{{SIGNATURE}}';

/** A named signature's HTML, ready to insert into a body; NOT_FOUND naming the ones there are. */
export async function resolveSignatureHtml(
    deps: Pick<ReplyBodyDeps, 'readOutlookSignatureHtml' | 'listOutlookSignatures'>,
    name: string,
): Promise<string> {
    const signatureHtml = await deps.readOutlookSignatureHtml(name);
    if (signatureHtml.trim() === '') {
        const available = (await deps.listOutlookSignatures()).join(', ') || '(none)';
        throw new NotFoundError('signature', `Outlook signature '${name}' not found. Signatures on this machine: ${available}.`);
    }
    return innerBodyHtml(signatureHtml);
}

/** Whether a body says where its signature goes. */
function asksForSignature(html: string): boolean {
    return findTokens(html, [SIGNATURE_TOKEN]).length > 0;
}

/**
 * Sign a body: in its `{{SIGNATURE}}` placeholder when it has one, otherwise at
 * the end, where Outlook itself would put a signature. In a whole document the
 * end is just inside `</body>`, since anything after `</html>` is outside the
 * email.
 */
export function placeSignature(html: string, signature: string): string {
    if (asksForSignature(html)) return replaceToken(html, SIGNATURE_TOKEN, signature);
    const block = `<br>${signature}`;
    const close = html.toLowerCase().lastIndexOf('</body');
    return close === -1 ? html + block : html.slice(0, close) + block + html.slice(close);
}

/**
 * A body composed: template resolved, section kept, placeholders and signature
 * filled. A new email keeps the document it came in, `<head>` styles and all;
 * a reply gets only the content of its `<body>`, for inserting into another
 * document.
 *
 * A body taken from a template is always put through composition, even with
 * nothing to fill, so a stray [[SECTION]] marker is caught here rather than
 * mailed out.
 */
async function composeBody(content: ReplyContent, deps: ReplyBodyDeps, kind: 'reply' | 'email'): Promise<string> {
    let html = content.htmlBody;
    let fromTemplate = false;
    if ((!html || html.trim() === '') && content.templateSubject) {
        html = await resolveTemplateHtml(deps, content.account, content.templateFolder, content.templateSubject);
        fromTemplate = true;
    }
    if (!html || html.trim() === '') {
        // A new email is only composed when it has a template, section or
        // placeholders, so one that gets here has the last two and nothing to
        // fill them in.
        throw new InvalidRequestError(kind === 'reply'
            ? 'A reply needs either htmlBody or templateSubject.'
            : 'An email with templateSection or templatePlaceholders needs either htmlBody or templateSubject.');
    }

    // A signature with a placeholder to go in is filled with the others, so
    // one that sat in a dropped section goes with it. Without one it is added
    // below the new text once everything else is composed.
    let placeholders = content.templatePlaceholders;
    const signature = content.signatureName ? await resolveSignatureHtml(deps, content.signatureName) : '';
    const signatureInPlace = signature !== '' && asksForSignature(html);
    if (signatureInPlace) placeholders = { ...placeholders, SIGNATURE: signature };

    const hasPlaceholders = !!placeholders && Object.keys(placeholders).length > 0;
    if (fromTemplate || content.templateSection || hasPlaceholders) {
        html = composeTemplateBody(html, {
            section: content.templateSection,
            placeholders,
            label: content.templateSubject,
        });
    }
    const body = kind === 'email' ? html : innerBodyHtml(html);
    return signature !== '' && !signatureInPlace ? placeSignature(body, signature) : body;
}

/** The HTML to insert above the quoted original, its document wrapper removed. */
export function composeReplyHtml(content: ReplyContent, deps: ReplyBodyDeps): Promise<string> {
    return composeBody(content, deps, 'reply');
}

/**
 * A new email's whole HTML, from a template or a body with sections or
 * placeholders to fill. Unlike a reply's, the document is kept: no other
 * email's wraps it, and a saved template's styles live in its `<head>`.
 */
export function composeEmailHtml(content: ReplyContent, deps: ReplyBodyDeps): Promise<string> {
    return composeBody(content, deps, 'email');
}

/**
 * Where the new text goes in an Outlook-built reply: straight after the opening
 * of `WordSection1`, the div holding the empty paragraph where a person's cursor
 * would be. A reply Outlook didn't build that way falls back to just inside
 * `<body>`. The insertion itself happens in each platform's own language, since
 * the reply HTML never leaves Outlook; this anchor is the part both must agree on.
 */
export const WORD_SECTION_ANCHOR = 'WordSection1>';
