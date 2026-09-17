// Everything a reply needs before it touches Outlook: resolving a named
// template, keeping its section, filling placeholders and a named signature,
// and reducing the result to HTML that can be inserted into another document.
// String work plus three reads each platform already provides, so it is done
// once here for both.
import { composeTemplateBody } from '../templateBody';
import { InvalidRequestError, NotFoundError } from '../errors';
import type { ReadTemplatesRequest } from '../backend';
import type { TemplateFolderResult } from '../types';

/** The reads a template- or signature-backed reply needs. */
export interface ReplyBodyDeps {
    readTemplateEmails(request: ReadTemplatesRequest): Promise<TemplateFolderResult>;

    readOutlookSignatureHtml(name: string): Promise<string>;

    listOutlookSignatures(): Promise<string[]>;
}

/** A reply's content, as the caller described it — already validated. */
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
export function innerBodyHtml(html: string): string {
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
export async function resolveTemplateHtml(
    deps: ReplyBodyDeps,
    account: string,
    folder: string,
    subject: string,
): Promise<string> {
    const result = await deps.readTemplateEmails({account, folder, limit: 50, includeBody: true, subject});
    if (!result.folderFound) {
        const folders = result.availableFolders.join(', ') || '(none)';
        throw new NotFoundError('folder', `Template folder '${folder}' not found in ${account}. Top-level folders: ${folders}.`);
    }
    const wanted = subject.trim().toLowerCase();
    const match = result.templates.find(t => t.subject.trim().toLowerCase() === wanted);
    if (!match) {
        // The read above was filtered, so it can't name the alternatives. A
        // second read without bodies can, cheaply.
        const all = await deps.readTemplateEmails({account, folder, limit: 200, includeBody: false});
        const names = all.templates.map(t => t.subject).filter(Boolean).join(', ') || '(none)';
        throw new NotFoundError('template', `Template '${subject}' not found in '${folder}'. Templates there: ${names}.`);
    }
    if (match.htmlBody.trim() === '') {
        throw new NotFoundError('template', `Template '${subject}' has an empty HTML body.`);
    }
    return match.htmlBody;
}

/**
 * The HTML to insert above the quoted original: template resolved, section
 * kept, placeholders and signature filled, document wrapper removed.
 *
 * A body taken from a template is always put through composition, even with
 * nothing to fill, so a stray [[SECTION]] marker is caught here rather than
 * mailed out.
 */
export async function composeReplyHtml(content: ReplyContent, deps: ReplyBodyDeps): Promise<string> {
    let html = content.htmlBody;
    let fromTemplate = false;
    if ((!html || html.trim() === '') && content.templateSubject) {
        html = await resolveTemplateHtml(deps, content.account, content.templateFolder, content.templateSubject);
        fromTemplate = true;
    }
    if (!html || html.trim() === '') {
        throw new InvalidRequestError('A reply needs either htmlBody or templateSubject.');
    }

    let placeholders = content.templatePlaceholders;
    if (content.signatureName) {
        const signatureHtml = await deps.readOutlookSignatureHtml(content.signatureName);
        if (signatureHtml.trim() === '') {
            const available = (await deps.listOutlookSignatures()).join(', ') || '(none)';
            throw new NotFoundError(
                'signature',
                `Outlook signature '${content.signatureName}' not found. Signatures on this machine: ${available}.`,
            );
        }
        placeholders = {...placeholders, SIGNATURE: innerBodyHtml(signatureHtml)};
    }

    const hasPlaceholders = !!placeholders && Object.keys(placeholders).length > 0;
    if (fromTemplate || content.templateSection || hasPlaceholders) {
        html = composeTemplateBody(html, {
            section: content.templateSection,
            placeholders,
            label: content.templateSubject,
        });
    }
    return innerBodyHtml(html);
}

/**
 * Where the new text goes in an Outlook-built reply: straight after the opening
 * of `WordSection1`, the div holding the empty paragraph where a person's cursor
 * would be. A reply Outlook didn't build that way falls back to just inside
 * `<body>`. The insertion itself happens in each platform's own language, since
 * the reply HTML never leaves Outlook; this anchor is the part both must agree on.
 */
export const WORD_SECTION_ANCHOR = 'WordSection1>';
