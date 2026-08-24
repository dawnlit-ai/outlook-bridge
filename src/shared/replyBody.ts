// Everything replyOutlookEmail does BEFORE it touches Outlook.
//
// Resolving a named template, filling {{PLACEHOLDER}}s, substituting a named
// signature and reducing the result to insertable HTML is identical work on both
// platforms — it is string handling plus two reads that each platform already
// implements. Keeping it here is what lets macOS answer the same
// ReplyEmailParams contract Windows does instead of a subset of it.
import { composeTemplateBody } from '../outlookTemplateSections';
import { InvalidRequestError, NotFoundError } from '../errors';
import type { ReplyEmailParams, TemplateFolderResult } from '../types';

/** The reads a platform must supply for a template- or signature-backed reply. */
export interface ReplyBodyDeps {
    readTemplateEmails(
        emailAccount: string,
        folderName?: string,
        limit?: number,
        includeBody?: boolean,
        subject?: string,
    ): Promise<TemplateFolderResult>;

    readOutlookSignatureHtml(name: string): Promise<string>;

    listOutlookSignatures(): Promise<string[]>;
}

/**
 * The content of an HTML document's `<body>`, or the input when there is no
 * document around it.
 *
 * Both things this is used on — a signature file and a saved template — are whole
 * HTML documents, and neither may nest its `<html>`/`<body>` inside the reply's.
 */
export function innerBodyHtml(html: string): string {
    const match = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
    return (match ? match[1] : html).trim();
}

/**
 * Resolve a saved template's HTML by subject, so a reply can name a template
 * instead of carrying its (often large, Word-generated) HTML.
 *
 * The read is filtered by subject in the platform layer rather than here:
 * unfiltered, EVERY template's full body crosses the interpreter's stdout on
 * every reply, and a folder of Word-sized templates overruns the buffer outright
 * — turning a growing Templates folder into a failure on a call that never
 * needed more than one body.
 */
export async function resolveTemplateHtmlBySubject(
    deps: ReplyBodyDeps,
    emailAccount: string,
    folderName: string,
    subject: string,
): Promise<string> {
    const result = await deps.readTemplateEmails(emailAccount, folderName, 50, true, subject);
    if (!result.folderFound) {
        const folders = result.availableFolders.join(', ') || '(none)';
        throw new NotFoundError(
            'folder',
            `Template folder '${folderName}' not found in ${emailAccount}. Available folders: ${folders}.`,
        );
    }
    const want = subject.trim().toLowerCase();
    const match = result.templates.find(t => t.subject.trim().toLowerCase() === want);
    if (!match) {
        // The read above was filtered, so it can't name the alternatives. List them
        // in a second pass with bodies omitted — which is what makes reading the
        // whole folder safe here, and keeps the error as useful as it was before
        // the filter.
        const all = await deps.readTemplateEmails(emailAccount, folderName, 50, false);
        const names = all.templates.map(t => t.subject).filter(Boolean).join(', ') || '(none)';
        throw new NotFoundError(
            'template',
            `Template '${subject}' not found in '${folderName}'. Available templates: ${names}.`,
        );
    }
    if (!match.htmlBody || match.htmlBody.trim() === '') {
        throw new NotFoundError('template', `Template '${subject}' has an empty HTML body.`);
    }
    return match.htmlBody;
}

/**
 * The HTML to insert above the quoted original: template resolved, section
 * chosen, placeholders and signature filled, document wrapper removed.
 *
 * Composition also runs (with no options) on any body taken straight from a
 * template, so a stray [[SECTION]] marker is caught here rather than mailed to a
 * customer.
 */
export async function composeReplyHtml(
    params: ReplyEmailParams,
    deps: ReplyBodyDeps,
): Promise<string> {
    let htmlBody = params.htmlBody;
    let fromTemplate = false;
    if ((!htmlBody || htmlBody.trim() === '') && params.templateSubject) {
        htmlBody = await resolveTemplateHtmlBySubject(
            deps,
            params.emailAccount,
            params.templateFolder || 'Templates',
            params.templateSubject,
        );
        fromTemplate = true;
    }
    if (!htmlBody || htmlBody.trim() === '') {
        throw new InvalidRequestError(
            'replyOutlookEmail needs either htmlBody or a resolvable templateSubject.',
        );
    }

    // A named signature fills {{SIGNATURE}} like any other placeholder, but is
    // read here so the caller never carries the signature HTML — images and all —
    // across the wire.
    let placeholders = params.templatePlaceholders;
    if (params.signatureName) {
        const signatureHtml = await deps.readOutlookSignatureHtml(params.signatureName);
        if (signatureHtml.trim() === '') {
            const available = (await deps.listOutlookSignatures()).join(', ') || '(none)';
            throw new NotFoundError(
                'signature',
                `Outlook signature '${params.signatureName}' not found. Available signatures: ${available}.`,
            );
        }
        placeholders = {...placeholders, SIGNATURE: innerBodyHtml(signatureHtml)};
    }

    const hasPlaceholders = !!placeholders && Object.keys(placeholders).length > 0;
    if (fromTemplate || params.templateSection || hasPlaceholders) {
        htmlBody = composeTemplateBody(htmlBody, {
            section: params.templateSection,
            placeholders,
            label: params.templateSubject,
        });
    }
    return innerBodyHtml(htmlBody);
}

/**
 * Put `insertHtml` at the top of a reply's HTML, above the quoted thread.
 *
 * Outlook builds the reply body as a Word document whose own (empty) paragraph
 * sits inside `WordSection1`; inserting straight after that div's opening tag is
 * what puts the new text where the user's cursor would have been. A reply
 * Outlook did not build that way falls back to just inside `<body>`, and a
 * fragment with neither gets the text prepended.
 */
export function insertAboveQuoted(replyHtml: string, insertHtml: string): string {
    const wordSection = replyHtml.indexOf('WordSection1>');
    if (wordSection >= 0) {
        const at = wordSection + 'WordSection1>'.length;
        return replyHtml.slice(0, at) + insertHtml + replyHtml.slice(at);
    }
    const body = /<body[^>]*>/i.exec(replyHtml);
    if (body) {
        const at = body.index + body[0].length;
        return replyHtml.slice(0, at) + insertHtml + replyHtml.slice(at);
    }
    return insertHtml + replyHtml;
}
