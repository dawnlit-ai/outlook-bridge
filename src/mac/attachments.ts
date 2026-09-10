// Saving attachments off an email, one at a time or in a batch.
//
// Two runs rather than one: the first reads the email's identity and attachment
// names, the matching happens in TypeScript against shared/attachmentMatch, and
// only then does the second run save what matched. That keeps one definition of
// "does this filename match" and one wording for the error, shared with Windows.
import { asEscape, asRow, field, runOsaScript, splitList, summaryFields } from './run';
import {
    macMessageId,
    MessageDetail,
    messageDetailFields,
    messageDetailSnippet,
    messageLookupSnippet
} from './scripts';
import { attachmentNotFoundMessage, findAttachmentIndex } from '../shared/attachmentMatch';
import { resolveDestDir } from '../runtime';
import { NotFoundError } from '../errors';
import path from 'path';
import type { SavedAttachment } from '../types';

interface ResolvedEmail {
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    attachmentNames: string[];
}

/** Read the email's identity and the names of the attachments it carries. */
async function resolveEmail(id: string): Promise<ResolvedEmail> {
    const script = `tell application "Microsoft Outlook"
${messageLookupSnippet(id)}
${messageDetailSnippet()}
    return ${asRow(messageDetailFields())}
end tell`;
    const parts = summaryFields(await runOsaScript(script, 30000));
    return {
        subject: field(parts, MessageDetail.subject),
        senderName: field(parts, MessageDetail.senderName),
        senderEmail: field(parts, MessageDetail.senderEmail),
        receivedTime: field(parts, MessageDetail.receivedTime),
        attachmentNames: splitList(field(parts, MessageDetail.attachmentNames)),
    };
}

/**
 * Match one requested filename to an attachment index, or throw the shared
 * "not found" error naming the email that actually resolved.
 */
function requireAttachmentIndex(email: ResolvedEmail, fileName: string): number {
    const index = findAttachmentIndex(email.attachmentNames, fileName);
    if (index === -1) {
        throw new NotFoundError('attachment', attachmentNotFoundMessage({
            fileName,
            senderEmail: email.senderEmail,
            subject: email.subject,
            receivedTime: email.receivedTime,
            present: email.attachmentNames.join(', ') || '(none)',
            idLabel: 'message id',
        }));
    }
    return index;
}

/** Save the attachments at the given 1-based indices to `outDir`. */
async function saveByIndex(
    id: string,
    picks: readonly { index: number; name: string }[],
    outDir: string,
): Promise<void> {
    const saveLines = picks.map(pick => `    save (item ${pick.index + 1} of attachments of theMsg) in POSIX file "${asEscape(path.join(outDir, pick.name))}"`);
    const script = `
tell application "Microsoft Outlook"
${messageLookupSnippet(id)}
${saveLines.join('\n')}
end tell`;
    await runOsaScript(script, 120000);
}

/**
 * Save an email attachment by message id and filename, returning the saved path
 * together with the resolved email's subject/sender so the caller can confirm the
 * file came from the email it intended.
 *
 * `destDir` is created if absent. Without one the file lands in a fresh directory
 * of its own, because attachments keep the name the sender gave them: a single
 * shared folder means two emails carrying "invoice.pdf" silently overwrite each
 * other. Pass `destDir` whenever you want the files somewhere you control.
 *
 * When the attachment isn't found, throws an error that names the email actually
 * resolved (from/subject/received) and the attachments it does carry — a message
 * id can go stale, and a bare "not found" would point at the wrong problem.
 */
export async function saveEmailAttachmentDetailed(
    entryId: string,
    fileName: string,
    storeId?: string,
    destDir?: string,
): Promise<SavedAttachment> {
    const results = await saveEmailAttachments(entryId, [fileName], storeId, destDir);
    if (results.length === 0) {
        throw new NotFoundError('attachment', 'Failed to save attachment.');
    }
    return results[0];
}

/** Save an attachment and return just the saved file path. */
export async function saveEmailAttachment(
    entryId: string,
    fileName: string,
    storeId?: string,
    destDir?: string,
): Promise<string> {
    return (await saveEmailAttachmentDetailed(entryId, fileName, storeId, destDir)).path;
}

/**
 * Save several attachments from one email in a single pair of runs — what a
 * caller working through searchInboxByFilter/readSelectedEmail results wants,
 * rather than paying an osascript spawn per attachment. Destination, matching
 * and the not-found error follow saveEmailAttachmentDetailed's rules exactly,
 * applied per name; results come back in the same order as `fileNames`.
 *
 * `storeId` is accepted for signature parity and ignored: macOS has no StoreID.
 */
export async function saveEmailAttachments(
    entryId: string,
    fileNames: string[],
    _storeId?: string,
    destDir?: string,
): Promise<SavedAttachment[]> {
    if (fileNames.length === 0) return [];
    const id = macMessageId(entryId);
    const email = await resolveEmail(id);
    // Resolve every name BEFORE saving anything, so a batch with one bad name
    // fails without having written half its files.
    const picks = fileNames.map(fileName => {
        const index = requireAttachmentIndex(email, fileName);
        return { index, name: email.attachmentNames[index] };
    });
    const outDir = resolveDestDir(destDir);
    await saveByIndex(id, picks, outDir);
    return picks.map(pick => ({
        path: path.join(outDir, pick.name),
        subject: email.subject,
        senderName: email.senderName,
        senderEmail: email.senderEmail,
        receivedTime: email.receivedTime,
    }));
}
