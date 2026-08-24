// Saving attachments off an email, one at a time or in a batch.
//
// Two runs rather than one: the first reads the email's identity and attachment
// names, the matching happens in TypeScript against shared/attachmentMatch, and
// only then does the second run save what matched. That keeps one definition of
// "does this filename match" and one wording for the error, shared with Windows.
import { AS_HANDLERS, AS_LIST_SEP, asEscape, asRow, field, runOsaScript, splitFields, splitList, splitRecords } from './run';
import { macMessageId, messageLookupSnippet } from './scripts';
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
    const script = `${AS_HANDLERS}
tell application "Microsoft Outlook"
${messageLookupSnippet(id)}
    set subj to ""
    try
        set subj to (subject of theMsg) as string
    end try
    set sndName to ""
    set sndAddr to ""
    try
        set snd to sender of theMsg
        try
            set sndAddr to (address of snd) as string
        end try
        try
            set sndName to (name of snd) as string
        end try
    end try
    set recvd to ""
    try
        set recvd to my isoDate(time received of theMsg)
    on error
        try
            set recvd to my isoDate(time sent of theMsg)
        end try
    end try
    set attNames to {}
    try
        set attNames to name of every attachment of theMsg
    end try
    return ${asRow([
        'subj',
        'sndName',
        'sndAddr',
        'recvd',
        'my sanitizeList(attNames)',
    ])}
end tell`;
    const parts = splitFields(splitRecords(await runOsaScript(script, 30000))[0] || '');
    return {
        subject: field(parts, 0),
        senderName: field(parts, 1),
        senderEmail: field(parts, 2),
        receivedTime: field(parts, 3),
        attachmentNames: splitList(field(parts, 4)),
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
    _storeId?: string,
    destDir?: string,
): Promise<SavedAttachment> {
    const results = await saveEmailAttachments(entryId, [fileName], _storeId, destDir);
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
        return {index, name: email.attachmentNames[index]};
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
