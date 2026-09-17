// Saving attachments off an email.
//
// Two runs: the first reads the email's identity and attachment names, the
// matching and file naming happen in TypeScript against shared/attachmentMatch,
// and the second run saves what matched. That keeps one definition of "does this
// name match", one wording for the error, and one file-naming rule.
import fs from 'fs';
import { asInt, asRow, asString, field, runOsaScript, splitList, summaryFields } from './run';
import {
    macMessageId,
    MessageDetail,
    messageDetailFields,
    messageDetailSnippet,
    messageLookupSnippet
} from './scripts';
import {
    attachmentNotFoundMessage,
    findAttachmentIndex,
    safeFileName,
    uniqueSavePath
} from '../shared/attachmentMatch';
import { NotFoundError } from '../errors';
import type { SaveAttachmentsRequest } from '../backend';
import type { SavedAttachment } from '../types';

/**
 * Save the named attachments of one email, in the order named. Every name is
 * matched before anything is written, so a batch with one bad name fails
 * without having saved half its files.
 */
export async function saveEmailAttachments(request: SaveAttachmentsRequest): Promise<SavedAttachment[]> {
    const id = macMessageId(request.email.entryId);
    const parts = summaryFields(await runOsaScript(`tell application "Microsoft Outlook"
${messageLookupSnippet(id)}
${messageDetailSnippet()}
    return ${asRow(messageDetailFields())}
end tell`, 'quick'));
    const email = {
        subject: field(parts, MessageDetail.subject).trim(),
        senderName: field(parts, MessageDetail.senderName),
        senderEmail: field(parts, MessageDetail.senderEmail),
        receivedTime: field(parts, MessageDetail.receivedTime),
    };
    const attachmentNames = splitList(field(parts, MessageDetail.attachmentNames));

    const picks = request.fileNames.map(fileName => {
        const index = findAttachmentIndex(attachmentNames, fileName);
        if (index === -1) {
            throw new NotFoundError('attachment', attachmentNotFoundMessage({
                fileName,
                ...email,
                present: attachmentNames.join(', ') || '(none)',
            }));
        }
        return {index, fileName: attachmentNames[index]};
    });

    // Paths are claimed as they are chosen, so two picks with one name don't collide.
    const claimed = new Set<string>();
    const saves = picks.map(pick => {
        const file = uniqueSavePath(request.destDir, safeFileName(pick.fileName), path => claimed.has(path) || fs.existsSync(path));
        claimed.add(file);
        return {...pick, path: file};
    });
    await runOsaScript(`tell application "Microsoft Outlook"
${messageLookupSnippet(id)}
${saves.map(save => `    save (item ${asInt(save.index + 1)} of attachments of theMsg) in POSIX file ${asString(save.path)}`).join('\n')}
end tell`, 'standard');
    return saves.map(save => ({path: save.path, fileName: save.fileName, ...email}));
}
