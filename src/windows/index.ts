// The Windows backend: Outlook driven through PowerShell and its COM object model.
//
// Split by feature, since every operation is "generate a script, run it, read
// its output" and the interesting part of each is the script, which reads best
// beside the rules it encodes. `run.ts` owns running and escaping; `scripts.ts`
// holds the fragments several features emit.
import { getOutlookAccounts } from './accounts';
import { replyOutlookEmail, sendOutlookEmail } from './send';
import { openOutlookEmail, readEmailBody, readInboxEmails, readSelectedEmail, searchInboxByFilter } from './read';
import { listInboxFolders, moveOutlookEmails } from './folders';
import { deleteOutlookDrafts, listOutlookDrafts, sendAllDrafts, sendDrafts } from './drafts';
import { deleteOutlookEmails, purgeDeletedItems } from './cleanup';
import { saveEmailAttachments } from './attachments';
import { cleanUndeliverableEmails, collectBouncedRecipients, readSentRecipientGroups } from './bounces';
import { listOutlookSignatures, readOutlookSignatureHtml } from './signatures';
import { editEmailTemplate, readTemplateEmails, saveTemplateEmail } from './templates';
import type { Backend } from '../backend';

export const windowsBackend: Backend = {
    getOutlookAccounts,
    sendOutlookEmail,
    replyOutlookEmail,
    readInboxEmails,
    searchInboxByFilter,
    readSelectedEmail,
    readEmailBody,
    openOutlookEmail,
    listInboxFolders,
    moveOutlookEmails,
    listOutlookDrafts,
    sendDrafts,
    sendAllDrafts,
    deleteOutlookDrafts,
    deleteOutlookEmails,
    purgeDeletedItems,
    saveEmailAttachments,
    cleanUndeliverableEmails,
    collectBouncedRecipients,
    readSentRecipientGroups,
    listOutlookSignatures,
    readOutlookSignatureHtml,
    readTemplateEmails,
    saveTemplateEmail,
    editEmailTemplate,
};
