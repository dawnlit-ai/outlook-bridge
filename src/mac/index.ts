// The macOS backend: Outlook driven through AppleScript (osascript).
//
// This targets LEGACY Outlook for Mac, which implements enough of the AppleScript
// dictionary to enumerate accounts, read mailboxes, reply, file and delete mail.
// New Outlook implements only a slice of it — accounts don't enumerate and the
// inbox reports no messages — so there the account lookup finds nothing and
// every operation fails loudly as ACCOUNT_NOT_FOUND rather than quietly doing
// the wrong thing.
//
// Three dictionary traps are handled where they bite:
//  - `move`/`delete`/`close` report success even when they do nothing, so every
//    mutation confirms with a count.
//  - A new outgoing message lands in Temporary Items; `open` only displays it,
//    so a windowless draft is moved into Drafts explicitly.
//  - One bad dictionary term fails the WHOLE script at compile time rather than
//    at the offending line, which is why the root terms in scripts.ts are the
//    verified ones.
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

export const macBackend: Backend = {
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
