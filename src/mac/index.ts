// The macOS implementation: Outlook driven through AppleScript (osascript).
//
// Scope note: this targets **legacy** Outlook for Mac (`defaults read
// com.microsoft.Outlook IsRunningNewOutlook` = 0), which implements enough of the
// AppleScript dictionary to enumerate accounts, read mailboxes, reply, file and
// delete mail. "New Outlook" implements only a slice of it: accounts don't
// enumerate, the inbox reports 0 messages, and mutations like `delete` silently
// no-op. Under New Outlook the account lookup finds nothing and every function
// here fails loudly with "Account '…' not found" rather than quietly doing the
// wrong thing.
//
// Three dictionary traps are worth knowing, and each is handled where it bites:
//  - `move`/`delete`/`close` report success even when they do nothing, so every
//    mutation confirms with a count.
//  - A new outgoing message lands in Temporary Items — `open` only displays it,
//    so a windowless draft has to be moved into Drafts explicitly.
//  - A bad dictionary term fails the WHOLE script at compile time rather than at
//    the offending line, which is why the root terms in scripts.ts are the
//    verified ones rather than the ones the Windows ids suggest.
import { getOutlookAccounts } from './accounts';
import { replyOutlookEmail, sendOutlookEmail } from './send';
import { deleteOutlookDrafts, listOutlookDrafts, sendAllDrafts } from './drafts';
import {
    openOutlookEmail,
    readEmailBody,
    readInboxEmails,
    readSelectedEmail,
    searchInboxByFilter,
} from './read';
import { listInboxFolders, moveOutlookEmails } from './folders';
import { deleteOutlookEmails, purgeDeletedItems } from './cleanup';
import {
    cleanUndeliverableEmails,
    collectBouncedRecipients,
    readSentRecipientGroups,
} from './bounces';
import {
    saveEmailAttachment,
    saveEmailAttachmentDetailed,
    saveEmailAttachments,
} from './attachments';
import { listOutlookSignatures, readOutlookSignatureHtml } from './signatures';
import { editEmailTemplate, readTemplateEmails, saveTemplateEmail } from './templates';
import type { CapabilityMap, OutlookBridge } from '../types';

/**
 * The whole contract in one object — and the compile-time proof that this
 * module answers all of it.
 *
 * Typed as `OutlookBridge` rather than inferred: a signature that drifts from
 * the Windows one has to fail here, in the package, rather than surface as a
 * per-platform union in the consumer's editor.
 */
const bridge: OutlookBridge = {
    getOutlookAccounts,
    sendOutlookEmail,
    replyOutlookEmail,
    sendAllDrafts,
    readInboxEmails,
    searchInboxByFilter,
    readSelectedEmail,
    readEmailBody,
    openOutlookEmail,
    listInboxFolders,
    moveOutlookEmails,
    listOutlookDrafts,
    deleteOutlookDrafts,
    deleteOutlookEmails,
    purgeDeletedItems,
    saveEmailAttachment,
    saveEmailAttachmentDetailed,
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

export { bridge };
export {
    getOutlookAccounts,
    sendOutlookEmail,
    replyOutlookEmail,
    sendAllDrafts,
    readInboxEmails,
    searchInboxByFilter,
    readSelectedEmail,
    readEmailBody,
    openOutlookEmail,
    listInboxFolders,
    moveOutlookEmails,
    listOutlookDrafts,
    deleteOutlookDrafts,
    deleteOutlookEmails,
    purgeDeletedItems,
    saveEmailAttachment,
    saveEmailAttachmentDetailed,
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

/**
 * What is ported to macOS.
 *
 * The whole contract is, now — the gaps this map existed to announce (replies,
 * drafts, attachments, deletion, templates, bounce handling and the Inbox-tree
 * search) are implemented against the AppleScript dictionary rather than
 * throwing NOT_IMPLEMENTED.
 *
 * Two differences survive that no capability flag can express, and a caller
 * moving between platforms has to know them:
 *  - **Ids are not portable.** macOS message ids are small integers, Windows
 *    EntryIDs are long MAPI strings; each platform rejects the other's. `storeId`
 *    has no macOS equivalent and is accepted-and-ignored throughout.
 *  - **`searchInboxByFilter` filters here, not in Outlook.** There is no
 *    AppleScript counterpart to COM's server-side `Items.Restrict`, so a
 *    `daysBack` of 0 over a large mailbox tree costs considerably more here.
 */
export const capabilities: CapabilityMap = {
    getOutlookAccounts: true,
    sendOutlookEmail: true,
    replyOutlookEmail: true,
    sendAllDrafts: true,
    readInboxEmails: true,
    searchInboxByFilter: true,
    readSelectedEmail: true,
    readEmailBody: true,
    openOutlookEmail: true,
    listInboxFolders: true,
    moveOutlookEmails: true,
    listOutlookDrafts: true,
    deleteOutlookDrafts: true,
    deleteOutlookEmails: true,
    purgeDeletedItems: true,
    saveEmailAttachment: true,
    saveEmailAttachmentDetailed: true,
    saveEmailAttachments: true,
    cleanUndeliverableEmails: true,
    collectBouncedRecipients: true,
    readSentRecipientGroups: true,
    listOutlookSignatures: true,
    readOutlookSignatureHtml: true,
    readTemplateEmails: true,
    saveTemplateEmail: true,
    editEmailTemplate: true,
};
