// The Windows implementation: Outlook driven through PowerShell + COM.
//
// Split by feature rather than kept as one file — every operation here is
// "generate a script, run it, coerce the result", and the interesting part of
// each is the script, which is only readable next to the rules it encodes.
// `scripts.ts` holds the fragments more than one feature emits; `run.ts` owns
// the run itself and the escaping.
//
// This is the reference implementation: everything in the platform contract
// works here, which is what the macOS capability map is measured against.
import { getOutlookAccounts } from './accounts';
import { replyOutlookEmail, sendOutlookEmail } from './send';
import { deleteOutlookDrafts, listOutlookDrafts, sendAllDrafts } from './drafts';
import { openOutlookEmail, readEmailBody, readInboxEmails, readSelectedEmail, searchInboxByFilter, } from './read';
import { listInboxFolders, moveOutlookEmails } from './folders';
import { deleteOutlookEmails, purgeDeletedItems } from './cleanup';
import { cleanUndeliverableEmails, collectBouncedRecipients, readSentRecipientGroups, } from './bounces';
import { saveEmailAttachment, saveEmailAttachmentDetailed, saveEmailAttachments, } from './attachments';
import { listOutlookSignatures, readOutlookSignatureHtml } from './signatures';
import { editEmailTemplate, readTemplateEmails, saveTemplateEmail } from './templates';
import { capabilityMap } from '../shared/capabilities';
import type { CapabilityMap, OutlookBridge } from '../types';

/**
 * The whole contract in one object — and the compile-time proof that this
 * module answers all of it.
 *
 * Typed as `OutlookBridge` rather than inferred: without that a signature could
 * drift from the macOS one and only surface as a union in the published .d.ts,
 * which is exactly how it drifted before.
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

// Re-exported one module at a time rather than by name: every feature module
// here exports exactly its bridge functions and nothing else, so the name list
// this replaces was a copy of the `bridge` literal above with no way to drift
// from it usefully.
export * from './accounts';
export * from './send';
export * from './drafts';
export * from './read';
export * from './folders';
export * from './cleanup';
export * from './bounces';
export * from './attachments';
export * from './signatures';
export * from './templates';

// The folder-scope emitter, exported so the generated script can be checked
// without an Outlook session — a syntax error in it would otherwise only ever
// surface as a failed live run.
export { mailScopeScript } from './scripts';

/**
 * Windows drives the full COM object model, so every operation is available —
 * this is the reference the macOS map is measured against.
 */
export const capabilities: CapabilityMap = capabilityMap(bridge);
