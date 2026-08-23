// Platform dispatcher for Outlook automation: Windows drives Outlook via
// PowerShell + COM, macOS via AppleScript (see each service for capability
// notes — New Outlook for Mac supports composing but not mailbox reading).
// Other platforms fall through to PowerShellService's not-supported behavior.
//
// Both implementations are pinned to `OutlookBridge` here rather than being
// destructured straight off the module namespaces. That is the difference
// between a consumer seeing `Promise<CleanUndeliverableResult>` and seeing
// `Promise<CleanUndeliverableResult | { matched: unknown[] }>` — a union of
// whatever the two platforms happened to declare, which is what this used to
// publish.
import * as PowerShellService from './PowerShellService';
import * as OutlookMacService from './OutlookMacService';
import type { OutlookBridge } from './types';

const windows: OutlookBridge = PowerShellService;
const mac: OutlookBridge = OutlookMacService;

const impl: OutlookBridge = process.platform === 'darwin' ? mac : windows;

export const {
    getOutlookAccounts,
    sendOutlookEmail,
    replyOutlookEmail,
    sendAllDrafts,
    sendReceivedConfirmation,
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
} = impl;

// Template editing drives a real Outlook compose window (COM inspector), which
// only Windows supports — macOS gets no fallback here since it isn't Outlook
// automation; callers needing an editor on macOS supply their own. Off Windows
// it throws rather than being absent, so it stays out of OutlookBridge.
export { editEmailTemplate } from './PowerShellService';
