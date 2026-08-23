// Platform dispatcher for Outlook automation: Windows drives Outlook/Excel via
// PowerShell + COM, macOS via AppleScript (see each service for capability
// notes — New Outlook for Mac supports composing but not mailbox reading).
// Other platforms fall through to PowerShellService's not-supported behavior.
import * as PowerShellService from './PowerShellService';
import * as OutlookMacService from './OutlookMacService';

const impl = process.platform === 'darwin' ? OutlookMacService : PowerShellService;

export const {
    getOutlookAccounts,
    sendOutlookEmail,
    sendAllDrafts,
    readAllPreAlerts,
    readSelectedPreAlert,
    readInboxEmails,
    readEmailBody,
    openOutlookEmail,
    sendReceivedConfirmation,
    editEmailTemplate,
    exportSheetAsPdf,
    saveEmailAttachment,
    saveEmailAttachmentDetailed,
    cleanUndeliverableEmails,
    collectBouncedRecipients,
    readSentRecipientGroups,
    readTemplateEmails,
    saveTemplateEmail,
    replyOutlookEmail,
    listInboxFolders,
    moveOutlookEmails,
    listOutlookDrafts,
    deleteOutlookDrafts,
    deleteOutlookEmails,
    purgeDeletedItems,
} = impl;

// Windows reads signatures synchronously from disk, macOS asynchronously from
// Outlook itself — expose both behind an async signature.
export async function listOutlookSignatures(): Promise<string[]> {
    return impl.listOutlookSignatures();
}

export async function readOutlookSignatureHtml(name: string): Promise<string> {
    return impl.readOutlookSignatureHtml(name);
}

export type {
    InboxEmail,
    EmailBodyResult,
    SavedAttachment,
    SendAllDraftsResult,
    UndeliverableEmail,
    CleanUndeliverableResult,
    SentRecipientGroup,
    TemplateEmail,
    TemplateFolderResult,
    ReplyEmailParams,
    ReplyEmailResult,
    InboxFolderInfo,
    MoveEmailsResult,
    OutlookDraft,
    ListDraftsResult,
    DeleteDraftsResult,
    DeleteMailOutcome,
    DeleteMailResult,
    PurgeDeletedItemsResult,
    MailFolderRef,
} from './PowerShellService';
