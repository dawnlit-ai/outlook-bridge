// Public API surface. Deliberately explicit (rather than `export *`) so the
// package only ever exposes Outlook-automation surface — PowerShellService
// also holds one unrelated Excel/PDF export helper that stays internal.
export {
    getOutlookAccounts,
    sendOutlookEmail,
    sendAllDrafts,
    searchInboxByFilter,
    readSelectedEmail,
    readInboxEmails,
    readEmailBody,
    openOutlookEmail,
    sendReceivedConfirmation,
    editEmailTemplate,
    saveEmailAttachment,
    saveEmailAttachmentDetailed,
    saveEmailAttachments,
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
    listOutlookSignatures,
    readOutlookSignatureHtml,
} from './OutlookService';

export type {
    InboxEmail,
    InboxSearchFilter,
    InboxSearchMatch,
    SelectedEmail,
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
} from './OutlookService';

export {
    findTemplateMarkers,
    replaceToken,
    removeTokenLine,
    findTokens,
    findUnfilledTokens,
    composeTemplateBody,
} from './outlookTemplateSections';

export type { TemplateMarkers, ComposeOptions } from './outlookTemplateSections';
