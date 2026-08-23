// Public API surface. Deliberately explicit (rather than `export *`) so the
// package only ever exposes what it means to support; `package.json`'s `exports`
// map closes the deep-import path that would otherwise reach around this.
export {
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
    editEmailTemplate,
} from './OutlookService';

// Process-wide settings: run timeouts, the stdout cap, where scratch files and
// saved attachments go, and a hook that hands you every generated script.
export { configure, getConfig } from './runtime';
export type { BridgeOptions, BridgeDebugEvent } from './runtime';

// Platform-neutral helpers, usable without an Outlook session — and the two
// pieces of parsing a caller most often needs to reproduce.
export { mailFolderRef, splitQuotedOriginal, WELL_KNOWN_FOLDERS } from './mail';

export type {
    OutlookBridge,
    SendEmailParams,
    ReplyEmailParams,
    ReplyEmailResult,
    DraftSendFailure,
    SendAllDraftsResult,
    OutlookDraft,
    ListDraftsResult,
    DeleteDraftsResult,
    DeleteMailOptions,
    DeleteMailOutcome,
    DeleteMailResult,
    PurgeDeletedItemsResult,
    UndeliverableEmail,
    CleanUndeliverableResult,
    SentRecipientGroup,
    InboxEmail,
    InboxSearchFilter,
    InboxSearchMatch,
    SelectedEmail,
    EmailBodyResult,
    MailFolderRef,
    InboxFolderInfo,
    MoveEmailsResult,
    SavedAttachment,
    TemplateEmail,
    TemplateFolderResult,
    SaveTemplateResult,
} from './types';

export {
    findTemplateMarkers,
    replaceToken,
    removeTokenLine,
    findTokens,
    findUnfilledTokens,
    composeTemplateBody,
} from './outlookTemplateSections';

export type { TemplateMarkers, ComposeOptions } from './outlookTemplateSections';
