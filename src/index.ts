// The public API. Listed explicitly rather than `export *`, so the package
// exposes only what it means to support; the `exports` map in package.json
// closes the deep-import path around it.
import { defaultBridge } from './bridge';

// Every operation as a plain function, running under the process-wide settings.
export const {
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
    saveEmailAttachment,
    saveEmailAttachments,
    cleanUndeliverableEmails,
    collectBouncedRecipients,
    readSentRecipientGroups,
    listOutlookSignatures,
    readOutlookSignatureHtml,
    readTemplateEmails,
    saveTemplateEmail,
    editEmailTemplate,
} = defaultBridge;

// A bridge with its own settings, and asking what works on this machine.
export { createOutlookBridge, capabilities, supports } from './bridge';
export type { OutlookBridgeInstance } from './bridge';

// Process-wide settings.
export { configure } from './runtime';
export type { BridgeOptions, BridgeDebugEvent, ResolvedConfig } from './runtime';

// Failures carry a stable `code`; the messages beside them are not an API.
export {
    OutlookError,
    UnsupportedPlatformError,
    NotImplementedError,
    AccountNotFoundError,
    NotFoundError,
    InvalidRequestError,
    ScriptError,
    OutputTooLargeError,
    TimeoutError,
    AbortedError,
} from './errors';
export type { OutlookErrorCode, NotFoundKind, ScriptRunner } from './errors';

// Platform-neutral helpers, usable without Outlook.
export { mailFolderRef, splitQuotedOriginal, WELL_KNOWN_FOLDERS } from './mail';
export {
    composeTemplateBody,
    findTemplateMarkers,
    findTokens,
    findUnfilledTokens,
    removeTokenLine,
    replaceToken,
} from './templateBody';
export type { ComposeOptions, TemplateMarkers } from './templateBody';

export type {
    OutlookBridge,
    BridgeCapability,
    CapabilityMap,
    EmailLocator,
    EmailRef,
    Recipients,
    SendEmailParams,
    ReplyEmailParams,
    ReplyEmailResult,
    ItemFailure,
    ReadInboxOptions,
    InboxEmail,
    InboxSearchFilter,
    InboxSearchMatch,
    SelectedEmail,
    ReadEmailBodyOptions,
    EmailBodyResult,
    MailFolderRef,
    ListFoldersOptions,
    InboxFolderInfo,
    MoveEmailsOptions,
    MoveEmailsResult,
    ListDraftsOptions,
    ListDraftsResult,
    OutlookDraft,
    SendDraftsResult,
    DeleteDraftsResult,
    DeleteMailOptions,
    DeleteMailOutcome,
    DeleteMailResult,
    PurgeDeletedItemsOptions,
    PurgeDeletedItemsResult,
    SaveAttachmentOptions,
    SavedAttachment,
    CleanUndeliverableOptions,
    CleanUndeliverableResult,
    UndeliverableEmail,
    CollectBouncedRecipientsOptions,
    SentRecipientGroupsOptions,
    SentRecipientGroup,
    ReadTemplatesOptions,
    TemplateEmail,
    TemplateFolderResult,
    SaveTemplateParams,
    SaveTemplateResult,
} from './types';
