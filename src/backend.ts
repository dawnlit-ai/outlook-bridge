// What a platform implements.
//
// The public operations (bridge.ts) check every argument, apply every default,
// and do all the work that is the same on both platforms — composing a reply
// from a template, splitting a body from its quoted thread, choosing where
// attachments go. What reaches a backend is a request in its final shape: counts
// are whole numbers in range, strings are trimmed, folders are parsed, and each
// field the platform's script needs is present. A backend's job is only to turn
// that request into Outlook automation and its output into the result.
import type {
    CleanUndeliverableResult,
    DeleteDraftsResult,
    DeleteMailResult,
    EmailLocator,
    InboxEmail,
    InboxFolderInfo,
    InboxSearchMatch,
    ListDraftsResult,
    MailFolderRef,
    MoveEmailsResult,
    PurgeDeletedItemsResult,
    ReplyEmailResult,
    SavedAttachment,
    SaveTemplateResult,
    SelectedEmail,
    SendDraftsResult,
    SentRecipientGroup,
    TemplateFolderResult,
} from './types';

/** What to do with a composed email. */
export type Disposition =
/** Send it now. */
    | 'send'
    /** Open it in a compose window for review. */
    | 'display'
    /** File it into Drafts without a window. */
    | 'save';

export interface AccountRequest {
    /** The mailbox's SMTP address, trimmed. */
    readonly account: string;
}

export interface SendRequest extends AccountRequest {
    /** One entry per recipient, each an address or a display-name form. */
    readonly to: readonly string[];
    readonly cc: readonly string[];
    readonly bcc: readonly string[];
    readonly subject: string;
    readonly htmlBody: string;
    /** Absolute paths, each verified to exist. */
    readonly attachments: readonly string[];
    readonly disposition: Disposition;
}

export interface ReplyRequest extends AccountRequest {
    readonly email: EmailLocator;
    /** The HTML to insert above the quoted original, fully composed. */
    readonly html: string;
    readonly replyAll: boolean;
    readonly disposition: Disposition;
}

export interface ReadInboxRequest extends AccountRequest {
    readonly folder: MailFolderRef;
    /** The folder as the caller wrote it, for messages. */
    readonly folderLabel: string;
    readonly daysBack: number;
    readonly limit: number;
    /** 0 means no body is read at all. */
    readonly previewChars: number;
}

export interface SearchRequest extends AccountRequest {
    readonly daysBack: number;
    readonly subjectLike?: string;
    readonly subjectPattern?: RegExp;
    readonly excludeReplies: boolean;
    readonly requireAttachment: boolean;
    readonly includeFolders: readonly string[];
    readonly excludeFolders: readonly string[];
    readonly includeBody: boolean;
}

/** One email as a backend reads it, before any splitting or capping. */
export interface RawEmail {
    readonly entryId: string;
    readonly subject: string;
    readonly senderName: string;
    readonly senderEmail: string;
    readonly receivedTime: string;
    /** The full plain-text body. */
    readonly body: string;
    readonly attachmentNames: readonly string[];
}

export interface ListFoldersRequest extends AccountRequest {
    readonly maxDepth: number;
}

export interface MoveRequest extends AccountRequest {
    /** Non-empty, deduplicated. */
    readonly entryIds: readonly string[];
    readonly folder: MailFolderRef;
    readonly folderLabel: string;
    readonly createIfMissing: boolean;
}

export interface ListDraftsRequest extends AccountRequest {
    readonly limit: number;
    readonly previewChars: number;
}

export interface EntryIdsRequest extends AccountRequest {
    /** Non-empty, deduplicated. */
    readonly entryIds: readonly string[];
}

export interface DeleteEmailsRequest extends EntryIdsRequest {
    readonly allowProtected: boolean;
    readonly dryRun: boolean;
}

export interface PurgeRequest extends AccountRequest {
    readonly olderThanDays: number;
    readonly dryRun: boolean;
}

export interface SaveAttachmentsRequest {
    readonly email: EmailLocator;
    /** Non-empty; matched exactly first, then ignoring case and whitespace. */
    readonly fileNames: readonly string[];
    /** An existing directory to save into. */
    readonly destDir: string;
}

export interface CleanUndeliverableRequest extends AccountRequest {
    readonly daysBack: number;
    readonly dryRun: boolean;
}

export interface CollectBouncesRequest extends AccountRequest {
    readonly daysBack: number;
    readonly includeDeletedItems: boolean;
}

export interface SentGroupsRequest extends AccountRequest {
    readonly daysBack: number;
    readonly limit: number;
}

export interface ReadTemplatesRequest extends AccountRequest {
    readonly folder: string;
    readonly limit: number;
    readonly includeBody: boolean;
    /** Only the template with this subject, case-insensitively. */
    readonly subject?: string;
}

export interface SaveTemplateRequest extends AccountRequest {
    readonly folder: string;
    readonly subject: string;
    readonly htmlBody: string;
}

export interface EditTemplateRequest {
    readonly label: string;
    readonly html: string;
}

/** One platform's Outlook automation. */
export interface Backend {
    getOutlookAccounts(): Promise<string[]>;

    sendOutlookEmail(request: SendRequest): Promise<void>;

    replyOutlookEmail(request: ReplyRequest): Promise<ReplyEmailResult>;

    readInboxEmails(request: ReadInboxRequest): Promise<InboxEmail[]>;

    searchInboxByFilter(request: SearchRequest): Promise<InboxSearchMatch[]>;

    readSelectedEmail(): Promise<SelectedEmail>;

    readEmailBody(email: EmailLocator): Promise<RawEmail>;

    openOutlookEmail(email: EmailLocator): Promise<void>;

    listInboxFolders(request: ListFoldersRequest): Promise<InboxFolderInfo[]>;

    moveOutlookEmails(request: MoveRequest): Promise<MoveEmailsResult>;

    listOutlookDrafts(request: ListDraftsRequest): Promise<ListDraftsResult>;

    sendDrafts(request: EntryIdsRequest): Promise<SendDraftsResult>;

    sendAllDrafts(request: AccountRequest): Promise<SendDraftsResult>;

    deleteOutlookDrafts(request: EntryIdsRequest): Promise<DeleteDraftsResult>;

    deleteOutlookEmails(request: DeleteEmailsRequest): Promise<DeleteMailResult>;

    purgeDeletedItems(request: PurgeRequest): Promise<PurgeDeletedItemsResult>;

    saveEmailAttachments(request: SaveAttachmentsRequest): Promise<SavedAttachment[]>;

    cleanUndeliverableEmails(request: CleanUndeliverableRequest): Promise<CleanUndeliverableResult>;

    collectBouncedRecipients(request: CollectBouncesRequest): Promise<string[]>;

    readSentRecipientGroups(request: SentGroupsRequest): Promise<SentRecipientGroup[]>;

    listOutlookSignatures(): Promise<string[]>;

    /** '' when there is no signature by that name. */
    readOutlookSignatureHtml(name: string): Promise<string>;

    readTemplateEmails(request: ReadTemplatesRequest): Promise<TemplateFolderResult>;

    saveTemplateEmail(request: SaveTemplateRequest): Promise<SaveTemplateResult>;

    editEmailTemplate(request: EditTemplateRequest): Promise<string>;
}
