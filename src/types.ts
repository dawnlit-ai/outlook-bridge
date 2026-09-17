// The public contract: what the package returns, what it takes, and the
// OutlookBridge interface every entry point answers to.
//
// Two conventions hold throughout:
//  - Timestamps are local time as 'yyyy-MM-dd HH:mm', which sorts correctly as
//    a plain string and is what both platforms can produce without a timezone
//    table.
//  - Ids are platform-specific (see EmailLocator). Don't persist one and look it
//    up on the other platform.

// ── Identifying an email ─────────────────────────────────────────────────
/**
 * Where an email lives, as a listing reported it.
 *
 * Every listing row (`InboxEmail`, `InboxSearchMatch`, `SelectedEmail`, …)
 * carries both fields, so a row can be passed straight back in wherever an
 * `EmailRef` is accepted — which is also the reliable way to keep the two ids
 * from different rows apart.
 */
export interface EmailLocator {
    /** Windows: the MAPI EntryID. macOS: Outlook for Mac's integer message id. */
    readonly entryId: string;
    /**
     * Windows: the StoreID the email was listed from, which lets it resolve in
     * any mounted mailbox rather than only the default one. macOS has no
     * equivalent and ignores it.
     */
    readonly storeId?: string;
}

/** An email to act on: its entry id alone, or anything carrying a locator. */
export type EmailRef = string | EmailLocator;

/**
 * Recipients: one address, a comma- or semicolon-separated list, or an array.
 * Display-name forms such as `"Doe, Jo" <jo@example.com>` are kept whole.
 */
export type Recipients = string | readonly string[];

// ── Sending and replying ─────────────────────────────────────────────────
export interface SendEmailParams {
    /** SMTP address of the mailbox to send from. Never falls back to the
     *  default account — an address that doesn't resolve is an error. */
    emailAccount: string;
    to?: Recipients;
    cc?: Recipients;
    bcc?: Recipients;
    subject: string;
    htmlBody: string;
    /** Absolute paths of files to attach. Each has to exist. */
    attachments?: readonly string[];
    /** Send now. Default false: the email is staged as a draft instead. */
    sendImmediately?: boolean;
    /**
     * Drafts only. True (the default) opens an Outlook compose window; false
     * files the draft silently into Drafts — the workable choice for batches.
     */
    openDraftWindow?: boolean;
}

export interface ReplyEmailParams extends EmailLocator {
    /** SMTP address of the mailbox to reply from. */
    emailAccount: string;
    /**
     * HTML inserted above the quoted original (a full document is reduced to
     * its <body> content). Optional when `templateSubject` is given.
     */
    htmlBody?: string;
    /**
     * Reply with a template email saved in the mailbox, found by subject, so
     * the caller never carries a large template's HTML itself.
     */
    templateSubject?: string;
    /** Folder holding the template. Default 'Templates'. */
    templateFolder?: string;
    /**
     * The [[SECTION]] of the template to keep, when one template holds several
     * reply variants (see composeTemplateBody). Required if it has sections.
     */
    templateSection?: string;
    /** `{{PLACEHOLDER}}` → HTML substituted into the composed body. */
    templatePlaceholders?: Readonly<Record<string, string>>;
    /**
     * An Outlook signature (by the name listOutlookSignatures reports) to
     * substitute into the body's `{{SIGNATURE}}` placeholder. Resolved here, so
     * the signature's HTML never crosses the caller's boundary.
     */
    signatureName?: string;
    /**
     * Answer every recipient of the original instead of only its sender.
     * Default false: widening a reply cannot be taken back, so it is opt-in.
     */
    replyAll?: boolean;
    /** Send now. Default false. */
    sendImmediately?: boolean;
    /** Drafts only: open a compose window (the default) or file silently. */
    openDraftWindow?: boolean;
}

export interface ReplyEmailResult {
    /** Who the reply is addressed to. Windows reports Outlook's To line
     *  (display names); macOS the resolved addresses. */
    to: string;
    subject: string;
    /** The original email's sender — check it is who you meant to answer. */
    repliedToSender: string;
}

// ── Per-item outcomes ────────────────────────────────────────────────────
/** One item a batch operation could not handle, and why. */
export interface ItemFailure {
    entryId: string;
    /** The item's subject when it could be read, else ''. */
    subject: string;
    error: string;
}

// ── Drafts ───────────────────────────────────────────────────────────────
/** One mail draft belonging to an account. */
export interface OutlookDraft {
    entryId: string;
    subject: string;
    /** The To line as Outlook renders it. */
    to: string;
    /** Recipient addresses, best effort — an unresolved Exchange entry falls back to its name. */
    toEmails: string[];
    /** Short plain-text preview, enough to tell similar drafts apart. */
    bodyPreview: string;
    hasAttachments: boolean;
    lastModified: string;
    /** The Drafts folder it sits in; one account's drafts can span two. */
    folderPath: string;
}

export interface ListDraftsOptions {
    /** Newest-first cap on the drafts returned. Default 100. */
    limit?: number;
    /** Characters of body preview per draft. Default 300; 0 omits it. */
    previewChars?: number;
}

export interface ListDraftsResult {
    account: string;
    /** The Drafts folders scanned, in scan order. */
    foldersScanned: string[];
    /** Matching drafts found, before `limit` was applied. */
    count: number;
    truncated: boolean;
    drafts: OutlookDraft[];
}

export interface SendDraftsResult {
    sent: number;
    failed: ItemFailure[];
}

export interface DeleteDraftsResult {
    deleted: number;
    failed: ItemFailure[];
}

// ── Deleting mail ────────────────────────────────────────────────────────
export interface DeleteMailOptions {
    /** Lift the refusal to touch Inbox and Sent Items (and their subfolders). */
    allowProtected?: boolean;
    /** Resolve and report every id without deleting anything. */
    dryRun?: boolean;
}

/** What happened to one id in a deleteOutlookEmails call. */
export interface DeleteMailOutcome {
    entryId: string;
    subject: string;
    /** The folder the item was actually in — the audit trail for a wrong id. */
    folderPath: string;
    status: 'deleted' | 'would-delete' | 'refused' | 'failed';
    reason: string;
}

export interface DeleteMailResult {
    dryRun: boolean;
    deleted: number;
    refused: number;
    failed: number;
    items: DeleteMailOutcome[];
}

export interface PurgeDeletedItemsOptions {
    /** Keep items newer than this many days. Default 0: purge everything. */
    olderThanDays?: number;
    /** Count what would be purged without destroying anything. */
    dryRun?: boolean;
}

export interface PurgeDeletedItemsResult {
    folderPath: string;
    dryRun: boolean;
    /** Items old enough to qualify. */
    matched: number;
    purged: number;
    /** Items left because they were newer than `olderThanDays`. */
    kept: number;
    failed: number;
}

// ── Bounces ──────────────────────────────────────────────────────────────
/** One bounce-back / non-delivery message found in a mailbox. */
export interface UndeliverableEmail {
    entryId: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    /** Which rule flagged it: an NDR message class, a mail-daemon sender, or a bounce subject phrase. */
    matchedReason: string;
    /** Best-effort addresses parsed from the bounce body — the sends that failed. */
    failedRecipients: string[];
}

export interface CleanUndeliverableOptions {
    /** How far back to scan. Default 30. */
    daysBack?: number;
    /** Report without deleting. Default TRUE — deleting has to be asked for. */
    dryRun?: boolean;
}

export interface CleanUndeliverableResult {
    account: string;
    scannedDays: number;
    dryRun: boolean;
    matchedCount: number;
    deletedCount: number;
    matched: UndeliverableEmail[];
    /** Matched items that could not be deleted (a delete run only). */
    failed: ItemFailure[];
}

export interface CollectBouncedRecipientsOptions {
    /** How far back to scan. Default 30. */
    daysBack?: number;
    /** Scan Deleted Items too, so bounces already cleaned away still count. Default true. */
    includeDeletedItems?: boolean;
}

export interface SentRecipientGroupsOptions {
    /** How far back to read Sent Items. Default 30. */
    daysBack?: number;
    /** Newest-first cap on messages. Default 3000. */
    limit?: number;
}

/** One sent message and every SMTP address it went to (To, CC and BCC). */
export interface SentRecipientGroup {
    entryId: string;
    subject: string;
    sentOn: string;
    recipients: string[];
}

// ── Reading ──────────────────────────────────────────────────────────────
export interface ReadInboxOptions {
    /**
     * Read this folder instead of the Inbox root: a folder under the Inbox by
     * name or path, a well-known folder ('Sent Items', 'Drafts', …), or a full
     * folder path — see mailFolderRef. Only that folder is read, not its
     * subfolders.
     */
    folder?: string;
    /** Only mail from the last this-many days. Default 60. */
    daysBack?: number;
    /** Newest-first cap on rows. Default 50. */
    limit?: number;
    /** Characters of plain-text body per row. Default 600; 0 skips reading bodies. */
    previewChars?: number;
}

export interface InboxEmail extends EmailLocator {
    entryId: string;
    storeId: string;
    subject: string;
    /** For outgoing folders (Sent Items, Outbox, Drafts) these carry the
     *  RECIPIENT, since a folder of mail "from" its own owner says nothing. */
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    bodyPreview: string;
    attachmentNames: string[];
    attachmentCount: number;
    /** The folder the email was read from. */
    folderPath: string;
}

export interface InboxSearchFilter {
    /** Only mail from the last this-many days. Default 60 — a scan has no
     *  unbounded mode; ask for 365 to reach back a year. */
    daysBack?: number;
    /** Glob-style subject pattern: `*` matches any run, `?` one character,
     *  case-insensitive, e.g. '*invoice*'. */
    subjectLike?: string;
    /**
     * Regular expression tested against each subject, case-insensitively
     * whatever the RegExp's own flags. On Windows the pattern runs in .NET's
     * regex engine, so keep to syntax JavaScript and .NET share.
     */
    subjectPattern?: RegExp;
    /** Drop subjects carrying a reply or forward prefix (RE:, FW:, AW:, 回复:, …). */
    excludeReplies?: boolean;
    /** Only return mail with at least one attachment. */
    requireAttachment?: boolean;
    /**
     * Folders under the Inbox to leave out, by full folder path or bare name,
     * case-insensitively. An excluded folder takes its whole subtree with it.
     * Sent, Drafts, Deleted and Junk are always skipped, even where the profile
     * nests them under the Inbox. macOS matches the leaf name only.
     */
    excludeFolders?: readonly string[];
    /**
     * Scan only these folders, given the same way. Matched exactly: naming a
     * folder does not bring its children, and the Inbox root is included only
     * when named. Exclusion wins where the two overlap.
     */
    includeFolders?: readonly string[];
    /** Return each match's full plain-text body. Default true; false makes a
     *  scan much lighter when only the metadata is needed. */
    includeBody?: boolean;
}

export interface InboxSearchMatch extends EmailLocator {
    entryId: string;
    storeId: string;
    subject: string;
    senderName: string;
    /** SMTP address, resolved from the Exchange DN where the sender is an Exchange user. */
    senderEmail: string;
    receivedTime: string;
    /** The plain-text body; '' when the search asked for no bodies. */
    body: string;
    /** Not saved to disk — pass the ones you want to saveEmailAttachments. */
    attachmentNames: string[];
    folderPath: string;
}

export interface SelectedEmail extends EmailLocator {
    entryId: string;
    storeId: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    body: string;
    /** Not saved to disk — pass the ones you want to saveEmailAttachments. */
    attachmentNames: string[];
}

export interface ReadEmailBodyOptions {
    /** Cap on the returned body. Default 8000. */
    maxChars?: number;
    /** Also return the quoted thread below the reply. Default false. */
    includeQuoted?: boolean;
    /** Cap on the quoted thread when it is returned. Default 4000. */
    maxQuotedChars?: number;
}

export interface EmailBodyResult {
    entryId: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    /** The sender's own new text — everything above the quoted thread. */
    body: string;
    /** True when `body` was cut at maxChars. */
    truncated: boolean;
    /** Length of the sender's new text before any cut. */
    bodyLength: number;
    /** What the quoted thread was split on; '' when none was found. */
    quoteSeparator: string;
    /** Length of the quoted thread, reported even when it isn't returned. */
    quotedLength: number;
    /** The quoted thread — only when asked for. */
    quotedOriginal: string;
    attachmentNames: string[];
    attachmentCount: number;
}

// ── Folders ──────────────────────────────────────────────────────────────
export interface MailFolderRef {
    /** olDefaultFolders id of the well-known root the path starts from. */
    rootId: number;
    /** That root as the caller spelled it, for messages. */
    rootLabel: string;
    /** Path segments below the root. */
    segments: string[];
}

export interface ListFoldersOptions {
    /** How many levels below the Inbox to list, 1–10. Default 2. */
    maxDepth?: number;
}

export interface InboxFolderInfo {
    name: string;
    /** Outlook folder path, e.g. `\\mailbox\Inbox\Invoices`. */
    folderPath: string;
    itemCount: number;
    /** 1 = direct child of the Inbox, 2 = grandchild, … */
    depth: number;
}

export interface MoveEmailsOptions {
    /** Create the destination — the whole missing chain — when absent. Default false. */
    createIfMissing?: boolean;
}

export interface MoveEmailsResult {
    folderPath: string;
    folderCreated: boolean;
    moved: number;
    failed: ItemFailure[];
}

// ── Attachments ──────────────────────────────────────────────────────────
export interface SaveAttachmentOptions {
    /**
     * Directory to save into, created if absent. Without one, each call saves
     * into a fresh private directory, so two emails that both carry
     * 'invoice.pdf' can never overwrite each other.
     */
    destDir?: string;
}

export interface SavedAttachment {
    /** Where the file was written. A name already taken in the directory gets
     *  a numbered suffix rather than overwriting what is there. */
    path: string;
    /** The attachment's name as the email carries it. */
    fileName: string;
    /** Who the email was from — confirm it is the email you meant. */
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
}

// ── Template emails ──────────────────────────────────────────────────────
export interface ReadTemplatesOptions {
    /** Mailbox folder holding the templates, searched up to 3 levels deep. Default 'Templates'. */
    folder?: string;
    /** Return only the template with this subject (case-insensitive). */
    subject?: string;
    /** Return each full HTML body. Default true; false returns previews. */
    includeBody?: boolean;
    /** Newest-first cap on templates, 1–200. Default 20. */
    limit?: number;
}

/** One template email stored in a mailbox folder. */
export interface TemplateEmail {
    entryId: string;
    subject: string;
    /** Full HTML body with embedded (cid:) images removed; '' without includeBody. */
    htmlBody: string;
    /** Short plain-text preview; '' with includeBody. */
    bodyPreview: string;
    /** [[SECTION]] names, returned with or without the body. */
    sections: string[];
    /** `{{PLACEHOLDER}}` names the template expects filled. */
    placeholders: string[];
    lastModified: string;
}

export interface TemplateFolderResult {
    folderFound: boolean;
    /** The folder's path when found. */
    folderPath: string;
    templates: TemplateEmail[];
    /** When the folder wasn't found: the mailbox's top-level folder names. */
    availableFolders: string[];
}

export interface SaveTemplateParams {
    subject: string;
    htmlBody: string;
    /** Folder to save into, created at the mailbox root if absent. Default 'Templates'. */
    folder?: string;
}

export interface SaveTemplateResult {
    folderPath: string;
    folderCreated: boolean;
}

// ── The contract ─────────────────────────────────────────────────────────
/**
 * Everything the package does, under one signature per operation whatever the
 * platform. A platform that cannot do something says so at run time — see
 * `CapabilityMap` — rather than through a narrower type.
 */
export interface OutlookBridge {
    /** Every mailbox this Outlook profile can reach, as SMTP addresses. */
    getOutlookAccounts(): Promise<string[]>;

    /** Send an email, or stage it as a draft. */
    sendOutlookEmail(params: SendEmailParams): Promise<void>;

    /** Reply to an email, with the new text above the quoted original. */
    replyOutlookEmail(params: ReplyEmailParams): Promise<ReplyEmailResult>;

    /** Recent mail from the Inbox root or one named folder, newest first. */
    readInboxEmails(emailAccount: string, options?: ReadInboxOptions): Promise<InboxEmail[]>;

    /** Walk every folder under the Inbox for the mail matching `filter`. */
    searchInboxByFilter(emailAccount: string, filter?: InboxSearchFilter): Promise<InboxSearchMatch[]>;

    /** The email currently selected, or open, in Outlook. */
    readSelectedEmail(): Promise<SelectedEmail>;

    /** One email's plain-text body, split from the thread it quotes. */
    readEmailBody(email: EmailRef, options?: ReadEmailBodyOptions): Promise<EmailBodyResult>;

    /** Open an email in Outlook. */
    openOutlookEmail(email: EmailRef): Promise<void>;

    /** The folders under an account's Inbox. */
    listInboxFolders(emailAccount: string, options?: ListFoldersOptions): Promise<InboxFolderInfo[]>;

    /** File emails into a folder. Moving rewrites each email's entry id. */
    moveOutlookEmails(
        emailAccount: string,
        entryIds: readonly string[],
        folder: string,
        options?: MoveEmailsOptions,
    ): Promise<MoveEmailsResult>;

    /** The mail drafts bound to an account, newest first. */
    listOutlookDrafts(emailAccount: string, options?: ListDraftsOptions): Promise<ListDraftsResult>;

    /** Send the named drafts — each has to be one of this account's drafts. */
    sendDrafts(emailAccount: string, entryIds: readonly string[]): Promise<SendDraftsResult>;

    /** Send every draft bound to an account. */
    sendAllDrafts(emailAccount: string): Promise<SendDraftsResult>;

    /** Delete the named drafts (into Deleted Items) — each has to be one of this account's drafts. */
    deleteOutlookDrafts(emailAccount: string, entryIds: readonly string[]): Promise<DeleteDraftsResult>;

    /** Delete mail by id into Deleted Items, refusing received and sent mail unless allowed. */
    deleteOutlookEmails(
        emailAccount: string,
        entryIds: readonly string[],
        options?: DeleteMailOptions,
    ): Promise<DeleteMailResult>;

    /** Permanently empty an account's Deleted Items. The one irreversible operation. */
    purgeDeletedItems(emailAccount: string, options?: PurgeDeletedItemsOptions): Promise<PurgeDeletedItemsResult>;

    /** Save one attachment to disk. */
    saveEmailAttachment(email: EmailRef, fileName: string, options?: SaveAttachmentOptions): Promise<SavedAttachment>;

    /** Save several attachments from one email, in the order named. */
    saveEmailAttachments(
        email: EmailRef,
        fileNames: readonly string[],
        options?: SaveAttachmentOptions,
    ): Promise<SavedAttachment[]>;

    /** Find bounce-backs in the Inbox and, unless dry-running, delete them. */
    cleanUndeliverableEmails(
        emailAccount: string,
        options?: CleanUndeliverableOptions,
    ): Promise<CleanUndeliverableResult>;

    /** The addresses bounce-backs report as having failed, deduplicated. */
    collectBouncedRecipients(emailAccount: string, options?: CollectBouncedRecipientsOptions): Promise<string[]>;

    /** Sent Items, each message with the full set of addresses it went to. */
    readSentRecipientGroups(emailAccount: string, options?: SentRecipientGroupsOptions): Promise<SentRecipientGroup[]>;

    /** The names of the Outlook signatures on this machine. */
    listOutlookSignatures(): Promise<string[]>;

    /** A signature's HTML, or '' when there is no signature by that name. */
    readOutlookSignatureHtml(name: string): Promise<string>;

    /** The template emails saved in a mailbox folder. */
    readTemplateEmails(emailAccount: string, options?: ReadTemplatesOptions): Promise<TemplateFolderResult>;

    /** Save a new template email. Always adds; never overwrites. */
    saveTemplateEmail(emailAccount: string, template: SaveTemplateParams): Promise<SaveTemplateResult>;

    /**
     * Open HTML in an Outlook compose window, wait for the person editing it to
     * save and close, and return what they saved. Has no timeout; cancel it
     * with an AbortSignal.
     */
    editEmailTemplate(label: string, currentHtml: string): Promise<string>;
}

// ── Capability discovery ─────────────────────────────────────────────────
/** Every operation whose availability can be asked about. */
export type BridgeCapability = keyof OutlookBridge;

/**
 * Which operations work on this machine: all of them on Windows and macOS,
 * none anywhere else. Lets a UI disable what it cannot back instead of
 * discovering the gap when a user clicks.
 */
export type CapabilityMap = Readonly<Record<BridgeCapability, boolean>>;
