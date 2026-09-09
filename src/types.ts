// The package's data contract, and the interface both platform implementations
// answer to.
//
// These types used to live in the Windows implementation, so macOS imported its
// own return shapes from it and its stubs re-declared them structurally. The two
// drifted, and because OutlookService dispatches by picking one implementation or
// the other, every drift surfaced in the PUBLIC types as a union — `matched:
// UndeliverableEmail[] | unknown[]` and the like. Both now implement
// `OutlookBridge` from here, so a drift is a compile error in the package rather
// than a union in the consumer's editor.

// ── Sending ──────────────────────────────────────────────────────────────
export interface SendEmailParams {
    /** SMTP address of the account to send from. Never falls back to the default
     *  account — an address that doesn't resolve is an error, not a silent switch. */
    emailAccount: string;
    /** One or more addresses, comma- or semicolon-separated. */
    to: string;
    cc?: string;
    subject: string;
    htmlBody: string;
    attachmentPath?: string;
    sendImmediately: boolean;
    /**
     * Draft mode only. True pops an Outlook compose window per email; false saves
     * straight to Drafts with no window — the only workable option for batches.
     * Defaults to true, matching the single-email callers that rely on the window.
     */
    openDraftWindow?: boolean;
}

// ── Replying ─────────────────────────────────────────────────────────────
export interface ReplyEmailParams {
    emailAccount: string;
    entryId: string;
    /** StoreID from the same listing row — disambiguates across mailboxes. */
    storeId?: string;
    /**
     * HTML inserted above the quoted original (a full document is reduced to its <body>
     * content). Optional when `templateSubject` is given — the body is then resolved from
     * the saved template server-side, so the caller never has to carry the (often large,
     * Word-generated) template HTML.
     */
    htmlBody?: string;
    /** Reply with a saved template's body, resolved by subject from the mailbox — an
     *  alternative to htmlBody that keeps a big template out of the caller's payload. */
    templateSubject?: string;
    /** Folder holding the template when templateSubject is used (default 'Templates'). */
    templateFolder?: string;
    /**
     * Section of the template to keep when one template holds several reply variants
     * between [[SECTION]] markers (see outlookTemplateSections). The other sections and
     * all markers are stripped before the reply is built.
     */
    templateSection?: string;
    /** `{{PLACEHOLDER}}` → HTML substituted into the composed body. */
    templatePlaceholders?: Record<string, string>;
    /**
     * Name of an Outlook signature (as listed by listOutlookSignatures) to substitute
     * into the template's `{{SIGNATURE}}` placeholder. Resolved by the package rather
     * than by the caller, so the signature HTML — images and all — never crosses the wire.
     */
    signatureName?: string;
    /**
     * Answer every recipient of the original (Reply All) rather than just its
     * sender. Off by default: widening a reply is the kind of mistake that can't
     * be taken back, so it has to be asked for.
     */
    replyAll?: boolean;
    sendImmediately?: boolean;
    openDraftWindow?: boolean;
}

export interface ReplyEmailResult {
    to: string;
    subject: string;
    /** The original email's sender — verify it's who you meant to answer. */
    repliedToSender: string;
}

// ── Drafts ───────────────────────────────────────────────────────────────
/** One draft that couldn't be sent, identified by its subject for the report. */
export interface DraftSendFailure {
    subject: string;
    error: string;
}

export interface SendAllDraftsResult {
    sent: number;
    failed: DraftSendFailure[];
}

/** One mail draft belonging to an account. */
export interface OutlookDraft {
    entryId: string;
    subject: string;
    /** The To line as Outlook renders it (display names). */
    to: string;
    /** Recipient addresses, best-effort — an unresolved Exchange entry falls back to its name. */
    toEmails: string[];
    /** Short plain-text preview — enough to tell one template section from another. */
    bodyPreview: string;
    hasAttachments: boolean;
    lastModified: string;
    /** Which Drafts folder it sits in, since two can hold one account's mail. */
    folderPath: string;
}

export interface ListDraftsResult {
    account: string;
    /** The Drafts folders scanned, in scan order. */
    foldersScanned: string[];
    /** Total matching drafts found, before `limit` was applied. */
    count: number;
    truncated: boolean;
    drafts: OutlookDraft[];
}

export interface DeleteDraftsResult {
    deleted: number;
    failed: { entryId: string; error: string }[];
}

/** One draft that couldn't be sent in a selective send, identified by both id and
 *  subject — id for the caller to retry or reconcile, subject for a human report. */
export interface SendDraftsFailure {
    entryId: string;
    subject: string;
    error: string;
}

export interface SendDraftsResult {
    sent: number;
    failed: SendDraftsFailure[];
}

// ── Deleting mail ────────────────────────────────────────────────────────
/** What happened to one id in a deleteOutlookEmails call. */
export interface DeleteMailOutcome {
    entryId: string;
    subject: string;
    /** The folder the item was actually in — the audit trail for what you deleted. */
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

export interface DeleteMailOptions {
    /** Lift the Inbox/Sent Items refusal, taking responsibility for reaching received mail. */
    allowProtected?: boolean;
    /** Resolve and report without deleting anything. */
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

// ── Bounce handling ──────────────────────────────────────────────────────
/** One bounce-back / non-delivery message found in the inbox. */
export interface UndeliverableEmail {
    entryId: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    /** Which rule flagged it (NDR class, mail-daemon sender, or a bounce subject phrase). */
    matchedReason: string;
    /** Best-effort recipient addresses parsed from the bounce body — which sends failed. */
    failedRecipients: string[];
}

export interface CleanUndeliverableResult {
    account: string;
    scannedDays: number;
    /** True when nothing was deleted (preview run). */
    dryRun: boolean;
    matchedCount: number;
    deletedCount: number;
    matched: UndeliverableEmail[];
    /** Matched items that could not be deleted (only populated on a delete run). */
    failed: DraftSendFailure[];
}

/** One sent message and the full set of SMTP addresses it went to. */
export interface SentRecipientGroup {
    entryId: string;
    subject: string;
    sentOn: string;
    recipients: string[];
}

// ── Reading ──────────────────────────────────────────────────────────────
export interface InboxEmail {
    entryId: string;
    storeId: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    bodyPreview: string;
    attachmentNames: string[];
    attachmentCount: number;
    /** Outlook path of the folder each email was read from, so a caller that
     *  scoped to a subfolder can confirm which folder actually matched.
     *  Populated on both platforms. */
    folderPath?: string;
}

export interface InboxSearchFilter {
    /** Glob-style subject pattern — `*` and `?` wildcards, e.g. '*invoice*'.
     *  Handed to Items.Restrict as the server-side prefilter that keeps a
     *  full-mailbox walk cheap, and re-checked against every subject that
     *  comes back, so it filters the result set on stores that won't run the
     *  query as well. Omit to restrict on date only. */
    subjectLike?: string;
    /** Regex re-checked client-side against each survivor's trimmed subject,
     *  since Restrict's `like` is a blunt substring match. A normal JS
     *  RegExp — only its `.source` crosses into the PowerShell/.NET regex
     *  engine, which reads the same syntax; `-match` is case-insensitive
     *  there by default regardless of the JS pattern's `i` flag. */
    subjectPattern?: RegExp;
    /** Drop subjects carrying a reply or forward prefix — RE:/FW:/Fwd: and the
     *  CJK equivalents. */
    excludeReplies?: boolean;
    /** Only return items carrying at least one attachment. */
    requireAttachment?: boolean;
    /** Folders under the Inbox to leave out of the walk, each given as a full
     *  folder path or as a bare folder name, matched case-insensitively. An
     *  excluded folder takes its whole subtree with it.
     *
     *  This is for the operator's OWN folders — an archive of sent copies, a
     *  "handled" pile — which no well-known-folder test can recognize because
     *  they are ordinary user folders that merely happen to be named like the
     *  special ones. The genuine Sent/Deleted/Drafts/Junk roots are skipped
     *  regardless and need no entry here.
     *
     *  macOS matches on the leaf name only; a path entry narrows to its last
     *  segment there. */
    excludeFolders?: string[];
    /** Restrict the scan to these folders, given the same way as
     *  `excludeFolders`. Empty (the default) scans the whole Inbox tree; the
     *  Inbox root itself is in scope only when it is named here.
     *
     *  Matched EXACTLY: naming a folder says nothing about its children, so a
     *  caller that wants a subtree lists the subtree. The walk still passes
     *  through folders that are not listed, so a nested folder can be reached
     *  without its parent being in scope.
     *
     *  This is deliberately not symmetric with `excludeFolders`, which does take
     *  its whole subtree — a pruned folder is never walked into, so nothing
     *  under it can be reached, let alone named back in. Exclusion therefore
     *  beats inclusion wherever the two overlap. */
    includeFolders?: string[];
}

export interface InboxSearchMatch {
    entryId: string;
    /** Store the item lives in — pass alongside entryId so it resolves unambiguously
     *  across mailboxes (replyOutlookEmail, saveEmailAttachments). */
    storeId: string;
    subject: string;
    senderName: string;
    /** SMTP address; resolved from the Exchange DN when the sender is an EX recipient. */
    senderEmail: string;
    /** 'yyyy-MM-dd HH:mm'. */
    receivedTime: string;
    body: string;
    /** Not saved to disk — pass the ones you want through saveEmailAttachments. */
    attachmentNames: string[];
    folderPath: string;
}

export interface SelectedEmail {
    entryId: string;
    storeId: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    body: string;
    /** Not saved to disk — pass the ones you want through saveEmailAttachments. */
    attachmentNames: string[];
}

export interface EmailBodyResult {
    entryId: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    /** The sender's own new text — everything above the quoted thread. */
    body: string;
    /** True when `body` hit maxChars and was cut. */
    truncated: boolean;
    /** Length of the sender's new text before any capping. */
    bodyLength: number;
    /** What the quoted thread was split on ("-----Original Message-----", …); empty when none was found. */
    quoteSeparator: string;
    /** Length of the quoted thread, reported even when it isn't returned. */
    quotedLength: number;
    /** The quoted thread — only populated when the caller asked for it. */
    quotedOriginal: string;
    attachmentNames: string[];
    attachmentCount: number;
}

// ── Folders ──────────────────────────────────────────────────────────────
export interface MailFolderRef {
    /** olDefaultFolders id of the root the walk starts from. */
    rootId: number;
    /** That root's name, for error messages. */
    rootLabel: string;
    /** Path segments below the root. */
    segments: string[];
}

export interface InboxFolderInfo {
    name: string;
    /** Outlook folder path, e.g. "\\\\mailbox\\Inbox\\Invoices". */
    folderPath: string;
    itemCount: number;
    /** 1 = direct child of Inbox, 2 = grandchild, … */
    depth: number;
}

export interface MoveEmailsResult {
    folderPath: string;
    folderCreated: boolean;
    moved: number;
    failed: { entryId: string; error: string }[];
}

// ── Attachments ──────────────────────────────────────────────────────────
export interface SavedAttachment {
    path: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
}

// ── Template emails ──────────────────────────────────────────────────────
/** One template email stored in a mailbox folder. */
export interface TemplateEmail {
    entryId: string;
    subject: string;
    /** Full HTML body — reusable verbatim as another email's htmlBody. Empty when the
     *  listing was requested without bodies (includeBody:false). */
    htmlBody: string;
    /** Short plain-text preview, returned instead of htmlBody when includeBody is false. */
    bodyPreview?: string;
    /** [[SECTION]] names when the template holds several reply variants (see
     *  outlookTemplateSections) — returned with or without the body, so a caller can
     *  confirm the sections it's about to ask for still exist. */
    sections?: string[];
    /** `{{PLACEHOLDER}}` names the template expects to have filled in. */
    placeholders?: string[];
    lastModified: string;
}

export interface TemplateFolderResult {
    folderFound: boolean;
    /** Outlook folder path when found, e.g. "\\\\mailbox\\Templates". */
    folderPath: string;
    templates: TemplateEmail[];
    /** When the folder wasn't found: the mailbox's folder names, so the caller can pick or create one. */
    availableFolders: string[];
}

export interface SaveTemplateResult {
    folderPath: string;
    folderCreated: boolean;
}

// ── The platform contract ────────────────────────────────────────────────
/**
 * What every platform implementation provides. `OutlookService` picks one at load
 * time and exposes it under these exact signatures.
 *
 * Not everything is implemented everywhere — macOS in particular still throws for
 * a good part of this. That is a runtime answer ("not implemented on this
 * platform"), deliberately not a type-level one: a consumer writing against the
 * package should see one signature per function, not a per-platform subset they
 * have to narrow.
 */
export interface OutlookBridge {
    getOutlookAccounts(): Promise<string[]>;

    sendOutlookEmail(params: SendEmailParams): Promise<void>;

    replyOutlookEmail(params: ReplyEmailParams): Promise<ReplyEmailResult>;

    sendAllDrafts(emailAccount: string): Promise<SendAllDraftsResult>;

    readInboxEmails(
        emailAccount: string,
        daysBack?: number,
        limit?: number,
        folder?: string,
    ): Promise<InboxEmail[]>;

    searchInboxByFilter(
        emailAccount: string,
        filter?: InboxSearchFilter,
        daysBack?: number,
    ): Promise<InboxSearchMatch[]>;

    readSelectedEmail(): Promise<SelectedEmail>;

    readEmailBody(
        entryId: string,
        storeId?: string,
        maxChars?: number,
        includeQuoted?: boolean,
    ): Promise<EmailBodyResult>;

    openOutlookEmail(entryId: string, storeId?: string): Promise<void>;

    listInboxFolders(emailAccount: string, maxDepth?: number): Promise<InboxFolderInfo[]>;

    moveOutlookEmails(
        emailAccount: string,
        entryIds: string[],
        folderName: string,
        createIfMissing?: boolean,
    ): Promise<MoveEmailsResult>;

    listOutlookDrafts(
        emailAccount: string,
        limit?: number,
        previewChars?: number,
    ): Promise<ListDraftsResult>;

    deleteOutlookDrafts(emailAccount: string, entryIds: string[]): Promise<DeleteDraftsResult>;

    /**
     * Send a chosen subset of an account's drafts by EntryID — e.g. after
     * `listOutlookDrafts` and a user review pass — rather than the account-wide
     * `sendAllDrafts`. Same "must be THIS account's draft" gate as
     * `deleteOutlookDrafts`: an id pointing at ordinary mail, or at another
     * account's draft, is refused and reported rather than sent.
     */
    sendDrafts(emailAccount: string, entryIds: string[]): Promise<SendDraftsResult>;

    deleteOutlookEmails(
        emailAccount: string,
        entryIds: string[],
        options?: DeleteMailOptions,
    ): Promise<DeleteMailResult>;

    purgeDeletedItems(
        emailAccount: string,
        olderThanDays?: number,
        dryRun?: boolean,
    ): Promise<PurgeDeletedItemsResult>;

    saveEmailAttachment(
        entryId: string,
        fileName: string,
        storeId?: string,
        destDir?: string,
    ): Promise<string>;

    saveEmailAttachmentDetailed(
        entryId: string,
        fileName: string,
        storeId?: string,
        destDir?: string,
    ): Promise<SavedAttachment>;

    saveEmailAttachments(
        entryId: string,
        fileNames: string[],
        storeId?: string,
        destDir?: string,
    ): Promise<SavedAttachment[]>;

    cleanUndeliverableEmails(
        emailAccount: string,
        daysBack?: number,
        dryRun?: boolean,
    ): Promise<CleanUndeliverableResult>;

    collectBouncedRecipients(
        emailAccount: string,
        daysBack?: number,
        scanDeleted?: boolean,
    ): Promise<string[]>;

    readSentRecipientGroups(
        emailAccount: string,
        daysBack?: number,
        limit?: number,
    ): Promise<SentRecipientGroup[]>;

    listOutlookSignatures(): Promise<string[]>;

    readOutlookSignatureHtml(name: string): Promise<string>;

    readTemplateEmails(
        emailAccount: string,
        folderName?: string,
        limit?: number,
        includeBody?: boolean,
        subject?: string,
    ): Promise<TemplateFolderResult>;

    saveTemplateEmail(
        emailAccount: string,
        subject: string,
        htmlBody: string,
        folderName?: string,
    ): Promise<SaveTemplateResult>;

    editEmailTemplate(label: string, currentHtml: string): Promise<string>;
}

// ── Capability discovery ─────────────────────────────────────────────────
/** Every operation whose availability can be asked about. */
export type BridgeCapability = keyof OutlookBridge;

/**
 * Which operations actually work on this machine.
 *
 * The alternative was matching on an error message: a consumer had no way to
 * learn a gap existed except by calling into it. A UI can now gray out the
 * buttons it can't back rather than discovering the gap when the user clicks one.
 *
 * Windows and macOS both answer the whole contract today, so every flag is
 * `true` on either; the map earns its place on the third case — any other OS,
 * where all of them are `false` and every call rejects with
 * `UNSUPPORTED_PLATFORM`. `false` is also how a platform would announce the
 * softer gap, an operation that answers emptily rather than throwing, so an
 * empty result there means "can't" rather than "none found".
 *
 * Each platform derives its map from its bridge object minus an explicit
 * not-ported list (see `shared/capabilities.ts`), and the key type is
 * `keyof OutlookBridge`, so a new function in the contract cannot be missed.
 */
export type CapabilityMap = Readonly<Record<BridgeCapability, boolean>>;
