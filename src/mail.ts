// Platform-neutral mail helpers: folder strings, reply prefixes, and splitting a
// reply from the thread it quotes. Pure functions, usable without Outlook, and
// shared by both platforms so neither can come to read these differently.
import { InvalidRequestError } from './errors';
import type { MailFolderRef } from './types';

/**
 * Outlook's olDefaultFolders ids for the folders this package addresses by
 * role rather than by name. The ids are Outlook's own, and a folder's NAME is
 * localized, so a role is always resolved through its id.
 */
export const FolderId = Object.freeze({
    DeletedItems: 3,
    Outbox: 4,
    SentMail: 5,
    Inbox: 6,
    Drafts: 16,
    Junk: 23,
});

/**
 * The folders addressable by name rather than by walking from the Inbox, with
 * their olDefaultFolders id. A well-known name wins over a user folder of the
 * same name; reach the latter by qualifying it (`Inbox\Drafts`).
 */
export const WELL_KNOWN_FOLDERS: Readonly<Record<string, number>> = Object.freeze({
    'inbox': FolderId.Inbox,
    'sent items': FolderId.SentMail,
    'sent': FolderId.SentMail,
    'sent mail': FolderId.SentMail,
    'drafts': FolderId.Drafts,
    'deleted items': FolderId.DeletedItems,
    'deleted': FolderId.DeletedItems,
    'trash': FolderId.DeletedItems,
    'junk email': FolderId.Junk,
    'junk': FolderId.Junk,
    'outbox': FolderId.Outbox,
});

/** The Inbox root, as a folder reference. */
export const INBOX_REF: Readonly<MailFolderRef> = Object.freeze({
    rootId: FolderId.Inbox,
    rootLabel: 'Inbox',
    segments: [],
});

/**
 * Why `deleteOutlookEmails` refuses an item, worded once for both platforms —
 * it reaches the caller verbatim as a per-item `reason`.
 */
export const PROTECTED_MAIL_REASON =
    'received or sent mail (Inbox/Sent Items or a subfolder) - pass allowProtected to override';

/**
 * Roots whose items are outgoing mail, which carries a sent time where incoming
 * mail carries a received one — so a date filter has to know which it's on.
 */
export function isOutgoingRoot(rootId: number): boolean {
    return rootId === FolderId.SentMail || rootId === FolderId.Outbox || rootId === FolderId.Drafts;
}

/**
 * Roots an Inbox walk must never descend into.
 *
 * On an Exchange profile these sit beside the Inbox and a walk from it can't
 * reach them anyway. On an IMAP profile Outlook nests them UNDER the Inbox, so
 * the same walk would hand back mail already sent, drafted or thrown away as if
 * it had just arrived. They are matched by id, never by name: the names are
 * localized.
 */
export const NON_INCOMING_ROOTS: readonly number[] = Object.freeze([
    FolderId.DeletedItems,
    FolderId.Outbox,
    FolderId.SentMail,
    FolderId.Drafts,
    FolderId.Junk,
]);

/**
 * Subject prefixes that mark a reply or a forward.
 *
 * Kept as a pattern SOURCE because it runs in two engines: macOS tests it as a
 * JavaScript RegExp, Windows inside PowerShell's .NET one, and both read every
 * construct used here the same way. Non-ASCII is written as \uXXXX so the
 * generated PowerShell stays plain ASCII.
 *
 * Not only English: mail threaded in Chinese, Japanese or Korean replies under
 * 回复:/返信:/답장: and forwards under 转发:/転送:/전달:, and the European
 * clients use AW:, WG:, SV:, TR:, RV:, ENC: and friends. The colon may be the
 * full-width one a CJK input method types, and Outlook's reply counter (RE[2]:)
 * can sit before it.
 */
const REPLY_PREFIXES: readonly string[] = [
    're', 'fwd?', 'aw', 'wg', 'sv', 'vs', 'vb', 'tr', 'res', 'rv', 'enc', 'odp',
    '\\u56DE\\u590D', '\\u56DE\\u8986', '\\u7B54\\u590D', '\\u7B54\\u8986',
    '\\u8F6C\\u53D1', '\\u8F49\\u767C', '\\u8FD4\\u4FE1', '\\u8EE2\\u9001',
    '\\uB2F5\\uC7A5', '\\uC804\\uB2EC',
];

/** The reply/forward prefixes as one case-insensitive pattern source. */
export const REPLY_PREFIX_SOURCE =
    '^\\s*(?:' + REPLY_PREFIXES.join('|') + ')\\s*(?:\\[\\d+\\])?\\s*[:\\uFF1A]';
export const REPLY_PREFIX = new RegExp(REPLY_PREFIX_SOURCE, 'i');

/**
 * Whether a subject opens on a reply or forward prefix.
 *
 * This reads what the subject claims, not how the email was made: a forward
 * carries a prefix too, and a client can leave one off a real reply.
 */
export function hasReplyPrefix(subject: string): boolean {
    return REPLY_PREFIX.test(subject);
}

/**
 * The subject every message in a thread shares: every reply and forward prefix
 * stripped, however many clients stacked them (`RE: FW: RE[2]: Rate request`
 * is `Rate request`), and whitespace collapsed. Case is kept, so lower-case
 * both sides to compare two threads.
 */
export function threadSubject(subject: string): string {
    let rest = subject;
    // Each pass removes at least the prefix's colon, so this always ends.
    while (REPLY_PREFIX.test(rest)) rest = rest.replace(REPLY_PREFIX, '');
    return rest.replace(/\s+/g, ' ').trim();
}

/**
 * A subject glob as an anchored regex source: `*` matches any run, `?` one
 * character, everything else itself. Written with constructs JavaScript and
 * .NET read identically, so both platforms apply exactly the same test.
 */
export function subjectGlobSource(glob: string): string {
    let source = '';
    for (const ch of glob) {
        if (ch === '*') source += '.*';
        else if (ch === '?') source += '.';
        else source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return `^${source}$`;
}

/** The leaf name of a folder path or bare name. */
export function folderLeafName(entry: string): string {
    return entry.trim().split(/[\\/]/).filter(Boolean).pop() ?? '';
}

/** A full Outlook folder path starts with the mailbox: `\\mailbox@example.com\…`. */
const MAILBOX_PREFIX = /^\\\\[^\\/]+[\\/]/;

function folderSegments(folder: string): string[] {
    return folder.trim().replace(MAILBOX_PREFIX, '').split(/[\\/]+/).map(s => s.trim()).filter(Boolean);
}

/**
 * Resolve a folder string to a well-known root plus the segments below it.
 *
 * Accepts every shape a person has to hand: a bare name ('Invoices'), a
 * relative path ('Invoices\Paid'), a well-known folder ('Sent Items'), a path
 * under one ('Deleted Items\2026'), and the full path listInboxFolders prints
 * ('\\team@example.com\Inbox\Invoices'). Anything not starting at a well-known
 * folder is taken as relative to the Inbox.
 */
export function mailFolderRef(folder: string): MailFolderRef {
    const segments = folderSegments(folder);
    const head = segments[0]?.toLowerCase();
    if (head !== undefined && Object.prototype.hasOwnProperty.call(WELL_KNOWN_FOLDERS, head)) {
        return {rootId: WELL_KNOWN_FOLDERS[head], rootLabel: segments[0], segments: segments.slice(1)};
    }
    return {rootId: FolderId.Inbox, rootLabel: 'Inbox', segments};
}

/**
 * A caller's folder option as a reference: none given means the Inbox root,
 * and a string that names no folder at all (only separators) is refused rather
 * than silently read as the Inbox root.
 */
export function folderOption(folder: string | undefined): { ref: MailFolderRef; label: string } {
    if (folder === undefined || folder.trim() === '') return {ref: INBOX_REF, label: 'Inbox'};
    if (folderSegments(folder).length === 0) {
        throw new InvalidRequestError(`Folder '${folder}' does not name a folder.`);
    }
    return {ref: mailFolderRef(folder), label: folder.trim()};
}

/**
 * Split a plain-text reply into the sender's own text and the quoted thread
 * below it.
 *
 * This matters more than it looks: the quoted original carries the text of the
 * message being answered, and handing a reader one blob invites a figure from
 * that message to be read back as the sender's own.
 *
 * Splits at the earliest separator Outlook or the common webmail clients emit.
 * If that would leave no new text at all — a bottom-posted reply, or a body
 * that opens on a quote — the split is abandoned and everything is `body`,
 * since losing the sender's actual words is the one outcome worth avoiding.
 */
export function splitQuotedOriginal(text: string): { body: string; quoted: string; separator: string } {
    const lines = text.split(/\r?\n/);
    const whole = {body: text, quoted: '', separator: ''};
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        let separator = '';
        if (/^-{2,}\s*original message\s*-{2,}$/i.test(line)) separator = '-----Original Message-----';
        // Outlook's HTML thread rule, directly above the From:/Sent: block.
        else if (/^_{10,}$/.test(line)) separator = 'Outlook thread rule';
        // "On Tue, Jul 21, 2026 at 9:14 AM Jo Doe <jo@example.com> wrote:"
        else if (/^on\b.{5,300}\bwrote:$/i.test(line)) separator = 'On … wrote:';
        else if (/^>/.test(lines[i])) separator = '> quoted lines';
        // A bare header block: "From: …" followed by another header line.
        else if (/^from:\s*\S/i.test(line) && /^(sent|date|to|subject|cc):/i.test((lines[i + 1] || '').trim())) {
            separator = 'From:/Sent: header block';
        }
        if (!separator) continue;
        const body = lines.slice(0, i).join('\n').trimEnd();
        if (!body.trim()) return whole;
        return {body, quoted: lines.slice(i).join('\n').trim(), separator};
    }
    return whole;
}
