// Platform-neutral mail helpers.
//
// These used to live in the Windows implementation, which meant the macOS reader
// imported them from it — the wrong direction, and the reason a folder string or
// a quote split could drift between platforms. Both import them from here now.
import type { MailFolderRef } from './types';

/**
 * Clamp a caller's count to the range an operation actually accepts.
 *
 * These bounds ARE the argument contract — how far back a scan reaches, how many
 * rows come back, how deep a folder walk goes — so they are stated here rather
 * than once per platform, which is how the two came to disagree about what a
 * `daysBack` of 0 means.
 */
export function clamp(value: number, min: number, max: number): number {
    return Math.max(min, Math.min(max, Math.floor(value)));
}

/**
 * The Outlook folders addressable by name rather than by walking from the Inbox,
 * with their olDefaultFolders id. Sent Items is the one that earns this: "has this
 * already gone out?" is otherwise unanswerable, and no amount of Inbox scanning
 * substitutes for it.
 *
 * A well-known name always wins over a user folder of the same name — reach the
 * latter by qualifying it (`Inbox\Drafts`).
 */
export const WELL_KNOWN_FOLDERS: Record<string, number> = {
    inbox: 6,
    'sent items': 5,
    sent: 5,
    'sent mail': 5,
    drafts: 16,
    'deleted items': 3,
    deleted: 3,
    trash: 3,
    'junk email': 23,
    junk: 23,
    outbox: 4,
};

/**
 * Why `deleteOutlookEmails` refuses a message, worded once for both platforms.
 *
 * It reaches the caller verbatim in a per-email `reason`, so the two generated
 * scripts must not word it independently — the same reasoning that put the
 * attachment "not found" sentence in `shared/attachmentMatch.ts`. It also names
 * the option a consumer of THIS package actually passes: the text used to say
 * `allow_protected`, which is the tool-layer spelling and not a name this
 * package exposes.
 */
export const PROTECTED_MAIL_REASON =
    'received or sent mail (Inbox/Sent Items or a subfolder) - pass allowProtected to override';

/**
 * Roots whose items are OUTGOING mail.
 *
 * Outgoing mail carries a sent timestamp where received mail carries a received
 * one, and both platforms have to pick the same folders out on that basis —
 * which is why it lives beside `WELL_KNOWN_FOLDERS` rather than as raw
 * `olDefaultFolders` ids compared inline in each reader.
 */
export function isOutgoingRoot(rootId: number): boolean {
    return rootId === WELL_KNOWN_FOLDERS['sent items']
        || rootId === WELL_KNOWN_FOLDERS.outbox
        || rootId === WELL_KNOWN_FOLDERS.drafts;
}

/**
 * Roots a content scan must never walk INTO, as olDefaultFolders ids.
 *
 * On an Exchange profile these sit beside the Inbox, and a walk from it cannot
 * reach them anyway. On an IMAP profile Outlook maps them UNDER the Inbox — a
 * mailbox whose sent and deleted mail is literally `Inbox\Sent Items` and
 * `Inbox\Deleted Items` — so that same walk hands back mail the operator already
 * sent, drafted or threw away as though it had just arrived.
 *
 * Matched by EntryID off `GetDefaultFolder`, never by folder name: the names are
 * localized (a Chinese profile files them under 已发送邮件 and 已删除邮件), which
 * is precisely the case a name test would miss.
 */
export const NON_INCOMING_ROOTS: number[] = [
    WELL_KNOWN_FOLDERS['deleted items'],
    WELL_KNOWN_FOLDERS.outbox,
    WELL_KNOWN_FOLDERS['sent items'],
    WELL_KNOWN_FOLDERS.drafts,
    WELL_KNOWN_FOLDERS['junk email'],
];

/**
 * How far back a filtered scan reaches when the caller names no window.
 *
 * There is no unbounded mode. A scan reaching back forever is never what a batch
 * run wants — it re-surfaces every pre-alert ever filed — and the one that used
 * to exist was reachable only by accident, depending on whether a given store
 * could answer the subject prefilter. A caller that really wants a year asks for
 * 365.
 */
export const DEFAULT_SCAN_DAYS = 60;

/**
 * Subject prefixes that mark a message as a reply or a forward.
 *
 * A pattern SOURCE rather than a RegExp because it must hold in two engines:
 * macOS tests it as a JavaScript RegExp, Windows concatenates it into a
 * PowerShell `-imatch`, and .NET and JavaScript agree on every construct used
 * here. Non-ASCII is spelled as \uXXXX escapes for the same reason — the Windows
 * script crosses a console codepage that plain ASCII survives unconditionally.
 *
 * The list is not only English. A mailbox run in Chinese threads its replies
 * under 回复: and its forwards under 转发:, so an `RE:|FW:` test reads every one
 * of them as a brand-new message — which is how one shipment's thread becomes a
 * dozen apparent arrivals.
 *
 * The colon may be the full-width one (U+FF1A) a CJK input method produces, and
 * Outlook's reply counter (`RE[2]:`) sits between the prefix and the colon.
 */
const BUILT_IN_REPLY_PREFIXES: readonly string[] = [
    're', 'fwd?', 'aw', 'wg', 'sv', 'vs', 'vb', 'tr', 'res', 'rv', 'enc', 'odp',
    '\\u56DE\\u590D', '\\u56DE\\u8986', '\\u7B54\\u590D', '\\u7B54\\u8986',
    '\\u8F6C\\u53D1', '\\u8F49\\u767C', '\\u8FD4\\u4FE1', '\\u8EE2\\u9001',
    '\\uB2F5\\uC7A5', '\\uC804\\uB2EC',
];

/**
 * One operator-supplied prefix as pattern source: regex metacharacters escaped
 * so it matches as the literal text they typed, and every non-ASCII character
 * spelled \uXXXX so the emitted PowerShell stays pure ASCII whatever the
 * console codepage happens to be.
 */
function literalPrefixSource(text: string): string {
    let out = '';
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        if (code < 0x20 || code > 0x7e) {
            out += '\\u' + code.toString(16).toUpperCase().padStart(4, '0');
        } else if ('\\^$.|?*+()[]{}'.includes(text[i])) {
            out += '\\' + text[i];
        } else {
            out += text[i];
        }
    }
    return out;
}

/**
 * The reply/forward prefix pattern, optionally widened with prefixes of the
 * caller's own — a house convention the built-in list has no way to know, like
 * an 'ACK:' an operator's own team puts on acknowledgements.
 *
 * Caller prefixes are matched as literal text sitting where RE: would sit: the
 * same optional reply counter and the same half- or full-width colon still have
 * to follow, so adding one cannot accidentally match mid-subject.
 */
export function replyPrefixSource(extra: readonly string[] = []): string {
    const own = extra.map(prefix => prefix.trim()).filter(Boolean).map(literalPrefixSource);
    const alternatives = [...BUILT_IN_REPLY_PREFIXES, ...own].join('|');
    return '^\\s*(?:' + alternatives + ')\\s*(?:\\[\\d+\\])?\\s*[:\\uFF1A]';
}

/**
 * The leaf name of an exclude-list entry, for the platforms that can only match
 * a folder by name. A full folder path narrows to its last segment; a bare name
 * is already one.
 */
export function folderLeafName(entry: string): string {
    return entry.trim().split(/[\\\\/]/).filter(Boolean).pop() ?? '';
}

/** The built-in prefixes as one pattern source. */
export const REPLY_PREFIX_SOURCE = replyPrefixSource();
export const REPLY_PREFIX = new RegExp(REPLY_PREFIX_SOURCE, 'i');

/**
 * Resolve a caller's folder string to a well-known root plus the segments below it.
 *
 * Accepts every shape a user has to hand: a bare name ("Invoices"), a relative
 * path ("Invoices\\Paid"), a well-known folder ("Sent Items"), a path under one
 * ("Deleted Items\\2026"), and the full FolderPath listInboxFolders prints
 * ("\\\\team@x.com\\Inbox\\Invoices").
 *
 * An unrecognized first segment means Inbox-relative, which is what keeps a bare
 * subfolder name working unchanged.
 */
export function mailFolderRef(folder: string): MailFolderRef {
    const rest = folder.trim().replace(/^\\\\[^\\/]+[\\/]/, '');
    const segments = rest.split(/[\\/]+/).map(s => s.trim()).filter(Boolean);
    if (segments.length > 0) {
        const head = segments[0].toLowerCase();
        if (Object.prototype.hasOwnProperty.call(WELL_KNOWN_FOLDERS, head)) {
            return { rootId: WELL_KNOWN_FOLDERS[head], rootLabel: segments[0], segments: segments.slice(1) };
        }
    }
    return { rootId: 6, rootLabel: 'Inbox', segments };
}

/**
 * Split a plain-text reply into the sender's own text and the quoted thread
 * below it.
 *
 * This matters more than it looks: the quoted original of a reply carries the
 * text of the message it answers. Handing a caller one blob invites a figure
 * from the outgoing message to be read back as the sender's answer, which is a
 * worse failure than the truncation this exists to fix.
 *
 * Scans for the earliest of the separators Outlook and the common webmail
 * clients emit. If the split would leave no new text at all — a bottom-posted
 * reply, or a body that opens on a quote — it is abandoned and everything is
 * returned as `body`, since dropping the sender's actual words is the one
 * outcome worth avoiding.
 */
export function splitQuotedOriginal(text: string): { body: string; quoted: string; separator: string } {
    const lines = text.split(/\r?\n/);
    const whole = { body: text, quoted: '', separator: '' };
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        let separator = '';
        if (/^-{2,}\s*original message\s*-{2,}$/i.test(line)) separator = '-----Original Message-----';
        // Outlook's HTML thread rule, which sits directly above the From:/Sent: block.
        else if (/^_{10,}$/.test(line)) separator = 'Outlook thread rule';
        // "On Tue, Jul 21, 2026 at 9:14 AM John Doe <j@x.com> wrote:"
        else if (/^on\b.{5,300}\bwrote:$/i.test(line)) separator = 'On … wrote:';
        else if (/^>/.test(lines[i])) separator = '> quoted lines';
        // A bare header block: "From: …" followed by another header line.
        else if (/^from:\s*\S/i.test(line) && /^(sent|date|to|subject|cc):/i.test((lines[i + 1] || '').trim())) {
            separator = 'From:/Sent: header block';
        }
        if (!separator) continue;
        const body = lines.slice(0, i).join('\n').trimEnd();
        if (!body.trim()) return whole;
        return { body, quoted: lines.slice(i).join('\n').trim(), separator };
    }
    return whole;
}
