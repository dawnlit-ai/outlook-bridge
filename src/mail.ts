// Platform-neutral mail helpers.
//
// These used to live in PowerShellService, which meant the macOS reader imported
// them from the Windows implementation — the wrong direction, and the reason a
// folder string or a quote split could drift between platforms. Both services
// import them from here instead.
import type { MailFolderRef } from './types';

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
            return {rootId: WELL_KNOWN_FOLDERS[head], rootLabel: segments[0], segments: segments.slice(1)};
        }
    }
    return {rootId: 6, rootLabel: 'Inbox', segments};
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
    const whole = {body: text, quoted: '', separator: ''};
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
        return {body, quoted: lines.slice(i).join('\n').trim(), separator};
    }
    return whole;
}
