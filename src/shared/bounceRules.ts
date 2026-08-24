// What counts as a bounce, defined once for both platforms.
//
// Windows classifies inside the generated PowerShell (the item body is only
// touched once a cheaper signal has matched, which is what keeps a large-inbox
// scan affordable) while macOS classifies here in TypeScript after a bulk
// property read. Two implementations of the RULES, though, is how "what is a
// bounce" quietly comes to mean two different things — so the lists live here
// and the PowerShell arrays are generated from them (see windows/bounces.ts).
//
// Deliberately specific, multi-word, mail-system wording: ordinary mail that
// merely mentions a "delivery" must never be flagged, because the caller's next
// move is to delete what this matches.

/** Subject fragments that only a mail system writes. Matched case-insensitively. */
export const BOUNCE_SUBJECT_PHRASES: readonly string[] = [
    'undeliverable',
    'message blocked',
    'mail delivery failed',
    'delivery status notification (failure)',
    'delivery has failed',
    'failure notice',
    'undelivered mail returned to sender',
    'returned mail',
];

/** Fragments of a sender ADDRESS that fingerprint a mail-delivery daemon. */
export const BOUNCE_DAEMON_ADDRESSES: readonly string[] = [
    'mailer-daemon',
    'mailer_daemon',
    'postmaster@',
    'mail-daemon',
    'maildelivery',
];

/** Fragments of a sender DISPLAY NAME that fingerprint a mail-delivery daemon. */
export const BOUNCE_DAEMON_NAMES: readonly string[] = [
    'mail delivery subsystem',
    'mail delivery system',
    'microsoft outlook',
    'postmaster',
    'mailer-daemon',
    'mail administrator',
    'internet mail delivery',
];

/** MessageClass prefix Exchange stamps on a non-delivery report. */
export const NDR_MESSAGE_CLASS_PREFIX = 'REPORT.IPM.Note.NDR';

/** Addresses never reported as a failed recipient — they are the bounce's author. */
const DAEMON_SELF_ADDRESSES = ['mailer-daemon', 'postmaster', 'mail-daemon'];

/** At most this many failed recipients per bounce; a list bounce can name hundreds. */
const MAX_FAILED_RECIPIENTS = 5;

/**
 * Why this item is a bounce, or '' if it isn't.
 *
 * Order matches the Windows script's: the structural signal (an NDR message
 * class) outranks the sender fingerprint, which outranks the subject phrase —
 * so the reason a caller is shown is the strongest one that applied, not the
 * first one that happened to be tested.
 */
export function bounceReason(item: {
    subject?: string;
    senderName?: string;
    senderEmail?: string;
    /** True when the item's MessageClass marks it a non-delivery report (Windows only). */
    isNonDeliveryReport?: boolean;
}): string {
    if (item.isNonDeliveryReport) return 'Non-delivery report (NDR)';

    const senderEmail = (item.senderEmail || '').toLowerCase();
    if (BOUNCE_DAEMON_ADDRESSES.some(fragment => senderEmail.includes(fragment))) {
        return 'From mail-delivery system';
    }
    const senderName = (item.senderName || '').toLowerCase();
    if (BOUNCE_DAEMON_NAMES.some(fragment => senderName.includes(fragment))) {
        return 'From mail-delivery system';
    }
    const subject = (item.subject || '').toLowerCase();
    const phrase = BOUNCE_SUBJECT_PHRASES.find(p => subject.includes(p));
    return phrase ? `Bounce subject phrase: '${phrase}'` : '';
}

/** The address pattern both platforms scan a bounce body with. */
export const BODY_ADDRESS_PATTERN = '[A-Za-z0-9._%+\\-]+@[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}';

/**
 * The addresses a bounce body names as having failed — best effort, since a
 * bounce body is prose with no agreed structure.
 *
 * The account's own address and the daemon's are dropped: both appear in nearly
 * every bounce, and neither is a send that failed.
 */
export function failedRecipients(body: string, accountAddress: string): string[] {
    if (!body) return [];
    const target = (accountAddress || '').toLowerCase();
    const found: string[] = [];
    for (const match of body.matchAll(new RegExp(BODY_ADDRESS_PATTERN, 'g'))) {
        const address = match[0].toLowerCase().replace(/\.+$/, '');
        if (address === target) continue;
        if (DAEMON_SELF_ADDRESSES.some(fragment => address.includes(fragment))) continue;
        if (found.includes(address)) continue;
        found.push(address);
        if (found.length >= MAX_FAILED_RECIPIENTS) break;
    }
    return found;
}
