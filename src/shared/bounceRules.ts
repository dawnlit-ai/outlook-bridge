// What counts as a bounce, defined once for both platforms.
//
// Windows classifies inside the generated PowerShell — an item's body is only
// read once a cheaper signal has matched, which keeps a large-inbox scan
// affordable — while macOS classifies here after a bulk property read. Two
// implementations of the RULES is how "what is a bounce" comes to mean two
// things, so every list and limit lives here and the PowerShell is generated
// from them (see windows/bounces.ts).
//
// Deliberately specific, mail-system wording: ordinary mail that merely
// mentions a "delivery" must never match, because the next step is often to
// delete what did.

/** Subject fragments only a mail system writes. Matched case-insensitively. */
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

/**
 * The same structural signal as it arrives on the wire, for macOS: a bounce is
 * a delivery-status notification (RFC 3464) and says so in its own headers.
 * Both markers must be present — `multipart/report` also carries read receipts,
 * and it is the report type that separates a delivery that FAILED from one that
 * was acknowledged.
 *
 * Outlook for Mac publishes no MessageClass in its AppleScript dictionary, so
 * this stands in for it: the message Exchange stamps REPORT.IPM.Note.NDR is the
 * message that arrived carrying these headers.
 *
 * The macOS index tests these in AppleScript, whose `contains` is already
 * case-insensitive, and returns one flag per message rather than the headers
 * themselves (see mac/bounces.ts).
 */
export const DSN_HEADER_MARKERS: readonly string[] = ['multipart/report', 'delivery-status'];

/** Address fragments never reported as a failed recipient — they are the bounce's own author. */
export const DAEMON_SELF_ADDRESSES: readonly string[] = ['mailer-daemon', 'postmaster', 'mail-daemon'];

/** At most this many failed recipients per bounce; a list bounce can name hundreds. */
export const MAX_FAILED_RECIPIENTS = 5;

/** The reasons a bounce is flagged, worded once. */
export const BOUNCE_REASON = {
    ndr: 'Non-delivery report (NDR)',
    daemon: 'From mail-delivery system',
    subjectPhrase: (phrase: string) => `Bounce subject phrase: '${phrase}'`,
} as const;

/**
 * Why this item is a bounce, or '' if it isn't.
 *
 * The structural signal (a non-delivery report) outranks the sender
 * fingerprint, which outranks a subject phrase — so the reason reported is the
 * strongest one that applied, not whichever was tested first.
 */
export function bounceReason(item: {
    subject?: string;
    senderName?: string;
    senderEmail?: string;
    /**
     * True when the item is STRUCTURALLY a non-delivery report, whichever way
     * its platform can see that: its MessageClass on Windows, the
     * delivery-status headers it arrived with on macOS.
     */
    isNonDeliveryReport?: boolean;
}): string {
    if (item.isNonDeliveryReport) return BOUNCE_REASON.ndr;

    const senderEmail = (item.senderEmail || '').toLowerCase();
    if (BOUNCE_DAEMON_ADDRESSES.some(fragment => senderEmail.includes(fragment))) return BOUNCE_REASON.daemon;
    const senderName = (item.senderName || '').toLowerCase();
    if (BOUNCE_DAEMON_NAMES.some(fragment => senderName.includes(fragment))) return BOUNCE_REASON.daemon;
    const subject = (item.subject || '').toLowerCase();
    const phrase = BOUNCE_SUBJECT_PHRASES.find(p => subject.includes(p));
    return phrase ? BOUNCE_REASON.subjectPhrase(phrase) : '';
}

/** The address pattern both platforms scan a bounce body with. */
export const BODY_ADDRESS_PATTERN = '[A-Za-z0-9._%+\\-]+@[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}';

/**
 * The addresses a bounce body names as having failed — best effort, since a
 * bounce body is prose with no agreed structure. The account's own address and
 * the daemon's are dropped: both appear in nearly every bounce and neither is a
 * send that failed.
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
