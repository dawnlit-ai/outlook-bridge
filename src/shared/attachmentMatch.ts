// Matching an attachment by name, and the error when nothing matches.
//
// Both are shared because the error is the useful half. Resolving an email by id
// happily returns a DIFFERENT message when the id is stale or wrong — common
// precisely when many replies share one subject — so a bare "not found" points
// the caller at the wrong problem. Naming the email actually resolved, and the
// attachments it does carry, is what makes the mistake visible.

/**
 * Collapse whitespace runs (non-breaking spaces included) and trim, so a
 * trivially reformatted filename still resolves. Both platforms' regex engines
 * treat `\s` as covering U+00A0, which is the character a mail client is most
 * likely to have substituted.
 */
export function normalizeAttachmentName(name: string): string {
    return name.replace(/\s+/g, ' ').trim();
}

/**
 * Index of the attachment matching `wanted`, or -1.
 *
 * Exact first, then whitespace-normalized and case-insensitive — in that order,
 * so an email carrying two names differing only in case still resolves the one
 * that was actually asked for.
 */
export function findAttachmentIndex(names: readonly string[], wanted: string): number {
    const exact = names.indexOf(wanted);
    if (exact !== -1) return exact;
    const normalized = normalizeAttachmentName(wanted).toLowerCase();
    return names.findIndex(name => normalizeAttachmentName(name).toLowerCase() === normalized);
}

/**
 * The "attachment not found" message, worded once for both platforms.
 *
 * Every field arrives already stringified, which is what lets the Windows script
 * build the identical sentence: it passes PowerShell interpolations
 * (`$($item.Subject)`) where macOS passes the values it read.
 */
export function attachmentNotFoundMessage(init: {
    fileName: string;
    senderEmail: string;
    subject: string;
    receivedTime: string;
    /** The attachments the email does carry, already joined for display. */
    present: string;
    /** What this platform calls the id, for the closing advice. */
    idLabel: string;
}): string {
    return `Attachment '${init.fileName}' not found. `
        + `Resolved email: from=${init.senderEmail}; subject=${init.subject}; received=${init.receivedTime}. `
        + `Attachments present: ${init.present}. `
        + `If this is not the email you expected, the ${init.idLabel} is likely wrong or stale - `
        + `re-run readInboxEmails to get a current ${init.idLabel}.`;
}
