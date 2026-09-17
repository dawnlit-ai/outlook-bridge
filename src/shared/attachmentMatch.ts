// Matching an attachment by name, the error when nothing matches, and the
// file name it is saved under.
//
// The error is the useful half. Resolving an email by id can return a
// DIFFERENT message when the id is stale or wrong — common exactly when many
// replies share one subject — so a bare "not found" points at the wrong
// problem. Naming the email that actually resolved, and the attachments it does
// carry, is what makes the mistake visible.
import path from 'path';

/**
 * Collapse whitespace runs (non-breaking spaces included) and trim, so a
 * trivially reformatted file name still resolves. Both platforms' regex engines
 * count U+00A0 as whitespace.
 */
export function normalizeAttachmentName(name: string): string {
    return name.replace(/\s+/g, ' ').trim();
}

/**
 * Index of the attachment matching `wanted`, or -1: exact first, then
 * whitespace-normalized and case-insensitive — in that order, so two names that
 * differ only in case still resolve the one asked for.
 */
export function findAttachmentIndex(names: readonly string[], wanted: string): number {
    const exact = names.indexOf(wanted);
    if (exact !== -1) return exact;
    const normalized = normalizeAttachmentName(wanted).toLowerCase();
    return names.findIndex(name => normalizeAttachmentName(name).toLowerCase() === normalized);
}

/**
 * The "attachment not found" message, worded once for both platforms. Fields
 * arrive as strings so the Windows script can build the same sentence from
 * PowerShell variables where macOS passes the values it read.
 */
export function attachmentNotFoundMessage(init: {
    fileName: string;
    senderEmail: string;
    subject: string;
    receivedTime: string;
    /** The attachments the email does carry, joined for display. */
    present: string;
}): string {
    return `Attachment '${init.fileName}' not found. `
        + `Resolved email: from=${init.senderEmail}; subject=${init.subject}; received=${init.receivedTime}. `
        + `Attachments present: ${init.present}. `
        + 'If this is not the email you expected, its id is likely wrong or stale - list the mail again for a current one.';
}

/** Characters no file name may carry on Windows (a superset of macOS's). */
export const UNSAFE_FILE_NAME_CHARS = '[<>:"/\\\\|?*\\x00-\\x1F]';

/** Names Windows reserves for devices, with or without an extension. */
export const RESERVED_FILE_NAME = '^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\\..*)?$';

/** How long a saved name's stem may run before it is cut. */
export const MAX_FILE_STEM = 150;

/**
 * A sender-chosen attachment name made safe to write: no path separators or
 * characters a file system rejects, no trailing dots or spaces, no reserved
 * device name, and a stem short enough to stay clear of path-length limits. A
 * name like `..\..\x.exe` therefore lands inside the destination as
 * `.._.._x.exe`, never outside it.
 */
export function safeFileName(name: string): string {
    let safe = name.replace(new RegExp(UNSAFE_FILE_NAME_CHARS, 'g'), '_').replace(/[. ]+$/, '');
    if (new RegExp(RESERVED_FILE_NAME, 'i').test(safe)) safe = `_${safe}`;
    const ext = path.extname(safe);
    const stem = safe.slice(0, safe.length - ext.length);
    if (stem.length > MAX_FILE_STEM) safe = stem.slice(0, MAX_FILE_STEM) + ext;
    return safe || 'attachment';
}

/**
 * A path in `dir` for `name` that `exists` reports free: the name itself, else
 * `name (1).ext`, `name (2).ext`, … Existing files are never overwritten.
 */
export function uniqueSavePath(dir: string, name: string, exists: (file: string) => boolean): string {
    const ext = path.extname(name);
    const stem = name.slice(0, name.length - ext.length);
    let candidate = path.join(dir, name);
    for (let n = 1; exists(candidate); n++) {
        candidate = path.join(dir, `${stem} (${n})${ext}`);
    }
    return candidate;
}
