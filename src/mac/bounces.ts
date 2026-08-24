// Bounce-backs: finding them, filing them away, and mining them for the
// addresses that failed.
//
// Windows classifies inside the generated PowerShell, where Restrict has already
// narrowed the set server-side. AppleScript has no Restrict, so the folder is
// indexed with bulk property reads and the SAME rules — imported from
// shared/bounceRules — are applied here in TypeScript. Only messages that match
// have their bodies read, which is the property that keeps either version
// affordable on a real inbox.
import { AS_HANDLERS, AS_LIST_SEP, asRow, field, runOsaScript, splitFields, splitList, splitRecords } from './run';
import { accountLookupSnippet, allRecipientsSnippet } from './scripts';
import { bounceReason, failedRecipients } from '../shared/bounceRules';
import type { CleanUndeliverableResult, SentRecipientGroup, UndeliverableEmail } from '../types';

/** Both scans clamp the window the same way; a year is as far back as either goes. */
function scanWindow(daysBack: number): number {
    return Math.max(1, Math.min(365, Math.floor(daysBack)));
}

interface IndexedMessage {
    id: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
}

/**
 * Index one well-known folder's messages within the window: everything the
 * bounce classifier needs, and nothing that costs a per-message event.
 */
async function indexFolder(
    emailAccount: string,
    folderTerm: string,
    days: number,
): Promise<IndexedMessage[]> {
    // Sender is a record, and a bulk `sender of every message` read hands back a
    // list of them that AppleScript then reads locally — so the whole index costs
    // four Apple events rather than one per message.
    const script = `${AS_HANDLERS}
tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
    set scanFolder to ${folderTerm} of targetAcct
    set idList to id of every message of scanFolder
    set subjList to subject of every message of scanFolder
    try
        set timeList to time received of every message of scanFolder
    on error
        set timeList to {}
    end try
    try
        set sndList to sender of every message of scanFolder
    on error
        set sndList to {}
    end try
    set cutoff to (current date) - (${days} * days)
    set out to ""
    repeat with i from 1 to (count of idList)
        set stamp to missing value
        if (count of timeList) is (count of idList) then set stamp to item i of timeList
        if stamp is missing value or stamp is greater than or equal to cutoff then
            set sndName to ""
            set sndAddr to ""
            if (count of sndList) is (count of idList) then
                set snd to item i of sndList
                try
                    set sndAddr to (address of snd) as string
                end try
                try
                    set sndName to (name of snd) as string
                end try
            end if
            set out to out & ${asRow([
        '(item i of idList as string)',
        '(item i of subjList)',
        'sndName',
        'sndAddr',
        'my isoDate(stamp)',
    ])}
        end if
    end repeat
    return out
end tell`;
    return splitRecords(await runOsaScript(script, 300000)).map(record => {
        const parts = splitFields(record);
        return {
            id: field(parts, 0),
            subject: field(parts, 1).trim(),
            senderName: field(parts, 2),
            senderEmail: field(parts, 3),
            receivedTime: field(parts, 4),
        };
    });
}

/** Read the plain-text bodies of the given messages, keyed by id. */
async function readBodies(ids: readonly string[]): Promise<Map<string, string>> {
    if (ids.length === 0) return new Map();
    const script = `${AS_HANDLERS}
tell application "Microsoft Outlook"
    set out to ""
    repeat with theId in {${ids.join(', ')}}
        set theMsg to missing value
        try
            set theMsg to message id theId
        end try
        if theMsg is not missing value then
            set bodyText to ""
            try
                set bodyText to (plain text content of theMsg) as string
            end try
            set out to out & ${asRow(['(theId as string)', 'bodyText'])}
        end if
    end repeat
    return out
end tell`;
    const bodies = new Map<string, string>();
    for (const record of splitRecords(await runOsaScript(script, 300000))) {
        const parts = splitFields(record);
        bodies.set(field(parts, 0), field(parts, 1));
    }
    return bodies;
}

/** Classify an index, then fill each match's failed recipients from its body. */
async function collectBounces(
    indexed: readonly IndexedMessage[],
    accountAddress: string,
): Promise<UndeliverableEmail[]> {
    const matched = indexed
        .map(message => ({message, reason: bounceReason(message)}))
        .filter(entry => entry.reason !== '');
    if (matched.length === 0) return [];
    const bodies = await readBodies(matched.map(entry => entry.message.id));
    return matched.map(({message, reason}) => ({
        entryId: message.id,
        subject: message.subject,
        senderName: message.senderName,
        senderEmail: message.senderEmail,
        receivedTime: message.receivedTime,
        matchedReason: reason,
        failedRecipients: failedRecipients(bodies.get(message.id) || '', accountAddress),
    }));
}

/**
 * Scan an account's inbox for bounce-back / non-delivery messages — mail-daemon
 * and postmaster rejections, and "Message blocked"-style failure notices — and,
 * unless previewing, move each to Deleted Items (recoverable).
 *
 * Classification is deliberately conservative so ordinary mail that merely
 * mentions "delivery" is never caught: an item matches only when its sender
 * fingerprints as a mail-delivery daemon/postmaster or its subject contains a
 * specific bounce phrase. macOS has no MessageClass, so the NDR-report signal
 * Windows can also use is unavailable here — an Exchange NDR still matches on
 * its "Undeliverable:" subject, which is what it carries.
 */
export async function cleanUndeliverableEmails(
    emailAccount: string,
    daysBack = 30,
    dryRun = true,
): Promise<CleanUndeliverableResult> {
    const days = scanWindow(daysBack);
    const indexed = await indexFolder(emailAccount, 'inbox', days);
    const matched = await collectBounces(indexed, emailAccount);

    let deletedCount = 0;
    const failed: { subject: string; error: string }[] = [];
    if (!dryRun && matched.length > 0) {
        const script = `${AS_HANDLERS}
tell application "Microsoft Outlook"
    set deletedCount to 0
    set out to ""
    repeat with theId in {${matched.map(m => m.entryId).join(', ')}}
        set subj to ""
        try
            set theMsg to missing value
            try
                set theMsg to message id theId
            end try
            if theMsg is missing value then error "message is no longer in the mailbox"
            try
                set subj to (subject of theMsg) as string
            end try
            delete theMsg
            set deletedCount to deletedCount + 1
        on error errText
            set out to out & ${asRow(['subj', 'errText'])}
        end try
    end repeat
    return ${asRow(['(deletedCount as string)'])} & out
end tell`;
        const records = splitRecords(await runOsaScript(script, 300000));
        deletedCount = Number.parseInt(field(splitFields(records[0] || ''), 0) || '0', 10) || 0;
        for (const record of records.slice(1)) {
            const parts = splitFields(record);
            failed.push({subject: field(parts, 0), error: field(parts, 1)});
        }
    }

    return {
        account: emailAccount,
        scannedDays: days,
        dryRun,
        matchedCount: matched.length,
        deletedCount,
        matched,
        failed,
    };
}

/**
 * Read-only scan of an account's Inbox (and, when scanDeleted, its Deleted Items)
 * for bounce messages, returning just the deduped set of failed recipient
 * addresses. Uses the same conservative classifier as cleanUndeliverableEmails.
 *
 * Deleted Items is included so this still works after cleanUndeliverableEmails has
 * already filed the bounces there — the blacklist step doesn't depend on running
 * before the cleanup. This never deletes anything.
 */
export async function collectBouncedRecipients(
    emailAccount: string,
    daysBack = 30,
    scanDeleted = true,
): Promise<string[]> {
    const days = scanWindow(daysBack);
    const terms = scanDeleted ? ['inbox', 'deleted items'] : ['inbox'];
    const found = new Set<string>();
    for (const term of terms) {
        const bounces = await collectBounces(await indexFolder(emailAccount, term, days), emailAccount);
        for (const bounce of bounces) {
            for (const address of bounce.failedRecipients) found.add(address);
        }
    }
    return [...found];
}

/**
 * Read the account's Sent Items within the window, returning each mail with the
 * full SMTP address set it was sent to (To + CC + BCC), newest first. This is how
 * the "was every address for this contact tried?" question gets answered: where a
 * blast sends one email per organization addressed to all of its addresses, a sent
 * message's recipient set IS that organization's full address set.
 */
export async function readSentRecipientGroups(
    emailAccount: string,
    daysBack = 30,
    limit = 3000,
): Promise<SentRecipientGroup[]> {
    const days = scanWindow(daysBack);
    const cap = Math.max(1, Math.min(10000, Math.floor(limit)));
    // Pass 1 — index Sent Items on `time sent`, which is the stamp outgoing mail
    // actually carries; indexing it on `time received` yields an empty set that
    // looks exactly like an empty folder.
    const indexScript = `${AS_HANDLERS}
tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
    set sentFolder to sent items of targetAcct
    set idList to id of every message of sentFolder
    try
        set timeList to time sent of every message of sentFolder
    on error
        set timeList to {}
    end try
    set subjList to subject of every message of sentFolder
    set cutoff to (current date) - (${days} * days)
    set out to ""
    repeat with i from 1 to (count of idList)
        set stamp to missing value
        if (count of timeList) is (count of idList) then set stamp to item i of timeList
        if stamp is missing value or stamp is greater than or equal to cutoff then
            set out to out & ${asRow([
        '(item i of idList as string)',
        '(item i of subjList)',
        'my isoDate(stamp)',
    ])}
        end if
    end repeat
    return out
end tell`;

    const indexed = splitRecords(await runOsaScript(indexScript, 300000)).map(record => {
        const parts = splitFields(record);
        return {id: field(parts, 0), subject: field(parts, 1).trim(), sentOn: field(parts, 2)};
    });
    indexed.sort((a, b) => b.sentOn.localeCompare(a.sentOn));
    const chosen = indexed.slice(0, cap);
    if (chosen.length === 0) return [];

    // Pass 2 — recipients, which can only be read per message.
    const detailScript = `${AS_HANDLERS}
tell application "Microsoft Outlook"
    set out to ""
    repeat with theId in {${chosen.map(row => row.id).join(', ')}}
        set theMsg to missing value
        try
            set theMsg to message id theId
        end try
        if theMsg is not missing value then
${allRecipientsSnippet('theMsg', '            ')}
            set out to out & ${asRow([
        '(theId as string)',
        'my sanitizeList(addrList)',
    ])}
        end if
    end repeat
    return out
end tell`;

    const recipientsById = new Map<string, string[]>();
    for (const record of splitRecords(await runOsaScript(detailScript, 300000))) {
        const parts = splitFields(record);
        recipientsById.set(
            field(parts, 0),
            splitList(field(parts, 1)).map(address => address.toLowerCase()),
        );
    }

    const groups: SentRecipientGroup[] = [];
    for (const row of chosen) {
        const recipients = recipientsById.get(row.id) || [];
        if (recipients.length === 0) continue;
        groups.push({entryId: row.id, subject: row.subject, sentOn: row.sentOn, recipients});
    }
    return groups;
}
