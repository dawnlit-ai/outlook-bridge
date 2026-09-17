// Bounce-backs: finding them, clearing them away, and mining them for the
// addresses that failed.
//
// Windows classifies inside PowerShell, after Restrict has narrowed the set.
// AppleScript has no Restrict, so a folder is indexed with bulk property reads
// and the SAME rules (shared/bounceRules) run here in TypeScript. Only matches
// have their bodies read, which keeps either version affordable.
import { asIdList, asInt, asRow, field, intField, runOsaScript, splitFields, splitList, splitRecords } from './run';
import { accountLookupSnippet, allRecipientsSnippet, resolveMacAccount, rootFolderSnippet } from './scripts';
import { bounceReason, failedRecipients } from '../shared/bounceRules';
import type { CleanUndeliverableRequest, CollectBouncesRequest, SentGroupsRequest } from '../backend';
import type { CleanUndeliverableResult, ItemFailure, SentRecipientGroup, UndeliverableEmail } from '../types';

interface IndexedMessage {
    id: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
}

/**
 * Index one well-known folder's messages within the window: everything the
 * classifier needs, in bulk reads — `sender of every message` hands back a list
 * of records AppleScript reads locally, so the whole index is four Apple events.
 */
async function indexFolder(emailAccount: string, folderTerm: string, daysBack: number): Promise<IndexedMessage[]> {
    const acct = await resolveMacAccount(emailAccount);
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, folderTerm, 'scanFolder')}
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
    set cutoff to (current date) - (${asInt(daysBack)} * days)
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
            set out to out & ${asRow(['(item i of idList as string)', '(item i of subjList)', 'sndName', 'sndAddr', 'my isoDate(stamp)'])}
        end if
    end repeat
    return out
end tell`, 'scan');
    return splitRecords(raw).map(record => {
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

/** The plain-text bodies of the given messages, by id. */
async function readBodies(ids: readonly string[]): Promise<Map<string, string>> {
    const bodies = new Map<string, string>();
    if (ids.length === 0) return bodies;
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
    set out to ""
    repeat with theId in ${asIdList(ids)}
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
end tell`, 'scan');
    for (const record of splitRecords(raw)) {
        const parts = splitFields(record);
        bodies.set(field(parts, 0), field(parts, 1));
    }
    return bodies;
}

/** Classify an index, then fill each match's failed recipients from its body. */
async function collectBounces(indexed: readonly IndexedMessage[], accountAddress: string): Promise<UndeliverableEmail[]> {
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
 * Find bounce-backs in the Inbox and, unless dry-running, move each to Deleted
 * Items. macOS has no MessageClass, so the NDR-class signal Windows also uses is
 * unavailable — an Exchange NDR still matches on its "Undeliverable:" subject.
 */
export async function cleanUndeliverableEmails(request: CleanUndeliverableRequest): Promise<CleanUndeliverableResult> {
    const matched = await collectBounces(await indexFolder(request.account, 'inbox', request.daysBack), request.account);
    let deletedCount = 0;
    const failed: ItemFailure[] = [];
    if (!request.dryRun && matched.length > 0) {
        const raw = await runOsaScript(`tell application "Microsoft Outlook"
    set deletedCount to 0
    set out to ""
    repeat with theId in ${asIdList(matched.map(m => m.entryId))}
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
            set out to out & ${asRow(['(theId as string)', 'subj', 'errText'])}
        end try
    end repeat
    return ${asRow(['(deletedCount as string)'])} & out
end tell`, 'scan');
        const records = splitRecords(raw);
        deletedCount = intField(splitFields(records[0] ?? ''), 0);
        for (const record of records.slice(1)) {
            const parts = splitFields(record);
            failed.push({entryId: field(parts, 0), subject: field(parts, 1), error: field(parts, 2)});
        }
    }
    return {
        account: request.account,
        scannedDays: request.daysBack,
        dryRun: request.dryRun,
        matchedCount: matched.length,
        deletedCount,
        matched,
        failed,
    };
}

/** The deduplicated addresses bounce-backs report as failed. Never deletes anything. */
export async function collectBouncedRecipients(request: CollectBouncesRequest): Promise<string[]> {
    const terms = request.includeDeletedItems ? ['inbox', 'deleted items'] : ['inbox'];
    const found = new Set<string>();
    for (const term of terms) {
        const bounces = await collectBounces(await indexFolder(request.account, term, request.daysBack), request.account);
        for (const bounce of bounces) {
            for (const address of bounce.failedRecipients) found.add(address);
        }
    }
    return [...found];
}

/**
 * Sent Items within the window, each message with every address it went to,
 * newest first. Indexed on `time sent` — outgoing mail's own stamp; indexing on
 * `time received` yields an empty set that looks exactly like an empty folder.
 */
export async function readSentRecipientGroups(request: SentGroupsRequest): Promise<SentRecipientGroup[]> {
    const acct = await resolveMacAccount(request.account);
    const indexed = splitRecords(await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'sent items', 'sentFolder')}
    set idList to id of every message of sentFolder
    try
        set timeList to time sent of every message of sentFolder
    on error
        set timeList to {}
    end try
    set subjList to subject of every message of sentFolder
    set cutoff to (current date) - (${asInt(request.daysBack)} * days)
    set out to ""
    repeat with i from 1 to (count of idList)
        set stamp to missing value
        if (count of timeList) is (count of idList) then set stamp to item i of timeList
        if stamp is missing value or stamp is greater than or equal to cutoff then
            set out to out & ${asRow(['(item i of idList as string)', '(item i of subjList)', 'my isoDate(stamp)'])}
        end if
    end repeat
    return out
end tell`, 'scan')).map(record => {
        const parts = splitFields(record);
        return {id: field(parts, 0), subject: field(parts, 1).trim(), sentOn: field(parts, 2)};
    });
    indexed.sort((a, b) => b.sentOn.localeCompare(a.sentOn));
    const chosen = indexed.slice(0, request.limit);
    if (chosen.length === 0) return [];

    // Pass 2 — recipients, which can only be read message by message.
    const recipientsById = new Map<string, string[]>();
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
    set out to ""
    repeat with theId in ${asIdList(chosen.map(row => row.id))}
        set theMsg to missing value
        try
            set theMsg to message id theId
        end try
        if theMsg is not missing value then
${allRecipientsSnippet('theMsg', '            ')}
            set out to out & ${asRow(['(theId as string)', 'my sanitizeList(addrList)'])}
        end if
    end repeat
    return out
end tell`, 'scan');
    for (const record of splitRecords(raw)) {
        const parts = splitFields(record);
        recipientsById.set(field(parts, 0), splitList(field(parts, 1)).map(address => address.toLowerCase()));
    }
    return chosen
        .map(row => ({
            entryId: row.id,
            subject: row.subject,
            sentOn: row.sentOn,
            recipients: recipientsById.get(row.id) ?? []
        }))
        .filter(group => group.recipients.length > 0);
}
