// Deleting mail, and emptying what was already deleted.
import {
    asEscape,
    asRow,
    boolField,
    field,
    intField,
    runOsaScript,
    splitFields,
    splitList,
    splitRecords,
    summaryFields
} from './run';
import { accountLookupSnippet, macFolderPath, partitionMessageIds } from './scripts';
import { PROTECTED_MAIL_REASON } from '../mail';
import type { DeleteMailOptions, DeleteMailOutcome, DeleteMailResult, PurgeDeletedItemsResult, } from '../types';

/**
 * Walk a folder's `container` chain to the top, as the list `chainNames`
 * (nearest first). Requires `theFolder`.
 *
 * This is what makes the Inbox/Sent Items refusal cover their SUBFOLDERS too: a
 * filed subfolder is still received mail. The chain's last element is the folder
 * directly under the account, and an account cannot hold two top-level folders
 * of one name, so comparing that name is exact rather than a guess.
 */
function folderChainSnippet(indent: string): string {
    return `${indent}set chainNames to {}
${indent}set cur to theFolder
${indent}set guard to 0
${indent}repeat while cur is not missing value and guard < 25
${indent}    -- The account's own root folder has a 'name' of missing value, and coercing
${indent}    -- that yields the STRING "missing value" — which sailed past an empty
${indent}    -- check and became the chain's root, so nothing was ever recognised as
${indent}    -- living under the Inbox. Test before coercing.
${indent}    set rawName to missing value
${indent}    try
${indent}        set rawName to name of cur
${indent}    end try
${indent}    if rawName is missing value then exit repeat
${indent}    set nm to rawName as string
${indent}    if nm is "" then exit repeat
${indent}    set end of chainNames to nm
${indent}    set nextUp to missing value
${indent}    try
${indent}        set nextUp to container of cur
${indent}    end try
${indent}    set cur to nextUp
${indent}    set guard to guard + 1
${indent}end repeat`;
}

/**
 * Delete mail by message id from anywhere in the account. `delete` files each
 * item in Deleted Items, so this is recoverable — purgeDeletedItems is what
 * destroys.
 *
 * ⚠️ This is the one tool here that can reach received mail. Three things hold
 * the line, exactly as on Windows:
 *
 *  - **Inbox and Sent Items are refused by default, INCLUDING their subfolders** —
 *    a filed subfolder is still received mail. `allowProtected` lifts that, and is
 *    the caller explicitly taking responsibility.
 *  - **`dryRun` resolves and reports without deleting**, so the exact subjects and
 *    folders can be shown to the user before anything happens. Use it first.
 *  - **Every outcome echoes the subject and folderPath** of the item actually
 *    resolved, so a wrong id is visible after the fact rather than silent.
 */
export async function deleteOutlookEmails(
    emailAccount: string,
    entryIds: string[],
    options: DeleteMailOptions = {},
): Promise<DeleteMailResult> {
    const {allowProtected = false, dryRun = false} = options;
    if (entryIds.length === 0) {
        return {dryRun, deleted: 0, refused: 0, failed: 0, items: []};
    }
    const {valid, invalid} = partitionMessageIds(entryIds);
    const items: DeleteMailOutcome[] = invalid.map(bad => ({
        entryId: bad.entryId,
        subject: '',
        folderPath: '',
        status: 'failed' as const,
        reason: bad.error,
    }));
    if (valid.length > 0) {
        const script = `tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
    set inboxName to (name of (inbox of targetAcct)) as string
    set sentName to (name of (sent items of targetAcct)) as string
    set out to ""
    repeat with theId in {${valid.join(', ')}}
        set subj to ""
        set pathText to ""
        try
            set theMsg to missing value
            try
                set theMsg to message id theId
            end try
            if theMsg is missing value then error "no item with that message id in this Outlook"
            try
                set subj to (subject of theMsg) as string
            end try
            set theFolder to missing value
            try
                set theFolder to folder of theMsg
            end try
            if theFolder is missing value then error "item has no parent folder - refusing to delete"
${folderChainSnippet('            ')}
            set pathText to my sanitizeList(my reverseList(chainNames))
            set rootName to item (count of chainNames) of chainNames
            set isProtected to (rootName is inboxName) or (rootName is sentName)
            if isProtected and not ${allowProtected} then
                set out to out & ${asRow([
            '(theId as string)',
            'subj',
            'pathText',
            '"refused"',
            `"${asEscape(PROTECTED_MAIL_REASON)}"`,
        ])}
            else if ${dryRun} then
                set out to out & ${asRow(['(theId as string)', 'subj', 'pathText', '"would-delete"', '""'])}
            else
                delete theMsg
                set out to out & ${asRow(['(theId as string)', 'subj', 'pathText', '"deleted"', '""'])}
            end if
        on error errText
            set out to out & ${asRow(['(theId as string)', 'subj', 'pathText', '"failed"', 'errText'])}
        end try
    end repeat
    return out
end tell

on reverseList(lst)
    set out to {}
    repeat with i from (count of lst) to 1 by -1
        set end of out to item i of lst
    end repeat
    return out
end reverseList`;

        for (const record of splitRecords(await runOsaScript(script, 300000))) {
            const parts = splitFields(record);
            const segments = splitList(field(parts, 2));
            items.push({
                entryId: field(parts, 0),
                subject: field(parts, 1).trim(),
                folderPath: segments.length
                    ? macFolderPath(emailAccount, segments[0], segments.slice(1))
                    : '',
                status: (field(parts, 3) || 'failed') as DeleteMailOutcome['status'],
                reason: field(parts, 4),
            });
        }
    }
    return {
        dryRun,
        deleted: items.filter(i => i.status === 'deleted').length,
        refused: items.filter(i => i.status === 'refused').length,
        failed: items.filter(i => i.status === 'failed').length,
        items,
    };
}

/**
 * Permanently remove items from the account's Deleted Items folder. This is the
 * ONE genuinely irreversible operation here — nothing recovers from it — which
 * is why it is folder-scoped rather than keyed on a message id: it can only ever
 * destroy what the user already threw away.
 *
 * `olderThanDays` keeps recent items (0 = purge everything). The folder is
 * indexed first and destroyed by id afterwards, so deleting can't shift the
 * collection out from under the walk.
 */
export async function purgeDeletedItems(
    emailAccount: string,
    olderThanDays = 0,
    dryRun = false,
): Promise<PurgeDeletedItemsResult> {
    const days = Math.max(0, Math.floor(olderThanDays));
    const folderPath = macFolderPath(emailAccount, 'Deleted Items', []);
    const indexScript = `tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
    set trashFolder to deleted items of targetAcct
    set idList to id of every message of trashFolder
    try
        set timeList to time received of every message of trashFolder
    on error
        set timeList to {}
    end try
    set cutoff to (current date) - (${days} * days)
    set out to ""
    repeat with i from 1 to (count of idList)
        set stamp to missing value
        if (count of timeList) is (count of idList) then set stamp to item i of timeList
        set keepIt to false
        if ${days} > 0 and stamp is not missing value and stamp is greater than cutoff then set keepIt to true
        set out to out & ${asRow(['(item i of idList as string)', '(keepIt as string)'])}
    end repeat
    return out
end tell`;

    const rows = splitRecords(await runOsaScript(indexScript, 300000)).map(record => {
        const parts = splitFields(record);
        return {id: field(parts, 0), keep: boolField(parts, 1)};
    });
    const doomed = rows.filter(row => !row.keep);
    const kept = rows.length - doomed.length;
    if (dryRun || doomed.length === 0) {
        return {folderPath, dryRun, matched: doomed.length, purged: 0, kept, failed: 0};
    }

    const purgeScript = `tell application "Microsoft Outlook"
    set purgedCount to 0
    set failedCount to 0
    repeat with theId in {${doomed.map(row => row.id).join(', ')}}
        try
            set theMsg to missing value
            try
                set theMsg to message id theId
            end try
            if theMsg is missing value then error "gone"
            permanently delete theMsg
            set purgedCount to purgedCount + 1
        on error
            set failedCount to failedCount + 1
        end try
    end repeat
    return ${asRow(['(purgedCount as string)', '(failedCount as string)'])}
end tell`;

    const summary = summaryFields(await runOsaScript(purgeScript, 600000));
    return {
        folderPath,
        dryRun,
        matched: doomed.length,
        purged: intField(summary, 0),
        kept,
        failed: intField(summary, 1),
    };
}
