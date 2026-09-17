// Deleting mail, and emptying what was already deleted.
import {
    asBool,
    asIdList,
    asInt,
    asRow,
    asString,
    boolField,
    field,
    intField,
    runOsaScript,
    splitFields,
    splitList,
    splitRecords,
    summaryFields
} from './run';
import {
    accountLookupSnippet,
    macFolderPath,
    partitionMessageIds,
    resolveMacAccount,
    rootFolderSnippet
} from './scripts';
import { PROTECTED_MAIL_REASON } from '../mail';
import type { DeleteEmailsRequest, PurgeRequest } from '../backend';
import type { DeleteMailOutcome, DeleteMailResult, PurgeDeletedItemsResult } from '../types';

const OUTCOME_STATUSES: readonly DeleteMailOutcome['status'][] = ['deleted', 'would-delete', 'refused', 'failed'];

/**
 * Walk a folder's `container` chain to the top as the list `chainNames`,
 * nearest first. Requires `theFolder`.
 *
 * This is what extends the Inbox/Sent Items refusal to their SUBFOLDERS. The
 * chain's last element is the folder directly under the account, and an account
 * can't hold two top-level folders of one name, so comparing that name is exact.
 */
function folderChainSnippet(indent: string): string {
    return `${indent}set chainNames to {}
${indent}set cur to theFolder
${indent}set guard to 0
${indent}repeat while cur is not missing value and guard < 25
${indent}    -- The account's root folder has a name of missing value, and coercing
${indent}    -- that yields the STRING "missing value" — so test before coercing.
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
 * Delete mail by message id from anywhere in the account, into Deleted Items.
 * Inbox and Sent Items (and their subfolders) are refused unless
 * `allowProtected`; `dryRun` reports without deleting; every outcome names the
 * subject and folder of the item that actually resolved.
 */
export async function deleteOutlookEmails(request: DeleteEmailsRequest): Promise<DeleteMailResult> {
    const {valid, invalid} = partitionMessageIds(request.entryIds);
    const items: DeleteMailOutcome[] = invalid.map(bad => ({
        entryId: bad.entryId,
        subject: '',
        folderPath: '',
        status: 'failed',
        reason: bad.error,
    }));
    if (valid.length > 0) {
        const acct = await resolveMacAccount(request.account);
        const raw = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'inbox', 'inboxFolder')}
${rootFolderSnippet(acct, 'sent items', 'sentFolder')}
    set inboxName to (name of inboxFolder) as string
    set sentName to (name of sentFolder) as string
    set out to ""
    repeat with theId in ${asIdList(valid)}
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
            if isProtected and not ${asBool(request.allowProtected)} then
                set out to out & ${asRow(['(theId as string)', 'subj', 'pathText', '"refused"', asString(PROTECTED_MAIL_REASON)])}
            else if ${asBool(request.dryRun)} then
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
end reverseList`, 'scan');

        for (const record of splitRecords(raw)) {
            const parts = splitFields(record);
            const segments = splitList(field(parts, 2));
            items.push({
                entryId: field(parts, 0),
                subject: field(parts, 1).trim(),
                folderPath: segments.length ? macFolderPath(request.account, segments[0], segments.slice(1)) : '',
                status: OUTCOME_STATUSES.find(s => s === field(parts, 3)) ?? 'failed',
                reason: field(parts, 4),
            });
        }
    }
    return {
        dryRun: request.dryRun,
        deleted: items.filter(i => i.status === 'deleted').length,
        refused: items.filter(i => i.status === 'refused').length,
        failed: items.filter(i => i.status === 'failed').length,
        items,
    };
}

/**
 * Permanently remove items from the account's Deleted Items. The folder is
 * indexed first and items destroyed by id afterwards, so deleting can't shift
 * the collection out from under the walk.
 */
export async function purgeDeletedItems(request: PurgeRequest): Promise<PurgeDeletedItemsResult> {
    const folderPath = macFolderPath(request.account, 'Deleted Items', []);
    const acct = await resolveMacAccount(request.account);
    const index = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'deleted items', 'trashFolder')}
    set idList to id of every message of trashFolder
    try
        set timeList to time received of every message of trashFolder
    on error
        set timeList to {}
    end try
    set olderThanDays to ${asInt(request.olderThanDays)}
    set cutoff to (current date) - (olderThanDays * days)
    set out to ""
    repeat with i from 1 to (count of idList)
        set stamp to missing value
        if (count of timeList) is (count of idList) then set stamp to item i of timeList
        set keepIt to false
        if olderThanDays > 0 and stamp is not missing value and stamp is greater than cutoff then set keepIt to true
        set out to out & ${asRow(['(item i of idList as string)', '(keepIt as string)'])}
    end repeat
    return out
end tell`, 'scan');

    const rows = splitRecords(index).map(record => {
        const parts = splitFields(record);
        return {id: field(parts, 0), keep: boolField(parts, 1)};
    });
    const doomed = rows.filter(row => !row.keep);
    const kept = rows.length - doomed.length;
    if (request.dryRun || doomed.length === 0) {
        return {folderPath, dryRun: request.dryRun, matched: doomed.length, purged: 0, kept, failed: 0};
    }

    const summary = summaryFields(await runOsaScript(`tell application "Microsoft Outlook"
    set purgedCount to 0
    set failedCount to 0
    repeat with theId in ${asIdList(doomed.map(row => row.id))}
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
end tell`, 'purge'));
    return {
        folderPath,
        dryRun: request.dryRun,
        matched: doomed.length,
        purged: intField(summary, 0),
        kept,
        failed: intField(summary, 1),
    };
}
