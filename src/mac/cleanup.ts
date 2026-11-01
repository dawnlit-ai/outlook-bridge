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
    summaryFields,
} from './run';
import {
    accountLookupSnippet,
    folderChainSnippet,
    macFolderPath,
    partitionMessageIds,
    resolveMacAccount,
    rootFolderSnippet,
} from './scripts';
import { deleteMailResult, deleteStatus, PROTECTED_MAIL_REASON } from '../mail';
import type { DeleteEmailsRequest, PurgeRequest } from '../backend';
import type { DeleteMailOutcome, DeleteMailResult, PurgeDeletedItemsResult } from '../types';

/**
 * Delete mail by message id from anywhere in the account, into Deleted Items.
 * Inbox and Sent Items (and their subfolders) are refused unless
 * `allowProtected`; `dryRun` reports without deleting; every outcome names the
 * subject and folder of the item that actually resolved.
 */
export async function deleteOutlookEmails(request: DeleteEmailsRequest): Promise<DeleteMailResult> {
    const { valid, invalid } = partitionMessageIds(request.entryIds);
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
            -- The chain's top is what extends the refusal to Inbox and Sent Items SUBFOLDERS.
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
                status: deleteStatus(field(parts, 3)),
                reason: field(parts, 4),
            });
        }
    }
    return deleteMailResult(request.dryRun, items);
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
        return { id: field(parts, 0), keep: boolField(parts, 1) };
    });
    const doomed = rows.filter(row => !row.keep);
    const kept = rows.length - doomed.length;
    if (request.dryRun || doomed.length === 0) {
        return { folderPath, dryRun: request.dryRun, matched: doomed.length, purged: 0, kept, failed: 0 };
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
