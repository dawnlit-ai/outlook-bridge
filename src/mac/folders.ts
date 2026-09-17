// Listing the folders under an Inbox, and filing mail into one.
import {
    AS_LIST_SEP,
    asIdList,
    asInt,
    asRow,
    boolField,
    field,
    intField,
    runOsaScript,
    splitFields,
    splitList,
    splitRecords
} from './run';
import {
    accountLookupSnippet,
    macFolderPath,
    mailScopeSnippet,
    partitionMessageIds,
    resolveMacAccount,
    rootFolderSnippet,
} from './scripts';
import type { ListFoldersRequest, MoveRequest } from '../backend';
import type { InboxFolderInfo, MoveEmailsResult } from '../types';

/**
 * The folders under an account's Inbox with their message counts. Path segments
 * travel joined by the LIST separator and are assembled into a folder path in
 * TypeScript, rather than pushed through two layers of backslash escaping.
 */
export async function listInboxFolders(request: ListFoldersRequest): Promise<InboxFolderInfo[]> {
    const acct = await resolveMacAccount(request.account);
    const raw = await runOsaScript(`on walkFolders(theFolder, level, maxLevel, prefix)
    set out to ""
    tell application "Microsoft Outlook"
        set subs to mail folders of theFolder
    end tell
    repeat with f in subs
        set nm to ""
        set cnt to 0
        tell application "Microsoft Outlook"
            try
                set nm to (name of f) as string
            end try
            try
                set cnt to (count of messages of f)
            end try
        end tell
        set thisPath to prefix & nm
        set out to out & ${asRow(['nm', 'thisPath', '(cnt as string)', '(level as string)'])}
        if level < maxLevel then
            set out to out & (my walkFolders(f, level + 1, maxLevel, thisPath & ${AS_LIST_SEP}))
        end if
    end repeat
    return out
end walkFolders

tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'inbox', 'inb')}
end tell
return my walkFolders(inb, 1, ${asInt(request.maxDepth)}, "")`, 'standard');
    return splitRecords(raw).map(record => {
        const parts = splitFields(record);
        return {
            name: field(parts, 0),
            folderPath: macFolderPath(request.account, 'Inbox', splitList(field(parts, 1))),
            itemCount: intField(parts, 2),
            depth: intField(parts, 3, 1),
        };
    });
}

/**
 * File emails into a folder of the account. Moving rewrites each message's id.
 *
 * Outlook for Mac's `move` reports success even where it silently does nothing,
 * so the destination's count is compared before and after: a run where every
 * move "succeeded" yet the folder never grew fails rather than reporting mail
 * that was never filed.
 */
export async function moveOutlookEmails(request: MoveRequest): Promise<MoveEmailsResult> {
    const folderPath = macFolderPath(request.account, request.folder.rootLabel, request.folder.segments);
    const {valid, invalid} = partitionMessageIds(request.entryIds);
    if (valid.length === 0) return {folderPath, folderCreated: false, moved: 0, failed: invalid};
    const acct = await resolveMacAccount(request.account);
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${mailScopeSnippet(acct, request.folder, request.folderLabel, request.createIfMissing)}
    set countBefore to count of messages of scopeFolder
    set movedCount to 0
    set out to ""
    repeat with theId in ${asIdList(valid)}
        set subj to ""
        try
            set theMsg to missing value
            try
                set theMsg to message id theId
            end try
            if theMsg is missing value then error "no message with that id in this Outlook"
            try
                set subj to (subject of theMsg) as string
            end try
            move theMsg to scopeFolder
            set movedCount to movedCount + 1
        on error errText
            set out to out & ${asRow(['(theId as string)', 'subj', 'errText'])}
        end try
    end repeat
    set countAfter to count of messages of scopeFolder
    if movedCount > 0 and countAfter is not greater than countBefore then
        error "Outlook reported " & (movedCount as string) & " move(s) but '" & (name of scopeFolder) & "' did not grow - nothing was filed."
    end if
    return ${asRow(['(movedCount as string)', '(scopeCreated as string)'])} & out
end tell`, 'scan');
    const records = splitRecords(raw);
    const summary = splitFields(records[0] ?? '');
    return {
        folderPath,
        folderCreated: boolField(summary, 1),
        moved: intField(summary, 0),
        failed: [
            ...invalid,
            ...records.slice(1).map(record => {
                const parts = splitFields(record);
                return {entryId: field(parts, 0), subject: field(parts, 1), error: field(parts, 2)};
            }),
        ],
    };
}
