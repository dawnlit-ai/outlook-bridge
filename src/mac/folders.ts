// Listing the folders under an Inbox, and filing mail into one.
import {
    AS_LIST_SEP,
    asIdList,
    asInt,
    asRow,
    boolField,
    field,
    intField,
    ITEM_FAILURE_ROW,
    runOsaScript,
    splitBatch,
    splitFields,
    splitList,
    splitRecords,
} from './run';
import {
    accountLookupSnippet,
    accountRootIdSnippet,
    folderChainSnippet,
    macFolderPath,
    macRefPath,
    mailScopeSnippet,
    messageByIdSnippet,
    ownMailGuardSnippet,
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
 * so each move is confirmed by the destination growing, and one that didn't
 * fails as its own item. A count compared once around the whole batch would let
 * one real move cover for silent failures, and give no per-item results when it
 * tripped. A message already in the destination counts as moved without being
 * moved, as Windows' Move() to its own folder does.
 *
 * Only the account's own mail is moved. A message id resolves against the whole
 * application, so each one is first shown to sit in this account — as Windows
 * resolves every id in the account's store — and one that doesn't fails as its
 * own item rather than being filed out of another mailbox.
 */
export async function moveOutlookEmails(request: MoveRequest): Promise<MoveEmailsResult> {
    const folderPath = macRefPath(request.account, request.folder);
    const { valid, invalid } = partitionMessageIds(request.entryIds);
    if (valid.length === 0) return { folderPath, folderCreated: false, moved: 0, failed: invalid };
    const acct = await resolveMacAccount(request.account);
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${accountRootIdSnippet(acct)}
${mailScopeSnippet(acct, request.folder, request.folderLabel, request.createIfMissing)}
    set scopeId to id of scopeFolder
    set movedCount to 0
    set out to ""
    repeat with theId in ${asIdList(valid)}
        set subj to ""
        try
${messageByIdSnippet('            ')}
            if theMsg is missing value then error "no message with that id in this Outlook"
            try
                set subj to (subject of theMsg) as string
            end try
            set theFolder to missing value
            try
                set theFolder to folder of theMsg
            end try
${folderChainSnippet('            ')}
${ownMailGuardSnippet(request.account, 'move', '            ')}
            set fromId to missing value
            try
                set fromId to id of theFolder
            end try
            if fromId is not scopeId then
                set countBefore to count of messages of scopeFolder
                move theMsg to scopeFolder
                if (count of messages of scopeFolder) is not greater than countBefore then error "Outlook reported the move but '" & (name of scopeFolder) & "' did not grow - not filed"
            end if
            set movedCount to movedCount + 1
        on error errText
            set out to out & ${ITEM_FAILURE_ROW}
        end try
    end repeat
    return ${asRow(['(movedCount as string)', '(scopeCreated as string)'])} & out
end tell`, 'scan');
    const { summary, failed } = splitBatch(raw);
    return {
        folderPath,
        folderCreated: boolField(summary, 1),
        moved: intField(summary, 0),
        failed: [...invalid, ...failed],
    };
}
