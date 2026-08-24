// Listing the folders under an Inbox, and filing mail into one.
import {
    AS_HANDLERS,
    AS_LIST_SEP,
    asRow,
    field,
    runOsaScript,
    splitFields,
    splitList,
    splitRecords,
} from './run';
import {
    accountLookupSnippet,
    macFolderPath,
    mailScopeSnippet,
    partitionMessageIds,
} from './scripts';
import { mailFolderRef } from '../mail';
import type { InboxFolderInfo, MoveEmailsResult } from '../types';

/**
 * List the folders under an account's Inbox with their item counts.
 *
 * The recursive walk joins path segments with the LIST separator rather than a
 * backslash: a literal backslash would have to survive a JS template literal and
 * then an AppleScript string literal, and the `folderPath` shape is assembled in
 * TypeScript anyway (see macFolderPath).
 */
export async function listInboxFolders(
    emailAccount: string,
    maxDepth = 2,
): Promise<InboxFolderInfo[]> {
    const depth = Math.max(1, Math.min(4, Math.floor(maxDepth)));
    const script = `${AS_HANDLERS}
on walkFolders(theFolder, level, maxLevel, prefix)
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
${accountLookupSnippet(emailAccount)}
    set inb to inbox of targetAcct
end tell
return my walkFolders(inb, 1, ${depth}, "")`;

    return splitRecords(await runOsaScript(script, 120000)).map(record => {
        const parts = splitFields(record);
        const segments = splitList(field(parts, 1));
        return {
            name: field(parts, 0),
            folderPath: macFolderPath(emailAccount, 'Inbox', segments),
            itemCount: Number.parseInt(field(parts, 2) || '0', 10) || 0,
            depth: Number.parseInt(field(parts, 3) || '1', 10) || 1,
        };
    });
}

/**
 * Move emails (by message id) into any folder of the account — a filing folder
 * under the Inbox, or a well-known folder by name (see WELL_KNOWN_FOLDERS).
 *
 * **Moving to "Deleted Items" is how mail gets deleted reversibly**, which is why
 * this takes well-known roots at all: it means the destructive path and the filing
 * path are one tool, and the destructive one is undoable from the folder it lands in.
 *
 * `createIfMissing` builds the WHOLE missing chain, so a nested destination
 * ("Clients\\Acme\\2026") is one call rather than a manual mkdir first.
 * NOTE: moving changes a message's id — the passed ids are dead afterwards.
 *
 * Outlook for Mac's `move` reports success even where it silently does nothing,
 * so the destination's message count is compared before and after. A run where
 * every move "succeeded" and yet the folder never grew fails loudly instead of
 * reporting mail that was never filed.
 */
export async function moveOutlookEmails(
    emailAccount: string,
    entryIds: string[],
    folderName: string,
    createIfMissing = false,
): Promise<MoveEmailsResult> {
    if (entryIds.length === 0) {
        return {folderPath: '', folderCreated: false, moved: 0, failed: []};
    }
    const ref = mailFolderRef(folderName);
    const {valid, invalid} = partitionMessageIds(entryIds);
    const folderPath = macFolderPath(emailAccount, ref.rootLabel, ref.segments);
    if (valid.length === 0) {
        return {folderPath, folderCreated: false, moved: 0, failed: invalid};
    }
    const script = `${AS_HANDLERS}
tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
${mailScopeSnippet(ref, folderName, createIfMissing)}
    set countBefore to count of messages of scopeFolder
    set movedCount to 0
    set out to ""
    repeat with theId in {${valid.join(', ')}}
        try
            set theMsg to missing value
            try
                set theMsg to message id theId
            end try
            if theMsg is missing value then error "no message with that id in this Outlook"
            move theMsg to scopeFolder
            set movedCount to movedCount + 1
        on error errText
            set out to out & ${asRow(['(theId as string)', 'errText'])}
        end try
    end repeat
    set countAfter to count of messages of scopeFolder
    if movedCount > 0 and countAfter is not greater than countBefore then
        error "Outlook reported " & (movedCount as string) & " move(s) but '" & (name of scopeFolder) & "' did not grow - nothing was filed."
    end if
    return ${asRow(['(movedCount as string)', '(scopeCreated as string)'])} & out
end tell`;

    const records = splitRecords(await runOsaScript(script, 300000));
    const summary = splitFields(records[0] || '');
    return {
        folderPath,
        folderCreated: field(summary, 1).trim() === 'true',
        moved: Number.parseInt(field(summary, 0) || '0', 10) || 0,
        failed: [
            ...invalid,
            ...records.slice(1).map(record => {
                const parts = splitFields(record);
                return {entryId: field(parts, 0), error: field(parts, 1)};
            }),
        ],
    };
}
