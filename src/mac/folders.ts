// Listing the folders under an Inbox, and filing mail into one.
import {
    AS_LIST_SEP,
    asIdList,
    asInt,
    asRow,
    asString,
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
    MAC_ROOT_TERMS,
    macFolderPath,
    macRefPath,
    mailScopeSnippet,
    messageByIdSnippet,
    ownMailGuardSnippet,
    partitionMessageIds,
    resolveMacAccount,
    rootFolderSnippet,
} from './scripts';
import { MAILBOX_ROOT } from '../mail';
import type { ListFoldersRequest, MoveRequest } from '../backend';
import type { InboxFolderInfo, MoveEmailsResult } from '../types';

/**
 * The folders under an account's Inbox, or from the top of its mailbox (its
 * `root folder`, whose `mail folders` are mail folders alone), with their
 * message counts. Path segments travel joined by the LIST separator and are
 * assembled into a folder path in TypeScript, rather than pushed through two
 * layers of backslash escaping.
 */
export async function listInboxFolders(request: ListFoldersRequest): Promise<InboxFolderInfo[]> {
    const top = request.from === 'mailbox';
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
${rootFolderSnippet(acct, top ? MAC_ROOT_TERMS[MAILBOX_ROOT] : 'inbox', 'inb')}
end tell
return my walkFolders(inb, 1, ${asInt(request.maxDepth)}, "")`, 'standard');
    return splitRecords(raw).map(record => {
        const parts = splitFields(record);
        const segments = splitList(field(parts, 1));
        return {
            name: field(parts, 0),
            folderPath: top
                ? macRefPath(request.account, { rootId: MAILBOX_ROOT, rootLabel: request.account, segments })
                : macFolderPath(request.account, 'Inbox', segments),
            itemCount: intField(parts, 2),
            depth: intField(parts, 3, 1),
        };
    });
}

/** Why the Move dialog can't be reached, for each of the two refusals macOS gives. Plain ASCII: they travel in script literals. */
export const MOVE_DIALOG_REFUSALS = {
    accessibility: "Outlook ignored the scripted move, and filing the email through Outlook's Move menu instead needs "
        + 'Accessibility access: allow the app this runs in (or the terminal or editor it was started from) in '
        + 'System Settings > Privacy & Security > Accessibility (Device Control and Data Access on macOS 27).',
    automation: "Outlook ignored the scripted move, and filing the email through Outlook's Move menu instead needs "
        + 'permission to control System Events: allow it in System Settings > Privacy & Security > Automation.',
} as const;

/**
 * File one email through Outlook's own Message > Move > Choose Folder... dialog,
 * as a person would — for a mailbox where the scripted `move` silently does
 * nothing, as one with no account object does (see profile.ts). Outlook comes
 * to the front while it works, and whatever was in front before is put back.
 *
 * The email is opened in its own window, recognised by its title, which starts
 * with the subject and ends "• <address>" (the window publishes no message to
 * check by). The dialog's search matches folder names only, and lists each match
 * as "<name>, <mailbox>/<path to it>", such as "Processed, me@example.com" for a
 * folder beside the Inbox or "CFL, me@example.com/Inbox" for one in it. Nothing
 * tells which row is selected, so Move is pressed only when the search leaves
 * one row and it is `scopeLabel`; otherwise the dialog is cancelled and the
 * rows it showed are named. Opening the email marks it read.
 *
 * Filed is what the scripted move's check calls filed — the destination grew —
 * or the message now sitting in the destination, since a move through the
 * dialog keeps its id.
 */
const MOVE_DIALOG_HANDLER = `
on fileThroughMoveDialog(theId, subj, leafName, scopeLabel, scopeId, countBefore, acctAddr)
    tell application "Microsoft Outlook"
        try
            if (id of folder of message id theId) is scopeId then return
        end try
    end tell
    set prevApp to missing value
    set winId to missing value
    set dlg to missing value
    try
        tell application "System Events"
            if not (UI elements enabled) then error number -25211
            try
                set prevApp to name of first application process whose frontmost is true
            end try
        end tell
        tell application "Microsoft Outlook"
            open (message id theId)
            activate
        end tell
        set winTitle to ""
        repeat 20 times
            tell application "Microsoft Outlook"
                try
                    set w1 to window 1
                    set nm to (name of w1) as string
                    if (class of w1) is not main window and nm ends with ((character id 8226) & " " & acctAddr) and (subj is "" or nm starts with subj) then
                        set winId to id of w1
                        set winTitle to nm
                    end if
                end try
            end tell
            if winId is not missing value then exit repeat
            delay 0.25
        end repeat
        if winId is missing value then error "Outlook opened no window for the email, so it could not be filed through its Move menu."
        tell application "System Events" to tell process "Microsoft Outlook"
            set frontmost to true
            delay 0.3
            if ((name of window 1) as string) is not winTitle then error "Another Outlook window came to the front, so the email was not filed through its Move menu."
            click menu item "Choose Folder..." of menu 1 of menu item "Move" of menu 1 of menu bar item "Message" of menu bar 1
            repeat 20 times
                try
                    set dlg to sheet 1 of window 1
                    exit repeat
                end try
                delay 0.25
            end repeat
            if dlg is missing value then error "Outlook's Move dialog did not open."
            set value of text field 1 of dlg to leafName
            -- The list fills as Outlook searches: read it until it holds still.
            set shown to {}
            set lastShown to missing value
            repeat 24 times
                delay 0.25
                set shown to {}
                repeat with g in (UI elements of scroll area 1 of dlg)
                    try
                        if (role of g) is "AXGroup" then set end of shown to (description of g) as string
                    end try
                end repeat
                if (count of shown) > 0 and shown is lastShown then exit repeat
                set lastShown to shown
            end repeat
            if shown is not {scopeLabel} then
                if (count of shown) is 0 then
                    error "Outlook's Move dialog found no folder named '" & leafName & "'."
                end if
                if (count of shown) is 1 then
                    error "Outlook's Move dialog finds only '" & (item 1 of shown) & "' for '" & leafName & "', not '" & scopeLabel & "', so the email was not filed."
                end if
                error "Outlook's Move dialog finds more than one folder for '" & leafName & "' (" & (my joinList(shown, "; ")) & ") and does not say which it would file into, so the email was not filed. A folder no other in the mailbox shares a name with can be filed into."
            end if
            click button "Move" of dlg
            set dlg to missing value
        end tell
        set landed to false
        repeat 20 times
            tell application "Microsoft Outlook"
                try
                    if (count of messages of mail folder id scopeId) > countBefore then set landed to true
                end try
                try
                    if (id of folder of message id theId) is scopeId then set landed to true
                end try
            end tell
            if landed then exit repeat
            delay 0.25
        end repeat
        if not landed then error "Outlook's Move dialog closed, but '" & leafName & "' did not gain the email - not filed."
    on error errText number errNum
        my leaveMoveDialog(dlg, winId, prevApp)
        if errNum is -25211 or errNum is -1719 then error ${asString(MOVE_DIALOG_REFUSALS.accessibility)}
        if errNum is -1743 then error ${asString(MOVE_DIALOG_REFUSALS.automation)}
        error errText
    end try
    my leaveMoveDialog(missing value, winId, prevApp)
end fileThroughMoveDialog

on leaveMoveDialog(dlg, winId, prevApp)
    if dlg is not missing value then
        try
            tell application "System Events" to click button "Cancel" of dlg
        end try
    end if
    if winId is not missing value then
        try
            tell application "Microsoft Outlook" to close window id winId
        end try
    end if
    if prevApp is not missing value and prevApp is not "Microsoft Outlook" then
        try
            tell application "System Events" to set frontmost of application process prevApp to true
        end try
    end if
end leaveMoveDialog
`;

/**
 * File emails into a folder of the account. Moving rewrites each message's id,
 * so the folder's ids are read before each move: its new id is that of the one
 * message the folder gained with its subject, and '' when there is no such
 * message or more than one, so another message landing there meanwhile is never
 * taken for it. Filed through the Move dialog, it keeps its id.
 *
 * Outlook for Mac's `move` reports success even where it silently does nothing,
 * so each move is confirmed by the destination growing. One that didn't is
 * filed through Outlook's Move dialog instead (see MOVE_DIALOG_HANDLER), and
 * fails as its own item if that can't file it either. A count compared once
 * around the whole batch would let one real move cover for silent failures,
 * and give no per-item results when it tripped. A message already in the
 * destination counts as moved without being moved, as Windows' Move() to its
 * own folder does.
 *
 * The dialog names a folder by its mailbox and the folders above it, so
 * `scopeLabel` is built from the destination's own chain of folders, the
 * nearest first, as the dialog would show it.
 *
 * Only the account's own mail is moved. A message id resolves against the whole
 * application, so each one is first shown to sit in this account — as Windows
 * resolves every id in the account's store — and one that doesn't fails as its
 * own item rather than being filed out of another mailbox.
 */
export async function moveOutlookEmails(request: MoveRequest): Promise<MoveEmailsResult> {
    const folderPath = macRefPath(request.account, request.folder);
    const { valid, invalid } = partitionMessageIds(request.entryIds);
    if (valid.length === 0) return { folderPath, folderCreated: false, moved: 0, movedEmails: [], failed: invalid };
    const acct = await resolveMacAccount(request.account);
    const raw = await runOsaScript(`${MOVE_DIALOG_HANDLER}
tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${accountRootIdSnippet(acct)}
${mailScopeSnippet(acct, request.folder, request.folderLabel, request.createIfMissing)}
    set scopeId to id of scopeFolder
    set theFolder to scopeFolder
${folderChainSnippet('    ')}
    set scopeLeaf to item 1 of chainNames
    set scopeLabel to scopeLeaf & ", " & ${asString(request.account)}
    repeat with i from (count of chainNames) to 2 by -1
        set scopeLabel to scopeLabel & "/" & (item i of chainNames)
    end repeat
    set movedCount to 0
    set movedIds to {}
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
            set newId to (contents of theId) as string
            if fromId is not scopeId then
                set idsBefore to id of every message of scopeFolder
                set countBefore to count of idsBefore
                move theMsg to scopeFolder
                if (count of messages of scopeFolder) is not greater than countBefore then
                    my fileThroughMoveDialog((contents of theId), subj, scopeLeaf, scopeLabel, scopeId, countBefore, ${asString(request.account)})
                end if
                -- The dialog keeps the message's id; a scripted move rewrites it, and
                -- then its id is the one the folder gained.
                set keptId to false
                try
                    if (id of folder of message id (contents of theId)) is scopeId then set keptId to true
                end try
                if not keptId then
                    set newId to ""
                    -- Bound before the loop, so it walks a local list rather than asking
                    -- Outlook for each item in turn.
                    set idsAfter to id of every message of scopeFolder
                    -- Gained, and with its subject, even when it is the only one gained:
                    -- another message landing there meanwhile is never taken for it.
                    set sameSubject to {}
                    repeat with anId in idsAfter
                        if idsBefore does not contain (contents of anId) then
                            try
                                if ((subject of message id (contents of anId)) as string) is subj then set end of sameSubject to (contents of anId)
                            end try
                        end if
                    end repeat
                    if (count of sameSubject) is 1 then set newId to (item 1 of sameSubject) as string
                end if
            end if
            set movedCount to movedCount + 1
            set end of movedIds to ((contents of theId) as string) & ":" & newId
        on error errText
            set out to out & ${ITEM_FAILURE_ROW}
        end try
    end repeat
    return ${asRow(['(movedCount as string)', '(scopeCreated as string)', `(my joinList(movedIds, ${AS_LIST_SEP}))`])} & out
end tell`, 'scan');
    const { summary, failed } = splitBatch(raw);
    return {
        folderPath,
        folderCreated: boolField(summary, 1),
        moved: intField(summary, 0),
        movedEmails: splitList(field(summary, 2)).map(pair => {
            const at = pair.indexOf(':');
            return { entryId: pair.slice(0, at), newEntryId: pair.slice(at + 1) };
        }),
        failed: [...invalid, ...failed],
    };
}
