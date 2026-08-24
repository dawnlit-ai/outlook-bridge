// Drafts: which ones belong to an account, and sending, listing or deleting them.
//
// Where Windows has to scan two Drafts folders and prove each item's
// SendUsingAccount, macOS asks the account for its own Drafts folder directly —
// `drafts of targetAcct` is already account-scoped, so a draft found there
// belongs to this account by construction. That is the same rule the Windows
// scan reaches by a longer route, not a weaker one.
import { asRow, field, intField, runOsaScript, splitFields, splitList, splitRecords, } from './run';
import { accountLookupSnippet, macFolderPath, partitionMessageIds } from './scripts';
import type { DeleteDraftsResult, ListDraftsResult, SendAllDraftsResult } from '../types';

/** The Drafts folder path this account's drafts are reported under. */
function draftsPath(emailAccount: string): string {
    return macFolderPath(emailAccount, 'Drafts', []);
}

/**
 * List the mail drafts belonging to `emailAccount`, newest first.
 *
 * Bodies are previewed, never returned whole — a templated reply body runs to
 * tens of thousands of characters and a folder's worth would blow the cap.
 */
export async function listOutlookDrafts(
    emailAccount: string,
    limit = 100,
    previewChars = 300,
): Promise<ListDraftsResult> {
    const cap = Math.max(1, Math.floor(limit));
    const preview = Math.max(1, Math.floor(previewChars));
    const script = `tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
    set draftsFolder to drafts of targetAcct
    set out to ""
    repeat with theMsg in (messages of draftsFolder)
        set subj to ""
        try
            set subj to (subject of theMsg) as string
        end try
        set toNames to {}
        set toAddrs to {}
        try
            repeat with r in (every to recipient of theMsg)
                try
                    set ea to email address of r
                    set oneAddr to ""
                    set oneName to ""
                    try
                        set oneAddr to (address of ea) as string
                    end try
                    try
                        set oneName to (name of ea) as string
                    end try
                    if oneName is "" then set oneName to oneAddr
                    if oneName is not "" then set end of toNames to oneName
                    if oneAddr is not "" then set end of toAddrs to oneAddr
                end try
            end repeat
        end try
        set bodyText to ""
        try
            set bodyText to (plain text content of theMsg) as string
        end try
        if (length of bodyText) > ${preview} then set bodyText to text 1 thru ${preview} of bodyText
        set attCount to 0
        try
            set attCount to count of attachments of theMsg
        end try
        set modAt to ""
        try
            set modAt to my isoDate(modification date of theMsg)
        end try
        set out to out & ${asRow([
        '(id of theMsg as string)',
        'subj',
        'my joinList(toNames, ", ")',
        'my sanitizeList(toAddrs)',
        'bodyText',
        '(attCount as string)',
        'modAt',
    ])}
    end repeat
    return out
end tell`;

    // One string for every row, rather than rebuilt per draft.
    const folderPath = draftsPath(emailAccount);
    const drafts = splitRecords(await runOsaScript(script, 120000)).map(record => {
        const parts = splitFields(record);
        return {
            entryId: field(parts, 0),
            subject: field(parts, 1),
            to: field(parts, 2),
            toEmails: splitList(field(parts, 3)),
            bodyPreview: field(parts, 4).replace(/\s+/g, ' ').trim(),
            hasAttachments: intField(parts, 5) > 0,
            lastModified: field(parts, 6),
            folderPath,
        };
    });
    // 'yyyy-MM-dd HH:mm' is lexicographically ordered, so a string compare sorts it.
    drafts.sort((a, b) => b.lastModified.localeCompare(a.lastModified));
    return {
        account: emailAccount,
        foldersScanned: [folderPath],
        count: drafts.length,
        truncated: drafts.length > cap,
        drafts: drafts.slice(0, cap),
    };
}

/**
 * Send the mail drafts that belong to `emailAccount`.
 *
 * The ids are snapshotted before the first send: sending moves an item out of
 * Drafts, and mutating the collection mid-enumeration would skip every other one
 * — the same trap the Windows implementation avoids by snapshotting item refs.
 */
export async function sendAllDrafts(emailAccount: string): Promise<SendAllDraftsResult> {
    const script = `tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
    set draftsFolder to drafts of targetAcct
    set wanted to id of every message of draftsFolder
    set sentCount to 0
    set out to ""
    repeat with theId in wanted
        set subj to ""
        try
            set theMsg to missing value
            try
                set theMsg to message id theId
            end try
            if theMsg is missing value then error "draft is no longer in the Drafts folder"
            try
                set subj to (subject of theMsg) as string
            end try
            send theMsg
            set sentCount to sentCount + 1
        on error errText
            set out to out & ${asRow(['subj', 'errText'])}
        end try
    end repeat
    return ${asRow(['(sentCount as string)'])} & out
end tell`;

    const records = splitRecords(await runOsaScript(script, 300000));
    return {
        sent: intField(splitFields(records[0] || ''), 0),
        failed: records.slice(1).map(record => {
            const parts = splitFields(record);
            return { subject: field(parts, 0), error: field(parts, 1) };
        }),
    };
}

/**
 * Delete mail drafts by message id. `delete` files the item in Deleted Items
 * rather than destroying it, so a mistaken call stays recoverable from there —
 * matching the Windows contract.
 *
 * Every id must resolve to a message sitting in THIS account's Drafts folder
 * before anything is deleted; an id pointing at ordinary mail, or at another
 * account's draft, is refused and reported. Without that gate this would be a
 * general-purpose "delete any email by id" tool, which is not what it is for.
 */
export async function deleteOutlookDrafts(
    emailAccount: string,
    entryIds: string[],
): Promise<DeleteDraftsResult> {
    if (entryIds.length === 0) return { deleted: 0, failed: [] };
    const { valid, invalid } = partitionMessageIds(entryIds);
    if (valid.length === 0) return { deleted: 0, failed: invalid };
    const script = `tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
    set draftsFolder to drafts of targetAcct
    set draftsName to (name of draftsFolder) as string
    set draftIds to id of every message of draftsFolder
    set deletedCount to 0
    set out to ""
    repeat with theId in {${valid.join(', ')}}
        try
            set theMsg to missing value
            try
                set theMsg to message id theId
            end try
            if theMsg is missing value then error "no such item in this account's Drafts folder"
            -- Prove it is in THIS account's Drafts before deleting: the folder it
            -- currently sits in, and the account it is bound to, must both match.
            set inFolder to ""
            try
                set inFolder to (name of (folder of theMsg)) as string
            end try
            if inFolder is not draftsName then error "item is not in this account's Drafts folder - refusing to delete"
            if draftIds does not contain (id of theMsg) then error "item is not in this account's Drafts folder - refusing to delete"
            set boundTo to ""
            try
                set boundTo to (email address of (account of theMsg)) as string
            end try
            if boundTo is not "" and boundTo is not (email address of targetAcct) then
                error "draft is not bound to this account - refusing to delete"
            end if
            delete theMsg
            set deletedCount to deletedCount + 1
        on error errText
            set out to out & ${asRow(['(theId as string)', 'errText'])}
        end try
    end repeat
    return ${asRow(['(deletedCount as string)'])} & out
end tell`;

    const records = splitRecords(await runOsaScript(script, 300000));
    return {
        deleted: intField(splitFields(records[0] || ''), 0),
        failed: [
            ...invalid,
            ...records.slice(1).map(record => {
                const parts = splitFields(record);
                return { entryId: field(parts, 0), error: field(parts, 1) };
            }),
        ],
    };
}
