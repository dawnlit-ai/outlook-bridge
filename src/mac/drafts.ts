// Drafts: which ones belong to an account, and listing, sending or deleting them.
//
// Where Windows scans two Drafts folders and proves each item's sending
// account, macOS resolves the one Drafts folder belonging to the account —
// through the account object, or by the folder's own id for a mailbox the
// dictionary won't name (see profile.ts) — so a draft found there belongs to
// the account by construction.
//
// What the folder can't vouch for is an item's own binding, which is why the
// guards compare against the requested address as a literal rather than asking
// `targetAcct`: there may be no account object to ask.
import {
    asIdList,
    asInt,
    asRow,
    asString,
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
    partitionMessageIds,
    resolveMacAccount,
    rootFolderSnippet
} from './scripts';
import type { AccountRequest, EntryIdsRequest, ListDraftsRequest } from '../backend';
import type { DeleteDraftsResult, ItemFailure, ListDraftsResult, SendDraftsResult } from '../types';

/** The account's drafts, newest first. Bodies are previewed, never returned whole. */
export async function listOutlookDrafts(request: ListDraftsRequest): Promise<ListDraftsResult> {
    const acct = await resolveMacAccount(request.account);
    const preview = request.previewChars > 0
        ? `        set bodyText to ""
        try
            set bodyText to (plain text content of theMsg) as string
        end try
        if (length of bodyText) > ${asInt(request.previewChars)} then set bodyText to text 1 thru ${asInt(request.previewChars)} of bodyText`
        : '        set bodyText to ""';
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'drafts', 'draftsFolder')}
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
${preview}
        set attCount to 0
        try
            set attCount to count of attachments of theMsg
        end try
        set modAt to ""
        try
            set modAt to my isoDate(modification date of theMsg)
        end try
        set out to out & ${asRow(['(id of theMsg as string)', 'subj', 'my joinList(toNames, ", ")', 'my sanitizeList(toAddrs)', 'bodyText', '(attCount as string)', 'modAt'])}
    end repeat
    return out
end tell`, 'standard');

    const folderPath = macFolderPath(request.account, 'Drafts', []);
    const drafts = splitRecords(raw).map(record => {
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
    drafts.sort((a, b) => b.lastModified.localeCompare(a.lastModified));
    return {
        account: request.account,
        foldersScanned: [folderPath],
        count: drafts.length,
        truncated: drafts.length > request.limit,
        drafts: drafts.slice(0, request.limit),
    };
}

/**
 * The AppleScript that proves `theMsg` is one of this account's drafts before
 * `verb` happens to it: it sits in the account's Drafts folder, and it is bound
 * to this account or to none.
 */
function ownDraftGuard(emailAccount: string, verb: 'send' | 'delete'): string {
    return `            set inFolder to ""
            try
                set inFolder to (name of (folder of theMsg)) as string
            end try
            if inFolder is not draftsName or draftIds does not contain (id of theMsg) then error "item is not in this account's Drafts folder - refusing to ${verb}"
            set boundTo to ""
            try
                set boundTo to (email address of (account of theMsg)) as string
            end try
            if boundTo is not "" and boundTo is not ${asString(emailAccount)} then error "draft is not bound to this account - refusing to ${verb}"`;
}

/** Run `verb` over named drafts of the account, each proved to be its own first. */
async function actOnDrafts(request: EntryIdsRequest, verb: 'send' | 'delete'): Promise<{
    done: number;
    failed: ItemFailure[]
}> {
    const {valid, invalid} = partitionMessageIds(request.entryIds);
    if (valid.length === 0) return {done: 0, failed: invalid};
    const acct = await resolveMacAccount(request.account);
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'drafts', 'draftsFolder')}
    set draftsName to (name of draftsFolder) as string
    set draftIds to id of every message of draftsFolder
    set doneCount to 0
    set out to ""
    repeat with theId in ${asIdList(valid)}
        set subj to ""
        try
            set theMsg to missing value
            try
                set theMsg to message id theId
            end try
            if theMsg is missing value then error "no such item in this account's Drafts folder"
            try
                set subj to (subject of theMsg) as string
            end try
${ownDraftGuard(request.account, verb)}
            ${verb} theMsg
            set doneCount to doneCount + 1
        on error errText
            set out to out & ${asRow(['(theId as string)', 'subj', 'errText'])}
        end try
    end repeat
    return ${asRow(['(doneCount as string)'])} & out
end tell`, 'scan');
    const records = splitRecords(raw);
    return {
        done: intField(splitFields(records[0] ?? ''), 0),
        failed: [
            ...invalid,
            ...records.slice(1).map(record => {
                const parts = splitFields(record);
                return {entryId: field(parts, 0), subject: field(parts, 1), error: field(parts, 2)};
            }),
        ],
    };
}

/** Send the named drafts. */
export async function sendDrafts(request: EntryIdsRequest): Promise<SendDraftsResult> {
    const {done, failed} = await actOnDrafts(request, 'send');
    return {sent: done, failed};
}

/**
 * Delete the named drafts. `delete` files an item in Deleted Items, so a
 * mistaken call stays recoverable.
 */
export async function deleteOutlookDrafts(request: EntryIdsRequest): Promise<DeleteDraftsResult> {
    const {done, failed} = await actOnDrafts(request, 'delete');
    return {deleted: done, failed};
}

/**
 * Send every draft belonging to the account. The ids are snapshotted before the
 * first send: sending moves an item out of Drafts, and mutating the collection
 * mid-enumeration would skip every other one.
 */
export async function sendAllDrafts(request: AccountRequest): Promise<SendDraftsResult> {
    const acct = await resolveMacAccount(request.account);
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'drafts', 'draftsFolder')}
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
            set out to out & ${asRow(['(theId as string)', 'subj', 'errText'])}
        end try
    end repeat
    return ${asRow(['(sentCount as string)'])} & out
end tell`, 'scan');
    const records = splitRecords(raw);
    return {
        sent: intField(splitFields(records[0] ?? ''), 0),
        failed: records.slice(1).map(record => {
            const parts = splitFields(record);
            return {entryId: field(parts, 0), subject: field(parts, 1), error: field(parts, 2)};
        }),
    };
}
