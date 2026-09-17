// The AppleScript fragments more than one macOS operation is built from.
import { asString } from './run';
import { readProfileAccounts } from './profile';
import { failureTag, InvalidRequestError, NotImplementedError } from '../errors';
import { FolderId, isOutgoingRoot } from '../mail';
import type { ItemFailure, MailFolderRef } from '../types';

/**
 * An account as the generated AppleScript reaches it.
 *
 * `folderIds` carries what Outlook's profile database knows about the mailbox's
 * well-known folders — the only handle on a mailbox the dictionary publishes no
 * account object for (see profile.ts). It is filled in for every account the
 * profile lists, because the generated script chooses between the account
 * object and the folder id at run time; nothing here has to predict which kinds
 * of account a given Outlook build exposes.
 */
export interface MacAccount {
    readonly emailAccount: string;
    readonly folderIds?: Readonly<Record<string, number>>;
}

/** Look an address up in Outlook's profile database, for the folder ids it can add. */
export async function resolveMacAccount(emailAccount: string): Promise<MacAccount> {
    const wanted = emailAccount.trim().toLowerCase();
    const found = (await readProfileAccounts())
        .find(account => account.emailAccount.trim().toLowerCase() === wanted);
    return found ? {emailAccount, folderIds: found.folderIds} : {emailAccount};
}

/**
 * Resolve the account whose address matches into `targetAcct`, inside a
 * `tell application "Microsoft Outlook"` block.
 *
 * `every account` errors even in legacy mode, so the typed classes are probed.
 * A `whose email address is …` filter can't be used either: `email address` is
 * also a class name, and AppleScript resolves it as one.
 *
 * A mailbox the dictionary publishes no account object for leaves `targetAcct`
 * as `missing value` rather than failing, provided the profile gave a folder id
 * to reach it by. `needsAccountObject` is for composing mail, which can't work
 * without the object: the account a message is sent from can only be set from
 * it, and a message composed without it goes out from the default mailbox.
 */
export function accountLookupSnippet(acct: MacAccount, needsAccountObject = false): string {
    const probe = `    set targetAcct to missing value
    try
        repeat with a in (exchange accounts & imap accounts & pop accounts)
            if (email address of a as string) is ${asString(acct.emailAccount)} then
                set targetAcct to a
                exit repeat
            end if
        end repeat
    end try`;
    if (acct.folderIds && !needsAccountObject) return probe;
    const complaint = acct.folderIds
        ? `${failureTag('INVALID_REQUEST')}Outlook publishes no account object for '${acct.emailAccount}', `
        + 'so the account a message is sent from cannot be set. Its folders can be read; mail cannot be composed from it.'
        : `${failureTag('ACCOUNT_NOT_FOUND')}Account '${acct.emailAccount}' not found in Outlook.`;
    return `${probe}
    if targetAcct is missing value then error ${asString(complaint)}`;
}

/**
 * Bind `variable` to one of the account's well-known folders: through the
 * account object when the probe found one, else by the folder's own id. `mail
 * folder id N` resolves against the application rather than an account, and N is
 * the id Outlook's profile database records for that folder.
 */
export function rootFolderSnippet(acct: MacAccount, term: string, variable: string): string {
    const viaAccount = `    set ${variable} to ${term} of targetAcct`;
    const id = acct.folderIds?.[term];
    if (id !== undefined) {
        return `    if targetAcct is missing value then
        set ${variable} to mail folder id ${id}
    else
${viaAccount}
    end if`;
    }
    // The profile listed the mailbox but not this root, so the probe is the only
    // way in — and it may have come up empty. Name the folder out of reach
    // rather than let `<term> of missing value` say it.
    if (acct.folderIds) {
        const complaint = `${failureTag('NOT_FOUND', 'folder')}Account '${acct.emailAccount}' has no ${term} folder AppleScript can reach.`;
        return `    if targetAcct is missing value then error ${asString(complaint)}
${viaAccount}`;
    }
    return viaAccount;
}

/**
 * Outlook for Mac's dictionary term for each well-known root, by olDefaultFolders
 * id. Verified against the running app: the terms are NOT what the Windows names
 * suggest — `sent mail` and `junk email` don't compile, `sent items` and `junk
 * mail` do — and one bad term fails the whole script at compile time.
 */
export const MAC_ROOT_TERMS: Readonly<Record<number, string>> = {
    [FolderId.Inbox]: 'inbox',
    [FolderId.SentMail]: 'sent items',
    [FolderId.Drafts]: 'drafts',
    [FolderId.DeletedItems]: 'deleted items',
    [FolderId.Outbox]: 'outbox',
    [FolderId.Junk]: 'junk mail',
};

/** The date property a root's items carry. */
export function dateProperty(rootId: number): string {
    return isOutgoingRoot(rootId) ? 'time sent' : 'time received';
}

/**
 * The Windows folder-path shape (`\\mailbox\Inbox\Invoices`), assembled in
 * TypeScript: escaping backslashes through a template literal and then an
 * AppleScript literal is unreadable, and the parts are known here anyway.
 */
export function macFolderPath(emailAccount: string, rootLabel: string, segments: readonly string[]): string {
    return ['\\\\' + emailAccount, rootLabel, ...segments].join('\\');
}

/** The dictionary term for a root, or NOT_IMPLEMENTED naming what was asked. */
export function rootTerm(ref: MailFolderRef, folderLabel = ''): string {
    const term = MAC_ROOT_TERMS[ref.rootId];
    if (!term) {
        const asked = folderLabel ? ` (asked for '${folderLabel}')` : '';
        throw new NotImplementedError(
            `Reading the '${ref.rootLabel}' folder${asked}`,
            `macOS, whose readable roots are ${Object.values(MAC_ROOT_TERMS).join(', ')}`,
        );
    }
    return term;
}

/**
 * Resolve `scopeFolder` from `targetAcct`: a well-known root walked down through
 * any further segments. Also binds `scopeCreated`.
 *
 * Names compare with `is`, which is case-insensitive in AppleScript, matching
 * the Windows walk. A missing segment fails as NOT_FOUND rather than falling
 * back to the root; `createMissing` builds the chain instead.
 */
export function mailScopeSnippet(acct: MacAccount, ref: MailFolderRef, folderLabel = '', createMissing = false): string {
    const term = rootTerm(ref, folderLabel);
    const walk = ref.segments.map(segment => {
        const complaint = `${failureTag('NOT_FOUND', 'folder')}Folder '${segment}' not found under '${ref.rootLabel}'.`;
        const create = createMissing
            ? `        set foundFolder to (make new mail folder at scopeFolder with properties {name:${asString(segment)}})
        set scopeCreated to true`
            : `        error ${asString(complaint)}`;
        return `
    set foundFolder to missing value
    repeat with sf in (mail folders of scopeFolder)
        if (name of sf as string) is ${asString(segment)} then
            set foundFolder to sf
            exit repeat
        end if
    end repeat
    if foundFolder is missing value then
${create}
    end if
    set scopeFolder to foundFolder`;
    }).join('');
    return `    set scopeCreated to false
${rootFolderSnippet(acct, term, 'scopeFolder')}${walk}`;
}

/**
 * The message id text for a script, or INVALID_REQUEST.
 *
 * Outlook for Mac ids are small integers; a Windows EntryID is refused up front
 * rather than left to fail as a confusing "not found", because a caller that
 * persisted one from a Windows run needs to be told exactly that.
 */
export function macMessageId(entryId: string): string {
    const id = String(entryId ?? '').trim();
    if (!/^\d+$/.test(id)) {
        throw new InvalidRequestError(
            `'${entryId}' is not an Outlook for Mac message id. Mac ids are small integers `
            + '(e.g. "1263") listed on this machine; a Windows EntryID cannot be resolved here.',
        );
    }
    return id;
}

/**
 * Split ids into those a script can look up and those that can only fail. A
 * batch reports an unusable id beside a genuine miss rather than throwing and
 * discarding the ids that would have worked.
 */
export function partitionMessageIds(entryIds: readonly string[]): { valid: string[]; invalid: ItemFailure[] } {
    const valid: string[] = [];
    const invalid: ItemFailure[] = [];
    for (const entryId of entryIds) {
        try {
            valid.push(macMessageId(entryId));
        } catch (error) {
            invalid.push({entryId: String(entryId), subject: '', error: (error as Error).message});
        }
    }
    return {valid, invalid};
}

/**
 * Resolve `variable` from a message id, or fail as NOT_FOUND. `message id N`
 * resolves against the application, so the message is found wherever it sits.
 */
export function messageLookupSnippet(id: string, variable = 'theMsg'): string {
    const complaint = `${failureTag('NOT_FOUND', 'email')}No email found for message id '${id}'. It may have been deleted - list the mail again for a current id.`;
    return `    set ${variable} to missing value
    try
        set ${variable} to message id ${id}
    end try
    if ${variable} is missing value then error ${asString(complaint)}`;
}

/**
 * The sender's display name and address as `sndName` / `sndAddr`. `sender` is a
 * record, and `address of sender of m` doesn't coerce, so the record is bound
 * first.
 */
export function senderSnippet(messageVariable = 'theMsg', indent = '    '): string {
    return `${indent}set sndName to ""
${indent}set sndAddr to ""
${indent}try
${indent}    set snd to sender of ${messageVariable}
${indent}    try
${indent}        set sndAddr to (address of snd) as string
${indent}    end try
${indent}    try
${indent}        set sndName to (name of snd) as string
${indent}    end try
${indent}end try`;
}

/**
 * Everything a reader wants off one message bound to `theMsg`, emitted together
 * with the row it produces (see `messageDetailFields`) so the read order and the
 * positions the decoders index into cannot drift apart. `bodyText` is emitted
 * LAST when asked for, so a stray separator in an earlier field can't shift it.
 */
export function messageDetailSnippet(withBody = false, indent = '    '): string {
    return `${indent}set subj to ""
${indent}try
${indent}    set subj to (subject of theMsg) as string
${indent}end try
${senderSnippet('theMsg', indent)}
${indent}set recvd to ""
${indent}try
${indent}    set recvd to my isoDate(time received of theMsg)
${indent}on error
${indent}    try
${indent}        set recvd to my isoDate(time sent of theMsg)
${indent}    end try
${indent}end try
${indent}set attNames to {}
${indent}try
${indent}    set attNames to name of every attachment of theMsg
${indent}end try` + (withBody ? `
${indent}set bodyText to ""
${indent}try
${indent}    set bodyText to (plain text content of theMsg) as string
${indent}end try` : '');
}

/** The fields `messageDetailSnippet` fills, in row order — see `MessageDetail`. */
export function messageDetailFields(withBody = false): string[] {
    const fields = ['(id of theMsg as string)', 'subj', 'sndName', 'sndAddr', 'recvd', 'my sanitizeList(attNames)'];
    if (withBody) fields.push('bodyText');
    return fields;
}

/** Where each of those fields lands in the row. */
export const MessageDetail = {
    id: 0,
    subject: 1,
    senderName: 2,
    senderEmail: 3,
    receivedTime: 4,
    attachmentNames: 5,
    body: 6,
} as const;

/**
 * The first recipient's name and address as `sndName` / `sndAddr` — what
 * outgoing mail reports in place of a sender.
 */
export function firstRecipientSnippet(messageVariable = 'theMsg', indent = '    '): string {
    return `${indent}set sndName to ""
${indent}set sndAddr to ""
${indent}try
${indent}    set rcps to to recipients of ${messageVariable}
${indent}    if (count of rcps) > 0 then
${indent}        set ea to email address of (item 1 of rcps)
${indent}        try
${indent}            set sndAddr to (address of ea) as string
${indent}        end try
${indent}        try
${indent}            set sndName to (name of ea) as string
${indent}        end try
${indent}    end if
${indent}end try`;
}

/** Every address a message was sent to — To, CC and BCC — as the list `addrList`. */
export function allRecipientsSnippet(messageVariable = 'theMsg', indent = '    '): string {
    return `${indent}set addrList to {}
${indent}try
${indent}    repeat with r in (every recipient of ${messageVariable})
${indent}        try
${indent}            -- The record is bound before its field is read: a nested
${indent}            -- 'address of (email address of r)' doesn't coerce, and fails
${indent}            -- inside the try as an empty list rather than an error.
${indent}            set ea to email address of r
${indent}            set oneAddr to (address of ea) as string
${indent}            if oneAddr is not "" then set end of addrList to oneAddr
${indent}        end try
${indent}    end repeat
${indent}end try`;
}

/**
 * Recursive, case-insensitive, depth-capped folder search from `startFolder` —
 * the macOS counterpart of the Windows Find-FolderByName.
 */
export const FIND_FOLDER_HANDLER = `
on findFolderByName(startFolder, wantedName, depth)
    tell application "Microsoft Outlook"
        set subs to mail folders of startFolder
    end tell
    repeat with f in subs
        if (name of f as string) is wantedName then return f
    end repeat
    if depth is less than or equal to 1 then return missing value
    repeat with f in subs
        set hit to my findFolderByName(f, wantedName, depth - 1)
        if hit is not missing value then return hit
    end repeat
    return missing value
end findFolderByName
`;

/**
 * Enumerate the accounts, probing each class separately: legacy Outlook errors
 * on `every account`, and New Outlook returns empty lists for the classes.
 */
export const LIST_ACCOUNTS_SNIPPET = `
set acctList to {}
tell application "Microsoft Outlook"
    repeat with acctClass in {"exchange", "imap", "pop"}
        try
            if (acctClass as string) is "exchange" then
                set accts to exchange accounts
            else if (acctClass as string) is "imap" then
                set accts to imap accounts
            else
                set accts to pop accounts
            end if
            repeat with a in accts
                try
                    set end of acctList to (email address of a as string)
                end try
            end repeat
        end try
    end repeat
end tell
`;
