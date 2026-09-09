// The AppleScript fragments more than one macOS operation is built from.
import { asBool, asEscape } from './run';
import { readProfileAccounts } from './profile';
import { InvalidRequestError, NotImplementedError } from '../errors';
import type { MailFolderRef } from '../types';
import { isOutgoingRoot } from '../mail';

/**
 * An account as the generated AppleScript will reach it.
 *
 * `folderIds` carries what Outlook's profile database knows about the mailbox's
 * well-known folders — the only handle there is on a mailbox the dictionary
 * publishes no account object for (see profile.ts). It is filled in for every
 * account the profile lists, including the ones the probe below resolves
 * perfectly well, because the generated script picks between the two at run
 * time. Nothing here has to predict which kinds of account a given Outlook build
 * chooses to expose.
 */
export interface MacAccount {
    readonly emailAccount: string;
    readonly folderIds?: Readonly<Record<string, number>>;
}

/** Look an address up in Outlook's profile database, for the ids it can add. */
export async function resolveMacAccount(emailAccount: string): Promise<MacAccount> {
    const wanted = emailAccount.trim().toLowerCase();
    const found = (await readProfileAccounts())
        .find(account => account.emailAccount.trim().toLowerCase() === wanted);
    return found ? { emailAccount, folderIds: found.folderIds } : { emailAccount };
}

/**
 * Resolve the account whose SMTP address matches, or raise. Emitted into a
 * `tell application "Microsoft Outlook"` block; binds `targetAcct`.
 *
 * `every account` errors even in legacy mode, so probe the typed classes. A
 * `whose email address is ...` filter can't be used either: `email address` is
 * also a class name, and AppleScript resolves it as one ("into type specifier").
 *
 * A mailbox the dictionary publishes no account object for leaves `targetAcct`
 * as `missing value` rather than failing here, provided `rootFolderSnippet` has
 * a folder id to reach it by instead. `needsAccountObject` is for the operations
 * that cannot work without the object itself — composing mail is the whole of
 * that list, because the account a message is sent from can only be set from it,
 * and a message composed without it goes out from whichever mailbox Outlook
 * considers default.
 *
 * The "not found" sentence is the one `classifyRunFailure` recognises to raise
 * `AccountNotFoundError`, so its wording is load-bearing. The sentence for a
 * mailbox that resolved but cannot compose deliberately is not: that account was
 * found, and reporting it missing would send the caller hunting the wrong thing.
 */
export function accountLookupSnippet(acct: MacAccount, needsAccountObject = false): string {
    const address = asEscape(acct.emailAccount);
    const probe = `    set targetAcct to missing value
    try
        repeat with a in (exchange accounts & imap accounts & pop accounts)
            if (email address of a as string) is "${address}" then
                set targetAcct to a
                exit repeat
            end if
        end repeat
    end try`;
    if (acct.folderIds && !needsAccountObject) return probe;
    const complaint = acct.folderIds
        ? `Outlook publishes no account object for '${address}', so the account a `
        + `message is sent from cannot be set. Its folders can be read; mail cannot `
        + `be composed from it.`
        : `Account '${address}' not found`;
    return `${probe}
    if targetAcct is missing value then error "${asEscape(complaint)}"`;
}

/**
 * Bind `variable` to one of the account's well-known folders.
 *
 * Two ways in, chosen at run time: through the account object when the probe
 * found one, and by the folder's own id when it didn't. `mail folder id N`
 * resolves against the application rather than through an account, which is what
 * makes the second way possible at all — and N is the same integer Outlook's
 * profile database records for that folder.
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
    // way left in — and it is the one that may have come up empty. Say which
    // folder is out of reach rather than let `<term> of missing value` say it.
    if (acct.folderIds) {
        return `    if targetAcct is missing value then error "Account '${asEscape(acct.emailAccount)}' has no ${term} folder that AppleScript can reach"
${viaAccount}`;
    }
    return viaAccount;
}

/**
 * Outlook for Mac's dictionary term for each well-known root, keyed by the
 * olDefaultFolders id `mailFolderRef` resolves.
 *
 * Verified against the running app. The terms are NOT the ones the Windows ids
 * suggest — `sent mail` and `junk email` do not compile at all, while
 * `sent items` and `junk mail` do, and a bad term fails the whole script at
 * compile time rather than at the offending line.
 */
export const MAC_ROOT_TERMS: Record<number, string> = {
    6: 'inbox',
    5: 'sent items',
    16: 'drafts',
    3: 'deleted items',
    4: 'outbox',
    23: 'junk mail',
};

/** The date property a folder's items actually carry. */
export function dateProperty(rootId: number): string {
    return isOutgoingRoot(rootId) ? 'time sent' : 'time received';
}

/**
 * The Windows `FolderPath` shape (`\\mailbox\Inbox\Invoices`), built here rather
 * than in AppleScript — escaping backslashes through a template literal and then
 * an AppleScript literal is unreadable, and TS already knows the parts.
 */
export function macFolderPath(emailAccount: string, rootLabel: string, segments: string[]): string {
    return ['\\\\' + emailAccount, rootLabel, ...segments].join('\\');
}

/** The dictionary term for a root, or a NOT_IMPLEMENTED naming what was asked. */
export function rootTerm(ref: MailFolderRef, folderLabel = ''): string {
    const term = MAC_ROOT_TERMS[ref.rootId];
    if (!term) {
        const asked = folderLabel ? ` (asked for '${folderLabel}')` : '';
        throw new NotImplementedError(
            `folder root '${ref.rootLabel}'${asked}`,
            `macOS — readable roots are ${Object.values(MAC_ROOT_TERMS).join(', ')}`,
        );
    }
    return term;
}

/**
 * Emit the AppleScript that resolves `scopeFolder` from `targetAcct`, walking a
 * well-known root down through any further path segments. Also binds
 * `scopeCreated`.
 *
 * Folder names are compared with `is`, which is case-insensitive in AppleScript —
 * matching the Windows walk's `-ieq` so 'invoices' finds 'Invoices' on both.
 * A missing segment errors by name instead of falling back to the root: silently
 * returning the Inbox for a caller that scoped to one folder hands back the wrong
 * emails under a name that says otherwise. `createMissing` builds the whole
 * chain instead, so a nested destination is one call.
 */
export function mailScopeSnippet(
    acct: MacAccount,
    ref: MailFolderRef,
    folderLabel = '',
    createMissing = false,
): string {
    const term = rootTerm(ref, folderLabel);
    const walk = ref.segments.map(seg => `
    set foundFolder to missing value
    repeat with sf in (mail folders of scopeFolder)
        if (name of sf as string) is "${asEscape(seg)}" then
            set foundFolder to sf
            exit repeat
        end if
    end repeat
    if foundFolder is missing value then
        if ${asBool(createMissing)} then
            set foundFolder to (make new mail folder at scopeFolder with properties {name:"${asEscape(seg)}"})
            set scopeCreated to true
        else
            error "Folder '${asEscape(seg)}' not found under '${asEscape(ref.rootLabel)}'"
        end if
    end if
    set scopeFolder to foundFolder`).join('');
    return `    set scopeCreated to false
${rootFolderSnippet(acct, term, 'scopeFolder')}${walk}`;
}

/**
 * The message-id text for a script, or an InvalidRequestError.
 *
 * Outlook for Mac ids are small integers; a Windows MAPI EntryID is rejected up
 * front rather than left to fail as a confusing "not found", because a caller
 * that persisted one from a Windows run needs to be told exactly that.
 */
export function macMessageId(entryId: string): string {
    const id = String(entryId ?? '').trim();
    if (!/^\d+$/.test(id)) {
        throw new InvalidRequestError(
            `'${entryId}' is not an Outlook for Mac message id. Mac ids are small integers `
            + `(e.g. "1263") returned by readInboxEmails on this machine; a Windows MAPI `
            + `EntryID cannot be resolved here.`,
        );
    }
    return id;
}

/**
 * Split ids into the ones a script can look up and the ones that can only fail.
 *
 * The per-id operations report failures rather than throwing, so an unusable id
 * belongs in the result's `failed` list beside a genuine lookup miss — not as an
 * exception that discards the ids that would have worked.
 */
export function partitionMessageIds(entryIds: readonly string[]): {
    valid: string[];
    invalid: { entryId: string; error: string }[];
} {
    const valid: string[] = [];
    const invalid: { entryId: string; error: string }[] = [];
    for (const entryId of entryIds) {
        try {
            valid.push(macMessageId(entryId));
        } catch (error) {
            invalid.push({ entryId: String(entryId), error: (error as Error).message });
        }
    }
    return { valid, invalid };
}

/**
 * Resolve `theMsg` from a macOS message id, or raise.
 *
 * `message id N` resolves against the application rather than one folder, so the
 * message is found wherever it currently sits — including a subfolder.
 */
export function messageLookupSnippet(id: string, variable = 'theMsg'): string {
    return `    set ${variable} to missing value
    try
        set ${variable} to message id ${id}
    end try
    if ${variable} is missing value then error "Email not found for message id '${id}'"`;
}

/**
 * The sender's display name and address, as `sndName` / `sndAddr`.
 *
 * `sender` yields a record, and `address of sender of m` fails to coerce, so the
 * record has to be bound first. `name` is absent when there is no display name.
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
 * Everything a reader wants off ONE message already bound to `theMsg`, and the
 * row it produces — emitted together so the read order and the positions the
 * TypeScript decoders index into cannot drift apart.
 *
 * Three callers share it (`readEmailBody`, `readSelectedEmail`, and the
 * attachment resolver) and differ only in how they bind `theMsg`. They used to
 * share it by copy: thirty lines repeated, with the `time received` → `time sent`
 * fallback and the field order restated each time, so a dictionary fix reached
 * one reader and not the others.
 *
 * `bodyText` is emitted LAST when asked for, so a stray separator in an earlier
 * field cannot shift it.
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

/**
 * The fields `messageDetailSnippet` fills, in row order. The decoders read them
 * back by the positions in `MessageDetail`.
 */
export function messageDetailFields(withBody = false): string[] {
    const fields = [
        '(id of theMsg as string)',
        'subj',
        'sndName',
        'sndAddr',
        'recvd',
        'my sanitizeList(attNames)',
    ];
    if (withBody) fields.push('bodyText');
    return fields;
}

/** Where each of those fields lands in the emitted row. */
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
 * The first recipient's name and address, as `sndName` / `sndAddr` — what
 * outgoing mail reports in place of a sender, matching the Windows reader's
 * contract for Sent Items (a folder of mail from yourself is unreadable).
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

/**
 * Every SMTP address a message was addressed to, as the list `addrList`.
 *
 * `every recipient` covers To, CC and BCC — the full set, which is the whole
 * point of readSentRecipientGroups.
 */
export function allRecipientsSnippet(messageVariable = 'theMsg', indent = '    '): string {
    return `${indent}set addrList to {}
${indent}try
${indent}    repeat with r in (every recipient of ${messageVariable})
${indent}        try
${indent}            -- The record has to be bound before its field is read: a nested
${indent}            -- 'address of (email address of r)' does not coerce, and fails
${indent}            -- inside the try as an empty recipient list rather than an error.
${indent}            set ea to email address of r
${indent}            set oneAddr to (address of ea) as string
${indent}            if oneAddr is not "" then set end of addrList to oneAddr
${indent}        end try
${indent}    end repeat
${indent}end try`;
}

/**
 * Recursive folder-by-name search from `startFolder`, case-insensitive and
 * depth-capped — the macOS counterpart of the Windows Find-FolderByName, used to
 * resolve a template folder anywhere in the mailbox tree.
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
 * Account enumeration only works when Outlook runs in legacy mode; New Outlook
 * errors on `every account` and returns empty lists for the typed classes, so
 * each class is probed independently and failures are ignored.
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
