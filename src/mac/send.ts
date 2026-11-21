// Composing outgoing mail on macOS: a new email, and a reply to an existing one.
import { asRow, asString, field, runOsaScript, summaryFields } from './run';
import {
    accountLookupSnippet,
    type MacAccount,
    macMessageId,
    messageLookupSnippet,
    resolveMacAccount,
    rootFolderSnippet,
} from './scripts';
import { failureTag } from '../errors';
import { parseRecipient } from '../shared/args';
import type { Disposition, ReplyRequest, SendRequest } from '../backend';
import type { ReplyEmailResult, SendEmailResult } from '../types';

/**
 * What to do with a composed message. A filed draft leaves its id in
 * `filedDraftId` ("" for the other dispositions, or when it can't be told).
 * `bind` ties the message to the account it is sent from.
 *
 * A new outgoing message lands in Temporary Items, not Drafts — `open` only
 * shows it — so a windowless draft is moved into Drafts explicitly. `move`
 * reports success even where it silently does nothing, so the folder is
 * confirmed to hold a message it didn't before rather than report a draft that
 * doesn't exist.
 *
 * A mailbox the dictionary publishes no account object for (see profile.ts)
 * can't be bound, and into its Drafts `move` and `duplicate` do nothing at all.
 * Such a draft is filed the way a person files one instead: opened in its
 * compose window and closed with saving, which puts it in the Drafts of the
 * mailbox Outlook sends it from. That mailbox is named nowhere else — `account`
 * is `missing value` and `sender` is only a header Outlook does not send by —
 * but the window's title ends "• <address>" with it, so that is checked first,
 * and a draft another mailbox would send is closed unsaved, deleted and refused.
 *
 * Filing rewrites a message's id, so the draft's id is the one Drafts gained.
 * Another draft can land in the same moment — an autosave of a compose window,
 * a sync — and then the subject tells them apart; a tie is left unnamed rather
 * than guessed, since a caller that keeps the id may later send by it.
 */
function dispose(variable: string, disposition: Disposition, acct: MacAccount): {
    prelude: string;
    bind: string;
    action: string;
} {
    const bind = `    set account of ${variable} to targetAcct`;
    switch (disposition) {
        case 'send':
            return { prelude: '    set filedDraftId to ""', bind, action: `    send ${variable}` };
        case 'display':
            return { prelude: '    set filedDraftId to ""', bind, action: `    open ${variable}` };
        case 'save':
            return {
                prelude: `    set filedDraftId to ""
${rootFolderSnippet(acct, 'drafts', 'draftsFolder')}
    set draftIdsBefore to id of every message of draftsFolder`,
                bind: acct.folderIds ? `    if targetAcct is not missing value then\n    ${bind}\n    end if` : bind,
                action: `    set filedSubject to ""
    try
        set filedSubject to (subject of ${variable}) as string
    end try
${acct.folderIds ? fileThroughWindow(variable, acct.emailAccount) : `    move ${variable} to draftsFolder`}
    -- Bound before the loop, so it walks a local list rather than asking
    -- Outlook for each item in turn.
    set draftIdsAfter to id of every message of draftsFolder
    set newIds to {}
    repeat with anId in draftIdsAfter
        if draftIdsBefore does not contain (contents of anId) then set end of newIds to (contents of anId)
    end repeat
    if (count of newIds) is 0 then error "Outlook did not file the draft in the Drafts folder."
    if (count of newIds) is 1 then
        set filedDraftId to (item 1 of newIds) as string
    else
        set sameSubject to {}
        repeat with anId in newIds
            try
                if ((subject of message id (contents of anId)) as string) is filedSubject then set end of sameSubject to (contents of anId)
            end try
        end repeat
        if (count of sameSubject) is 1 then set filedDraftId to (item 1 of sameSubject) as string
    end if`,
            };
    }
}

/**
 * File `variable` into `draftsFolder` through its compose window, when the probe
 * left `targetAcct` missing; otherwise `move` it, as for any account. See
 * `dispose`. Leaves nothing behind when it fails: a draft this can't file in the
 * right mailbox's Drafts is deleted, so a refused draft can't be sent by hand
 * from the wrong mailbox later.
 *
 * Windows are found by id, not index: opening and closing one renumbers the rest.
 * Saving is not instant, so Drafts is watched for the draft for a few seconds.
 */
function fileThroughWindow(variable: string, emailAccount: string): string {
    const notFrom = `${failureTag('INVALID_REQUEST')}Outlook would send this draft from another mailbox than `
        + `'${emailAccount}', so it was not filed. Its compose window read: `;
    return `    if targetAcct is not missing value then
        move ${variable} to draftsFolder
    else
        set composedId to id of ${variable}
        open ${variable}
        set draftWinId to missing value
        repeat 20 times
            repeat with i from 1 to (count of windows)
                try
                    if (class of window i) is draft window then
                        if (id of (object of window i)) is composedId then set draftWinId to id of window i
                    end if
                end try
            end repeat
            if draftWinId is not missing value then exit repeat
            delay 0.25
        end repeat
        if draftWinId is missing value then
            try
                delete ${variable}
            end try
            error "Outlook opened no compose window for the draft, so it could not be filed. Nothing was created."
        end if
        set winTitle to (name of window id draftWinId) as string
        if winTitle does not end with ((character id 8226) & " " & ${asString(emailAccount)}) then
            close window id draftWinId saving no
            try
                delete ${variable}
            end try
            error ${asString(notFrom)} & winTitle
        end if
        close window id draftWinId saving yes
        repeat 20 times
            if (id of every message of draftsFolder) is not draftIdsBefore then exit repeat
            delay 0.25
        end repeat
        if (id of every message of draftsFolder) is draftIdsBefore then
            try
                delete ${variable}
            end try
        end if
    end if`;
}

/** One `make new … recipient` line per entry, keeping a display name when given. */
function recipientLines(kind: 'to' | 'cc' | 'bcc', entries: readonly string[]): string[] {
    return entries.map(entry => {
        const { name, address } = parseRecipient(entry);
        const emailAddress = name
            ? `{name:${asString(name)}, address:${asString(address)}}`
            : `{address:${asString(address)}}`;
        return `    make new ${kind} recipient at newMsg with properties {email address:${emailAddress}}`;
    });
}

/** Send, display or file a new email; a filed draft reports its id. */
export async function sendOutlookEmail(request: SendRequest): Promise<SendEmailResult> {
    const acct = await resolveMacAccount(request.account);
    const { prelude, bind, action } = dispose('newMsg', request.disposition, acct);
    const lines = [
        ...recipientLines('to', request.to),
        ...recipientLines('cc', request.cc),
        ...recipientLines('bcc', request.bcc),
        ...request.attachments.map(file => `    make new attachment at newMsg with properties {file:POSIX file ${asString(file)}}`),
    ];
    const raw = await runOsaScript(`
tell application "Microsoft Outlook"
${accountLookupSnippet(acct, request.disposition !== 'save')}
${prelude}
    set newMsg to make new outgoing message with properties {subject:${asString(request.subject)}, content:${asString(request.htmlBody)}}
${lines.join('\n')}
${bind}
${action}
    return filedDraftId
end tell`, 'standard');
    return { draftEntryId: raw.trim() };
}

/**
 * Reply to an email with the composed HTML above the quoted original.
 *
 * `reply to` builds the same quoted body COM's Reply() returns, so the insertion
 * point is the one thing the platforms must agree on (see WORD_SECTION_ANCHOR).
 * `to` comes back as the reply's resolved addresses: Outlook for Mac exposes
 * recipients as records, and the address is the half worth checking. A filed
 * draft is named by the id Drafts gained, as a new email's is.
 */
export async function replyOutlookEmail(request: ReplyRequest): Promise<ReplyEmailResult> {
    const id = macMessageId(request.email.entryId);
    const acct = await resolveMacAccount(request.account);
    const { prelude, bind, action } = dispose('theReply', request.disposition, acct);
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct, request.disposition !== 'save')}
${messageLookupSnippet(id)}
    set repliedTo to ""
    try
        -- Bind the record before reading its field: a nested
        -- 'address of (sender of theMsg)' does not coerce.
        set origSender to sender of theMsg
        set repliedTo to (address of origSender) as string
    end try
${prelude}
    set theReply to reply to theMsg opening window false ${request.replyAll ? 'reply to all true' : 'without reply to all'}
    -- Neither of these is wrapped in a try. A swallowed account assignment sends
    -- from the default mailbox, and a swallowed insertion files a reply missing
    -- its text — both invisible to the caller if tolerated here. A mailbox with
    -- no account object to assign is checked when the draft is filed instead.
${bind}
    set content of theReply to my insertAboveQuoted(content of theReply, ${asString(request.html)})
    -- Read the outgoing fields BEFORE acting: sending moves the item, and the
    -- reference can't be read afterwards.
    set toList to {}
    try
        repeat with r in (every to recipient of theReply)
            try
                set ea to email address of r
                set end of toList to (address of ea) as string
            end try
        end repeat
    end try
    set toLine to my joinList(toList, ", ")
    set subj to ""
    try
        set subj to (subject of theReply) as string
    end try
${action}
    return ${asRow(['toLine', 'subj', 'repliedTo', 'filedDraftId'])}
end tell`, 'standard');
    const parts = summaryFields(raw);
    return {
        to: field(parts, 0),
        subject: field(parts, 1),
        repliedToSender: field(parts, 2),
        draftEntryId: field(parts, 3),
    };
}
