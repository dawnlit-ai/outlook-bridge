// Composing outgoing mail on macOS: a new email, and a reply to an existing one.
import { asString, field, runOsaScript, splitFields } from './run';
import { accountLookupSnippet, macMessageId, messageLookupSnippet, resolveMacAccount } from './scripts';
import { parseRecipient } from '../shared/args';
import type { Disposition, ReplyRequest, SendRequest } from '../backend';
import type { ReplyEmailResult, SendEmailResult } from '../types';

/**
 * What to do with a composed message. A filed draft leaves its id in
 * `filedDraftId` ("" for the other dispositions, or when it can't be told).
 *
 * A new outgoing message lands in Temporary Items, not Drafts — `open` only
 * shows it — so a windowless draft is moved into Drafts explicitly. `move`
 * reports success even where it silently does nothing, so the folder is
 * confirmed to hold a message it didn't before rather than report a draft that
 * doesn't exist.
 *
 * Moving rewrites a message's id, so the draft's id is the one Drafts gained.
 * Another draft can land in the same moment — an autosave of a compose window,
 * a sync — and then the subject tells them apart; a tie is left unnamed rather
 * than guessed, since a caller that keeps the id may later send by it.
 */
function dispose(variable: string, disposition: Disposition): { prelude: string; action: string } {
    switch (disposition) {
        case 'send':
            return {prelude: '    set filedDraftId to ""', action: `    send ${variable}`};
        case 'display':
            return {prelude: '    set filedDraftId to ""', action: `    open ${variable}`};
        case 'save':
            return {
                prelude: `    set filedDraftId to ""
    set draftsFolder to drafts of targetAcct
    set draftIdsBefore to id of every message of draftsFolder`,
                action: `    set filedSubject to ""
    try
        set filedSubject to (subject of ${variable}) as string
    end try
    move ${variable} to draftsFolder
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

/** One `make new … recipient` line per entry, keeping a display name when given. */
function recipientLines(kind: 'to' | 'cc' | 'bcc', entries: readonly string[]): string[] {
    return entries.map(entry => {
        const {name, address} = parseRecipient(entry);
        const emailAddress = name
            ? `{name:${asString(name)}, address:${asString(address)}}`
            : `{address:${asString(address)}}`;
        return `    make new ${kind} recipient at newMsg with properties {email address:${emailAddress}}`;
    });
}

/** Send, display or file a new email; a filed draft reports its id. */
export async function sendOutlookEmail(request: SendRequest): Promise<SendEmailResult> {
    const acct = await resolveMacAccount(request.account);
    const {prelude, action} = dispose('newMsg', request.disposition);
    const lines = [
        ...recipientLines('to', request.to),
        ...recipientLines('cc', request.cc),
        ...recipientLines('bcc', request.bcc),
        ...request.attachments.map(file => `    make new attachment at newMsg with properties {file:POSIX file ${asString(file)}}`),
    ];
    const raw = await runOsaScript(`
tell application "Microsoft Outlook"
${accountLookupSnippet(acct, true)}
${prelude}
    set newMsg to make new outgoing message with properties {subject:${asString(request.subject)}, content:${asString(request.htmlBody)}}
${lines.join('\n')}
    set account of newMsg to targetAcct
${action}
    return filedDraftId
end tell`, 'standard');
    return {draftEntryId: raw.trim()};
}

/**
 * Reply to an email with the composed HTML above the quoted original.
 *
 * `reply to` builds the same quoted body COM's Reply() returns, so the insertion
 * point is the one thing the platforms must agree on (see WORD_SECTION_ANCHOR).
 * `to` comes back as the reply's resolved addresses: Outlook for Mac exposes
 * recipients as records, and the address is the half worth checking.
 */
export async function replyOutlookEmail(request: ReplyRequest): Promise<ReplyEmailResult> {
    const id = macMessageId(request.email.entryId);
    const acct = await resolveMacAccount(request.account);
    const {prelude, action} = dispose('theReply', request.disposition);
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct, true)}
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
    -- its text — both invisible to the caller if tolerated here.
    set account of theReply to targetAcct
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
    return my sanitize(toLine) & (character id 31) & my sanitize(subj) & (character id 31) & my sanitize(repliedTo)
end tell`, 'standard');
    const parts = splitFields(raw);
    return {to: field(parts, 0), subject: field(parts, 1), repliedToSender: field(parts, 2)};
}
