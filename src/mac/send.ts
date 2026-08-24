// Composing outgoing mail on macOS: a new email, and a reply to an existing one.
import { asEscape, field, runOsaScript, splitFields } from './run';
import { accountLookupSnippet, macMessageId, messageLookupSnippet } from './scripts';
import { composeReplyHtml } from '../shared/replyBody';
import { listOutlookSignatures, readOutlookSignatureHtml } from './signatures';
import { readTemplateEmails } from './templates';
import type { ReplyEmailParams, ReplyEmailResult, SendEmailParams } from '../types';

/** Split a To/CC string the way the Windows contract accepts it. */
function splitAddresses(s: string): string[] {
    return s.split(/[,;]+/).map(x => x.trim()).filter(Boolean);
}

/**
 * What to do with a composed item.
 *
 * A new outgoing message lands in Temporary Items, not Drafts — `open` is what
 * surfaces it, so a windowless draft has to be moved into the Drafts folder
 * explicitly. (`save` is no help: it demands a file destination, not a folder.)
 * `move` reports success even where it silently does nothing, so the folder is
 * confirmed to have grown rather than risk reporting a draft that doesn't exist.
 */
function composeAction(
    variable: string,
    sendImmediately: boolean | undefined,
    openDraftWindow: boolean | undefined,
): { prelude: string; action: string } {
    if (sendImmediately) return {prelude: '', action: `    send ${variable}`};
    if (openDraftWindow !== false) return {prelude: '', action: `    open ${variable}`};
    return {
        prelude: `    set draftsFolder to drafts of targetAcct
    set draftsBefore to count of messages of draftsFolder`,
        action: `    move ${variable} to draftsFolder
    if (count of messages of draftsFolder) is not greater than draftsBefore then error "Outlook did not file the draft in the Drafts folder."`,
    };
}

/**
 * Create an Outlook email as a draft window (or send it) via AppleScript.
 * Mirrors the Windows COM contract from windows/send.ts.
 */
export async function sendOutlookEmail(params: SendEmailParams): Promise<void> {
    // Same safety property as Windows: never silently send from the wrong mailbox.
    // The account lookup raises when the address doesn't resolve, so there's no
    // pre-flight getOutlookAccounts() round-trip — it cost an extra osascript
    // launch on every email in a batch.
    const recipientLines = [
        ...splitAddresses(params.to).map(addr =>
            `    make new to recipient at newMsg with properties {email address:{address:"${asEscape(addr)}"}}`),
        ...splitAddresses(params.cc || '').map(addr =>
            `    make new cc recipient at newMsg with properties {email address:{address:"${asEscape(addr)}"}}`),
    ].join('\n');

    const attachLine = params.attachmentPath
        ? `    make new attachment at newMsg with properties {file:POSIX file "${asEscape(params.attachmentPath)}"}`
        : '';

    const {prelude, action} = composeAction('newMsg', params.sendImmediately, params.openDraftWindow);
    const script = `
tell application "Microsoft Outlook"
${accountLookupSnippet(params.emailAccount)}
${prelude}
    set newMsg to make new outgoing message with properties {subject:"${asEscape(params.subject)}", content:"${asEscape(params.htmlBody)}"}
${recipientLines}
    set account of newMsg to targetAcct
${attachLine}
${action}
end tell`;
    await runOsaScript(script, 120000);
}

/**
 * Reply to an email, inserting the caller's HTML above the quoted original.
 *
 * `entryId` is Outlook for Mac's small integer message id (see readInboxEmails),
 * not a Windows MAPI EntryID; `storeId` is accepted for signature parity and
 * ignored, since macOS has no StoreID and `message id N` resolves against the
 * application rather than one folder.
 *
 * `reply to` builds the quoted body Outlook itself would have — the same thing
 * COM's Reply() returns — so the insertion point is the one place the two
 * platforms have to agree, and both look for WordSection1 then `<body>`.
 *
 * The `to` line is the reply's resolved recipient addresses rather than the
 * display-name string Windows reports: Outlook for Mac exposes recipients as
 * records, and an address is the more useful half to verify against.
 */
export async function replyOutlookEmail(params: ReplyEmailParams): Promise<ReplyEmailResult> {
    const id = macMessageId(params.entryId);
    const insertHtml = await composeReplyHtml(params, {
        readTemplateEmails,
        readOutlookSignatureHtml,
        listOutlookSignatures,
    });
    const {prelude, action} = composeAction('theReply', params.sendImmediately, params.openDraftWindow);
    const replyAll = params.replyAll ? 'reply to all true' : 'without reply to all';
    const script = `tell application "Microsoft Outlook"
${accountLookupSnippet(params.emailAccount)}
${messageLookupSnippet(id)}
    set repliedTo to ""
    try
        -- Bind the record before reading its field: a nested
        -- 'address of (sender of theMsg)' does not coerce.
        set origSender to sender of theMsg
        set repliedTo to (address of origSender) as string
    end try
${prelude}
    set theReply to reply to theMsg opening window false ${replyAll}
    -- Neither of these is wrapped in a try, and both would read more defensively
    -- if they were. A swallowed account assignment sends from whichever mailbox
    -- Outlook considers default, and a swallowed insertion files a reply with the
    -- caller's text missing — the two failures this package exists to prevent,
    -- both of them invisible to the caller if tolerated here.
    set account of theReply to targetAcct
    set content of theReply to my insertAboveQuoted(content of theReply, "${asEscape(insertHtml)}")
    -- Read the outgoing fields BEFORE acting: sending moves the item, after which
    -- the reference is no longer readable.
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
end tell`;
    const parts = splitFields(await runOsaScript(script, 120000));
    return {
        to: field(parts, 0),
        subject: field(parts, 1),
        repliedToSender: field(parts, 2),
    };
}
