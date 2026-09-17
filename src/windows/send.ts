// Composing outgoing mail: a new email, and a reply to an existing one.
import { psString, runPowerShell, runPowerShellJson, scriptInput } from './run';
import { accountScript, itemLookupScript, SENDER_SMTP_PS } from './scripts';
import { WORD_SECTION_ANCHOR } from '../shared/replyBody';
import { failureTag } from '../errors';
import { record, str } from '../shared/json';
import type { Disposition, ReplyRequest, SendRequest } from '../backend';
import type { ReplyEmailResult } from '../types';

/**
 * Set the sending account through IDispatch reflection.
 *
 * Plain assignment (`$mail.SendUsingAccount = $account`) is a silent no-op under
 * PowerShell's COM binding — the mail then goes out from whichever account
 * Outlook considers default, which is exactly what this package must never do.
 *
 * Refuses a mailbox that resolved to a store with no Account: SendUsingAccount
 * takes an Account, and sending with none falls through to the default account.
 * The usual cause is a stale session — Outlook builds its Accounts list at
 * startup, so a mailbox added since has its store mounted and readable but no
 * Account until Outlook restarts — and the message says so.
 */
function sendUsingAccount(variable: string): string {
    return `if ($account -eq $null) {
    throw "${failureTag('INVALID_REQUEST')}Mailbox '$target' has no sending account in this Outlook session. If it was added to the profile after Outlook started, restart Outlook and try again."
}
[void]$${variable}.GetType().InvokeMember('SendUsingAccount', [Reflection.BindingFlags]::SetProperty, $null, $${variable}, @($account))`;
}

/**
 * What to do with a composed item. `Display()` only shows a compose window and
 * never files the item, so a windowless draft has to `Save()` into Drafts.
 */
function dispose(variable: string, disposition: Disposition): string {
    switch (disposition) {
        case 'send':
            return `$${variable}.Send()`;
        case 'display':
            return `$${variable}.Display()`;
        case 'save':
            return `$${variable}.Save()`;
    }
}

/** Outlook takes recipients as one semicolon-separated line. */
function recipientLine(entries: readonly string[]): string {
    return psString(entries.join('; '));
}

/** Send, display or file a new email. */
export async function sendOutlookEmail(request: SendRequest): Promise<void> {
    // The body goes through a file: inline, a large email bloats the script and
    // every error that carries it.
    const body = scriptInput('email-body', request.htmlBody);
    const attach = request.attachments
        .map(file => `[void]$mail.Attachments.Add(${psString(file)})`)
        .join('\n');
    try {
        await runPowerShell(`${accountScript(request.account)}
$mail = $outlook.CreateItem(0)
${sendUsingAccount('mail')}
$mail.To = ${recipientLine(request.to)}
$mail.CC = ${recipientLine(request.cc)}
$mail.BCC = ${recipientLine(request.bcc)}
$mail.Subject = ${psString(request.subject)}
$mail.HTMLBody = [IO.File]::ReadAllText(${psString(body.path)}, [Text.Encoding]::UTF8)
${attach}
${dispose('mail', request.disposition)}
`, 'standard');
    } finally {
        body.cleanup();
    }
}

/** Reply to an email, with the composed HTML above the quoted original. */
export async function replyOutlookEmail(request: ReplyRequest): Promise<ReplyEmailResult> {
    const insert = scriptInput('reply-insert', request.html);
    try {
        const output = await runPowerShellJson(`${accountScript(request.account)}
${itemLookupScript(request.email)}
$reply = $item.${request.replyAll ? 'ReplyAll' : 'Reply'}()
${sendUsingAccount('reply')}
$insertHtml = [IO.File]::ReadAllText(${psString(insert.path)}, [Text.Encoding]::UTF8)
$anchor = ${psString(WORD_SECTION_ANCHOR)}
# Read once: every HTMLBody read marshals the whole generated reply across COM,
# and a reply on a long thread runs to hundreds of KB.
$html = [string]$reply.HTMLBody
$at = $html.IndexOf($anchor)
if ($at -ge 0) { $at += $anchor.Length } else {
    $bodyTag = [regex]::Match($html, '<body[^>]*>')
    $at = if ($bodyTag.Success) { $bodyTag.Index + $bodyTag.Length } else { 0 }
}
$reply.HTMLBody = $html.Insert($at, $insertHtml)
$to = [string]$reply.To
$subject = [string]$reply.Subject
${SENDER_SMTP_PS}
${dispose('reply', request.disposition)}
ConvertTo-Json -Compress @{ to = $to; subject = $subject; repliedToSender = $senderSmtp }
`, 'standard');
        const result = record(output);
        return {to: str(result.to), subject: str(result.subject), repliedToSender: str(result.repliedToSender)};
    } finally {
        insert.cleanup();
    }
}
