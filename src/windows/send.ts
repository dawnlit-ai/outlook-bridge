// Composing outgoing mail: a new email, and a reply to an existing one.
import { psEscape, requireWindows, runPowerShell, runPowerShellFile, scriptInput } from './run';
import { accountScript, getItemScript } from './scripts';
import { composeReplyHtml } from '../shared/replyBody';
import { parseObject, str } from '../shared/json';
import { listOutlookSignatures, readOutlookSignatureHtml } from './signatures';
import { readTemplateEmails } from './templates';
import type { ReplyEmailParams, ReplyEmailResult, SendEmailParams } from '../types';

/**
 * Set the sending account through IDispatch reflection.
 *
 * Direct assignment (`$mail.SendUsingAccount = $account`) is a silent no-op under
 * PowerShell's COM binding — the mail then goes out from whichever account
 * Outlook considers default, which is the failure this package exists to prevent.
 */
function sendUsingAccount(variable: string): string {
    return `[void]$${variable}.GetType().InvokeMember('SendUsingAccount', [Reflection.BindingFlags]::SetProperty, $null, $${variable}, @($account))`;
}

/**
 * What to do with a composed item: send it, show it, or file it silently.
 *
 * `Display()` only shows the compose window — it never files the item, so a
 * windowless draft has to `Save()` into Drafts explicitly.
 */
function composeAction(
    variable: string,
    sendImmediately: boolean | undefined,
    openDraftWindow: boolean | undefined,
): string {
    if (sendImmediately) return `$${variable}.Send()`;
    return openDraftWindow !== false ? `$${variable}.Display()` : `$${variable}.Save()`;
}

/** Send (or display for review) an email through Outlook COM. */
export async function sendOutlookEmail(params: SendEmailParams): Promise<void> {
    requireWindows();
    const attachLine = params.attachmentPath
        ? `$mail.Attachments.Add('${psEscape(params.attachmentPath)}') | Out-Null`
        : '';
    // The HTML body goes through a file: inline, a large email is what pushes the
    // generated script past the command-line length limit.
    const body = scriptInput('email-body', params.htmlBody);
    const script = `${accountScript(params.emailAccount)}
$mail = $outlook.CreateItem(0)
${sendUsingAccount('mail')}
$mail.To = '${psEscape(params.to)}'
$mail.CC = '${psEscape(params.cc || '')}'
$mail.Subject = '${psEscape(params.subject)}'
$mail.HTMLBody = [IO.File]::ReadAllText('${psEscape(body.path)}', [Text.Encoding]::UTF8)
${attachLine}
${composeAction('mail', params.sendImmediately, params.openDraftWindow)}
`;
    try {
        await runPowerShellFile(script, 120000);
    } finally {
        body.cleanup();
    }
}

/**
 * Reply to an email, inserting the caller's HTML above the quoted original.
 *
 * The body may be given directly or named as a saved template — see
 * `composeReplyHtml`, which resolves the template, its section and its
 * placeholders before anything reaches Outlook.
 */
export async function replyOutlookEmail(params: ReplyEmailParams): Promise<ReplyEmailResult> {
    requireWindows();
    const insertHtml = await composeReplyHtml(params, {
        readTemplateEmails,
        readOutlookSignatureHtml,
        listOutlookSignatures,
    });
    const body = scriptInput('reply-insert', insertHtml);
    const replyMethod = params.replyAll ? 'ReplyAll' : 'Reply';
    // Outlook builds the reply as a Word document whose own empty paragraph sits
    // inside WordSection1; inserting straight after that div's opening tag puts
    // the new text where the user's cursor would have been. A reply Outlook did
    // not build that way falls back to just inside <body>.
    const script = `${accountScript(params.emailAccount)}
${getItemScript(params.entryId, params.storeId)}
$reply = $item.${replyMethod}()
${sendUsingAccount('reply')}

$insertHtml = [IO.File]::ReadAllText('${psEscape(body.path)}', [Text.Encoding]::UTF8)
$keyWord = 'WordSection1>'
$idx = $reply.HTMLBody.IndexOf($keyWord)
if ($idx -ge 0) { $idx += $keyWord.Length } else {
    $m = [regex]::Match($reply.HTMLBody, '<body[^>]*>')
    $idx = if ($m.Success) { $m.Index + $m.Length } else { 0 }
}
$reply.HTMLBody = $reply.HTMLBody.Insert($idx, $insertHtml)
$to = [string]$reply.To
$subject = [string]$reply.Subject
$sender = ''
try { $sender = [string]$item.SenderEmailAddress } catch {}
${composeAction('reply', params.sendImmediately, params.openDraftWindow)}
ConvertTo-Json @{ to = $to; subject = $subject; repliedToSender = $sender }
`;
    try {
        const parsed = parseObject(await runPowerShell(script, 60000));
        return {
            to: str(parsed.to),
            subject: str(parsed.subject),
            repliedToSender: str(parsed.repliedToSender),
        };
    } finally {
        body.cleanup();
    }
}
