// Saving attachments off an email, one at a time or in a batch.
import { psEscape, psList, requireWindows, runPowerShell } from './run';
import { getItemScript } from './scripts';
import { parseArray, parseObject, record, str } from '../shared/json';
import { attachmentNotFoundMessage } from '../shared/attachmentMatch';
import { resolveDestDir } from '../runtime';
import { NotFoundError } from '../errors';
import type { SavedAttachment } from '../types';

/**
 * The "not found" throw, built from the shared wording with PowerShell
 * interpolations in place of the values — so the sentence a Windows caller sees
 * is literally the same one macOS composes.
 */
const NOT_FOUND_THROW = `throw "${attachmentNotFoundMessage({
    fileName: '$target',
    senderEmail: '$($item.SenderEmailAddress)',
    subject: '$($item.Subject)',
    receivedTime: '$($item.ReceivedTime)',
    present: '$haveStr',
    idLabel: 'EntryID',
})}"`;

/**
 * Locate one attachment by name into `$found`, or throw naming the email that
 * actually resolved. Matching is exact first, then whitespace-normalized (trim +
 * collapse runs, incl. non-breaking spaces) so a trivially reformatted filename
 * still resolves.
 */
const FIND_ATTACHMENT_PS = `
    $found = $null
    foreach ($att in $item.Attachments) {
        if ($att.FileName -eq $target) { $found = $att; break }
    }
    if ($found -eq $null) {
        $tnorm = ($target -replace '\\s+', ' ').Trim()
        foreach ($att in $item.Attachments) {
            if ((($att.FileName -replace '\\s+', ' ').Trim()) -ieq $tnorm) { $found = $att; break }
        }
    }
    if ($found -eq $null) {
        $have = @(); foreach ($a in $item.Attachments) { $have += $a.FileName }
        $haveStr = if ($have.Count -gt 0) { $have -join ', ' } else { '(none)' }
        ${NOT_FOUND_THROW}
    }
`;

/** The row every saver returns: the file, plus who the email was really from. */
const SAVED_ROW_PS = `[PSCustomObject]@{
        path         = $savePath
        subject      = if ($item.Subject) { $item.Subject } else { '' }
        senderName   = if ($item.SenderName) { $item.SenderName } else { '' }
        senderEmail  = if ($item.SenderEmailAddress) { $item.SenderEmailAddress } else { '' }
        receivedTime = if ($item.ReceivedTime) { $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } else { '' }
    }`;

function savedAttachment(value: unknown): SavedAttachment {
    const e = record(value);
    return {
        path: str(e.path),
        subject: str(e.subject),
        senderName: str(e.senderName),
        senderEmail: str(e.senderEmail),
        receivedTime: str(e.receivedTime),
    };
}

/**
 * Save an email attachment by entryId and filename, returning the saved path
 * together with the resolved email's subject/sender so the caller can confirm the
 * file came from the email it intended.
 *
 * `destDir` is created if absent. Without one the file lands in a fresh directory
 * of its own, because attachments keep the name the sender gave them: a single
 * shared folder means two emails carrying "invoice.pdf" silently overwrite each
 * other. Pass `destDir` whenever you want the files somewhere you control.
 *
 * When the attachment isn't found, throws an error that names the email actually
 * resolved (from/subject/received) and the attachments it does carry. Because
 * GetItemFromID happily returns a *different* message when handed a stale or wrong
 * EntryID (common when many replies share one subject), a bare "not found" is
 * misleading — this surfaces which email you're really looking at.
 */
export async function saveEmailAttachmentDetailed(
    entryId: string,
    fileName: string,
    storeId?: string,
    destDir?: string,
): Promise<SavedAttachment> {
    requireWindows();
    const outDir = resolveDestDir(destDir);
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
${getItemScript(entryId, storeId)}
if ($item -eq $null) { throw "Email not found for EntryID '${psEscape(entryId)}'" }
$target = '${psEscape(fileName)}'
${FIND_ATTACHMENT_PS}
$savePath = [IO.Path]::Combine('${psEscape(outDir)}', $found.FileName)
$found.SaveAsFile($savePath)
$out = ${SAVED_ROW_PS}
ConvertTo-Json $out -Compress
`;
    const raw = await runPowerShell(script, 15000);
    if (!raw || !raw.trim()) throw new NotFoundError('attachment', 'Failed to save attachment.');
    return savedAttachment(parseObject(raw));
}

/** Save an attachment and return just the saved file path. */
export async function saveEmailAttachment(
    entryId: string,
    fileName: string,
    storeId?: string,
    destDir?: string,
): Promise<string> {
    return (await saveEmailAttachmentDetailed(entryId, fileName, storeId, destDir)).path;
}

/**
 * Save several attachments from one email in a single COM round trip — what a
 * caller working through searchInboxByFilter/readSelectedEmail results wants,
 * rather than paying a PowerShell process spawn per attachment. Destination,
 * matching and the not-found error follow saveEmailAttachmentDetailed's rules
 * exactly, applied per name; results come back in the same order as `fileNames`.
 */
export async function saveEmailAttachments(
    entryId: string,
    fileNames: string[],
    storeId?: string,
    destDir?: string,
): Promise<SavedAttachment[]> {
    requireWindows();
    if (fileNames.length === 0) return [];
    const outDir = resolveDestDir(destDir);
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
${getItemScript(entryId, storeId)}
if ($item -eq $null) { throw "Email not found for EntryID '${psEscape(entryId)}'" }
$targets = @(${psList(fileNames)})
$results = @()
foreach ($target in $targets) {
${FIND_ATTACHMENT_PS}
    $savePath = [IO.Path]::Combine('${psEscape(outDir)}', $found.FileName)
    $found.SaveAsFile($savePath)
    $results += ${SAVED_ROW_PS}
}
ConvertTo-Json $results -Depth 3 -Compress
`;
    const raw = await runPowerShell(script, 15000);
    if (!raw || !raw.trim()) throw new NotFoundError('attachment', 'Failed to save attachments.');
    return parseArray(raw).map(savedAttachment);
}
