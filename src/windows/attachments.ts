// Saving attachments off an email.
import { psArray, psInt, psString, runPowerShellJson } from './run';
import { itemLookupScript, SENDER_SMTP_PS, SESSION_PS } from './scripts';
import { record, str, toArray } from '../shared/json';
import {
    attachmentNotFoundMessage,
    MAX_FILE_STEM,
    RESERVED_FILE_NAME,
    UNSAFE_FILE_NAME_CHARS,
} from '../shared/attachmentMatch';
import { failureTag } from '../errors';
import type { SaveAttachmentsRequest } from '../backend';
import type { SavedAttachment } from '../types';

/**
 * The file-name rules from shared/attachmentMatch.ts, as PowerShell: the same
 * sanitizing `safeFileName` does, and the same never-overwrite numbering as
 * `uniqueSavePath`.
 */
const SAVE_PATH_PS = `
$unsafeChars = ${psString(UNSAFE_FILE_NAME_CHARS)}
$reservedName = ${psString(RESERVED_FILE_NAME)}
$maxStem = ${psInt(MAX_FILE_STEM)}
function Get-SafeFileName([string]$name) {
    $safe = [regex]::Replace($name, $unsafeChars, '_').TrimEnd([char[]]'. ')
    if ($safe -match $reservedName) { $safe = '_' + $safe }
    $ext = [IO.Path]::GetExtension($safe)
    $stem = $safe.Substring(0, $safe.Length - $ext.Length)
    if ($stem.Length -gt $maxStem) { $safe = $stem.Substring(0, $maxStem) + $ext }
    if (-not $safe) { $safe = 'attachment' }
    return $safe
}
function Get-UniqueSavePath([string]$dir, [string]$name) {
    $ext = [IO.Path]::GetExtension($name)
    $stem = $name.Substring(0, $name.Length - $ext.Length)
    $candidate = [IO.Path]::Combine($dir, $name)
    $n = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = [IO.Path]::Combine($dir, "$stem ($n)$ext")
        $n++
    }
    return $candidate
}
`;

/**
 * Save the named attachments of one email, in the order named.
 *
 * Every name is matched before anything is written — exact first, then ignoring
 * case and whitespace runs — so a batch with one bad name fails without having
 * saved half its files. A name that matches nothing fails as NOT_FOUND naming
 * the email that actually resolved and the attachments it does carry: an id can
 * resolve to a different email than the caller meant, and a bare "not found"
 * would point at the wrong problem.
 */
export async function saveEmailAttachments(request: SaveAttachmentsRequest): Promise<SavedAttachment[]> {
    const notFound = attachmentNotFoundMessage({
        fileName: '$wanted',
        senderEmail: '$senderSmtp',
        subject: '$subject',
        receivedTime: '$received',
        present: '$present',
    });
    const output = await runPowerShellJson(`${SESSION_PS}
${itemLookupScript(request.email)}
${SENDER_SMTP_PS}
${SAVE_PATH_PS}
$destDir = ${psString(request.destDir)}
$subject = ''
try { $subject = [string]$item.Subject } catch {}
$senderName = ''
try { $senderName = [string]$item.SenderName } catch {}
$received = ''
try { $received = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } catch {}
$attachments = @()
foreach ($att in $item.Attachments) { $attachments += $att }
$picks = @()
foreach ($wanted in ${psArray(request.fileNames)}) {
    $found = $null
    foreach ($att in $attachments) { if ([string]$att.FileName -ceq $wanted) { $found = $att; break } }
    if ($found -eq $null) {
        $normalized = ($wanted -replace '\\s+', ' ').Trim()
        foreach ($att in $attachments) {
            if ((([string]$att.FileName) -replace '\\s+', ' ').Trim() -ieq $normalized) { $found = $att; break }
        }
    }
    if ($found -eq $null) {
        $present = (@($attachments | ForEach-Object { [string]$_.FileName }) -join ', ')
        if (-not $present) { $present = '(none)' }
        throw "${failureTag('NOT_FOUND', 'attachment')}${notFound}"
    }
    $picks += $found
}
$rows = @()
foreach ($att in $picks) {
    $fileName = [string]$att.FileName
    $savePath = Get-UniqueSavePath $destDir (Get-SafeFileName $fileName)
    $att.SaveAsFile($savePath)
    $rows += [PSCustomObject]@{
        path         = $savePath
        fileName     = $fileName
        subject      = $subject.Trim()
        senderName   = $senderName
        senderEmail  = $senderSmtp
        receivedTime = $received
    }
}
ConvertTo-Json -Compress -Depth 3 -InputObject @($rows)
`, 'standard');
    return toArray(output).map(row => {
        const e = record(row);
        return {
            path: str(e.path),
            fileName: str(e.fileName),
            subject: str(e.subject),
            senderName: str(e.senderName),
            senderEmail: str(e.senderEmail),
            receivedTime: str(e.receivedTime),
        };
    });
}
