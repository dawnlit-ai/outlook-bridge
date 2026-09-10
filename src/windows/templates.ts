// Template emails: ordinary mail items kept in a mailbox folder and reused as
// reply bodies, plus the compose-window editor for one.
import {
    psBool,
    psEscape,
    readScriptOutput,
    requireWindows,
    runPowerShell,
    runPowerShellFile,
    scriptInput
} from './run';
import { accountScript, DELIVERY_STORE_PS, FIND_FOLDER_PS } from './scripts';
import { parseObject, record, str, toArray } from '../shared/json';
import { findTemplateMarkers } from '../outlookTemplateSections';
import { clamp } from '../mail';
import { NotFoundError } from '../errors';
import { tempFile } from '../runtime';
import type { SaveTemplateResult, TemplateFolderResult } from '../types';

/**
 * Read the template emails saved in a mailbox folder (default "Templates"),
 * returning each item's full HTML body. When the folder doesn't exist, returns
 * folderFound:false plus the mailbox's folder names instead of throwing, so the
 * caller can ask the user whether to create it rather than fail.
 */
export async function readTemplateEmails(
    emailAccount: string,
    folderName = 'Templates',
    limit = 20,
    includeBody = true,
    subject = '',
): Promise<TemplateFolderResult> {
    requireWindows();
    const cap = clamp(limit, 1, 50);
    const wanted = (subject || '').trim();
    const script = `${accountScript(emailAccount)}
${DELIVERY_STORE_PS}
$root = $store.GetRootFolder()
${FIND_FOLDER_PS}
$folder = Find-FolderByName $root '${psEscape(folderName)}' 3
if ($folder -eq $null) {
    $names = @()
    foreach ($f in $root.Folders) { $names += $f.Name }
    ConvertTo-Json @{ folderFound = $false; folderPath = ''; templates = @(); availableFolders = $names } -Depth 3
    exit 0
}
$includeBody = ${psBool(includeBody)}
$wanted = '${psEscape(wanted)}'
$items = $folder.Items
$items.Sort('[LastModificationTime]', $true)
$results = @()
# Without a subject filter the sort order does the capping (newest first). With one,
# scan the whole folder — the wanted template need not be among the newest — and stop
# as soon as the cap is filled.
$count = $items.Count
for ($i = 1; $i -le $count; $i++) {
    if ($results.Count -ge ${cap}) { break }
    $item = $null
    try { $item = $items.Item($i) } catch { continue }
    if ($item -eq $null) { continue }
    $cls = 0
    try { $cls = [int]$item.Class } catch {}
    if ($cls -ne 43) { continue }  # olMail only
    if ($wanted -ne '') {
        $subj = ''
        try { $subj = [string]$item.Subject } catch {}
        if ($subj -eq $null) { $subj = '' }
        if ($subj.Trim() -ine $wanted) { continue }
    }
    $html = ''
    $preview = ''
    $markers = ''
    $plain = ''
    if ($includeBody) {
        try { $html = [string]$item.HTMLBody } catch {}
        # Embedded (cid:) images live as attachments on the TEMPLATE item: reusing this HTML
        # on a new email shows a broken "linked image" placeholder, and rewriting to file:
        # URIs renders invisibly in modern Outlook. Neither works, so strip the img tags —
        # template emails are text/HTML only.
        $html = [regex]::Replace($html, '<img[^>]*src="cid:[^"]*"[^>]*>', '')
        if ($html.Length -gt 100000) { $html = $html.Substring(0, 100000) }
    } else {
        # Bodies omitted: return a short plain-text preview so the caller can tell templates
        # apart by subject without pulling the full (often huge) Word-generated HTML.
        try { $plain = [string]$item.Body } catch { $plain = '' }
        if ($plain) {
            # ...plus the [[SECTION]] / {{PLACEHOLDER}} markers, scanned from the full text:
            # that's what lets a caller confirm a sectioned template is intact without
            # fetching its body. Word can split a marker across tags in the HTML but not
            # in the plain-text body, so a literal scan is right here.
            # (Backslashes in the regex are doubled because this whole script is a JS
            # template literal — a single one would be eaten before PowerShell saw it.)
            $found = @()
            foreach ($x in [regex]::Matches($plain, '\\[\\[\\s*/?\\s*[A-Za-z0-9_-]{1,40}\\s*\\]\\]|\\{\\{\\s*[A-Za-z0-9_-]{1,40}\\s*\\}\\}')) {
                $t = ($x.Value -replace '\\s', '')
                if ($found -notcontains $t) { $found += $t }
            }
            $markers = ($found -join ' ')
            $preview = $plain.Trim()
            if ($preview.Length -gt 200) { $preview = $preview.Substring(0, 200) }
        }
    }
    $results += [PSCustomObject]@{
        entryId      = $item.EntryID
        subject      = if ($item.Subject) { $item.Subject.Trim() } else { '' }
        htmlBody     = $html
        bodyPreview  = $preview
        markers      = $markers
        lastModified = $item.LastModificationTime.ToString('yyyy-MM-dd HH:mm')
    }
}
ConvertTo-Json @{ folderFound = $true; folderPath = $folder.FolderPath; templates = $results; availableFolders = @() } -Depth 4
`;
    const parsed = parseObject(await runPowerShell(script, 60000));
    return {
        folderFound: Boolean(parsed.folderFound),
        folderPath: str(parsed.folderPath),
        templates: toArray(parsed.templates).map(item => {
            const e = record(item);
            const html = str(e.htmlBody);
            // With a body in hand the HTML is authoritative (it survives Word splitting a
            // marker across tags); without one, fall back to what the script scanned out
            // of the plain-text body.
            const markers = findTemplateMarkers(html || str(e.markers));
            return {
                entryId: str(e.entryId),
                subject: str(e.subject),
                htmlBody: html,
                bodyPreview: str(e.bodyPreview),
                sections: markers.sections,
                placeholders: markers.placeholders,
                lastModified: str(e.lastModified),
            };
        }),
        availableFolders: toArray(parsed.availableFolders).map(str),
    };
}

/**
 * Save a new template email into a mailbox folder (default "Templates"),
 * creating the folder at the mailbox root if it doesn't exist. The item is a
 * plain unsent mail (subject + HTML body) that can be edited in Outlook.
 * Never overwrites an existing item — it always adds a new one, so check with
 * readTemplateEmails first and only call after the user has agreed to create it.
 */
export async function saveTemplateEmail(
    emailAccount: string,
    subject: string,
    htmlBody: string,
    folderName = 'Templates',
): Promise<SaveTemplateResult> {
    requireWindows();
    const body = scriptInput('template-body', htmlBody);
    const script = `${accountScript(emailAccount)}
${DELIVERY_STORE_PS}
$root = $store.GetRootFolder()
${FIND_FOLDER_PS}
$folder = Find-FolderByName $root '${psEscape(folderName)}' 3
$folderCreated = $false
if ($folder -eq $null) {
    $folder = $root.Folders.Add('${psEscape(folderName)}')
    $folderCreated = $true
}
$mail = $outlook.CreateItem(0)
$mail.Subject = '${psEscape(subject)}'
$mail.HTMLBody = [IO.File]::ReadAllText('${psEscape(body.path)}', [Text.Encoding]::UTF8)
$mail.Save()
[void]$mail.Move($folder)
ConvertTo-Json @{ folderPath = $folder.FolderPath; folderCreated = $folderCreated }
`;
    try {
        const parsed = parseObject(await runPowerShell(script, 60000));
        return {
            folderPath: str(parsed.folderPath),
            folderCreated: Boolean(parsed.folderCreated),
        };
    } finally {
        body.cleanup();
    }
}

/**
 * Open an Outlook compose window pre-filled with `currentHtml`, wait for the
 * user to close it, and return the HTML they saved.
 *
 * Deliberately unbounded: the run lasts as long as the person is editing. Give
 * it an AbortSignal (`bridge.withOptions({ signal })`) if the caller needs a way
 * to give up on them.
 */
export async function editEmailTemplate(label: string, currentHtml: string): Promise<string> {
    requireWindows();
    const input = scriptInput('template-input', currentHtml);
    const outputFile = tempFile('template-output', 'html');
    const subject = `${label} - Save (Ctrl+S) and close when done`;
    const script = `
$ErrorActionPreference = 'Stop'
$inputPath = '${psEscape(input.path)}'
$outputPath = '${psEscape(outputFile)}'
$subject = '${psEscape(subject)}'

$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$mail = $outlook.CreateItem(0)
$mail.Subject = $subject
$mail.HTMLBody = [IO.File]::ReadAllText($inputPath, [Text.Encoding]::UTF8)
$mail.Display()

# Give the inspector time to fully initialize
Start-Sleep -Seconds 2

# Poll until the compose window is closed, keeping the last live body as a
# fallback for the case where the user edits but never presses Ctrl+S.
$liveBody = ''
while ($true) {
    Start-Sleep -Milliseconds 700
    $open = $false
    try {
        foreach ($insp in $outlook.Inspectors) {
            $ci = $null
            try { $ci = $insp.CurrentItem } catch {}
            if ($ci -ne $null -and $ci.Subject -eq $subject) {
                $open = $true
                try { $liveBody = $ci.HTMLBody } catch {}
                break
            }
        }
    } catch {}
    if (-not $open) { break }
}

# Prefer the saved draft — it reflects the user's Ctrl+S and survives the
# inspector closing, unlike the original mail reference. Also removes any
# leftover editor drafts so they don't pile up in the Drafts folder.
$savedBody = ''
try {
    $drafts = $ns.GetDefaultFolder(16)
    $draftItems = $drafts.Items
    for ($i = $draftItems.Count; $i -ge 1; $i--) {
        $it = $draftItems.Item($i)
        if ($it.Subject -eq $subject) {
            if (-not $savedBody) { try { $savedBody = $it.HTMLBody } catch {} }
            try { $it.Delete() } catch {}
        }
    }
} catch {}

$finalBody = if ($savedBody) { $savedBody } else { $liveBody }
if ($finalBody) {
    [IO.File]::WriteAllText($outputPath, $finalBody, [Text.Encoding]::UTF8)
}
`;
    try {
        // 0 disables the timeout: this run is as long as the edit takes.
        await runPowerShellFile(script, 0);
    } finally {
        input.cleanup();
    }
    const result = readScriptOutput(outputFile);
    if (!result) {
        throw new NotFoundError('template', 'No template saved. Did you save (Ctrl+S) before closing?');
    }
    return result;
}
