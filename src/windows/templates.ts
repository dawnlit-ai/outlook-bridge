// Template emails: ordinary mail items kept in a mailbox folder and reused as
// email bodies, plus the compose-window editor for one.
import { psBool, psInt, psString, readScriptOutput, runPowerShell, runPowerShellJson, scriptInput } from './run';
import { accountScript, DELIVERY_STORE_PS, FIND_FOLDER_PS, SESSION_PS } from './scripts';
import { bool, record, str, toArray } from '../shared/json';
import { findTemplateMarkers } from '../templateBody';
import { FolderId } from '../mail';
import { NotFoundError } from '../errors';
import { tempFile } from '../runtime';
import type { EditTemplateRequest, ReadTemplatesRequest, SaveTemplateRequest } from '../backend';
import type { SaveTemplateResult, TemplateFolderResult } from '../types';

/** How deep under the mailbox root a template folder is searched for. */
const TEMPLATE_SEARCH_DEPTH = 3;

/** Cap on one template body; a Word-generated one is large, and a runaway one shouldn't sink the call. */
const MAX_TEMPLATE_HTML = 100_000;

/** Plain text standing in for a body that wasn't asked for. */
const PREVIEW_CHARS = 200;

/**
 * The template emails in a mailbox folder, newest first. A missing folder is
 * reported as folderFound:false with the mailbox's top-level folder names
 * rather than thrown, so a caller can offer to create it.
 */
export async function readTemplateEmails(request: ReadTemplatesRequest): Promise<TemplateFolderResult> {
    const output = await runPowerShellJson(`${accountScript(request.account)}
${DELIVERY_STORE_PS}
$root = $store.GetRootFolder()
${FIND_FOLDER_PS}
$folder = Find-FolderByName $root ${psString(request.folder)} ${psInt(TEMPLATE_SEARCH_DEPTH)}
if ($folder -eq $null) {
    $names = @()
    foreach ($f in $root.Folders) { $names += [string]$f.Name }
    ConvertTo-Json -Compress -Depth 3 -InputObject ([PSCustomObject]@{ folderFound = $false; folderPath = ''; templates = @(); availableFolders = @($names) })
    exit 0
}
$includeBody = ${psBool(request.includeBody)}
$wanted = ${psString(request.subject ?? '')}
$limit = ${psInt(request.limit)}
$maxHtml = ${psInt(MAX_TEMPLATE_HTML)}
$previewChars = ${psInt(PREVIEW_CHARS)}
$items = $folder.Items
$items.Sort('[LastModificationTime]', $true)
$rows = @()
$count = $items.Count
# Without a subject the sort does the capping, newest first. With one, the whole
# folder is scanned — the template wanted need not be among the newest.
for ($i = 1; $i -le $count -and $rows.Count -lt $limit; $i++) {
    $item = $null
    try { $item = $items.Item($i) } catch { continue }
    if ($item -eq $null) { continue }
    $class = 0
    try { $class = [int]$item.Class } catch {}
    if ($class -ne 43) { continue }
    $subject = ''
    try { $subject = ([string]$item.Subject).Trim() } catch {}
    if ($wanted -and $subject -ine $wanted) { continue }
    $html = ''
    $preview = ''
    $markers = ''
    if ($includeBody) {
        try { $html = [string]$item.HTMLBody } catch {}
        # Embedded (cid:) images live as attachments on the TEMPLATE item, so its
        # HTML reused elsewhere shows broken-image placeholders. Template emails
        # are text and markup only.
        $html = [regex]::Replace($html, '<img[^>]*src="cid:[^"]*"[^>]*>', '')
        if ($html.Length -gt $maxHtml) { $html = $html.Substring(0, $maxHtml) }
    } else {
        # No body asked for: a short preview, plus the section and placeholder
        # markers scanned here from the full plain text — so a caller can confirm
        # a template is intact without its body crossing stdout. Word can split a
        # marker across tags in the HTML but not in the plain text.
        $plain = ''
        try { $plain = [string]$item.Body } catch {}
        $found = @()
        foreach ($m in [regex]::Matches($plain, '\\[\\[\\s*/?\\s*[A-Za-z0-9_-]{1,40}\\s*\\]\\]|\\{\\{\\s*[A-Za-z0-9_-]{1,40}\\s*\\}\\}')) {
            $token = $m.Value -replace '\\s', ''
            if ($found -notcontains $token) { $found += $token }
        }
        $markers = $found -join ' '
        $preview = $plain.Trim()
        if ($preview.Length -gt $previewChars) { $preview = $preview.Substring(0, $previewChars) }
    }
    $modified = ''
    try { $modified = $item.LastModificationTime.ToString('yyyy-MM-dd HH:mm') } catch {}
    $rows += [PSCustomObject]@{
        entryId      = [string]$item.EntryID
        subject      = $subject
        htmlBody     = $html
        bodyPreview  = $preview
        markers      = $markers
        lastModified = $modified
    }
}
ConvertTo-Json -Depth 4 -InputObject ([PSCustomObject]@{
    folderFound      = $true
    folderPath       = [string]$folder.FolderPath
    templates        = @($rows)
    availableFolders = @()
})
`, 'standard');
    const e = record(output);
    return {
        folderFound: bool(e.folderFound),
        folderPath: str(e.folderPath),
        templates: toArray(e.templates).map(row => {
            const t = record(row);
            const html = str(t.htmlBody);
            // With the HTML in hand it is authoritative (markers survive Word
            // splitting them across tags); without it, the markers the script
            // scanned out of the plain text stand in.
            const markers = findTemplateMarkers(request.includeBody ? html : str(t.markers));
            return {
                entryId: str(t.entryId),
                subject: str(t.subject),
                htmlBody: html,
                bodyPreview: str(t.bodyPreview),
                sections: markers.sections,
                placeholders: markers.placeholders,
                lastModified: str(t.lastModified),
            };
        }),
        availableFolders: toArray(e.availableFolders).map(str),
    };
}

/**
 * Save a new template email into a folder, creating the folder at the mailbox
 * root when absent. Always adds a new item — never overwrites one.
 */
export async function saveTemplateEmail(request: SaveTemplateRequest): Promise<SaveTemplateResult> {
    const body = scriptInput('template-body', request.htmlBody);
    try {
        const output = await runPowerShellJson(`${accountScript(request.account)}
${DELIVERY_STORE_PS}
$root = $store.GetRootFolder()
${FIND_FOLDER_PS}
$folderName = ${psString(request.folder)}
$folder = Find-FolderByName $root $folderName ${psInt(TEMPLATE_SEARCH_DEPTH)}
$folderCreated = $false
if ($folder -eq $null) {
    $folder = $root.Folders.Add($folderName)
    $folderCreated = $true
}
$mail = $outlook.CreateItem(0)
$mail.Subject = ${psString(request.subject)}
$mail.HTMLBody = [IO.File]::ReadAllText(${psString(body.path)}, [Text.Encoding]::UTF8)
$mail.Save()
[void]$mail.Move($folder)
ConvertTo-Json -Compress -InputObject ([PSCustomObject]@{ folderPath = [string]$folder.FolderPath; folderCreated = $folderCreated })
`, 'standard');
        const e = record(output);
        return {folderPath: str(e.folderPath), folderCreated: bool(e.folderCreated)};
    } finally {
        body.cleanup();
    }
}

/**
 * Open a compose window pre-filled with the HTML, wait for it to close, and
 * return what was saved (Ctrl+S) — or, failing a save, the body as it last
 * stood while the window was open. Leftover editor drafts are removed.
 */
export async function editEmailTemplate(request: EditTemplateRequest): Promise<string> {
    const input = scriptInput('template-input', request.html);
    const outputFile = tempFile('template-output', 'html');
    try {
        await runPowerShell(`${SESSION_PS}
$subject = ${psString(`${request.label} - Save (Ctrl+S) and close when done`)}
$outputPath = ${psString(outputFile)}
$mail = $outlook.CreateItem(0)
$mail.Subject = $subject
$mail.HTMLBody = [IO.File]::ReadAllText(${psString(input.path)}, [Text.Encoding]::UTF8)
$mail.Display()
Start-Sleep -Seconds 2
# Poll until the compose window closes, keeping the last live body as the
# fallback for an edit that was never saved.
$liveBody = ''
while ($true) {
    Start-Sleep -Milliseconds 700
    $open = $false
    try {
        foreach ($inspector in $outlook.Inspectors) {
            $current = $null
            try { $current = $inspector.CurrentItem } catch {}
            if ($current -ne $null -and [string]$current.Subject -eq $subject) {
                $open = $true
                try { $liveBody = [string]$current.HTMLBody } catch {}
                break
            }
        }
    } catch {}
    if (-not $open) { break }
}
# Prefer the saved draft: it is what the person chose to keep, and it outlives
# the window. Every editor draft is removed so none pile up in Drafts.
$savedBody = ''
try {
    $draftItems = $ns.GetDefaultFolder(${psInt(FolderId.Drafts)}).Items
    for ($i = $draftItems.Count; $i -ge 1; $i--) {
        $draft = $draftItems.Item($i)
        if ([string]$draft.Subject -eq $subject) {
            if (-not $savedBody) { try { $savedBody = [string]$draft.HTMLBody } catch {} }
            try { $draft.Delete() } catch {}
        }
    }
} catch {}
$finalBody = if ($savedBody) { $savedBody } else { $liveBody }
if ($finalBody) { [IO.File]::WriteAllText($outputPath, $finalBody, [Text.Encoding]::UTF8) }
`, 'interactive');
    } finally {
        input.cleanup();
    }
    const saved = readScriptOutput(outputFile);
    if (!saved) {
        throw new NotFoundError('template', 'No template was saved. Save (Ctrl+S) before closing the window.');
    }
    return saved;
}
