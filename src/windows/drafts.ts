// Drafts: which ones belong to an account, and listing, sending or deleting them.
import { psArray, psInt, runPowerShellJson } from './run';
import { accountScript, DELIVERY_STORE_PS } from './scripts';
import { bool, itemFailures, num, record, str, strList, toArray } from '../shared/json';
import { FolderId } from '../mail';
import type { AccountRequest, EntryIdsRequest, ListDraftsRequest } from '../backend';
import type { DeleteDraftsResult, ListDraftsResult, SendDraftsResult } from '../types';

/**
 * Resolve the Drafts folders belonging to one account into `$scan`, and define
 * `Test-DraftMatches` over them. Emitted into every drafts script, so listing,
 * sending and deleting can never disagree about which drafts are this account's.
 *
 * Two Drafts folders can hold one account's drafts, so both are scanned:
 *  - the account's OWN store's Drafts, where a draft composed in that mailbox
 *    lands. A draft there with no explicit sending account still belongs to it.
 *  - the DEFAULT store's Drafts, where a program's `Save()` files a draft
 *    whatever account it sends from. Only drafts stamped with this account
 *    match there; an unstamped one belongs to the default account.
 * When the two are the same folder it is scanned once, under the first rule.
 *
 * Non-mail items (meeting requests, reports) never match.
 */
const DRAFTS_SCAN_PS = `
$scan = @()
$seenFolders = @{}
$homeDrafts = $null
try { $homeDrafts = $store.GetDefaultFolder(${psInt(FolderId.Drafts)}) } catch {}
if ($homeDrafts -ne $null) {
    $seenFolders["$($homeDrafts.StoreID)|$($homeDrafts.EntryID)"] = $true
    $scan += [PSCustomObject]@{ folder = $homeDrafts; includeUnstamped = $true }
}
$defaultDrafts = $null
try { $defaultDrafts = $ns.GetDefaultFolder(${psInt(FolderId.Drafts)}) } catch {}
if ($defaultDrafts -ne $null -and -not $seenFolders.ContainsKey("$($defaultDrafts.StoreID)|$($defaultDrafts.EntryID)")) {
    $scan += [PSCustomObject]@{ folder = $defaultDrafts; includeUnstamped = $false }
}
function Test-DraftMatches($item, [bool]$includeUnstamped) {
    if ($item.Class -ne 43) { return $false }  # olMail only
    $sender = $null
    try { $sender = $item.SendUsingAccount } catch {}
    if ($sender -eq $null) { return $includeUnstamped }
    return ([string]$sender.SmtpAddress -ieq $target)
}
# An id is acted on only once it is PROVED to be one of this account's drafts:
# resolved in one of the scanned folders' stores, sitting in that folder, and
# bound to this account (or unbound, in its own store).
$draftFolderRules = @{}
foreach ($entry in $scan) { $draftFolderRules[[string]$entry.folder.EntryID] = $entry.includeUnstamped }
function Resolve-OwnDraft([string]$id, [string]$verb) {
    $draft = $null
    foreach ($entry in $scan) {
        try { $draft = $ns.GetItemFromID($id, $entry.folder.StoreID) } catch { $draft = $null }
        if ($draft -ne $null) { break }
    }
    if ($draft -eq $null) { throw "no such item in this account's Drafts folders" }
    $parentId = ''
    try { $parentId = [string]$draft.Parent.EntryID } catch {}
    if (-not $draftFolderRules.ContainsKey($parentId)) { throw "item is not in this account's Drafts folder - refusing to $verb" }
    if (-not (Test-DraftMatches $draft $draftFolderRules[$parentId])) { throw "draft is not bound to $target - refusing to $verb" }
    return $draft
}
`;

function draftsScript(account: string): string {
    return accountScript(account) + DELIVERY_STORE_PS + DRAFTS_SCAN_PS;
}

/** The account's drafts, newest first. Bodies are previewed, never returned whole. */
export async function listOutlookDrafts(request: ListDraftsRequest): Promise<ListDraftsResult> {
    const output = await runPowerShellJson(`${draftsScript(request.account)}
$previewChars = ${psInt(request.previewChars)}
$limit = ${psInt(request.limit)}
$rows = @()
$folderPaths = @()
foreach ($entry in $scan) {
    $folderPath = ''
    try { $folderPath = [string]$entry.folder.FolderPath } catch {}
    $folderPaths += $folderPath
    foreach ($it in $entry.folder.Items) {
        if (-not (Test-DraftMatches $it $entry.includeUnstamped)) { continue }
        $addresses = @()
        try {
            foreach ($r in $it.Recipients) {
                $address = ''
                try { $address = [string]$r.Address } catch {}
                # An unresolved Exchange entry has a DN for an address; its name reads better.
                if ($address.StartsWith('/')) { try { $address = [string]$r.Name } catch {} }
                if ($address) { $addresses += $address }
            }
        } catch {}
        $preview = ''
        if ($previewChars -gt 0) {
            try { $preview = ([string]$it.Body -replace '\\s+', ' ').Trim() } catch {}
            if ($preview.Length -gt $previewChars) { $preview = $preview.Substring(0, $previewChars) }
        }
        $modified = ''
        try { $modified = $it.LastModificationTime.ToString('yyyy-MM-dd HH:mm') } catch {}
        $hasAttachments = $false
        try { $hasAttachments = ($it.Attachments.Count -gt 0) } catch {}
        $rows += [PSCustomObject]@{
            entryId        = [string]$it.EntryID
            subject        = [string]$it.Subject
            to             = [string]$it.To
            toEmails       = @($addresses)
            bodyPreview    = $preview
            hasAttachments = $hasAttachments
            lastModified   = $modified
            folderPath     = $folderPath
        }
    }
}
$sorted = @($rows | Sort-Object -Property lastModified -Descending)
$total = $sorted.Count
if ($total -gt $limit) { $sorted = @($sorted[0..($limit - 1)]) }
ConvertTo-Json -Depth 4 -InputObject ([PSCustomObject]@{
    account        = $target
    foldersScanned = @($folderPaths)
    count          = $total
    truncated      = ($total -gt $limit)
    drafts         = @($sorted)
})
`, 'scan');
    const e = record(output);
    return {
        account: str(e.account) || request.account,
        foldersScanned: toArray(e.foldersScanned).map(str),
        count: num(e.count),
        truncated: bool(e.truncated),
        drafts: toArray(e.drafts).map(row => {
            const d = record(row);
            return {
                entryId: str(d.entryId),
                subject: str(d.subject),
                to: str(d.to),
                toEmails: strList(d.toEmails),
                bodyPreview: str(d.bodyPreview),
                hasAttachments: bool(d.hasAttachments),
                lastModified: str(d.lastModified),
                folderPath: str(d.folderPath),
            };
        }),
    };
}

/** Send the named drafts, each proved to be this account's first. */
export async function sendDrafts(request: EntryIdsRequest): Promise<SendDraftsResult> {
    const output = await runPowerShellJson(`${draftsScript(request.account)}
$sent = 0
$failed = @()
foreach ($id in ${psArray(request.entryIds)}) {
    $subject = ''
    try {
        $draft = Resolve-OwnDraft $id 'send'
        try { $subject = [string]$draft.Subject } catch {}
        $draft.Send()
        $sent++
    } catch {
        $failed += [PSCustomObject]@{ entryId = $id; subject = $subject; error = $_.Exception.Message }
    }
}
ConvertTo-Json -Compress -Depth 3 -InputObject ([PSCustomObject]@{ sent = $sent; failed = @($failed) })
`, 'scan');
    const e = record(output);
    return {sent: num(e.sent), failed: itemFailures(e.failed)};
}

/**
 * Send every draft belonging to the account.
 *
 * The matching items are snapshotted before the first send: sending moves an
 * item out of Drafts, and sending during a live enumeration would skip every
 * other one.
 */
export async function sendAllDrafts(request: AccountRequest): Promise<SendDraftsResult> {
    const output = await runPowerShellJson(`${draftsScript(request.account)}
$drafts = @()
foreach ($entry in $scan) {
    foreach ($it in $entry.folder.Items) {
        if (Test-DraftMatches $it $entry.includeUnstamped) { $drafts += $it }
    }
}
$sent = 0
$failed = @()
foreach ($draft in $drafts) {
    $entryId = ''
    $subject = ''
    try { $entryId = [string]$draft.EntryID } catch {}
    try { $subject = [string]$draft.Subject } catch {}
    try {
        $draft.Send()
        $sent++
    } catch {
        $failed += [PSCustomObject]@{ entryId = $entryId; subject = $subject; error = $_.Exception.Message }
    }
}
ConvertTo-Json -Compress -Depth 3 -InputObject ([PSCustomObject]@{ sent = $sent; failed = @($failed) })
`, 'scan');
    const e = record(output);
    return {sent: num(e.sent), failed: itemFailures(e.failed)};
}

/**
 * Delete the named drafts, each proved to be this account's first. Outlook's
 * Delete() moves an item to Deleted Items, so a mistake stays recoverable.
 * Without the proof this would be a general "delete any email by id" call,
 * which is not what it is for.
 */
export async function deleteOutlookDrafts(request: EntryIdsRequest): Promise<DeleteDraftsResult> {
    const output = await runPowerShellJson(`${draftsScript(request.account)}
$deleted = 0
$failed = @()
foreach ($id in ${psArray(request.entryIds)}) {
    $subject = ''
    try {
        $draft = Resolve-OwnDraft $id 'delete'
        try { $subject = [string]$draft.Subject } catch {}
        $draft.Delete()
        $deleted++
    } catch {
        $failed += [PSCustomObject]@{ entryId = $id; subject = $subject; error = $_.Exception.Message }
    }
}
ConvertTo-Json -Compress -Depth 3 -InputObject ([PSCustomObject]@{ deleted = $deleted; failed = @($failed) })
`, 'scan');
    const e = record(output);
    return {deleted: num(e.deleted), failed: itemFailures(e.failed)};
}
