// Deleting mail, and emptying what was already deleted.
import { psArray, psBool, psInt, psString, runPowerShellJson } from './run';
import { accountScript, cutoffScript, DELIVERY_STORE_PS } from './scripts';
import { num, record, str, toArray } from '../shared/json';
import { FolderId, PROTECTED_MAIL_REASON } from '../mail';
import type { DeleteEmailsRequest, PurgeRequest } from '../backend';
import type { DeleteMailOutcome, DeleteMailResult, PurgeDeletedItemsResult } from '../types';

/** Folders whose contents are received or sent mail rather than working state. */
const PROTECTED_FOLDER_IDS = [FolderId.Inbox, FolderId.SentMail];

const OUTCOME_STATUSES: readonly DeleteMailOutcome['status'][] = ['deleted', 'would-delete', 'refused', 'failed'];

/**
 * Delete mail by entry id from anywhere in the account. Delete() moves each item
 * to Deleted Items, so this is recoverable — purgeDeletedItems is what destroys.
 *
 * This is the one operation that can reach received mail, and an entry id is an
 * unsafe key for it: a stale or wrong id can resolve to a DIFFERENT message.
 * Three things hold the line:
 *  - Inbox and Sent Items are refused, INCLUDING their subfolders (a filed
 *    folder is still received mail), unless `allowProtected`.
 *  - `dryRun` resolves and reports without deleting, so the exact subjects and
 *    folders can be checked before anything happens.
 *  - Every outcome echoes the subject and folder of the item actually resolved.
 */
export async function deleteOutlookEmails(request: DeleteEmailsRequest): Promise<DeleteMailResult> {
    const output = await runPowerShellJson(`${accountScript(request.account)}
${DELIVERY_STORE_PS}
$allowProtected = ${psBool(request.allowProtected)}
$dryRun = ${psBool(request.dryRun)}
$protectedReason = ${psString(PROTECTED_MAIL_REASON)}
$protectedIds = @{}
foreach ($role in @(${PROTECTED_FOLDER_IDS.map(psInt).join(', ')})) {
    try { $protectedIds[[string]$store.GetDefaultFolder($role).EntryID] = $true } catch {}
}
function Test-UnderProtectedFolder($folder) {
    $current = $folder
    for ($guard = 0; $current -ne $null -and $guard -lt 25; $guard++) {
        $folderId = $null
        try { $folderId = [string]$current.EntryID } catch {}
        if (-not $folderId) { return $false }
        if ($protectedIds.ContainsKey($folderId)) { return $true }
        $parent = $null
        try { $parent = $current.Parent } catch {}
        $current = $parent
    }
    return $false
}
$items = @()
foreach ($id in ${psArray(request.entryIds)}) {
    $subject = ''
    $folderPath = ''
    try {
        $it = $null
        try { $it = $ns.GetItemFromID($id, $store.StoreID) } catch { $it = $null }
        if ($it -eq $null) { try { $it = $ns.GetItemFromID($id) } catch { $it = $null } }
        if ($it -eq $null) { throw "no item with that entry id in this account" }
        try { $subject = [string]$it.Subject } catch {}
        $parent = $null
        try { $parent = $it.Parent } catch {}
        if ($parent -eq $null) { throw "item has no parent folder - refusing to delete" }
        try { $folderPath = [string]$parent.FolderPath } catch {}
        if ((Test-UnderProtectedFolder $parent) -and -not $allowProtected) {
            $items += [PSCustomObject]@{ entryId = $id; subject = $subject; folderPath = $folderPath; status = 'refused'; reason = $protectedReason }
        } elseif ($dryRun) {
            $items += [PSCustomObject]@{ entryId = $id; subject = $subject; folderPath = $folderPath; status = 'would-delete'; reason = '' }
        } else {
            $it.Delete()
            $items += [PSCustomObject]@{ entryId = $id; subject = $subject; folderPath = $folderPath; status = 'deleted'; reason = '' }
        }
    } catch {
        $items += [PSCustomObject]@{ entryId = $id; subject = $subject; folderPath = $folderPath; status = 'failed'; reason = $_.Exception.Message }
    }
}
ConvertTo-Json -Depth 3 -InputObject @($items)
`, 'scan');
    const items: DeleteMailOutcome[] = toArray(output).map(row => {
        const e = record(row);
        const status = OUTCOME_STATUSES.find(s => s === str(e.status)) ?? 'failed';
        return {
            entryId: str(e.entryId),
            subject: str(e.subject),
            folderPath: str(e.folderPath),
            status,
            reason: str(e.reason),
        };
    });
    return {
        dryRun: request.dryRun,
        deleted: items.filter(i => i.status === 'deleted').length,
        refused: items.filter(i => i.status === 'refused').length,
        failed: items.filter(i => i.status === 'failed').length,
        items,
    };
}

/**
 * Permanently remove items from the account's Deleted Items. The ONE irreversible
 * operation here, which is why it is folder-scoped rather than keyed on ids: it
 * can only ever destroy what was already thrown away.
 *
 * Walks backwards, since deleting shifts the collection and a forward walk would
 * skip every other item.
 */
export async function purgeDeletedItems(request: PurgeRequest): Promise<PurgeDeletedItemsResult> {
    const output = await runPowerShellJson(`${accountScript(request.account)}
${DELIVERY_STORE_PS}
$folder = $store.GetDefaultFolder(${psInt(FolderId.DeletedItems)})
$olderThanDays = ${psInt(request.olderThanDays)}
$dryRun = ${psBool(request.dryRun)}
${cutoffScript(request.olderThanDays)}
$matched = 0
$purged = 0
$kept = 0
$failed = 0
# Bound once: each $folder.Items is a COM call that builds a fresh collection.
$folderItems = $folder.Items
for ($i = $folderItems.Count; $i -ge 1; $i--) {
    $it = $null
    try { $it = $folderItems.Item($i) } catch { continue }
    $stamp = $null
    foreach ($property in @('ReceivedTime', 'LastModificationTime', 'CreationTime')) {
        try { $stamp = $it.$property; if ($stamp -ne $null) { break } } catch {}
    }
    if ($olderThanDays -gt 0 -and $stamp -ne $null -and $stamp -gt $cutoff) { $kept++; continue }
    $matched++
    if ($dryRun) { continue }
    try { $it.Delete(); $purged++ } catch { $failed++ }
}
ConvertTo-Json -Compress -InputObject ([PSCustomObject]@{
    folderPath = [string]$folder.FolderPath
    matched    = $matched
    purged     = $purged
    kept       = $kept
    failed     = $failed
})
`, 'purge');
    const e = record(output);
    return {
        folderPath: str(e.folderPath),
        dryRun: request.dryRun,
        matched: num(e.matched),
        purged: num(e.purged),
        kept: num(e.kept),
        failed: num(e.failed),
    };
}
