// Deleting mail, and emptying what was already deleted.
import { psBool, psList, requireWindows, runPowerShell } from './run';
import { accountScript, DELIVERY_STORE_PS } from './scripts';
import { num, parseObject, record, str, toArray } from '../shared/json';
import type {
    DeleteMailOptions,
    DeleteMailOutcome,
    DeleteMailResult,
    PurgeDeletedItemsResult,
} from '../types';

/** Folders whose contents are received or already-sent mail, not working state. */
const PROTECTED_FOLDER_IDS = [6, 5];

/** Walks an item's parent chain, so Inbox/Sent protection covers their subfolders too. */
const FOLDER_CHAIN_PS = `
function Get-FolderChainIds($folder) {
    $ids = @()
    $cur = $folder
    $guard = 0
    while ($cur -ne $null -and $guard -lt 25) {
        $cid = $null
        try { $cid = $cur.EntryID } catch { $cid = $null }
        if (-not $cid) { break }
        $ids += $cid
        $parent = $null
        try { $parent = $cur.Parent } catch { $parent = $null }
        if ($parent -eq $null) { break }
        $cur = $parent
        $guard++
    }
    return $ids
}
`;

/**
 * Delete mail by EntryID from anywhere in the account. Delete() moves each item to
 * Deleted Items, so this is recoverable — purgeDeletedItems is what destroys.
 *
 * ⚠️ This is the one tool here that can reach received mail, and EntryIDs are a
 * genuinely unsafe key for it: GetItemFromID silently returns a DIFFERENT message
 * when handed a stale or wrong id, which is common precisely when many replies share
 * one subject (see saveEmailAttachmentDetailed). Three things hold the line:
 *
 *  - **Inbox and Sent Items are refused by default, INCLUDING their subfolders** —
 *    a filed subfolder is still received mail. `allowProtected` lifts that, and is
 *    the caller explicitly taking responsibility.
 *  - **`dryRun` resolves and reports without deleting**, so the exact subjects and
 *    folders can be shown to the user before anything happens. Use it first.
 *  - **Every outcome echoes the subject and folderPath** of the item actually
 *    resolved, so a wrong id is visible after the fact rather than silent.
 */
export async function deleteOutlookEmails(
    emailAccount: string,
    entryIds: string[],
    options: DeleteMailOptions = {},
): Promise<DeleteMailResult> {
    requireWindows();
    const {allowProtected = false, dryRun = false} = options;
    if (entryIds.length === 0) {
        return {dryRun, deleted: 0, refused: 0, failed: 0, items: []};
    }
    const script = `${accountScript(emailAccount)}
${DELIVERY_STORE_PS}
${FOLDER_CHAIN_PS}
$protectedIds = @{}
foreach ($fid in @(${PROTECTED_FOLDER_IDS.join(',')})) {
    try { $protectedIds[$store.GetDefaultFolder($fid).EntryID] = $true } catch {}
}

$items = @()
$deleted = 0
$refused = 0
$failed = 0
foreach ($id in @(${psList(entryIds)})) {
    $subject = ''
    $folderPath = ''
    try {
        $it = $null
        try { $it = $ns.GetItemFromID($id, $store.StoreID) } catch { $it = $null }
        if ($it -eq $null) { try { $it = $ns.GetItemFromID($id) } catch { $it = $null } }
        if ($it -eq $null) { throw "no item with that EntryID in this account" }
        try { $subject = $it.Subject } catch {}
        $parent = $null
        try { $parent = $it.Parent } catch {}
        if ($parent -eq $null) { throw "item has no parent folder - refusing to delete" }
        try { $folderPath = $parent.FolderPath } catch {}

        $chain = Get-FolderChainIds $parent
        $isProtected = $false
        foreach ($cid in $chain) { if ($protectedIds.ContainsKey($cid)) { $isProtected = $true; break } }

        if ($isProtected -and -not ${psBool(allowProtected)}) {
            $refused++
            $items += [pscustomobject]@{ entryId = $id; subject = $subject; folderPath = $folderPath; status = 'refused'; reason = 'received or sent mail (Inbox/Sent Items or a subfolder) - pass allow_protected to override' }
        } elseif (${psBool(dryRun)}) {
            $items += [pscustomobject]@{ entryId = $id; subject = $subject; folderPath = $folderPath; status = 'would-delete'; reason = '' }
        } else {
            $it.Delete()
            $deleted++
            $items += [pscustomobject]@{ entryId = $id; subject = $subject; folderPath = $folderPath; status = 'deleted'; reason = '' }
        }
    } catch {
        $failed++
        $items += [pscustomobject]@{ entryId = $id; subject = $subject; folderPath = $folderPath; status = 'failed'; reason = $_.Exception.Message }
    }
}
ConvertTo-Json @{ dryRun = ${psBool(dryRun)}; deleted = $deleted; refused = $refused; failed = $failed; items = @($items) } -Depth 4
`;
    const parsed = parseObject(await runPowerShell(script, 300000));
    return {
        dryRun,
        deleted: num(parsed.deleted),
        refused: num(parsed.refused),
        failed: num(parsed.failed),
        items: toArray(parsed.items).map(i => {
            const e = record(i);
            return {
                entryId: str(e.entryId),
                subject: str(e.subject),
                folderPath: str(e.folderPath),
                status: (str(e.status) || 'failed') as DeleteMailOutcome['status'],
                reason: str(e.reason),
            };
        }),
    };
}

/**
 * Permanently remove items from the account's Deleted Items folder. This is the ONE
 * genuinely irreversible operation here — nothing recovers from it — which is why it
 * is folder-scoped rather than keyed on an EntryID: it can only ever destroy what the
 * user already threw away.
 *
 * `olderThanDays` keeps recent items (0 = purge everything). Iterates backwards, as
 * deleting mutates the collection and a forward walk would skip every other item.
 */
export async function purgeDeletedItems(
    emailAccount: string,
    olderThanDays = 0,
    dryRun = false,
): Promise<PurgeDeletedItemsResult> {
    requireWindows();
    const script = `${accountScript(emailAccount)}
$folder = $account.DeliveryStore.GetDefaultFolder(3)
$cutoff = (Get-Date).AddDays(-${olderThanDays})
$matched = 0
$purged = 0
$kept = 0
$failed = 0
for ($i = $folder.Items.Count; $i -ge 1; $i--) {
    $it = $null
    try { $it = $folder.Items.Item($i) } catch { continue }
    $stamp = $null
    foreach ($p in @('ReceivedTime','LastModificationTime','CreationTime')) {
        try { $stamp = $it.$p; if ($stamp -ne $null) { break } } catch {}
    }
    if (${olderThanDays} -gt 0 -and $stamp -ne $null -and $stamp -gt $cutoff) { $kept++; continue }
    $matched++
    if (${psBool(dryRun)}) { continue }
    try { $it.Delete(); $purged++ } catch { $failed++ }
}
ConvertTo-Json @{ folderPath = $folder.FolderPath; matched = $matched; purged = $purged; kept = $kept; failed = $failed } -Depth 3
`;
    const parsed = parseObject(await runPowerShell(script, 600000));
    return {
        folderPath: str(parsed.folderPath),
        dryRun,
        matched: num(parsed.matched),
        purged: num(parsed.purged),
        kept: num(parsed.kept),
        failed: num(parsed.failed),
    };
}
