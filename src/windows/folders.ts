// Listing the folders under an Inbox, and filing mail into one.
import { psArray, psInt, runPowerShellJson } from './run';
import { accountScript, DELIVERY_STORE_PS, mailScopeScript } from './scripts';
import { bool, itemFailures, num, record, str, toArray } from '../shared/json';
import { FolderId } from '../mail';
import type { ListFoldersRequest, MoveRequest } from '../backend';
import type { InboxFolderInfo, MoveEmailsResult } from '../types';

/** The folders under an account's Inbox, depth-first. */
export async function listInboxFolders(request: ListFoldersRequest): Promise<InboxFolderInfo[]> {
    const output = await runPowerShellJson(`${accountScript(request.account)}
${DELIVERY_STORE_PS}
$inbox = $store.GetDefaultFolder(${psInt(FolderId.Inbox)})
$maxDepth = ${psInt(request.maxDepth)}
$rows = New-Object System.Collections.ArrayList
function Add-Folders($folder, [int]$level) {
    foreach ($f in $folder.Folders) {
        $itemCount = 0
        try { $itemCount = [int]$f.Items.Count } catch {}
        [void]$rows.Add([PSCustomObject]@{
            name       = [string]$f.Name
            folderPath = [string]$f.FolderPath
            itemCount  = $itemCount
            depth      = $level
        })
        if ($level -lt $maxDepth) { Add-Folders $f ($level + 1) }
    }
}
Add-Folders $inbox 1
ConvertTo-Json -Depth 3 -InputObject @($rows)
`, 'standard');
    return toArray(output).map(row => {
        const e = record(row);
        return {
            name: str(e.name),
            folderPath: str(e.folderPath),
            itemCount: num(e.itemCount),
            depth: num(e.depth) || 1,
        };
    });
}

/**
 * File emails into a folder of the account: one under the Inbox, or a
 * well-known one by name. Moving to Deleted Items is how mail is deleted
 * recoverably. Moving rewrites each item's entry id, so the ids passed in are
 * dead afterwards.
 */
export async function moveOutlookEmails(request: MoveRequest): Promise<MoveEmailsResult> {
    const output = await runPowerShellJson(`${accountScript(request.account)}
${DELIVERY_STORE_PS}
${mailScopeScript(request.folder, request.folderLabel, request.createIfMissing)}
$moved = 0
$failed = @()
foreach ($id in ${psArray(request.entryIds)}) {
    $subject = ''
    try {
        $it = $ns.GetItemFromID($id, $store.StoreID)
        try { $subject = [string]$it.Subject } catch {}
        [void]$it.Move($scope)
        $moved++
    } catch {
        $failed += [PSCustomObject]@{ entryId = $id; subject = $subject; error = $_.Exception.Message }
    }
}
ConvertTo-Json -Depth 3 -InputObject ([PSCustomObject]@{
    folderPath    = [string]$scope.FolderPath
    folderCreated = $scopeCreated
    moved         = $moved
    failed        = @($failed)
})
`, 'scan');
    const e = record(output);
    return {
        folderPath: str(e.folderPath),
        folderCreated: bool(e.folderCreated),
        moved: num(e.moved),
        failed: itemFailures(e.failed),
    };
}
