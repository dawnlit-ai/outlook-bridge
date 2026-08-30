// Listing the folders under an Inbox, and filing mail into one.
import { psList, requireWindows, runPowerShell } from './run';
import { accountScript, DELIVERY_STORE_PS, mailScopeScript } from './scripts';
import { num, parseArray, parseObject, record, str, toArray } from '../shared/json';
import { clamp, mailFolderRef } from '../mail';
import type { InboxFolderInfo, MoveEmailsResult } from '../types';

/** List the folders under an account's Inbox (the user's filing folders). */
export async function listInboxFolders(
    emailAccount: string,
    maxDepth = 2,
): Promise<InboxFolderInfo[]> {
    if (process.platform !== 'win32') return [];
    const depth = clamp(maxDepth, 1, 4);
    const script = `${accountScript(emailAccount)}
${DELIVERY_STORE_PS}
$inbox = $store.GetDefaultFolder(6)
function Walk-Folders($folder, $level) {
    foreach ($f in $folder.Folders) {
        [PSCustomObject]@{
            name       = $f.Name
            folderPath = $f.FolderPath
            itemCount  = $f.Items.Count
            depth      = $level
        }
        if ($level -lt ${depth}) { Walk-Folders $f ($level + 1) }
    }
}
$results = @(Walk-Folders $inbox 1)
ConvertTo-Json $results -Depth 3
`;
    return parseArray(await runPowerShell(script, 60000)).map(item => {
        const e = record(item);
        return {
            name: str(e.name),
            folderPath: str(e.folderPath),
            itemCount: num(e.itemCount),
            depth: num(e.depth) || 1,
        };
    });
}

/**
 * Move emails (by EntryID) into any folder of the account — a filing folder under
 * the Inbox, or a well-known folder by name (see WELL_KNOWN_FOLDERS).
 *
 * **Moving to "Deleted Items" is how mail gets deleted reversibly**, which is why
 * this takes well-known roots at all: it means the destructive path and the filing
 * path are one tool, and the destructive one is undoable from the folder it lands in.
 *
 * `createIfMissing` builds the WHOLE missing chain, so a nested destination
 * ("Clients\\Acme\\2026") is one call rather than a manual mkdir first.
 * NOTE: moving changes an item's EntryID — the passed ids are dead afterwards.
 */
export async function moveOutlookEmails(
    emailAccount: string,
    entryIds: string[],
    folderName: string,
    createIfMissing = false,
): Promise<MoveEmailsResult> {
    requireWindows();
    if (entryIds.length === 0) {
        return { folderPath: '', folderCreated: false, moved: 0, failed: [] };
    }
    const ref = mailFolderRef(folderName);
    const script = `${accountScript(emailAccount)}
${DELIVERY_STORE_PS}
${mailScopeScript(ref, folderName, createIfMissing)}
$folder = $scope
$folderCreated = $scopeCreated
$moved = 0
$failed = @()
foreach ($id in @(${psList(entryIds)})) {
    try {
        $it = $ns.GetItemFromID($id, $store.StoreID)
        [void]$it.Move($folder)
        $moved++
    } catch {
        $failed += [PSCustomObject]@{ entryId = $id; error = $_.Exception.Message }
    }
}
ConvertTo-Json @{ folderPath = $folder.FolderPath; folderCreated = $folderCreated; moved = $moved; failed = @($failed) } -Depth 3
`;
    const parsed = parseObject(await runPowerShell(script, 300000));
    return {
        folderPath: str(parsed.folderPath),
        folderCreated: Boolean(parsed.folderCreated),
        moved: num(parsed.moved),
        failed: toArray(parsed.failed).map(f => {
            const e = record(f);
            return { entryId: str(e.entryId), error: str(e.error) };
        }),
    };
}
