// The PowerShell fragments more than one Windows operation is built from.
//
// A fragment lives here once at least two modules emit it; one used by a single
// operation stays beside that operation, where its rules are readable next to
// the code that depends on them.
import { psEscape } from './run';
import type { MailFolderRef } from '../types';

/**
 * Open a MAPI session. `$outlook` and `$ns` are what every later fragment reads.
 *
 * `Logon()` is a no-op against an Outlook that is already running and starts one
 * that isn't, so it costs nothing and removes a class of "works only when
 * Outlook happens to be open" failure.
 */
const PS_PRELUDE = `
$ErrorActionPreference = 'Stop'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
`;

/**
 * Session plus the mailbox the address names; binds `$target`, `$account` and
 * `$storeFolder`.
 *
 * An address can name a mailbox two different ways, and only one of them is an
 * Account. A mailbox opened as a secondary store has no Account object, so
 * resolving through Accounts alone closes every operation to it — including the
 * reads, which never needed an Account in the first place. Both are resolved
 * here and EITHER is enough: `$account` is null for a store-only mailbox, and
 * the operations that genuinely require an identity to act as (sending) check
 * for it themselves rather than being denied the mailbox up front.
 *
 * Fails rather than falling back to the default account. A bad address must
 * never silently send from — or read — the wrong mailbox, and the thrown
 * sentence is the one `classifyRunFailure` recognises to raise
 * `AccountNotFoundError`, so the wording is load-bearing.
 *
 * The store is matched on the target address BEFORE the account's display name,
 * which is what keeps the named-store route (see `namedStoreScript`) selecting
 * the folders it selected before on a profile where the two differ.
 */
export function accountScript(emailAccount: string): string {
    return `${PS_PRELUDE}
$target = '${psEscape(emailAccount)}'
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
$storeFolder = $null
foreach ($f in $ns.Folders) {
    if ($f.Name -ieq $target) { $storeFolder = $f; break }
}
if ($storeFolder -eq $null -and $account -ne $null) {
    foreach ($f in $ns.Folders) {
        if ($f.Name -ieq $account.DisplayName) { $storeFolder = $f; break }
    }
}
if ($account -eq $null -and $storeFolder -eq $null) { throw "Account '$target' not found" }
`;
}

/**
 * The mailbox's own store, as `$store`. Requires `accountScript`.
 *
 * Falls back to the resolved store folder because a store-only mailbox has no
 * account to read a DeliveryStore off — reaching through `$account` directly is
 * what makes an operation fail on exactly those mailboxes.
 */
export const DELIVERY_STORE_PS = `
$store = if ($account -ne $null) { $account.DeliveryStore } else { $storeFolder.Store }
$storeId = $store.StoreID
`;

/**
 * The store `accountScript` resolved through the namespace's top-level folders,
 * as `$store`.
 *
 * Kept alongside `DELIVERY_STORE_PS` rather than replaced by it: the reading
 * operations resolve the store this way, and on a profile where a mailbox is
 * open under a display name that differs from the delivery store's, the two do
 * not select the same folders. Changing which one a reader uses changes what it
 * returns, so each keeps the route it was verified against.
 *
 * The folder search itself now lives in `accountScript`, which has to attempt it
 * anyway to resolve a mailbox that has no account. This throws when it found
 * nothing, which is the case where the address named an account whose store is
 * not open in this profile.
 */
export function namedStoreScript(emailAccount: string): string {
    return `
if ($storeFolder -eq $null) { throw "Store folder not found for account '${psEscape(emailAccount)}'" }
$store = $storeFolder.Store
$storeId = $store.StoreID
`;
}

/**
 * Resolve one item into `$item`. A StoreID disambiguates across mailboxes, so
 * pass it whenever the listing that produced the id carried one.
 */
export function getItemScript(entryId: string, storeId?: string): string {
    return storeId
        ? `$item = $ns.GetItemFromID('${psEscape(entryId)}', '${psEscape(storeId)}')`
        : `$item = $ns.GetItemFromID('${psEscape(entryId)}')`;
}

/**
 * The sender's real SMTP address, as `$senderSmtp`. Requires `$item`.
 *
 * An Exchange sender's `SenderEmailAddress` is an X500 DN rather than an
 * address; a caller that means to reply to it, or correlate it with anything,
 * needs the address it resolves to.
 */
export const SENDER_SMTP_PS = `
$senderSmtp = ''
try {
    if ($item.SenderEmailType -eq 'EX') { $senderSmtp = $item.Sender.GetExchangeUser().PrimarySmtpAddress }
} catch {}
if (-not $senderSmtp) { $senderSmtp = if ($item.SenderEmailAddress) { $item.SenderEmailAddress } else { '' } }
`;

/**
 * Recursive folder-by-name search. Case-insensitive, depth-capped so a huge
 * mailbox tree can't hang the scan.
 */
export const FIND_FOLDER_PS = `
function Find-FolderByName($root, $name, $depth) {
    foreach ($f in $root.Folders) {
        if ($f.Name -ieq $name) { return $f }
    }
    if ($depth -le 1) { return $null }
    foreach ($f in $root.Folders) {
        $hit = Find-FolderByName $f $name ($depth - 1)
        if ($hit -ne $null) { return $hit }
    }
    return $null
}
`;

/**
 * Emit the PowerShell that resolves `$scope` from `$store`, walking a well-known
 * root down through any further segments. Also binds `$scopeCreated`.
 *
 * Split out so the emitted script can be parse-checked without an Outlook
 * session — the folder walk is the only part that varies per call, and a syntax
 * error in it would only ever surface as a failed live run.
 *
 * `folderLabel` is the caller's original string, used verbatim in the error so
 * the message names what they typed rather than the parsed segments. Set
 * `createMissing` to build absent segments instead of throwing — the whole
 * chain, so a nested destination can be created in one call.
 */
export function mailScopeScript(ref: MailFolderRef, folderLabel = '', createMissing = false): string {
    const root = `$store.GetDefaultFolder(${ref.rootId})`;
    if (ref.segments.length === 0) return `$scope = ${root}\n$scopeCreated = $false`;
    return `
${FIND_FOLDER_PS}
$segments = @(${ref.segments.map(s => `'${psEscape(s)}'`).join(',')})
$scopeRoot = ${root}
$scope = $scopeRoot
$deepest = $scopeRoot
$scopeCreated = $false
foreach ($seg in $segments) {
    $next = $null
    foreach ($f in $scope.Folders) { if ($f.Name -ieq $seg) { $next = $f; break } }
    if ($next -eq $null) { $deepest = $scope; $scope = $null; break }
    $scope = $next
}
# A bare name that isn't a direct child still resolves by recursive search, so one
# folder string keeps working across tools.
if ($scope -eq $null -and $segments.Count -eq 1) {
    $scope = Find-FolderByName $scopeRoot $segments[0] 3
}
if ($scope -eq $null -and ${createMissing ? '$true' : '$false'}) {
    # Rebuild the whole chain from the root, creating only what is genuinely absent.
    $scope = $scopeRoot
    foreach ($seg in $segments) {
        $next = $null
        foreach ($f in $scope.Folders) { if ($f.Name -ieq $seg) { $next = $f; break } }
        if ($next -eq $null) { $next = $scope.Folders.Add($seg); $scopeCreated = $true }
        $scope = $next
    }
}
if ($scope -eq $null) {
    $names = @()
    foreach ($f in $deepest.Folders) { $names += $f.Name }
    throw "Folder '${psEscape(folderLabel)}' not found. Folders under '$($deepest.Name)': $($names -join ', ')"
}`;
}
