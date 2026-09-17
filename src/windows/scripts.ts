// The PowerShell fragments more than one Windows operation is built from.
//
// A fragment lives here once two modules emit it; one used by a single
// operation stays beside that operation, where its rules read next to the code
// relying on them. Every fragment follows run.ts's escaping rule: caller text
// only ever appears as a psString literal.
import { psInt, psString } from './run';
import { failureTag } from '../errors';
import type { EmailLocator, MailFolderRef } from '../types';

/**
 * Open the MAPI session, binding `$outlook` and `$ns`.
 *
 * `Logon()` is a no-op against an Outlook that is already running and starts
 * one that isn't, which removes a class of "works only while Outlook is open"
 * failure.
 */
export const SESSION_PS = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('MAPI')
$ns.Logon()
`;

/**
 * The session plus the mailbox an address names; binds `$target`, `$account`
 * and `$storeFolder`.
 *
 * An address can name a mailbox two ways, and only one is an Account. A mailbox
 * mounted as a secondary store — a shared mailbox, or one added after Outlook
 * started — has no Account object, so both are resolved and EITHER is enough:
 * `$account` is null for a store-only mailbox, and the operations that need an
 * identity to act as (sending) check for it themselves.
 *
 * Never falls back to the default account: a bad address must not silently
 * read or send from the wrong mailbox.
 *
 * The store is matched on the address BEFORE the account's display name, so a
 * profile where the two differ selects the same folders it always has.
 */
export function accountScript(emailAccount: string): string {
    return `${SESSION_PS}
$target = ${psString(emailAccount)}
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
if ($account -eq $null -and $storeFolder -eq $null) {
    throw "${failureTag('ACCOUNT_NOT_FOUND')}Account '$target' not found in this Outlook profile."
}
`;
}

/**
 * The mailbox's store as `$store` (and its id as `$storeId`), reached through
 * the account's delivery store. Requires `accountScript`.
 *
 * Falls back to the resolved store folder, because a store-only mailbox has no
 * account to read a delivery store off.
 */
export const DELIVERY_STORE_PS = `
$store = if ($account -ne $null) { $account.DeliveryStore } else { $storeFolder.Store }
$storeId = $store.StoreID
`;

/**
 * The mailbox's store as `$store`, reached through the namespace's top-level
 * folder for it. Requires `accountScript`.
 *
 * Kept beside DELIVERY_STORE_PS rather than replaced by it: the readers resolve
 * their store this way, and on a profile where a mailbox is open under a display
 * name that differs from its delivery store's, the two routes do not select the
 * same folders. Each operation keeps the route it was verified against.
 */
export const NAMED_STORE_PS = `
if ($storeFolder -eq $null) {
    throw "${failureTag('ACCOUNT_NOT_FOUND')}Account '$target' has no mailbox open in this Outlook profile."
}
$store = $storeFolder.Store
$storeId = $store.StoreID
`;

/**
 * Resolve one email into `$<variable>`, or fail as NOT_FOUND.
 *
 * The StoreID is tried first, since it lets an id from any mounted mailbox
 * resolve; the bare lookup (default store) is the fallback for a caller holding
 * a stale StoreID or none.
 */
export function itemLookupScript(email: EmailLocator, variable = 'item'): string {
    const byStore = email.storeId
        ? `try { $${variable} = $ns.GetItemFromID($lookupId, ${psString(email.storeId)}) } catch { $${variable} = $null }`
        : '';
    return `
$lookupId = ${psString(email.entryId)}
$${variable} = $null
${byStore}
if ($${variable} -eq $null) { try { $${variable} = $ns.GetItemFromID($lookupId) } catch { $${variable} = $null } }
if ($${variable} -eq $null) {
    throw "${failureTag('NOT_FOUND', 'email')}No email found for entry id '$lookupId'. It may have been moved or deleted - list the mail again for a current id."
}
`;
}

/**
 * The sender's SMTP address as `$senderSmtp`. Requires `$item`.
 *
 * An Exchange sender's SenderEmailAddress is an X.500 DN, not an address; a
 * caller replying to it or matching it against anything needs the address.
 */
export const SENDER_SMTP_PS = `
$senderSmtp = ''
try {
    if ($item.SenderEmailType -eq 'EX') { $senderSmtp = [string]$item.Sender.GetExchangeUser().PrimarySmtpAddress }
} catch {}
if (-not $senderSmtp) { try { $senderSmtp = [string]$item.SenderEmailAddress } catch {} }
`;

/** Recursive, case-insensitive folder search by name, depth-capped so a huge tree can't hang it. */
export const FIND_FOLDER_PS = `
function Find-FolderByName($root, [string]$name, [int]$depth) {
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
 * `$cutoff`, the moment `daysBack` days ago, and `$cutoffFilter`, that moment
 * as a date string for an `Items.Restrict` filter.
 *
 * Restrict parses date strings in the machine's own regional format, so the
 * string is formatted in it ('g': short date and time) rather than a fixed
 * pattern — 'MM/dd/yyyy' reads as a different day on a day-first locale. Every
 * reader re-checks each item against `$cutoff` too, so the window holds even
 * where a store reads the filter loosely.
 */
export function cutoffScript(daysBack: number): string {
    return `
$cutoff = (Get-Date).AddDays(-${psInt(daysBack)})
$cutoffFilter = $cutoff.ToString('g').Replace("'", "''")
`;
}

/**
 * `$<variable>` = the items of `$<folderVariable>` with `property` on or after
 * `$cutoff`. Requires `cutoffScript`. A store that rejects the filter outright
 * yields every item instead, which the per-item re-check then narrows.
 */
export function itemsSinceScript(folderVariable: string, property: string, variable: string): string {
    return `
$${variable} = $null
try { $${variable} = $${folderVariable}.Items.Restrict("[${property}] >= '$cutoffFilter'") } catch { $${variable} = $${folderVariable}.Items }
`;
}

/**
 * Resolve `$scope` from `$store`: a well-known root walked down through any
 * further segments. Also binds `$scopeCreated`.
 *
 * Segments match direct children; a bare name that isn't a direct child falls
 * back to a recursive search, so one folder string works across operations. A
 * missing folder fails as NOT_FOUND naming what the caller typed and what the
 * deepest folder reached does hold — never a silent fall back to the root. With
 * `createMissing`, the whole missing chain is created instead.
 */
export function mailScopeScript(ref: MailFolderRef, folderLabel: string, createMissing = false): string {
    const root = `$store.GetDefaultFolder(${psInt(ref.rootId)})`;
    if (ref.segments.length === 0) return `$scope = ${root}\n$scopeCreated = $false`;
    return `
${FIND_FOLDER_PS}
$folderLabel = ${psString(folderLabel)}
$segments = @(${ref.segments.map(psString).join(', ')})
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
if ($scope -eq $null -and $segments.Count -eq 1) {
    $scope = Find-FolderByName $scopeRoot $segments[0] 3
}
if ($scope -eq $null -and ${createMissing ? '$true' : '$false'}) {
    # Rebuild the chain from the root, creating only what is genuinely absent.
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
    throw "${failureTag('NOT_FOUND', 'folder')}Folder '$folderLabel' not found. Folders under '$($deepest.Name)': $($names -join ', ')"
}`;
}
