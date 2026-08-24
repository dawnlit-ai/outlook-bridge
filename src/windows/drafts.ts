// Drafts: which ones belong to an account, and sending, listing or deleting them.
import { psList, requireWindows, runPowerShell } from './run';
import { accountScript } from './scripts';
import { num, parseObject, record, str, strList, toArray } from '../shared/json';
import type { DeleteDraftsResult, ListDraftsResult, SendAllDraftsResult, SendDraftsResult } from '../types';

/**
 * Resolve the Drafts folders belonging to one account into `$scan`, and define
 * `Test-DraftMatches` over them. Emitted into every drafts script so listing,
 * deleting and sending can never disagree about which drafts are "this account's".
 * Expects `$target`, `$outlook`, `$ns` and `$account` to already be set.
 *
 * Two Drafts folders can hold drafts for one account, so both are scanned:
 *  - The account's OWN mailbox store Drafts folder (account.DeliveryStore) — where
 *    a draft composed while that mailbox is selected lands. A draft there with no
 *    explicit SendUsingAccount still belongs to this account, so null counts.
 *  - The DEFAULT store's Drafts folder — where the in-app "Create Drafts" flow
 *    files drafts via $mail.Save() regardless of send-account. Only drafts stamped
 *    with this account match here; a null SendUsingAccount means the default
 *    account, not necessarily this one, so it's left alone.
 * (When those two resolve to the same folder — the account IS the default — it is
 * scanned once, under the more permissive "own store" rule.)
 *
 * Drafts bound to another account never match, so nothing here can reach unrelated
 * mail. Non-mail items (meeting requests, reports) are ignored.
 */
const DRAFTS_SCAN_PS = `
# Collect the Drafts folders to scan, deduped by id. includeNull marks the account's
# own store, where a draft with no explicit send-account still belongs to it.
$scan = @()
$seen = @{}

$homeDrafts = $null
try { $homeDrafts = $account.DeliveryStore.GetDefaultFolder(16) } catch {}  # olFolderDrafts
if ($homeDrafts -ne $null) {
    $seen["$($homeDrafts.StoreID)|$($homeDrafts.EntryID)"] = $true
    $scan += [pscustomobject]@{ folder = $homeDrafts; includeNull = $true }
}

$defDrafts = $null
try { $defDrafts = $ns.GetDefaultFolder(16) } catch {}
if ($defDrafts -ne $null) {
    $k = "$($defDrafts.StoreID)|$($defDrafts.EntryID)"
    if (-not $seen.ContainsKey($k)) {
        $seen[$k] = $true
        $scan += [pscustomobject]@{ folder = $defDrafts; includeNull = $false }
    }
}

# SendUsingAccount is null for a draft that never set one, so guard the access.
function Test-DraftMatches($item, $includeNull) {
    if ($item.Class -ne 43) { return $false }  # olMail only
    $acct = $null
    try { $acct = $item.SendUsingAccount } catch {}
    if ($acct -eq $null) { return $includeNull }
    return ($acct.SmtpAddress -ieq $target)
}
`;

/** Session, account and the drafts scan rule — the prelude all three share. */
function draftsScript(emailAccount: string): string {
    return accountScript(emailAccount) + DRAFTS_SCAN_PS;
}

/**
 * Send the mail drafts that belong to `emailAccount`, wherever Outlook filed them.
 * Which drafts those are is DRAFTS_SCAN_PS's rule, not this function's.
 *
 * Item references are snapshotted before any Send() call: sending moves an item out
 * of Drafts, and mutating the collection mid-enumeration would skip every other one.
 */
export async function sendAllDrafts(emailAccount: string): Promise<SendAllDraftsResult> {
    requireWindows();
    const script = `${draftsScript(emailAccount)}
# Snapshot the matching mail items first — Send() removes each from Drafts, so
# sending inside a live enumeration would skip every other item.
$items = @()
foreach ($entry in $scan) {
    foreach ($it in $entry.folder.Items) {
        if (Test-DraftMatches $it $entry.includeNull) { $items += $it }
    }
}

$sent = 0
$failed = @()
foreach ($m in $items) {
    try {
        $m.Send()
        $sent++
    } catch {
        $subj = ''
        try { $subj = $m.Subject } catch {}
        $failed += [pscustomobject]@{ subject = $subj; error = $_.Exception.Message }
    }
}

[pscustomobject]@{ sent = $sent; failed = @($failed) } | ConvertTo-Json -Compress -Depth 4
`;
    try {
        const parsed = parseObject(await runPowerShell(script, 300000));
        return {
            sent: num(parsed.sent),
            failed: toArray(parsed.failed).map(f => {
                const e = record(f);
                return { subject: str(e.subject), error: str(e.error) };
            }),
        };
    } catch (error) {
        // A malformed report is not worth failing a send that already happened;
        // a run that never got that far has already rejected above.
        if (error instanceof SyntaxError) return { sent: 0, failed: [] };
        throw error;
    }
}

/**
 * List the mail drafts belonging to `emailAccount`, newest first. Which drafts
 * those are is DRAFTS_SCAN_PS's rule.
 *
 * Deliberately returns no StoreID: deleteOutlookDrafts re-resolves each EntryID
 * against the same folders, so repeating a ~700-char store id on every row would be
 * pure payload. Bodies are previewed, never returned whole — a templated reply body
 * runs to tens of thousands of characters and a folder's worth would blow the cap.
 */
export async function listOutlookDrafts(
    emailAccount: string,
    limit = 100,
    previewChars = 300,
): Promise<ListDraftsResult> {
    requireWindows();
    const script = `${draftsScript(emailAccount)}
$previewChars = ${previewChars}
$rows = @()
$folders = @()
foreach ($entry in $scan) {
    $fp = ''
    try { $fp = $entry.folder.FolderPath } catch {}
    $folders += $fp
    foreach ($it in $entry.folder.Items) {
        if (-not (Test-DraftMatches $it $entry.includeNull)) { continue }
        $subj = ''; try { $subj = $it.Subject } catch {}
        $to = ''; try { $to = $it.To } catch {}
        $addrs = @()
        try {
            foreach ($r in $it.Recipients) {
                $a = ''
                try { $a = $r.Address } catch {}
                if ($a -and $a.StartsWith('/')) { try { $a = $r.Name } catch {} }
                if ($a) { $addrs += $a }
            }
        } catch {}
        $body = ''
        try { $body = $it.Body } catch {}
        if ($body -eq $null) { $body = '' }
        $body = ($body -replace '\\s+', ' ').Trim()
        if ($body.Length -gt $previewChars) { $body = $body.Substring(0, $previewChars) }
        $mod = ''
        try { $mod = $it.LastModificationTime.ToString('yyyy-MM-dd HH:mm') } catch {}
        $att = $false
        try { $att = ($it.Attachments.Count -gt 0) } catch {}
        $rows += [pscustomobject]@{
            entryId = $it.EntryID
            subject = $subj
            to = $to
            toEmails = @($addrs)
            bodyPreview = $body
            hasAttachments = $att
            lastModified = $mod
            folderPath = $fp
        }
    }
}
$sorted = @($rows | Sort-Object -Property lastModified -Descending)
$total = $sorted.Count
if ($total -gt ${limit}) { $sorted = @($sorted[0..(${limit} - 1)]) }
ConvertTo-Json @{ account = $target; foldersScanned = @($folders); count = $total; truncated = ($total -gt ${limit}); drafts = @($sorted) } -Depth 4
`;
    const parsed = parseObject(await runPowerShell(script, 300000));
    return {
        account: str(parsed.account) || emailAccount,
        foldersScanned: toArray(parsed.foldersScanned).map(str),
        count: num(parsed.count),
        truncated: Boolean(parsed.truncated),
        drafts: toArray(parsed.drafts).map(d => {
            const e = record(d);
            return {
                entryId: str(e.entryId),
                subject: str(e.subject),
                to: str(e.to),
                toEmails: strList(e.toEmails),
                bodyPreview: str(e.bodyPreview),
                hasAttachments: Boolean(e.hasAttachments),
                lastModified: str(e.lastModified),
                folderPath: str(e.folderPath),
            };
        }),
    };
}

/**
 * Delete mail drafts by EntryID. Outlook's Delete() moves the item to Deleted Items
 * rather than destroying it, so a mistaken call stays recoverable from there.
 *
 * Every id must resolve to a mail item sitting in one of THIS account's Drafts
 * folders before anything is deleted; an id pointing at ordinary mail, or at another
 * account's draft, is refused and reported. Without that gate this would be a
 * general-purpose "delete any email by id" tool, which is not what it is for.
 */
export async function deleteOutlookDrafts(
    emailAccount: string,
    entryIds: string[],
): Promise<DeleteDraftsResult> {
    requireWindows();
    if (entryIds.length === 0) return { deleted: 0, failed: [] };
    const script = `${draftsScript(emailAccount)}
# Index this account's Drafts folders by EntryID, so a resolved item can be PROVED
# to live in one of them before it is deleted.
$draftFolderIds = @{}
foreach ($entry in $scan) { $draftFolderIds[$entry.folder.EntryID] = $entry.includeNull }

$deleted = 0
$failed = @()
foreach ($id in @(${psList(entryIds)})) {
    try {
        $it = $null
        foreach ($entry in $scan) {
            try {
                $it = $ns.GetItemFromID($id, $entry.folder.StoreID)
                if ($it -ne $null) { break }
            } catch { $it = $null }
        }
        if ($it -eq $null) { throw "no such item in this account's Drafts folders" }
        $parentId = ''
        try { $parentId = $it.Parent.EntryID } catch {}
        if (-not $draftFolderIds.ContainsKey($parentId)) {
            throw "item is not in this account's Drafts folder - refusing to delete"
        }
        if (-not (Test-DraftMatches $it $draftFolderIds[$parentId])) {
            throw "draft is not bound to $target - refusing to delete"
        }
        $it.Delete()
        $deleted++
    } catch {
        $failed += [pscustomobject]@{ entryId = $id; error = $_.Exception.Message }
    }
}
ConvertTo-Json @{ deleted = $deleted; failed = @($failed) } -Depth 3
`;
    const parsed = parseObject(await runPowerShell(script, 300000));
    return {
        deleted: num(parsed.deleted),
        failed: toArray(parsed.failed).map(f => {
            const e = record(f);
            return { entryId: str(e.entryId), error: str(e.error) };
        }),
    };
}

/**
 * Send a chosen subset of `emailAccount`'s drafts by EntryID — e.g. after
 * `listOutlookDrafts` and a user review pass — instead of `sendAllDrafts`'s
 * account-wide sweep.
 *
 * Same ownership gate as `deleteOutlookDrafts`: every id must resolve to a mail
 * item sitting in one of THIS account's Drafts folders before anything is sent.
 * An id pointing at ordinary inbox mail, or at another account's draft, is
 * refused and reported rather than sent.
 */
export async function sendDrafts(
    emailAccount: string,
    entryIds: string[],
): Promise<SendDraftsResult> {
    requireWindows();
    if (entryIds.length === 0) return { sent: 0, failed: [] };
    const script = `${draftsScript(emailAccount)}
# Index this account's Drafts folders by EntryID, so a resolved item can be PROVED
# to live in one of them before it is sent.
$draftFolderIds = @{}
foreach ($entry in $scan) { $draftFolderIds[$entry.folder.EntryID] = $entry.includeNull }

$sent = 0
$failed = @()
foreach ($id in @(${psList(entryIds)})) {
    $subj = ''
    try {
        $it = $null
        foreach ($entry in $scan) {
            try {
                $it = $ns.GetItemFromID($id, $entry.folder.StoreID)
                if ($it -ne $null) { break }
            } catch { $it = $null }
        }
        if ($it -eq $null) { throw "no such item in this account's Drafts folders" }
        try { $subj = $it.Subject } catch {}
        $parentId = ''
        try { $parentId = $it.Parent.EntryID } catch {}
        if (-not $draftFolderIds.ContainsKey($parentId)) {
            throw "item is not in this account's Drafts folder - refusing to send"
        }
        if (-not (Test-DraftMatches $it $draftFolderIds[$parentId])) {
            throw "draft is not bound to $target - refusing to send"
        }
        $it.Send()
        $sent++
    } catch {
        $failed += [pscustomobject]@{ entryId = $id; subject = $subj; error = $_.Exception.Message }
    }
}
[pscustomobject]@{ sent = $sent; failed = @($failed) } | ConvertTo-Json -Compress -Depth 4
`;
    const parsed = parseObject(await runPowerShell(script, 300000));
    return {
        sent: num(parsed.sent),
        failed: toArray(parsed.failed).map(f => {
            const e = record(f);
            return { entryId: str(e.entryId), subject: str(e.subject), error: str(e.error) };
        }),
    };
}
