// Bounce-backs: finding them, clearing them away, and mining them for the
// addresses that failed.
import { psArray, psBool, psInt, psString, runPowerShellJson } from './run';
import { accountScript, cutoffScript, DELIVERY_STORE_PS, itemsSinceScript, NAMED_STORE_PS } from './scripts';
import { itemFailures, num, record, str, strList, toArray } from '../shared/json';
import {
    BODY_ADDRESS_PATTERN,
    BOUNCE_DAEMON_ADDRESSES,
    BOUNCE_DAEMON_NAMES,
    BOUNCE_REASON,
    BOUNCE_SUBJECT_PHRASES,
    DAEMON_SELF_ADDRESSES,
    MAX_FAILED_RECIPIENTS,
    NDR_MESSAGE_CLASS_PREFIX,
} from '../shared/bounceRules';
import { FolderId } from '../mail';
import type { CleanUndeliverableRequest, CollectBouncesRequest, SentGroupsRequest } from '../backend';
import type { CleanUndeliverableResult, SentRecipientGroup } from '../types';

/**
 * The classifier, generated from the shared rules so Windows and macOS can't
 * disagree about what a bounce is. Defines `Get-BounceReport`, which returns
 * null for an item that isn't a bounce and a report row for one that is.
 *
 * The body is read only after a cheaper signal (class, sender, subject) has
 * matched — that ordering is what keeps a large-inbox scan affordable.
 */
function bounceClassifierScript(): string {
    return `
$bouncePhrases = ${psArray(BOUNCE_SUBJECT_PHRASES)}
$daemonAddresses = ${psArray(BOUNCE_DAEMON_ADDRESSES)}
$daemonNames = ${psArray(BOUNCE_DAEMON_NAMES)}
$daemonSelf = ${psArray(DAEMON_SELF_ADDRESSES)}
$ndrPrefix = ${psString(NDR_MESSAGE_CLASS_PREFIX)}
$addressPattern = ${psString(BODY_ADDRESS_PATTERN)}
$maxFailed = ${psInt(MAX_FAILED_RECIPIENTS)}
$accountLower = $target.ToLower()
function Get-BounceReport($item) {
    $class = 0
    try { $class = [int]$item.Class } catch {}
    if ($class -ne 43 -and $class -ne 46) { return $null }
    $subject = ''
    try { $subject = [string]$item.Subject } catch {}
    $senderName = ''
    try { $senderName = [string]$item.SenderName } catch {}
    $senderEmail = ''
    try { $senderEmail = [string]$item.SenderEmailAddress } catch {}
    $messageClass = ''
    try { $messageClass = [string]$item.MessageClass } catch {}
    $reason = ''
    if ($messageClass.StartsWith($ndrPrefix, [StringComparison]::OrdinalIgnoreCase)) { $reason = ${psString(BOUNCE_REASON.ndr)} }
    if (-not $reason) { foreach ($d in $daemonAddresses) { if ($senderEmail.ToLower().Contains($d)) { $reason = ${psString(BOUNCE_REASON.daemon)}; break } } }
    if (-not $reason) { foreach ($d in $daemonNames) { if ($senderName.ToLower().Contains($d)) { $reason = ${psString(BOUNCE_REASON.daemon)}; break } } }
    if (-not $reason) { foreach ($p in $bouncePhrases) { if ($subject.ToLower().Contains($p)) { $reason = "${BOUNCE_REASON.subjectPhrase('$p')}"; break } } }
    if (-not $reason) { return $null }
    $failedRecipients = @()
    $body = ''
    try { $body = [string]$item.Body } catch {}
    foreach ($m in [regex]::Matches($body, $addressPattern)) {
        $address = $m.Value.ToLower().TrimEnd('.')
        if ($address -eq $accountLower) { continue }
        $isDaemon = $false
        foreach ($d in $daemonSelf) { if ($address.Contains($d)) { $isDaemon = $true; break } }
        if ($isDaemon -or $failedRecipients -contains $address) { continue }
        $failedRecipients += $address
        if ($failedRecipients.Count -ge $maxFailed) { break }
    }
    $received = ''
    try { $received = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } catch {}
    return [PSCustomObject]@{
        entryId          = [string]$item.EntryID
        subject          = $subject.Trim()
        senderName       = $senderName
        senderEmail      = $senderEmail
        receivedTime     = $received
        matchedReason    = $reason
        failedRecipients = @($failedRecipients)
    }
}
`;
}

/**
 * Scan an account's Inbox for bounce-backs and, unless dry-running, move each to
 * Deleted Items (recoverable).
 *
 * Matched items are collected before the first Delete(): deleting shifts the
 * collection, and deleting mid-enumeration would skip items.
 */
export async function cleanUndeliverableEmails(request: CleanUndeliverableRequest): Promise<CleanUndeliverableResult> {
    const output = await runPowerShellJson(`${accountScript(request.account)}
${NAMED_STORE_PS}
${bounceClassifierScript()}
$dryRun = ${psBool(request.dryRun)}
$inbox = $store.GetDefaultFolder(${psInt(FolderId.Inbox)})
${cutoffScript(request.daysBack)}
${itemsSinceScript('inbox', 'ReceivedTime', 'filtered')}
$reports = @()
$matchedItems = @()
$count = $filtered.Count
for ($i = 1; $i -le $count; $i++) {
    $item = $null
    try { $item = $filtered.Item($i) } catch { continue }
    if ($item -eq $null) { continue }
    $received = $null
    try { $received = $item.ReceivedTime } catch {}
    if ($received -ne $null -and $received -lt $cutoff) { continue }
    $report = Get-BounceReport $item
    if ($report -eq $null) { continue }
    $reports += $report
    $matchedItems += $item
}
$deleted = 0
$failed = @()
if (-not $dryRun) {
    foreach ($m in $matchedItems) {
        $entryId = ''
        $subject = ''
        try { $entryId = [string]$m.EntryID } catch {}
        try { $subject = [string]$m.Subject } catch {}
        try { $m.Delete(); $deleted++ }
        catch { $failed += [PSCustomObject]@{ entryId = $entryId; subject = $subject; error = $_.Exception.Message } }
    }
}
ConvertTo-Json -Depth 4 -InputObject ([PSCustomObject]@{
    deletedCount = $deleted
    matched      = @($reports)
    failed       = @($failed)
})
`, 'scan');
    const e = record(output);
    const matched = toArray(e.matched).map(row => {
        const m = record(row);
        return {
            entryId: str(m.entryId),
            subject: str(m.subject),
            senderName: str(m.senderName),
            senderEmail: str(m.senderEmail),
            receivedTime: str(m.receivedTime),
            matchedReason: str(m.matchedReason),
            failedRecipients: strList(m.failedRecipients),
        };
    });
    return {
        account: request.account,
        scannedDays: request.daysBack,
        dryRun: request.dryRun,
        matchedCount: matched.length,
        deletedCount: num(e.deletedCount),
        matched,
        failed: itemFailures(e.failed),
    };
}

/**
 * The deduplicated addresses bounce-backs report as failed, from the Inbox and,
 * optionally, Deleted Items — so it still works after the bounces were cleaned
 * away. Never deletes anything.
 */
export async function collectBouncedRecipients(request: CollectBouncesRequest): Promise<string[]> {
    const roles = request.includeDeletedItems ? [FolderId.Inbox, FolderId.DeletedItems] : [FolderId.Inbox];
    const output = await runPowerShellJson(`${accountScript(request.account)}
${DELIVERY_STORE_PS}
${bounceClassifierScript()}
${cutoffScript(request.daysBack)}
$found = [ordered]@{}
foreach ($role in @(${roles.map(psInt).join(', ')})) {
    $scanFolder = $null
    try { $scanFolder = $store.GetDefaultFolder($role) } catch {}
    if ($scanFolder -eq $null) { continue }
    ${itemsSinceScript('scanFolder', 'ReceivedTime', 'filtered').trim()}
    $count = $filtered.Count
    for ($i = 1; $i -le $count; $i++) {
        $item = $null
        try { $item = $filtered.Item($i) } catch { continue }
        if ($item -eq $null) { continue }
        $received = $null
        try { $received = $item.ReceivedTime } catch {}
        if ($received -ne $null -and $received -lt $cutoff) { continue }
        $report = Get-BounceReport $item
        if ($report -eq $null) { continue }
        foreach ($address in $report.failedRecipients) { $found[$address] = $true }
    }
}
ConvertTo-Json -Compress -InputObject @($found.Keys)
`, 'scan');
    return strList(output);
}

/**
 * Sent Items within the window, each message with the full SMTP address set it
 * went to (To + CC + BCC), newest first.
 */
export async function readSentRecipientGroups(request: SentGroupsRequest): Promise<SentRecipientGroup[]> {
    const output = await runPowerShellJson(`${accountScript(request.account)}
${DELIVERY_STORE_PS}
$sentFolder = $store.GetDefaultFolder(${psInt(FolderId.SentMail)})
${cutoffScript(request.daysBack)}
${itemsSinceScript('sentFolder', 'SentOn', 'filtered')}
$filtered.Sort('[SentOn]', $true)
$limit = ${psInt(request.limit)}
$rows = @()
$count = $filtered.Count
for ($i = 1; $i -le $count -and $rows.Count -lt $limit; $i++) {
    $m = $null
    try { $m = $filtered.Item($i) } catch { continue }
    if ($m -eq $null) { continue }
    $sentOn = $null
    try { $sentOn = $m.SentOn } catch {}
    if ($sentOn -ne $null -and $sentOn -lt $cutoff) { break }
    $class = 0
    try { $class = [int]$m.Class } catch {}
    if ($class -ne 43) { continue }
    $addresses = @()
    try {
        foreach ($r in $m.Recipients) {
            $address = ''
            try { $address = [string]$r.Address } catch {}
            # An internal recipient resolves to an Exchange DN; recover its SMTP address.
            if ($address -notlike '*@*') {
                try { $address = [string]$r.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch {}
            }
            if ($address -like '*@*') { $addresses += $address.ToLower() }
        }
    } catch {}
    if ($addresses.Count -eq 0) { continue }
    $rows += [PSCustomObject]@{
        entryId    = [string]$m.EntryID
        subject    = ([string]$m.Subject).Trim()
        sentOn     = if ($sentOn -ne $null) { $sentOn.ToString('yyyy-MM-dd HH:mm') } else { '' }
        recipients = @($addresses)
    }
}
ConvertTo-Json -Depth 4 -InputObject @($rows)
`, 'scan');
    return toArray(output).map(row => {
        const e = record(row);
        return {
            entryId: str(e.entryId),
            subject: str(e.subject),
            sentOn: str(e.sentOn),
            recipients: strList(e.recipients),
        };
    });
}
