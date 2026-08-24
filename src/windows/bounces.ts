// Bounce-backs: finding them, filing them away, and mining them for the
// addresses that failed.
import { psBool, psList, requireWindows, runPowerShell } from './run';
import { accountScript, DELIVERY_STORE_PS, namedStoreScript } from './scripts';
import { num, parseArray, parseObject, record, str, toArray } from '../shared/json';
import {
    BODY_ADDRESS_PATTERN,
    BOUNCE_DAEMON_ADDRESSES,
    BOUNCE_DAEMON_NAMES,
    BOUNCE_SUBJECT_PHRASES,
    NDR_MESSAGE_CLASS_PREFIX,
} from '../shared/bounceRules';
import { clamp } from '../mail';
import type { CleanUndeliverableResult, SentRecipientGroup } from '../types';

/**
 * The classifier lists, generated from the shared rules so Windows and macOS
 * cannot come to disagree about what a bounce is. Emitted once per script,
 * before the item loop.
 */
const BOUNCE_LISTS_PS = `
$phrases = @(${psList(BOUNCE_SUBJECT_PHRASES)})
$daemonAddr = @(${psList(BOUNCE_DAEMON_ADDRESSES)})
$daemonName = @(${psList(BOUNCE_DAEMON_NAMES)})
`;

/**
 * Per-item classifier. Given an Outlook item in `$item` and the account address
 * (lowercased) in `$targetLc`, sets `$reason` to why it's a bounce ('' if not)
 * and fills `$failedRcpts` with the recipient addresses parsed from the bounce
 * body. Also leaves `$subj`/`$sName`/`$sEmail` populated for the caller's report.
 *
 * The body is only read once a cheaper signal (class/sender/subject) has already
 * matched — that ordering is what keeps a large-inbox scan affordable.
 */
const BOUNCE_CLASSIFY_PS = `
$reason = ''
$failedRcpts = @()
$subj = ''
$sName = ''
$sEmail = ''
$cls = 0
try { $cls = [int]$item.Class } catch {}
if ($cls -eq 43 -or $cls -eq 46) {
    try { if ($item.Subject) { $subj = [string]$item.Subject } } catch {}
    $subjLc = $subj.ToLower()
    try { if ($item.SenderName) { $sName = [string]$item.SenderName } } catch {}
    $sNameLc = $sName.ToLower()
    try { if ($item.SenderEmailAddress) { $sEmail = [string]$item.SenderEmailAddress } } catch {}
    $sEmailLc = $sEmail.ToLower()
    $msgClass = ''
    try { $msgClass = [string]$item.MessageClass } catch {}
    if ($msgClass -like '${NDR_MESSAGE_CLASS_PREFIX}*') { $reason = 'Non-delivery report (NDR)' }
    if ($reason -eq '') { foreach ($d in $daemonAddr) { if ($sEmailLc.Contains($d)) { $reason = 'From mail-delivery system'; break } } }
    if ($reason -eq '') { foreach ($d in $daemonName) { if ($sNameLc.Contains($d)) { $reason = 'From mail-delivery system'; break } } }
    if ($reason -eq '') { foreach ($p in $phrases) { if ($subjLc.Contains($p)) { $reason = "Bounce subject phrase: '$p'"; break } } }
    if ($reason -ne '') {
        $bodyRaw = ''
        try { if ($item.Body) { $bodyRaw = [string]$item.Body } } catch {}
        if ($bodyRaw -ne '') {
            $mm = [regex]::Matches($bodyRaw, '${BODY_ADDRESS_PATTERN}')
            foreach ($m in $mm) {
                $addr = $m.Value.ToLower().TrimEnd('.')
                if ($addr -eq $targetLc) { continue }
                if ($addr.Contains('mailer-daemon') -or $addr.Contains('postmaster') -or $addr.Contains('mail-daemon')) { continue }
                if ($failedRcpts -notcontains $addr) { $failedRcpts += $addr }
                if ($failedRcpts.Count -ge 5) { break }
            }
        }
    }
}
`;

/**
 * Scan an account's inbox for bounce-back / non-delivery messages — Outlook NDRs,
 * mail-daemon/postmaster rejections, and "Message blocked"-style Google/O365
 * failure notices — and, unless previewing, move each to Deleted Items (recoverable).
 *
 * Classification is deliberately conservative so ordinary mail that merely
 * mentions "delivery" is never caught: an item matches only when its
 * MessageClass is an NDR report, its sender fingerprints as a mail-delivery
 * daemon/postmaster, or its subject contains a specific bounce phrase.
 *
 * Matched item references are snapshotted before any Delete() call — deleting
 * mutates the folder collection, so deleting mid-enumeration would skip items
 * (same pattern as sendAllDrafts).
 */
export async function cleanUndeliverableEmails(
    emailAccount: string,
    daysBack = 30,
    dryRun = true,
): Promise<CleanUndeliverableResult> {
    requireWindows();
    const days = clamp(daysBack, 1, 365);
    const script = `${accountScript(emailAccount)}
$dryRun = ${psBool(dryRun)}
${namedStoreScript(emailAccount)}
$inbox = $store.GetDefaultFolder(6)
$cutoff = (Get-Date).AddDays(-${days}).ToString('MM/dd/yyyy HH:mm')
$filtered = $inbox.Items.Restrict("[ReceivedTime] >= '$cutoff'")
${BOUNCE_LISTS_PS}
$targetLc = $target.ToLower()
$matched = @()   # snapshot of COM item refs to (optionally) delete afterward
$report = @()
$count = $filtered.Count
for ($i = 1; $i -le $count; $i++) {
    $item = $null
    try { $item = $filtered.Item($i) } catch { continue }
    if ($item -eq $null) { continue }
${BOUNCE_CLASSIFY_PS}
    if ($reason -eq '') { continue }

    $rt = ''
    try { $rt = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } catch {}
    $report += [PSCustomObject]@{
        entryId = $item.EntryID
        subject = $subj.Trim()
        senderName = $sName
        senderEmail = $sEmail
        receivedTime = $rt
        matchedReason = $reason
        failedRecipients = @($failedRcpts)
    }
    $matched += $item
}

$deleted = 0
$failed = @()
if (-not $dryRun) {
    foreach ($m in $matched) {
        $s = ''
        try { $s = [string]$m.Subject } catch {}
        try { $m.Delete(); $deleted++ }
        catch { $failed += [PSCustomObject]@{ subject = $s; error = $_.Exception.Message } }
    }
}

[PSCustomObject]@{
    account = $target
    scannedDays = ${days}
    dryRun = $dryRun
    matchedCount = $report.Count
    deletedCount = $deleted
    matched = @($report)
    failed = @($failed)
} | ConvertTo-Json -Depth 4
`;
    const parsed = parseObject(await runPowerShell(script, 300000));
    return {
        account: str(parsed.account) || emailAccount,
        scannedDays: num(parsed.scannedDays) || days,
        dryRun: parsed.dryRun !== false,
        matchedCount: num(parsed.matchedCount),
        deletedCount: num(parsed.deletedCount),
        matched: toArray(parsed.matched).map(m => {
            const e = record(m);
            return {
                entryId: str(e.entryId),
                subject: str(e.subject),
                senderName: str(e.senderName),
                senderEmail: str(e.senderEmail),
                receivedTime: str(e.receivedTime),
                matchedReason: str(e.matchedReason),
                failedRecipients: toArray(e.failedRecipients).map(str),
            };
        }),
        failed: toArray(parsed.failed).map(f => {
            const e = record(f);
            return {subject: str(e.subject), error: str(e.error)};
        }),
    };
}

/**
 * Read-only scan of an account's Inbox (and, when scanDeleted, its Deleted Items)
 * for bounce messages, returning just the deduped set of failed recipient
 * addresses. Uses the same conservative classifier as cleanUndeliverableEmails.
 *
 * Deleted Items is included so this still works after cleanUndeliverableEmails has
 * already filed the bounces there — the blacklist step doesn't depend on running
 * before the cleanup. This never deletes anything.
 */
export async function collectBouncedRecipients(
    emailAccount: string,
    daysBack = 30,
    scanDeleted = true,
): Promise<string[]> {
    if (process.platform !== 'win32') return [];
    const days = clamp(daysBack, 1, 365);
    // 6 = Inbox, 3 = Deleted Items.
    const folderIds = scanDeleted ? '@(6, 3)' : '@(6)';
    const script = `${accountScript(emailAccount)}
${DELIVERY_STORE_PS}
$cutoff = (Get-Date).AddDays(-${days}).ToString('MM/dd/yyyy HH:mm')
${BOUNCE_LISTS_PS}
$targetLc = $target.ToLower()
$found = @{}
foreach ($fid in ${folderIds}) {
    $folder = $null
    try { $folder = $store.GetDefaultFolder($fid) } catch {}
    if ($folder -eq $null) { continue }
    $filtered = $folder.Items.Restrict("[ReceivedTime] >= '$cutoff'")
    $count = $filtered.Count
    for ($i = 1; $i -le $count; $i++) {
        $item = $null
        try { $item = $filtered.Item($i) } catch { continue }
        if ($item -eq $null) { continue }
${BOUNCE_CLASSIFY_PS}
        if ($reason -eq '') { continue }
        foreach ($a in $failedRcpts) { $found[$a] = $true }
    }
}
ConvertTo-Json @($found.Keys) -Depth 2
`;
    return parseArray(await runPowerShell(script, 300000)).map(str).filter(Boolean);
}

/**
 * Read the account's Sent Items within the window, returning each mail with the
 * full SMTP address set it was sent to (To + CC + BCC), newest first. This is how
 * the "was every address for this contact tried?" question gets answered: where a
 * blast sends one email per organization addressed to all of its addresses, a sent
 * message's recipient set IS that organization's full address set.
 */
export async function readSentRecipientGroups(
    emailAccount: string,
    daysBack = 30,
    limit = 3000,
): Promise<SentRecipientGroup[]> {
    if (process.platform !== 'win32') return [];
    const days = clamp(daysBack, 1, 365);
    const cap = clamp(limit, 1, 10000);
    const script = `${accountScript(emailAccount)}
$sent = $account.DeliveryStore.GetDefaultFolder(5)  # olFolderSentMail
$cutoff = (Get-Date).AddDays(-${days}).ToString('MM/dd/yyyy HH:mm')
$items = $sent.Items.Restrict("[SentOn] >= '$cutoff'")
$items.Sort('[SentOn]', $true)
$results = @()
$count = [Math]::Min($items.Count, ${cap})
for ($i = 1; $i -le $count; $i++) {
    $m = $null
    try { $m = $items.Item($i) } catch { continue }
    if ($m -eq $null) { continue }
    $cls = 0
    try { $cls = [int]$m.Class } catch {}
    if ($cls -ne 43) { continue }  # olMail only
    $addrs = @()
    try {
        foreach ($r in $m.Recipients) {
            $addr = ''
            try { $addr = [string]$r.Address } catch {}
            # Internal recipients resolve to an Exchange DN, not SMTP — recover the
            # real address so external correlation still works.
            if ($addr -notlike '*@*') {
                try { $addr = [string]$r.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch {}
            }
            if ($addr -like '*@*') { $addrs += $addr.ToLower() }
        }
    } catch {}
    if ($addrs.Count -eq 0) { continue }
    $sentOn = ''
    try { $sentOn = $m.SentOn.ToString('yyyy-MM-dd HH:mm') } catch {}
    $subj = ''
    try { if ($m.Subject) { $subj = [string]$m.Subject } } catch {}
    $results += [PSCustomObject]@{
        entryId = $m.EntryID
        subject = $subj.Trim()
        sentOn = $sentOn
        recipients = @($addrs)
    }
}
ConvertTo-Json @($results) -Depth 4
`;
    return parseArray(await runPowerShell(script, 300000)).map(it => {
        const e = record(it);
        return {
            entryId: str(e.entryId),
            subject: str(e.subject),
            sentOn: str(e.sentOn),
            recipients: toArray(e.recipients).map(str),
        };
    });
}
