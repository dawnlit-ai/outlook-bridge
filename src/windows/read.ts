// Reading mail: one folder, one item, a whole Inbox tree, or whatever is
// selected in the running Outlook.
import { psArray, psBool, psInt, psString, runPowerShell, runPowerShellJson } from './run';
import {
    accountScript,
    cutoffScript,
    itemLookupScript,
    itemsSinceScript,
    mailScopeScript,
    NAMED_STORE_PS,
    SENDER_SMTP_PS,
    SESSION_PS,
} from './scripts';
import { base64Text, record, str, strList, toArray } from '../shared/json';
import { FolderId, isOutgoingRoot, NON_INCOMING_ROOTS, REPLY_PREFIX_SOURCE, subjectGlobSource } from '../mail';
import { failureTag } from '../errors';
import type { RawEmail, ReadInboxRequest, SearchRequest } from '../backend';
import type { EmailLocator, InboxEmail, InboxSearchMatch, SelectedEmail } from '../types';

/**
 * The COM property carrying each kind of folder's timestamp. Sent Items and
 * Drafts have no ReceivedTime, and filtering them on it returns an empty set
 * that looks exactly like an empty folder.
 */
function dateProperty(rootId: number): string {
    if (rootId === FolderId.Drafts) return 'LastModificationTime';
    return isOutgoingRoot(rootId) ? 'SentOn' : 'ReceivedTime';
}

/**
 * Bodies cross the PowerShell boundary base64-encoded: they carry quotes,
 * control characters and non-ASCII that would otherwise have to survive JSON
 * escaping in two languages intact.
 */
const BODY_B64_PS = `[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes([string]$bodyRaw))`;

/** Recent mail from the Inbox root or one folder, newest first. */
export async function readInboxEmails(request: ReadInboxRequest): Promise<InboxEmail[]> {
    const dateProp = dateProperty(request.folder.rootId);
    const outgoing = isOutgoingRoot(request.folder.rootId);
    const output = await runPowerShellJson(`${accountScript(request.account)}
${NAMED_STORE_PS}
${mailScopeScript(request.folder, request.folderLabel)}
$scopePath = [string]$scope.FolderPath
${cutoffScript(request.daysBack)}
${itemsSinceScript('scope', dateProp, 'filtered')}
$filtered.Sort('[${dateProp}]', $true)
$limit = ${psInt(request.limit)}
$previewChars = ${psInt(request.previewChars)}
$rows = @()
$count = $filtered.Count
for ($i = 1; $i -le $count -and $rows.Count -lt $limit; $i++) {
    $item = $null
    try { $item = $filtered.Item($i) } catch {}
    if ($item -eq $null) { continue }
    $stamp = $null
    try { $stamp = $item.${dateProp} } catch {}
    # Sorted newest first, so the first item older than the window ends it.
    if ($stamp -ne $null -and $stamp -lt $cutoff) { break }
    $preview = ''
    if ($previewChars -gt 0) {
        $bodyText = ''
        try { $bodyText = [string]$item.Body } catch {}
        $preview = if ($bodyText.Length -gt $previewChars) { $bodyText.Substring(0, $previewChars) } else { $bodyText }
    }
    $attachmentNames = @()
    try { foreach ($att in $item.Attachments) { $attachmentNames += [string]$att.FileName } } catch {}
    $who = ''
    $whoEmail = ''
    if (${psBool(outgoing)}) {
        # Outgoing mail reports who it went TO; "from" is always the mailbox itself.
        try { $who = [string]$item.To } catch {}
        $whoEmail = $who
    } else {
        try { $who = [string]$item.SenderName } catch {}
        try { $whoEmail = [string]$item.SenderEmailAddress } catch {}
    }
    $rows += [PSCustomObject]@{
        entryId         = [string]$item.EntryID
        storeId         = $storeId
        folderPath      = $scopePath
        subject         = ([string]$item.Subject).Trim()
        senderName      = $who
        senderEmail     = $whoEmail
        receivedTime    = if ($stamp -ne $null) { $stamp.ToString('yyyy-MM-dd HH:mm') } else { '' }
        bodyPreview     = $preview
        attachmentNames = @($attachmentNames)
    }
}
ConvertTo-Json -Depth 4 -InputObject @($rows)
`, 'standard');
    return toArray(output).map(row => {
        const e = record(row);
        const attachmentNames = toArray(e.attachmentNames).map(str);
        return {
            entryId: str(e.entryId),
            storeId: str(e.storeId),
            subject: str(e.subject),
            senderName: str(e.senderName),
            senderEmail: str(e.senderEmail),
            receivedTime: str(e.receivedTime),
            bodyPreview: str(e.bodyPreview),
            attachmentNames,
            attachmentCount: attachmentNames.length,
            folderPath: str(e.folderPath),
        };
    });
}

/**
 * One email's full plain-text body, by id.
 *
 * `Body`, not `HTMLBody`: Outlook's own plain-text rendering is what every other
 * reader here consumes, and markup costs an order of magnitude more for the same
 * sentences. The tradeoff is that an HTML table's rows flatten.
 */
export async function readEmailBody(email: EmailLocator): Promise<RawEmail> {
    const output = await runPowerShellJson(`${SESSION_PS}
${itemLookupScript(email)}
$class = 0
try { $class = [int]$item.Class } catch {}
# 43 = olMail, 46 = olReport (a non-delivery report is still worth reading).
if ($class -ne 43 -and $class -ne 46) {
    throw "${failureTag('NOT_FOUND', 'email')}The item with entry id '$lookupId' is not an email (class $class)."
}
$bodyRaw = ''
try { $bodyRaw = [string]$item.Body } catch {}
$attachmentNames = @()
try { foreach ($att in $item.Attachments) { $attachmentNames += [string]$att.FileName } } catch {}
${SENDER_SMTP_PS}
$received = ''
try { $received = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } catch {}
ConvertTo-Json -Compress -Depth 3 -InputObject ([PSCustomObject]@{
    entryId         = [string]$item.EntryID
    subject         = ([string]$item.Subject).Trim()
    senderName      = [string]$item.SenderName
    senderEmail     = $senderSmtp
    receivedTime    = $received
    body            = ${BODY_B64_PS}
    attachmentNames = @($attachmentNames)
})
`, 'quick');
    const e = record(output);
    return {
        entryId: str(e.entryId) || email.entryId,
        subject: str(e.subject),
        senderName: str(e.senderName),
        senderEmail: str(e.senderEmail),
        receivedTime: str(e.receivedTime),
        body: base64Text(e.body),
        attachmentNames: toArray(e.attachmentNames).map(str),
    };
}

/**
 * Walk every folder under the Inbox for the mail matching the request.
 *
 * The window is enforced twice — a date Restrict per folder, then a per-item
 * check — because Restrict is not reliable across stores, and a scan that
 * silently widens on some mailboxes is the failure this guards against.
 * `subjectLike` also drives a server-side prefilter, purely as an optimization:
 * every subject is re-checked, so a store that can't run the query (IMAP stores
 * accept it and match nothing) returns the same set, only slower.
 */
export async function searchInboxByFilter(request: SearchRequest): Promise<InboxSearchMatch[]> {
    const subjectFiltered = !!(request.subjectLike || request.subjectPattern || request.excludeReplies);
    // The server-side prefilter speaks DASL, whose `like` wildcards are % and _.
    // A literal % or _ in the glob is left as a wildcard there: that can only
    // widen the prefilter, and the exact test below narrows it again.
    const subjectDasl = request.subjectLike ? request.subjectLike.replace(/\*/g, '%').replace(/\?/g, '_') : '';
    const output = await runPowerShellJson(`${accountScript(request.account)}
${NAMED_STORE_PS}
$inbox = $store.GetDefaultFolder(${psInt(FolderId.Inbox)})
$subjectLike = ${psString(request.subjectLike ? subjectGlobSource(request.subjectLike) : '')}
$subjectDasl = ${psString(subjectDasl)}
$subjectPattern = ${psString(request.subjectPattern?.source ?? '')}
$replyPattern = ${psString(request.excludeReplies ? REPLY_PREFIX_SOURCE : '')}
$requireSubject = ${psBool(subjectFiltered)}
$requireAttachment = ${psBool(request.requireAttachment)}
$includeBody = ${psBool(request.includeBody)}
${cutoffScript(request.daysBack)}
# Sent, Drafts, Deleted and Junk are siblings of the Inbox on an Exchange profile
# and CHILDREN of it on an IMAP one. They are cut by id — their names are
# localized — before the walk, so each takes its whole subtree with it.
$skipIds = @{}
foreach ($role in @(${NON_INCOMING_ROOTS.map(psInt).join(', ')})) {
    try {
        $roleFolder = $store.GetDefaultFolder($role)
        if ($roleFolder -ne $null) { $skipIds[[string]$roleFolder.EntryID] = $true }
    } catch {}
}
$skipNames = ${psArray(request.excludeFolders)}
$onlyNames = ${psArray(request.includeFolders)}
function Test-FolderNamed([string]$name, [string]$path, $list) {
    foreach ($entry in $list) {
        if ($name -ieq $entry -or $path -ieq $entry) { return $true }
    }
    return $false
}
# Traversal and scanning are separate, and the two lists are not symmetric: an
# excluded folder is never queued, so its subtree goes with it, while an
# included one only marks a folder the walk reached anyway — the walk still
# passes through folders nobody named to reach a nested one somebody did.
$rootName = ''
try { $rootName = [string]$inbox.Name } catch {}
$rootPath = ''
try { $rootPath = [string]$inbox.FolderPath } catch {}
$folders = [System.Collections.ArrayList]@($inbox)
$inScope = [System.Collections.ArrayList]@(($onlyNames.Count -eq 0) -or (Test-FolderNamed $rootName $rootPath $onlyNames))
$walked = 0
while ($walked -lt $folders.Count) {
    try {
        foreach ($sub in $folders[$walked].Folders) {
            if ($skipIds.ContainsKey([string]$sub.EntryID)) { continue }
            $subName = ''
            try { $subName = [string]$sub.Name } catch {}
            $subPath = ''
            try { $subPath = [string]$sub.FolderPath } catch {}
            if (Test-FolderNamed $subName $subPath $skipNames) { continue }
            [void]$folders.Add($sub)
            [void]$inScope.Add(($onlyNames.Count -eq 0) -or (Test-FolderNamed $subName $subPath $onlyNames))
        }
    } catch {}
    $walked++
}
# The subject prefilter is DASL rather than Jet: some stores reject Jet's
# "[Subject] like" outright, while DASL is accepted wherever Restrict is. The
# two syntaxes can't share one filter, so it is a second Restrict on top of the
# date one. The pattern is concatenated into the query, never interpolated.
$subjectQuery = '@SQL=' + [char]34 + 'urn:schemas:httpmail:subject' + [char]34 + ' like ' + [char]39 + $subjectDasl.Replace("'", "''") + [char]39
# A list, not an array: += copies the whole array per row, and a wide scan returns thousands.
$rows = New-Object System.Collections.ArrayList
$seen = @{}
for ($fx = 0; $fx -lt $folders.Count; $fx++) {
    if (-not $inScope[$fx]) { continue }
    $folder = $folders[$fx]
    ${itemsSinceScript('folder', 'ReceivedTime', 'filtered').trim()}
    if ($subjectDasl) {
        try {
            $candidate = $filtered.Restrict($subjectQuery)
            if ($candidate.Count -gt 0) { $filtered = $candidate }
        } catch {}
    }
    $folderPath = ''
    try { $folderPath = [string]$folder.FolderPath } catch {}
    $count = $filtered.Count
    for ($i = 1; $i -le $count; $i++) {
        $item = $null
        try { $item = $filtered.Item($i) } catch {}
        # A restricted collection can hand back nothing for an index, and an item
        # with no id can't be acted on by anything this package does.
        if ($item -eq $null) { continue }
        $entryId = [string]$item.EntryID
        if (-not $entryId -or $seen.ContainsKey($entryId)) { continue }
        $seen[$entryId] = $true
        $subject = ''
        try { $subject = ([string]$item.Subject).Trim() } catch {}
        if ($requireSubject -and -not $subject) { continue }
        if ($subjectLike -and $subject -notmatch $subjectLike) { continue }
        if ($subjectPattern -and $subject -notmatch $subjectPattern) { continue }
        if ($replyPattern -and $subject -imatch $replyPattern) { continue }
        $received = $null
        try { $received = $item.ReceivedTime } catch {}
        if ($received -ne $null -and $received -lt $cutoff) { continue }
        $attachmentNames = @()
        try { foreach ($att in $item.Attachments) { $attachmentNames += [string]$att.FileName } } catch {}
        if ($requireAttachment -and $attachmentNames.Count -eq 0) { continue }
        $bodyRaw = ''
        if ($includeBody) { try { $bodyRaw = [string]$item.Body } catch {} }
${SENDER_SMTP_PS}
        [void]$rows.Add([PSCustomObject]@{
            entryId         = $entryId
            storeId         = $storeId
            subject         = $subject
            senderName      = [string]$item.SenderName
            senderEmail     = $senderSmtp
            receivedTime    = if ($received -ne $null) { $received.ToString('yyyy-MM-dd HH:mm') } else { '' }
            body            = ${BODY_B64_PS}
            attachmentNames = @($attachmentNames)
            folderPath      = $folderPath
        })
    }
}
ConvertTo-Json -Depth 4 -InputObject @($rows)
`, 'scan');
    return toArray(output).map(row => {
        const e = record(row);
        return {
            entryId: str(e.entryId),
            storeId: str(e.storeId),
            subject: str(e.subject),
            senderName: str(e.senderName),
            senderEmail: str(e.senderEmail),
            receivedTime: str(e.receivedTime),
            body: base64Text(e.body),
            attachmentNames: toArray(e.attachmentNames).map(str),
            folderPath: str(e.folderPath),
        };
    });
}

/** The email currently selected — or open — in Outlook. */
export async function readSelectedEmail(): Promise<SelectedEmail> {
    const output = await runPowerShellJson(`
$outlook = New-Object -ComObject Outlook.Application
$item = $null
try {
    $explorer = $outlook.ActiveExplorer()
    if ($explorer -ne $null -and $explorer.Selection.Count -ge 1) { $item = $explorer.Selection.Item(1) }
} catch {}
if ($item -eq $null) {
    try {
        $inspector = $outlook.ActiveInspector()
        if ($inspector -ne $null) { $item = $inspector.CurrentItem }
    } catch {}
}
if ($item -eq $null) {
    throw "${failureTag('NOT_FOUND', 'email')}No email is selected in Outlook. Select or open an email, then try again."
}
if ([string]$item.MessageClass -notlike 'IPM.Note*') {
    throw "${failureTag('INVALID_REQUEST')}The item selected in Outlook is not an email."
}
$attachmentNames = @()
try { foreach ($att in $item.Attachments) { $attachmentNames += [string]$att.FileName } } catch {}
$bodyRaw = ''
try { $bodyRaw = [string]$item.Body } catch {}
$storeId = ''
try { $storeId = [string]$item.Parent.Store.StoreID } catch {}
${SENDER_SMTP_PS}
$received = ''
try { $received = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } catch {}
ConvertTo-Json -Compress -Depth 3 -InputObject ([PSCustomObject]@{
    entryId         = [string]$item.EntryID
    storeId         = $storeId
    subject         = ([string]$item.Subject).Trim()
    senderName      = [string]$item.SenderName
    senderEmail     = $senderSmtp
    receivedTime    = $received
    body            = ${BODY_B64_PS}
    attachmentNames = @($attachmentNames)
})
`, 'quick');
    const e = record(output);
    return {
        entryId: str(e.entryId),
        storeId: str(e.storeId),
        subject: str(e.subject),
        senderName: str(e.senderName),
        senderEmail: str(e.senderEmail),
        receivedTime: str(e.receivedTime),
        body: base64Text(e.body),
        attachmentNames: strList(e.attachmentNames),
    };
}

/**
 * Open an email in Outlook.
 *
 * Entry ids are rewritten when an item MOVES between folders, so an id recorded
 * before its mail was filed can stop resolving. That fails as NOT_FOUND rather
 * than guessing — the caller is the one who knows what the email was.
 */
export async function openOutlookEmail(email: EmailLocator): Promise<void> {
    await runPowerShell(`${SESSION_PS}
${itemLookupScript(email)}
$item.Display()
`, 'quick');
}
