// Reading mail: one folder, one item, a whole Inbox tree, or whatever is
// selected in the running Outlook.
import { psBool, psEscape, requireWindows, runPowerShell } from './run';
import { accountScript, getItemScript, mailScopeScript, namedStoreScript, SENDER_SMTP_PS } from './scripts';
import { parseArray, parseObject, record, str, toArray } from '../shared/json';
import { isOutgoingRoot, mailFolderRef, splitQuotedOriginal, WELL_KNOWN_FOLDERS } from '../mail';
import { NotFoundError } from '../errors';
import type { EmailBodyResult, InboxEmail, InboxSearchFilter, InboxSearchMatch, SelectedEmail, } from '../types';

/**
 * An arbitrary email body crosses the PowerShell boundary base64'd: it carries
 * quotes, control characters and non-ASCII that would otherwise have to survive
 * both the console codepage and JSON string escaping intact.
 */
function decodeBody(value: unknown): string {
    return Buffer.from(str(value), 'base64').toString('utf8');
}

// Outgoing roots keep their timestamp under a different COM property; anything
// else carries ReceivedTime.
const OUTGOING_DATE_PROPS: Record<number, string> = {
    [WELL_KNOWN_FOLDERS['sent items']]: 'SentOn',
    [WELL_KNOWN_FOLDERS.outbox]: 'SentOn',
    [WELL_KNOWN_FOLDERS.drafts]: 'LastModificationTime',
};

/**
 * Read recent emails from the Outlook inbox for the given account.
 * No subject filtering — returns everything within the date window, up to the limit.
 * Does NOT download attachments; returns attachment names only.
 *
 * `folder` scopes the read to one folder under the Inbox instead of the Inbox
 * root. Without it, mail already filed away (Inbox\\Invoices) is unreachable —
 * which is the normal state of any mailbox its owner keeps tidy. Segments are
 * matched as direct children; a bare name that isn't a direct child falls back to
 * the same recursive by-name search moveOutlookEmails uses, so one folder string
 * works in both. Only the named folder is read — its own subfolders are not.
 */
export async function readInboxEmails(
    emailAccount: string,
    daysBack: number = 60,
    limit: number = 50,
    folder?: string,
): Promise<InboxEmail[]> {
    if (process.platform !== 'win32') return [];
    const ref = folder ? mailFolderRef(folder) : { rootId: 6, rootLabel: 'Inbox', segments: [] };
    // A folder argument that trims away to nothing ("\\", "  ") would otherwise
    // read the Inbox root and look like it had scoped — the exact silent
    // mis-scoping this parameter exists to prevent. A bare well-known name
    // ("Sent Items") legitimately has no segments, so it is not that case.
    if (folder && folder.trim() && ref.segments.length === 0 && ref.rootId === 6
        && !Object.prototype.hasOwnProperty.call(WELL_KNOWN_FOLDERS, folder.trim().toLowerCase())) {
        throw new NotFoundError('folder', `Folder '${folder}' does not name a folder under the Inbox.`);
    }
    // Which timestamp the folder's items actually carry (see the note in the script).
    const dateProp = OUTGOING_DATE_PROPS[ref.rootId] ?? 'ReceivedTime';
    const isOutgoing = psBool(isOutgoingRoot(ref.rootId));
    const script = `${accountScript(emailAccount)}
${namedStoreScript(emailAccount)}
${mailScopeScript(ref, folder || '')}
$scopePath = $scope.FolderPath
$cutoff = (Get-Date).AddDays(-${daysBack}).ToString('MM/dd/yyyy HH:mm')
# Sent Items and Drafts carry no ReceivedTime, so restricting on it there returns
# an empty set that looks exactly like an empty folder. Filter each on the date
# property it actually has.
$filtered = $scope.Items.Restrict("[${dateProp}] >= '$cutoff'")
$filtered.Sort('[${dateProp}]', $true)
$results = @()
$cap = [Math]::Min($filtered.Count, ${limit})
for ($i = 1; $i -le $cap; $i++) {
    $item = $filtered.Item($i)
    $subj = if ($item.Subject) { $item.Subject.Trim() } else { '' }
    $bodyText = if ($item.Body) { $item.Body } else { '' }
    $bodyPreview = if ($bodyText.Length -gt 600) { $bodyText.Substring(0, 600) } else { $bodyText }
    $attNames = @()
    foreach ($att in $item.Attachments) { $attNames += $att.FileName }
    $stamp = ''
    try { $stamp = $item.${dateProp}.ToString('yyyy-MM-dd HH:mm') } catch {}
    # Outgoing mail has no meaningful sender line of its own — report who it is TO,
    # or a Sent Items listing reads as a folder of mail from yourself.
    $who = ''
    try { $who = if ($item.To) { $item.To } else { '' } } catch {}
    $results += [PSCustomObject]@{
        entryId       = $item.EntryID
        storeId       = $storeId
        folderPath    = $scopePath
        subject       = $subj
        senderName    = if (${isOutgoing}) { $who } elseif ($item.SenderName) { $item.SenderName } else { '' }
        senderEmail   = if (${isOutgoing}) { $who } elseif ($item.SenderEmailAddress) { $item.SenderEmailAddress } else { '' }
        receivedTime  = $stamp
        bodyPreview   = $bodyPreview
        attachmentNames = $attNames
        attachmentCount = $item.Attachments.Count
    }
}
ConvertTo-Json $results -Depth 3
`;
    return parseArray(await runPowerShell(script, 30000)).map(item => {
        const e = record(item);
        return {
            entryId: str(e.entryId),
            storeId: str(e.storeId),
            folderPath: str(e.folderPath),
            subject: str(e.subject),
            senderName: str(e.senderName),
            senderEmail: str(e.senderEmail),
            receivedTime: str(e.receivedTime),
            bodyPreview: str(e.bodyPreview),
            attachmentNames: toArray(e.attachmentNames).map(str),
            attachmentCount: typeof e.attachmentCount === 'number' ? e.attachmentCount : 0,
        };
    });
}

/**
 * Read one email's full plain-text body by EntryID.
 *
 * Uses `MailItem.Body`, not `HTMLBody`: Outlook's own plain-text rendering is
 * what every other reader here consumes, and the markup costs an order of
 * magnitude more for the same sentences. The tradeoff is that HTML tables
 * flatten, so figures broken out in a table lose their row/column pairing — read
 * those from the attachment where there is one.
 *
 * Exists because `readInboxEmails` caps its preview at 600 chars, and a reply
 * whose first 600 chars read as complete can still carry a material detail below
 * the cut. Nothing detects that from the preview alone.
 *
 * `includeQuoted` is off by default — see splitQuotedOriginal for why the
 * quoted thread is a hazard to a caller reading figures out of a reply.
 * `quotedLength` is reported either way so a caller can tell it exists and ask again.
 */
export async function readEmailBody(
    entryId: string,
    storeId?: string,
    maxChars: number = 8000,
    includeQuoted: boolean = false,
): Promise<EmailBodyResult> {
    requireWindows();
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
${getItemScript(entryId, storeId)}
if ($item -eq $null) { throw "Email not found for EntryID '${psEscape(entryId)}'" }
$cls = 0
try { $cls = [int]$item.Class } catch {}
if ($cls -ne 43 -and $cls -ne 46) { throw "The item for this EntryID is not an email (Class=$cls). Re-run readInboxEmails for a current EntryID." }
$bodyRaw = ''
try { if ($item.Body) { $bodyRaw = [string]$item.Body } } catch {}
$attNames = @()
try { foreach ($att in $item.Attachments) { $attNames += $att.FileName } } catch {}
${SENDER_SMTP_PS}
$received = ''
try { $received = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } catch {}
$out = [PSCustomObject]@{
    entryId         = $item.EntryID
    subject         = if ($item.Subject) { $item.Subject.Trim() } else { '' }
    senderName      = if ($item.SenderName) { $item.SenderName } else { '' }
    senderEmail     = $senderSmtp
    receivedTime    = $received
    body            = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($bodyRaw))
    attachmentNames = $attNames
}
ConvertTo-Json $out -Depth 3 -Compress
`;
    const raw = await runPowerShell(script, 20000);
    if (!raw || !raw.trim()) throw new NotFoundError('email', 'Failed to read email body.');
    const parsed = parseObject(raw);
    const { body, quoted, separator } = splitQuotedOriginal(decodeBody(parsed.body));
    const attachmentNames = toArray(parsed.attachmentNames).map(str);
    // The quoted thread is context, never the sender's own answer, so it is capped
    // harder than the reply itself — its useful part (what was asked) is at the
    // top of it.
    const quotedCap = Math.min(maxChars, 4000);
    return {
        entryId: str(parsed.entryId),
        subject: str(parsed.subject),
        senderName: str(parsed.senderName),
        senderEmail: str(parsed.senderEmail),
        receivedTime: str(parsed.receivedTime),
        body: body.length > maxChars ? body.slice(0, maxChars) : body,
        truncated: body.length > maxChars,
        bodyLength: body.length,
        quoteSeparator: separator,
        quotedLength: quoted.length,
        quotedOriginal: includeQuoted ? quoted.slice(0, quotedCap) : '',
        attachmentNames,
        attachmentCount: attachmentNames.length,
    };
}

/**
 * Walk every folder under the Inbox (recursively) for one account and return
 * the emails matching `filter` — full body included, attachments listed by
 * name but not saved.
 *
 * Built for a scan that doesn't know in advance which subfolder holds what
 * it's after; `readInboxEmails` covers the cheaper "one known folder" case.
 *
 * `daysBack` bounds the scan to items received within that many days — what a
 * daily batch run wants, since a subject filter alone walks the whole Inbox
 * tree and returns every match ever received. 0 keeps that unbounded
 * behavior, and is only sane paired with `filter.subjectLike` so Restrict can
 * narrow the set server-side before anything crosses COM.
 */
export async function searchInboxByFilter(
    emailAccount: string,
    filter: InboxSearchFilter = {},
    daysBack = 0,
): Promise<InboxSearchMatch[]> {
    if (process.platform !== 'win32') return [];
    const subjectLike = filter.subjectLike ? psEscape(filter.subjectLike) : '';
    const subjectPatternSrc = filter.subjectPattern ? psEscape(filter.subjectPattern.source) : '';
    // The subject filter reaches PowerShell as a single-quoted literal assigned to
    // $subjLike, and the Restrict query is CONCATENATED from it. Interpolating it
    // into the double-quoted query string directly — which is what the query needs
    // to be, so `$cutoff` expands — would let a caller's `$(...)` execute, and
    // would break the query outright on an apostrophe.
    const boundedRestrict = subjectLike
        ? `$folder.Items.Restrict("[ReceivedTime] >= '$cutoff' AND [Subject] like '" + $subjLike + "'")`
        : `$folder.Items.Restrict("[ReceivedTime] >= '$cutoff'")`;
    const unboundedRestrict = subjectLike
        ? `$folder.Items.Restrict("[Subject] like '" + $subjLike + "'")`
        : `$folder.Items`;
    const script = `
$subjLike = '${subjectLike}'
${accountScript(emailAccount)}
${namedStoreScript(emailAccount)}
$inbox = $store.GetDefaultFolder(6)
$results = @()
$seen = @{}
$folders = [System.Collections.ArrayList]@($inbox)
$fi = 0
while ($fi -lt $folders.Count) {
    try { foreach ($sub in $folders[$fi].Folders) { [void]$folders.Add($sub) } } catch {}
    $fi++
}
$daysBack = ${daysBack}
foreach ($folder in $folders) {
    if ($daysBack -gt 0) {
        # Bounded scan. Filter on BOTH date and subject in the Restrict so Outlook does the
        # work server-side — a date-only restrict hands back every item in the window for
        # every folder in the tree, which is thousands of COM round-trips on a busy mailbox.
        # Any stricter subject regex still runs on whatever survives.
        $cutoff = (Get-Date).AddDays(-$daysBack).ToString('MM/dd/yyyy HH:mm')
        try {
            $filtered = ${boundedRestrict}
            $fCount = $filtered.Count
        } catch {
            # Some stores reject the compound query; fall back to date-only.
            $filtered = $folder.Items.Restrict("[ReceivedTime] >= '$cutoff'")
            $fCount = $filtered.Count
        }
    } else {
        $filtered = ${unboundedRestrict}
        $fCount = $filtered.Count
        ${subjectLike ? `if ($fCount -eq 0) {
            $cutoff = (Get-Date).AddDays(-60).ToString('MM/dd/yyyy HH:mm')
            $filtered = $folder.Items.Restrict("[ReceivedTime] >= '$cutoff'")
            $fCount = $filtered.Count
        }` : ''}
    }
    for ($i = 1; $i -le $fCount; $i++) {
        $item = $filtered.Item($i)
        if ($seen.ContainsKey($item.EntryID)) { continue }
        $seen[$item.EntryID] = $true
        $subject = $item.Subject
        if (-not $subject) { continue }
        $subject = $subject.Trim()
        ${subjectPatternSrc ? `if ($subject -notmatch '${subjectPatternSrc}') { continue }` : ''}
        ${filter.excludeReplies ? `if ($subject -imatch '^(re|fw[d]?)\\s*:') { continue }` : ''}
        ${filter.requireAttachment ? `if ($item.Attachments.Count -eq 0) { continue }` : ''}
        $attNames = @()
        foreach ($att in $item.Attachments) { $attNames += $att.FileName }
        $bodyRaw = if ($item.Body) { $item.Body } else { '' }
        $bodyB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($bodyRaw))
${SENDER_SMTP_PS}
        $received = ''
        try { $received = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } catch {}
        $results += [PSCustomObject]@{
            entryId = $item.EntryID
            storeId = $storeId
            subject = $item.Subject.Trim()
            senderName = if ($item.SenderName) { $item.SenderName } else { '' }
            senderEmail = $senderSmtp
            receivedTime = $received
            body = $bodyB64
            attachmentNames = $attNames
            folderPath = $folder.FolderPath
        }
    }
}
ConvertTo-Json $results -Depth 3
`;
    // Returns whole message bodies, so this is the call most likely to push against
    // the stdout cap — raise `maxBufferBytes` via configure() before a scan that
    // matches thousands of emails.
    return parseArray(await runPowerShell(script, 300000)).map(item => {
        const e = record(item);
        return {
            entryId: str(e.entryId),
            storeId: str(e.storeId),
            subject: str(e.subject),
            senderName: str(e.senderName),
            senderEmail: str(e.senderEmail),
            receivedTime: str(e.receivedTime),
            body: decodeBody(e.body),
            attachmentNames: toArray(e.attachmentNames).map(str),
            folderPath: str(e.folderPath),
        };
    });
}

/**
 * Read the email currently selected (or open) in Outlook — full body,
 * attachments listed by name but not saved.
 */
export async function readSelectedEmail(): Promise<SelectedEmail> {
    requireWindows();
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$item = $null
try {
    $explorer = $outlook.ActiveExplorer()
    if ($explorer -ne $null) {
        $sel = $explorer.Selection
        if ($sel -ne $null -and $sel.Count -ge 1) { $item = $sel.Item(1) }
    }
} catch {}
if ($item -eq $null) {
    try {
        $insp = $outlook.ActiveInspector()
        if ($insp -ne $null) { $item = $insp.CurrentItem }
    } catch {}
}
if ($item -eq $null) { throw 'No email is selected in Outlook. Open Outlook, select (or open) an email, then try again.' }
if ($item.MessageClass -notlike 'IPM.Note*') { throw 'The selected Outlook item is not an email.' }
$attNames = @()
foreach ($att in $item.Attachments) { $attNames += $att.FileName }
$subject = ''
if ($item.Subject) { $subject = $item.Subject.Trim() }
$bodyRaw = if ($item.Body) { $item.Body } else { '' }
$bodyB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($bodyRaw))
$storeId = ''
try { $storeId = $item.Parent.Store.StoreID } catch {}
${SENDER_SMTP_PS}
$received = ''
try { $received = $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } catch {}
$result = [PSCustomObject]@{
    entryId = $item.EntryID
    storeId = $storeId
    subject = $subject
    senderName = if ($item.SenderName) { $item.SenderName } else { '' }
    senderEmail = $senderSmtp
    receivedTime = $received
    body = $bodyB64
    attachmentNames = $attNames
}
ConvertTo-Json $result -Depth 3
`;
    const raw = await runPowerShell(script, 30000);
    if (!raw || raw.trim() === '' || raw.trim() === 'null') {
        throw new NotFoundError('email', 'Failed to read the selected Outlook email.');
    }
    const e = parseObject(raw);
    return {
        entryId: str(e.entryId),
        storeId: str(e.storeId),
        subject: str(e.subject),
        senderName: str(e.senderName),
        senderEmail: str(e.senderEmail),
        receivedTime: str(e.receivedTime),
        body: decodeBody(e.body),
        attachmentNames: toArray(e.attachmentNames).map(str),
    };
}

/**
 * Open an email in Outlook by its EntryID.
 *
 * `storeId` disambiguates across mailboxes — without one, `GetItemFromID` only
 * looks in the default store, so an id from a shared or secondary mailbox
 * simply isn't found. It's tried first and the bare lookup is the fallback, so
 * a caller holding a stale StoreID (or none) still resolves whatever it can.
 *
 * EntryIDs are rewritten when an item MOVES between folders, so an id recorded
 * before its mail was filed can resolve to nothing — or, worse, to a different
 * message. The caller is the one that knows what the email was supposed to be;
 * this throws a plain not-found rather than guessing.
 */
export async function openOutlookEmail(entryId: string, storeId?: string): Promise<void> {
    requireWindows();
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$item = $null
${storeId ? `try { ${getItemScript(entryId, storeId)} } catch { $item = $null }` : ''}
if ($item -eq $null) { try { ${getItemScript(entryId)} } catch { $item = $null } }
if ($item -eq $null) { throw "Email not found for EntryID '${psEscape(entryId)}'" }
$item.Display()
`;
    await runPowerShell(script);
}
