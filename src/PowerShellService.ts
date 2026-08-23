import { execFile } from 'child_process';
import fs from 'fs';
import path from 'path';
import { pathToFileURL } from 'url';
import { composeTemplateBody, findTemplateMarkers } from './outlookTemplateSections';
import { mailFolderRef, splitQuotedOriginal, WELL_KNOWN_FOLDERS } from './mail';
import { getConfig, reportRun, resolveDestDir, tempFile } from './runtime';
import type {
    CleanUndeliverableResult,
    DeleteDraftsResult,
    DeleteMailOptions,
    DeleteMailOutcome,
    DeleteMailResult,
    EmailBodyResult,
    InboxEmail,
    InboxFolderInfo,
    InboxSearchFilter,
    InboxSearchMatch,
    ListDraftsResult,
    MailFolderRef,
    MoveEmailsResult,
    OutlookBridge,
    PurgeDeletedItemsResult,
    ReplyEmailParams,
    ReplyEmailResult,
    SavedAttachment,
    SaveTemplateResult,
    SelectedEmail,
    SendAllDraftsResult,
    SendEmailParams,
    SentRecipientGroup,
    TemplateFolderResult,
} from './types';

/**
 * Run a generated script through Windows PowerShell.
 *
 * `timeout` overrides the configured default for this one call — pass it where a
 * call has a genuinely different budget (a full-mailbox walk, a purge) rather
 * than relying on the global. See `configure()` for both knobs.
 */
function runPowerShell(script: string, timeout?: number): Promise<string> {
    if (process.platform !== 'win32') {
        return Promise.reject(new Error('PowerShell and COM automation are only supported on Windows.'));
    }
    // Force UTF-8 on stdout. Windows PowerShell 5.1 otherwise encodes output in the
    // OEM/ANSI console code page, whose "best-fit" mapping silently rewrites non-ASCII
    // characters — e.g. a curly quote (U+201C) in an email body becomes a plain " —
    // which corrupts the JSON these scripts emit (an unescaped quote) and breaks
    // JSON.parse. Setting the output encoding first makes non-ASCII survive as real
    // UTF-8 bytes, which Node then decodes correctly.
    const utf8Script = `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n${script}`;
    const {timeoutMs, maxBufferBytes} = getConfig();
    const startedAt = Date.now();
    return new Promise((resolve, reject) => {
        execFile(
            'powershell.exe',
            ['-NoProfile', '-NonInteractive', '-Command', utf8Script],
            {maxBuffer: maxBufferBytes, timeout: timeout ?? timeoutMs},
            (error, stdout, stderr) => {
                const durationMs = Date.now() - startedAt;
                if (error) {
                    const message = stderr || error.message;
                    reportRun({runner: 'powershell', script: utf8Script, durationMs, error: message});
                    reject(new Error(message));
                } else {
                    reportRun({runner: 'powershell', script: utf8Script, durationMs});
                    resolve(stdout.trim());
                }
            }
        );
    });
}

/**
 * Escape a string for embedding inside a PowerShell SINGLE-quoted literal — the
 * only context caller-supplied text may go in. A double-quoted PowerShell string
 * expands `$(...)` subexpressions, so text placed there would execute; where a
 * generated script needs user text inside one, assign it to a variable with this
 * first and concatenate (see searchInboxByFilter).
 */
function psEscape(s: string): string {
    return s.replace(/'/g, "''");
}

/** Send (or display for review) an email through Outlook COM. */
export async function sendOutlookEmail(params: SendEmailParams): Promise<void> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const attachLine = params.attachmentPath
        ? `$mail.Attachments.Add('${psEscape(params.attachmentPath)}') | Out-Null`
        : '';
    // Display() only shows the compose window — it never files the item, so a
    // silent draft has to Save() into the default Drafts folder explicitly.
    const openDraftWindow = params.openDraftWindow !== false;
    const actionLine = params.sendImmediately
        ? '$mail.Send()'
        : openDraftWindow ? '$mail.Display()' : '$mail.Save()';

    // Write HTML body to a temp file to avoid ENAMETOOLONG on large emails
    const bodyFile = tempFile('email-body', 'html');
    const scriptFile = tempFile('email-script', 'ps1');
    fs.writeFileSync(bodyFile, params.htmlBody, 'utf-8');

    const psBodyPath = psEscape(bodyFile);
    const script = `
$ErrorActionPreference = 'Stop'
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)

# Set sending account. Fail rather than fall back to the default account, so a
# bad address can't silently send from the wrong mailbox.
$account = $null
foreach ($a in $outlook.Session.Accounts) {
    if ($a.SmtpAddress -ieq '${psEscape(params.emailAccount)}') { $account = $a; break }
}
if ($account -eq $null) { throw "Account '${psEscape(params.emailAccount)}' not found" }
# Direct assignment ($mail.SendUsingAccount = $account) is a silent no-op under
# PowerShell's COM binding; set the property through IDispatch reflection instead.
[void]$mail.GetType().InvokeMember('SendUsingAccount', [Reflection.BindingFlags]::SetProperty, $null, $mail, @($account))

$mail.To = '${psEscape(params.to)}'
$mail.CC = '${psEscape(params.cc || '')}'
$mail.Subject = '${psEscape(params.subject)}'
$mail.HTMLBody = [IO.File]::ReadAllText('${psBodyPath}', [Text.Encoding]::UTF8)
${attachLine}
${actionLine}
`;
    fs.writeFileSync(scriptFile, script, 'utf-8');

    try {
        await new Promise<void>((resolve, reject) => {
            execFile(
                'powershell.exe',
                ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptFile],
                {timeout: 120000},
                (error, _stdout, stderr) => {
                    if (error) reject(new Error(stderr || error.message));
                    else resolve();
                }
            );
        });
    } finally {
        try {
            fs.unlinkSync(bodyFile);
        } catch { /* ignore */
        }
        try {
            fs.unlinkSync(scriptFile);
        } catch { /* ignore */
        }
    }
}

/**
 * Resolve a saved template's HTML body by subject, so reply flows can name a template
 * instead of carrying its (often large, Word-generated) HTML. Reuses readTemplateEmails,
 * so the same cid-image stripping and size cap apply. Throws a helpful error naming what
 * is available when the folder or the subject can't be found.
 */
async function resolveTemplateHtmlBySubject(
    emailAccount: string,
    folderName: string,
    subject: string,
): Promise<string> {
    // Filter by subject in the script rather than here: unfiltered, EVERY template's
    // full body (up to 100k chars each) crosses stdout on every reply, and ~20
    // Word-sized templates overrun runPowerShell's 1 MB maxBuffer outright — turning
    // a growing Templates folder into "stdout maxBuffer length exceeded" on a call
    // that never needed more than one body.
    const result = await readTemplateEmails(emailAccount, folderName, 50, true, subject);
    if (!result.folderFound) {
        const folders = result.availableFolders.join(', ') || '(none)';
        throw new Error(`Template folder '${folderName}' not found in ${emailAccount}. Available folders: ${folders}.`);
    }
    const want = subject.trim().toLowerCase();
    const match = result.templates.find(t => t.subject.trim().toLowerCase() === want);
    if (!match) {
        // The read above was filtered, so it can't name the alternatives. List them in a
        // second pass with bodies omitted — which is what makes reading the whole folder
        // safe here, and keeps the error as useful as it was before the filter.
        const all = await readTemplateEmails(emailAccount, folderName, 50, false);
        const names = all.templates.map(t => t.subject).filter(Boolean).join(', ') || '(none)';
        throw new Error(`Template '${subject}' not found in '${folderName}'. Available templates: ${names}.`);
    }
    if (!match.htmlBody || match.htmlBody.trim() === '') {
        throw new Error(`Template '${subject}' has an empty HTML body.`);
    }
    return match.htmlBody;
}

/**
 * A signature `.htm` is a whole HTML document; only its body belongs inside the reply.
 * Substituting the document whole would nest `<html>`/`<body>` inside the template's own.
 */
function signatureInnerHtml(html: string): string {
    const match = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
    return (match ? match[1] : html).trim();
}

export async function replyOutlookEmail(params: ReplyEmailParams): Promise<ReplyEmailResult> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    // Resolve the reply body: either the caller passed htmlBody, or it named a saved
    // template (templateSubject) that we fetch here — so a large template body is
    // resolved server-side and never round-trips through the caller.
    let htmlBody = params.htmlBody;
    let fromTemplate = false;
    if ((!htmlBody || htmlBody.trim() === '') && params.templateSubject) {
        htmlBody = await resolveTemplateHtmlBySubject(
            params.emailAccount,
            params.templateFolder || 'Templates',
            params.templateSubject,
        );
        fromTemplate = true;
    }
    if (!htmlBody || htmlBody.trim() === '') {
        throw new Error('replyOutlookEmail needs either htmlBody or a resolvable templateSubject.');
    }
    // One template can hold several reply variants between [[SECTION]] markers: keep the
    // requested one, drop the rest, fill any {{PLACEHOLDER}}. Composing also runs (with no
    // options) on any body taken straight from a template, so a stray marker is caught
    // here rather than mailed to a customer.
    // A named signature fills {{SIGNATURE}} like any other placeholder, but is read from
    // the Signatures folder here so the caller never carries the signature HTML either.
    let placeholders = params.templatePlaceholders;
    if (params.signatureName) {
        const signatureHtml = await readOutlookSignatureHtml(params.signatureName);
        if (signatureHtml.trim() === '') {
            const available = (await listOutlookSignatures()).join(', ') || '(none)';
            throw new Error(
                `Outlook signature '${params.signatureName}' not found. Available signatures: ${available}.`,
            );
        }
        placeholders = {...placeholders, SIGNATURE: signatureInnerHtml(signatureHtml)};
    }
    const hasPlaceholders = !!placeholders && Object.keys(placeholders).length > 0;
    if (fromTemplate || params.templateSection || hasPlaceholders) {
        htmlBody = composeTemplateBody(htmlBody, {
            section: params.templateSection,
            placeholders,
            label: params.templateSubject,
        });
    }
    // A template read back from Outlook is a full Word HTML document; nesting one
    // document inside the reply's confuses no renderer but bloats the item — keep
    // just the body content (inline styles carry the meaningful formatting).
    const bodyMatch = /<body[^>]*>([\s\S]*)<\/body>/i.exec(htmlBody);
    const insertHtml = bodyMatch ? bodyMatch[1] : htmlBody;
    const openDraftWindow = params.openDraftWindow !== false;
    const actionLine = params.sendImmediately
        ? '$reply.Send()'
        : openDraftWindow ? '$reply.Display()' : '$reply.Save()';
    const bodyFile = tempFile('reply-insert', 'html');
    fs.writeFileSync(bodyFile, insertHtml, 'utf-8');
    const getItemLine = params.storeId
        ? `$item = $ns.GetItemFromID('${psEscape(params.entryId)}', '${psEscape(params.storeId)}')`
        : `$item = $ns.GetItemFromID('${psEscape(params.entryId)}')`;
    const replyMethod = params.replyAll ? 'ReplyAll' : 'Reply';
    const script = `
$ErrorActionPreference = 'Stop'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
${getItemLine}
$reply = $item.${replyMethod}()

# Fail rather than fall back to the default account, like sendOutlookEmail.
$account = $null
foreach ($a in $outlook.Session.Accounts) {
    if ($a.SmtpAddress -ieq '${psEscape(params.emailAccount)}') { $account = $a; break }
}
if ($account -eq $null) { throw "Account '${psEscape(params.emailAccount)}' not found" }
# Direct assignment is a silent no-op under PowerShell's COM binding.
[void]$reply.GetType().InvokeMember('SendUsingAccount', [Reflection.BindingFlags]::SetProperty, $null, $reply, @($account))

$insertHtml = [IO.File]::ReadAllText('${psEscape(bodyFile)}', [Text.Encoding]::UTF8)
$keyWord = 'WordSection1>'
$idx = $reply.HTMLBody.IndexOf($keyWord)
if ($idx -ge 0) { $idx += $keyWord.Length } else {
    $m = [regex]::Match($reply.HTMLBody, '<body[^>]*>')
    $idx = if ($m.Success) { $m.Index + $m.Length } else { 0 }
}
$reply.HTMLBody = $reply.HTMLBody.Insert($idx, $insertHtml)
$to = [string]$reply.To
$subject = [string]$reply.Subject
$sender = ''
try { $sender = [string]$item.SenderEmailAddress } catch {}
${actionLine}
ConvertTo-Json @{ to = $to; subject = $subject; repliedToSender = $sender }
`;
    try {
        const raw = await runPowerShell(script, 60000);
        const parsed = raw && raw.trim() && raw.trim() !== 'null'
            ? (JSON.parse(raw) as Record<string, unknown>)
            : {};
        return {
            to: String(parsed.to || ''),
            subject: String(parsed.subject || ''),
            repliedToSender: String(parsed.repliedToSender || ''),
        };
    } finally {
        try {
            fs.unlinkSync(bodyFile);
        } catch { /* ignore */
        }
    }
}

/** One draft that couldn't be sent, identified by its subject for the report. */
/**
 * Resolve the Drafts folders belonging to one account into `$scan`, and define
 * `Test-DraftMatches` over them. Emitted into every drafts script so listing,
 * deleting and sending can never disagree about which drafts are "this account's".
 * Expects `$target` (SMTP address), `$outlook` and `$ns` to already be set.
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
$account = $null
foreach ($a in $outlook.Session.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }

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

/**
 * Send the mail drafts that belong to `emailAccount`, wherever Outlook filed them.
 * Which drafts those are is DRAFTS_SCAN_PS's rule, not this function's.
 *
 * Item references are snapshotted before any Send() call: sending moves an item out
 * of Drafts, and mutating the collection mid-enumeration would skip every other one.
 */
export async function sendAllDrafts(emailAccount: string): Promise<SendAllDraftsResult> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')

${DRAFTS_SCAN_PS}
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
    const out = await runPowerShell(script, 300000);
    try {
        const parsed = JSON.parse(out || '{}');
        // ConvertTo-Json unwraps a single-element array to a bare object and an
        // empty one to null, so normalize both back to an array here.
        const raw = parsed.failed;
        const failedList: unknown[] = Array.isArray(raw) ? raw : raw ? [raw] : [];
        return {
            sent: Number(parsed.sent) || 0,
            failed: failedList.map((f) => {
                const rec = f as { subject?: unknown; error?: unknown };
                return {subject: String(rec?.subject ?? ''), error: String(rec?.error ?? '')};
            }),
        };
    } catch {
        return {sent: 0, failed: []};
    }
}

/** One mail draft belonging to an account. */
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
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
${DRAFTS_SCAN_PS}
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
    const raw = await runPowerShell(script, 300000);
    const parsed = raw && raw.trim() && raw.trim() !== 'null'
        ? (JSON.parse(raw) as Record<string, unknown>)
        : {};
    const toArray = (v: unknown): unknown[] => (Array.isArray(v) ? v : v == null ? [] : [v]);
    return {
        account: String(parsed.account || emailAccount),
        foldersScanned: toArray(parsed.foldersScanned).map(f => String(f || '')),
        count: typeof parsed.count === 'number' ? parsed.count : 0,
        truncated: Boolean(parsed.truncated),
        drafts: toArray(parsed.drafts).map(d => {
            const e = d as Record<string, unknown>;
            return {
                entryId: String(e.entryId || ''),
                subject: String(e.subject || ''),
                to: String(e.to || ''),
                toEmails: toArray(e.toEmails).map(a => String(a || '')),
                bodyPreview: String(e.bodyPreview || ''),
                hasAttachments: Boolean(e.hasAttachments),
                lastModified: String(e.lastModified || ''),
                folderPath: String(e.folderPath || ''),
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
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    if (entryIds.length === 0) {
        return {deleted: 0, failed: []};
    }
    const psIds = entryIds.map(id => `'${psEscape(id)}'`).join(',');
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
${DRAFTS_SCAN_PS}
# Index this account's Drafts folders by EntryID, so a resolved item can be PROVED
# to live in one of them before it is deleted.
$draftFolderIds = @{}
foreach ($entry in $scan) { $draftFolderIds[$entry.folder.EntryID] = $entry.includeNull }

$deleted = 0
$failed = @()
foreach ($id in @(${psIds})) {
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
    const raw = await runPowerShell(script, 300000);
    const parsed = raw && raw.trim() && raw.trim() !== 'null'
        ? (JSON.parse(raw) as Record<string, unknown>)
        : {};
    const toArray = (v: unknown): unknown[] => (Array.isArray(v) ? v : v == null ? [] : [v]);
    return {
        deleted: typeof parsed.deleted === 'number' ? parsed.deleted : 0,
        failed: toArray(parsed.failed).map(f => {
            const e = f as Record<string, unknown>;
            return {entryId: String(e.entryId || ''), error: String(e.error || '')};
        }),
    };
}

/** What happened to one id in a deleteOutlookEmails call. */
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
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const {allowProtected = false, dryRun = false} = options;
    if (entryIds.length === 0) {
        return {dryRun, deleted: 0, refused: 0, failed: 0, items: []};
    }
    const psIds = entryIds.map(id => `'${psEscape(id)}'`).join(',');
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }
$store = $account.DeliveryStore
${FOLDER_CHAIN_PS}
$protectedIds = @{}
foreach ($fid in @(${[...PROTECTED_FOLDER_IDS].join(',')})) {
    try { $protectedIds[$store.GetDefaultFolder($fid).EntryID] = $true } catch {}
}

$items = @()
$deleted = 0
$refused = 0
$failed = 0
foreach ($id in @(${psIds})) {
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

        if ($isProtected -and -not ${allowProtected ? '$true' : '$false'}) {
            $refused++
            $items += [pscustomobject]@{ entryId = $id; subject = $subject; folderPath = $folderPath; status = 'refused'; reason = 'received or sent mail (Inbox/Sent Items or a subfolder) - pass allow_protected to override' }
        } elseif (${dryRun ? '$true' : '$false'}) {
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
ConvertTo-Json @{ dryRun = ${dryRun ? '$true' : '$false'}; deleted = $deleted; refused = $refused; failed = $failed; items = @($items) } -Depth 4
`;
    const raw = await runPowerShell(script, 300000);
    const parsed = raw && raw.trim() && raw.trim() !== 'null'
        ? (JSON.parse(raw) as Record<string, unknown>)
        : {};
    const toArray = (v: unknown): unknown[] => (Array.isArray(v) ? v : v == null ? [] : [v]);
    return {
        dryRun,
        deleted: typeof parsed.deleted === 'number' ? parsed.deleted : 0,
        refused: typeof parsed.refused === 'number' ? parsed.refused : 0,
        failed: typeof parsed.failed === 'number' ? parsed.failed : 0,
        items: toArray(parsed.items).map(i => {
            const e = i as Record<string, unknown>;
            return {
                entryId: String(e.entryId || ''),
                subject: String(e.subject || ''),
                folderPath: String(e.folderPath || ''),
                status: String(e.status || 'failed') as DeleteMailOutcome['status'],
                reason: String(e.reason || ''),
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
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }
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
    if (${dryRun ? '$true' : '$false'}) { continue }
    try { $it.Delete(); $purged++ } catch { $failed++ }
}
ConvertTo-Json @{ folderPath = $folder.FolderPath; matched = $matched; purged = $purged; kept = $kept; failed = $failed } -Depth 3
`;
    const raw = await runPowerShell(script, 600000);
    const parsed = raw && raw.trim() && raw.trim() !== 'null'
        ? (JSON.parse(raw) as Record<string, unknown>)
        : {};
    const num = (v: unknown): number => (typeof v === 'number' ? v : 0);
    return {
        folderPath: String(parsed.folderPath || ''),
        dryRun,
        matched: num(parsed.matched),
        purged: num(parsed.purged),
        kept: num(parsed.kept),
        failed: num(parsed.failed),
    };
}

// ── Undeliverable / bounce cleanup types ────────────────────────────────────
/** One bounce-back / non-delivery message found in the inbox. */
// Bounce-classifier lists, defined once per script (before the item loop). Kept
// specific — multi-word, mail-system wording — so ordinary mail that merely
// mentions a 'delivery' is never flagged.
const BOUNCE_LISTS_PS = `
$phrases = @(
    'undeliverable',
    'message blocked',
    'mail delivery failed',
    'delivery status notification (failure)',
    'delivery has failed',
    'failure notice',
    'undelivered mail returned to sender',
    'returned mail'
)
$daemonAddr = @('mailer-daemon', 'mailer_daemon', 'postmaster@', 'mail-daemon', 'maildelivery')
$daemonName = @('mail delivery subsystem', 'mail delivery system', 'microsoft outlook', 'postmaster', 'mailer-daemon', 'mail administrator', 'internet mail delivery')
`;

// Per-item classifier. Given an Outlook item in $item and the account address
// (lowercased) in $targetLc, sets $reason to why it's a bounce ('' if not) and
// fills $failedRcpts with the recipient addresses parsed from the bounce body.
// Also leaves $subj/$sName/$sEmail populated for the caller's report. The body
// is only read once a cheaper signal (class/sender/subject) has already matched.
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
    if ($msgClass -like 'REPORT.IPM.Note.NDR*') { $reason = 'Non-delivery report (NDR)' }
    if ($reason -eq '') { foreach ($d in $daemonAddr) { if ($sEmailLc.Contains($d)) { $reason = 'From mail-delivery system'; break } } }
    if ($reason -eq '') { foreach ($d in $daemonName) { if ($sNameLc.Contains($d)) { $reason = 'From mail-delivery system'; break } } }
    if ($reason -eq '') { foreach ($p in $phrases) { if ($subjLc.Contains($p)) { $reason = "Bounce subject phrase: '$p'"; break } } }
    if ($reason -ne '') {
        $bodyRaw = ''
        try { if ($item.Body) { $bodyRaw = [string]$item.Body } } catch {}
        if ($bodyRaw -ne '') {
            $mm = [regex]::Matches($bodyRaw, '[A-Za-z0-9._%+\\-]+@[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}')
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
 * daemon/postmaster, or its subject contains a specific bounce phrase. The item
 * body is read only after a match, so scanning a large inbox stays cheap.
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
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const days = Math.max(1, Math.min(365, Math.floor(daysBack)));
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$dryRun = ${dryRun ? '$true' : '$false'}
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }
$storeFolder = $null
foreach ($f in $ns.Folders) {
    if ($f.Name -ieq $account.DisplayName) { $storeFolder = $f; break }
}
if ($storeFolder -eq $null) { throw "Store folder not found for account '$target'" }
$inbox = $storeFolder.Store.GetDefaultFolder(6)
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
    const out = await runPowerShell(script, 300000);
    const parsed = JSON.parse(out || '{}');
    const toArray = (v: unknown): unknown[] => (Array.isArray(v) ? v : v ? [v] : []);
    return {
        account: String(parsed.account ?? emailAccount),
        scannedDays: Number(parsed.scannedDays) || days,
        dryRun: parsed.dryRun !== false,
        matchedCount: Number(parsed.matchedCount) || 0,
        deletedCount: Number(parsed.deletedCount) || 0,
        matched: toArray(parsed.matched).map((m) => {
            const e = m as Record<string, unknown>;
            return {
                entryId: String(e.entryId ?? ''),
                subject: String(e.subject ?? ''),
                senderName: String(e.senderName ?? ''),
                senderEmail: String(e.senderEmail ?? ''),
                receivedTime: String(e.receivedTime ?? ''),
                matchedReason: String(e.matchedReason ?? ''),
                failedRecipients: toArray(e.failedRecipients).map(String),
            };
        }),
        failed: toArray(parsed.failed).map((f) => {
            const rec = f as { subject?: unknown; error?: unknown };
            return {subject: String(rec?.subject ?? ''), error: String(rec?.error ?? '')};
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
    const days = Math.max(1, Math.min(365, Math.floor(daysBack)));
    // 6 = Inbox, 3 = Deleted Items.
    const folderIds = scanDeleted ? '@(6, 3)' : '@(6)';
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }
$store = $account.DeliveryStore
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
    const raw = await runPowerShell(script, 300000);
    if (!raw || raw.trim() === '' || raw.trim() === 'null') return [];
    const parsed = JSON.parse(raw);
    const arr: unknown[] = Array.isArray(parsed) ? parsed : [parsed];
    return arr.map(String).filter(Boolean);
}

// ── Sent-recipient groups ───────────────────────────────────────────────────
/** One sent message and the full set of SMTP addresses it went to. */
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
    const days = Math.max(1, Math.min(365, Math.floor(daysBack)));
    const cap = Math.max(1, Math.min(10000, Math.floor(limit)));
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }
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
    const raw = await runPowerShell(script, 300000);
    if (!raw || raw.trim() === '' || raw.trim() === 'null') return [];
    const parsed = JSON.parse(raw);
    const arr: unknown[] = Array.isArray(parsed) ? parsed : [parsed];
    return arr.map((it) => {
        const e = it as Record<string, unknown>;
        const recips = Array.isArray(e.recipients)
            ? (e.recipients as unknown[]).map(String)
            : e.recipients ? [String(e.recipients)] : [];
        return {
            entryId: String(e.entryId ?? ''),
            subject: String(e.subject ?? ''),
            sentOn: String(e.sentOn ?? ''),
            recipients: recips,
        };
    });
}

// Outlook stores each signature as "<name>.htm" (plus .rtf/.txt and a "<name>_files"
// folder for images) under %APPDATA%\Microsoft\Signatures.
const SIGNATURES_DIR = process.env.APPDATA
    ? path.join(process.env.APPDATA, 'Microsoft', 'Signatures')
    : '';

/**
 * Names of the user's Outlook signatures (the ".htm" files), sorted. Empty off
 * Windows. Reads from disk synchronously but is declared async to match macOS,
 * which has to ask Outlook itself — one signature for both platforms.
 */
export async function listOutlookSignatures(): Promise<string[]> {
    if (process.platform !== 'win32' || !SIGNATURES_DIR) return [];
    try {
        return fs.readdirSync(SIGNATURES_DIR)
            .filter(f => f.toLowerCase().endsWith('.htm'))
            .map(f => f.slice(0, -4))
            .sort((a, b) => a.localeCompare(b));
    } catch {
        return [];
    }
}

/**
 * Read a named signature's HTML with its image references rewritten to absolute
 * file: URIs. Outlook stores signature images relative to a "<name>_files" folder;
 * once the refs are absolute, assigning the HTML to a mail body lets Outlook resolve
 * and embed the images on Display/Send. Returns '' if the signature can't be found.
 */
export async function readOutlookSignatureHtml(name: string): Promise<string> {
    if (process.platform !== 'win32' || !SIGNATURES_DIR) return '';
    // Only accept a bare signature name — never a path — so a crafted name can't
    // escape the Signatures folder.
    const safe = path.basename(name);
    const file = path.join(SIGNATURES_DIR, `${safe}.htm`);
    if (!fs.existsSync(file)) return '';
    // Classic Outlook signatures are saved as windows-1252, not UTF-8; decode by the
    // charset the file declares so accented text / smart quotes survive.
    const buf = fs.readFileSync(file);
    const head = buf.toString('latin1', 0, 2048);
    const charset = head.match(/charset=["']?([\w-]+)/i)?.[1] || 'utf-8';
    let html: string;
    try {
        html = new TextDecoder(charset).decode(buf);
    } catch {
        html = buf.toString('utf-8');
    }
    // pathToFileURL encodes spaces/specials the same way Outlook's relative refs are,
    // so prefixing the (already relative, already-encoded) src keeps a valid URI.
    const dirUri = pathToFileURL(SIGNATURES_DIR + path.sep).href;
    html = html.replace(
        /(src|background)=(["'])(?!https?:|cid:|data:|file:|mailto:|#)/gi,
        `$1=$2${dirUri}`
    );
    return html;
}

// ── Generic inbox search ──────────────────────────────────────────────────
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
    const subjectLikeDecl = `$subjLike = '${subjectLike}'`;
    const boundedRestrict = subjectLike
        ? `$folder.Items.Restrict("[ReceivedTime] >= '$cutoff' AND [Subject] like '" + $subjLike + "'")`
        : `$folder.Items.Restrict("[ReceivedTime] >= '$cutoff'")`;
    const unboundedRestrict = subjectLike
        ? `$folder.Items.Restrict("[Subject] like '" + $subjLike + "'")`
        : `$folder.Items`;
    const script = `
${subjectLikeDecl}
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq '${psEscape(emailAccount)}') { $account = $a; break }
}
if ($account -eq $null) { throw "Account '${psEscape(emailAccount)}' not found" }
$storeFolder = $null
foreach ($f in $ns.Folders) {
    if ($f.Name -ieq $account.DisplayName) { $storeFolder = $f; break }
}
if ($storeFolder -eq $null) { throw "Store folder not found for account '${psEscape(emailAccount)}'" }
$inbox = $storeFolder.Store.GetDefaultFolder(6)
$storeId = $storeFolder.Store.StoreID
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
        # An Exchange sender's SenderEmailAddress is an X500 DN, not an address —
        # resolve the real SMTP address when we can, fall back to whatever is there.
        $senderSmtp = ''
        try {
            if ($item.SenderEmailType -eq 'EX') {
                $senderSmtp = $item.Sender.GetExchangeUser().PrimarySmtpAddress
            }
        } catch {}
        if (-not $senderSmtp) {
            $senderSmtp = if ($item.SenderEmailAddress) { $item.SenderEmailAddress } else { '' }
        }
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
    const raw = await runPowerShell(script, 300000);
    if (!raw || raw.trim() === '' || raw.trim() === 'null') return [];
    const parsed = JSON.parse(raw);
    const arr: unknown[] = Array.isArray(parsed) ? parsed : [parsed];
    return arr.map(item => {
        const e = item as Record<string, unknown>;
        return {
            entryId: String(e.entryId || ''),
            storeId: String(e.storeId || ''),
            subject: String(e.subject || ''),
            senderName: String(e.senderName || ''),
            senderEmail: String(e.senderEmail || ''),
            receivedTime: String(e.receivedTime || ''),
            body: Buffer.from(String(e.body || ''), 'base64').toString('utf8'),
            attachmentNames: Array.isArray(e.attachmentNames)
                ? (e.attachmentNames as unknown[]).map(String)
                : [],
            folderPath: String(e.folderPath || ''),
        };
    });
}

/**
 * Read the email currently selected (or open) in Outlook — full body,
 * attachments listed by name but not saved.
 */
export async function readSelectedEmail(): Promise<SelectedEmail> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
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
$senderSmtp = ''
try {
    if ($item.SenderEmailType -eq 'EX') {
        $senderSmtp = $item.Sender.GetExchangeUser().PrimarySmtpAddress
    }
} catch {}
if (-not $senderSmtp) {
    $senderSmtp = if ($item.SenderEmailAddress) { $item.SenderEmailAddress } else { '' }
}
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
        throw new Error('Failed to read the selected Outlook email.');
    }
    const e = JSON.parse(raw) as Record<string, unknown>;
    return {
        entryId: String(e.entryId || ''),
        storeId: String(e.storeId || ''),
        subject: String(e.subject || ''),
        senderName: String(e.senderName || ''),
        senderEmail: String(e.senderEmail || ''),
        receivedTime: String(e.receivedTime || ''),
        body: Buffer.from(String(e.body || ''), 'base64').toString('utf8'),
        attachmentNames: Array.isArray(e.attachmentNames)
            ? (e.attachmentNames as unknown[]).map(String)
            : e.attachmentNames ? [String(e.attachmentNames)] : [],
    };
}

/**
 * Open an email in Outlook by its EntryID.
 */
export async function openOutlookEmail(entryId: string): Promise<void> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const script = `(New-Object -ComObject Outlook.Application).GetNamespace('mapi').GetItemFromID('${psEscape(entryId)}').Display()`;
    await runPowerShell(script);
}

/**
 * Open an Outlook draft pre-filled with the current reply template HTML.
 * Polls until the user closes the draft window, then returns the edited HTML body.
 */
export async function editEmailTemplate(label: string, currentHtml: string): Promise<string> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const inputFile = tempFile('template-input', 'html');
    const outputFile = tempFile('template-output', 'html');
    const scriptFile = tempFile('template-script', 'ps1');

    fs.writeFileSync(inputFile, currentHtml, 'utf-8');

    const subject = `${label} - Save (Ctrl+S) and close when done`;
    const psScript = `
$ErrorActionPreference = 'Stop'
$inputPath = '${psEscape(inputFile)}'
$outputPath = '${psEscape(outputFile)}'
$subject = '${psEscape(subject)}'

$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$mail = $outlook.CreateItem(0)
$mail.Subject = $subject
$mail.HTMLBody = [IO.File]::ReadAllText($inputPath, [Text.Encoding]::UTF8)
$mail.Display()

# Give the inspector time to fully initialize
Start-Sleep -Seconds 2

# Poll until the compose window is closed, keeping the last live body as a
# fallback for the case where the user edits but never presses Ctrl+S.
$liveBody = ''
while ($true) {
    Start-Sleep -Milliseconds 700
    $open = $false
    try {
        foreach ($insp in $outlook.Inspectors) {
            $ci = $null
            try { $ci = $insp.CurrentItem } catch {}
            if ($ci -ne $null -and $ci.Subject -eq $subject) {
                $open = $true
                try { $liveBody = $ci.HTMLBody } catch {}
                break
            }
        }
    } catch {}
    if (-not $open) { break }
}

# Prefer the saved draft — it reflects the user's Ctrl+S and survives the
# inspector closing, unlike the original mail reference. Also removes any
# leftover editor drafts so they don't pile up in the Drafts folder.
$savedBody = ''
try {
    $drafts = $ns.GetDefaultFolder(16)
    for ($i = $drafts.Items.Count; $i -ge 1; $i--) {
        $it = $drafts.Items.Item($i)
        if ($it.Subject -eq $subject) {
            if (-not $savedBody) { try { $savedBody = $it.HTMLBody } catch {} }
            try { $it.Delete() } catch {}
        }
    }
} catch {}

$finalBody = if ($savedBody) { $savedBody } else { $liveBody }
if ($finalBody) {
    [IO.File]::WriteAllText($outputPath, $finalBody, [Text.Encoding]::UTF8)
}
`;
    fs.writeFileSync(scriptFile, psScript, 'utf-8');

    await new Promise<void>((resolve, reject) => {
        execFile(
            'powershell.exe',
            ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptFile],
            {timeout: 0},
            (error) => {
                if (error) reject(new Error(error.message));
                else resolve();
            }
        );
    });

    try {
        fs.unlinkSync(scriptFile);
        fs.unlinkSync(inputFile);
    } catch { /* ignore */
    }

    try {
        const result = fs.readFileSync(outputFile, 'utf-8');
        fs.unlinkSync(outputFile);
        return result;
    } catch {
        throw new Error('No template saved. Did you save (Ctrl+S) before closing?');
    }
}

// ── Inbox email types ───────────────────────────────────────────────────
/** Folders whose contents are received or already-sent mail, not working state. */
const PROTECTED_FOLDER_IDS = new Set([6, 5]);

/**
 * Emit the PowerShell that resolves `$scope` from `$store`, walking a well-known
 * root down through any further segments.
 *
 * Split out so the emitted script can be parse-checked without an Outlook session —
 * the folder walk is the only part that varies per call, and a syntax error in it
 * would only ever surface as a failed live run.
 *
 * `folderLabel` is the caller's original string, used verbatim in the error so the
 * message names what they typed rather than the parsed segments. Set `createMissing`
 * to build absent segments instead of throwing — the whole chain, so a nested
 * destination can be created in one call.
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
    const ref = folder ? mailFolderRef(folder) : {rootId: 6, rootLabel: 'Inbox', segments: []};
    // A folder argument that trims away to nothing ("\\", "  ") would otherwise
    // read the Inbox root and look like it had scoped — the exact silent
    // mis-scoping this parameter exists to prevent. A bare well-known name
    // ("Sent Items") legitimately has no segments, so it is not that case.
    if (folder && folder.trim() && ref.segments.length === 0 && ref.rootId === 6
        && !Object.prototype.hasOwnProperty.call(WELL_KNOWN_FOLDERS, folder.trim().toLowerCase())) {
        throw new Error(`Folder '${folder}' does not name a folder under the Inbox.`);
    }
    const resolveScope = mailScopeScript(ref, folder || '');
    // Which timestamp the folder's items actually carry (see the note in the script).
    const dateProp = ref.rootId === 5 || ref.rootId === 4 ? 'SentOn'
        : ref.rootId === 16 ? 'LastModificationTime'
            : 'ReceivedTime';
    const isOutgoing = ref.rootId === 5 || ref.rootId === 4 || ref.rootId === 16 ? '$true' : '$false';
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq '${psEscape(emailAccount)}') { $account = $a; break }
}
if ($account -eq $null) { throw "Account '${psEscape(emailAccount)}' not found" }
$storeFolder = $null
foreach ($f in $ns.Folders) {
    if ($f.Name -ieq $account.DisplayName) { $storeFolder = $f; break }
}
if ($storeFolder -eq $null) { throw "Store folder not found for account '${psEscape(emailAccount)}'" }
$store = $storeFolder.Store
$storeId = $store.StoreID
${resolveScope}
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
    const raw = await runPowerShell(script, 30000);
    if (!raw || raw.trim() === '' || raw.trim() === 'null') return [];
    const parsed = JSON.parse(raw);
    const arr: unknown[] = Array.isArray(parsed) ? parsed : [parsed];
    return arr.map(item => {
        const e = item as Record<string, unknown>;
        return {
            entryId: String(e.entryId || ''),
            storeId: String(e.storeId || ''),
            folderPath: String(e.folderPath || ''),
            subject: String(e.subject || ''),
            senderName: String(e.senderName || ''),
            senderEmail: String(e.senderEmail || ''),
            receivedTime: String(e.receivedTime || ''),
            bodyPreview: String(e.bodyPreview || ''),
            attachmentNames: Array.isArray(e.attachmentNames)
                ? (e.attachmentNames as unknown[]).map(String)
                : typeof e.attachmentNames === 'string' ? [e.attachmentNames] : [],
            attachmentCount: typeof e.attachmentCount === 'number' ? e.attachmentCount : 0,
        };
    });
}

// ── Full email body ─────────────────────────────────────────────────────
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
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const getItemCall = storeId
        ? `$ns.GetItemFromID('${psEscape(entryId)}', '${psEscape(storeId)}')`
        : `$ns.GetItemFromID('${psEscape(entryId)}')`;
    // The body is base64'd across the PowerShell boundary: an arbitrary email body
    // carries quotes, control characters and non-ASCII that the console codepage
    // would otherwise mangle on the way into JSON.
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$item = ${getItemCall}
if ($item -eq $null) { throw "Email not found for EntryID '${psEscape(entryId)}'" }
$cls = 0
try { $cls = [int]$item.Class } catch {}
if ($cls -ne 43 -and $cls -ne 46) { throw "The item for this EntryID is not an email (Class=$cls). Re-run readInboxEmails for a current EntryID." }
$bodyRaw = ''
try { if ($item.Body) { $bodyRaw = [string]$item.Body } } catch {}
$attNames = @()
try { foreach ($att in $item.Attachments) { $attNames += $att.FileName } } catch {}
# An Exchange sender's SenderEmailAddress is an X500 DN, not an address — resolve
# the real SMTP address where we can so the caller can confirm the sender.
$senderSmtp = ''
try {
    if ($item.SenderEmailType -eq 'EX') { $senderSmtp = $item.Sender.GetExchangeUser().PrimarySmtpAddress }
} catch {}
if (-not $senderSmtp) { $senderSmtp = if ($item.SenderEmailAddress) { $item.SenderEmailAddress } else { '' } }
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
    if (!raw || !raw.trim()) throw new Error('Failed to read email body');
    const parsed = JSON.parse(raw.trim()) as Record<string, unknown>;
    const full = Buffer.from(String(parsed.body || ''), 'base64').toString('utf8');
    const {body, quoted, separator} = splitQuotedOriginal(full);
    const attachmentNames = Array.isArray(parsed.attachmentNames)
        ? (parsed.attachmentNames as unknown[]).map(String)
        : typeof parsed.attachmentNames === 'string' ? [String(parsed.attachmentNames)] : [];
    // The quoted thread is context, never the sender's own answer, so it is capped
    // harder than the reply itself — its useful part (what was asked) is at the
    // top of it.
    const quotedCap = Math.min(maxChars, 4000);
    return {
        entryId: String(parsed.entryId || ''),
        subject: String(parsed.subject || ''),
        senderName: String(parsed.senderName || ''),
        senderEmail: String(parsed.senderEmail || ''),
        receivedTime: String(parsed.receivedTime || ''),
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

// ── Outlook template-email folder ────────────────────────────────────────
/** One template email stored in the mailbox's Templates folder. */
/** Recursive folder-by-name search, emitted into the template-folder scripts.
 *  Case-insensitive, depth-capped so a huge mailbox tree can't hang the scan. */
const FIND_FOLDER_PS = `
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
 * Read the template emails saved in a mailbox folder (default "Templates"):
 * regular mail items kept as reusable reply bodies. Returns each
 * item's full HTML body. When the folder doesn't exist, returns folderFound:false
 * plus the mailbox's folder names instead of throwing, so the caller can ask the
 * user whether to create it rather than fail.
 */
export async function readTemplateEmails(
    emailAccount: string,
    folderName = 'Templates',
    limit = 20,
    includeBody = true,
    subject = '',
): Promise<TemplateFolderResult> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const cap = Math.max(1, Math.min(50, Math.floor(limit)));
    const wanted = (subject || '').trim();
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }
$root = $account.DeliveryStore.GetRootFolder()
${FIND_FOLDER_PS}
$folder = Find-FolderByName $root '${psEscape(folderName)}' 3
if ($folder -eq $null) {
    $names = @()
    foreach ($f in $root.Folders) { $names += $f.Name }
    ConvertTo-Json @{ folderFound = $false; folderPath = ''; templates = @(); availableFolders = $names } -Depth 3
    exit 0
}
$includeBody = ${includeBody ? '$true' : '$false'}
$wanted = '${psEscape(wanted)}'
$items = $folder.Items
$items.Sort('[LastModificationTime]', $true)
$results = @()
# Without a subject filter the sort order does the capping (newest first). With one,
# scan the whole folder — the wanted template need not be among the newest — and stop
# as soon as the cap is filled.
$count = $items.Count
for ($i = 1; $i -le $count; $i++) {
    if ($results.Count -ge ${cap}) { break }
    $item = $null
    try { $item = $items.Item($i) } catch { continue }
    if ($item -eq $null) { continue }
    $cls = 0
    try { $cls = [int]$item.Class } catch {}
    if ($cls -ne 43) { continue }  # olMail only
    if ($wanted -ne '') {
        $subj = ''
        try { $subj = [string]$item.Subject } catch {}
        if ($subj -eq $null) { $subj = '' }
        if ($subj.Trim() -ine $wanted) { continue }
    }
    $html = ''
    $preview = ''
    $markers = ''
    $plain = ''
    if ($includeBody) {
        try { $html = [string]$item.HTMLBody } catch {}
        # Embedded (cid:) images live as attachments on the TEMPLATE item: reusing this HTML
        # on a new email shows a broken "linked image" placeholder, and rewriting to file:
        # URIs renders invisibly in modern Outlook. Neither works, so strip the img tags —
        # template emails are text/HTML only.
        $html = [regex]::Replace($html, '<img[^>]*src="cid:[^"]*"[^>]*>', '')
        if ($html.Length -gt 100000) { $html = $html.Substring(0, 100000) }
    } else {
        # Bodies omitted: return a short plain-text preview so the caller can tell templates
        # apart by subject without pulling the full (often huge) Word-generated HTML.
        try { $plain = [string]$item.Body } catch { $plain = '' }
        if ($plain) {
            # ...plus the [[SECTION]] / {{PLACEHOLDER}} markers, scanned from the full text:
            # that's what lets a caller confirm a sectioned template is intact without
            # fetching its body. Word can split a marker across tags in the HTML but not
            # in the plain-text body, so a literal scan is right here.
            # (Backslashes in the regex are doubled because this whole script is a JS
            # template literal — a single one would be eaten before PowerShell saw it.)
            $found = @()
            foreach ($x in [regex]::Matches($plain, '\\[\\[\\s*/?\\s*[A-Za-z0-9_-]{1,40}\\s*\\]\\]|\\{\\{\\s*[A-Za-z0-9_-]{1,40}\\s*\\}\\}')) {
                $t = ($x.Value -replace '\\s', '')
                if ($found -notcontains $t) { $found += $t }
            }
            $markers = ($found -join ' ')
            $preview = $plain.Trim()
            if ($preview.Length -gt 200) { $preview = $preview.Substring(0, 200) }
        }
    }
    $results += [PSCustomObject]@{
        entryId      = $item.EntryID
        subject      = if ($item.Subject) { $item.Subject.Trim() } else { '' }
        htmlBody     = $html
        bodyPreview  = $preview
        markers      = $markers
        lastModified = $item.LastModificationTime.ToString('yyyy-MM-dd HH:mm')
    }
}
ConvertTo-Json @{ folderFound = $true; folderPath = $folder.FolderPath; templates = $results; availableFolders = @() } -Depth 4
`;
    const raw = await runPowerShell(script, 60000);
    if (!raw || raw.trim() === '' || raw.trim() === 'null') {
        return {folderFound: false, folderPath: '', templates: [], availableFolders: []};
    }
    const parsed = JSON.parse(raw) as Record<string, unknown>;
    const toArray = (v: unknown): unknown[] => (Array.isArray(v) ? v : v == null ? [] : [v]);
    return {
        folderFound: Boolean(parsed.folderFound),
        folderPath: String(parsed.folderPath || ''),
        templates: toArray(parsed.templates).map(item => {
            const e = item as Record<string, unknown>;
            const html = String(e.htmlBody || '');
            // With a body in hand the HTML is authoritative (it survives Word splitting a
            // marker across tags); without one, fall back to what the script scanned out
            // of the plain-text body.
            const markers = findTemplateMarkers(html || String(e.markers || ''));
            return {
                entryId: String(e.entryId || ''),
                subject: String(e.subject || ''),
                htmlBody: html,
                bodyPreview: String(e.bodyPreview || ''),
                sections: markers.sections,
                placeholders: markers.placeholders,
                lastModified: String(e.lastModified || ''),
            };
        }),
        availableFolders: toArray(parsed.availableFolders).map(String),
    };
}

/**
 * Save a new template email into a mailbox folder (default "Templates"),
 * creating the folder at the mailbox root if it doesn't exist. The item is a
 * plain unsent mail (subject + HTML body) that can be edited in Outlook.
 * Never overwrites an existing item — it always adds a new one, so check with
 * readTemplateEmails first and only call after the user has agreed to create it.
 */
export async function saveTemplateEmail(
    emailAccount: string,
    subject: string,
    htmlBody: string,
    folderName = 'Templates',
): Promise<SaveTemplateResult> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    // Body goes through a temp file (same reason as sendOutlookEmail): inline
    // -Command scripts hit the command-line length limit on large HTML.
    const bodyFile = tempFile('template-body', 'html');
    fs.writeFileSync(bodyFile, htmlBody, 'utf-8');
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }
$root = $account.DeliveryStore.GetRootFolder()
${FIND_FOLDER_PS}
$folder = Find-FolderByName $root '${psEscape(folderName)}' 3
$folderCreated = $false
if ($folder -eq $null) {
    $folder = $root.Folders.Add('${psEscape(folderName)}')
    $folderCreated = $true
}
$mail = $outlook.CreateItem(0)
$mail.Subject = '${psEscape(subject)}'
$mail.HTMLBody = [IO.File]::ReadAllText('${psEscape(bodyFile)}', [Text.Encoding]::UTF8)
$mail.Save()
$moved = $mail.Move($folder)
ConvertTo-Json @{ folderPath = $folder.FolderPath; folderCreated = $folderCreated }
`;
    try {
        const raw = await runPowerShell(script, 60000);
        const parsed = raw && raw.trim() && raw.trim() !== 'null'
            ? (JSON.parse(raw) as Record<string, unknown>)
            : {};
        return {
            folderPath: String(parsed.folderPath || ''),
            folderCreated: Boolean(parsed.folderCreated),
        };
    } finally {
        try {
            fs.unlinkSync(bodyFile);
        } catch { /* ignore */
        }
    }
}

// ── Inbox folder listing & filing ────────────────────────────────────────
/** One folder under the account's Inbox. */
/** List the folders under an account's Inbox (the user's filing folders). */
export async function listInboxFolders(
    emailAccount: string,
    maxDepth = 2,
): Promise<InboxFolderInfo[]> {
    if (process.platform !== 'win32') return [];
    const depth = Math.max(1, Math.min(4, Math.floor(maxDepth)));
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }
$inbox = $account.DeliveryStore.GetDefaultFolder(6)
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
    const raw = await runPowerShell(script, 60000);
    if (!raw || raw.trim() === '' || raw.trim() === 'null') return [];
    const parsed = JSON.parse(raw);
    const arr: unknown[] = Array.isArray(parsed) ? parsed : [parsed];
    return arr.map(item => {
        const e = item as Record<string, unknown>;
        return {
            name: String(e.name || ''),
            folderPath: String(e.folderPath || ''),
            itemCount: typeof e.itemCount === 'number' ? e.itemCount : 0,
            depth: typeof e.depth === 'number' ? e.depth : 1,
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
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    if (entryIds.length === 0) {
        return {folderPath: '', folderCreated: false, moved: 0, failed: []};
    }
    const ref = mailFolderRef(folderName);
    const psIds = entryIds.map(id => `'${psEscape(id)}'`).join(',');
    const script = `
$ErrorActionPreference = 'Stop'
$target = '${psEscape(emailAccount)}'
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$account = $null
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress -ieq $target) { $account = $a; break }
}
if ($account -eq $null) { throw "Account '$target' not found" }
$store = $account.DeliveryStore
${mailScopeScript(ref, folderName, createIfMissing)}
$folder = $scope
$folderCreated = $scopeCreated
$moved = 0
$failed = @()
foreach ($id in @(${psIds})) {
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
    const raw = await runPowerShell(script, 300000);
    const parsed = raw && raw.trim() && raw.trim() !== 'null'
        ? (JSON.parse(raw) as Record<string, unknown>)
        : {};
    const toArray = (v: unknown): unknown[] => (Array.isArray(v) ? v : v == null ? [] : [v]);
    return {
        folderPath: String(parsed.folderPath || ''),
        folderCreated: Boolean(parsed.folderCreated),
        moved: typeof parsed.moved === 'number' ? parsed.moved : 0,
        failed: toArray(parsed.failed).map(f => {
            const e = f as Record<string, unknown>;
            return {entryId: String(e.entryId || ''), error: String(e.error || '')};
        }),
    };
}

/**
 * Save an email attachment by entryId and filename, returning the saved path
 * together with the resolved email's subject/sender so the caller can confirm the
 * file came from the email it intended.
 *
 * `destDir` is created if absent. Without one the file lands in a fresh directory
 * of its own, because attachments keep the name the sender gave them: a single
 * shared folder means two emails carrying "invoice.pdf" silently overwrite each
 * other. Pass `destDir` whenever you want the files somewhere you control.
 *
 * When the attachment isn't found, throws an error that names the email actually
 * resolved (from/subject/received) and the attachments it does carry. Because
 * GetItemFromID happily returns a *different* message when handed a stale or wrong
 * EntryID (common when many replies share one subject), a bare "not found" is
 * misleading — this surfaces which email you're really looking at. Matching is
 * exact first, then whitespace-normalized (trim + collapse runs, incl. non-breaking
 * spaces) so a trivially reformatted filename still resolves.
 */
export async function saveEmailAttachmentDetailed(
    entryId: string,
    fileName: string,
    storeId?: string,
    destDir?: string,
): Promise<SavedAttachment> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    const outDir = resolveDestDir(destDir);
    const getItemCall = storeId
        ? `$ns.GetItemFromID('${psEscape(entryId)}', '${psEscape(storeId)}')`
        : `$ns.GetItemFromID('${psEscape(entryId)}')`;
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$item = ${getItemCall}
if ($item -eq $null) { throw "Email not found for EntryID '${psEscape(entryId)}'" }
$target = '${psEscape(fileName)}'
$found = $null
foreach ($att in $item.Attachments) {
    if ($att.FileName -eq $target) { $found = $att; break }
}
if ($found -eq $null) {
    $tnorm = ($target -replace '\\s+', ' ').Trim()
    foreach ($att in $item.Attachments) {
        if ((($att.FileName -replace '\\s+', ' ').Trim()) -ieq $tnorm) { $found = $att; break }
    }
}
if ($found -eq $null) {
    $have = @(); foreach ($a in $item.Attachments) { $have += $a.FileName }
    $haveStr = if ($have.Count -gt 0) { $have -join ', ' } else { '(none)' }
    throw "Attachment '$target' not found. Resolved email: from=$($item.SenderEmailAddress); subject=$($item.Subject); received=$($item.ReceivedTime). Attachments present: $haveStr. If this is not the email you expected, the EntryID is likely wrong or stale - re-run readInboxEmails to get a current EntryID."
}
$savePath = [IO.Path]::Combine('${psEscape(outDir)}', $found.FileName)
$found.SaveAsFile($savePath)
$out = [PSCustomObject]@{
    path         = $savePath
    subject      = if ($item.Subject) { $item.Subject } else { '' }
    senderName   = if ($item.SenderName) { $item.SenderName } else { '' }
    senderEmail  = if ($item.SenderEmailAddress) { $item.SenderEmailAddress } else { '' }
    receivedTime = if ($item.ReceivedTime) { $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } else { '' }
}
ConvertTo-Json $out -Compress
`;
    const raw = await runPowerShell(script, 15000);
    if (!raw || !raw.trim()) throw new Error('Failed to save attachment');
    const parsed = JSON.parse(raw.trim());
    return {
        path: String(parsed.path || ''),
        subject: String(parsed.subject || ''),
        senderName: String(parsed.senderName || ''),
        senderEmail: String(parsed.senderEmail || ''),
        receivedTime: String(parsed.receivedTime || ''),
    };
}

/** Save an attachment and return just the saved file path. */
export async function saveEmailAttachment(
    entryId: string,
    fileName: string,
    storeId?: string,
    destDir?: string,
): Promise<string> {
    return (await saveEmailAttachmentDetailed(entryId, fileName, storeId, destDir)).path;
}

/**
 * Save several attachments from one email in a single COM round trip — what a
 * caller working through searchInboxByFilter/readSelectedEmail results wants,
 * rather than paying a PowerShell process spawn per attachment. Destination,
 * matching and the not-found error follow saveEmailAttachmentDetailed's rules
 * exactly, applied per name; results come back in the same order as `fileNames`.
 */
export async function saveEmailAttachments(
    entryId: string,
    fileNames: string[],
    storeId?: string,
    destDir?: string,
): Promise<SavedAttachment[]> {
    if (process.platform !== 'win32') {
        throw new Error('Outlook COM automation is only supported on Windows.');
    }
    if (fileNames.length === 0) return [];
    const outDir = resolveDestDir(destDir);
    const getItemCall = storeId
        ? `$ns.GetItemFromID('${psEscape(entryId)}', '${psEscape(storeId)}')`
        : `$ns.GetItemFromID('${psEscape(entryId)}')`;
    const targetsList = fileNames.map(f => `'${psEscape(f)}'`).join(',');
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$item = ${getItemCall}
if ($item -eq $null) { throw "Email not found for EntryID '${psEscape(entryId)}'" }
$targets = @(${targetsList})
$results = @()
foreach ($target in $targets) {
    $found = $null
    foreach ($att in $item.Attachments) {
        if ($att.FileName -eq $target) { $found = $att; break }
    }
    if ($found -eq $null) {
        $tnorm = ($target -replace '\\s+', ' ').Trim()
        foreach ($att in $item.Attachments) {
            if ((($att.FileName -replace '\\s+', ' ').Trim()) -ieq $tnorm) { $found = $att; break }
        }
    }
    if ($found -eq $null) {
        $have = @(); foreach ($a in $item.Attachments) { $have += $a.FileName }
        $haveStr = if ($have.Count -gt 0) { $have -join ', ' } else { '(none)' }
        throw "Attachment '$target' not found. Resolved email: from=$($item.SenderEmailAddress); subject=$($item.Subject); received=$($item.ReceivedTime). Attachments present: $haveStr. If this is not the email you expected, the EntryID is likely wrong or stale - re-run readInboxEmails to get a current EntryID."
    }
    $savePath = [IO.Path]::Combine('${psEscape(outDir)}', $found.FileName)
    $found.SaveAsFile($savePath)
    $results += [PSCustomObject]@{
        path         = $savePath
        subject      = if ($item.Subject) { $item.Subject } else { '' }
        senderName   = if ($item.SenderName) { $item.SenderName } else { '' }
        senderEmail  = if ($item.SenderEmailAddress) { $item.SenderEmailAddress } else { '' }
        receivedTime = if ($item.ReceivedTime) { $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } else { '' }
    }
}
ConvertTo-Json $results -Depth 3 -Compress
`;
    const raw = await runPowerShell(script, 15000);
    if (!raw || !raw.trim()) throw new Error('Failed to save attachments');
    const parsed = JSON.parse(raw.trim());
    const arr: unknown[] = Array.isArray(parsed) ? parsed : [parsed];
    return arr.map(item => {
        const p = item as Record<string, unknown>;
        return {
            path: String(p.path || ''),
            subject: String(p.subject || ''),
            senderName: String(p.senderName || ''),
            senderEmail: String(p.senderEmail || ''),
            receivedTime: String(p.receivedTime || ''),
        };
    });
}

/**
 * Retrieve all Outlook email accounts (SMTP addresses).
 */
export async function getOutlookAccounts(): Promise<string[]> {
    if (process.platform !== 'win32') return [];
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$accounts = @()
foreach ($a in $outlook.Session.Accounts) {
    $accounts += $a.SmtpAddress
}
$accounts -join '|'
`;
    const result = await runPowerShell(script);
    if (!result) return [];
    return result.split('|').filter(Boolean);
}

// Compile-time proof that this module answers the whole platform contract. Without
// it a signature could drift from OutlookMacService's and only surface as a union
// in the published .d.ts — which is exactly how it drifted before.
const _conformance: OutlookBridge = {
    getOutlookAccounts,
    sendOutlookEmail,
    replyOutlookEmail,
    sendAllDrafts,
    readInboxEmails,
    searchInboxByFilter,
    readSelectedEmail,
    readEmailBody,
    openOutlookEmail,
    listInboxFolders,
    moveOutlookEmails,
    listOutlookDrafts,
    deleteOutlookDrafts,
    deleteOutlookEmails,
    purgeDeletedItems,
    saveEmailAttachment,
    saveEmailAttachmentDetailed,
    saveEmailAttachments,
    cleanUndeliverableEmails,
    collectBouncedRecipients,
    readSentRecipientGroups,
    listOutlookSignatures,
    readOutlookSignatureHtml,
    readTemplateEmails,
    saveTemplateEmail,
};
