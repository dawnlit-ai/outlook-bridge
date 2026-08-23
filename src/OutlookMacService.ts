import { execFile } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { pathToFileURL } from 'url';
import type {
    EmailBodyResult,
    InboxEmail,
    InboxFolderInfo,
    MailFolderRef,
    TemplateFolderResult,
} from './PowerShellService';
// Path parsing and quote splitting are shared with the Windows reader on purpose:
// a folder string and a quoted thread must resolve identically on both platforms,
// and a second implementation here is how the two contracts drift apart.
import { mailFolderRef, splitQuotedOriginal } from './PowerShellService';

// macOS Outlook automation via AppleScript (osascript).
//
// Scope note: this targets **legacy** Outlook for Mac (`defaults read
// com.microsoft.Outlook IsRunningNewOutlook` = 0), which implements enough of the
// AppleScript dictionary to enumerate accounts, read mailboxes, and file drafts.
// "New Outlook" implements only a slice of it: accounts don't enumerate, the
// inbox reports 0 messages, and mutations like `delete` silently no-op. Under New
// Outlook the account lookup below finds nothing and every function here fails
// loudly with "Account '…' not found" rather than quietly doing the wrong thing.
//
// Two dictionary traps worth knowing: `move`/`delete`/`close` report success even
// when they do nothing (always confirm with a count), and a new outgoing message
// lands in Temporary Items — `open` only displays it, so a windowless draft has to
// be moved into the Drafts folder explicitly.

const MAC_NOT_IMPLEMENTED =
    'This feature is not implemented for Outlook on macOS yet.';

/**
 * Resolve the account whose SMTP address matches, or raise. Emitted into a
 * `tell application "Microsoft Outlook"` block; binds `targetAcct`.
 * `every account` errors even in legacy mode, so probe the typed classes.
 * A `whose email address is …` filter can't be used: `email address` is also a
 * class name, and AppleScript resolves it as one ("into type specifier" error).
 */
function accountLookupSnippet(emailAccount: string): string {
    return `    set targetAcct to missing value
    try
        repeat with a in (exchange accounts & imap accounts & pop accounts)
            if (email address of a as string) is "${asEscape(emailAccount)}" then
                set targetAcct to a
                exit repeat
            end if
        end repeat
    end try
    if targetAcct is missing value then error "Account '${asEscape(emailAccount)}' not found"`;
}

/**
 * Outlook for Mac's dictionary term for each well-known root, keyed by the
 * olDefaultFolders id `mailFolderRef` resolves. Verified against the running app:
 * `sent mail` and `junk email` do NOT compile (a bad term fails the whole script
 * at compile time, not at the offending line), so Junk — olDefaultFolders 23 — is
 * deliberately absent and reported as unsupported rather than silently retargeted.
 */
const MAC_ROOT_TERMS: Record<number, string> = {
    6: 'inbox',
    5: 'sent items',
    16: 'drafts',
    3: 'deleted items',
    4: 'outbox',
};

/** The Windows `FolderPath` shape (`\\mailbox\Inbox\Savannah. GA`), built here
 *  rather than in AppleScript — escaping backslashes through a template literal
 *  and then an AppleScript literal is unreadable, and TS already knows the parts. */
function macFolderPath(emailAccount: string, rootLabel: string, segments: string[]): string {
    return ['\\\\' + emailAccount, rootLabel, ...segments].join('\\');
}

/**
 * Emit the AppleScript that resolves `scopeFolder` from `targetAcct`, walking a
 * well-known root down through any further path segments.
 *
 * Folder names are compared with `is`, which is case-insensitive in AppleScript —
 * matching the Windows walk's `-ieq` so 'mobile, al' finds 'Mobile, AL' on both.
 * A missing segment errors by name instead of falling back to the root: silently
 * returning the Inbox for a caller that scoped to one folder hands back the wrong
 * emails under a name that says otherwise.
 */
function mailScopeSnippet(ref: MailFolderRef, folderLabel = ''): string {
    const term = MAC_ROOT_TERMS[ref.rootId];
    if (!term) {
        const asked = folderLabel ? ` (asked for '${folderLabel}')` : '';
        throw new Error(
            `Outlook for Mac has no '${ref.rootLabel}' folder in its AppleScript dictionary${asked}. `
            + `Readable roots: Inbox, Sent Items, Drafts, Deleted Items, Outbox.`,
        );
    }
    const walk = ref.segments.map(seg => `
    set foundFolder to missing value
    repeat with sf in (mail folders of scopeFolder)
        if (name of sf as string) is "${asEscape(seg)}" then
            set foundFolder to sf
            exit repeat
        end if
    end repeat
    if foundFolder is missing value then error "Folder '${asEscape(seg)}' not found under '${asEscape(ref.rootLabel)}'"
    set scopeFolder to foundFolder`).join('');
    return `    set scopeFolder to ${term} of targetAcct${walk}`;
}

/** Zero-padded `yyyy-MM-dd HH:mm`, matching the Windows receivedTime contract. */
const DATE_HANDLERS = `
on pad2(n)
    set s to (n as integer) as string
    if (length of s) < 2 then set s to "0" & s
    return s
end pad2

on isoDate(d)
    return (year of d as string) & "-" & pad2(month of d as integer) & "-" & pad2(day of d) & " " & pad2(hours of d) & ":" & pad2(minutes of d)
end isoDate
`;

// Control characters as separators: subjects and bodies contain tabs, newlines and
// commas, but never these.
const FIELD_SEP = '\u001f';
const RECORD_SEP = '\u001e';
const LIST_SEP = '\u001d';

/** Run an AppleScript via a temp file (avoids arg-length and quoting limits). */
function runOsaScript(script: string, timeout?: number): Promise<string> {
    const scriptFile = path.join(os.tmpdir(), `sla-osa-${Date.now()}-${Math.random().toString(36).slice(2)}.applescript`);
    fs.writeFileSync(scriptFile, script, 'utf-8');
    return new Promise((resolve, reject) => {
        execFile(
            'osascript',
            [scriptFile],
            { maxBuffer: 8 * 1024 * 1024, timeout: timeout ?? 60000 },
            (error, stdout, stderr) => {
                try {
                    fs.unlinkSync(scriptFile);
                } catch { /* ignore */
                }
                if (error) {
                    // Running from a file, osascript prefixes the script path:
                    // "/tmp/sla-osa-1.applescript:12:34: execution error: Microsoft
                    // Outlook got an error: … (-1728)". Keep just the human part —
                    // these strings reach the user in per-email failure lists.
                    const msg = (stderr || error.message)
                        .replace(/^(?:.*?:)?\d+:\d+:\s*execution error:\s*/m, '')
                        .trim();
                    reject(new Error(msg));
                } else {
                    resolve(stdout.replace(/\n$/, ''));
                }
            }
        );
    });
}

/** Escape a string for embedding inside an AppleScript double-quoted literal. */
function asEscape(s: string): string {
    return s
        .replace(/\\/g, '\\\\')
        .replace(/"/g, '\\"')
        .replace(/\r\n|\r|\n/g, '\\n');
}

// Account enumeration only works when Outlook runs in legacy mode; New Outlook
// errors on `every account` and returns empty lists for the typed classes, so
// each class is probed independently and failures are ignored.
const LIST_ACCOUNTS_SNIPPET = `
set acctList to {}
tell application "Microsoft Outlook"
    repeat with acctClass in {"exchange", "imap", "pop"}
        try
            if (acctClass as string) is "exchange" then
                set accts to exchange accounts
            else if (acctClass as string) is "imap" then
                set accts to imap accounts
            else
                set accts to pop accounts
            end if
            repeat with a in accts
                try
                    set end of acctList to (email address of a as string)
                end try
            end repeat
        end try
    end repeat
end tell
`;

/** SMTP addresses of Outlook accounts. Empty under New Outlook (not exposed). Used by the UI's account picker. */
export async function getOutlookAccounts(): Promise<string[]> {
    const script = `${LIST_ACCOUNTS_SNIPPET}
set AppleScript's text item delimiters to linefeed
return acctList as string`;
    const raw = await runOsaScript(script, 15000);
    return raw.split('\n').map(s => s.trim()).filter(Boolean);
}

interface OutlookEmailParams {
    emailAccount: string;
    to: string;
    cc?: string;
    subject: string;
    htmlBody: string;
    attachmentPath?: string;
    sendImmediately: boolean;
    /**
     * Draft mode only. True opens an Outlook compose window per email; false
     * files the draft into the Drafts folder with no window — the only workable
     * option for batches. Defaults to true, matching the single-email callers.
     */
    openDraftWindow?: boolean;
}

/**
 * Create an Outlook email as a draft window (or send it) via AppleScript.
 * Mirrors the Windows COM contract from PowerShellService.sendOutlookEmail.
 */
export async function sendOutlookEmail(params: OutlookEmailParams): Promise<void> {
    // Same safety property as Windows: never silently send from the wrong mailbox.
    // The account lookup raises when the address doesn't resolve, so there's no
    // pre-flight getOutlookAccounts() round-trip — it cost an extra osascript
    // launch on every email in a batch.
    const splitAddresses = (s: string): string[] =>
        s.split(/[,;]+/).map(x => x.trim()).filter(Boolean);

    const recipientLines = [
        ...splitAddresses(params.to).map(addr =>
            `    make new to recipient at newMsg with properties {email address:{address:"${asEscape(addr)}"}}`),
        ...splitAddresses(params.cc || '').map(addr =>
            `    make new cc recipient at newMsg with properties {email address:{address:"${asEscape(addr)}"}}`),
    ].join('\n');

    const attachLine = params.attachmentPath
        ? `    make new attachment at newMsg with properties {file:POSIX file "${asEscape(params.attachmentPath)}"}`
        : '';

    // A new outgoing message lands in Temporary Items, not Drafts — `open` is what
    // surfaces it, so a windowless draft has to be moved into the Drafts folder
    // explicitly. (`save` is no help: it demands a file destination, not a folder.)
    // `move` reports success even where it silently does nothing, so confirm the
    // folder actually grew rather than risk reporting a draft that doesn't exist.
    const openDraftWindow = params.openDraftWindow !== false;
    const fileDraft = !params.sendImmediately && !openDraftWindow;

    const draftsPrelude = fileDraft
        ? `    set draftsFolder to drafts of targetAcct
    set draftsBefore to count of messages of draftsFolder`
        : '';

    const actionLine = params.sendImmediately
        ? '    send newMsg'
        : openDraftWindow
            ? '    open newMsg'
            : `    move newMsg to draftsFolder
    if (count of messages of draftsFolder) is not greater than draftsBefore then error "Outlook did not file the draft in the Drafts folder."`;

    const script = `
tell application "Microsoft Outlook"
${accountLookupSnippet(params.emailAccount)}
${draftsPrelude}
    set newMsg to make new outgoing message with properties {subject:"${asEscape(params.subject)}", content:"${asEscape(params.htmlBody)}"}
${recipientLines}
    set account of newMsg to targetAcct
${attachLine}
${actionLine}
end tell`;
    await runOsaScript(script, 120000);
}

/** Names of the user's Outlook signatures. */
export async function listOutlookSignatures(): Promise<string[]> {
    const script = `
set sigNames to {}
tell application "Microsoft Outlook"
    try
        repeat with s in signatures
            set end of sigNames to (name of s as string)
        end repeat
    end try
end tell
set AppleScript's text item delimiters to linefeed
return sigNames as string`;
    const raw = await runOsaScript(script, 15000);
    return raw.split('\n').map(s => s.trim()).filter(Boolean).sort((a, b) => a.localeCompare(b));
}

/**
 * Read a named signature's HTML. Unlike Windows (files on disk), Outlook for
 * Mac serves signature content directly through AppleScript, images included
 * (inline data or remote refs). Returns '' if the signature can't be found.
 */
export async function readOutlookSignatureHtml(name: string): Promise<string> {
    const script = `
tell application "Microsoft Outlook"
    try
        return content of (first signature whose name is "${asEscape(name)}")
    on error
        return ""
    end try
end tell`;
    return runOsaScript(script, 15000);
}

/**
 * In-app template editor. Windows edits templates through an Outlook draft and
 * polls until the inspector closes; New Outlook for Mac invalidates the
 * scripting object on `open`, making that flow impossible. Instead, open a
 * small always-WYSIWYG BrowserWindow with the template in a contentEditable
 * region — Save (button or Cmd+S) resolves with the edited HTML, closing
 * without saving rejects, matching the Windows contract.
 */
export async function editEmailTemplate(label: string, currentHtml: string): Promise<string> {
    // Lazy require: the MCP bundle stubs 'electron' with {}, and MCP tools never
    // call this — guard so a stray call fails with a clear message, not a crash.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const electron = require('electron') as typeof import('electron');
    if (!electron.BrowserWindow) {
        throw new Error('Template editing requires the desktop app.');
    }

    const SAVE_TITLE = 'sla-template-save';
    const pageHtml = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>${label.replace(/[<>&]/g, '')}</title>
</head>
<body style="margin:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f8fafc;display:flex;flex-direction:column;height:100vh;">
<div style="display:flex;align-items:center;gap:12px;padding:10px 16px;background:#1e293b;color:#f1f5f9;flex:none;">
    <div style="font-size:14px;font-weight:600;flex:1;">${label.replace(/[<>&]/g, '')}</div>
    <button id="sla-cancel" style="padding:6px 14px;border:1px solid #475569;border-radius:6px;background:transparent;color:#cbd5e1;font-size:13px;cursor:pointer;">Cancel</button>
    <button id="sla-save" style="padding:6px 18px;border:none;border-radius:6px;background:#2563eb;color:#fff;font-size:13px;font-weight:600;cursor:pointer;">Save (&#8984;S)</button>
</div>
<div id="sla-editor" contenteditable="true" style="flex:1;overflow:auto;margin:16px;padding:20px;background:#fff;border:1px solid #cbd5e1;border-radius:8px;outline:none;font-size:14px;color:#0f172a;">${currentHtml}</div>
<script>
document.getElementById('sla-save').addEventListener('click', () => { document.title = '${SAVE_TITLE}'; });
document.getElementById('sla-cancel').addEventListener('click', () => { window.close(); });
document.addEventListener('keydown', (e) => {
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        document.title = '${SAVE_TITLE}';
    }
});
</script>
</body>
</html>`;

    const pageFile = path.join(os.tmpdir(), `sla-template-editor-${Date.now()}.html`);
    fs.writeFileSync(pageFile, pageHtml, 'utf-8');

    try {
        return await new Promise<string>((resolve, reject) => {
            const win = new electron.BrowserWindow({
                width: 720,
                height: 640,
                title: label,
                webPreferences: { nodeIntegration: false, contextIsolation: true },
            });
            let settled = false;
            win.webContents.on('page-title-updated', async (e, title) => {
                e.preventDefault();
                if (title !== SAVE_TITLE || settled) return;
                try {
                    const html: string = await win.webContents.executeJavaScript(
                        "document.getElementById('sla-editor').innerHTML"
                    );
                    settled = true;
                    resolve(html);
                } catch (err) {
                    settled = true;
                    reject(err instanceof Error ? err : new Error(String(err)));
                } finally {
                    win.destroy();
                }
            });
            win.on('closed', () => {
                if (!settled) {
                    settled = true;
                    reject(new Error('No template saved. Click Save (Cmd+S) before closing.'));
                }
            });
            win.loadURL(pathToFileURL(pageFile).href);
        });
    } finally {
        try {
            fs.unlinkSync(pageFile);
        } catch { /* ignore */
        }
    }
}

/**
 * Read recent inbox messages for the given account, newest first.
 *
 * `entryId` here is Outlook for Mac's small integer message id (e.g. "779"), not
 * the MAPI EntryID string Windows returns. The two are not interchangeable, so
 * don't persist one and look it up on the other platform.
 */
export async function readInboxEmails(
    emailAccount: string,
    daysBack: number = 60,
    limit: number = 50,
    folder?: string,
): Promise<InboxEmail[]> {
    const days = Math.max(0, Math.floor(daysBack));
    const cap = Math.max(0, Math.floor(limit));
    if (cap === 0) return [];
    const ref = folder ? mailFolderRef(folder) : { rootId: 6, rootLabel: 'Inbox', segments: [] };
    // A folder argument that trims away to nothing ("\\", "  ") would otherwise
    // read the Inbox root and look like it had scoped — the exact silent
    // mis-scoping this parameter exists to prevent. Mirrors the Windows check.
    if (folder && folder.trim() && ref.segments.length === 0 && ref.rootId === 6
        && folder.trim().toLowerCase() !== 'inbox') {
        throw new Error(`Folder '${folder}' does not name a folder under the Inbox.`);
    }
    const resolveScope = mailScopeSnippet(ref, folder || '');
    const folderPath = macFolderPath(emailAccount, ref.rootLabel, ref.segments);
    // Which timestamp the folder's items actually carry: outgoing mail has no
    // `time received`, so indexing Sent Items on it yields an empty set that looks
    // exactly like an empty folder. The try/on error keeps a folder holding a mix
    // (a draft filed into Inbox) from failing the whole bulk read.
    const isOutgoing = ref.rootId === 5 || ref.rootId === 4 || ref.rootId === 16;
    const dateProp = isOutgoing ? 'time sent' : 'time received';

    // Pass 1 — index the folder with two bulk property reads (one Apple event
    // each). Reading per-message in a loop costs an event per property and is
    // unusably slow on a real inbox. Folder order isn't documented, so sort here
    // instead of trusting Outlook to hand back newest-first.
    const indexScript = `${DATE_HANDLERS}
tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
${resolveScope}
    set cutoff to (current date) - (${days} * days)
    set idList to id of every message of scopeFolder
    try
        set timeList to ${dateProp} of every message of scopeFolder
    on error
        set timeList to time received of every message of scopeFolder
    end try
    set out to ""
    repeat with i from 1 to (count of idList)
        set d to item i of timeList
        if d is greater than or equal to cutoff then
            set out to out & (item i of idList as string) & tab & my isoDate(d) & linefeed
        end if
    end repeat
    return out
end tell`;

    const index = (await runOsaScript(indexScript, 60000))
        .split('\n')
        .map(line => line.trim())
        .filter(Boolean)
        .map(line => {
            const [id, receivedTime] = line.split('\t');
            return { id, receivedTime: receivedTime || '' };
        });
    // 'yyyy-MM-dd HH:mm' is lexicographically ordered, so plain string compare sorts it.
    index.sort((a, b) => b.receivedTime.localeCompare(a.receivedTime));
    const chosen = index.slice(0, cap);
    if (chosen.length === 0) return [];

    // Pass 2 — the expensive reads (body, attachments) only for messages we keep.
    // `sender` yields a record, and `address of sender of m` fails to coerce, so
    // bind the record first. `name` is absent when there's no display name.
    const detailScript = `
tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
${resolveScope}
    set wanted to {${chosen.map(c => c.id).join(', ')}}
    set out to ""
    repeat with k from 1 to (count of wanted)
        set theId to item k of wanted
        set m to message id theId of scopeFolder
        set subj to ""
        try
            set subj to (subject of m) as string
        end try
        set sndName to ""
        set sndAddr to ""
${isOutgoing ? `        -- Outgoing mail has no meaningful sender of its own, so these carry the
        -- first recipient instead — the Windows reader's contract for Sent Items.
        try
            set rcps to to recipients of m
            if (count of rcps) > 0 then
                set ea to email address of (item 1 of rcps)
                try
                    set sndAddr to (address of ea) as string
                end try
                try
                    set sndName to (name of ea) as string
                end try
            end if
        end try` : `        try
            set snd to sender of m
            try
                set sndAddr to (address of snd) as string
            end try
            try
                set sndName to (name of snd) as string
            end try
        end try`}
        set bodyText to ""
        try
            set bodyText to (plain text content of m) as string
        end try
        if (length of bodyText) > 600 then set bodyText to text 1 thru 600 of bodyText
        set attNames to {}
        try
            set attNames to name of every attachment of m
        end try
        set AppleScript's text item delimiters to (character id 29)
        set attJoined to attNames as string
        set AppleScript's text item delimiters to ""
        set out to out & (theId as string) & (character id 31) & subj & (character id 31) & sndName & (character id 31) & sndAddr & (character id 31) & bodyText & (character id 31) & attJoined & (character id 30)
    end repeat
    return out
end tell`;

    const byId = new Map(chosen.map(c => [c.id, c.receivedTime]));
    return (await runOsaScript(detailScript, 120000))
        .split(RECORD_SEP)
        .filter(Boolean)
        .map(record => {
            const [id, subject, senderName, senderEmail, bodyPreview, attJoined] = record.split(FIELD_SEP);
            const attachmentNames = (attJoined || '').split(LIST_SEP).filter(Boolean);
            return {
                entryId: id || '',
                storeId: '', // macOS AppleScript has no StoreID equivalent
                subject: (subject || '').trim(),
                senderName: senderName || '',
                senderEmail: senderEmail || '',
                receivedTime: byId.get(id) || '',
                bodyPreview: bodyPreview || '',
                attachmentNames,
                attachmentCount: attachmentNames.length,
                folderPath,
            };
        });
}

// ── Mailbox features not yet ported to macOS ─────────────────────────────
// These mirror the Windows contract: list readers return empty, explicit
// single-item actions throw. Legacy Outlook exposes what they need (messages,
// attachments, `reply to`), so they are implementable — just not implemented.

interface PreAlertEntry {
    hblPath: string;
    entryId: string;
    storeId: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
    body: string;
    attachmentPaths: string[];
}

export async function readAllPreAlerts(
    _emailAccount: string,
    _daysBack = 0,
): Promise<PreAlertEntry[]> {
    return [];
}

export async function readSelectedPreAlert(): Promise<PreAlertEntry> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function openOutlookEmail(_entryId: string): Promise<void> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

/**
 * Read one email's full plain-text body by message id.
 *
 * `entryId` is Outlook for Mac's small integer message id (see readInboxEmails),
 * not a Windows MAPI EntryID — a digit check rejects the latter up front rather
 * than letting it fail as a confusing "not found".
 *
 * `storeId` is accepted for signature parity and ignored: macOS has no StoreID,
 * and `message id N` resolves against the application rather than one folder, so
 * the message is found wherever it currently sits — including a lane subfolder.
 *
 * Uses `plain text content`, matching the Windows reader's use of `.Body`: the
 * same tradeoff applies, so an HTML table's rows flatten and a tabular rate is
 * better read from its attachment.
 */
export async function readEmailBody(
    entryId: string,
    _storeId?: string,
    maxChars: number = 8000,
    includeQuoted: boolean = false,
): Promise<EmailBodyResult> {
    const id = String(entryId ?? '').trim();
    if (!/^\d+$/.test(id)) {
        throw new Error(
            `'${entryId}' is not an Outlook for Mac message id. Mac ids are small integers `
            + `(e.g. "1263") returned by list_outlook_inbox on this machine; a Windows MAPI `
            + `EntryID cannot be resolved here.`,
        );
    }
    const script = `${DATE_HANDLERS}
tell application "Microsoft Outlook"
    set m to missing value
    try
        set m to message id ${id}
    end try
    if m is missing value then error "Email not found for message id '${id}'"
    set subj to ""
    try
        set subj to (subject of m) as string
    end try
    set sndName to ""
    set sndAddr to ""
    try
        set snd to sender of m
        try
            set sndAddr to (address of snd) as string
        end try
        try
            set sndName to (name of snd) as string
        end try
    end try
    set recvd to ""
    try
        set recvd to my isoDate(time received of m)
    on error
        try
            set recvd to my isoDate(time sent of m)
        end try
    end try
    set attNames to {}
    try
        set attNames to name of every attachment of m
    end try
    set AppleScript's text item delimiters to (character id 29)
    set attJoined to attNames as string
    set AppleScript's text item delimiters to ""
    set bodyText to ""
    try
        set bodyText to (plain text content of m) as string
    end try
    return (id of m as string) & (character id 31) & subj & (character id 31) & sndName & (character id 31) & sndAddr & (character id 31) & recvd & (character id 31) & attJoined & (character id 31) & bodyText
end tell`;

    // The body is emitted LAST so a stray separator in earlier fields can't shift it.
    const parts = (await runOsaScript(script, 60000)).split(FIELD_SEP);
    const [rid, subject, senderName, senderEmail, receivedTime, attJoined] = parts;
    const full = parts.slice(6).join(FIELD_SEP);
    const attachmentNames = (attJoined || '').split(LIST_SEP).filter(Boolean);
    const { body, quoted, separator } = splitQuotedOriginal(full);
    // The quoted thread is context, never the priced content, so it is capped
    // harder than the reply itself — matching the Windows reader.
    const quotedCap = Math.min(maxChars, 4000);
    return {
        entryId: rid || id,
        subject: (subject || '').trim(),
        senderName: senderName || '',
        senderEmail: senderEmail || '',
        receivedTime: receivedTime || '',
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

export async function sendReceivedConfirmation(_emailAccount: string, _entryId: string, _customHtml?: string): Promise<void> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function sendAllDrafts(_emailAccount: string): Promise<{
    sent: number;
    failed: { subject: string; error: string }[]
}> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function saveEmailAttachment(_entryId: string, _fileName: string, _storeId?: string): Promise<string> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function saveEmailAttachmentDetailed(_entryId: string, _fileName: string, _storeId?: string): Promise<{
    path: string;
    subject: string;
    senderName: string;
    senderEmail: string;
    receivedTime: string;
}> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function cleanUndeliverableEmails(
    _emailAccount: string,
    _daysBack?: number,
    _dryRun?: boolean,
): Promise<{
    account: string;
    scannedDays: number;
    dryRun: boolean;
    matchedCount: number;
    deletedCount: number;
    matched: unknown[];
    failed: { subject: string; error: string }[];
}> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function collectBouncedRecipients(
    _emailAccount: string,
    _daysBack?: number,
    _scanDeleted?: boolean,
): Promise<string[]> {
    return [];
}

export async function readSentRecipientGroups(
    _emailAccount: string,
    _daysBack?: number,
    _limit?: number,
): Promise<{ entryId: string; subject: string; sentOn: string; recipients: string[] }[]> {
    return [];
}

/**
 * List the folders under an account's Inbox with their item counts.
 *
 * The recursive walk joins path segments with the LIST separator rather than a
 * backslash: a literal backslash would have to survive a JS template literal and
 * then an AppleScript string literal, and the `folderPath` shape is assembled in
 * TypeScript anyway (see macFolderPath).
 */
export async function listInboxFolders(
    emailAccount: string,
    maxDepth = 2,
): Promise<InboxFolderInfo[]> {
    const depth = Math.max(1, Math.min(4, Math.floor(maxDepth)));
    const script = `
on walkFolders(theFolder, level, maxLevel, prefix)
    set sep to (character id 29)
    set fs to (character id 31)
    set rs to (character id 30)
    set out to ""
    tell application "Microsoft Outlook"
        set subs to mail folders of theFolder
    end tell
    repeat with f in subs
        set nm to ""
        set cnt to 0
        tell application "Microsoft Outlook"
            try
                set nm to (name of f) as string
            end try
            try
                set cnt to (count of messages of f)
            end try
        end tell
        set thisPath to prefix & nm
        set out to out & nm & fs & thisPath & fs & (cnt as string) & fs & (level as string) & rs
        if level < maxLevel then
            set out to out & (my walkFolders(f, level + 1, maxLevel, thisPath & sep))
        end if
    end repeat
    return out
end walkFolders

tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
    set inb to inbox of targetAcct
end tell
return my walkFolders(inb, 1, ${depth}, "")`;

    return (await runOsaScript(script, 120000))
        .split(RECORD_SEP)
        .filter(Boolean)
        .map(record => {
            const [name, rawPath, itemCount, level] = record.split(FIELD_SEP);
            const segments = (rawPath || '').split(LIST_SEP).filter(Boolean);
            return {
                name: name || '',
                folderPath: macFolderPath(emailAccount, 'Inbox', segments),
                itemCount: Number.parseInt(itemCount || '0', 10) || 0,
                depth: Number.parseInt(level || '1', 10) || 1,
            };
        });
}

export async function moveOutlookEmails(
    _emailAccount: string,
    _entryIds: string[],
    _folderName: string,
    _createIfMissing?: boolean,
): Promise<{
    folderPath: string;
    folderCreated: boolean;
    moved: number;
    failed: { entryId: string; error: string }[]
}> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function listOutlookDrafts(
    emailAccount: string,
    _limit?: number,
    _previewChars?: number,
): Promise<{
    account: string;
    foldersScanned: string[];
    count: number;
    truncated: boolean;
    drafts: never[];
}> {
    return { account: emailAccount, foldersScanned: [], count: 0, truncated: false, drafts: [] };
}

export async function deleteOutlookDrafts(
    _emailAccount: string,
    _entryIds: string[],
): Promise<{ deleted: number; failed: { entryId: string; error: string }[] }> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function deleteOutlookEmails(
    _emailAccount: string,
    _entryIds: string[],
    _options?: { allowProtected?: boolean; dryRun?: boolean },
): Promise<never> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function purgeDeletedItems(
    _emailAccount: string,
    _olderThanDays?: number,
    _dryRun?: boolean,
): Promise<never> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function replyOutlookEmail(_params: {
    emailAccount: string;
    entryId: string;
    storeId?: string;
    htmlBody?: string;
    templateSubject?: string;
    templateFolder?: string;
    templateSection?: string;
    templatePlaceholders?: Record<string, string>;
    sendImmediately?: boolean;
    openDraftWindow?: boolean;
}): Promise<{ to: string; subject: string; repliedToSender: string }> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function readTemplateEmails(
    _emailAccount: string,
    _folderName?: string,
    _limit?: number,
    _includeBody?: boolean,
    _subject?: string,
): Promise<TemplateFolderResult> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function saveTemplateEmail(
    _emailAccount: string,
    _subject: string,
    _htmlBody: string,
    _folderName?: string,
): Promise<{ folderPath: string; folderCreated: boolean }> {
    throw new Error(MAC_NOT_IMPLEMENTED);
}

export async function exportSheetAsPdf(_excelPath: string, _sheetName: string, _pdfPath: string): Promise<void> {
    throw new Error('Excel PDF export is not yet supported on macOS.');
}
