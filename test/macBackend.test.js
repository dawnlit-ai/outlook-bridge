// The macOS backend's two halves that need no Outlook: what goes INTO a script,
// and what comes back OUT of one.
//
// macScripts.test.js compiles every generated script, which needs Outlook
// installed and so skips in CI — leaving the macOS backend with no coverage
// there at all, while Windows gets its parse and injection checks on every run.
// These are the macOS counterparts that do run anywhere: the runner is stubbed,
// so no Outlook, no mailbox and (except where noted) no osascript is involved.
//
// Two properties are checked. Caller text must never escape its AppleScript
// literal — the counterpart of the PowerShell expandable-string check, since a
// value that breaks out of a literal becomes code either way. And a script's
// framed output must decode into the same result shape Windows returns, which
// is the half compiling a script can say nothing about.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');

const macRun = require('../dist/mac/run.js');
const {macBackend} = require('../dist/mac/index.js');
const {createOperations} = require('../dist/bridge.js');
const {captureScripts, tempDir, tempFileWith} = require('./support.js');

const {FIELD_SEP: FS, RECORD_SEP: RS, LIST_SEP: LS} = macRun;
const ACCOUNT = 'someone@example.com';
const outlook = createOperations(macBackend);

/** One framed record, the way the scripts emit them. */
function row(...fields) {
    return fields.join(FS) + RS;
}

/** A list field written by sanitizePositionalList, as an email's attachment names are. */
function names(...items) {
    return items.map(item => item + LS).join('');
}

/** Drive an operation with canned runner output, returning its result. */
async function decode(operation, responses) {
    let value;
    const {error} = await captureScripts(macRun, ['runOsaScript'], async () => {
        value = await operation();
    }, responses);
    if (error) throw error;
    return value;
}

// ── Caller text never escapes its AppleScript literal ────────────────────

/**
 * Text built to break out of a double-quoted AppleScript literal and run as
 * code. `OBPROBE` is the sentinel: plain letters, so escaping leaves it intact
 * and it can be found again in the generated script.
 */
const PROBE = 'OBPROBE" & (do shell script "echo pwned") & "tail \\ back "dq" ’curly’';

/**
 * Split a script into its double-quoted literals and the code between them,
 * honouring `\\` escapes and `--` comments (a comment may quote the very
 * patterns this guards against). Returns the code half and the decoded value of
 * every literal.
 */
function partition(script) {
    let code = '';
    const literals = [];
    let inString = false;
    let inComment = false;
    let current = '';
    for (let i = 0; i < script.length; i++) {
        const ch = script[i];
        if (inComment) {
            if (ch === '\n') inComment = false;
            continue;
        }
        if (inString) {
            if (ch === '\\') {
                const next = script[++i];
                current += next === 'n' ? '\n' : next === 't' ? '\t' : next === 'r' ? '\r' : next;
            } else if (ch === '"') {
                literals.push(current);
                current = '';
                inString = false;
            } else {
                current += ch;
            }
            continue;
        }
        if (ch === '"') {
            inString = true;
        } else if (ch === '-' && script[i + 1] === '-') {
            inComment = true;
            i++;
        } else {
            code += ch;
        }
    }
    return {code, literals};
}

test('partition finds a literal, its escapes, and the code around it', () => {
    const {code, literals} = partition('set x to "a\\"b\\\\c" -- "not code"\nreturn x');
    assert.deepEqual(literals, ['a"b\\c']);
    assert.match(code, /set x to/);
    assert.doesNotMatch(code, /not code/, 'a comment is neither code nor a literal');
});

/** Every operation, driven with the probe wherever caller text is accepted. */
function probeCases() {
    const attachment = tempFileWith('report.pdf');
    const destDir = tempDir('ob-mac-probe-');
    return {
        sendOutlookEmail: {
            run: () => outlook.sendOutlookEmail({
                emailAccount: PROBE,
                to: `"${PROBE.replace(/"/g, '')}" <a@example.com>`,
                cc: 'c@example.com',
                subject: PROBE,
                htmlBody: `<p>${PROBE}</p>`,
                attachments: [attachment],
                openDraftWindow: false,
            }),
        },
        replyOutlookEmail: {
            run: () => outlook.replyOutlookEmail({
                emailAccount: PROBE, entryId: '1263', htmlBody: `<p>${PROBE}</p>`, replyAll: true,
            }),
        },
        readInboxEmails: {
            run: () => outlook.readInboxEmails(PROBE, {folder: `Inbox\\${PROBE}`}),
            responses: ['1263\t2026-08-01 09:30\n'],
        },
        searchInboxByFilter: {
            run: () => outlook.searchInboxByFilter(PROBE, {
                subjectLike: `*${PROBE}*`, includeFolders: [PROBE], excludeFolders: [PROBE],
            }),
            responses: [row('INBOX', '1263', PROBE, '2026-08-01 09:30')],
        },
        listInboxFolders: {run: () => outlook.listInboxFolders(PROBE)},
        listMailboxFolders: {run: () => outlook.listInboxFolders(PROBE, {from: 'mailbox'})},
        moveOutlookEmails: {
            run: () => outlook.moveOutlookEmails(PROBE, ['1263'], `Clients\\${PROBE}`, {createIfMissing: true}),
            responses: [row('1', 'true')],
        },
        listOutlookDrafts: {run: () => outlook.listOutlookDrafts(PROBE)},
        sendDrafts: {run: () => outlook.sendDrafts(PROBE, ['1263']), responses: [row('1')]},
        deleteOutlookDrafts: {run: () => outlook.deleteOutlookDrafts(PROBE, ['1263']), responses: [row('1')]},
        sendAllDrafts: {run: () => outlook.sendAllDrafts(PROBE)},
        deleteOutlookEmails: {run: () => outlook.deleteOutlookEmails(PROBE, ['1263'])},
        purgeDeletedItems: {
            run: () => outlook.purgeDeletedItems(PROBE, {olderThanDays: 30}),
            responses: [row('1263', 'false'), row('1', '0')],
        },
        saveEmailAttachments: {
            run: () => outlook.saveEmailAttachments('1263', [PROBE], {destDir}),
            responses: [row('1263', 'Subject', 'Name', 'a@example.com', '2026-08-01 09:30', names(PROBE))],
        },
        cleanUndeliverableEmails: {
            run: () => outlook.cleanUndeliverableEmails(PROBE, {dryRun: false}),
            responses: [
                row('1263', 'Undeliverable: x', 'Mail Delivery Subsystem', 'daemon@x.com', '2026-08-01 09:30'),
                row('1263', 'failed for bob@example.com'),
                row('1'),
            ],
        },
        collectBouncedRecipients: {run: () => outlook.collectBouncedRecipients(PROBE)},
        readSentRecipientGroups: {
            run: () => outlook.readSentRecipientGroups(PROBE),
            responses: [row('1263', PROBE, '2026-08-01 09:30')],
        },
        readOutlookSignatureHtml: {run: () => outlook.readOutlookSignatureHtml(PROBE)},
        readTemplateEmails: {
            run: () => outlook.readTemplateEmails(PROBE, {folder: PROBE, subject: PROBE}),
            responses: [row('1', 'Templates') + row('1263', PROBE, '2026-08-01 09:30')],
        },
        saveTemplateEmail: {
            run: () => outlook.saveTemplateEmail(PROBE, {subject: PROBE, folder: PROBE, htmlBody: `<p>${PROBE}</p>`}),
            responses: [row('Templates', 'false')],
        },
        editEmailTemplate: {
            run: () => outlook.editEmailTemplate(PROBE, `<p>${PROBE}</p>`),
            responses: ['<p>edited</p>'],
        },
    };
}

test('no generated script lets caller text out of its AppleScript literal', async () => {
    for (const [name, {run, responses}] of Object.entries(probeCases())) {
        const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], run, responses);
        if (error) throw error;
        assert.ok(scripts.length > 0, `${name} generated no script`);
        for (const script of scripts) {
            const {code, literals} = partition(macRun.AS_HANDLERS + script);
            // The sentinel is plain letters, so it survives escaping verbatim:
            // finding it in the code half means it left its literal.
            assert.doesNotMatch(code, /OBPROBE/, `${name}: caller text reached the code half of the script`);
            if (script.includes('OBPROBE')) {
                // Inside a literal is the whole requirement. Not every operation
                // passes the text through untouched — an attachment's name is
                // sanitized into a safe file name first — so what is asserted is
                // containment, not that the bytes survived. `asString` keeping
                // the text exact is checked on its own, under osascript, below.
                assert.ok(
                    literals.some(value => value.includes('OBPROBE')),
                    `${name}: the probe is in the script but in no literal`,
                );
            }
        }
    }
});

test('asEscape closes every way out of a literal', () => {
    assert.equal(macRun.asEscape('a"b'), 'a\\"b');
    assert.equal(macRun.asEscape('a\\b'), 'a\\\\b');
    // A lone backslash before a quote must not leave the quote unescaped.
    assert.equal(macRun.asEscape('a\\"b'), 'a\\\\\\"b');
    assert.equal(macRun.asEscape('a\r\nb\rc\nd'), 'a\\nb\\nc\\nd');
});

test('asInt and asIdList refuse anything that is not a whole number', () => {
    assert.equal(macRun.asInt(42), '42');
    for (const bad of [1.5, NaN, Infinity]) {
        assert.throws(() => macRun.asInt(bad), error => error.code === 'INVALID_REQUEST');
    }
    assert.equal(macRun.asIdList(['1', '2']), '{1, 2}');
    assert.throws(() => macRun.asIdList(['1; do shell script "x"']), error => error.code === 'INVALID_REQUEST');
});

// osascript ships with macOS, so this runs on any Mac — Outlook is not involved.
const needsOsascript = process.platform === 'darwin' ? false : 'needs macOS osascript';

test('the probe round-trips through asString exactly, under osascript', {skip: needsOsascript}, () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-osa-esc-'));
    const file = path.join(dir, 'probe.applescript');
    try {
        // '& (2 + 2)' is the tell: if the probe escaped its literal, the script
        // would return a 4 where the text should be — or fail to compile.
        fs.writeFileSync(file, `${macRun.AS_HANDLERS}\nreturn ${macRun.asString(PROBE)} & (2 + 2)`, 'utf8');
        const out = execFileSync('osascript', [file], {encoding: 'utf8'}).replace(/\n$/, '');
        // Coming back byte for byte IS the proof: had the probe escaped, the
        // `do shell script` it carries would have run and replaced itself with
        // its own output, and the trailing `& (2 + 2)` would have been consumed
        // by the injected expression rather than appending a 4.
        assert.equal(out, `${PROBE}4`);
    } finally {
        fs.rmSync(dir, {recursive: true, force: true});
    }
});

test('a script file is written as UTF-8 that osascript reads back unchanged', {skip: needsOsascript}, () => {
    // The Windows runner needs a byte-order mark for PowerShell 5.1 to read its
    // file as UTF-8; this is the check that osascript needs no such help, and
    // that a subject in any script is safe to pass through a file.
    const text = 'Ångström — ünïcode “curly” 中文 café';
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-osa-utf8-'));
    const file = path.join(dir, 'utf8.applescript');
    try {
        fs.writeFileSync(file, `return ${macRun.asString(text)}`, 'utf8');
        assert.equal(execFileSync('osascript', [file], {encoding: 'utf8'}).replace(/\n$/, ''), text);
    } finally {
        fs.rmSync(dir, {recursive: true, force: true});
    }
});

// ── A script's output decodes into the shape Windows returns ─────────────

test('readInboxEmails pairs the index with the details and keeps the folder path', async () => {
    const mail = await decode(() => outlook.readInboxEmails(ACCOUNT, {folder: 'Inbox\\Clients'}), [
        '1263\t2026-08-01 09:30\n1264\t2026-08-02 11:00\n',
        row('1264', ' Later ', 'Al Roe', 'al@x.com', 'preview B', '') +
        row('1263', ' Subject A ', 'Jo Doe', 'jo@x.com', 'preview A', names('a.pdf', 'b.pdf')),
    ]);
    // Newest first, whatever order the index and the details came back in.
    assert.deepEqual(mail.map(m => m.entryId), ['1264', '1263']);
    assert.deepEqual(mail[1], {
        entryId: '1263',
        storeId: '',
        subject: 'Subject A',
        senderName: 'Jo Doe',
        senderEmail: 'jo@x.com',
        receivedTime: '2026-08-01 09:30',
        bodyPreview: 'preview A',
        attachmentNames: ['a.pdf', 'b.pdf'],
        attachmentCount: 2,
        folderPath: '\\\\someone@example.com\\Inbox\\Clients',
    });
});

test('readInboxEmails takes the newest, not the first, when the limit bites', async () => {
    const mail = await decode(() => outlook.readInboxEmails(ACCOUNT, {limit: 1}), [
        '1\t2026-08-01 09:30\n2\t2026-08-05 09:30\n3\t2026-08-03 09:30\n',
        row('2', 'Newest', '', '', '', ''),
    ]);
    assert.deepEqual(mail.map(m => m.entryId), ['2']);
});

test('readEmailBody splits the quoted original off the reply', async () => {
    const body = 'Thanks, noted.\n\nOn 1 Aug 2026, Jo wrote:\n> the original';
    const result = await decode(() => outlook.readEmailBody('1263', {includeQuoted: true}), [
        row('1263', ' Re: Rates ', 'Jo Doe', 'jo@x.com', '2026-08-01 09:30', names('a.pdf'), body),
    ]);
    assert.equal(result.subject, 'Re: Rates');
    assert.equal(result.body, 'Thanks, noted.');
    assert.match(result.quotedOriginal, /the original/);
    assert.deepEqual(result.attachmentNames, ['a.pdf']);
    assert.equal(result.attachmentCount, 1);
});

test('readSelectedEmail rebuilds the folder it sits in from its container chain', async () => {
    const detail = row('1263', ' BL types ', 'Jo Doe', 'jo@x.com', '2026-08-17 09:00', '', 'body text');
    const selected = await decode(() => outlook.readSelectedEmail(), [
        detail + row(ACCOUNT, `Drafts${LS}[Gmail]`, '7'),
    ]);
    assert.equal(selected.subject, 'BL types');
    assert.equal(selected.body, 'body text', 'the detail row keeps its body last');
    assert.equal(selected.folderPath, '\\\\someone@example.com\\[Gmail]\\Drafts');

    const profile = require('../dist/mac/profile.js');
    const real = profile.readProfileAccounts;
    // A mailbox with no account object: only its root folder's id names it.
    profile.readProfileAccounts = async () => [{emailAccount: ACCOUNT, folderIds: {'root folder': 7}}];
    try {
        const byRoot = await decode(() => outlook.readSelectedEmail(), [detail + row('', 'Inbox', '7')]);
        assert.equal(byRoot.folderPath, '\\\\someone@example.com\\Inbox');
        const unnamed = await decode(() => outlook.readSelectedEmail(), [detail + row('', 'Inbox', '8')]);
        assert.equal(unnamed.folderPath, '', 'a mailbox nothing names reports no path');
    } finally {
        profile.readProfileAccounts = real;
    }
});

test('searchInboxByFilter rebuilds each folder path and drops what the details deny', async () => {
    const matches = await decode(
        () => outlook.searchInboxByFilter(ACCOUNT, {subjectLike: '*invoice*', requireAttachment: true}),
        [
            row(`INBOX${LS}Clients`, '1263', 'Invoice 42', '2026-08-01 09:30') +
            row('INBOX', '1264', 'Invoice 43', '2026-08-02 09:30') +
            row('INBOX', '1265', 'Unrelated', '2026-08-02 09:30'),
            // 1264 comes back with no attachment, so requireAttachment drops it.
            row('1263', 'Jo Doe', 'jo@x.com', names('inv.pdf'), 'body text') +
            row('1264', 'Al Roe', 'al@x.com', '', 'body text'),
        ],
    );
    assert.deepEqual(matches.map(m => m.entryId), ['1263']);
    assert.equal(matches[0].folderPath, '\\\\someone@example.com\\INBOX\\Clients');
    assert.equal(matches[0].senderEmail, 'jo@x.com');
    assert.deepEqual(matches[0].attachmentNames, ['inv.pdf']);
});

test('searchInboxByFilter keeps the flags that decide how a pattern parses, and no others', async () => {
    const index = row('INBOX', '1', 'Invoice 42', '2026-08-01 09:30') + row('INBOX', '2', '42 invoice', '2026-08-02 09:30');
    const details = row('1', 'Jo Doe', 'jo@x.com', '', '') + row('2', 'Jo Doe', 'jo@x.com', '', '');
    const ids = async pattern => (await decode(
        () => outlook.searchInboxByFilter(ACCOUNT, {subjectPattern: pattern}), [index, details],
    )).map(m => m.entryId).sort();
    // Without u, \p{L} is an identity escape and matches only the text "p{L}".
    assert.deepEqual(await ids(/^\p{L}/u), ['1']);
    // A kept g would carry lastIndex into the next subject and miss it.
    assert.deepEqual(await ids(/invoice/g), ['1', '2']);
});

test('searchInboxByFilter reads the Inbox root when it is one of the folders named, as on Windows', async () => {
    for (const filter of [{}, {includeFolders: ['Inbox']}, {includeFolders: ['Clients']}]) {
        const {
            scripts,
            error
        } = await captureScripts(macRun, ['runOsaScript'], () => outlook.searchInboxByFilter(ACCOUNT, filter));
        if (error) throw error;
        // Decided in the script, against the root's own name: only Outlook
        // knows what this mailbox calls its Inbox.
        assert.match(scripts[0], /set rootScope to \(count of onlyNames\) is 0\s+try\s+if onlyNames contains rootName then set rootScope to true/);
        assert.match(scripts[0], /return my scanFolder\(rootInbox, rootName, cutoff, skipIds, skipNames, onlyNames, rootScope\)/);
    }
});

test('searchInboxByFilter never walks into a root that holds no incoming mail', async () => {
    const {NON_INCOMING_ROOTS} = require('../dist/mail.js');
    const {MAC_ROOT_TERMS} = require('../dist/mac/scripts.js');
    const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], () => outlook.searchInboxByFilter(ACCOUNT));
    if (error) throw error;
    // The roots Windows cuts, from the one list both platforms read.
    for (const rootId of NON_INCOMING_ROOTS) {
        const term = MAC_ROOT_TERMS[rootId];
        assert.ok(scripts[0].includes(`set end of skipIds to (id of (${term} of targetAcct))`), term);
    }
});

test('listOutlookDrafts sorts newest first and reports truncation against the whole folder', async () => {
    const result = await decode(() => outlook.listOutlookDrafts(ACCOUNT, {limit: 1}), [
        row('1', 'Older', 'Jo Doe', 'jo@x.com', '  spaced   body  ', '2', '2026-08-01 09:30') +
        row('2', 'Newer', 'Al Roe', 'al@x.com', 'b', '0', '2026-08-03 09:30'),
    ]);
    assert.equal(result.count, 2);
    assert.equal(result.truncated, true);
    assert.deepEqual(result.drafts.map(d => d.entryId), ['2']);
    assert.deepEqual(result.foldersScanned, ['\\\\someone@example.com\\Drafts']);
    const older = await decode(() => outlook.listOutlookDrafts(ACCOUNT), [
        row('1', 'Older', 'Jo Doe', 'jo@x.com', '  spaced   body  ', '2', '2026-08-01 09:30'),
    ]);
    assert.equal(older.truncated, false);
    assert.equal(older.drafts[0].bodyPreview, 'spaced body', 'whitespace runs collapse');
    assert.equal(older.drafts[0].hasAttachments, true);
    assert.deepEqual(older.drafts[0].toEmails, ['jo@x.com']);
});

test('a draft set to send from another account is neither listed nor sent by sendAllDrafts', async () => {
    const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], async () => {
        await outlook.listOutlookDrafts(ACCOUNT);
        await outlook.sendAllDrafts(ACCOUNT);
    }, ['', row('0')]);
    if (error) throw error;
    const [list, sendAll] = scripts;
    for (const script of scripts) {
        assert.match(script, /set boundElsewhere to \(boundTo is not "" and boundTo is not "someone@example\.com"\)/);
    }
    assert.match(list, /if not boundElsewhere then set out to out &/);
    // Skipped, as Windows skips it, rather than sent from the account it names.
    assert.match(sendAll, /if not boundElsewhere then\s+try\s+set subj[^]*?send theMsg[^]*?end if/);
});

test('a batch reports the summary and the per-item failures apart', async () => {
    const sent = await decode(() => outlook.sendDrafts(ACCOUNT, ['1263', '1264']), [
        row('1') + row('1264', 'Stuck', 'draft is not bound to this account'),
    ]);
    assert.equal(sent.sent, 1);
    assert.deepEqual(sent.failed, [{entryId: '1264', subject: 'Stuck', error: 'draft is not bound to this account'}]);

    const moved = await decode(() => outlook.moveOutlookEmails(ACCOUNT, ['1263', '1264'], 'Archive'), [
        row('2', 'true'),
    ]);
    assert.deepEqual(moved, {
        folderPath: '\\\\someone@example.com\\Inbox\\Archive',
        folderCreated: true,
        moved: 2,
        failed: [],
    });
});

test('each move is confirmed on its own, and a silent no-op goes through the Move dialog as that item', async () => {
    let moved;
    const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], async () => {
        moved = await outlook.moveOutlookEmails(ACCOUNT, ['1263', '1264'], 'Archive');
    }, [row('1', 'false') + row('1264', 'Stuck', "Outlook's Move dialog found no folder named 'Archive'.")]);
    if (error) throw error;
    const [script] = scripts;
    const start = script.indexOf('repeat with theId');
    // The last one: the loop holds the folder-chain walk, a repeat of its own.
    const end = script.lastIndexOf('end repeat');
    const loop = script.slice(start, end);
    assert.match(loop, /if fromId is not scopeId then/, 'mail already in the folder is not moved again');
    assert.match(loop, /set countBefore to count of messages of scopeFolder\s+move theMsg to scopeFolder\s+if \(count of messages of scopeFolder\) is not greater than countBefore then\s+my fileThroughMoveDialog\(\(contents of theId\), subj, scopeLeaf, scopeLabel, scopeId, countBefore, "someone@example\.com"\)/,
        'the dialog only when the scripted move filed nothing, for that one item');
    assert.doesNotMatch(script.slice(end), /did not grow/, 'no batch-wide check is left to discard the per-item results');
    assert.equal(moved.moved, 1);
    assert.deepEqual(moved.failed.map(f => f.entryId), ['1264']);
});

test('the Move dialog is named the way it lists a folder: its name, then the mailbox and the folders above it', async () => {
    const {
        scripts,
        error
    } = await captureScripts(macRun, ['runOsaScript'], () => outlook.moveOutlookEmails(ACCOUNT, ['1263'], 'Archive'), [row('1', 'false')]);
    if (error) throw error;
    const [script] = scripts;
    const label = script.slice(script.indexOf('set scopeId to id of scopeFolder'), script.indexOf('set movedCount to 0'));
    assert.match(label, /set theFolder to scopeFolder/, 'the chain walked is the destination\'s');
    assert.match(label, /set scopeLeaf to item 1 of chainNames\s+set scopeLabel to scopeLeaf & ", " & "someone@example\.com"/);
    // Nearest first, so the folders above it are added from the top down:
    // "Archive, someone@example.com/Inbox".
    assert.match(label, /repeat with i from \(count of chainNames\) to 2 by -1\s+set scopeLabel to scopeLabel & "\/" & \(item i of chainNames\)/);
});

test('the Move dialog is pressed only when its search shows the one folder, and is left as it was found otherwise', async () => {
    const {
        scripts,
        error
    } = await captureScripts(macRun, ['runOsaScript'], () => outlook.moveOutlookEmails(ACCOUNT, ['1263'], 'Archive'), [row('1', 'false')]);
    if (error) throw error;
    const [script] = scripts;
    const handler = script.slice(script.indexOf('on fileThroughMoveDialog('), script.indexOf('end fileThroughMoveDialog'));
    assert.ok(handler.length > 0, 'the script carries the handler it calls');
    assert.match(handler, /if \(id of folder of message id theId\) is scopeId then return/, 'mail the scripted move did file is not opened');
    assert.match(handler, /nm ends with \(\(character id 8226\) & " " & acctAddr\) and \(subj is "" or nm starts with subj\)/, 'the email\'s own window, by its title');
    assert.match(handler, /if \(\(name of window 1\) as string\) is not winTitle then error/, 'and still in front when the menu is used');
    const guard = handler.indexOf('if shown is not {scopeLabel} then');
    const press = handler.indexOf('click button "Move" of dlg');
    assert.ok(guard > 0 && press > guard, 'Move is pressed only past the check that the search left that one row');
    assert.ok(handler.slice(guard, press).includes('error "Outlook\'s Move dialog finds more than one folder for '), 'other rows are named, and nothing is pressed');
    assert.match(handler.slice(press), /if \(count of messages of mail folder id scopeId\) > countBefore then set landed to true/);
    assert.match(handler.slice(press), /if not landed then error/, 'a press that filed nothing fails as the scripted move would');
    // Whatever fails, the dialog is cancelled, the email's window closed and
    // the app that was in front put back.
    assert.match(handler, /on error errText number errNum\s+my leaveMoveDialog\(dlg, winId, prevApp\)/);
    const leave = script.slice(script.indexOf('on leaveMoveDialog('), script.indexOf('end leaveMoveDialog'));
    assert.match(leave, /click button "Cancel" of dlg/);
    assert.match(leave, /close window id winId/);
    assert.match(leave, /set frontmost of application process prevApp to true/);
});

test('the Move dialog says which privacy setting macOS wants, in plain ASCII', async () => {
    const {MOVE_DIALOG_REFUSALS} = require('../dist/mac/folders.js');
    const {
        scripts,
        error
    } = await captureScripts(macRun, ['runOsaScript'], () => outlook.moveOutlookEmails(ACCOUNT, ['1263'], 'Archive'), [row('1', 'false')]);
    if (error) throw error;
    const [script] = scripts;
    assert.match(script, /if not \(UI elements enabled\) then error number -25211/, 'asked before Outlook is touched');
    assert.match(script, new RegExp(String.raw`if errNum is -25211 or errNum is -1719 then error "${MOVE_DIALOG_REFUSALS.accessibility.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`));
    assert.match(script, /if errNum is -1743 then error "Outlook ignored the scripted move/);
    for (const [name, text] of Object.entries(MOVE_DIALOG_REFUSALS)) {
        assert.doesNotMatch(text, /[^\x20-\x7e]/, `${name}: plain ASCII, since it travels inside a script literal`);
        assert.match(text, /System Settings > Privacy & Security > /, name);
    }
});

test('an id macOS cannot use is reported beside the ids that worked', async () => {
    const moved = await decode(() => outlook.moveOutlookEmails(ACCOUNT, ['1263', 'AABBCC'], 'Archive'), [
        row('1', 'false'),
    ]);
    assert.equal(moved.moved, 1);
    assert.deepEqual(moved.failed.map(f => f.entryId), ['AABBCC']);
    assert.match(moved.failed[0].error, /Outlook for Mac message id/);
});

test('a delete or a move refuses mail it cannot show to be the account\'s, before touching it', async () => {
    // A message id resolves against the whole application, so this is the
    // counterpart of Windows resolving every id in the account's own store.
    const cases = [
        ['delete', 'delete theMsg', () => outlook.deleteOutlookEmails(ACCOUNT, ['1263'], {allowProtected: true}), ''],
        ['move', 'move theMsg to scopeFolder', () => outlook.moveOutlookEmails(ACCOUNT, ['1263'], 'Archive'), row('1', 'false')],
    ];
    for (const [verb, action, run, response] of cases) {
        const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], run, [response]);
        if (error) throw error;
        const [script] = scripts;
        const refusal = script.indexOf(`if not inAccount then error "item is in another mailbox, not this account - refusing to ${verb}"`);
        assert.ok(refusal > 0, `${verb}: every item is checked against the account`);
        assert.ok(refusal < script.indexOf(action), `${verb}: and before it is touched, whatever allowProtected says`);
        // Either sign vouches for an item: its folder chain ending at the
        // account's root folder, or its own account carrying the address.
        assert.match(script, /\(id of cur\) is acctRootId then set inAccount to true/, verb);
        assert.match(script, /\(ownerAddr as string\) is "someone@example\.com" then set inAccount to true/, verb);
    }
});

test('a mailbox with no account object is vouched for by the root folder the profile names', async () => {
    const profiled = {emailAccount: ACCOUNT, folderIds: {'inbox': 148, 'sent items': 149, 'root folder': 142}};
    await withProfile([profiled], null, async () => {
        const {
            scripts,
            error
        } = await captureScripts(macRun, ['runOsaScript'], () => outlook.deleteOutlookEmails(ACCOUNT, ['1263']));
        if (error) throw error;
        assert.match(scripts[0], /if acctRootId is missing value then set acctRootId to 142/);
    });
    const {
        scripts,
        error
    } = await captureScripts(macRun, ['runOsaScript'], () => outlook.deleteOutlookEmails(ACCOUNT, ['1263']));
    if (error) throw error;
    assert.doesNotMatch(scripts[0], /set acctRootId to \d/, 'with no profile row, only the account object can name the root');
});

// ── Composing from a mailbox with no account object ──────────────────────

const PROFILED_MAILBOX = {
    emailAccount: ACCOUNT,
    folderIds: {
        'inbox': 148,
        'sent items': 149,
        'deleted items': 146,
        'drafts': 147,
        'junk mail': 150,
        'root folder': 142
    },
};

/** The one script `operation` generates while the profile lists `accounts`. */
async function composedScript(accounts, operation) {
    let script;
    await withProfile(accounts, null, async () => {
        const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], operation);
        if (error) throw error;
        script = scripts[0];
    });
    return script;
}

const COMPOSERS = {
    replyOutlookEmail: extra => outlook.replyOutlookEmail({
        emailAccount: ACCOUNT,
        entryId: '1263',
        htmlBody: '<p>ok</p>', ...extra
    }),
    sendOutlookEmail: extra => outlook.sendOutlookEmail({
        emailAccount: ACCOUNT, to: 'a@example.com', subject: 'Rates', htmlBody: '<p>x</p>', ...extra,
    }),
};

test('a draft from a mailbox with no account object is filed through its window, once it is from that mailbox', async () => {
    for (const [name, compose] of Object.entries(COMPOSERS)) {
        const variable = name === 'replyOutlookEmail' ? 'theReply' : 'newMsg';
        const script = await composedScript([PROFILED_MAILBOX], () => compose({openDraftWindow: false}));
        assert.doesNotMatch(script, /INVALID_REQUEST\] Outlook publishes no account object/, `${name}: a draft is not refused`);
        assert.match(script, /set draftsFolder to mail folder id 147/, `${name}: the mailbox's own Drafts, by id`);
        assert.match(
            script,
            new RegExp(`if targetAcct is not missing value then\\s+set account of ${variable} to targetAcct`),
            `${name}: bound only where there is an account object to bind`,
        );
        assert.match(script, new RegExp(`open ${variable}\\n`), `${name}: filed through its compose window`);
        assert.match(script, /does not end with \(\(character id 8226\) & " " & "someone@example\.com"\)/, `${name}: the window names the sender`);
        assert.ok(
            script.indexOf('close window id draftWinId saving no') < script.indexOf('close window id draftWinId saving yes'),
            `${name}: a draft from another mailbox is closed unsaved before any is saved`,
        );
        assert.match(script, /Outlook would send this draft from another mailbox than 'someone@example\.com'/, name);
        assert.match(script, /if targetAcct is not missing value then\s+move/, `${name}: an account object still files by move`);
    }
});

test('a mailbox with no account object still cannot send or show mail as it is composed', async () => {
    for (const [name, compose] of Object.entries(COMPOSERS)) {
        for (const extra of [{sendImmediately: true}, {openDraftWindow: true}]) {
            const script = await composedScript([PROFILED_MAILBOX], () => compose(extra));
            assert.match(script, /INVALID_REQUEST\] Outlook publishes no account object for 'someone@example\.com'/, `${name} ${JSON.stringify(extra)}`);
        }
    }
});

test('a draft from an account the profile does not list is filed by move, with no window', async () => {
    for (const [name, compose] of Object.entries(COMPOSERS)) {
        const script = await composedScript([], () => compose({openDraftWindow: false}));
        assert.match(script, /set draftsFolder to drafts of targetAcct/, name);
        assert.match(script, /^ {4}set account of (theReply|newMsg) to targetAcct$/m, `${name}: bound unconditionally`);
        assert.doesNotMatch(script, /draftWinId/, `${name}: no window`);
    }
});

test('the window-filed drafts compile', {skip: fs.existsSync('/Applications/Microsoft Outlook.app') ? false : 'needs Outlook installed'}, async () => {
    for (const [name, compose] of Object.entries(COMPOSERS)) {
        const script = await composedScript([PROFILED_MAILBOX], () => compose({openDraftWindow: false}));
        const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-osa-'));
        try {
            const input = path.join(dir, 'generated.applescript');
            fs.writeFileSync(input, macRun.AS_HANDLERS + script, 'utf-8');
            execFileSync('osacompile', ['-o', path.join(dir, 'generated.scpt'), input], {stdio: 'pipe'});
        } catch (error) {
            assert.fail(`${name}: ${error.stderr || error.message}`);
        } finally {
            fs.rmSync(dir, {recursive: true, force: true});
        }
    }
});

test('deleteOutlookEmails protects the roots Windows does, by the names this account gives them', async () => {
    const {PROTECTED_ROOTS} = require('../dist/mail.js');
    const {MAC_ROOT_TERMS} = require('../dist/mac/scripts.js');
    const {
        scripts,
        error
    } = await captureScripts(macRun, ['runOsaScript'], () => outlook.deleteOutlookEmails(ACCOUNT, ['1263']));
    if (error) throw error;
    for (const rootId of PROTECTED_ROOTS) {
        assert.ok(scripts[0].includes(`set protectedFolder to ${MAC_ROOT_TERMS[rootId]} of targetAcct`), MAC_ROOT_TERMS[rootId]);
    }
    assert.match(scripts[0], /set isProtected to \(protectedNames contains rootName\)/);
});

test('deleteOutlookEmails counts each outcome and rebuilds the folder path it was found in', async () => {
    const result = await decode(() => outlook.deleteOutlookEmails(ACCOUNT, ['1', '2', '3']), [
        row('1', ' Kept ', `Inbox${LS}Clients`, 'refused', 'protected') +
        row('2', 'Gone', `Archive`, 'deleted', '') +
        row('3', 'Broken', '', 'failed', 'no such item'),
    ]);
    assert.deepEqual(
        {deleted: result.deleted, refused: result.refused, failed: result.failed},
        {deleted: 1, refused: 1, failed: 1},
    );
    assert.equal(result.items[0].folderPath, '\\\\someone@example.com\\Inbox\\Clients');
    assert.equal(result.items[0].subject, 'Kept');
    assert.equal(result.items[2].folderPath, '', 'an item with no chain reports no path');
});

test('purgeDeletedItems keeps what the window spares and purges the rest', async () => {
    const result = await decode(() => outlook.purgeDeletedItems(ACCOUNT, {olderThanDays: 30}), [
        row('1', 'true') + row('2', 'false') + row('3', 'false'),
        row('2', '0'),
    ]);
    assert.deepEqual(result, {
        folderPath: '\\\\someone@example.com\\Deleted Items',
        dryRun: false,
        matched: 2,
        purged: 2,
        kept: 1,
        failed: 0,
    });
});

test('a dry-run purge never runs its second script', async () => {
    const {scripts} = await captureScripts(
        macRun,
        ['runOsaScript'],
        () => outlook.purgeDeletedItems(ACCOUNT, {dryRun: true}),
        [row('1', 'false')],
    );
    assert.equal(scripts.length, 1);
});

test('readSentRecipientGroups lowercases recipients and drops a message with none', async () => {
    const groups = await decode(() => outlook.readSentRecipientGroups(ACCOUNT), [
        row('1', 'Rate request', '2026-08-01 09:30') + row('2', 'No recipients', '2026-08-02 09:30'),
        row('1', `Bob@Example.com${LS}CAROL@example.com`),
    ]);
    assert.deepEqual(groups, [{
        entryId: '1',
        subject: 'Rate request',
        sentOn: '2026-08-01 09:30',
        recipients: ['bob@example.com', 'carol@example.com'],
    }]);
});

test('cleanUndeliverableEmails classifies the bounce and mines its failed recipients', async () => {
    const result = await decode(() => outlook.cleanUndeliverableEmails(ACCOUNT, {dryRun: false}), [
        row('1', 'Undeliverable: Rate request', 'Mail Delivery Subsystem', 'daemon@x.com', '2026-08-01 09:30') +
        row('2', 'An ordinary email', 'Jo Doe', 'jo@x.com', '2026-08-01 09:30'),
        row('1', 'Delivery failed for bob@example.com'),
        row('1'),
    ]);
    assert.equal(result.matchedCount, 1);
    assert.equal(result.deletedCount, 1);
    assert.equal(result.dryRun, false);
    assert.deepEqual(result.matched[0].failedRecipients, ['bob@example.com']);
    assert.ok(result.matched[0].matchedReason, 'a match says why it matched');
});

// The index's sixth field is the structural bounce signal: whether the message's
// own headers declare it a delivery-status notification. It is what stands in
// for the MessageClass Windows reads, so these cover the parity it buys.
test('a bounce only its headers give away is still caught', async () => {
    // A localized subject and an unremarkable sender: no subject phrase and no
    // daemon fingerprint matches, so before the header read this was a miss.
    const result = await decode(() => outlook.cleanUndeliverableEmails(ACCOUNT, {dryRun: true}), [
        row('1', 'Unzustellbar: Ratenanfrage', 'Mail Gateway', 'noreply@corp.example', '2026-08-01 09:30', 'true') +
        row('2', 'An ordinary email', 'Jo Doe', 'jo@x.com', '2026-08-01 09:30', 'false'),
        row('1', 'Delivery failed for bob@example.com'),
    ]);
    assert.equal(result.matchedCount, 1);
    assert.equal(result.matched[0].entryId, '1', 'the ordinary email is left alone');
    assert.equal(result.matched[0].matchedReason, 'Non-delivery report (NDR)');
    assert.deepEqual(result.matched[0].failedRecipients, ['bob@example.com']);
});

test('the structural signal outranks the sender fingerprint, as it does on Windows', async () => {
    // Both signals fire. Windows reports the NDR because the message class is
    // tested first, so the same message must report the same reason here.
    const result = await decode(() => outlook.cleanUndeliverableEmails(ACCOUNT, {dryRun: true}), [
        row('1', 'Undeliverable: Rate request', 'Mail Delivery Subsystem', 'daemon@x.com', '2026-08-01 09:30', 'true'),
        row('1', 'Delivery failed for bob@example.com'),
    ]);
    assert.equal(result.matched[0].matchedReason, 'Non-delivery report (NDR)');
});

test('a folder whose headers cannot be read still classifies on sender and subject', async () => {
    // The header read is wrapped in a try, so an unreadable folder yields no
    // flags at all — the cheaper rules must still find the bounce.
    const result = await decode(() => outlook.cleanUndeliverableEmails(ACCOUNT, {dryRun: true}), [
        row('1', 'Undeliverable: Rate request', 'Mail Delivery Subsystem', 'daemon@x.com', '2026-08-01 09:30'),
        row('1', 'Delivery failed for bob@example.com'),
    ]);
    assert.equal(result.matchedCount, 1);
    assert.equal(result.matched[0].matchedReason, 'From mail-delivery system');
});

test('readTemplateEmails reports a missing folder with the folders there are', async () => {
    const result = await decode(() => outlook.readTemplateEmails(ACCOUNT), [
        row('0', `Inbox${LS}Archive`),
    ]);
    assert.equal(result.folderFound, false);
    assert.deepEqual(result.availableFolders, ['Inbox', 'Archive']);
    assert.deepEqual(result.templates, []);
});

test('readTemplateEmails reports the whole path of a folder found below the top level', async () => {
    const result = await decode(() => outlook.readTemplateEmails(ACCOUNT, {folder: 'Templates'}), [
        row('1', `Templates${LS}Work${LS}Inbox`),
    ]);
    assert.equal(result.folderFound, true);
    assert.equal(result.folderPath, '\\\\someone@example.com\\Inbox\\Work\\Templates', 'as Windows FolderPath reports it');
});

test('readTemplateEmails strips embedded images and reports the markers a template uses', async () => {
    const html = '<html><body><p>Hi {{NAME}}</p><img src="cid:logo123" width="10"><IMG SRC="CID:seal"><p>[[QUOTE]]body[[/QUOTE]]</p></body></html>';
    const result = await decode(() => outlook.readTemplateEmails(ACCOUNT, {subject: 'Rate reply'}), [
        row('1', 'Templates') + row('10', ' Rate reply ', '2026-08-01 09:30') + row('11', 'Other', '2026-08-02 09:30'),
        row('10', html),
    ]);
    assert.equal(result.folderFound, true);
    assert.equal(result.folderPath, '\\\\someone@example.com\\Templates');
    assert.deepEqual(result.templates.map(t => t.subject), ['Rate reply'], 'the subject filter picks one');
    assert.doesNotMatch(result.templates[0].htmlBody, /cid:/i, 'an embedded image would show as a broken placeholder');
    assert.deepEqual(result.templates[0].placeholders, ['NAME']);
    assert.deepEqual(result.templates[0].sections, ['QUOTE']);
});

test('replyOutlookEmail reports the addresses the reply actually resolved to', async () => {
    const result = await decode(
        () => outlook.replyOutlookEmail({emailAccount: ACCOUNT, entryId: '1263', htmlBody: '<p>ok</p>'}),
        [['a@example.com, b@example.com', 'RE: Rates', 'jo@x.com', ''].join(FS)],
    );
    assert.deepEqual(result, {
        to: 'a@example.com, b@example.com',
        subject: 'RE: Rates',
        repliedToSender: 'jo@x.com',
        draftEntryId: '',
    });
});

test('replyOutlookEmail reports the id of the reply it filed in Drafts', async () => {
    const reply = {emailAccount: ACCOUNT, entryId: '1263', htmlBody: '<p>ok</p>', openDraftWindow: false};
    const result = await decode(
        () => outlook.replyOutlookEmail(reply),
        [['jo@x.com', 'RE: Rates', 'jo@x.com', '1290'].join(FS)],
    );
    assert.equal(result.draftEntryId, '1290');

    const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], () => outlook.replyOutlookEmail(reply));
    if (error) throw error;
    const [script] = scripts;
    assert.ok(
        script.indexOf('set draftIdsBefore') < script.indexOf('set theReply to reply to'),
        'Drafts is listed before the reply exists, so the reply is among the ids it gains',
    );
    assert.doesNotMatch(script, /id of theReply/, 'moving rewrites the id, so the one read before it is stale');
});

test('sendOutlookEmail reports the id of the draft it filed, and none for a window', async () => {
    const base = {emailAccount: ACCOUNT, to: 'a@example.com', subject: 'Rates', htmlBody: '<p>x</p>'};
    assert.deepEqual(await decode(() => outlook.sendOutlookEmail({...base, openDraftWindow: false}), ['1263']), {
        draftEntryId: '1263',
    });
    assert.deepEqual(await decode(() => outlook.sendOutlookEmail(base), ['']), {draftEntryId: ''});
});

test('a filed draft is named by the id Drafts gained, never the one it had before the move', async () => {
    const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], () => outlook.sendOutlookEmail({
        emailAccount: ACCOUNT, to: 'a@example.com', subject: 'Rates', htmlBody: '<p>x</p>', openDraftWindow: false,
    }));
    if (error) throw error;
    const [script] = scripts;
    assert.match(script, /draftIdsBefore does not contain/);
    assert.doesNotMatch(script, /id of newMsg/, 'moving rewrites the id, so the one read before it is stale');
    assert.match(script, /return filedDraftId/);
});

test('saveEmailAttachments never overwrites, and refuses a name the email does not carry', async () => {
    const destDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-mac-att-'));
    fs.writeFileSync(path.join(destDir, 'inv.pdf'), 'already here');
    const detail = row('1263', 'Subject', 'Jo Doe', 'jo@x.com', '2026-08-01 09:30', names('inv.pdf', 'second.pdf'));

    const saved = await decode(() => outlook.saveEmailAttachments('1263', ['inv.pdf'], {destDir}), [detail, '']);
    // The counter goes before the extension, so the file stays openable.
    assert.equal(path.basename(saved[0].path), 'inv (1).pdf', 'the existing file is left alone');
    assert.equal(saved[0].fileName, 'inv.pdf');
    assert.equal(saved[0].senderEmail, 'jo@x.com');

    await assert.rejects(
        decode(() => outlook.saveEmailAttachments('1263', ['missing.pdf'], {destDir}), [detail, '']),
        error => {
            assert.equal(error.code, 'NOT_FOUND');
            // The message names the email that resolved, which is what makes a
            // stale id visible rather than looking like a missing attachment.
            assert.match(error.message, /Attachments present: inv\.pdf, second\.pdf/);
            return true;
        },
    );
    fs.rmSync(destDir, {recursive: true, force: true});
});

test('an unnamed attachment keeps its place, so the attachment saved is the one named', async () => {
    const destDir = tempDir('ob-mac-att-');
    const detail = row('1263', 'Subject', 'Jo Doe', 'jo@x.com', '2026-08-01 09:30', names('', 'invoice.pdf'));
    const {scripts, error} = await captureScripts(macRun, ['runOsaScript'],
        () => outlook.saveEmailAttachments('1263', ['invoice.pdf'], {destDir}), [detail, '']);
    if (error) throw error;
    assert.match(scripts[1], /save \(item 2 of attachments of theMsg\)/, 'item 1 is the unnamed inline image');

    const body = await decode(() => outlook.readEmailBody('1263'), [detail]);
    assert.deepEqual(body.attachmentNames, ['', 'invoice.pdf']);
    assert.equal(body.attachmentCount, 2, 'an unnamed attachment still counts, as on Windows');
});

test('a positional list keeps empty elements, and an empty list stays empty, under osascript', {skip: needsOsascript}, () => {
    const file = path.join(tempDir('ob-osa-list-'), 'list.applescript');
    for (const items of [[], [''], ['', 'invoice.pdf'], ['a.pdf', 'b.pdf']]) {
        fs.writeFileSync(file, `${macRun.AS_HANDLERS}\nreturn my sanitizePositionalList({${items.map(macRun.asString).join(', ')}})`, 'utf8');
        const out = execFileSync('osascript', [file], {encoding: 'utf8'}).replace(/\n$/, '');
        assert.deepEqual(macRun.splitPositionalList(out), items);
    }
});

test('a framed field carrying the separators still decodes into one record', () => {
    // sanitize() strips FIELD and RECORD from a field but leaves LIST alone, so
    // this is what a subject full of separators looks like by the time it lands.
    const raw = row('1263', 'Subject with a list sep', `a.pdf${LS}b.pdf`);
    const [record] = macRun.splitRecords(raw);
    const fields = macRun.splitFields(record);
    assert.equal(macRun.field(fields, 1), 'Subject with a list sep');
    assert.deepEqual(macRun.splitList(macRun.field(fields, 2)), ['a.pdf', 'b.pdf']);
    assert.equal(macRun.intField(fields, 9, 7), 7, 'an absent numeric field falls back');
    assert.equal(macRun.boolField(macRun.splitFields(macRun.splitRecords(row('true'))[0]), 0), true);
});

// ── The profile database is an enrichment pass, never a failure ──────────

test('a profile database that cannot be read loses no mailbox AppleScript found', async () => {
    const profile = require('../dist/mac/profile.js');
    const real = {readdirSync: fs.readdirSync, existsSync: fs.existsSync, statSync: fs.statSync};
    // readdir succeeds and the file is there, but stat is refused — the shape a
    // half-granted Full Disk Access, or Outlook rewriting a profile mid-call, takes.
    fs.readdirSync = () => ['Main Profile', 'Other Profile'];
    fs.existsSync = () => true;
    fs.statSync = () => {
        const error = new Error('EPERM: operation not permitted, stat');
        error.code = 'EPERM';
        throw error;
    };
    try {
        assert.deepEqual(await profile.readProfileAccounts(), []);
    } finally {
        Object.assign(fs, real);
    }
});

// ── A profile macOS refuses is said, not taken for an empty one ──────────

const REFUSAL = "macOS is not letting Terminal read Outlook's profile. Turn on Terminal under Full Disk Access.";

/** Runs `check` with the profile reporting `accounts` and `refusal`, then puts the real ones back. */
async function withProfile(accounts, refusal, check) {
    const profile = require('../dist/mac/profile.js');
    const real = {readProfileAccounts: profile.readProfileAccounts, profileRefusal: profile.profileRefusal};
    profile.readProfileAccounts = async () => accounts;
    profile.profileRefusal = async () => refusal;
    try {
        await check();
    } finally {
        Object.assign(profile, real);
    }
}

test('nothing to list while macOS refuses the profile is that refusal, not an empty list', async () => {
    await withProfile([], REFUSAL, async () => {
        const {error} = await captureScripts(macRun, ['runOsaScript'], () => outlook.getOutlookAccounts(), ['']);
        assert.equal(error?.code, 'PERMISSION_DENIED');
        assert.equal(error.message, `No Outlook mailbox could be listed. ${REFUSAL}`);
        assert.deepEqual(
            await decode(() => outlook.getOutlookAccounts(), ['ops@example.com']),
            ['ops@example.com'],
            'a list that found something is kept: what it holds works',
        );
    });
    await withProfile([], null, async () => {
        assert.deepEqual(await decode(() => outlook.getOutlookAccounts(), ['']), [], 'nothing refused, nothing found');
    });
});

test('an address the refused profile could not confirm fails as PERMISSION_DENIED when the probe misses it too', async () => {
    const lookup = async () => {
        const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], () => outlook.readInboxEmails(ACCOUNT));
        if (error) throw error;
        return scripts[0];
    };
    await withProfile([], REFUSAL, async () => {
        const script = await lookup();
        assert.match(script, /\[outlook-bridge:PERMISSION_DENIED\] Account 'someone@example\.com' is not one/);
        assert.ok(script.includes(REFUSAL.replace(/"/g, '\\"')), 'the refusal says what to do');
    });
    await withProfile([], null, async () => {
        assert.match(await lookup(), /\[outlook-bridge:ACCOUNT_NOT_FOUND\]/, 'with nothing refused, not found is not found');
    });
    await withProfile([{emailAccount: ACCOUNT, folderIds: {'inbox': 148}}], REFUSAL, async () => {
        assert.doesNotMatch(await lookup(), /PERMISSION_DENIED/, 'a mailbox the profile names needs no access said about it');
    });
});

test('profileRefusal tells macOS refusing the profile apart from there being none', async () => {
    const profile = require('../dist/mac/profile.js');
    const real = fs.readdirSync;
    const failWith = code => () => {
        const error = new Error(`${code}: on the profiles folder`);
        error.code = code;
        throw error;
    };
    try {
        fs.readdirSync = failWith('ENOENT');
        assert.equal(await profile.profileRefusal(), null, 'no profiles folder: New Outlook, or none at all');
        for (const code of ['EPERM', 'EACCES']) {
            fs.readdirSync = failWith(code);
            const refusal = await profile.profileRefusal();
            assert.match(refusal, /^macOS is not letting .+ read Outlook's profile/, code);
            assert.match(refusal, /System Settings > Privacy & Security > Full Disk Access/, code);
            assert.doesNotMatch(refusal, /[^\x20-\x7e]/, 'plain ASCII, since it travels inside a script literal');
        }
        fs.readdirSync = () => ['Main Profile'];
        assert.equal(await profile.profileRefusal(), null, 'a profile that can be read is not refused');
    } finally {
        fs.readdirSync = real;
    }
});

test('a move beside the Inbox walks from the root folder the profile names, and reports that path', async () => {
    let moved;
    const script = await composedScript([PROFILED_MAILBOX], async () => {
        moved = await outlook.moveOutlookEmails(ACCOUNT, ['1263'], '\\\\someone@example.com\\Processed');
    });
    assert.match(script, /set scopeFolder to mail folder id 142/, 'the top of the mailbox, by id');
    assert.doesNotMatch(script, /mail folder id 148/, 'not the Inbox');
    assert.match(script, /if \(name of sf as string\) is "Processed" then/);
    assert.equal(moved.folderPath, '\\\\someone@example.com\\Processed');
});
