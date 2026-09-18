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
const {captureScripts, tempFileWith} = require('./support.js');

const {FIELD_SEP: FS, RECORD_SEP: RS, LIST_SEP: LS} = macRun;
const ACCOUNT = 'someone@example.com';
const outlook = createOperations(macBackend);

/** One framed record, the way the scripts emit them. */
function row(...fields) {
    return fields.join(FS) + RS;
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
    const destDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-mac-probe-'));
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
            responses: [row('1263', 'Subject', 'Name', 'a@example.com', '2026-08-01 09:30', PROBE)],
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
        row('1263', ' Subject A ', 'Jo Doe', 'jo@x.com', 'preview A', `a.pdf${LS}b.pdf`),
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
        row('1263', ' Re: Rates ', 'Jo Doe', 'jo@x.com', '2026-08-01 09:30', 'a.pdf', body),
    ]);
    assert.equal(result.subject, 'Re: Rates');
    assert.equal(result.body, 'Thanks, noted.');
    assert.match(result.quotedOriginal, /the original/);
    assert.deepEqual(result.attachmentNames, ['a.pdf']);
    assert.equal(result.attachmentCount, 1);
});

test('searchInboxByFilter rebuilds each folder path and drops what the details deny', async () => {
    const matches = await decode(
        () => outlook.searchInboxByFilter(ACCOUNT, {subjectLike: '*invoice*', requireAttachment: true}),
        [
            row(`INBOX${LS}Clients`, '1263', 'Invoice 42', '2026-08-01 09:30') +
            row('INBOX', '1264', 'Invoice 43', '2026-08-02 09:30') +
            row('INBOX', '1265', 'Unrelated', '2026-08-02 09:30'),
            // 1264 comes back with no attachment, so requireAttachment drops it.
            row('1263', 'Jo Doe', 'jo@x.com', 'inv.pdf', 'body text') +
            row('1264', 'Al Roe', 'al@x.com', '', 'body text'),
        ],
    );
    assert.deepEqual(matches.map(m => m.entryId), ['1263']);
    assert.equal(matches[0].folderPath, '\\\\someone@example.com\\INBOX\\Clients');
    assert.equal(matches[0].senderEmail, 'jo@x.com');
    assert.deepEqual(matches[0].attachmentNames, ['inv.pdf']);
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

test('an id macOS cannot use is reported beside the ids that worked', async () => {
    const moved = await decode(() => outlook.moveOutlookEmails(ACCOUNT, ['1263', 'AABBCC'], 'Archive'), [
        row('1', 'false'),
    ]);
    assert.equal(moved.moved, 1);
    assert.deepEqual(moved.failed.map(f => f.entryId), ['AABBCC']);
    assert.match(moved.failed[0].error, /Outlook for Mac message id/);
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

test('readTemplateEmails reports a missing folder with the folders there are', async () => {
    const result = await decode(() => outlook.readTemplateEmails(ACCOUNT), [
        row('0', `Inbox${LS}Archive`),
    ]);
    assert.equal(result.folderFound, false);
    assert.deepEqual(result.availableFolders, ['Inbox', 'Archive']);
    assert.deepEqual(result.templates, []);
});

test('readTemplateEmails strips embedded images and reports the markers a template uses', async () => {
    const html = '<html><body><p>Hi {{NAME}}</p><img src="cid:logo123" width="10"><p>[[QUOTE]]body[[/QUOTE]]</p></body></html>';
    const result = await decode(() => outlook.readTemplateEmails(ACCOUNT, {subject: 'Rate reply'}), [
        row('1', 'Templates') + row('10', ' Rate reply ', '2026-08-01 09:30') + row('11', 'Other', '2026-08-02 09:30'),
        row('10', html),
    ]);
    assert.equal(result.folderFound, true);
    assert.equal(result.folderPath, '\\\\someone@example.com\\Templates');
    assert.deepEqual(result.templates.map(t => t.subject), ['Rate reply'], 'the subject filter picks one');
    assert.doesNotMatch(result.templates[0].htmlBody, /cid:/, 'an embedded image would show as a broken placeholder');
    assert.deepEqual(result.templates[0].placeholders, ['NAME']);
    assert.deepEqual(result.templates[0].sections, ['QUOTE']);
});

test('replyOutlookEmail reports the addresses the reply actually resolved to', async () => {
    const result = await decode(
        () => outlook.replyOutlookEmail({emailAccount: ACCOUNT, entryId: '1263', htmlBody: '<p>ok</p>'}),
        [['a@example.com, b@example.com', 'RE: Rates', 'jo@x.com'].join(FS)],
    );
    assert.deepEqual(result, {
        to: 'a@example.com, b@example.com',
        subject: 'RE: Rates',
        repliedToSender: 'jo@x.com',
    });
});

test('saveEmailAttachments never overwrites, and refuses a name the email does not carry', async () => {
    const destDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-mac-att-'));
    fs.writeFileSync(path.join(destDir, 'inv.pdf'), 'already here');
    const detail = row('1263', 'Subject', 'Jo Doe', 'jo@x.com', '2026-08-01 09:30', `inv.pdf${LS}second.pdf`);

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
