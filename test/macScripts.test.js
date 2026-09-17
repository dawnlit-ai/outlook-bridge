// Every complete AppleScript the macOS backend generates, compiled.
//
// A single bad dictionary term fails the WHOLE script at compile time rather
// than at the offending line, so a live error never names its cause — and the
// automation itself can only run against a real mailbox. So the runner is
// stubbed, every operation is driven through the public operations, and each
// captured script goes through osacompile, which resolves terminology against
// the installed Outlook without opening a session or touching a message.
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

const OUTLOOK_APP = '/Applications/Microsoft Outlook.app';
const runnable = process.platform === 'darwin' && fs.existsSync(OUTLOOK_APP);
const skip = runnable ? false : 'needs macOS with Outlook installed';

const ACCOUNT = 'someone@example.com';
const {FIELD_SEP: FS, RECORD_SEP: RS, LIST_SEP: LS} = macRun;

/** One framed record, the way the scripts emit them. */
function row(...fields) {
    return fields.join(FS) + RS;
}

/** Compile (never run) a script, returning osacompile's complaint or ''. The
 *  shared handlers are prepended because the runner prepends them. */
function compileError(body) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-osa-'));
    const input = path.join(dir, 'generated.applescript');
    try {
        fs.writeFileSync(input, macRun.AS_HANDLERS + body, 'utf-8');
        execFileSync('osacompile', ['-o', path.join(dir, 'generated.scpt'), input], {stdio: 'pipe'});
        return '';
    } catch (error) {
        return String(error.stderr || error.message);
    } finally {
        fs.rmSync(dir, {recursive: true, force: true});
    }
}

const outlook = createOperations(macBackend);

/**
 * Every operation, with the canned runner output that walks it through all of
 * its passes. Ids are the small integers Outlook for Mac uses.
 */
const CASES = {
    getOutlookAccounts: {run: () => outlook.getOutlookAccounts()},
    sendOutlookEmail: {
        run: () => outlook.sendOutlookEmail({
            emailAccount: ACCOUNT,
            to: '"Doe, Jo" <a@example.com>, b@example.com',
            cc: 'c@example.com',
            bcc: ['d@example.com'],
            subject: "Quarter's figures",
            htmlBody: '<p>Hello "world" \\ backslash</p>',
            attachments: [tempFileWith('report.pdf')],
            openDraftWindow: false,
        }),
    },
    replyOutlookEmail: {
        run: () => outlook.replyOutlookEmail({
            emailAccount: ACCOUNT,
            entryId: '1263',
            htmlBody: '<p>Thanks — noted.</p>',
            replyAll: true,
            openDraftWindow: false,
        }),
    },
    sendAllDrafts: {run: () => outlook.sendAllDrafts(ACCOUNT)},
    readInboxEmails: {
        run: () => outlook.readInboxEmails(ACCOUNT, {daysBack: 30, limit: 10, folder: 'Inbox\\Clients'}),
        responses: ['1263\t2026-08-01 09:30\n'],
    },
    readInboxEmailsNoPreview: {
        run: () => outlook.readInboxEmails(ACCOUNT, {folder: 'Sent Items', previewChars: 0}),
        responses: ['1263\t2026-08-01 09:30\n'],
    },
    searchInboxByFilter: {
        run: () => outlook.searchInboxByFilter(ACCOUNT, {
            daysBack: 30,
            subjectLike: '*invoice*',
            subjectPattern: /invoice\s+\d+/,
            excludeReplies: true,
            requireAttachment: true,
        }),
        responses: [row('INBOX' + LS + 'Clients', '1263', 'Invoice 42', '2026-08-01 09:30')],
    },
    searchInboxByFilterNoBody: {
        run: () => outlook.searchInboxByFilter(ACCOUNT, {includeBody: false, includeFolders: ['Clients']}),
        responses: [row('INBOX', '1263', 'Hello', '2026-08-01 09:30')],
    },
    readSelectedEmail: {
        run: () => outlook.readSelectedEmail(),
        responses: [row('1263', 'Subject', 'Name', 'a@example.com', '2026-08-01 09:30', '', 'body')],
    },
    readEmailBody: {
        run: () => outlook.readEmailBody('1263', {includeQuoted: true}),
        responses: [row('1263', 'Subject', 'Name', 'a@example.com', '2026-08-01 09:30', '', 'body')],
    },
    openOutlookEmail: {run: () => outlook.openOutlookEmail('1263')},
    listInboxFolders: {run: () => outlook.listInboxFolders(ACCOUNT, {maxDepth: 3})},
    moveOutlookEmails: {
        run: () => outlook.moveOutlookEmails(ACCOUNT, ['1263', '1264'], "Clients\\Bob's mail\\2026", {createIfMissing: true}),
        responses: [row('2', 'true')],
    },
    listOutlookDrafts: {run: () => outlook.listOutlookDrafts(ACCOUNT, {limit: 50, previewChars: 200})},
    deleteOutlookDrafts: {run: () => outlook.deleteOutlookDrafts(ACCOUNT, ['1263']), responses: [row('1')]},
    sendDrafts: {run: () => outlook.sendDrafts(ACCOUNT, ['1263']), responses: [row('1')]},
    deleteOutlookEmails: {run: () => outlook.deleteOutlookEmails(ACCOUNT, ['1263'], {allowProtected: true})},
    deleteOutlookEmailsDryRun: {run: () => outlook.deleteOutlookEmails(ACCOUNT, ['1263'], {dryRun: true})},
    purgeDeletedItems: {
        run: () => outlook.purgeDeletedItems(ACCOUNT, {olderThanDays: 30}),
        responses: [row('1263', 'false'), row('1', '0')],
    },
    saveEmailAttachments: {
        run: () => outlook.saveEmailAttachments('1263', ['invoice.pdf']),
        responses: [row('1263', 'Subject', 'Name', 'a@example.com', '2026-08-01 09:30', 'invoice.pdf')],
    },
    cleanUndeliverableEmails: {
        run: () => outlook.cleanUndeliverableEmails(ACCOUNT, {daysBack: 30, dryRun: false}),
        responses: [
            row('1263', 'Undeliverable: Rate request', 'Mail Delivery Subsystem', 'mailer-daemon@x.com', '2026-08-01 09:30'),
            row('1263', 'failed for bob@example.com'),
            row('1'),
        ],
    },
    collectBouncedRecipients: {
        run: () => outlook.collectBouncedRecipients(ACCOUNT, {daysBack: 30}),
        responses: [
            row('1263', 'Undeliverable: Rate request', 'Mail Delivery Subsystem', 'mailer-daemon@x.com', '2026-08-01 09:30'),
            row('1263', 'failed for bob@example.com'),
        ],
    },
    readSentRecipientGroups: {
        run: () => outlook.readSentRecipientGroups(ACCOUNT, {daysBack: 30, limit: 100}),
        responses: [row('1263', 'Rate request', '2026-08-01 09:30')],
    },
    listOutlookSignatures: {run: () => outlook.listOutlookSignatures()},
    readOutlookSignatureHtml: {run: () => outlook.readOutlookSignatureHtml('Default "work"')},
    readTemplateEmails: {
        run: () => outlook.readTemplateEmails(ACCOUNT, {subject: 'Rate reply'}),
        responses: [row('1', 'Templates') + row('1263', 'Rate reply', '2026-08-01 09:30')],
    },
    readTemplateEmailsNoBody: {
        run: () => outlook.readTemplateEmails(ACCOUNT, {includeBody: false}),
        responses: [row('1', 'Templates') + row('1263', 'Rate reply', '2026-08-01 09:30')],
    },
    saveTemplateEmail: {
        run: () => outlook.saveTemplateEmail(ACCOUNT, {subject: 'Rate reply', htmlBody: '<p>{{SIGNATURE}}</p>'}),
        responses: [row('Templates', 'false')],
    },
    editEmailTemplate: {
        run: () => outlook.editEmailTemplate('Rate reply', '<p>body</p>'),
        responses: ['<p>edited</p>'],
    },
};

async function scriptsOf(name) {
    const {run, responses} = CASES[name];
    const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], run, responses);
    if (error) throw error;
    return scripts;
}

for (const name of Object.keys(CASES)) {
    test(`${name} generates AppleScript that compiles`, {skip}, async () => {
        const scripts = await scriptsOf(name);
        assert.ok(scripts.length > 0, 'the operation generated no script at all');
        scripts.forEach((script, index) => {
            assert.equal(compileError(script), '', `${name} script ${index + 1}/${scripts.length}`);
        });
    });
}

test('every operation generates its scripts without throwing', async () => {
    for (const name of Object.keys(CASES)) {
        assert.ok((await scriptsOf(name)).length > 0, name);
    }
});

test('the two-pass readers really do reach their second pass', async () => {
    for (const name of ['readInboxEmails', 'searchInboxByFilter', 'purgeDeletedItems', 'readSentRecipientGroups', 'readTemplateEmails', 'saveEmailAttachments']) {
        assert.ok((await scriptsOf(name)).length >= 2, `${name} stopped after one script`);
    }
});

// ── Guards for the failures compiling cannot catch ───────────────────────
// Both shipped once, and neither is a compile error: the script builds, then
// fails only against a live Outlook. They need no Outlook to check, since the
// scripts are captured rather than run.

/** AppleScript terms that can't be variable names. Assigning to one compiles,
 *  then fails at run time with a message naming the value, not the variable. */
const RESERVED_WORDS = new Set([
    'at', 'rest', 'end', 'count', 'length', 'text', 'item', 'id', 'name', 'contents',
    'result', 'first', 'last', 'front', 'back', 'middle', 'every', 'some', 'it', 'me',
    'my', 'its', 'error', 'script', 'run', 'div', 'mod', 'and', 'or', 'not', 'as',
    'ref', 'reference', 'of', 'in', 'on', 'to', 'from', 'by', 'with', 'without',
    'into', 'through', 'thru', 'given', 'timeout', 'property', 'return', 'copy',
    'beginning', 'above', 'below', 'since', 'until', 'while', 'repeat', 'tell',
]);

/** A script without its comments, which quote the very patterns these guards forbid. */
function withoutComments(script) {
    return script.split('\n').filter(line => !line.trim().startsWith('--')).join('\n');
}

async function everyGeneratedScript() {
    const all = [];
    for (const name of Object.keys(CASES)) {
        for (const script of await scriptsOf(name)) all.push({name, script: withoutComments(script)});
    }
    return all;
}

test('no generated script assigns to a reserved AppleScript word', async () => {
    for (const {name, script} of await everyGeneratedScript()) {
        for (const [, variable] of script.matchAll(/^\s*set ([A-Za-z_]\w*) to /gm)) {
            assert.ok(!RESERVED_WORDS.has(variable.toLowerCase()), `${name}: "set ${variable} to ..." uses a reserved word`);
        }
    }
});

test('no generated script reads a record field through a nested accessor', async () => {
    for (const {name, script} of await everyGeneratedScript()) {
        const nested = script.match(/\b(?:address|name) of \((?:sender|email address) of /);
        assert.equal(nested, null, `${name}: ${nested && nested[0]}... must bind the record first`);
    }
});

test('a Windows EntryID is refused before any script is built', async () => {
    const windowsEntryId = '00000000AABBCCDD1122334455667788';
    for (const call of [
        () => outlook.readEmailBody(windowsEntryId),
        () => outlook.openOutlookEmail(windowsEntryId),
        () => outlook.replyOutlookEmail({emailAccount: ACCOUNT, entryId: windowsEntryId, htmlBody: '<p>x</p>'}),
    ]) {
        const {scripts, error} = await captureScripts(macRun, ['runOsaScript'], call);
        assert.equal(error.code, 'INVALID_REQUEST');
        assert.match(error.message, /Outlook for Mac message id/);
        assert.equal(scripts.length, 0);
    }
});

test('a batch reports an unusable id beside the others rather than discarding them', async () => {
    const {error} = await captureScripts(macRun, ['runOsaScript'], async () => {
        const result = await outlook.moveOutlookEmails(ACCOUNT, ['not-a-mac-id'], 'Archive');
        assert.equal(result.moved, 0);
        assert.deepEqual(result.failed.map(f => f.entryId), ['not-a-mac-id']);
    });
    assert.equal(error, undefined);
});
