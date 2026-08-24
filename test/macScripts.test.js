// Every complete AppleScript the macOS implementation generates, compiled.
//
// This is the test that earns its keep on this platform. A single bad
// dictionary term fails the WHOLE script at compile time rather than at the
// offending line, so the live error never names the thing that caused it — and
// the automation itself can only be exercised against a real mailbox, which CI
// does not have.
//
// So: stub the runner, drive every operation with plausible arguments, capture
// the scripts they build, and put each through osacompile. Compiling resolves
// terminology against the installed Outlook without opening a session, sending
// anything, or touching a single message.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');

const OUTLOOK_APP = '/Applications/Microsoft Outlook.app';
const runnable = process.platform === 'darwin' && fs.existsSync(OUTLOOK_APP);
const skip = runnable ? false : 'needs macOS with Outlook installed';

const macRun = require('../dist/mac/run.js');
const mac = require('../dist/mac/index.js');

const ACCOUNT = 'someone@example.com';
const FS = macRun.FIELD_SEP;
const RS = macRun.RECORD_SEP;
const LS = macRun.LIST_SEP;

/** Build one framed record, the way the scripts emit them. */
function row(...fields) {
    return fields.join(FS) + RS;
}

/**
 * Run `operation` with the runner stubbed out, returning every script it built.
 *
 * `responses` are handed back in order, so a two-pass reader can be driven into
 * its second pass — which is the only way that script gets generated at all.
 */
async function scriptsFrom(operation, responses = []) {
    const captured = [];
    const real = macRun.runOsaScript;
    let call = 0;
    macRun.runOsaScript = (script) => {
        captured.push(script);
        return Promise.resolve(responses[call++] ?? '');
    };
    try {
        await operation();
    } finally {
        macRun.runOsaScript = real;
    }
    return captured;
}

/** Compile (never run) a script, returning osacompile's complaint or ''. */
function compileError(source) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-osa-'));
    const input = path.join(dir, 'generated.applescript');
    try {
        fs.writeFileSync(input, source, 'utf-8');
        execFileSync('osacompile', ['-o', path.join(dir, 'generated.scpt'), input], {stdio: 'pipe'});
        return '';
    } catch (error) {
        return String(error.stderr || error.message);
    } finally {
        fs.rmSync(dir, {recursive: true, force: true});
    }
}

/**
 * Every operation, paired with the canned runner output that walks it through
 * all of its passes. Ids are the small integers Outlook for Mac uses.
 */
const CASES = {
    getOutlookAccounts: {
        run: () => mac.getOutlookAccounts(),
    },
    sendOutlookEmail: {
        run: () => mac.sendOutlookEmail({
            emailAccount: ACCOUNT,
            to: 'a@example.com, b@example.com',
            cc: 'c@example.com',
            subject: "Quarter's figures",
            htmlBody: '<p>Hello "world" \\ backslash</p>',
            attachmentPath: '/tmp/report.pdf',
            sendImmediately: false,
            openDraftWindow: false,
        }),
    },
    replyOutlookEmail: {
        run: () => mac.replyOutlookEmail({
            emailAccount: ACCOUNT,
            entryId: '1263',
            htmlBody: '<p>Thanks — noted.</p>',
            replyAll: true,
            sendImmediately: false,
            openDraftWindow: false,
        }),
    },
    sendAllDrafts: {
        run: () => mac.sendAllDrafts(ACCOUNT),
    },
    readInboxEmails: {
        run: () => mac.readInboxEmails(ACCOUNT, 30, 10, 'Inbox\\Clients'),
        responses: ['1263\t2026-08-01 09:30\n'],
    },
    searchInboxByFilter: {
        run: () => mac.searchInboxByFilter(ACCOUNT, {
            subjectLike: '*invoice*',
            subjectPattern: /invoice\s+\d+/,
            excludeReplies: true,
            requireAttachment: true,
        }, 30),
        responses: [row('INBOX' + LS + 'Clients', '1263', 'Invoice 42', '2026-08-01 09:30')],
    },
    readSelectedEmail: {
        run: () => mac.readSelectedEmail(),
        responses: [row('1263', 'Subject', 'Name', 'a@example.com', '2026-08-01 09:30', '', 'body')],
    },
    readEmailBody: {
        run: () => mac.readEmailBody('1263', undefined, 8000, true),
        responses: [row('1263', 'Subject', 'Name', 'a@example.com', '2026-08-01 09:30', '', 'body')],
    },
    openOutlookEmail: {
        run: () => mac.openOutlookEmail('1263'),
    },
    listInboxFolders: {
        run: () => mac.listInboxFolders(ACCOUNT, 3),
    },
    moveOutlookEmails: {
        run: () => mac.moveOutlookEmails(ACCOUNT, ['1263', '1264'], "Clients\\Bob's mail\\2026", true),
        responses: [row('2', 'true')],
    },
    listOutlookDrafts: {
        run: () => mac.listOutlookDrafts(ACCOUNT, 50, 200),
    },
    deleteOutlookDrafts: {
        run: () => mac.deleteOutlookDrafts(ACCOUNT, ['1263']),
        responses: [row('1')],
    },
    deleteOutlookEmails: {
        run: () => mac.deleteOutlookEmails(ACCOUNT, ['1263'], {allowProtected: true, dryRun: false}),
    },
    deleteOutlookEmailsDryRun: {
        run: () => mac.deleteOutlookEmails(ACCOUNT, ['1263'], {dryRun: true}),
    },
    purgeDeletedItems: {
        run: () => mac.purgeDeletedItems(ACCOUNT, 30, false),
        responses: [row('1263', 'false'), row('1', '0')],
    },
    saveEmailAttachments: {
        run: () => mac.saveEmailAttachments('1263', ['invoice.pdf'], undefined, undefined),
        responses: [row('Subject', 'Name', 'a@example.com', '2026-08-01 09:30', 'invoice.pdf')],
    },
    cleanUndeliverableEmails: {
        run: () => mac.cleanUndeliverableEmails(ACCOUNT, 30, false),
        responses: [
            row('1263', 'Undeliverable: Rate request', 'Mail Delivery Subsystem', 'mailer-daemon@x.com', '2026-08-01 09:30'),
            row('1263', 'failed for bob@example.com'),
            row('1'),
        ],
    },
    collectBouncedRecipients: {
        run: () => mac.collectBouncedRecipients(ACCOUNT, 30, true),
        responses: [
            row('1263', 'Undeliverable: Rate request', 'Mail Delivery Subsystem', 'mailer-daemon@x.com', '2026-08-01 09:30'),
            row('1263', 'failed for bob@example.com'),
        ],
    },
    readSentRecipientGroups: {
        run: () => mac.readSentRecipientGroups(ACCOUNT, 30, 100),
        responses: [row('1263', 'Rate request', '2026-08-01 09:30')],
    },
    listOutlookSignatures: {
        run: () => mac.listOutlookSignatures(),
    },
    readOutlookSignatureHtml: {
        run: () => mac.readOutlookSignatureHtml('Default "work"'),
    },
    readTemplateEmails: {
        run: () => mac.readTemplateEmails(ACCOUNT, 'Templates', 20, true, 'Rate reply'),
        responses: [row('1', 'Templates') + row('1263', 'Rate reply', '2026-08-01 09:30')],
    },
    readTemplateEmailsNoBody: {
        run: () => mac.readTemplateEmails(ACCOUNT, 'Templates', 20, false),
        responses: [row('1', 'Templates') + row('1263', 'Rate reply', '2026-08-01 09:30')],
    },
    saveTemplateEmail: {
        run: () => mac.saveTemplateEmail(ACCOUNT, 'Rate reply', '<p>{{SIGNATURE}}</p>', 'Templates'),
        responses: [row('Templates', 'false')],
    },
    editEmailTemplate: {
        run: () => mac.editEmailTemplate('Rate reply', '<p>body</p>'),
        responses: ['<p>edited</p>'],
    },
};

for (const [name, {run, responses}] of Object.entries(CASES)) {
    test(`${name} generates AppleScript that compiles`, {skip}, async () => {
        const scripts = await scriptsFrom(run, responses);
        assert.ok(scripts.length > 0, 'the operation generated no script at all');
        scripts.forEach((script, index) => {
            const error = compileError(script);
            assert.equal(error, '', `${name} script ${index + 1}/${scripts.length}:\n${error}`);
        });
    });
}

test('the two-pass readers really do reach their second pass', {skip}, async () => {
    // Otherwise the cases above would quietly only ever check pass one, and the
    // detail scripts — the ones that read bodies and senders — would go
    // uncompiled while the suite stayed green.
    for (const name of [
        'readInboxEmails',
        'searchInboxByFilter',
        'purgeDeletedItems',
        'readSentRecipientGroups',
        'readTemplateEmails',
        'saveEmailAttachments',
    ]) {
        const {run, responses} = CASES[name];
        const scripts = await scriptsFrom(run, responses);
        assert.ok(scripts.length >= 2, `${name} stopped after ${scripts.length} script(s)`);
    }
});

test('caller text stays inside its AppleScript literal', {skip}, async () => {
    // A subject carrying a quote or a backslash must not be able to close the
    // literal it sits in — the AppleScript counterpart of the PowerShell
    // single-quote doubling, and the reason compiling these matters.
    const scripts = await scriptsFrom(() => mac.sendOutlookEmail({
        emailAccount: 'a"b@example.com',
        to: 'x@example.com',
        subject: 'He said "hello" \\ then left',
        htmlBody: '<p>line one\nline two "quoted"</p>',
        sendImmediately: false,
        openDraftWindow: false,
    }));
    assert.equal(compileError(scripts[0]), '');
});

// ── Guards for the failures compiling cannot catch ───────────────────────
// Both of these shipped once. Neither is a compile error: the script builds,
// osacompile accepts it, and it fails only against a live Outlook — which is
// exactly the kind of bug that needs a cheap test rather than a lucky run.
// These need no Outlook at all, since the scripts are captured, not run.

/**
 * AppleScript terms that cannot be used as variable names. Assigning to one
 * compiles and then fails at run time with a message that names the value
 * rather than the variable, so it reads like a data problem.
 */
const RESERVED_WORDS = new Set([
    'at', 'rest', 'end', 'count', 'length', 'text', 'item', 'id', 'name', 'contents',
    'result', 'first', 'last', 'front', 'back', 'middle', 'every', 'some', 'it', 'me',
    'my', 'its', 'error', 'script', 'run', 'div', 'mod', 'and', 'or', 'not', 'as',
    'ref', 'reference', 'of', 'in', 'on', 'to', 'from', 'by', 'with', 'without',
    'into', 'through', 'thru', 'given', 'timeout', 'property', 'return', 'copy',
    'beginning', 'above', 'below', 'since', 'until', 'while', 'repeat', 'tell',
]);

/** A script with its AppleScript comments removed — those quote the very
 *  patterns these guards forbid, and are not code. */
function withoutComments(script) {
    return script
        .split('\n')
        .filter(line => !line.trim().startsWith('--'))
        .join('\n');
}

async function everyGeneratedScript() {
    const all = [];
    for (const [name, {run, responses}] of Object.entries(CASES)) {
        for (const script of await scriptsFrom(run, responses)) {
            all.push({name, script: withoutComments(script)});
        }
    }
    return all;
}

test('no generated script assigns to a reserved AppleScript word', async () => {
    // `set rest to ...` and `set at to ...` both shipped and both failed live.
    for (const {name, script} of await everyGeneratedScript()) {
        for (const [, variable] of script.matchAll(/^\s*set ([A-Za-z_]\w*) to /gm)) {
            assert.ok(
                !RESERVED_WORDS.has(variable.toLowerCase()),
                `${name}: "set ${variable} to ..." uses a reserved AppleScript word`,
            );
        }
    }
});

test('no generated script reads a record field through a nested accessor', async () => {
    // `address of (sender of m)` and `address of (email address of r)` do not
    // coerce — the record has to be bound to a variable first. Inside a try, the
    // failure is silent: an empty sender, or no recipients at all.
    for (const {name, script} of await everyGeneratedScript()) {
        const nested = script.match(/\b(?:address|name) of \((?:sender|email address) of /);
        assert.equal(nested, null, `${name}: ${nested && nested[0]}... must bind the record first`);
    }
});
