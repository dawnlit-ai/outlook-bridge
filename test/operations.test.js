// The public operations over a fake backend: what each one validates, which
// defaults it applies, and the shared work it does before a platform sees the
// request. Runs anywhere — this is the layer both platforms have in common.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const {createOperations} = require('../dist/bridge.js');
const {fakeBackend, tempFileWith} = require('./support.js');

const ACCOUNT = 'me@example.com';

function setup(overrides) {
    const fake = fakeBackend(overrides);
    return {...fake, outlook: createOperations(fake.backend)};
}

async function rejectsWith(promise, code) {
    await assert.rejects(promise, error => {
        assert.equal(error.code, code, error.message);
        return true;
    });
}

// ── Reading ──────────────────────────────────────────────────────────────

test('readInboxEmails applies its defaults and trims the account', async () => {
    const {outlook, last} = setup();
    await outlook.readInboxEmails(`  ${ACCOUNT} `);
    const [request] = last('readInboxEmails').args;
    assert.equal(request.account, ACCOUNT);
    assert.deepEqual(request.folder, {rootId: 6, rootLabel: 'Inbox', segments: []});
    assert.equal(request.daysBack, 60);
    assert.equal(request.limit, 50);
    assert.equal(request.previewChars, 600);
});

test('readInboxEmails resolves a folder and clamps counts into range', async () => {
    const {outlook, last} = setup();
    await outlook.readInboxEmails(ACCOUNT, {folder: 'Sent Items\\2026', daysBack: 0, limit: 7.9, previewChars: -3});
    const [request] = last('readInboxEmails').args;
    assert.deepEqual(request.folder, {rootId: 5, rootLabel: 'Sent Items', segments: ['2026']});
    assert.equal(request.folderLabel, 'Sent Items\\2026');
    assert.equal(request.daysBack, 1);
    assert.equal(request.limit, 7);
    assert.equal(request.previewChars, 0);
});

test('a blank folder means the Inbox root; a folder of only separators is refused', async () => {
    const {outlook, last} = setup();
    await outlook.readInboxEmails(ACCOUNT, {folder: '   '});
    assert.equal(last('readInboxEmails').args[0].folder.rootId, 6);
    await rejectsWith(outlook.readInboxEmails(ACCOUNT, {folder: '\\\\'}), 'INVALID_REQUEST');
});

test('a full folder path to the Inbox itself names the Inbox root', async () => {
    const {outlook, last} = setup();
    await outlook.readInboxEmails(ACCOUNT, {folder: '\\\\me@example.com\\Inbox'});
    assert.deepEqual(last('readInboxEmails').args[0].folder, {rootId: 6, rootLabel: 'Inbox', segments: []});
});

test('arguments of the wrong type never reach the backend', async () => {
    const {outlook, calls} = setup();
    await rejectsWith(outlook.readInboxEmails(ACCOUNT, {daysBack: '7; Remove-Item x'}), 'INVALID_REQUEST');
    await rejectsWith(outlook.readInboxEmails(''), 'INVALID_REQUEST');
    await rejectsWith(outlook.searchInboxByFilter(ACCOUNT, {subjectPattern: 'not a regexp'}), 'INVALID_REQUEST');
    await rejectsWith(outlook.moveOutlookEmails(ACCOUNT, 'not-an-array', 'Archive'), 'INVALID_REQUEST');
    assert.equal(calls.length, 0);
});

test('searchInboxByFilter normalizes its filter', async () => {
    const {outlook, last} = setup();
    await outlook.searchInboxByFilter(ACCOUNT, {
        subjectLike: ' *invoice* ',
        subjectPattern: /inv\d+/,
        includeFolders: [' Clients '],
        includeBody: false,
    });
    const [request] = last('searchInboxByFilter').args;
    assert.equal(request.daysBack, 60);
    assert.equal(request.subjectLike, '*invoice*');
    assert.equal(request.subjectPattern.source, 'inv\\d+');
    assert.equal(request.excludeReplies, false);
    assert.equal(request.requireAttachment, false);
    assert.deepEqual(request.includeFolders, ['Clients']);
    assert.deepEqual(request.excludeFolders, []);
    assert.equal(request.includeBody, false);
});

test('readEmailBody splits the reply from its quote and caps both', async () => {
    const body = `${'a'.repeat(30)}\n\n-----Original Message-----\n${'q'.repeat(50)}`;
    const {outlook, last} = setup({
        readEmailBody: () => ({
            entryId: 'E1',
            subject: 'S',
            senderName: 'N',
            senderEmail: 'n@x.com',
            receivedTime: 'T',
            body,
            attachmentNames: ['a.pdf']
        }),
    });
    const result = await outlook.readEmailBody({entryId: 'E1', storeId: ' S1 ', subject: 'ignored'}, {
        maxChars: 10,
        includeQuoted: true,
        maxQuotedChars: 20
    });
    assert.deepEqual(last('readEmailBody').args[0], {entryId: 'E1', storeId: 'S1'});
    assert.equal(result.body, 'a'.repeat(10));
    assert.equal(result.truncated, true);
    assert.equal(result.bodyLength, 30);
    assert.equal(result.quoteSeparator, '-----Original Message-----');
    assert.equal(result.quotedOriginal.length, 10, 'capped by the tighter of maxChars and maxQuotedChars');
    assert.equal(result.attachmentCount, 1);
});

test('a listing row can be passed back in as the email', async () => {
    const {outlook, last} = setup();
    const row = {entryId: 'E9', storeId: 'S9', subject: 'Hello', body: 'Hi'};
    await outlook.openOutlookEmail(row);
    assert.deepEqual(last('openOutlookEmail').args[0], {entryId: 'E9', storeId: 'S9'});
    await outlook.openOutlookEmail('E10');
    assert.deepEqual(last('openOutlookEmail').args[0], {entryId: 'E10'});
});

// ── Sending ──────────────────────────────────────────────────────────────

test('sendOutlookEmail splits recipients, keeping display names whole', async () => {
    const {outlook, last} = setup();
    await outlook.sendOutlookEmail({
        emailAccount: ACCOUNT,
        to: '"Doe, Jo" <jo@example.com>; b@example.com, c@example.com',
        cc: ['d@example.com', 'e@example.com;f@example.com'],
        subject: 'Hi',
        htmlBody: '<p>x</p>',
    });
    const [request] = last('sendOutlookEmail').args;
    assert.deepEqual(request.to, ['"Doe, Jo" <jo@example.com>', 'b@example.com', 'c@example.com']);
    assert.deepEqual(request.cc, ['d@example.com', 'e@example.com', 'f@example.com']);
    assert.deepEqual(request.bcc, []);
    assert.equal(request.disposition, 'display', 'a draft with a window by default');
});

test('sendOutlookEmail maps the two flags onto one disposition', async () => {
    const {outlook, last} = setup();
    const base = {emailAccount: ACCOUNT, to: 'a@example.com', subject: 's', htmlBody: 'b'};
    await outlook.sendOutlookEmail({...base, openDraftWindow: false});
    assert.equal(last('sendOutlookEmail').args[0].disposition, 'save');
    await outlook.sendOutlookEmail({...base, sendImmediately: true, openDraftWindow: false});
    assert.equal(last('sendOutlookEmail').args[0].disposition, 'send');
});

test('sendOutlookEmail resolves attachments to absolute paths and refuses missing ones', async () => {
    const {outlook, last} = setup();
    const file = tempFileWith('report.pdf');
    const relative = path.relative(process.cwd(), file);
    await outlook.sendOutlookEmail({
        emailAccount: ACCOUNT,
        to: 'a@example.com',
        subject: 's',
        htmlBody: 'b',
        attachments: [relative]
    });
    assert.deepEqual(last('sendOutlookEmail').args[0].attachments, [file]);
    await rejectsWith(
        outlook.sendOutlookEmail({
            emailAccount: ACCOUNT,
            to: 'a@example.com',
            subject: 's',
            htmlBody: 'b',
            attachments: ['/no/such/file.pdf']
        }),
        'NOT_FOUND',
    );
});

test('an email sent immediately needs a recipient; a draft does not', async () => {
    const {outlook} = setup();
    await rejectsWith(outlook.sendOutlookEmail({
        emailAccount: ACCOUNT,
        subject: 's',
        htmlBody: 'b',
        sendImmediately: true
    }), 'INVALID_REQUEST');
    await outlook.sendOutlookEmail({emailAccount: ACCOUNT, subject: 's', htmlBody: 'b'});
});

test('replyOutlookEmail composes a template section, placeholders and signature before the backend sees it', async () => {
    const template = [
        '<html><body>',
        '<p>Hello,</p>',
        '<p>[[YES]]Accepted.[[/YES]]</p>',
        '<p>[[NO]]Declined.[[/NO]]</p>',
        '<p>{{NOTE}}</p>',
        '<p>{{SIGNATURE}}</p>',
        '</body></html>',
    ].join('');
    const {outlook, last, calls} = setup({
        readTemplateEmails: request => ({
            folderFound: true,
            folderPath: '\\\\me\\Templates',
            templates: [{
                entryId: 'T1',
                subject: 'Reply',
                htmlBody: template,
                bodyPreview: '',
                sections: ['YES', 'NO'],
                placeholders: ['NOTE', 'SIGNATURE'],
                lastModified: ''
            }],
            availableFolders: [],
            request,
        }),
        readOutlookSignatureHtml: () => '<html><body><b>Jo</b></body></html>',
    });
    await outlook.replyOutlookEmail({
        emailAccount: ACCOUNT,
        entryId: 'E1',
        templateSubject: 'Reply',
        templateSection: 'yes',
        templatePlaceholders: {NOTE: 'See you.'},
        signatureName: 'Work',
        replyAll: true,
        openDraftWindow: false,
    });
    const templateRead = calls.find(call => call.name === 'readTemplateEmails').args[0];
    assert.deepEqual(templateRead, {
        account: ACCOUNT,
        folder: 'Templates',
        limit: 50,
        includeBody: true,
        subject: 'Reply'
    });
    const [request] = last('replyOutlookEmail').args;
    assert.match(request.html, /Accepted\./);
    assert.doesNotMatch(request.html, /Declined|\[\[|\{\{|<html|<body/);
    assert.match(request.html, /See you\./);
    assert.match(request.html, /<b>Jo<\/b>/);
    assert.equal(request.replyAll, true);
    assert.equal(request.disposition, 'save');
});

test('replyOutlookEmail names the signatures there are when one is missing', async () => {
    const {outlook} = setup({
        readOutlookSignatureHtml: () => '',
        listOutlookSignatures: () => ['Home', 'Work'],
    });
    await assert.rejects(
        outlook.replyOutlookEmail({
            emailAccount: ACCOUNT,
            entryId: 'E1',
            htmlBody: '<p>{{SIGNATURE}}</p>',
            signatureName: 'Nope'
        }),
        error => error.code === 'NOT_FOUND' && error.kind === 'signature' && /Home, Work/.test(error.message),
    );
});

test('replyOutlookEmail needs a body or a template', async () => {
    const {outlook, calls} = setup();
    await rejectsWith(outlook.replyOutlookEmail({emailAccount: ACCOUNT, entryId: 'E1'}), 'INVALID_REQUEST');
    assert.equal(calls.length, 0);
});

// ── Batches ──────────────────────────────────────────────────────────────

test('batch operations deduplicate ids and skip the backend for an empty batch', async () => {
    const {outlook, last, calls} = setup();
    await outlook.sendDrafts(ACCOUNT, ['a', ' b ', 'a']);
    assert.deepEqual(last('sendDrafts').args[0].entryIds, ['a', 'b']);
    const before = calls.length;
    assert.deepEqual(await outlook.deleteOutlookDrafts(ACCOUNT, []), {deleted: 0, failed: []});
    assert.deepEqual(await outlook.moveOutlookEmails(ACCOUNT, [], 'Archive'), {
        folderPath: '',
        folderCreated: false,
        moved: 0,
        failed: []
    });
    assert.equal(calls.length, before);
});

test('cleanUndeliverableEmails dry-runs unless deleting is asked for', async () => {
    const {outlook, last} = setup();
    await outlook.cleanUndeliverableEmails(ACCOUNT);
    assert.equal(last('cleanUndeliverableEmails').args[0].dryRun, true);
    await outlook.cleanUndeliverableEmails(ACCOUNT, {dryRun: false, daysBack: 7});
    assert.equal(last('cleanUndeliverableEmails').args[0].dryRun, false);
    assert.equal(last('cleanUndeliverableEmails').args[0].daysBack, 7);
});

test('deleteOutlookEmails refuses protected mail and deletes for real only when asked', async () => {
    const {outlook, last} = setup();
    await outlook.deleteOutlookEmails(ACCOUNT, ['a']);
    assert.deepEqual(last('deleteOutlookEmails').args[0], {
        account: ACCOUNT,
        entryIds: ['a'],
        allowProtected: false,
        dryRun: false
    });
});

// ── Attachments ──────────────────────────────────────────────────────────

test('saveEmailAttachment saves into a private directory by default and returns the one file', async () => {
    const {outlook, last} = setup();
    const saved = await outlook.saveEmailAttachment({entryId: 'E1', storeId: 'S1'}, 'invoice.pdf');
    const [request] = last('saveEmailAttachments').args;
    assert.deepEqual(request.email, {entryId: 'E1', storeId: 'S1'});
    assert.deepEqual(request.fileNames, ['invoice.pdf']);
    assert.ok(fs.statSync(request.destDir).isDirectory());
    assert.equal(saved.fileName, 'invoice.pdf');
    fs.rmSync(request.destDir, {recursive: true, force: true});
});

test('a failed save removes the private directory it created', async () => {
    let destDir;
    const {outlook} = setup({
        saveEmailAttachments: request => {
            destDir = request.destDir;
            const {NotFoundError} = require('../dist/errors.js');
            throw new NotFoundError('attachment', 'nope');
        },
    });
    await rejectsWith(outlook.saveEmailAttachments('E1', ['x.pdf']), 'NOT_FOUND');
    assert.equal(fs.existsSync(destDir), false);
});

test('a named destination is created and resolved to an absolute path', async () => {
    const {outlook, last} = setup();
    const base = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-dest-'));
    const relative = path.relative(process.cwd(), path.join(base, 'nested', 'dir'));
    await outlook.saveEmailAttachments('E1', ['a.pdf'], {destDir: relative});
    assert.equal(last('saveEmailAttachments').args[0].destDir, path.join(base, 'nested', 'dir'));
    assert.ok(fs.statSync(path.join(base, 'nested', 'dir')).isDirectory());
    fs.rmSync(base, {recursive: true, force: true});
});

// ── Templates ────────────────────────────────────────────────────────────

test('template operations apply their defaults', async () => {
    const {outlook, last} = setup();
    await outlook.readTemplateEmails(ACCOUNT);
    assert.deepEqual(last('readTemplateEmails').args[0], {
        account: ACCOUNT,
        folder: 'Templates',
        limit: 20,
        includeBody: true,
        subject: undefined
    });
    await outlook.saveTemplateEmail(ACCOUNT, {subject: ' Reply ', htmlBody: '<p>x</p>'});
    assert.deepEqual(last('saveTemplateEmail').args[0], {
        account: ACCOUNT,
        folder: 'Templates',
        subject: 'Reply',
        htmlBody: '<p>x</p>'
    });
});
