// Folder strings, subject globs, reply prefixes and quote splitting: pure,
// platform-neutral, and shared by both backends — so a regression here is wrong
// on Windows AND macOS at once.
const test = require('node:test');
const assert = require('node:assert/strict');

const {
    folderOption,
    mailFolderRef,
    REPLY_PREFIX,
    splitQuotedOriginal,
    subjectGlobSource,
    WELL_KNOWN_FOLDERS,
} = require('../dist/mail.js');

test('mailFolderRef: a bare name is Inbox-relative', () => {
    assert.deepEqual(mailFolderRef('Invoices'), {rootId: 6, rootLabel: 'Inbox', segments: ['Invoices']});
});

test('mailFolderRef: a well-known root is recognized and consumed', () => {
    assert.deepEqual(mailFolderRef('Sent Items'), {rootId: 5, rootLabel: 'Sent Items', segments: []});
    assert.deepEqual(mailFolderRef('Deleted Items\\2026'), {rootId: 3, rootLabel: 'Deleted Items', segments: ['2026']});
});

test('mailFolderRef: root matching ignores case but keeps the label', () => {
    const ref = mailFolderRef('sENT iTEMS\\Archive');
    assert.equal(ref.rootId, 5);
    assert.equal(ref.rootLabel, 'sENT iTEMS');
    assert.deepEqual(ref.segments, ['Archive']);
});

test('mailFolderRef: a full folder path drops the mailbox prefix', () => {
    assert.deepEqual(mailFolderRef('\\\\team@x.com\\Inbox\\Invoices\\Paid'), {
        rootId: 6,
        rootLabel: 'Inbox',
        segments: ['Invoices', 'Paid']
    });
});

test('mailFolderRef: forward slashes and stray separators are tolerated', () => {
    assert.deepEqual(mailFolderRef('Invoices/Paid').segments, ['Invoices', 'Paid']);
    assert.deepEqual(mailFolderRef('  Invoices \\\\ Paid  ').segments, ['Invoices', 'Paid']);
});

test('folderOption: nothing given is the Inbox root; only separators is refused', () => {
    assert.deepEqual(folderOption(undefined).ref, {rootId: 6, rootLabel: 'Inbox', segments: []});
    assert.deepEqual(folderOption('  ').ref.segments, []);
    assert.throws(() => folderOption('\\\\'), error => error.code === 'INVALID_REQUEST');
    assert.throws(() => folderOption(' / '), error => error.code === 'INVALID_REQUEST');
});

test('WELL_KNOWN_FOLDERS maps the roots both platforms agree on, and cannot be changed', () => {
    assert.equal(WELL_KNOWN_FOLDERS.inbox, 6);
    assert.equal(WELL_KNOWN_FOLDERS['sent items'], 5);
    assert.equal(WELL_KNOWN_FOLDERS.drafts, 16);
    assert.equal(WELL_KNOWN_FOLDERS['deleted items'], 3);
    assert.equal(WELL_KNOWN_FOLDERS.outbox, 4);
    assert.ok(Object.isFrozen(WELL_KNOWN_FOLDERS));
});

test('subjectGlobSource: * and ? are wildcards, everything else is literal', () => {
    const re = new RegExp(subjectGlobSource('*[EXTERNAL] Invoice 10?.5*'), 'i');
    assert.ok(re.test('re: [external] invoice 104.5 attached'));
    assert.ok(!re.test('Invoice 104x5'), 'the dot is literal');
    assert.ok(!re.test('EXTERNAL Invoice 104.5'), 'the brackets are literal');
    assert.equal(subjectGlobSource('a$b'), '^a\\$b$');
});

test('REPLY_PREFIX recognizes replies and forwards across languages', () => {
    for (const subject of ['RE: x', 'Fwd: x', 'FW:x', 'AW: x', 'RE[2]: x', '回复：x', '转发: x', '답장: x']) {
        assert.ok(REPLY_PREFIX.test(subject), subject);
    }
    for (const subject of ['Regarding the invoice', 'Fwding soon', 'Rate request']) {
        assert.ok(!REPLY_PREFIX.test(subject), subject);
    }
});

test('splitQuotedOriginal: Outlook "-----Original Message-----"', () => {
    const r = splitQuotedOriginal('Our price is $850.\n\n-----Original Message-----\nFrom: us\nWhat is your price?');
    assert.equal(r.body, 'Our price is $850.');
    assert.equal(r.separator, '-----Original Message-----');
    assert.match(r.quoted, /What is your price\?/);
});

test('splitQuotedOriginal: the "On … wrote:" line', () => {
    const r = splitQuotedOriginal('Confirmed.\nOn Tue, Jul 21, 2026 at 9:14 AM Jo Doe <j@x.com> wrote:\n> anything?');
    assert.equal(r.body, 'Confirmed.');
    assert.equal(r.separator, 'On … wrote:');
});

test('splitQuotedOriginal: a From:/Sent: header block', () => {
    const r = splitQuotedOriginal('Sounds good.\n\nFrom: someone@x.com\nSent: Monday\nTo: us');
    assert.equal(r.body, 'Sounds good.');
    assert.equal(r.separator, 'From:/Sent: header block');
});

test('splitQuotedOriginal: a bare "From:" with no second header is not a separator', () => {
    const r = splitQuotedOriginal('From: the yard, the price is $850.');
    assert.equal(r.separator, '');
    assert.equal(r.quoted, '');
});

test('splitQuotedOriginal: a bottom-posted reply is never split away', () => {
    const text = '-----Original Message-----\nFrom: us\nWhat is your price?\n\n$850.';
    const r = splitQuotedOriginal(text);
    assert.equal(r.body, text);
    assert.equal(r.separator, '');
});

test('splitQuotedOriginal: the EARLIEST separator wins, and CRLF splits like LF', () => {
    assert.equal(splitQuotedOriginal('Answer.\n> quoted line\n-----Original Message-----\nolder').separator, '> quoted lines');
    assert.equal(splitQuotedOriginal('Answer.\r\n\r\n-----Original Message-----\r\nFrom: us').body, 'Answer.');
});
