// Folder-string parsing and reply-quote splitting: pure, platform-neutral, and
// shared by both implementations — so a regression here is wrong on Windows AND
// macOS, silently, with no Outlook in sight to catch it.
const test = require('node:test');
const assert = require('node:assert/strict');

const {mailFolderRef, splitQuotedOriginal, WELL_KNOWN_FOLDERS} = require('../dist/mail.js');

test('mailFolderRef: a bare name is Inbox-relative', () => {
    assert.deepEqual(mailFolderRef('Invoices'), {
        rootId: 6,
        rootLabel: 'Inbox',
        segments: ['Invoices'],
    });
});

test('mailFolderRef: a well-known root is recognized and consumed', () => {
    assert.deepEqual(mailFolderRef('Sent Items'), {
        rootId: 5,
        rootLabel: 'Sent Items',
        segments: [],
    });
    assert.deepEqual(mailFolderRef('Deleted Items\\2026'), {
        rootId: 3,
        rootLabel: 'Deleted Items',
        segments: ['2026'],
    });
});

test('mailFolderRef: root matching is case-insensitive but preserves the label', () => {
    const ref = mailFolderRef('sENT iTEMS\\Archive');
    assert.equal(ref.rootId, 5);
    assert.equal(ref.rootLabel, 'sENT iTEMS');
    assert.deepEqual(ref.segments, ['Archive']);
});

test('mailFolderRef: a full FolderPath drops the mailbox prefix', () => {
    assert.deepEqual(mailFolderRef('\\\\team@x.com\\Inbox\\Invoices\\Paid'), {
        rootId: 6,
        rootLabel: 'Inbox',
        segments: ['Invoices', 'Paid'],
    });
});

test('mailFolderRef: forward slashes and stray separators are tolerated', () => {
    assert.deepEqual(mailFolderRef('Invoices/Paid').segments, ['Invoices', 'Paid']);
    assert.deepEqual(mailFolderRef('  Invoices \\\\ Paid  ').segments, ['Invoices', 'Paid']);
});

test('mailFolderRef: an empty-ish string yields the Inbox root with no segments', () => {
    // readInboxEmails relies on this to reject a folder argument that trims to
    // nothing rather than silently reading the Inbox root.
    assert.deepEqual(mailFolderRef('   ').segments, []);
    assert.deepEqual(mailFolderRef('\\\\').segments, []);
});

test('WELL_KNOWN_FOLDERS maps the roots both platforms agree on', () => {
    assert.equal(WELL_KNOWN_FOLDERS.inbox, 6);
    assert.equal(WELL_KNOWN_FOLDERS['sent items'], 5);
    assert.equal(WELL_KNOWN_FOLDERS.drafts, 16);
    assert.equal(WELL_KNOWN_FOLDERS['deleted items'], 3);
    assert.equal(WELL_KNOWN_FOLDERS.outbox, 4);
});

test('splitQuotedOriginal: Outlook "-----Original Message-----"', () => {
    const r = splitQuotedOriginal(
        'Our price is $850.\n\n-----Original Message-----\nFrom: us\nWhat is your price?',
    );
    assert.equal(r.body, 'Our price is $850.');
    assert.equal(r.separator, '-----Original Message-----');
    assert.match(r.quoted, /What is your price\?/);
});

test('splitQuotedOriginal: the gmail-style "On … wrote:" line', () => {
    const r = splitQuotedOriginal(
        'Confirmed.\nOn Tue, Jul 21, 2026 at 9:14 AM John Doe <j@x.com> wrote:\n> anything?',
    );
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
    // The split is abandoned rather than returning an empty body — losing the
    // sender's own words is the one outcome worth avoiding.
    const text = '-----Original Message-----\nFrom: us\nWhat is your price?\n\n$850.';
    const r = splitQuotedOriginal(text);
    assert.equal(r.body, text);
    assert.equal(r.separator, '');
});

test('splitQuotedOriginal: text with no quote comes back whole', () => {
    const r = splitQuotedOriginal('Just a note.\nNothing quoted.');
    assert.equal(r.body, 'Just a note.\nNothing quoted.');
    assert.equal(r.quoted, '');
    assert.equal(r.separator, '');
});

test('splitQuotedOriginal: the EARLIEST separator wins', () => {
    const r = splitQuotedOriginal(
        'Answer.\n> quoted line\n-----Original Message-----\nolder',
    );
    assert.equal(r.body, 'Answer.');
    assert.equal(r.separator, '> quoted lines');
});

test('splitQuotedOriginal: CRLF input splits the same as LF', () => {
    const r = splitQuotedOriginal('Answer.\r\n\r\n-----Original Message-----\r\nFrom: us');
    assert.equal(r.body, 'Answer.');
    assert.equal(r.separator, '-----Original Message-----');
});
