// Argument checking and the file-name rules — the boundary where caller input
// becomes something a generated script can safely hold.
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

const {
    count,
    emailLocator,
    entryIdList,
    flag,
    parseRecipient,
    recipientList,
    requiredText
} = require('../dist/shared/args.js');
const {findAttachmentIndex, safeFileName, uniqueSavePath} = require('../dist/shared/attachmentMatch.js');

const invalid = error => error.code === 'INVALID_REQUEST';

test('count: omitted takes the default, fractions round down, out-of-range clamps', () => {
    assert.equal(count(undefined, 'n', 5, 1, 10), 5);
    assert.equal(count(null, 'n', 5, 1, 10), 5);
    assert.equal(count(7.9, 'n', 5, 1, 10), 7);
    assert.equal(count(0, 'n', 5, 1, 10), 1);
    assert.equal(count(99, 'n', 5, 1, 10), 10);
});

test('count: anything that is not a finite number is refused', () => {
    for (const value of ['7', '7; evil', NaN, Infinity, {}, [3]]) {
        assert.throws(() => count(value, 'n', 5, 1, 10), invalid, String(value));
    }
});

test('flag and requiredText refuse the wrong type', () => {
    assert.equal(flag(undefined, 'f', true), true);
    assert.throws(() => flag('true', 'f', false), invalid);
    assert.equal(requiredText('  a@b.com ', 'account'), 'a@b.com');
    assert.throws(() => requiredText('   ', 'account'), invalid);
    assert.throws(() => requiredText(42, 'account'), invalid);
});

test('entryIdList trims, deduplicates and keeps order', () => {
    assert.deepEqual(entryIdList([' b', 'a', 'b ']), ['b', 'a']);
    assert.throws(() => entryIdList(['a', '']), invalid);
    assert.throws(() => entryIdList('a'), invalid);
});

test('emailLocator accepts an id or a row, dropping a blank store id', () => {
    assert.deepEqual(emailLocator('E1'), {entryId: 'E1'});
    assert.deepEqual(emailLocator({entryId: 'E1', storeId: '', subject: 'x'}), {entryId: 'E1'});
    assert.deepEqual(emailLocator({entryId: 'E1', storeId: 'S1'}), {entryId: 'E1', storeId: 'S1'});
    assert.throws(() => emailLocator({storeId: 'S1'}), invalid);
    assert.throws(() => emailLocator(null), invalid);
});

test('recipientList splits on commas and semicolons outside quotes and angle brackets', () => {
    assert.deepEqual(
        recipientList('"Doe, Jo" <jo@x.com>; a@x.com,b@x.com ;; ', 'to'),
        ['"Doe, Jo" <jo@x.com>', 'a@x.com', 'b@x.com'],
    );
    assert.deepEqual(recipientList(['a@x.com', 'b@x.com; c@x.com'], 'to'), ['a@x.com', 'b@x.com', 'c@x.com']);
    assert.deepEqual(recipientList(undefined, 'to'), []);
    assert.throws(() => recipientList([42], 'to'), invalid);
});

test('parseRecipient separates a display name from its address', () => {
    assert.deepEqual(parseRecipient('"Doe, Jo" <jo@x.com>'), {name: 'Doe, Jo', address: 'jo@x.com'});
    assert.deepEqual(parseRecipient('Jo Doe <jo@x.com>'), {name: 'Jo Doe', address: 'jo@x.com'});
    assert.deepEqual(parseRecipient('jo@x.com'), {name: '', address: 'jo@x.com'});
});

test('findAttachmentIndex: exact first, then ignoring case and whitespace', () => {
    const names = ['Invoice.pdf', 'invoice.pdf', 'Rate\u00A0 Sheet.xlsx'];
    assert.equal(findAttachmentIndex(names, 'invoice.pdf'), 1);
    assert.equal(findAttachmentIndex(names, 'INVOICE.PDF'), 0);
    assert.equal(findAttachmentIndex(names, ' rate sheet.xlsx '), 2);
    assert.equal(findAttachmentIndex(names, 'missing.pdf'), -1);
});

test('safeFileName keeps a hostile name inside its directory', () => {
    assert.equal(safeFileName('..\\..\\evil.exe'), '.._.._evil.exe');
    assert.equal(safeFileName('../../etc/passwd'), '.._.._etc_passwd');
    assert.equal(safeFileName('a:b*c?"d<e>f|g.txt'), 'a_b_c__d_e_f_g.txt');
    assert.equal(safeFileName('report.pdf. . '), 'report.pdf');
    assert.equal(safeFileName('CON.txt'), '_CON.txt');
    assert.equal(safeFileName('...'), 'attachment');
    assert.equal(safeFileName(`${'x'.repeat(300)}.pdf`), `${'x'.repeat(150)}.pdf`);
});

test('uniqueSavePath never reuses a taken name', () => {
    const taken = new Set([path.join('d', 'a.pdf'), path.join('d', 'a (1).pdf')]);
    assert.equal(uniqueSavePath('d', 'a.pdf', file => taken.has(file)), path.join('d', 'a (2).pdf'));
    assert.equal(uniqueSavePath('d', 'b.pdf', file => taken.has(file)), path.join('d', 'b.pdf'));
});
