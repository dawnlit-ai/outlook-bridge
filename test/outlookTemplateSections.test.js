// The template toolkit is pure string work, so it is testable without Outlook —
// and it is the part most likely to break silently, because a regression here
// shows up as a customer receiving an email with "[[QUOTE]]" in it.
const test = require('node:test');
const assert = require('node:assert/strict');

const {
    findTemplateMarkers,
    findTokens,
    findUnfilledTokens,
    replaceToken,
    removeTokenLine,
    composeTemplateBody,
} = require('../dist/outlookTemplateSections.js');
const {OutlookError} = require('../dist/errors.js');

const TEMPLATE = [
    '<p>Hello,</p>',
    '<p>[[QUOTE]]Thanks for your quote.[[/QUOTE]]</p>',
    '<p>[[NOQUOTE]]Thanks for letting us know.[[/NOQUOTE]]</p>',
    '<p>[[CLARIFY]]{{QUESTIONS}}[[/CLARIFY]]</p>',
    '<p>Best regards,</p>',
    '<p>{{SIGNATURE}}</p>',
].join('\n');

test('findTemplateMarkers lists sections and placeholders', () => {
    const found = findTemplateMarkers(TEMPLATE);
    assert.deepEqual(found.sections, ['QUOTE', 'NOQUOTE', 'CLARIFY']);
    assert.deepEqual(found.placeholders, ['QUESTIONS', 'SIGNATURE']);
});

test('composeTemplateBody keeps one section and drops the others', () => {
    const out = composeTemplateBody(TEMPLATE, {
        section: 'QUOTE',
        placeholders: {SIGNATURE: '<i>Jo</i>'},
    });
    assert.match(out, /Thanks for your quote\./);
    assert.doesNotMatch(out, /letting us know/);
    assert.doesNotMatch(out, /\{\{QUESTIONS\}\}/);
    assert.match(out, /<i>Jo<\/i>/);
    // No marker of any kind survives into a body that is about to be mailed.
    assert.doesNotMatch(out, /\[\[|\]\]|\{\{|\}\}/);
});

test('composeTemplateBody matches the section name case-insensitively', () => {
    const out = composeTemplateBody(TEMPLATE, {
        section: 'quote',
        placeholders: {signature: 'Jo'},
    });
    assert.match(out, /Thanks for your quote\./);
});

test('composeTemplateBody survives Word splitting a marker across tags', () => {
    // Word breaks runs mid-word across <span>s and wraps long lines; markers are
    // matched character by character precisely so this still resolves.
    const mangled = '<p>Hi</p><p>[<span>[QU</span>OTE]]Priced.[[/QU<span>OTE]]</span></p>';
    const out = composeTemplateBody(mangled, {section: 'QUOTE'});
    assert.match(out, /Priced\./);
    assert.doesNotMatch(out, /QUOTE/);
});

test('composeTemplateBody refuses a section the template does not have', () => {
    assert.throws(
        () => composeTemplateBody(TEMPLATE, {section: 'NOPE'}),
        /has no \[\[NOPE\]\] section.*QUOTE, NOQUOTE, CLARIFY/s,
    );
});

test('composeTemplateBody refuses to send a sectioned template unnamed', () => {
    // Otherwise every variant plus the raw markers would go to the recipient.
    // The message names `templateSection` — the parameter this package exposes —
    // rather than the tool-layer spelling it was extracted from.
    assert.throws(() => composeTemplateBody(TEMPLATE, {}), /name one with templateSection/);
});

test('a template failure is an OutlookError, like every other deliberate one', () => {
    // These reach callers through replyOutlookEmail, so a bare Error would land
    // them in the "bug" branch of the `err instanceof OutlookError` split that
    // errors.ts tells consumers to write.
    assert.throws(() => composeTemplateBody(TEMPLATE, {}), (error) => {
        assert.ok(error instanceof OutlookError);
        assert.equal(error.code, 'INVALID_REQUEST');
        return true;
    });
});

test('composeTemplateBody refuses a placeholder the template lacks', () => {
    assert.throws(
        () => composeTemplateBody(TEMPLATE, {section: 'QUOTE', placeholders: {NOPE: 'x'}}),
        /has no \{\{NOPE\}\} placeholder/,
    );
});

test('composeTemplateBody accepts a placeholder that lived in a dropped section', () => {
    // Over-supplying is the caller being generic, not an error.
    const out = composeTemplateBody(TEMPLATE, {
        section: 'QUOTE',
        placeholders: {QUESTIONS: 'ignored', SIGNATURE: 'Jo'},
    });
    assert.doesNotMatch(out, /ignored/);
});

test('composeTemplateBody refuses a body left holding unresolved markers', () => {
    assert.throws(
        () => composeTemplateBody(TEMPLATE, {section: 'CLARIFY'}),
        /unresolved markers.*QUESTIONS/s,
    );
});

test('composeTemplateBody rejects a malformed section', () => {
    const broken = '<p>[[/QUOTE]]body[[QUOTE]]</p>';
    assert.throws(() => composeTemplateBody(broken, {section: 'QUOTE'}), /malformed/);
});

test('composeTemplateBody passes a single-body template through unchanged', () => {
    const plain = '<p>Hello,</p><p>Thanks.</p>';
    assert.equal(composeTemplateBody(plain), plain);
});

test('replaceToken substitutes every occurrence and never rescans the value', () => {
    const out = replaceToken('<p>{Ref} and {Ref}</p>', '{Ref}', '{Ref}');
    assert.equal(out, '<p>{Ref} and {Ref}</p>');
    assert.equal(replaceToken('<p>{Ref}</p>', '{Ref}', 'ABC-1'), '<p>ABC-1</p>');
});

test('removeTokenLine takes out the table row a token sits in', () => {
    const html = '<table><tr><td>Weight</td><td>{Weight}</td></tr><tr><td>Keep</td></tr></table>';
    const out = removeTokenLine(html, '{Weight}');
    assert.doesNotMatch(out, /Weight/);
    assert.match(out, /Keep/);
});

test('findTokens reports only the tokens the template actually uses', () => {
    assert.deepEqual(
        findTokens('<p>{Ref} only</p>', ['{Ref}', '{Weight}']),
        ['{Ref}'],
    );
});

test('findUnfilledTokens spots a leftover token but ignores CSS braces', () => {
    const html = '<style>p { margin: 0 }</style><p>{Weight} missing</p>';
    assert.deepEqual(findUnfilledTokens(html), ['{Weight}']);
});
