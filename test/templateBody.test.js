// The template toolkit is pure string work, so it is testable without Outlook —
// and it is the part most likely to break silently, since a regression shows up
// as a recipient getting an email with "[[ACCEPT]]" in it.
const test = require('node:test');
const assert = require('node:assert/strict');

const {
    findTemplateMarkers,
    findTokens,
    findUnfilledTokens,
    replaceToken,
    removeTokenLine,
    composeTemplateBody,
} = require('../dist/templateBody.js');
const {OutlookError} = require('../dist/errors.js');

const TEMPLATE = [
    '<p>Hello,</p>',
    '<p>[[ACCEPT]]Thanks — accepted.[[/ACCEPT]]</p>',
    '<p>[[DECLINE]]Thanks for letting us know.[[/DECLINE]]</p>',
    '<p>[[CLARIFY]]{{QUESTIONS}}[[/CLARIFY]]</p>',
    '<p>Best regards,</p>',
    '<p>{{SIGNATURE}}</p>',
].join('\n');

test('findTemplateMarkers lists sections and placeholders', () => {
    const found = findTemplateMarkers(TEMPLATE);
    assert.deepEqual(found.sections, ['ACCEPT', 'DECLINE', 'CLARIFY']);
    assert.deepEqual(found.placeholders, ['QUESTIONS', 'SIGNATURE']);
});

test('composeTemplateBody keeps one section and drops the others', () => {
    const out = composeTemplateBody(TEMPLATE, {section: 'ACCEPT', placeholders: {SIGNATURE: '<i>Jo</i>'}});
    assert.match(out, /Thanks — accepted\./);
    assert.doesNotMatch(out, /letting us know/);
    assert.doesNotMatch(out, /\{\{QUESTIONS\}\}/);
    assert.match(out, /<i>Jo<\/i>/);
    assert.doesNotMatch(out, /\[\[|\]\]|\{\{|\}\}/, 'no marker survives into a body about to be sent');
});

test('composeTemplateBody matches the section and placeholder names case-insensitively', () => {
    const out = composeTemplateBody(TEMPLATE, {section: 'accept', placeholders: {signature: 'Jo'}});
    assert.match(out, /Thanks — accepted\./);
});

test('composeTemplateBody survives Word splitting a marker across tags', () => {
    const mangled = '<p>Hi</p><p>[<span>[ACC</span>EPT]]Priced.[[/ACC<span>EPT]]</span></p>';
    const out = composeTemplateBody(mangled, {section: 'ACCEPT'});
    assert.match(out, /Priced\./);
    assert.doesNotMatch(out, /ACCEPT/);
});

test('composeTemplateBody refuses a section the template does not have', () => {
    assert.throws(() => composeTemplateBody(TEMPLATE, {section: 'NOPE'}), /has no \[\[NOPE\]\] section.*ACCEPT, DECLINE, CLARIFY/s);
});

test('composeTemplateBody refuses to send a sectioned template with no section named', () => {
    assert.throws(() => composeTemplateBody(TEMPLATE, {}), /split into sections.*name the one to send/s);
});

test('a template failure is an INVALID_REQUEST OutlookError', () => {
    assert.throws(() => composeTemplateBody(TEMPLATE, {}), error => {
        assert.ok(error instanceof OutlookError);
        assert.equal(error.code, 'INVALID_REQUEST');
        return true;
    });
});

test('composeTemplateBody refuses a placeholder the template lacks', () => {
    assert.throws(() => composeTemplateBody(TEMPLATE, {
        section: 'ACCEPT',
        placeholders: {NOPE: 'x'}
    }), /has no \{\{NOPE\}\} placeholder/);
});

test('composeTemplateBody accepts a placeholder that lived in a dropped section', () => {
    const out = composeTemplateBody(TEMPLATE, {
        section: 'ACCEPT',
        placeholders: {QUESTIONS: 'ignored', SIGNATURE: 'Jo'}
    });
    assert.doesNotMatch(out, /ignored/);
});

test('composeTemplateBody refuses a body left holding unresolved markers', () => {
    assert.throws(() => composeTemplateBody(TEMPLATE, {section: 'CLARIFY'}), /unresolved markers.*QUESTIONS/s);
});

test('composeTemplateBody rejects a malformed section', () => {
    assert.throws(() => composeTemplateBody('<p>[[/ACCEPT]]body[[ACCEPT]]</p>', {section: 'ACCEPT'}), /malformed/);
});

test('composeTemplateBody passes a single-body template through unchanged', () => {
    const plain = '<p>Hello,</p><p>Thanks.</p>';
    assert.equal(composeTemplateBody(plain), plain);
});

test('replaceToken substitutes every occurrence and never rescans the value', () => {
    assert.equal(replaceToken('<p>{Ref} and {Ref}</p>', '{Ref}', '{Ref}'), '<p>{Ref} and {Ref}</p>');
    assert.equal(replaceToken('<p>{Ref}</p>', '{Ref}', 'ABC-1'), '<p>ABC-1</p>');
});

test('removeTokenLine takes out the table row a token sits in, and only that row', () => {
    const html = '<table><tr><td>Weight</td><td>{Weight}</td></tr><tr><td>Keep</td></tr></table>';
    const out = removeTokenLine(html, '{Weight}');
    assert.doesNotMatch(out, /Weight/);
    assert.match(out, /Keep/);
});

test('removeTokenLine does not mistake a closed row before a paragraph for its line', () => {
    const html = '<table><tr><td>Keep</td></tr></table><p>Note: {Note}</p>';
    assert.equal(removeTokenLine(html, '{Note}'), '<table><tr><td>Keep</td></tr></table>');
});

test('findTokens reports only the tokens the template actually uses', () => {
    assert.deepEqual(findTokens('<p>{Ref} only</p>', ['{Ref}', '{Weight}']), ['{Ref}']);
});

test('findUnfilledTokens spots a leftover token but ignores CSS braces', () => {
    assert.deepEqual(findUnfilledTokens('<style>p { margin: 0 }</style><p>{Weight} missing</p>'), ['{Weight}']);
});
