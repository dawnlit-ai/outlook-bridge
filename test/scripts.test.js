// The script fragments themselves.
//
// Neither platform's automation can run in CI, but its scripts can be checked.
// The PowerShell fragments are checked for shape and escaping here (and every
// whole script is parsed in windowsScripts.test.js); on a Mac with Outlook, the
// AppleScript fragments go through osacompile, which resolves dictionary terms
// without opening a session or touching a mailbox.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');

const {accountScript, itemLookupScript, mailScopeScript} = require('../dist/windows/scripts.js');
const {buildScript, psEscape, psInt} = require('../dist/windows/run.js');
const {mailFolderRef} = require('../dist/mail.js');
const {errorFromTaggedMessage} = require('../dist/errors.js');
const macScripts = require('../dist/mac/scripts.js');
const macRun = require('../dist/mac/run.js');

// ── PowerShell ───────────────────────────────────────────────────────────

test('psEscape doubles every quote PowerShell accepts as a single quote', () => {
    assert.equal(psEscape("o'brien"), "o''brien");
    assert.equal(psEscape('\u2018a\u2019 \u201Ab\u201B'), '\u2018\u2018a\u2019\u2019 \u201A\u201Ab\u201B\u201B');
    assert.equal(psEscape('$(Get-Date) "x"'), '$(Get-Date) "x"', 'nothing else is special in a single-quoted literal');
});

test('psInt refuses anything but a whole number', () => {
    assert.equal(psInt(42), '42');
    for (const bad of [1.5, NaN, '7', Infinity]) assert.throws(() => psInt(bad), error => error.code === 'INVALID_REQUEST');
});

test('every script runs inside a catch that reports one JSON failure line', () => {
    const script = buildScript('$x = 1');
    assert.match(script, /\$ErrorActionPreference = 'Stop'/);
    assert.match(script, /try \{\s*\$x = 1\s*\} catch \{/);
    assert.match(script, /ConvertTo-Json -Compress @\{ message = \$message; line = \$line \}/);
});

test('a mailbox resolves through Accounts or through store folders', () => {
    const script = accountScript('team@example.com');
    assert.match(script, /\$a\.SmtpAddress -ieq \$target/);
    assert.match(script, /\$f\.Name -ieq \$target/);
    assert.match(script, /if \(\$account -eq \$null -and \$storeFolder -eq \$null\)/);
    // The address is matched before the account's display name.
    assert.ok(script.indexOf('$f.Name -ieq $target') < script.indexOf('$f.Name -ieq $account.DisplayName'));
});

test('a missing account fails with a tag the classifier reads as ACCOUNT_NOT_FOUND', () => {
    const thrown = /throw "([^"]*)"/.exec(accountScript('nobody@example.com'))[1].replace('$target', 'nobody@example.com');
    const error = errorFromTaggedMessage(thrown);
    assert.equal(error.code, 'ACCOUNT_NOT_FOUND');
    assert.equal(error.account, 'nobody@example.com');
});

test('caller text only ever appears in a single-quoted literal', () => {
    assert.match(accountScript("o'brien@example.com"), /\$target = 'o''brien@example\.com'/);
    const lookup = itemLookupScript({entryId: "$(evil)'", storeId: 'S'});
    assert.match(lookup, /\$lookupId = '\$\(evil\)'''/);
    assert.doesNotMatch(lookup, /"[^"]*\$\(evil\)/, 'never inside a double-quoted string');
});

test('a bare well-known root needs no walk', () => {
    const script = mailScopeScript(mailFolderRef('Sent Items'), 'Sent Items');
    assert.match(script, /\$scope = \$store\.GetDefaultFolder\(5\)/);
    assert.match(script, /\$scopeCreated = \$false/);
    assert.doesNotMatch(script, /Find-FolderByName/);
});

test('a folder path walks its segments and names the caller string on failure, by variable', () => {
    const script = mailScopeScript(mailFolderRef('Inbox\\Clients\\$(evil)'), 'Inbox\\Clients\\$(evil)');
    assert.match(script, /\$segments = @\('Clients', '\$\(evil\)'\)/);
    assert.match(script, /\$folderLabel = 'Inbox\\Clients\\\$\(evil\)'/);
    assert.match(script, /Folder '\$folderLabel' not found/);
    assert.doesNotMatch(script, /throw "[^"]*\$\(evil\)/);
});

test('createMissing rebuilds the whole chain rather than only the leaf', () => {
    const ref = mailFolderRef('Clients\\Acme\\2026');
    assert.match(mailScopeScript(ref, 'x', true), /if \(\$scope -eq \$null -and \$true\)/);
    assert.match(mailScopeScript(ref, 'x', true), /\$scope\.Folders\.Add\(\$seg\)/);
    assert.match(mailScopeScript(ref, 'x', false), /if \(\$scope -eq \$null -and \$false\)/);
});

// ── AppleScript ──────────────────────────────────────────────────────────

const OUTLOOK_APP = '/Applications/Microsoft Outlook.app';
const canCompile = process.platform === 'darwin' && fs.existsSync(OUTLOOK_APP);
const macOnly = {skip: !canCompile && 'needs macOS with Outlook installed'};

/** Compile (never run) a script, returning osacompile's complaint or ''. */
function compileError(source) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-osa-'));
    const input = path.join(dir, 'fragment.applescript');
    fs.writeFileSync(input, source, 'utf-8');
    try {
        execFileSync('osacompile', ['-o', path.join(dir, 'fragment.scpt'), input], {stdio: 'pipe'});
        return '';
    } catch (error) {
        return String(error.stderr || error.message);
    } finally {
        fs.rmSync(dir, {recursive: true, force: true});
    }
}

/** A snippet inside the tell block its terms resolve against. */
function inTell(body) {
    return `${macRun.AS_HANDLERS}
tell application "Microsoft Outlook"
${body}
end tell`;
}

/** An account only the AppleScript probe can reach. */
const PROBED = {emailAccount: 'someone@example.com'};

/** An account the profile database also describes, by folder id. */
const PROFILED = {
    emailAccount: 'someone@example.com',
    folderIds: {
        'inbox': 148,
        'sent items': 149,
        'deleted items': 146,
        'drafts': 147,
        'junk mail': 150,
        'root folder': 142
    },
};

test('the shared AppleScript handlers compile', macOnly, () => {
    assert.equal(compileError(macRun.AS_HANDLERS), '');
    assert.equal(compileError(macScripts.FIND_FOLDER_HANDLER), '');
    assert.equal(compileError(macScripts.LIST_ACCOUNTS_SNIPPET), '');
});

test('the account and message lookups compile', macOnly, () => {
    assert.equal(compileError(inTell(macScripts.accountLookupSnippet(PROBED))), '');
    assert.equal(compileError(inTell(macScripts.accountLookupSnippet(PROFILED))), '');
    assert.equal(compileError(inTell(macScripts.accountLookupSnippet(PROFILED, true))), '');
    assert.equal(compileError(inTell(macScripts.messageLookupSnippet('123'))), '');
});

test('only a genuinely absent account raises ACCOUNT_NOT_FOUND', () => {
    const tagged = snippet => {
        const literal = /error "((?:[^"\\]|\\.)*)"/.exec(snippet);
        return literal && errorFromTaggedMessage(literal[1]);
    };
    assert.equal(tagged(macScripts.accountLookupSnippet(PROBED)).code, 'ACCOUNT_NOT_FOUND');
    assert.equal(tagged(macScripts.accountLookupSnippet(PROFILED)), null, 'a profiled mailbox is found by folder id');
    assert.equal(tagged(macScripts.accountLookupSnippet(PROFILED, true)).code, 'INVALID_REQUEST', 'found, but cannot compose');
});

test('a well-known root resolves by account, by id, or names the folder out of reach', macOnly, () => {
    for (const term of Object.values(macScripts.MAC_ROOT_TERMS)) {
        for (const acct of [PROBED, PROFILED]) {
            const source = inTell(`    set targetAcct to item 1 of imap accounts
${macScripts.rootFolderSnippet(acct, term, 'f')}`);
            assert.equal(compileError(source), '', term);
        }
    }
});

test('a root the profile has no id for is reported as a missing folder', () => {
    const partial = {emailAccount: 'someone@example.com', folderIds: {'inbox': 148}};
    assert.match(macScripts.rootFolderSnippet(partial, 'drafts', 'f'), /NOT_FOUND:folder.*has no drafts folder/);
    assert.match(macScripts.rootFolderSnippet(PROFILED, 'drafts', 'f'), /mail folder id 147/);
    assert.doesNotMatch(macScripts.rootFolderSnippet(PROBED, 'drafts', 'f'), /mail folder id/);
});

test('the per-message field snippets compile', macOnly, () => {
    for (const snippet of [macScripts.senderSnippet(), macScripts.firstRecipientSnippet(), macScripts.allRecipientsSnippet()]) {
        assert.equal(compileError(inTell(`    set theMsg to missing value\n${snippet}`)), '');
    }
});

test('a folder scope compiles for every root, walked and created', macOnly, () => {
    for (const acct of [PROBED, PROFILED]) {
        for (const rootId of Object.keys(macScripts.MAC_ROOT_TERMS)) {
            const ref = {rootId: Number(rootId), rootLabel: 'Root', segments: []};
            assert.equal(compileError(inTell(`    set targetAcct to item 1 of imap accounts
${macScripts.mailScopeSnippet(acct, ref)}`)), '', `root ${rootId}`);
        }
        const nested = {rootId: 6, rootLabel: 'Inbox', segments: ['Clients', 'Acme']};
        for (const create of [false, true]) {
            assert.equal(compileError(inTell(`    set targetAcct to item 1 of imap accounts
${macScripts.mailScopeSnippet(acct, nested, 'Inbox\\Clients\\Acme', create)}`)), '');
        }
    }
});

test('an unsupported root is refused in TypeScript, not by a broken script', () => {
    assert.throws(
        () => macScripts.mailScopeSnippet(PROBED, {rootId: 99, rootLabel: 'Journal', segments: []}, 'Journal'),
        error => error.code === 'NOT_IMPLEMENTED' && /Journal/.test(error.message),
    );
});

test('asEscape keeps caller text inside its AppleScript literal', () => {
    assert.equal(macRun.asEscape('He said "hi" \\ then\nleft'), 'He said \\"hi\\" \\\\ then\\nleft');
});

test('the framing survives a body that contains the separators', () => {
    const raw = ['one', 'two', ['a', 'b'].join(macRun.LIST_SEP)].join(macRun.FIELD_SEP) + macRun.RECORD_SEP;
    const records = macRun.splitRecords(raw);
    assert.equal(records.length, 1);
    const fields = macRun.splitFields(records[0]);
    assert.deepEqual(fields.slice(0, 2), ['one', 'two']);
    assert.deepEqual(macRun.splitList(fields[2]), ['a', 'b']);
});
