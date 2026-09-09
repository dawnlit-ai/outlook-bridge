// The generated scripts themselves.
//
// Neither platform's scripts can be run in CI, but they can be checked, and the
// two failure modes worth catching are exactly the ones a live run reports
// badly. A PowerShell folder walk that emits the wrong shape only shows up as a
// failed run against a real mailbox; an AppleScript with one bad dictionary term
// fails the WHOLE script at compile time rather than at the offending line, so
// on a Mac with Outlook installed every fragment is put through osacompile here
// — no Outlook session, no mailbox touched, just terminology resolution.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');

const {accountScript, mailScopeScript} = require('../dist/windows/index.js');
const {mailFolderRef} = require('../dist/mail.js');
const macScripts = require('../dist/mac/scripts.js');
const macRun = require('../dist/mac/run.js');

// -- The Windows mailbox resolver --------------------------------------

// A mailbox open as a secondary store has no Account object. Resolving through
// Accounts alone made every operation unreachable for it, so both routes are
// tried and either one is enough.
test('a mailbox resolves through Accounts or through Stores', () => {
    const script = accountScript('team@example.com');
    assert.match(script, /\$a\.SmtpAddress -ieq \$target/);
    assert.match(script, /\$f\.Name -ieq \$target/);
    assert.match(script, /if \(\$account -eq \$null -and \$storeFolder -eq \$null\)/);
});

// classifyRunFailure recognises this exact sentence to raise AccountNotFoundError.
// Reword it and the most common operational mistake degrades to SCRIPT_FAILED.
test('the not-found sentence stays the one the error classifier matches', () => {
    assert.match(accountScript('nobody@example.com'), /throw "Account '\$target' not found"/);
});

// The address is matched before the account's display name, which is what keeps
// the named-store route selecting the folders it selected before on a profile
// where the two differ.
test('the target address is matched before the account display name', () => {
    const script = accountScript('team@example.com');
    const byTarget = script.indexOf('$f.Name -ieq $target');
    const byDisplay = script.indexOf('$f.Name -ieq $account.DisplayName');
    assert.ok(byTarget > -1 && byDisplay > -1);
    assert.ok(byTarget < byDisplay, 'target match must come first');
});

// Caller text goes into a single-quoted PowerShell literal; a quote in it would
// otherwise close the string and run whatever followed.
test('the target address is escaped into its literal', () => {
    assert.match(accountScript("o'brien@example.com"), /\$target = 'o''brien@example\.com'/);
});


// ── The Windows folder-scope emitter ─────────────────────────────────────
test('a bare well-known root needs no walk', () => {
    const script = mailScopeScript(mailFolderRef('Sent Items'));
    assert.match(script, /\$scope = \$store\.GetDefaultFolder\(5\)/);
    assert.match(script, /\$scopeCreated = \$false/);
    assert.doesNotMatch(script, /Find-FolderByName/, 'no recursive search is needed');
});

test('a path under a root walks its segments and names the caller string on failure', () => {
    const script = mailScopeScript(mailFolderRef('Inbox\\Clients\\Acme'), 'Inbox\\Clients\\Acme');
    assert.match(script, /\$segments = @\('Clients','Acme'\)/);
    assert.match(script, /GetDefaultFolder\(6\)/);
    assert.match(script, /Folder 'Inbox\\Clients\\Acme' not found/);
});

test('createMissing rebuilds the whole chain rather than only the leaf', () => {
    const withCreate = mailScopeScript(mailFolderRef('Clients\\Acme\\2026'), 'Clients\\Acme\\2026', true);
    // The creation branch is emitted either way and gated on a literal, so what
    // distinguishes the two calls is the gate — not the presence of Folders.Add.
    assert.match(withCreate, /if \(\$scope -eq \$null -and \$true\)/);
    assert.match(withCreate, /foreach \(\$seg in \$segments\)[\s\S]*\$scope\.Folders\.Add\(\$seg\)/);
    const without = mailScopeScript(mailFolderRef('Clients\\Acme\\2026'), 'Clients\\Acme\\2026', false);
    assert.match(without, /if \(\$scope -eq \$null -and \$false\)/);
});

test('a folder name carrying an apostrophe stays inside its literal', () => {
    // The single-quote doubling is the whole defence for caller text in a
    // generated script; a name like "Bob's mail" must not close the literal.
    const script = mailScopeScript(mailFolderRef("Inbox\\Bob's mail"), "Inbox\\Bob's mail");
    assert.match(script, /'Bob''s mail'/);
});

// ── The macOS fragments, compiled ────────────────────────────────────────
const OUTLOOK_APP = '/Applications/Microsoft Outlook.app';
const canCompile = process.platform === 'darwin' && fs.existsSync(OUTLOOK_APP);

/** Compile (never run) a script, returning osacompile's complaint or ''. */
function compileError(source) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-osa-'));
    const input = path.join(dir, 'fragment.applescript');
    const output = path.join(dir, 'fragment.scpt');
    fs.writeFileSync(input, source, 'utf-8');
    try {
        execFileSync('osacompile', ['-o', output, input], {stdio: 'pipe'});
        return '';
    } catch (error) {
        return String(error.stderr || error.message);
    } finally {
        fs.rmSync(dir, {recursive: true, force: true});
    }
}

/** Wrap a snippet in the tell block its terms are resolved against. */
function inTell(body) {
    return `${macRun.AS_HANDLERS}
tell application "Microsoft Outlook"
${body}
end tell`;
}

test('the shared AppleScript handlers compile', {skip: !canCompile && 'needs macOS with Outlook installed'}, () => {
    assert.equal(compileError(macRun.AS_HANDLERS), '');
    assert.equal(compileError(macScripts.FIND_FOLDER_HANDLER), '');
    assert.equal(compileError(macScripts.LIST_ACCOUNTS_SNIPPET), '');
});

/** An account only the AppleScript probe can reach — no profile folder ids. */
const PROBED = {emailAccount: 'someone@example.com'};

/**
 * An account the profile database also describes, which is the only handle on a
 * mailbox Outlook publishes no account object for. Every root is present, since
 * that is what the profile reports for a mailbox in normal shape.
 */
const PROFILED = {
    emailAccount: 'someone@example.com',
    folderIds: {
        'inbox': 148,
        'sent items': 149,
        'deleted items': 146,
        'drafts': 147,
        'junk mail': 150,
        'root folder': 142,
    },
};

test('the account and message lookups compile', {skip: !canCompile && 'needs macOS with Outlook installed'}, () => {
    assert.equal(compileError(inTell(macScripts.accountLookupSnippet(PROBED))), '');
    assert.equal(compileError(inTell(macScripts.accountLookupSnippet(PROFILED))), '');
    assert.equal(compileError(inTell(macScripts.accountLookupSnippet(PROFILED, true))), '');
    assert.equal(compileError(inTell(macScripts.messageLookupSnippet('123'))), '');
});

// classifyRunFailure recognises this exact sentence to raise AccountNotFoundError,
// and only a mailbox that genuinely isn't there should produce it. One the profile
// can reach by folder id is found — it just can't be composed from, which is a
// different failure and says so.
test('only a genuinely absent account emits the not-found sentence', () => {
    assert.match(macScripts.accountLookupSnippet(PROBED), /error "Account '[^']*' not found"/);
    assert.doesNotMatch(macScripts.accountLookupSnippet(PROFILED), /not found/);
    const composing = macScripts.accountLookupSnippet(PROFILED, true);
    assert.doesNotMatch(composing, /not found/);
    assert.match(composing, /publishes no account object/);
});

test('a well-known root resolves by account, by id, or refuses', {skip: !canCompile && 'needs macOS with Outlook installed'}, () => {
    for (const term of Object.values(macScripts.MAC_ROOT_TERMS)) {
        for (const acct of [PROBED, PROFILED]) {
            const source = inTell(`    set targetAcct to item 1 of imap accounts
${macScripts.rootFolderSnippet(acct, term, 'f')}`);
            assert.equal(compileError(source), '', `${term} (${acct === PROBED ? 'probed' : 'profiled'})`);
        }
    }
    // The probe is the only way in, and it is the one that may have come up
    // empty — so the folder that is out of reach gets named.
    const partial = {emailAccount: 'someone@example.com', folderIds: {'inbox': 148}};
    assert.match(macScripts.rootFolderSnippet(partial, 'drafts', 'f'), /has no drafts folder/);
    assert.match(macScripts.rootFolderSnippet(PROFILED, 'drafts', 'f'), /mail folder id 147/);
    assert.doesNotMatch(macScripts.rootFolderSnippet(PROBED, 'drafts', 'f'), /mail folder id/);
});

test('the per-message field snippets compile', {skip: !canCompile && 'needs macOS with Outlook installed'}, () => {
    for (const snippet of [
        macScripts.senderSnippet(),
        macScripts.firstRecipientSnippet(),
        macScripts.allRecipientsSnippet(),
    ]) {
        assert.equal(compileError(inTell(`    set theMsg to missing value\n${snippet}`)), '');
    }
});

test('every well-known root term Outlook actually accepts', {skip: !canCompile && 'needs macOS with Outlook installed'}, () => {
    // A wrong term here is the expensive kind of mistake: it fails the whole
    // script at compile time, so the error never names the folder that caused it.
    for (const [rootId, term] of Object.entries(macScripts.MAC_ROOT_TERMS)) {
        const source = inTell(`    set targetAcct to item 1 of imap accounts
    set f to ${term} of targetAcct`);
        assert.equal(compileError(source), '', `root ${rootId} (${term})`);
    }
});

test('a folder scope compiles for every root, walked and created', {skip: !canCompile && 'needs macOS with Outlook installed'}, () => {
    for (const acct of [PROBED, PROFILED]) {
        const how = acct === PROBED ? 'probed' : 'profiled';
        for (const rootId of Object.keys(macScripts.MAC_ROOT_TERMS)) {
            const ref = {rootId: Number(rootId), rootLabel: 'Root', segments: []};
            assert.equal(compileError(inTell(`    set targetAcct to item 1 of imap accounts
${macScripts.mailScopeSnippet(acct, ref)}`)), '', `root ${rootId} (${how})`);
        }
        const nested = {rootId: 6, rootLabel: 'Inbox', segments: ['Clients', 'Acme']};
        assert.equal(compileError(inTell(`    set targetAcct to item 1 of imap accounts
${macScripts.mailScopeSnippet(acct, nested, 'Inbox\\Clients\\Acme')}`)), '', how);
        assert.equal(compileError(inTell(`    set targetAcct to item 1 of imap accounts
${macScripts.mailScopeSnippet(acct, nested, 'Inbox\\Clients\\Acme', true)}`)), '', `createMissing (${how})`);
    }
});

test('an unsupported root is refused in TypeScript, not by a broken script', () => {
    // olDefaultFolders ids with no Outlook-for-Mac term must fail as
    // NOT_IMPLEMENTED naming the folder, rather than emitting a script that
    // cannot compile.
    assert.throws(
        () => macScripts.mailScopeSnippet(PROBED, {rootId: 99, rootLabel: 'Journal', segments: []}, 'Journal'),
        (error) => {
            assert.equal(error.code, 'NOT_IMPLEMENTED');
            assert.match(error.message, /Journal/);
            return true;
        },
    );
});

test('an emitted row compiles and round-trips through the splitters', {skip: !canCompile && 'needs macOS with Outlook installed'}, () => {
    const row = macRun.asRow(['"alpha"', '"beta"', 'my joinList({"x", "y"}, ' + macRun.AS_LIST_SEP + ')']);
    assert.equal(compileError(`${macRun.AS_HANDLERS}\nreturn ${row}`), '');
});

test('the framing survives a body that contains the separators', () => {
    // sanitize() strips these inside AppleScript; this pins the TypeScript half
    // of the contract — that a record splits into exactly the fields emitted.
    const raw = ['one', 'two', ['a', 'b'].join(macRun.LIST_SEP)].join(macRun.FIELD_SEP) + macRun.RECORD_SEP;
    const records = macRun.splitRecords(raw);
    assert.equal(records.length, 1);
    const fields = macRun.splitFields(records[0]);
    assert.deepEqual(fields.slice(0, 2), ['one', 'two']);
    assert.deepEqual(macRun.splitList(fields[2]), ['a', 'b']);
});
