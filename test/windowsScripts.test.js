// Every complete PowerShell script the Windows backend generates, parsed.
//
// The counterpart of the macOS compile tests. The runner is stubbed, every
// operation is driven through the public operations with arguments carrying a
// hostile probe — quotes of every kind, a `$(...)` subexpression, a backtick —
// and PowerShell's own parser reads each captured script. Parsing needs no
// Outlook and runs nothing, so this runs on any Windows machine, CI included.
//
// Two things are checked. The script must parse: a syntax error otherwise only
// surfaces as a failed run against a live mailbox. And no double-quoted
// (expandable) string may contain caller text, since PowerShell evaluates
// `$(...)` inside one — every caller value has to sit in a single-quoted literal,
// and the probe has to come back out of one byte for byte.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');

const run = require('../dist/windows/run.js');
const {windowsBackend} = require('../dist/windows/index.js');
const {createOperations} = require('../dist/bridge.js');
const {captureScripts, tempFileWith} = require('./support.js');

const skip = process.platform === 'win32' ? false : 'needs Windows PowerShell';

/** Caller text built to break out of any literal that doesn't escape it properly. */
const PROBE = 'PROBE\u2019 $(Write-Output INJECTED) "dq" \u201Cdq\u201D \'sq\' `tick';
const EMAIL = {entryId: PROBE, storeId: PROBE};

function probeCases() {
    const attachment = tempFileWith('report.pdf');
    const destDir = fs.mkdtempSync(path.join(os.tmpdir(), "ob-probe-\u2019$(x)'-"));
    const escapedProbe = PROBE.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const outlook = createOperations(windowsBackend);
    return {
        getOutlookAccounts: () => outlook.getOutlookAccounts(),
        sendOutlookEmail: () => outlook.sendOutlookEmail({
            emailAccount: PROBE,
            to: [`"${PROBE.replace(/"/g, '')}" <a@example.com>`, 'b@example.com'],
            cc: PROBE,
            bcc: 'c@example.com',
            subject: PROBE,
            htmlBody: `<p>${PROBE}</p>`,
            attachments: [attachment],
            openDraftWindow: false,
        }),
        replyOutlookEmail: () => outlook.replyOutlookEmail({
            emailAccount: PROBE, ...EMAIL, htmlBody: `<p>${PROBE}</p>`, replyAll: true, sendImmediately: true,
        }),
        readInboxEmails: () => outlook.readInboxEmails(PROBE, {
            folder: `Inbox\\${PROBE}`,
            daysBack: 7,
            limit: 5,
            previewChars: 100
        }),
        readInboxEmailsSentItems: () => outlook.readInboxEmails(PROBE, {folder: 'Sent Items', previewChars: 0}),
        searchInboxByFilter: () => outlook.searchInboxByFilter(PROBE, {
            daysBack: 14,
            subjectLike: `*${PROBE}*`,
            subjectPattern: new RegExp(escapedProbe),
            excludeReplies: true,
            requireAttachment: true,
            includeFolders: [PROBE],
            excludeFolders: [`\\\\${PROBE}\\Inbox\\${PROBE}`],
            includeBody: false,
        }),
        readSelectedEmail: () => outlook.readSelectedEmail(),
        readEmailBody: () => outlook.readEmailBody(EMAIL),
        openOutlookEmail: () => outlook.openOutlookEmail(PROBE),
        listInboxFolders: () => outlook.listInboxFolders(PROBE, {maxDepth: 3}),
        moveOutlookEmails: () => outlook.moveOutlookEmails(PROBE, [PROBE, 'other'], `Clients\\${PROBE}`, {createIfMissing: true}),
        listOutlookDrafts: () => outlook.listOutlookDrafts(PROBE, {limit: 10, previewChars: 50}),
        sendDrafts: () => outlook.sendDrafts(PROBE, [PROBE]),
        sendAllDrafts: () => outlook.sendAllDrafts(PROBE),
        deleteOutlookDrafts: () => outlook.deleteOutlookDrafts(PROBE, [PROBE]),
        deleteOutlookEmails: () => outlook.deleteOutlookEmails(PROBE, [PROBE], {allowProtected: true}),
        purgeDeletedItems: () => outlook.purgeDeletedItems(PROBE, {olderThanDays: 30, dryRun: true}),
        saveEmailAttachments: () => outlook.saveEmailAttachments(EMAIL, [PROBE, 'b.pdf'], {destDir}),
        cleanUndeliverableEmails: () => outlook.cleanUndeliverableEmails(PROBE, {daysBack: 10, dryRun: false}),
        collectBouncedRecipients: () => outlook.collectBouncedRecipients(PROBE, {includeDeletedItems: true}),
        readSentRecipientGroups: () => outlook.readSentRecipientGroups(PROBE, {limit: 5}),
        readTemplateEmails: () => outlook.readTemplateEmails(PROBE, {
            folder: PROBE,
            subject: PROBE,
            includeBody: false
        }),
        readTemplateEmailsWithBody: () => outlook.readTemplateEmails(PROBE, {folder: PROBE, includeBody: true}),
        saveTemplateEmail: () => outlook.saveTemplateEmail(PROBE, {
            subject: PROBE,
            htmlBody: '<p>x</p>',
            folder: PROBE
        }),
        editEmailTemplate: () => outlook.editEmailTemplate(PROBE, '<p>x</p>'),
    };
}

/**
 * Parse every script in one PowerShell process, reporting each one's parse
 * errors, its expandable strings, and its single-quoted string values.
 */
function parseScripts(scripts) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-ps-parse-'));
    try {
        const files = scripts.map((script, index) => {
            const file = path.join(dir, `script-${index}.ps1`);
            fs.writeFileSync(file, '\uFEFF' + script, 'utf8');
            return file;
        });
        const listFile = path.join(dir, 'files.json');
        fs.writeFileSync(listFile, JSON.stringify(files), 'utf8');
        const parser = path.join(dir, 'parse.ps1');
        fs.writeFileSync(parser, '\uFEFF' + `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$files = [IO.File]::ReadAllText('${run.psEscape(listFile)}', [Text.Encoding]::UTF8) | ConvertFrom-Json
$results = @()
foreach ($file in $files) {
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors)
    $expandable = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true) | ForEach-Object { $_.Extent.Text })
    $literals = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.StringConstantType -eq 'SingleQuoted' }, $true) | ForEach-Object { $_.Value })
    $results += [PSCustomObject]@{
        errors     = @($errors | ForEach-Object { "$($_.Message) (line $($_.Extent.StartLineNumber))" })
        expandable = $expandable
        literals   = $literals
    }
}
ConvertTo-Json -Depth 4 -Compress -InputObject @($results)
`, 'utf8');
        const output = execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', parser], {
            encoding: 'utf8',
            maxBuffer: 64 * 1024 * 1024,
            windowsHide: true,
        });
        return JSON.parse(output);
    } finally {
        fs.rmSync(dir, {recursive: true, force: true});
    }
}

test('every Windows script parses, and caller text never reaches an expandable string', {skip}, async () => {
    const generated = [];
    for (const [name, operation] of Object.entries(probeCases())) {
        const {scripts, error} = await captureScripts(run, ['runPowerShell', 'runPowerShellJson'], operation);
        // editEmailTemplate finds no saved output when its run is stubbed; any
        // other operation failing means it never built its script.
        if (error && !(name === 'editEmailTemplate' && error.code === 'NOT_FOUND')) throw error;
        assert.ok(scripts.length > 0, `${name} generated no script`);
        for (const body of scripts) generated.push({name, script: run.buildScript(body)});
    }

    const parsed = parseScripts(generated.map(g => g.script));
    parsed.forEach((result, index) => {
        const {name, script} = generated[index];
        assert.deepEqual(result.errors, [], `${name} does not parse`);
        const leaked = (result.expandable ?? []).filter(text => text.includes('PROBE'));
        assert.deepEqual(leaked, [], `${name} put caller text in an expandable string`);
        if (script.includes('PROBE')) {
            assert.ok(
                (result.literals ?? []).some(value => value.includes(PROBE)),
                `${name}: the probe did not survive its single-quoted literal intact`,
            );
        }
    });
});

test('the probe round-trips through psString exactly', {skip}, () => {
    const [result] = parseScripts([`$value = ${run.psString(PROBE)}`]);
    assert.deepEqual(result.errors, []);
    assert.deepEqual(result.literals, [PROBE]);
});
