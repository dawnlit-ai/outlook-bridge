// Shared test helpers. Not a test file itself (node --test runs *.test.js).
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

/** Every method a backend implements, with a plausible empty answer. */
const BACKEND_DEFAULTS = {
    getOutlookAccounts: () => [],
    sendOutlookEmail: () => undefined,
    replyOutlookEmail: () => ({to: '', subject: '', repliedToSender: ''}),
    readInboxEmails: () => [],
    searchInboxByFilter: () => [],
    readSelectedEmail: () => ({
        entryId: '1',
        storeId: '',
        subject: '',
        senderName: '',
        senderEmail: '',
        receivedTime: '',
        body: '',
        attachmentNames: []
    }),
    readEmailBody: () => ({
        entryId: '1',
        subject: '',
        senderName: '',
        senderEmail: '',
        receivedTime: '',
        body: '',
        attachmentNames: []
    }),
    openOutlookEmail: () => undefined,
    listInboxFolders: () => [],
    moveOutlookEmails: () => ({folderPath: '', folderCreated: false, moved: 0, failed: []}),
    listOutlookDrafts: () => ({account: '', foldersScanned: [], count: 0, truncated: false, drafts: []}),
    sendDrafts: () => ({sent: 0, failed: []}),
    sendAllDrafts: () => ({sent: 0, failed: []}),
    deleteOutlookDrafts: () => ({deleted: 0, failed: []}),
    deleteOutlookEmails: () => ({dryRun: true, deleted: 0, refused: 0, failed: 0, items: []}),
    purgeDeletedItems: () => ({folderPath: '', dryRun: true, matched: 0, purged: 0, kept: 0, failed: 0}),
    saveEmailAttachments: request => request.fileNames.map(fileName => ({
        path: path.join(request.destDir, fileName),
        fileName,
        subject: '',
        senderName: '',
        senderEmail: '',
        receivedTime: '',
    })),
    cleanUndeliverableEmails: request => ({
        account: request.account, scannedDays: request.daysBack, dryRun: request.dryRun,
        matchedCount: 0, deletedCount: 0, matched: [], failed: [],
    }),
    collectBouncedRecipients: () => [],
    readSentRecipientGroups: () => [],
    listOutlookSignatures: () => [],
    readOutlookSignatureHtml: () => '',
    readTemplateEmails: () => ({folderFound: true, folderPath: '', templates: [], availableFolders: []}),
    saveTemplateEmail: () => ({folderPath: '', folderCreated: false}),
    editEmailTemplate: request => request.html,
};

/**
 * A backend that records every request it receives. `overrides` replaces any
 * method's answer; `calls` lists `{ name, args }` in order.
 */
function fakeBackend(overrides = {}) {
    const calls = [];
    const backend = {};
    for (const [name, answer] of Object.entries({...BACKEND_DEFAULTS, ...overrides})) {
        backend[name] = async (...args) => {
            calls.push({name, args});
            return answer(...args);
        };
    }
    return {backend, calls, last: name => [...calls].reverse().find(call => call.name === name)};
}

/**
 * Run `operation` with a module's runner functions replaced, returning every
 * script body the operation handed them and the error it ended with, if any.
 * `responses` are returned in order, so a multi-pass operation can be driven
 * into its later passes.
 */
async function captureScripts(runModule, runnerNames, operation, responses = []) {
    const captured = [];
    const originals = runnerNames.map(name => [name, runModule[name]]);
    let call = 0;
    for (const name of runnerNames) {
        runModule[name] = async body => {
            captured.push(body);
            const response = responses[call++];
            return response === undefined ? (name.endsWith('Json') ? null : '') : response;
        };
    }
    let error;
    try {
        await operation();
    } catch (caught) {
        error = caught;
    } finally {
        for (const [name, original] of originals) runModule[name] = original;
    }
    return {scripts: captured, error};
}

/** A file that exists, for operations that check attachment paths. */
function tempFileWith(name, contents = 'x') {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ob-test-'));
    const file = path.join(dir, name);
    fs.writeFileSync(file, contents);
    return file;
}

module.exports = {fakeBackend, captureScripts, tempFileWith};
