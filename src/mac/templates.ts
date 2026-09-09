// Template emails: ordinary mail items kept in a mailbox folder and reused as
// reply bodies, plus the compose-window editor for one.
//
// Outlook for Mac serves a message's HTML through `content` when `has html` is
// true — the same thing COM's HTMLBody returns — which is what makes templates
// portable between the two platforms rather than a Windows-only feature.
import {
    asEscape,
    asRow,
    boolField,
    field,
    runOsaScript,
    splitFields,
    splitList,
    splitRecords,
    summaryFields,
} from './run';
import {
    accountLookupSnippet,
    FIND_FOLDER_HANDLER,
    macFolderPath,
    resolveMacAccount,
    rootFolderSnippet,
} from './scripts';
import { findTemplateMarkers } from '../outlookTemplateSections';
import { clamp } from '../mail';
import { NotFoundError } from '../errors';
import type { SaveTemplateResult, TemplateEmail, TemplateFolderResult } from '../types';

/** How deep under the mailbox root a template folder is searched for. */
const TEMPLATE_SEARCH_DEPTH = 3;

/** The cap Windows applies to a template body; a Word-generated one is huge. */
const MAX_TEMPLATE_HTML = 100000;

/** How much plain text stands in for a body when one wasn't asked for. */
const PREVIEW_CHARS = 200;

/**
 * Embedded (cid:) images live as attachments on the TEMPLATE item, so reusing
 * the HTML on a new email shows a broken "linked image" placeholder, and
 * rewriting the refs renders invisibly. Neither works, so the img tags go —
 * template emails are text/HTML only.
 */
function stripEmbeddedImages(html: string): string {
    return html.replace(/<img[^>]*src="cid:[^"]*"[^>]*>/gi, '');
}

/**
 * Read the template emails saved in a mailbox folder (default "Templates"),
 * returning each item's full HTML body. When the folder doesn't exist, returns
 * folderFound:false plus the mailbox's folder names instead of throwing, so the
 * caller can ask the user whether to create it rather than fail.
 *
 * Bodies are read in a second pass, for the chosen templates only. That is not
 * a micro-optimisation: a folder of Word-sized templates is megabytes, and
 * pulling all of them through stdout on a call that wanted one is what overruns
 * the buffer — the same reason the Windows reader filters before reading.
 */
export async function readTemplateEmails(
    emailAccount: string,
    folderName = 'Templates',
    limit = 20,
    includeBody = true,
    subject = '',
): Promise<TemplateFolderResult> {
    const cap = clamp(limit, 1, 50);
    const wanted = (subject || '').trim().toLowerCase();
    const acct = await resolveMacAccount(emailAccount);

    // Pass 1 — find the folder and index it: three bulk reads, no bodies.
    const indexScript = `${FIND_FOLDER_HANDLER}
tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'root folder', 'rootFolder')}
end tell
set theFolder to my findFolderByName(rootFolder, "${asEscape(folderName)}", ${TEMPLATE_SEARCH_DEPTH})
tell application "Microsoft Outlook"
    if theFolder is missing value then
        set names to {}
        repeat with f in (mail folders of rootFolder)
            try
                set end of names to (name of f) as string
            end try
        end repeat
        return ${asRow(['"0"', 'my sanitizeList(names)'])}
    end if
    set idList to id of every message of theFolder
    set subjList to subject of every message of theFolder
    try
        set modList to modification date of every message of theFolder
    on error
        set modList to {}
    end try
    set hasMods to ((count of modList) is (count of idList))
    set out to ${asRow(['"1"', '(name of theFolder as string)'])}
    repeat with i from 1 to (count of idList)
        set modAt to ""
        if hasMods then set modAt to my isoDate(item i of modList)
        set out to out & ${asRow([
        '(item i of idList as string)',
        '(item i of subjList)',
        'modAt',
    ])}
    end repeat
    return out
end tell`;

    const records = splitRecords(await runOsaScript(indexScript, 120000));
    const header = splitFields(records[0] || '');
    if (field(header, 0) !== '1') {
        return {
            folderFound: false,
            folderPath: '',
            templates: [],
            availableFolders: splitList(field(header, 1)),
        };
    }
    const folderPath = macFolderPath(emailAccount, field(header, 1) || folderName, []);

    const indexed = records.slice(1).map(record => {
        const parts = splitFields(record);
        return {
            id: field(parts, 0),
            subject: field(parts, 1).trim(),
            lastModified: field(parts, 2),
        };
    });
    // Newest first, matching the Windows reader's Sort on LastModificationTime —
    // and done before the cap, so "the newest 20" means that rather than an
    // arbitrary 20. 'yyyy-MM-dd HH:mm' sorts lexicographically.
    indexed.sort((a, b) => b.lastModified.localeCompare(a.lastModified));
    const chosen = indexed
        .filter(item => !wanted || item.subject.toLowerCase() === wanted)
        .slice(0, cap);
    if (chosen.length === 0) {
        return { folderFound: true, folderPath, templates: [], availableFolders: [] };
    }

    // Pass 2 — the bodies, for the chosen templates only.
    const bodyProperty = includeBody ? 'content' : 'plain text content';
    const bodyScript = `tell application "Microsoft Outlook"
    set out to ""
    repeat with theId in {${chosen.map(item => item.id).join(', ')}}
        set theMsg to missing value
        try
            set theMsg to message id theId
        end try
        if theMsg is not missing value then
            set bodyText to ""
            try
                set bodyText to (${bodyProperty} of theMsg) as string
            end try
            set out to out & ${asRow(['(theId as string)', 'bodyText'])}
        end if
    end repeat
    return out
end tell`;

    const bodies = new Map<string, string>();
    for (const record of splitRecords(await runOsaScript(bodyScript, 120000))) {
        const parts = splitFields(record);
        bodies.set(field(parts, 0), field(parts, 1));
    }

    const templates: TemplateEmail[] = chosen.map(item => {
        const body = bodies.get(item.id) || '';
        const html = includeBody
            ? stripEmbeddedImages(body).slice(0, MAX_TEMPLATE_HTML)
            : '';
        // With a body in hand the HTML is authoritative (it survives Word splitting
        // a marker across tags); without one, scan the plain-text body, where a
        // marker cannot have been split.
        const markers = findTemplateMarkers(includeBody ? html : body);
        return {
            entryId: item.id,
            subject: item.subject,
            htmlBody: html,
            bodyPreview: includeBody ? '' : body.trim().slice(0, PREVIEW_CHARS),
            sections: markers.sections,
            placeholders: markers.placeholders,
            lastModified: item.lastModified,
        };
    });
    return { folderFound: true, folderPath, templates, availableFolders: [] };
}

/**
 * Save a new template email into a mailbox folder (default "Templates"),
 * creating the folder at the mailbox root if it doesn't exist. The item is a
 * plain unsent mail (subject + HTML body) that can be edited in Outlook.
 * Never overwrites an existing item — it always adds a new one, so check with
 * readTemplateEmails first and only call after the user has agreed to create it.
 */
export async function saveTemplateEmail(
    emailAccount: string,
    subject: string,
    htmlBody: string,
    folderName = 'Templates',
): Promise<SaveTemplateResult> {
    const acct = await resolveMacAccount(emailAccount);
    const script = `${FIND_FOLDER_HANDLER}
tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'root folder', 'rootFolder')}
end tell
set theFolder to my findFolderByName(rootFolder, "${asEscape(folderName)}", ${TEMPLATE_SEARCH_DEPTH})
tell application "Microsoft Outlook"
    set wasCreated to false
    if theFolder is missing value then
        set theFolder to (make new mail folder at rootFolder with properties {name:"${asEscape(folderName)}"})
        set wasCreated to true
    end if
    set countBefore to count of messages of theFolder
    set newMsg to make new outgoing message with properties {subject:"${asEscape(subject)}", content:"${asEscape(htmlBody)}"}
    try
        set account of newMsg to targetAcct
    end try
    -- 'move' reports success even where it silently does nothing, so a template
    -- that was never filed must not be reported as saved.
    move newMsg to theFolder
    if (count of messages of theFolder) is not greater than countBefore then error "Outlook did not file the template in '" & (name of theFolder) & "'."
    return ${asRow(['(name of theFolder as string)', '(wasCreated as string)'])}
end tell`;
    const parts = summaryFields(await runOsaScript(script, 120000));
    return {
        folderPath: macFolderPath(emailAccount, field(parts, 0) || folderName, []),
        folderCreated: boolField(parts, 1),
    };
}

/**
 * Open an Outlook compose window pre-filled with `currentHtml`, wait for the
 * user to close it, and return the HTML they saved.
 *
 * Deliberately unbounded: the run lasts as long as the person is editing. Give
 * it an AbortSignal (`bridge.withOptions({ signal })`) if the caller needs a way
 * to give up on them.
 *
 * The macOS half of the round trip differs from Windows in one visible way: an
 * edited draft is filed by answering Outlook's own "save this message?" prompt
 * on close, so the subject tells the user to save rather than to press a key.
 */
export async function editEmailTemplate(label: string, currentHtml: string): Promise<string> {
    const subject = `${label} - Save and close when done`;
    const script = `tell application "Microsoft Outlook"
    set editSubject to "${asEscape(subject)}"
    set draftMsg to make new outgoing message with properties {subject:editSubject, content:"${asEscape(currentHtml)}"}
    open draftMsg
    activate
end tell

-- Poll until the compose window is gone. A draft window carries the message it
-- is editing as its object, so the one we opened is identified by subject rather
-- than by counting windows the user may also have open.
set stillOpen to true
repeat while stillOpen
    delay 0.7
    set stillOpen to false
    tell application "Microsoft Outlook"
        try
            repeat with w in draft windows
                try
                    if ((subject of (object of w)) as string) is editSubject then
                        set stillOpen to true
                        exit repeat
                    end if
                end try
            end repeat
        end try
    end tell
end repeat

tell application "Microsoft Outlook"
    -- Prefer the saved draft: it reflects what the user chose to keep, and it
    -- survives the window closing. Any leftover editor drafts are removed so
    -- they don't pile up in the Drafts folder.
    set savedHtml to ""
    try
        repeat with acctList in {exchange accounts, imap accounts, pop accounts}
            repeat with a in acctList
                try
                    repeat with m in (messages of (drafts of a))
                        try
                            if ((subject of m) as string) is editSubject then
                                if savedHtml is "" then set savedHtml to (content of m) as string
                                delete m
                            end if
                        end try
                    end repeat
                end try
            end repeat
        end repeat
    end try
    return my sanitize(savedHtml)
end tell`;
    // 0 disables the timeout: this run is as long as the edit takes.
    const saved = await runOsaScript(script, 0);
    if (!saved) {
        throw new NotFoundError('template', 'No template saved. Did you save the draft before closing?');
    }
    return saved;
}
