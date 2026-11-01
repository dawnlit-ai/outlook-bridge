// Template emails: ordinary mail items kept in a mailbox folder and reused as
// email bodies, plus the compose-window editor for one.
//
// Outlook for Mac serves a message's HTML through `content` — the same thing
// COM's HTMLBody returns — which is what makes templates portable between the
// two platforms.
import {
    asIdList,
    asInt,
    asRow,
    asString,
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
    folderChainSnippet,
    macFolderPath,
    resolveMacAccount,
    rootFolderSnippet,
} from './scripts';
import {
    EMBEDDED_IMAGE_SOURCE,
    MAX_TEMPLATE_HTML,
    TEMPLATE_PREVIEW_CHARS,
    TEMPLATE_SEARCH_DEPTH,
} from '../shared/templateLimits';
import { findTemplateMarkers } from '../templateBody';
import { NotFoundError } from '../errors';
import type { EditTemplateRequest, ReadTemplatesRequest, SaveTemplateRequest } from '../backend';
import type { SaveTemplateResult, TemplateEmail, TemplateFolderResult } from '../types';

/** A template body without its embedded images — see EMBEDDED_IMAGE_SOURCE. */
function stripEmbeddedImages(html: string): string {
    return html.replace(new RegExp(EMBEDDED_IMAGE_SOURCE, 'gi'), '');
}

/**
 * The template emails in a mailbox folder, newest first; a missing folder is
 * reported as folderFound:false with the mailbox's top-level folder names.
 *
 * Bodies are read in a second pass, for the chosen templates only: a folder of
 * Word-sized templates is megabytes, and pulling all of them through stdout on
 * a call that wanted one is what overruns the output limit.
 */
export async function readTemplateEmails(request: ReadTemplatesRequest): Promise<TemplateFolderResult> {
    const acct = await resolveMacAccount(request.account);
    // Pass 1 — find the folder and index it: three bulk reads, no bodies.
    const records = splitRecords(await runOsaScript(`${FIND_FOLDER_HANDLER}
tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'root folder', 'rootFolder')}
end tell
set theFolder to my findFolderByName(rootFolder, ${asString(request.folder)}, ${asInt(TEMPLATE_SEARCH_DEPTH)})
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
    -- The folder may sit up to the search depth down, so its whole chain is
    -- reported, not just its name: the path is the one Windows' FolderPath gives.
${folderChainSnippet('    ')}
    set out to ${asRow(['"1"', 'my sanitizeList(chainNames)'])}
    repeat with i from 1 to (count of idList)
        set modAt to ""
        if hasMods then set modAt to my isoDate(item i of modList)
        set out to out & ${asRow(['(item i of idList as string)', '(item i of subjList)', 'modAt'])}
    end repeat
    return out
end tell`, 'standard'));

    const header = splitFields(records[0] ?? '');
    if (field(header, 0) !== '1') {
        return { folderFound: false, folderPath: '', templates: [], availableFolders: splitList(field(header, 1)) };
    }
    const chain = splitList(field(header, 1)).reverse();
    const folderPath = chain.length
        ? macFolderPath(request.account, chain[0], chain.slice(1))
        : macFolderPath(request.account, request.folder, []);
    const wanted = (request.subject ?? '').toLowerCase();
    const chosen = records.slice(1)
        .map(record => {
            const parts = splitFields(record);
            return { id: field(parts, 0), subject: field(parts, 1).trim(), lastModified: field(parts, 2) };
        })
        // Newest first, before the cap, so "the newest 20" means that.
        .sort((a, b) => b.lastModified.localeCompare(a.lastModified))
        .filter(item => !wanted || item.subject.toLowerCase() === wanted)
        .slice(0, request.limit);
    if (chosen.length === 0) return { folderFound: true, folderPath, templates: [], availableFolders: [] };

    // Pass 2 — the bodies, for the chosen templates only.
    const bodyProperty = request.includeBody ? 'content' : 'plain text content';
    const bodies = new Map<string, string>();
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
    set out to ""
    repeat with theId in ${asIdList(chosen.map(item => item.id))}
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
end tell`, 'standard');
    for (const record of splitRecords(raw)) {
        const parts = splitFields(record);
        bodies.set(field(parts, 0), field(parts, 1));
    }

    const templates: TemplateEmail[] = chosen.map(item => {
        const body = bodies.get(item.id) ?? '';
        const html = request.includeBody ? stripEmbeddedImages(body).slice(0, MAX_TEMPLATE_HTML) : '';
        // With the HTML in hand it is authoritative (markers survive Word
        // splitting them across tags); without it, the plain text is scanned,
        // where a marker can't have been split.
        const markers = findTemplateMarkers(request.includeBody ? html : body);
        return {
            entryId: item.id,
            subject: item.subject,
            htmlBody: html,
            bodyPreview: request.includeBody ? '' : body.trim().slice(0, TEMPLATE_PREVIEW_CHARS),
            sections: markers.sections,
            placeholders: markers.placeholders,
            lastModified: item.lastModified,
        };
    });
    return { folderFound: true, folderPath, templates, availableFolders: [] };
}

/**
 * Save a new template email into a folder, creating the folder at the mailbox
 * root when absent. Always adds a new item — never overwrites one.
 */
export async function saveTemplateEmail(request: SaveTemplateRequest): Promise<SaveTemplateResult> {
    const acct = await resolveMacAccount(request.account);
    const parts = summaryFields(await runOsaScript(`${FIND_FOLDER_HANDLER}
tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'root folder', 'rootFolder')}
end tell
set theFolder to my findFolderByName(rootFolder, ${asString(request.folder)}, ${asInt(TEMPLATE_SEARCH_DEPTH)})
tell application "Microsoft Outlook"
    set wasCreated to false
    if theFolder is missing value then
        set theFolder to (make new mail folder at rootFolder with properties {name:${asString(request.folder)}})
        set wasCreated to true
    end if
    set countBefore to count of messages of theFolder
    set newMsg to make new outgoing message with properties {subject:${asString(request.subject)}, content:${asString(request.htmlBody)}}
    try
        set account of newMsg to targetAcct
    end try
    -- 'move' reports success even where it silently does nothing, so a template
    -- that was never filed must not be reported as saved.
    move newMsg to theFolder
    if (count of messages of theFolder) is not greater than countBefore then error "Outlook did not file the template in '" & (name of theFolder) & "'."
    return ${asRow(['(name of theFolder as string)', '(wasCreated as string)'])}
end tell`, 'standard'));
    return {
        folderPath: macFolderPath(request.account, field(parts, 0) || request.folder, []),
        folderCreated: boolField(parts, 1),
    };
}

/**
 * Open a compose window pre-filled with the HTML, wait for it to close, and
 * return the saved draft's HTML. On macOS the edit is kept by answering
 * Outlook's own "save this message?" prompt on close.
 */
export async function editEmailTemplate(request: EditTemplateRequest): Promise<string> {
    const subject = `${request.label} - Save and close when done`;
    const saved = await runOsaScript(`tell application "Microsoft Outlook"
    set editSubject to ${asString(subject)}
    set draftMsg to make new outgoing message with properties {subject:editSubject, content:${asString(request.html)}}
    open draftMsg
    activate
end tell

-- Poll until the compose window is gone, identifying ours by subject rather
-- than counting windows the person may also have open.
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
    -- Prefer the saved draft: it is what the person chose to keep, and it
    -- outlives the window. Every editor draft is removed so none pile up.
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
end tell`, 'interactive');
    if (!saved) {
        throw new NotFoundError('template', 'No template was saved. Save the draft before closing the window.');
    }
    return saved;
}
