// Reading mail on macOS: one folder, one item, a whole Inbox tree, or whatever
// is selected in the running Outlook.
//
// Every reader here is two passes. Pass 1 indexes a folder with bulk property
// reads (one Apple event each); pass 2 fetches the expensive fields — body,
// attachments, sender — only for the messages that survived. Reading
// per-message in a loop costs an event per property and is unusably slow on a
// real mailbox, which is the whole reason for the shape.
import { AS_LIST_SEP, asRow, field, runOsaScript, splitFields, splitList, splitRecords, summaryFields, } from './run';
import {
    accountLookupSnippet,
    dateProperty,
    firstRecipientSnippet,
    macFolderPath,
    macMessageId,
    mailScopeSnippet,
    MessageDetail,
    messageDetailFields,
    messageDetailSnippet,
    messageLookupSnippet,
    senderSnippet,
} from './scripts';
import { isOutgoingRoot, mailFolderRef, splitQuotedOriginal } from '../mail';
import { NotFoundError } from '../errors';
import type { EmailBodyResult, InboxEmail, InboxSearchFilter, InboxSearchMatch, SelectedEmail, } from '../types';

/** How much of a body readInboxEmails previews, matching the Windows reader. */
const PREVIEW_CHARS = 600;

/**
 * Read recent messages for the given account, newest first.
 *
 * `entryId` here is Outlook for Mac's small integer message id (e.g. "779"), not
 * the MAPI EntryID string Windows returns. The two are not interchangeable, so
 * don't persist one and look it up on the other platform.
 */
export async function readInboxEmails(
    emailAccount: string,
    daysBack: number = 60,
    limit: number = 50,
    folder?: string,
): Promise<InboxEmail[]> {
    const days = Math.max(0, Math.floor(daysBack));
    const cap = Math.max(0, Math.floor(limit));
    if (cap === 0) return [];
    const ref = folder ? mailFolderRef(folder) : {rootId: 6, rootLabel: 'Inbox', segments: []};
    // A folder argument that trims away to nothing ("\\", "  ") would otherwise
    // read the Inbox root and look like it had scoped — the exact silent
    // mis-scoping this parameter exists to prevent. Mirrors the Windows check.
    if (folder && folder.trim() && ref.segments.length === 0 && ref.rootId === 6
        && folder.trim().toLowerCase() !== 'inbox') {
        throw new NotFoundError('folder', `Folder '${folder}' does not name a folder under the Inbox.`);
    }
    const resolveScope = mailScopeSnippet(ref, folder || '');
    const folderPath = macFolderPath(emailAccount, ref.rootLabel, ref.segments);
    const isOutgoing = isOutgoingRoot(ref.rootId);

    // Pass 1 — index the folder. Folder order isn't documented, so sort here
    // rather than trusting Outlook to hand back newest-first.
    const indexScript = `tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
${resolveScope}
    set cutoff to (current date) - (${days} * days)
    set idList to id of every message of scopeFolder
    try
        set timeList to ${dateProperty(ref.rootId)} of every message of scopeFolder
    on error
        set timeList to {}
    end try
    set hasTimes to ((count of timeList) is (count of idList))
    set out to ""
    repeat with i from 1 to (count of idList)
        set d to missing value
        if hasTimes then set d to item i of timeList
        -- An item carrying no timestamp of its own (an unsent draft filed into a
        -- mail folder) is skipped rather than crashing the walk, matching what the
        -- Windows Restrict on the same property does with it.
        if d is not missing value and d is greater than or equal to cutoff then
            set out to out & (item i of idList as string) & tab & my isoDate(d) & linefeed
        end if
    end repeat
    return out
end tell`;

    const index = (await runOsaScript(indexScript, 60000))
        .split('\n')
        .map(line => line.trim())
        .filter(Boolean)
        .map(line => {
            const [id, receivedTime] = line.split('\t');
            return {id, receivedTime: receivedTime || ''};
        });
    // 'yyyy-MM-dd HH:mm' is lexicographically ordered, so plain string compare sorts it.
    index.sort((a, b) => b.receivedTime.localeCompare(a.receivedTime));
    const chosen = index.slice(0, cap);
    if (chosen.length === 0) return [];

    // Pass 2 — the expensive reads, only for messages we keep.
    const detailScript = `tell application "Microsoft Outlook"
    set wanted to {${chosen.map(c => c.id).join(', ')}}
    set out to ""
    repeat with k from 1 to (count of wanted)
        set theMsg to missing value
        try
            set theMsg to message id (item k of wanted)
        end try
        if theMsg is not missing value then
            set subj to ""
            try
                set subj to (subject of theMsg) as string
            end try
${isOutgoing ? firstRecipientSnippet('theMsg', '            ') : senderSnippet('theMsg', '            ')}
            set bodyText to ""
            try
                set bodyText to (plain text content of theMsg) as string
            end try
            if (length of bodyText) > ${PREVIEW_CHARS} then set bodyText to text 1 thru ${PREVIEW_CHARS} of bodyText
            set attNames to {}
            try
                set attNames to name of every attachment of theMsg
            end try
            set out to out & ${asRow([
        '(id of theMsg as string)',
        'subj',
        'sndName',
        'sndAddr',
        'bodyText',
        'my sanitizeList(attNames)',
    ])}
        end if
    end repeat
    return out
end tell`;

    const byId = new Map(chosen.map(c => [c.id, c.receivedTime]));
    return splitRecords(await runOsaScript(detailScript, 120000)).map(record => {
        const parts = splitFields(record);
        const id = field(parts, 0);
        const attachmentNames = splitList(field(parts, 5));
        return {
            entryId: id,
            storeId: '', // macOS AppleScript has no StoreID equivalent
            subject: field(parts, 1).trim(),
            senderName: field(parts, 2),
            senderEmail: field(parts, 3),
            receivedTime: byId.get(id) || '',
            bodyPreview: field(parts, 4),
            attachmentNames,
            attachmentCount: attachmentNames.length,
            folderPath,
        };
    });
}

/**
 * Read one email's full plain-text body by message id.
 *
 * `storeId` is accepted for signature parity and ignored: macOS has no StoreID,
 * and `message id N` resolves against the application rather than one folder, so
 * the message is found wherever it currently sits — including a subfolder.
 *
 * Uses `plain text content`, matching the Windows reader's use of `.Body`: the
 * same tradeoff applies, so an HTML table's rows flatten and tabular figures are
 * better read from an attachment.
 */
export async function readEmailBody(
    entryId: string,
    _storeId?: string,
    maxChars: number = 8000,
    includeQuoted: boolean = false,
): Promise<EmailBodyResult> {
    const id = macMessageId(entryId);
    const script = `tell application "Microsoft Outlook"
${messageLookupSnippet(id)}
${messageDetailSnippet(true)}
    return ${asRow(messageDetailFields(true))}
end tell`;

    // The body is emitted LAST so a stray separator in earlier fields can't shift it.
    const parts = summaryFields(await runOsaScript(script, 60000));
    const attachmentNames = splitList(field(parts, MessageDetail.attachmentNames));
    const {body, quoted, separator} = splitQuotedOriginal(field(parts, MessageDetail.body));
    // The quoted thread is context, never the priced content, so it is capped
    // harder than the reply itself — matching the Windows reader.
    const quotedCap = Math.min(maxChars, 4000);
    return {
        entryId: field(parts, MessageDetail.id) || id,
        subject: field(parts, MessageDetail.subject).trim(),
        senderName: field(parts, MessageDetail.senderName),
        senderEmail: field(parts, MessageDetail.senderEmail),
        receivedTime: field(parts, MessageDetail.receivedTime),
        body: body.length > maxChars ? body.slice(0, maxChars) : body,
        truncated: body.length > maxChars,
        bodyLength: body.length,
        quoteSeparator: separator,
        quotedLength: quoted.length,
        quotedOriginal: includeQuoted ? quoted.slice(0, quotedCap) : '',
        attachmentNames,
        attachmentCount: attachmentNames.length,
    };
}

/**
 * Translate a Restrict-style `like` pattern to a regex.
 *
 * Windows hands `subjectLike` to Outlook's server-side Restrict; AppleScript has
 * no equivalent, so the same pattern is applied here instead. `%` and `*` are
 * the any-run wildcards, `_` and `?` match one character, and the match is
 * anchored because that is what SQL `like` means.
 */
function likePattern(pattern: string): RegExp {
    let source = '';
    for (const ch of pattern) {
        if (ch === '%' || ch === '*') source += '.*';
        else if (ch === '_' || ch === '?') source += '.';
        else source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp(`^${source}$`, 'i');
}

/**
 * Walk every folder under the Inbox (recursively) for one account and return the
 * emails matching `filter` — full body included, attachments listed by name but
 * not saved.
 *
 * The filtering happens here rather than in Outlook: `Items.Restrict` is a COM
 * facility with no AppleScript counterpart, so the folder walk indexes each
 * folder in two bulk reads and the subject/date tests run on the index. The
 * observable contract is the same; the cost profile is not, and a `daysBack` of
 * 0 over a large mailbox tree is correspondingly heavier here.
 *
 * `subjectPattern` is applied case-insensitively regardless of the RegExp's own
 * flags, matching PowerShell's `-match`.
 */
export async function searchInboxByFilter(
    emailAccount: string,
    filter: InboxSearchFilter = {},
    daysBack = 0,
): Promise<InboxSearchMatch[]> {
    const days = Math.max(0, Math.floor(daysBack));
    // Pass 1 — index every folder under the Inbox: path, id, subject, received.
    const indexScript = `on scanFolder(theFolder, prefix, cutoff, useCutoff)
    set out to ""
    tell application "Microsoft Outlook"
        set idList to id of every message of theFolder
        set subjList to subject of every message of theFolder
        try
            set timeList to time received of every message of theFolder
        on error
            set timeList to {}
        end try
        set subs to mail folders of theFolder
    end tell
    set hasTimes to ((count of timeList) is (count of idList))
    repeat with i from 1 to (count of idList)
        set d to missing value
        if hasTimes then set d to item i of timeList
        set keep to true
        if d is missing value then
            -- A folder whose items carry no readable receive time can only be
            -- included when the caller isn't bounding by date; silently dropping
            -- it would look exactly like an empty folder.
            set keep to not useCutoff
        else if useCutoff and d is less than cutoff then
            set keep to false
        end if
        if keep then
            set out to out & ${asRow([
        'prefix',
        '(item i of idList as string)',
        '(item i of subjList)',
        'my isoDate(d)',
    ])}
        end if
    end repeat
    repeat with f in subs
        set childName to ""
        try
            set childName to (name of f) as string
        end try
        set out to out & my scanFolder(f, prefix & ${AS_LIST_SEP} & childName, cutoff, useCutoff)
    end repeat
    return out
end scanFolder

tell application "Microsoft Outlook"
${accountLookupSnippet(emailAccount)}
    set rootInbox to inbox of targetAcct
    set rootName to (name of rootInbox) as string
end tell
set cutoff to (current date) - (${days} * days)
return my scanFolder(rootInbox, rootName, cutoff, ${days > 0 ? 'true' : 'false'})`;

    const likeRe = filter.subjectLike ? likePattern(filter.subjectLike) : null;
    // PowerShell's -match is case-insensitive by default, so the pattern is
    // rebuilt with `i` regardless of the flags the caller's RegExp carries.
    const patternRe = filter.subjectPattern
        ? new RegExp(filter.subjectPattern.source, 'i')
        : null;

    const candidates = splitRecords(await runOsaScript(indexScript, 300000))
        .map(record => {
            const parts = splitFields(record);
            const segments = splitList(field(parts, 0));
            return {
                folderPath: macFolderPath(emailAccount, segments[0] || 'Inbox', segments.slice(1)),
                id: field(parts, 1),
                subject: field(parts, 2).trim(),
                receivedTime: field(parts, 3),
            };
        })
        .filter(candidate => {
            if (!candidate.subject) return false;
            if (likeRe && !likeRe.test(candidate.subject)) return false;
            if (patternRe && !patternRe.test(candidate.subject)) return false;
            if (filter.excludeReplies && /^(re|fwd?)\s*:/i.test(candidate.subject)) return false;
            return true;
        });
    if (candidates.length === 0) return [];

    // Pass 2 — bodies, senders and attachment names for the survivors only.
    const detailScript = `tell application "Microsoft Outlook"
    set wanted to {${candidates.map(c => c.id).join(', ')}}
    set out to ""
    repeat with k from 1 to (count of wanted)
        set theMsg to missing value
        try
            set theMsg to message id (item k of wanted)
        end try
        if theMsg is not missing value then
${senderSnippet('theMsg', '            ')}
            set attNames to {}
            try
                set attNames to name of every attachment of theMsg
            end try
            set bodyText to ""
            try
                set bodyText to (plain text content of theMsg) as string
            end try
            set out to out & ${asRow([
        '(id of theMsg as string)',
        'sndName',
        'sndAddr',
        'my sanitizeList(attNames)',
        'bodyText',
    ])}
        end if
    end repeat
    return out
end tell`;

    const details = new Map<string, {
        senderName: string;
        senderEmail: string;
        attachmentNames: string[];
        body: string
    }>();
    for (const record of splitRecords(await runOsaScript(detailScript, 300000))) {
        const parts = splitFields(record);
        details.set(field(parts, 0), {
            senderName: field(parts, 1),
            senderEmail: field(parts, 2),
            attachmentNames: splitList(field(parts, 3)),
            body: field(parts, 4),
        });
    }

    const matches: InboxSearchMatch[] = [];
    for (const candidate of candidates) {
        const detail = details.get(candidate.id);
        if (!detail) continue;
        // requireAttachment can only be answered once the attachment names are in
        // hand, so it filters here rather than during the index pass.
        if (filter.requireAttachment && detail.attachmentNames.length === 0) continue;
        matches.push({
            entryId: candidate.id,
            storeId: '', // macOS AppleScript has no StoreID equivalent
            subject: candidate.subject,
            senderName: detail.senderName,
            senderEmail: detail.senderEmail,
            receivedTime: candidate.receivedTime,
            body: detail.body,
            attachmentNames: detail.attachmentNames,
            folderPath: candidate.folderPath,
        });
    }
    return matches;
}

/**
 * Read the email currently selected (or open) in Outlook — full body,
 * attachments listed by name but not saved.
 *
 * `current messages` covers both cases the Windows reader handles separately
 * (an explorer selection and an open item), so there is no second lookup here.
 */
export async function readSelectedEmail(): Promise<SelectedEmail> {
    const script = `tell application "Microsoft Outlook"
    set sel to {}
    try
        set sel to current messages
    end try
    if (count of sel) is 0 then error "No email is selected in Outlook. Open Outlook, select (or open) an email, then try again."
    set theMsg to item 1 of sel
${messageDetailSnippet(true)}
    return ${asRow(messageDetailFields(true))}
end tell`;
    const records = splitRecords(await runOsaScript(script, 30000));
    if (records.length === 0) {
        throw new NotFoundError('email', 'Failed to read the selected Outlook email.');
    }
    const parts = splitFields(records[0]);
    return {
        entryId: field(parts, MessageDetail.id),
        storeId: '', // macOS AppleScript has no StoreID equivalent
        subject: field(parts, MessageDetail.subject).trim(),
        senderName: field(parts, MessageDetail.senderName),
        senderEmail: field(parts, MessageDetail.senderEmail),
        receivedTime: field(parts, MessageDetail.receivedTime),
        body: field(parts, MessageDetail.body),
        attachmentNames: splitList(field(parts, MessageDetail.attachmentNames)),
    };
}

/** Open an email in Outlook by its message id, and bring Outlook forward. */
export async function openOutlookEmail(entryId: string): Promise<void> {
    const id = macMessageId(entryId);
    const script = `
tell application "Microsoft Outlook"
${messageLookupSnippet(id)}
    open theMsg
    activate
end tell`;
    await runOsaScript(script, 30000);
}
