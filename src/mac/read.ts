// Reading mail on macOS: one folder, one item, a whole Inbox tree, or whatever
// is selected in the running Outlook.
//
// Every reader is two passes. Pass 1 indexes a folder with bulk property reads
// (one Apple event each); pass 2 fetches the expensive fields — body,
// attachments, sender — only for the messages that survived. Reading
// per-message in a loop costs an event per property and is unusably slow on a
// real mailbox, which is the whole reason for the shape.
import {
    AS_LIST_SEP,
    asIdList,
    asInt,
    asRow,
    asString,
    field,
    runOsaScript,
    splitFields,
    splitList,
    splitRecords,
    summaryFields
} from './run';
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
    resolveMacAccount,
    rootFolderSnippet,
    senderSnippet,
} from './scripts';
import { failureTag, NotFoundError } from '../errors';
import { folderLeafName, isOutgoingRoot, REPLY_PREFIX, subjectGlobSource } from '../mail';
import type { RawEmail, ReadInboxRequest, SearchRequest } from '../backend';
import type { EmailLocator, InboxEmail, InboxSearchMatch, SelectedEmail } from '../types';

/**
 * Recent mail from the Inbox root or one folder, newest first. Entry ids here
 * are Outlook for Mac's integer message ids, not Windows EntryIDs.
 */
export async function readInboxEmails(request: ReadInboxRequest): Promise<InboxEmail[]> {
    const acct = await resolveMacAccount(request.account);
    const folderPath = macFolderPath(request.account, request.folder.rootLabel, request.folder.segments);
    const outgoing = isOutgoingRoot(request.folder.rootId);

    // Pass 1 — index the folder. Its order isn't documented, so it is sorted here.
    const index = (await runOsaScript(`tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${mailScopeSnippet(acct, request.folder, request.folderLabel)}
    set cutoff to (current date) - (${asInt(request.daysBack)} * days)
    set idList to id of every message of scopeFolder
    try
        set timeList to ${dateProperty(request.folder.rootId)} of every message of scopeFolder
    on error
        set timeList to {}
    end try
    set hasTimes to ((count of timeList) is (count of idList))
    set out to ""
    repeat with i from 1 to (count of idList)
        set d to missing value
        if hasTimes then set d to item i of timeList
        -- An item with no timestamp of its own (an unsent draft filed into a
        -- mail folder) is skipped, as the Windows date filter skips it.
        if d is not missing value and d is greater than or equal to cutoff then
            set out to out & (item i of idList as string) & tab & my isoDate(d) & linefeed
        end if
    end repeat
    return out
end tell`, 'standard'))
        .split('\n')
        .map(line => line.trim())
        .filter(Boolean)
        .map(line => {
            const [id, receivedTime] = line.split('\t');
            return {id, receivedTime: receivedTime || ''};
        });
    // 'yyyy-MM-dd HH:mm' sorts correctly as a string.
    index.sort((a, b) => b.receivedTime.localeCompare(a.receivedTime));
    const chosen = index.slice(0, request.limit);
    if (chosen.length === 0) return [];

    // Pass 2 — the expensive reads, only for the messages kept.
    const preview = request.previewChars > 0
        ? `            set bodyText to ""
            try
                set bodyText to (plain text content of theMsg) as string
            end try
            if (length of bodyText) > ${asInt(request.previewChars)} then set bodyText to text 1 thru ${asInt(request.previewChars)} of bodyText`
        : '            set bodyText to ""';
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
    set out to ""
    repeat with theId in ${asIdList(chosen.map(c => c.id))}
        set theMsg to missing value
        try
            set theMsg to message id theId
        end try
        if theMsg is not missing value then
            set subj to ""
            try
                set subj to (subject of theMsg) as string
            end try
${outgoing ? firstRecipientSnippet('theMsg', '            ') : senderSnippet('theMsg', '            ')}
${preview}
            set attNames to {}
            try
                set attNames to name of every attachment of theMsg
            end try
            set out to out & ${asRow(['(id of theMsg as string)', 'subj', 'sndName', 'sndAddr', 'bodyText', 'my sanitizeList(attNames)'])}
        end if
    end repeat
    return out
end tell`, 'standard');

    const receivedById = new Map(chosen.map(c => [c.id, c.receivedTime]));
    return splitRecords(raw).map(record => {
        const parts = splitFields(record);
        const id = field(parts, 0);
        const attachmentNames = splitList(field(parts, 5));
        return {
            entryId: id,
            storeId: '',
            subject: field(parts, 1).trim(),
            senderName: field(parts, 2),
            senderEmail: field(parts, 3),
            receivedTime: receivedById.get(id) || '',
            bodyPreview: field(parts, 4),
            attachmentNames,
            attachmentCount: attachmentNames.length,
            folderPath,
        };
    });
}

/**
 * One email's full plain-text body. `storeId` has no macOS meaning: `message id
 * N` resolves against the application, wherever the message sits.
 */
export async function readEmailBody(email: EmailLocator): Promise<RawEmail> {
    const id = macMessageId(email.entryId);
    const parts = summaryFields(await runOsaScript(`tell application "Microsoft Outlook"
${messageLookupSnippet(id)}
${messageDetailSnippet(true)}
    return ${asRow(messageDetailFields(true))}
end tell`, 'quick'));
    return {
        entryId: field(parts, MessageDetail.id) || id,
        subject: field(parts, MessageDetail.subject).trim(),
        senderName: field(parts, MessageDetail.senderName),
        senderEmail: field(parts, MessageDetail.senderEmail),
        receivedTime: field(parts, MessageDetail.receivedTime),
        body: field(parts, MessageDetail.body),
        attachmentNames: splitList(field(parts, MessageDetail.attachmentNames)),
    };
}

/**
 * Walk every folder under the Inbox for the mail matching the request.
 *
 * The filtering happens here rather than in Outlook: there is no AppleScript
 * counterpart to COM's server-side `Items.Restrict`, so each folder is indexed in
 * bulk reads and the subject and date tests run on the index. The results match
 * Windows; the cost doesn't, and a wide window over a big tree is heavier here.
 */
export async function searchInboxByFilter(request: SearchRequest): Promise<InboxSearchMatch[]> {
    const acct = await resolveMacAccount(request.account);
    // AppleScript reaches a child folder by name, so a path entry narrows to its leaf.
    const asNameList = (list: readonly string[]) => `{${list.map(folderLeafName).filter(Boolean).map(asString).join(', ')}}`;
    const rootInScope = request.includeFolders.length === 0 ? 'true' : 'false';
    // Sent, Drafts, Deleted and Junk hang UNDER the Inbox on an IMAP profile, so
    // their ids are collected first and the walk never descends into them — by
    // the profile's own id where it records one, through the account otherwise.
    const skipIds = ['sent items', 'deleted items', 'drafts', 'junk mail']
        .map(term => {
            const id = acct.folderIds?.[term];
            return id !== undefined
                ? `    set end of skipIds to ${id}`
                : `    try
        set end of skipIds to (id of (${term} of targetAcct))
    end try`;
        })
        .join('\n');

    // Pass 1 — index every folder under the Inbox: path, id, subject, received.
    const indexed = splitRecords(await runOsaScript(`on scanFolder(theFolder, prefix, cutoff, skipIds, skipNames, onlyNames, inScope)
    set out to ""
    set idList to {}
    set subjList to {}
    set timeList to {}
    -- A folder the walk only passes through is never read: the bulk reads are
    -- the expensive part.
    if inScope then
        tell application "Microsoft Outlook"
            set idList to id of every message of theFolder
            set subjList to subject of every message of theFolder
            try
                set timeList to time received of every message of theFolder
            on error
                set timeList to {}
            end try
        end tell
    end if
    tell application "Microsoft Outlook"
        set subs to mail folders of theFolder
    end tell
    set hasTimes to ((count of timeList) is (count of idList))
    repeat with i from 1 to (count of idList)
        set d to missing value
        if hasTimes then set d to item i of timeList
        -- An item with no readable receive time is kept whatever the window
        -- says; dropping it would make a folder of them look empty.
        if d is missing value or d is not less than cutoff then
            set out to out & ${asRow(['prefix', '(item i of idList as string)', '(item i of subjList)', 'my isoDate(d)'])}
        end if
    end repeat
    repeat with f in subs
        set skipThis to false
        try
            if skipIds contains (id of f) then set skipThis to true
        end try
        try
            if skipNames contains ((name of f) as string) then set skipThis to true
        end try
        if not skipThis then
            set childName to ""
            try
                set childName to (name of f) as string
            end try
            set childScope to (count of onlyNames) is 0
            try
                if onlyNames contains childName then set childScope to true
            end try
            set out to out & my scanFolder(f, prefix & ${AS_LIST_SEP} & childName, cutoff, skipIds, skipNames, onlyNames, childScope)
        end if
    end repeat
    return out
end scanFolder

tell application "Microsoft Outlook"
${accountLookupSnippet(acct)}
${rootFolderSnippet(acct, 'inbox', 'rootInbox')}
    set rootName to (name of rootInbox) as string
    set skipIds to {}
${skipIds}
    set skipNames to ${asNameList(request.excludeFolders)}
    set onlyNames to ${asNameList(request.includeFolders)}
end tell
set cutoff to (current date) - (${asInt(request.daysBack)} * days)
return my scanFolder(rootInbox, rootName, cutoff, skipIds, skipNames, onlyNames, ${rootInScope})`, 'scan'));

    const subjectFiltered = !!(request.subjectLike || request.subjectPattern || request.excludeReplies);
    const likeRe = request.subjectLike ? new RegExp(subjectGlobSource(request.subjectLike), 'i') : null;
    // PowerShell's -match is case-insensitive, so the Windows test is too.
    const patternRe = request.subjectPattern ? new RegExp(request.subjectPattern.source, 'i') : null;
    const candidates = indexed
        .map(record => {
            const parts = splitFields(record);
            const segments = splitList(field(parts, 0));
            return {
                folderPath: macFolderPath(request.account, segments[0] || 'Inbox', segments.slice(1)),
                id: field(parts, 1),
                subject: field(parts, 2).trim(),
                receivedTime: field(parts, 3),
            };
        })
        .filter(candidate => {
            if (subjectFiltered && !candidate.subject) return false;
            if (likeRe && !likeRe.test(candidate.subject)) return false;
            if (patternRe && !patternRe.test(candidate.subject)) return false;
            if (request.excludeReplies && REPLY_PREFIX.test(candidate.subject)) return false;
            return true;
        });
    if (candidates.length === 0) return [];

    // Pass 2 — senders, attachment names and (when asked) bodies, for the survivors only.
    const body = request.includeBody
        ? `            set bodyText to ""
            try
                set bodyText to (plain text content of theMsg) as string
            end try`
        : '            set bodyText to ""';
    const details = new Map<string, {
        senderName: string;
        senderEmail: string;
        attachmentNames: string[];
        body: string
    }>();
    const raw = await runOsaScript(`tell application "Microsoft Outlook"
    set out to ""
    repeat with theId in ${asIdList(candidates.map(c => c.id))}
        set theMsg to missing value
        try
            set theMsg to message id theId
        end try
        if theMsg is not missing value then
${senderSnippet('theMsg', '            ')}
            set attNames to {}
            try
                set attNames to name of every attachment of theMsg
            end try
${body}
            set out to out & ${asRow(['(id of theMsg as string)', 'sndName', 'sndAddr', 'my sanitizeList(attNames)', 'bodyText'])}
        end if
    end repeat
    return out
end tell`, 'scan');
    for (const record of splitRecords(raw)) {
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
        // Attachment names only arrive with the details, so this test waits for them.
        if (request.requireAttachment && detail.attachmentNames.length === 0) continue;
        matches.push({
            entryId: candidate.id,
            storeId: '',
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
 * The email currently selected — or open — in Outlook. `current messages`
 * covers both cases the Windows reader handles separately.
 */
export async function readSelectedEmail(): Promise<SelectedEmail> {
    const records = splitRecords(await runOsaScript(`tell application "Microsoft Outlook"
    set sel to {}
    try
        set sel to current messages
    end try
    if (count of sel) is 0 then error ${asString(`${failureTag('NOT_FOUND', 'email')}No email is selected in Outlook. Select or open an email, then try again.`)}
    set theMsg to item 1 of sel
${messageDetailSnippet(true)}
    return ${asRow(messageDetailFields(true))}
end tell`, 'quick'));
    if (records.length === 0) throw new NotFoundError('email', 'Outlook reported no selected email.');
    const parts = splitFields(records[0]);
    return {
        entryId: field(parts, MessageDetail.id),
        storeId: '',
        subject: field(parts, MessageDetail.subject).trim(),
        senderName: field(parts, MessageDetail.senderName),
        senderEmail: field(parts, MessageDetail.senderEmail),
        receivedTime: field(parts, MessageDetail.receivedTime),
        body: field(parts, MessageDetail.body),
        attachmentNames: splitList(field(parts, MessageDetail.attachmentNames)),
    };
}

/** Open an email in Outlook and bring Outlook forward. */
export async function openOutlookEmail(email: EmailLocator): Promise<void> {
    const id = macMessageId(email.entryId);
    await runOsaScript(`
tell application "Microsoft Outlook"
${messageLookupSnippet(id)}
    open theMsg
    activate
end tell`, 'quick');
}
