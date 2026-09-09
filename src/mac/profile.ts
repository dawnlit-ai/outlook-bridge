// Outlook's own profile database, read for the mailboxes the dictionary omits.
//
// `getOutlookAccounts` can only enumerate `exchange accounts`, `imap accounts`
// and `pop accounts`, and a Microsoft 365 mailbox added by the modern sync
// engine is none of those: Outlook files it in `AccountsMail` under a 'SHDO'
// server type, publishes no account object for it, and answers `every account`
// with -1728 rather than listing it. Even `account of <folder>` and
// `account of <message>` come back `missing value` inside that mailbox's own
// tree, so there is nothing to match an address against — the mailbox is
// invisible to every function in this directory while its mail sits in plain
// sight in the folder list.
//
// The folders themselves are reachable, because `mail folder id N` resolves
// against the application rather than through an account, and N is exactly the
// `Record_RecordID` Outlook's profile database gives that folder. So the
// addresses AppleScript can't report, and the ids of their well-known folders,
// are read from that database instead, and the generated scripts fall back to a
// folder id wherever the account probe comes up empty.
//
// Strictly read-only, and never fatal: every failure here leaves the caller with
// exactly what AppleScript could see on its own.
import { execFile } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';

/** Where Outlook keeps its profiles, and the profile used when there are several. */
const PROFILES_DIR = path.join(
    os.homedir(),
    'Library', 'Group Containers', 'UBF8T346G9.Office', 'Outlook', 'Outlook 15 Profiles',
);
const DEFAULT_PROFILE = 'Main Profile';

/**
 * `Folder_SpecialFolderType` for each root the generated scripts ask for, keyed
 * by the dictionary term `MAC_ROOT_TERMS` resolves.
 *
 * These are Outlook's own numbering, confirmed against a profile's local "On My
 * Computer" store, where every one of them appears exactly once: 1 Inbox,
 * 2 Outbox, 8 Sent Items, 9 Deleted Items, 10 Drafts, 12 Junk, 99 the account
 * root. Outbox is absent here because no script asks for one.
 */
const ROOT_SPECIAL_TYPES: Readonly<Record<string, number>> = {
    'inbox': 1,
    'sent items': 8,
    'deleted items': 9,
    'drafts': 10,
    'junk mail': 12,
    'root folder': 99,
};

/** One mail account as the profile database records it. */
export interface ProfileAccount {
    emailAccount: string;
    /** AppleScript folder id per root term — the keys of `ROOT_SPECIAL_TYPES`. */
    folderIds: Record<string, number>;
}

/** Column separator for the query output - a unit separator cannot occur in a row. */
const SEP = '\x1f';

/**
 * Address and well-known folder ids for every mail account in the profile.
 *
 * `Record_AccountUID` is a tagged id — the account's `Record_RecordID` in its low
 * 32 bits, a table tag above — so the join masks rather than assuming the tag,
 * and drops UID 0, which is the local "On My Computer" store rather than an
 * account.
 */
const ACCOUNTS_SQL = `
    SELECT a.Account_EmailAddress, f.Folder_SpecialFolderType, f.Record_RecordID
    FROM AccountsMail a
             JOIN Folders f ON (f.Record_AccountUID & 4294967295) = a.Record_RecordID
    WHERE f.Record_AccountUID <> 0
      AND a.Account_EmailAddress IS NOT NULL
      AND a.Account_EmailAddress <> ''
      AND f.Folder_SpecialFolderType IN (${Object.values(ROOT_SPECIAL_TYPES).join(', ')})
    ORDER BY a.Record_RecordID, f.Folder_SpecialFolderType`;

/** The profile whose database to read, or '' when there is none to read. */
function profileDatabase(): string {
    let names: string[];
    try {
        names = fs.readdirSync(PROFILES_DIR);
    } catch {
        return '';
    }
    // Outlook records no "default profile" anywhere reachable — not in its
    // preferences, not beside the profiles — so with more than one the newest
    // wins, after the name Outlook itself creates and all but everyone keeps.
    const dbOf = (name: string) => path.join(PROFILES_DIR, name, 'Data', 'Outlook.sqlite');
    const candidates = names
        .map(dbOf)
        .filter(db => fs.existsSync(db))
        .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
    const preferred = dbOf(DEFAULT_PROFILE);
    if (candidates.includes(preferred)) return preferred;
    return candidates[0] ?? '';
}

/** Run one read-only query, or return '' if anything at all gets in the way. */
function query(database: string, sql: string): Promise<string> {
    return new Promise(resolve => {
        execFile(
            'sqlite3',
            ['-readonly', '-list', '-noheader', '-separator', SEP, database, sql],
            { timeout: 5000, maxBuffer: 1024 * 1024 },
            (error, stdout) => resolve(error ? '' : stdout),
        );
    });
}

/**
 * Every mail account the profile knows, whether or not AppleScript can see it.
 *
 * An account whose row is missing a given root simply has no id for it, and the
 * caller keeps whatever the account probe would have done.
 */
export async function readProfileAccounts(): Promise<ProfileAccount[]> {
    const database = profileDatabase();
    if (!database) return [];
    const termOf = new Map(Object.entries(ROOT_SPECIAL_TYPES).map(([term, type]) => [type, term]));
    const byAddress = new Map<string, ProfileAccount>();
    for (const line of (await query(database, ACCOUNTS_SQL)).split('\n')) {
        const [address, specialType, folderId] = line.split(SEP);
        const term = termOf.get(Number(specialType));
        if (!address || !term || !/^\d+$/.test(folderId ?? '')) continue;
        const account = byAddress.get(address.toLowerCase())
            ?? { emailAccount: address, folderIds: {} };
        account.folderIds[term] = Number(folderId);
        byAddress.set(address.toLowerCase(), account);
    }
    return [...byAddress.values()];
}
