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
// exactly what AppleScript could see on its own. When macOS privacy control is
// what failed, `profileRefusal` says so, for the callers that would otherwise
// report a mailbox in plain sight as missing.
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
 * root. Outbox is left out: a caller can still name it (`MAC_ROOT_TERMS` maps it),
 * but nothing here files mail into one, so it is not worth an id — that request
 * falls back to the account probe like any root the profile has no row for.
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
    // Outlook rewrites these files while it runs, and the whole directory sits
    // behind macOS privacy control, so a database that existed a moment ago need
    // not stat. An unreadable timestamp sorts last rather than throwing: this
    // runs on nearly every macOS operation (resolveMacAccount), and a mailbox
    // AppleScript could have named on its own must not be lost to it.
    const mtimeOf = (db: string) => {
        try {
            return fs.statSync(db).mtimeMs;
        } catch {
            return -Infinity;
        }
    };
    const candidates = names
        .map(dbOf)
        .filter(db => fs.existsSync(db))
        .sort((a, b) => mtimeOf(b) - mtimeOf(a));
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

/** The app macOS holds this process's file access against, once asked. */
let responsible: Promise<string> | null = null;

/**
 * The app macOS privacy control asks about this process's file access, by name:
 * the outermost app bundle among its ancestors, the one launchd started. A
 * process started from a terminal or an editor is that app's to grant, not its
 * own, so naming this process would send the person to the wrong switch. '' when
 * it can't be told.
 */
function responsibleApp(): Promise<string> {
    responsible ??= new Promise(resolve => {
        execFile('ps', ['-A', '-o', 'pid=,ppid=,comm='], {
            timeout: 5000,
            maxBuffer: 4 * 1024 * 1024
        }, (error, stdout) => {
            if (error) return resolve('');
            const processes = new Map<number, { ppid: number; command: string }>();
            for (const line of stdout.split('\n')) {
                const row = /^\s*(\d+)\s+(\d+)\s+(.*)$/.exec(line);
                if (row) processes.set(Number(row[1]), { ppid: Number(row[2]), command: row[3] ?? '' });
            }
            let app = '';
            const climbed = new Set<number>();
            for (let pid = process.pid; pid > 1 && !climbed.has(pid);) {
                climbed.add(pid);
                const entry = processes.get(pid);
                if (!entry) break;
                // The first bundle in a path is the outer one: a helper app sits
                // inside its parent's. Each one higher up replaces the last.
                const bundle = /\/([^/]+)\.app(?:\/|$)/.exec(entry.command)?.[1];
                if (bundle) app = bundle;
                pid = entry.ppid;
            }
            resolve(app);
        });
    });
    return responsible;
}

/**
 * Why Outlook's profile can't be read, when macOS privacy control is the
 * reason: the sentence to act on, naming the app to grant access to. null when
 * the profile can be read, or simply isn't there — New Outlook, or no Outlook
 * profile on this Mac.
 *
 * Everything else in this module stays silent on purpose, which is right for an
 * enrichment pass and wrong when the pass was the only way to see a mailbox:
 * the caller then hears "no accounts" or "not found" about a mailbox in plain
 * sight in Outlook. The callers that would say that ask here first.
 */
export async function profileRefusal(): Promise<string | null> {
    try {
        fs.readdirSync(PROFILES_DIR);
        return null;
    } catch (error) {
        const code = (error as NodeJS.ErrnoException).code;
        if (code !== 'EPERM' && code !== 'EACCES') return null;
    }
    const app = await responsibleApp();
    const who = app || 'the app this runs in (the terminal or editor it was started from, if it was)';
    const it = app || 'that app';
    // Plain ASCII: the sentence also travels inside a generated script's string literal.
    return `macOS is not letting ${who} read Outlook's profile, the only place Outlook for Mac lists `
        + "a Microsoft 365 mailbox it syncs through Microsoft's cloud. "
        + `Turn on ${it} under System Settings > Privacy & Security > Full Disk Access, then quit ${it} and open it again.`;
}

/**
 * Every mail account the profile knows, whether or not AppleScript can see it.
 *
 * An account whose row is missing a given root simply has no id for it, and the
 * caller keeps whatever the account probe would have done.
 */
export async function readProfileAccounts(): Promise<ProfileAccount[]> {
    try {
        return await readAccounts();
    } catch {
        // The module's contract, enforced at its one exit: this is an enrichment
        // pass over what AppleScript already found, so nothing that happens in it
        // is worth failing an operation that would otherwise have worked.
        return [];
    }
}

async function readAccounts(): Promise<ProfileAccount[]> {
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
