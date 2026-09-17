// Finding an email again from what was kept about it: the rules, as pure
// functions. The search itself is an ordinary searchInboxByFilter per mailbox;
// what lives here is how narrow that search can be made, and which of its
// results is the email.
//
// The rules never guess. Every field the description gives has to agree, and at
// least two have to be given, because a subject alone names every message in a
// thread. A miss stays a miss rather than returning somebody else's email.
import { threadSubject } from '../mail';
import { InvalidRequestError } from '../errors';
import type { InboxSearchMatch } from '../types';

/** The fields a match is judged on, trimmed; blank means not given. */
export interface WantedEmail {
    readonly subject?: string;
    readonly sender?: string;
    readonly receivedTime?: string;
}

/** `yyyy-MM-dd`, optionally followed by ` HH:mm` — how both platforms stamp a received time. */
const STAMP = /^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2})?/;

/** An address inside a sender string, which may be a name, an address, or both. */
const ADDRESS = /[^\s<>,;"']+@[^\s<>,;"']+/;

const DAY_MS = 86_400_000;

const subjectKey = (subject: string): string => threadSubject(subject).toLowerCase();

/**
 * The stamp in the listings' own form (`yyyy-MM-dd HH:mm`), or '' when it isn't
 * a stamp at all — which is then treated as not given, since it can't disagree
 * with anything.
 */
function stampOf(receivedTime: string | undefined): string {
    const match = STAMP.exec((receivedTime ?? '').trim());
    return match ? match[0].replace('T', ' ') : '';
}

/** How many of the three fields the description actually gives. */
function recordedCount(wanted: WantedEmail): number {
    return (wanted.subject ? 1 : 0) + (wanted.sender ? 1 : 0) + (stampOf(wanted.receivedTime) ? 1 : 0);
}

/** INVALID_REQUEST unless the description gives enough to tell one email from its thread. */
export function requireDescribed(wanted: WantedEmail): void {
    if (recordedCount(wanted) < 2) {
        throw new InvalidRequestError(
            'Describe the email by at least two of subject, sender and receivedTime (as yyyy-MM-dd HH:mm): '
            + 'a subject alone names every message in its thread.',
        );
    }
}

/**
 * The subject glob to prefilter a search with: the longest run of plain
 * characters in the subject, or undefined when there's no run of four.
 *
 * A subject can hold an apostrophe or a `*`, and the prefilter becomes a query
 * on some stores; it only has to be selective, since every result is judged on
 * the whole subject afterwards.
 */
export function subjectPrefilter(subject: string | undefined): string | undefined {
    const longest = (threadSubject(subject ?? '').match(/[A-Za-z0-9 ._-]+/g) ?? [])
        .map(run => run.trim())
        .sort((a, b) => b.length - a.length)[0] ?? '';
    return longest.length >= 4 ? `*${longest}*` : undefined;
}

/**
 * How far back a search has to reach to cover the day the email arrived, or
 * `fallback` when no day is known. Two days of slack cover whichever timezone
 * the stamp was written in, and a search started either side of midnight.
 */
export function daysBackFor(receivedTime: string | undefined, fallback: number, now = Date.now()): number {
    const stamp = stampOf(receivedTime);
    if (!stamp) return fallback;
    const day = Date.parse(`${stamp.slice(0, 10)}T00:00:00Z`);
    if (Number.isNaN(day)) return fallback;
    return Math.max(1, Math.floor((now - day) / DAY_MS) + 2);
}

/** Whether a found email came from the sender described. */
function senderMatches(wanted: string, match: InboxSearchMatch): boolean {
    const want = wanted.toLowerCase();
    const address = ADDRESS.exec(want)?.[0];
    if (address) return address === match.senderEmail.trim().toLowerCase();
    const name = match.senderName.trim().toLowerCase();
    return name !== '' && (want.includes(name) || name.includes(want));
}

/**
 * How well a found email answers the description, or 0 for "not this one".
 *
 * Disagreement on any given field disqualifies outright. Agreement is then
 * ranked, so an exact minute beats the same day: a sender who wrote twice in
 * one afternoon still resolves to the email described.
 */
export function matchScore(wanted: WantedEmail, match: InboxSearchMatch): number {
    if (recordedCount(wanted) < 2) return 0;

    const wantSubject = wanted.subject ? subjectKey(wanted.subject) : '';
    if (wantSubject && subjectKey(match.subject) !== wantSubject) return 0;
    if (wanted.sender && !senderMatches(wanted.sender, match)) return 0;

    const stamp = stampOf(wanted.receivedTime);
    const found = match.receivedTime.trim();
    const sameDay = stamp !== '' && stamp.slice(0, 10) === found.slice(0, 10);
    const sameMinute = sameDay && stamp.length >= 16 && stamp.slice(0, 16) === found.slice(0, 16);
    if (stamp && !sameDay) return 0;

    return (wantSubject ? 1 : 0) + (wanted.sender ? 2 : 0) + (sameMinute ? 3 : sameDay ? 1 : 0);
}

/** The best match, or null when nothing agrees with the description. */
export function bestMatch(wanted: WantedEmail, matches: readonly InboxSearchMatch[]): InboxSearchMatch | null {
    let best: InboxSearchMatch | null = null;
    let bestScore = 0;
    for (const match of matches) {
        const score = matchScore(wanted, match);
        if (score > bestScore) {
            best = match;
            bestScore = score;
        }
    }
    return best;
}

/**
 * Accounts in the order to search them: the one a listing's folderPath names
 * first (`\\team@example.com\Inbox\…`), then the rest as given.
 */
export function accountSearchOrder(accounts: readonly string[], folderPath: string | undefined): string[] {
    const path = (folderPath ?? '').toLowerCase();
    if (!path) return [...accounts];
    const named = accounts.filter(account => path.includes(account.toLowerCase()));
    return [...named, ...accounts.filter(account => !named.includes(account))];
}
