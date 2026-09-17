// The mailboxes Outlook for Mac can reach.
import { runOsaScript } from './run';
import { readProfileAccounts } from './profile';
import { LIST_ACCOUNTS_SNIPPET } from './scripts';

/**
 * Every mailbox's SMTP address. Empty under New Outlook, which doesn't expose
 * accounts to AppleScript.
 *
 * The dictionary's three account classes aren't the whole profile: a Microsoft
 * 365 mailbox added by the modern sync engine belongs to none of them. So the
 * profile database is asked as well and the answers merged — the script's
 * first, since those are the accounts every operation supports fully.
 */
export async function getOutlookAccounts(): Promise<string[]> {
    const [raw, profile] = await Promise.all([
        runOsaScript(`${LIST_ACCOUNTS_SNIPPET}
set AppleScript's text item delimiters to linefeed
return acctList as string`, 'quick'),
        readProfileAccounts(),
    ]);
    const accounts = raw.split('\n').map(s => s.trim()).filter(Boolean);
    const seen = new Set(accounts.map(a => a.toLowerCase()));
    for (const {emailAccount} of profile) {
        if (seen.has(emailAccount.toLowerCase())) continue;
        seen.add(emailAccount.toLowerCase());
        accounts.push(emailAccount);
    }
    return accounts;
}
