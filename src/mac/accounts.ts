import { runOsaScript } from './run';
import { readProfileAccounts } from './profile';
import { LIST_ACCOUNTS_SNIPPET } from './scripts';

/**
 * SMTP addresses of Outlook accounts. Empty under New Outlook (not exposed).
 * Used by the UI's account picker.
 *
 * The dictionary's three account classes are not the whole profile: a Microsoft
 * 365 mailbox added by the modern sync engine belongs to none of them and is
 * simply absent from the script's answer, which showed an operator with two
 * mailboxes a picker holding one. So the profile database is asked as well, and
 * the two answers are merged. The script's come first — those are the accounts
 * every operation here supports fully.
 */
export async function getOutlookAccounts(): Promise<string[]> {
    const script = `${LIST_ACCOUNTS_SNIPPET}
set AppleScript's text item delimiters to linefeed
return acctList as string`;
    const [raw, profile] = await Promise.all([
        runOsaScript(script, 15000),
        readProfileAccounts(),
    ]);
    const accounts = raw.split('\n').map(s => s.trim()).filter(Boolean);
    const seen = new Set(accounts.map(a => a.toLowerCase()));
    for (const { emailAccount } of profile) {
        if (seen.has(emailAccount.toLowerCase())) continue;
        seen.add(emailAccount.toLowerCase());
        accounts.push(emailAccount);
    }
    return accounts;
}
