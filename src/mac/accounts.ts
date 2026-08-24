import { runOsaScript } from './run';
import { LIST_ACCOUNTS_SNIPPET } from './scripts';

/**
 * SMTP addresses of Outlook accounts. Empty under New Outlook (not exposed).
 * Used by the UI's account picker.
 */
export async function getOutlookAccounts(): Promise<string[]> {
    const script = `${LIST_ACCOUNTS_SNIPPET}
set AppleScript's text item delimiters to linefeed
return acctList as string`;
    const raw = await runOsaScript(script, 15000);
    return raw.split('\n').map(s => s.trim()).filter(Boolean);
}
