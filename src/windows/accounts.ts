// The mailboxes a Windows Outlook profile can reach.
import { runPowerShellJson } from './run';
import { SESSION_PS } from './scripts';
import { strList } from '../shared/json';

/**
 * Every mailbox this profile can reach, as SMTP addresses.
 *
 * Walks Accounts AND the top-level store folders, because they are not the same
 * set: a mailbox mounted as a secondary store has no Account object, and an
 * Accounts-only walk would report a profile missing exactly the mailboxes an
 * operator added on purpose. A store contributes only when its name is an
 * address; a data file's store is named whatever it was called ("Archive") and
 * names no mailbox a caller could pass back in.
 */
export async function getOutlookAccounts(): Promise<string[]> {
    const output = await runPowerShellJson(`${SESSION_PS}
$seen = @{}
$accounts = @()
foreach ($a in $ns.Accounts) {
    $address = [string]$a.SmtpAddress
    if ($address -and -not $seen.ContainsKey($address.ToLower())) { $seen[$address.ToLower()] = $true; $accounts += $address }
}
foreach ($f in $ns.Folders) {
    $name = [string]$f.Name
    if ($name -match '@' -and -not $seen.ContainsKey($name.ToLower())) { $seen[$name.ToLower()] = $true; $accounts += $name }
}
ConvertTo-Json -Compress -InputObject @($accounts)
`, 'quick');
    return strList(output);
}
