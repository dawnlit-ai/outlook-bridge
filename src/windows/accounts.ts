import { runPowerShell } from './run';

/**
 * Retrieve every mailbox this profile can reach, as SMTP addresses.
 *
 * Walks Accounts AND Stores, because they are not the same set. A mailbox
 * opened as a secondary store — a shared mailbox, or a delegated one — has no
 * `Account` object at all, so an Accounts-only walk reports a profile that is
 * missing exactly the mailboxes an operator added on purpose. Worse, it reports
 * them silently: the caller scopes a read to the one address it was given and
 * gets somebody else's inbox back, with no error to notice.
 *
 * A store contributes only when its name looks like an address. Outlook names a
 * store by its display name, which for a mailbox is its address but for a data
 * file is whatever it was called — "Archive", "Personal Folders" — and those
 * name no mailbox a caller could pass back in.
 */
export async function getOutlookAccounts(): Promise<string[]> {
    if (process.platform !== 'win32') return [];
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('mapi')
$ns.Logon()
$seen = @{}
$out = @()
foreach ($a in $ns.Accounts) {
    if ($a.SmtpAddress) {
        $k = $a.SmtpAddress.ToLower()
        if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $out += $a.SmtpAddress }
    }
}
foreach ($f in $ns.Folders) {
    $n = $f.Name
    if ($n -and $n -match '@') {
        $k = $n.ToLower()
        if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $out += $n }
    }
}
$out -join '|'
`;
    const result = await runPowerShell(script);
    if (!result) return [];
    return result.split('|').filter(Boolean);
}
