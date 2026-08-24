import { runPowerShell } from './run';

/** Retrieve all Outlook email accounts (SMTP addresses). */
export async function getOutlookAccounts(): Promise<string[]> {
    if (process.platform !== 'win32') return [];
    const script = `
$outlook = New-Object -ComObject Outlook.Application
$accounts = @()
foreach ($a in $outlook.Session.Accounts) {
    $accounts += $a.SmtpAddress
}
$accounts -join '|'
`;
    const result = await runPowerShell(script);
    if (!result) return [];
    return result.split('|').filter(Boolean);
}
