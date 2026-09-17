// Outlook for Mac serves signatures through AppleScript, images included
// (inline data or remote references), where Windows keeps them as files.
import { asString, runOsaScript } from './run';

/** The signature names, sorted. */
export async function listOutlookSignatures(): Promise<string[]> {
    const raw = await runOsaScript(`
set sigNames to {}
tell application "Microsoft Outlook"
    try
        repeat with s in signatures
            set end of sigNames to (name of s as string)
        end repeat
    end try
end tell
set AppleScript's text item delimiters to linefeed
return sigNames as string`, 'quick');
    return raw.split('\n').map(s => s.trim()).filter(Boolean).sort((a, b) => a.localeCompare(b));
}

/** A signature's HTML; '' when there is no signature by that name. */
export async function readOutlookSignatureHtml(name: string): Promise<string> {
    return runOsaScript(`
tell application "Microsoft Outlook"
    try
        return content of (first signature whose name is ${asString(name)})
    on error
        return ""
    end try
end tell`, 'quick');
}
