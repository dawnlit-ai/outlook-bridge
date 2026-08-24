// Unlike Windows (files on disk), Outlook for Mac serves signature content
// directly through AppleScript, images included (inline data or remote refs).
import { asEscape, runOsaScript } from './run';

/** Names of the user's Outlook signatures. */
export async function listOutlookSignatures(): Promise<string[]> {
    const script = `
set sigNames to {}
tell application "Microsoft Outlook"
    try
        repeat with s in signatures
            set end of sigNames to (name of s as string)
        end repeat
    end try
end tell
set AppleScript's text item delimiters to linefeed
return sigNames as string`;
    const raw = await runOsaScript(script, 15000);
    return raw.split('\n').map(s => s.trim()).filter(Boolean).sort((a, b) => a.localeCompare(b));
}

/** Read a named signature's HTML. Returns '' if the signature can't be found. */
export async function readOutlookSignatureHtml(name: string): Promise<string> {
    const script = `
tell application "Microsoft Outlook"
    try
        return content of (first signature whose name is "${asEscape(name)}")
    on error
        return ""
    end try
end tell`;
    return runOsaScript(script, 15000);
}
