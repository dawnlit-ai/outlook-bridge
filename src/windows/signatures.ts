// Outlook for Windows keeps each signature as "<name>.htm" (beside .rtf/.txt
// copies and a "<name>_files" folder of images) under
// %APPDATA%\Microsoft\Signatures, so these read the disk rather than COM.
import fs from 'fs';
import path from 'path';
import { pathToFileURL } from 'url';

function signaturesDir(): string {
    return process.env.APPDATA ? path.join(process.env.APPDATA, 'Microsoft', 'Signatures') : '';
}

/** The signature names — the ".htm" files — sorted. */
export async function listOutlookSignatures(): Promise<string[]> {
    const dir = signaturesDir();
    if (!dir) return [];
    try {
        return fs.readdirSync(dir)
            .filter(file => file.toLowerCase().endsWith('.htm'))
            .map(file => file.slice(0, -'.htm'.length))
            .sort((a, b) => a.localeCompare(b));
    } catch {
        return [];
    }
}

/**
 * A signature's HTML with its image references made absolute file: URIs, so an
 * email body built from it resolves and embeds the images; '' when there is no
 * signature by that name.
 */
export async function readOutlookSignatureHtml(name: string): Promise<string> {
    const dir = signaturesDir();
    if (!dir) return '';
    // A bare name only — never a path — so a crafted name can't reach outside
    // the Signatures folder.
    const file = path.join(dir, `${path.basename(name)}.htm`);
    if (!fs.existsSync(file)) return '';
    // Signatures are saved in whatever charset Outlook chose (often
    // windows-1252), so decode by the one the file declares.
    const bytes = fs.readFileSync(file);
    const charset = bytes.toString('latin1', 0, 2048).match(/charset=["']?([\w-]+)/i)?.[1] || 'utf-8';
    let html: string;
    try {
        html = new TextDecoder(charset).decode(bytes);
    } catch {
        html = bytes.toString('utf8');
    }
    // pathToFileURL encodes the directory the way Outlook's relative refs are
    // already encoded, so prefixing it keeps each reference a valid URI.
    const dirUri = pathToFileURL(dir + path.sep).href;
    return html.replace(/(src|background)=(["'])(?!https?:|cid:|data:|file:|mailto:|#)/gi, `$1=$2${dirUri}`);
}
