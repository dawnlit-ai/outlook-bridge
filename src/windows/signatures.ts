// Outlook stores each signature as "<name>.htm" (plus .rtf/.txt and a
// "<name>_files" folder for images) under %APPDATA%\Microsoft\Signatures, so
// both of these read the disk rather than driving COM.
import fs from 'fs';
import path from 'path';
import { pathToFileURL } from 'url';

const SIGNATURES_DIR = process.env.APPDATA
    ? path.join(process.env.APPDATA, 'Microsoft', 'Signatures')
    : '';

/**
 * Names of the user's Outlook signatures (the ".htm" files), sorted. Empty off
 * Windows. Reads from disk synchronously but is declared async to match macOS,
 * which has to ask Outlook itself — one signature for both platforms.
 */
export async function listOutlookSignatures(): Promise<string[]> {
    if (process.platform !== 'win32' || !SIGNATURES_DIR) return [];
    try {
        return fs.readdirSync(SIGNATURES_DIR)
            .filter(f => f.toLowerCase().endsWith('.htm'))
            .map(f => f.slice(0, -4))
            .sort((a, b) => a.localeCompare(b));
    } catch {
        return [];
    }
}

/**
 * Read a named signature's HTML with its image references rewritten to absolute
 * file: URIs. Outlook stores signature images relative to a "<name>_files"
 * folder; once the refs are absolute, assigning the HTML to a mail body lets
 * Outlook resolve and embed the images on Display/Send. Returns '' if the
 * signature can't be found.
 */
export async function readOutlookSignatureHtml(name: string): Promise<string> {
    if (process.platform !== 'win32' || !SIGNATURES_DIR) return '';
    // Only accept a bare signature name — never a path — so a crafted name can't
    // escape the Signatures folder.
    const safe = path.basename(name);
    const file = path.join(SIGNATURES_DIR, `${safe}.htm`);
    if (!fs.existsSync(file)) return '';
    // Classic Outlook signatures are saved as windows-1252, not UTF-8; decode by
    // the charset the file declares so accented text / smart quotes survive.
    const buf = fs.readFileSync(file);
    const head = buf.toString('latin1', 0, 2048);
    const charset = head.match(/charset=["']?([\w-]+)/i)?.[1] || 'utf-8';
    let html: string;
    try {
        html = new TextDecoder(charset).decode(buf);
    } catch {
        html = buf.toString('utf-8');
    }
    // pathToFileURL encodes spaces/specials the same way Outlook's relative refs
    // are, so prefixing the (already relative, already-encoded) src keeps a valid URI.
    const dirUri = pathToFileURL(SIGNATURES_DIR + path.sep).href;
    html = html.replace(
        /(src|background)=(["'])(?!https?:|cid:|data:|file:|mailto:|#)/gi,
        `$1=$2${dirUri}`,
    );
    return html;
}
