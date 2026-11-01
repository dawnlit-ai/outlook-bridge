// How far both platforms go when reading a folder of template emails, stated
// once so a template that reads one way on Windows can't read another on macOS.

/** How deep under the mailbox root a template folder is searched for. */
export const TEMPLATE_SEARCH_DEPTH = 3;

/** Cap on one template body; a Word-generated one is large, and a runaway one shouldn't sink the call. */
export const MAX_TEMPLATE_HTML = 100_000;

/** Characters of plain text standing in for a body that wasn't asked for. */
export const TEMPLATE_PREVIEW_CHARS = 200;
