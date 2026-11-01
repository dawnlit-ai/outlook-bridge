// How far both platforms go when reading a folder of template emails, and what
// they strip from a body, stated once so a template that reads one way on
// Windows can't read another on macOS.

/** How deep under the mailbox root a template folder is searched for. */
export const TEMPLATE_SEARCH_DEPTH = 3;

/** Cap on one template body; a Word-generated one is large, and a runaway one shouldn't sink the call. */
export const MAX_TEMPLATE_HTML = 100_000;

/** Characters of plain text standing in for a body that wasn't asked for. */
export const TEMPLATE_PREVIEW_CHARS = 200;

/**
 * An embedded (cid:) image tag, removed from a template body as it is read.
 * Those images live as attachments on the TEMPLATE item, so its HTML reused
 * elsewhere would show broken-image placeholders; template emails are text and
 * markup only. Matched ignoring case on both platforms, since editors write
 * both `<img src=` and `<IMG SRC=`.
 */
export const EMBEDDED_IMAGE_SOURCE = '<img[^>]*src="cid:[^"]*"[^>]*>';
