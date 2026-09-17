// Coercing the JSON a PowerShell script prints into the package's types.
//
// ConvertTo-Json has two habits a caller must never see: it unwraps a
// single-element array into a bare object, and it renders an empty one as
// nothing at all. Every list crossing that boundary goes through `toArray`, and
// every scalar through a coercion that survives a missing property.

/** One value as a list: an array stays, null and undefined become empty, anything else wraps. */
export function toArray(value: unknown): unknown[] {
    return Array.isArray(value) ? value : value == null ? [] : [value];
}

/** A value as a plain object, for reading fields off one row. */
export function record(value: unknown): Record<string, unknown> {
    return value !== null && typeof value === 'object' && !Array.isArray(value)
        ? value as Record<string, unknown>
        : {};
}

export function str(value: unknown): string {
    return value == null ? '' : String(value);
}

export function num(value: unknown): number {
    return typeof value === 'number' && Number.isFinite(value) ? value : 0;
}

export function bool(value: unknown): boolean {
    return value === true;
}

/** A list of non-empty strings. */
export function strList(value: unknown): string[] {
    return toArray(value).map(str).filter(Boolean);
}

/** A base64 field decoded as UTF-8 — how scripts pass bodies that must survive intact. */
export function base64Text(value: unknown): string {
    return Buffer.from(str(value), 'base64').toString('utf8');
}

/** A list of per-item failures, as the scripts report them. */
export function itemFailures(value: unknown): { entryId: string; subject: string; error: string }[] {
    return toArray(value).map(row => {
        const e = record(row);
        return {entryId: str(e.entryId), subject: str(e.subject), error: str(e.error)};
    });
}
