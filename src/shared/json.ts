// Coercion helpers for the JSON the PowerShell scripts emit.
//
// PowerShell's ConvertTo-Json has two habits that a consumer must never see:
// it unwraps a single-element array to a bare object, and renders an empty one
// as null. Every list crossing that boundary therefore goes through `toArray`,
// and every scalar through a coercion that survives a missing property.

/** Parse a script's stdout as an object, tolerating empty output and 'null'. */
export function parseObject(raw: string): Record<string, unknown> {
    if (!raw || !raw.trim() || raw.trim() === 'null') return {};
    return JSON.parse(raw.trim()) as Record<string, unknown>;
}

/** Parse a script's stdout as a list, absorbing ConvertTo-Json's unwrapping. */
export function parseArray(raw: string): unknown[] {
    if (!raw || !raw.trim() || raw.trim() === 'null') return [];
    return toArray(JSON.parse(raw.trim()));
}

/** One value as a list: an array stays, null becomes empty, anything else wraps. */
export function toArray(value: unknown): unknown[] {
    return Array.isArray(value) ? value : value == null ? [] : [value];
}

/** A record's field as a plain object, for mapping a list of rows. */
export function record(value: unknown): Record<string, unknown> {
    return (value ?? {}) as Record<string, unknown>;
}

export function str(value: unknown): string {
    return value == null ? '' : String(value);
}

export function num(value: unknown): number {
    return typeof value === 'number' ? value : 0;
}

export function strList(value: unknown): string[] {
    return toArray(value).map(str).filter(Boolean);
}
