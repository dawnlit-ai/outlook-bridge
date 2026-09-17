// Checking and normalizing what a caller passes, once, before either platform
// sees it.
//
// This is a security boundary as much as a convenience. Arguments end up
// inside generated scripts, and TypeScript's types vanish at run time: a
// JavaScript caller, or a tool layer passing JSON through, can hand a string
// where a count belongs. So every value is checked for its type here, every
// count is made a whole number within its documented range, and the platform
// code only ever receives values of the shape its script templates assume.
import { InvalidRequestError } from '../errors';
import type { EmailLocator, EmailRef, Recipients } from '../types';

function describe(value: unknown): string {
    if (value === null) return 'null';
    if (Array.isArray(value)) return 'an array';
    return typeof value === 'string' ? `'${value}'` : typeof value;
}

/** A required parameters object. */
export function requireObject<T extends object>(value: T | undefined | null, name: string): Partial<T> {
    if (value === null || typeof value !== 'object' || Array.isArray(value)) {
        throw new InvalidRequestError(`${name} must be an object (got ${describe(value)}).`);
    }
    return value;
}

/** An optional options object; undefined and null are an empty one. */
export function optionsObject<T extends object>(value: T | undefined | null, name = 'options'): Partial<T> {
    if (value === undefined || value === null) return {};
    return requireObject(value, name);
}

/** A required string, trimmed. Blank is an error. */
export function requiredText(value: unknown, name: string): string {
    if (typeof value !== 'string' || value.trim() === '') {
        throw new InvalidRequestError(`${name} is required and must be a non-empty string (got ${describe(value)}).`);
    }
    return value.trim();
}

/** A required string kept exactly as given — content, not an identifier. May be empty. */
export function requiredContent(value: unknown, name: string): string {
    if (typeof value !== 'string') {
        throw new InvalidRequestError(`${name} must be a string (got ${describe(value)}).`);
    }
    return value;
}

/** Optional content kept exactly as given; undefined and null mean "not given". */
export function optionalContent(value: unknown, name: string): string | undefined {
    if (value === undefined || value === null) return undefined;
    return requiredContent(value, name);
}

/** An optional regular expression. */
export function optionalRegExp(value: unknown, name: string): RegExp | undefined {
    if (value === undefined || value === null) return undefined;
    // Checked by tag rather than instanceof, so a RegExp from another realm passes.
    if (Object.prototype.toString.call(value) !== '[object RegExp]') {
        throw new InvalidRequestError(`${name} must be a RegExp (got ${describe(value)}).`);
    }
    return value as RegExp;
}

/** An optional string, trimmed; undefined, null and blank all mean "not given". */
export function optionalText(value: unknown, name: string): string | undefined {
    if (value === undefined || value === null) return undefined;
    if (typeof value !== 'string') {
        throw new InvalidRequestError(`${name} must be a string (got ${describe(value)}).`);
    }
    return value.trim() || undefined;
}

/** An optional flag; undefined and null take the fallback. */
export function flag(value: unknown, name: string, fallback: boolean): boolean {
    if (value === undefined || value === null) return fallback;
    if (typeof value !== 'boolean') {
        throw new InvalidRequestError(`${name} must be true or false (got ${describe(value)}).`);
    }
    return value;
}

/**
 * An optional count: undefined and null take the fallback, a fraction is
 * rounded down, and the result is clamped into [min, max] — the documented
 * range is a guard, not a trap. Anything that isn't a finite number is an error.
 */
export function count(value: unknown, name: string, fallback: number, min: number, max: number): number {
    if (value === undefined || value === null) return fallback;
    if (typeof value !== 'number' || !Number.isFinite(value)) {
        throw new InvalidRequestError(`${name} must be a finite number (got ${describe(value)}).`);
    }
    return Math.min(max, Math.max(min, Math.floor(value)));
}

/** A list of strings, each non-blank and trimmed. Omitted is an empty list. */
export function textList(value: unknown, name: string): string[] {
    if (value === undefined || value === null) return [];
    if (!Array.isArray(value)) {
        throw new InvalidRequestError(`${name} must be an array of strings (got ${describe(value)}).`);
    }
    return value.map((item, index) => requiredText(item, `${name}[${index}]`));
}

/** A list of names matched exactly as given — attachment file names, say. */
export function nameList(value: unknown, name: string): string[] {
    if (!Array.isArray(value)) {
        throw new InvalidRequestError(`${name} must be an array of strings (got ${describe(value)}).`);
    }
    return value.map((item, index) => {
        if (typeof item !== 'string' || item.trim() === '') {
            throw new InvalidRequestError(`${name}[${index}] must be a non-empty string (got ${describe(item)}).`);
        }
        return item;
    });
}

/** The entry ids a batch operation acts on. Duplicates are dropped; order is kept. */
export function entryIdList(value: unknown, name = 'entryIds'): string[] {
    if (!Array.isArray(value)) {
        throw new InvalidRequestError(`${name} must be an array of entry ids (got ${describe(value)}).`);
    }
    return [...new Set(value.map((id, index) => requiredText(id, `${name}[${index}]`)))];
}

/** Where an email lives, from an id or anything carrying a locator. */
export function emailLocator(email: EmailRef | undefined | null, name = 'email'): EmailLocator {
    if (typeof email === 'string') return {entryId: requiredText(email, name)};
    if (email && typeof email === 'object') {
        const entryId = requiredText(email.entryId, `${name}.entryId`);
        const storeId = optionalText(email.storeId, `${name}.storeId`);
        return storeId ? {entryId, storeId} : {entryId};
    }
    throw new InvalidRequestError(`${name} must be an entry id or an object with an entryId (got ${describe(email)}).`);
}

/** A string→string map, e.g. template placeholders. */
export function stringRecord(value: unknown, name: string): Record<string, string> | undefined {
    if (value === undefined || value === null) return undefined;
    if (typeof value !== 'object' || Array.isArray(value)) {
        throw new InvalidRequestError(`${name} must be an object of strings (got ${describe(value)}).`);
    }
    const out: Record<string, string> = {};
    for (const [key, entry] of Object.entries(value)) {
        if (typeof entry !== 'string') {
            throw new InvalidRequestError(`${name}.${key} must be a string (got ${describe(entry)}).`);
        }
        out[key] = entry;
    }
    return out;
}

/**
 * Split a recipient list into its entries.
 *
 * Separators are commas and semicolons — but not inside a quoted display name
 * or between angle brackets, where `"Doe, Jo" <jo@example.com>` has a comma
 * that belongs to the name.
 */
export function recipientList(value: Recipients | undefined | null, name: string): string[] {
    if (value === undefined || value === null) return [];
    const parts = Array.isArray(value) ? value : [value];
    const out: string[] = [];
    parts.forEach((part, index) => {
        if (typeof part !== 'string') {
            throw new InvalidRequestError(`${name}[${index}] must be a string (got ${describe(part)}).`);
        }
        out.push(...splitAddressList(part));
    });
    return out;
}

function splitAddressList(text: string): string[] {
    const entries: string[] = [];
    let current = '';
    let quoted = false;
    let angled = false;
    for (const ch of text) {
        if (ch === '"' && !angled) quoted = !quoted;
        else if (ch === '<' && !quoted) angled = true;
        else if (ch === '>' && !quoted) angled = false;
        if ((ch === ',' || ch === ';') && !quoted && !angled) {
            entries.push(current);
            current = '';
        } else {
            current += ch;
        }
    }
    entries.push(current);
    return entries.map(entry => entry.trim()).filter(Boolean);
}

/** One recipient entry as a display name and an address. */
export interface ParsedRecipient {
    name: string;
    address: string;
}

/** `"Doe, Jo" <jo@example.com>` → name + address; a bare address has no name. */
export function parseRecipient(entry: string): ParsedRecipient {
    const angled = /^(.*)<([^<>]+)>\s*$/.exec(entry);
    if (!angled) return {name: '', address: entry.trim()};
    const name = angled[1].trim().replace(/^"(.*)"$/, '$1').trim();
    return {name, address: angled[2].trim()};
}
