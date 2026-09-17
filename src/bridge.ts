// The public operations, implemented once for both platforms.
//
// Each operation checks its arguments, applies its defaults, does whatever work
// is the same everywhere — composing a reply from a template, splitting a body
// from its quoted thread, choosing where attachments go — and hands a finished
// request to this machine's backend. Keeping that here means a default, a
// range or a validation rule can't come to differ between Windows and macOS.
import fs from 'fs';
import path from 'path';
import { windowsBackend } from './windows';
import { macBackend } from './mac';
import type { Backend, Disposition } from './backend';
import {
    count,
    emailLocator,
    entryIdList,
    flag,
    nameList,
    optionalContent,
    optionalRegExp,
    optionalText,
    optionsObject,
    recipientList,
    requiredContent,
    requiredText,
    requireObject,
    stringRecord,
    textList,
} from './shared/args';
import { composeReplyHtml } from './shared/replyBody';
import { folderOption, splitQuotedOriginal } from './mail';
import { InvalidRequestError, NotFoundError, UnsupportedPlatformError } from './errors';
import {
    attachmentDestination,
    type BridgeOptions,
    getGlobalConfig,
    mergeOptions,
    type ResolvedConfig,
    withConfig,
} from './runtime';
import type { BridgeCapability, CapabilityMap, OutlookBridge, SavedAttachment } from './types';

/** Every default, stated once. */
const DEFAULTS = {
    daysBack: 60,
    inboxLimit: 50,
    previewChars: 600,
    bodyMaxChars: 8000,
    quotedMaxChars: 4000,
    folderDepth: 2,
    draftsLimit: 100,
    draftsPreviewChars: 300,
    bounceDaysBack: 30,
    sentGroupsLimit: 3000,
    templateFolder: 'Templates',
    templatesLimit: 20,
} as const;

/** The widest window any scan accepts, in days. */
const MAX_DAYS = 3650;

/**
 * Every operation name. `satisfies` makes this list fail to compile the moment
 * it and the OutlookBridge interface disagree, in either direction.
 */
const OPERATION_NAMES = {
    getOutlookAccounts: true,
    sendOutlookEmail: true,
    replyOutlookEmail: true,
    readInboxEmails: true,
    searchInboxByFilter: true,
    readSelectedEmail: true,
    readEmailBody: true,
    openOutlookEmail: true,
    listInboxFolders: true,
    moveOutlookEmails: true,
    listOutlookDrafts: true,
    sendDrafts: true,
    sendAllDrafts: true,
    deleteOutlookDrafts: true,
    deleteOutlookEmails: true,
    purgeDeletedItems: true,
    saveEmailAttachment: true,
    saveEmailAttachments: true,
    cleanUndeliverableEmails: true,
    collectBouncedRecipients: true,
    readSentRecipientGroups: true,
    listOutlookSignatures: true,
    readOutlookSignatureHtml: true,
    readTemplateEmails: true,
    saveTemplateEmail: true,
    editEmailTemplate: true,
} satisfies Record<BridgeCapability, true>;

export const OPERATIONS = Object.freeze(Object.keys(OPERATION_NAMES) as BridgeCapability[]);

/** This machine's backend, or null where Outlook can't be automated. */
function platformBackend(): Backend | null {
    switch (process.platform) {
        case 'win32':
            return windowsBackend;
        case 'darwin':
            return macBackend;
        default:
            return null;
    }
}

const PLATFORM_BACKEND = platformBackend();

function dispositionOf(sendImmediately: unknown, openDraftWindow: unknown): Disposition {
    if (flag(sendImmediately, 'sendImmediately', false)) return 'send';
    return flag(openDraftWindow, 'openDraftWindow', true) ? 'display' : 'save';
}

/** Attachment paths as absolute paths to files that exist — Outlook resolves a relative path against its own directory, not the caller's. */
function attachmentFiles(value: unknown): string[] {
    return textList(value, 'attachments').map(file => {
        const absolute = path.resolve(file);
        let isFile = false;
        try {
            isFile = fs.statSync(absolute).isFile();
        } catch {
            // Reported below.
        }
        if (!isFile) throw new NotFoundError('file', `Attachment file '${absolute}' does not exist.`);
        return absolute;
    });
}

/**
 * The operations over one backend. A null backend is a platform without Outlook
 * automation: every operation rejects with UNSUPPORTED_PLATFORM before looking
 * at its arguments, since nothing about them could make it work.
 */
export function createOperations(backendOrNull: Backend | null): OutlookBridge {
    const backend = (): Backend => {
        if (!backendOrNull) throw new UnsupportedPlatformError(process.platform);
        return backendOrNull;
    };

    async function saveEmailAttachments(
        email: Parameters<OutlookBridge['saveEmailAttachments']>[0],
        fileNames: readonly string[],
        options?: Parameters<OutlookBridge['saveEmailAttachments']>[2],
    ): Promise<SavedAttachment[]> {
        const platform = backend();
        const target = emailLocator(email);
        const names = nameList(fileNames, 'fileNames');
        const destDir = optionalText(optionsObject(options).destDir, 'destDir');
        if (names.length === 0) return [];
        const destination = attachmentDestination(destDir === undefined ? undefined : path.resolve(destDir));
        try {
            return await platform.saveEmailAttachments({email: target, fileNames: names, destDir: destination.dir});
        } catch (error) {
            destination.discardIfUnused();
            throw error;
        }
    }

    return {
        async getOutlookAccounts() {
            return backend().getOutlookAccounts();
        },

        async sendOutlookEmail(params) {
            const platform = backend();
            const p = requireObject(params, 'params');
            const request = {
                account: requiredText(p.emailAccount, 'emailAccount'),
                to: recipientList(p.to, 'to'),
                cc: recipientList(p.cc, 'cc'),
                bcc: recipientList(p.bcc, 'bcc'),
                subject: requiredContent(p.subject, 'subject'),
                htmlBody: requiredContent(p.htmlBody, 'htmlBody'),
                attachments: attachmentFiles(p.attachments),
                disposition: dispositionOf(p.sendImmediately, p.openDraftWindow),
            };
            if (request.disposition === 'send' && request.to.length + request.cc.length + request.bcc.length === 0) {
                throw new InvalidRequestError('An email sent immediately needs at least one recipient.');
            }
            await platform.sendOutlookEmail(request);
        },

        async replyOutlookEmail(params) {
            const platform = backend();
            const p = requireObject(params, 'params');
            const account = requiredText(p.emailAccount, 'emailAccount');
            const email = emailLocator({entryId: p.entryId as string, storeId: p.storeId}, 'params');
            const html = await composeReplyHtml({
                account,
                htmlBody: optionalContent(p.htmlBody, 'htmlBody'),
                templateSubject: optionalText(p.templateSubject, 'templateSubject'),
                templateFolder: optionalText(p.templateFolder, 'templateFolder') ?? DEFAULTS.templateFolder,
                templateSection: optionalText(p.templateSection, 'templateSection'),
                templatePlaceholders: stringRecord(p.templatePlaceholders, 'templatePlaceholders'),
                signatureName: optionalText(p.signatureName, 'signatureName'),
            }, platform);
            return platform.replyOutlookEmail({
                account,
                email,
                html,
                replyAll: flag(p.replyAll, 'replyAll', false),
                disposition: dispositionOf(p.sendImmediately, p.openDraftWindow),
            });
        },

        async readInboxEmails(emailAccount, options) {
            const platform = backend();
            const o = optionsObject(options);
            const {ref, label} = folderOption(optionalText(o.folder, 'folder'));
            return platform.readInboxEmails({
                account: requiredText(emailAccount, 'emailAccount'),
                folder: ref,
                folderLabel: label,
                daysBack: count(o.daysBack, 'daysBack', DEFAULTS.daysBack, 1, MAX_DAYS),
                limit: count(o.limit, 'limit', DEFAULTS.inboxLimit, 1, 100_000),
                previewChars: count(o.previewChars, 'previewChars', DEFAULTS.previewChars, 0, 1_000_000),
            });
        },

        async searchInboxByFilter(emailAccount, filter) {
            const platform = backend();
            const f = optionsObject(filter, 'filter');
            return platform.searchInboxByFilter({
                account: requiredText(emailAccount, 'emailAccount'),
                daysBack: count(f.daysBack, 'daysBack', DEFAULTS.daysBack, 1, MAX_DAYS),
                subjectLike: optionalText(f.subjectLike, 'subjectLike'),
                subjectPattern: optionalRegExp(f.subjectPattern, 'subjectPattern'),
                excludeReplies: flag(f.excludeReplies, 'excludeReplies', false),
                requireAttachment: flag(f.requireAttachment, 'requireAttachment', false),
                includeFolders: textList(f.includeFolders, 'includeFolders'),
                excludeFolders: textList(f.excludeFolders, 'excludeFolders'),
                includeBody: flag(f.includeBody, 'includeBody', true),
            });
        },

        async readSelectedEmail() {
            return backend().readSelectedEmail();
        },

        async readEmailBody(email, options) {
            const platform = backend();
            const target = emailLocator(email);
            const o = optionsObject(options);
            const maxChars = count(o.maxChars, 'maxChars', DEFAULTS.bodyMaxChars, 1, 100_000_000);
            const maxQuotedChars = count(o.maxQuotedChars, 'maxQuotedChars', DEFAULTS.quotedMaxChars, 0, 100_000_000);
            const includeQuoted = flag(o.includeQuoted, 'includeQuoted', false);
            const raw = await platform.readEmailBody(target);
            // The quoted thread is context, never the sender's own answer, so it is
            // capped by the tighter of the two limits.
            const {body, quoted, separator} = splitQuotedOriginal(raw.body);
            return {
                entryId: raw.entryId,
                subject: raw.subject,
                senderName: raw.senderName,
                senderEmail: raw.senderEmail,
                receivedTime: raw.receivedTime,
                body: body.slice(0, maxChars),
                truncated: body.length > maxChars,
                bodyLength: body.length,
                quoteSeparator: separator,
                quotedLength: quoted.length,
                quotedOriginal: includeQuoted ? quoted.slice(0, Math.min(maxChars, maxQuotedChars)) : '',
                attachmentNames: [...raw.attachmentNames],
                attachmentCount: raw.attachmentNames.length,
            };
        },

        async openOutlookEmail(email) {
            const platform = backend();
            await platform.openOutlookEmail(emailLocator(email));
        },

        async listInboxFolders(emailAccount, options) {
            const platform = backend();
            const o = optionsObject(options);
            return platform.listInboxFolders({
                account: requiredText(emailAccount, 'emailAccount'),
                maxDepth: count(o.maxDepth, 'maxDepth', DEFAULTS.folderDepth, 1, 10),
            });
        },

        async moveOutlookEmails(emailAccount, entryIds, folder, options) {
            const platform = backend();
            const account = requiredText(emailAccount, 'emailAccount');
            const ids = entryIdList(entryIds);
            const {ref, label} = folderOption(requiredText(folder, 'folder'));
            const createIfMissing = flag(optionsObject(options).createIfMissing, 'createIfMissing', false);
            if (ids.length === 0) return {folderPath: '', folderCreated: false, moved: 0, failed: []};
            return platform.moveOutlookEmails({
                account,
                entryIds: ids,
                folder: ref,
                folderLabel: label,
                createIfMissing
            });
        },

        async listOutlookDrafts(emailAccount, options) {
            const platform = backend();
            const o = optionsObject(options);
            return platform.listOutlookDrafts({
                account: requiredText(emailAccount, 'emailAccount'),
                limit: count(o.limit, 'limit', DEFAULTS.draftsLimit, 1, 100_000),
                previewChars: count(o.previewChars, 'previewChars', DEFAULTS.draftsPreviewChars, 0, 1_000_000),
            });
        },

        async sendDrafts(emailAccount, entryIds) {
            const platform = backend();
            const account = requiredText(emailAccount, 'emailAccount');
            const ids = entryIdList(entryIds);
            if (ids.length === 0) return {sent: 0, failed: []};
            return platform.sendDrafts({account, entryIds: ids});
        },

        async sendAllDrafts(emailAccount) {
            const platform = backend();
            return platform.sendAllDrafts({account: requiredText(emailAccount, 'emailAccount')});
        },

        async deleteOutlookDrafts(emailAccount, entryIds) {
            const platform = backend();
            const account = requiredText(emailAccount, 'emailAccount');
            const ids = entryIdList(entryIds);
            if (ids.length === 0) return {deleted: 0, failed: []};
            return platform.deleteOutlookDrafts({account, entryIds: ids});
        },

        async deleteOutlookEmails(emailAccount, entryIds, options) {
            const platform = backend();
            const account = requiredText(emailAccount, 'emailAccount');
            const ids = entryIdList(entryIds);
            const o = optionsObject(options);
            const allowProtected = flag(o.allowProtected, 'allowProtected', false);
            const dryRun = flag(o.dryRun, 'dryRun', false);
            if (ids.length === 0) return {dryRun, deleted: 0, refused: 0, failed: 0, items: []};
            return platform.deleteOutlookEmails({account, entryIds: ids, allowProtected, dryRun});
        },

        async purgeDeletedItems(emailAccount, options) {
            const platform = backend();
            const o = optionsObject(options);
            return platform.purgeDeletedItems({
                account: requiredText(emailAccount, 'emailAccount'),
                olderThanDays: count(o.olderThanDays, 'olderThanDays', 0, 0, 36_500),
                dryRun: flag(o.dryRun, 'dryRun', false),
            });
        },

        async saveEmailAttachment(email, fileName, options) {
            const [saved] = await saveEmailAttachments(email, [fileName], options);
            if (!saved) throw new NotFoundError('attachment', `Attachment '${fileName}' was not saved.`);
            return saved;
        },

        saveEmailAttachments,

        async cleanUndeliverableEmails(emailAccount, options) {
            const platform = backend();
            const o = optionsObject(options);
            return platform.cleanUndeliverableEmails({
                account: requiredText(emailAccount, 'emailAccount'),
                daysBack: count(o.daysBack, 'daysBack', DEFAULTS.bounceDaysBack, 1, MAX_DAYS),
                // Deleting has to be asked for.
                dryRun: flag(o.dryRun, 'dryRun', true),
            });
        },

        async collectBouncedRecipients(emailAccount, options) {
            const platform = backend();
            const o = optionsObject(options);
            return platform.collectBouncedRecipients({
                account: requiredText(emailAccount, 'emailAccount'),
                daysBack: count(o.daysBack, 'daysBack', DEFAULTS.bounceDaysBack, 1, MAX_DAYS),
                includeDeletedItems: flag(o.includeDeletedItems, 'includeDeletedItems', true),
            });
        },

        async readSentRecipientGroups(emailAccount, options) {
            const platform = backend();
            const o = optionsObject(options);
            return platform.readSentRecipientGroups({
                account: requiredText(emailAccount, 'emailAccount'),
                daysBack: count(o.daysBack, 'daysBack', DEFAULTS.bounceDaysBack, 1, MAX_DAYS),
                limit: count(o.limit, 'limit', DEFAULTS.sentGroupsLimit, 1, 100_000),
            });
        },

        async listOutlookSignatures() {
            return backend().listOutlookSignatures();
        },

        async readOutlookSignatureHtml(name) {
            const platform = backend();
            return platform.readOutlookSignatureHtml(requiredText(name, 'name'));
        },

        async readTemplateEmails(emailAccount, options) {
            const platform = backend();
            const o = optionsObject(options);
            return platform.readTemplateEmails({
                account: requiredText(emailAccount, 'emailAccount'),
                folder: optionalText(o.folder, 'folder') ?? DEFAULTS.templateFolder,
                limit: count(o.limit, 'limit', DEFAULTS.templatesLimit, 1, 200),
                includeBody: flag(o.includeBody, 'includeBody', true),
                subject: optionalText(o.subject, 'subject'),
            });
        },

        async saveTemplateEmail(emailAccount, template) {
            const platform = backend();
            const account = requiredText(emailAccount, 'emailAccount');
            const t = requireObject(template, 'template');
            return platform.saveTemplateEmail({
                account,
                folder: optionalText(t.folder, 'folder') ?? DEFAULTS.templateFolder,
                subject: requiredText(t.subject, 'subject'),
                htmlBody: requiredContent(t.htmlBody, 'htmlBody'),
            });
        },

        async editEmailTemplate(label, currentHtml) {
            const platform = backend();
            return platform.editEmailTemplate({
                label: requiredText(label, 'label'),
                html: requiredContent(currentHtml, 'currentHtml'),
            });
        },
    };
}

// ── Instances and capabilities ───────────────────────────────────────────

/**
 * A bridge carrying its own settings, plus capability discovery.
 *
 * The exported functions are the same operations under the process-wide
 * settings; an instance is what to use when the process holds more than one
 * caller, or when a call needs its own timeout or AbortSignal.
 */
export interface OutlookBridgeInstance extends OutlookBridge {
    /** The settings this bridge runs with. */
    readonly options: ResolvedConfig;

    /** Which operations work on this machine. */
    capabilities(): CapabilityMap;

    /** Whether one operation works on this machine. */
    supports(operation: BridgeCapability): boolean;

    /**
     * A bridge like this one with some settings changed — how one call gets its
     * own budget or cancellation: `bridge.withOptions({ signal }).readInboxEmails(...)`.
     */
    withOptions(options: BridgeOptions): OutlookBridgeInstance;
}

function capabilityMap(backend: Backend | null): CapabilityMap {
    return Object.freeze(Object.fromEntries(OPERATIONS.map(name => [name, backend !== null]))) as CapabilityMap;
}

type AnyOperation = (...args: unknown[]) => Promise<unknown>;

/**
 * A bridge over `backend` whose every call runs with `config` in force.
 *
 * The platform code reads settings through `getConfig()`, so the config is put
 * in place around each call rather than passed into it: `withConfig` scopes it
 * to the call's whole async subtree, which is why two bridges with different
 * settings can run concurrently without seeing each other's.
 */
export function createBridge(backend: Backend | null, config: ResolvedConfig): OutlookBridgeInstance {
    const operations = createOperations(backend) as unknown as Record<BridgeCapability, AnyOperation>;
    const map = capabilityMap(backend);
    const instance: Record<string, unknown> = {
        options: Object.freeze({...config}),
        capabilities: () => map,
        supports: (operation: BridgeCapability) => map[operation] === true,
        withOptions: (options: BridgeOptions) => createBridge(backend, mergeOptions(config, options)),
    };
    for (const name of OPERATIONS) {
        instance[name] = (...args: unknown[]) => withConfig(config, () => operations[name](...args));
    }
    return instance as unknown as OutlookBridgeInstance;
}

/**
 * A bridge with its own settings, isolated from `configure()` and from every
 * other bridge. Options left out take the process-wide values as they stand
 * when this is called.
 *
 * ```ts
 * const outlook = createOutlookBridge({ timeoutMs: 30_000 });
 * const recent = await outlook.readInboxEmails('me@example.com', { daysBack: 7 });
 * ```
 */
export function createOutlookBridge(options: BridgeOptions = {}): OutlookBridgeInstance {
    return createBridge(PLATFORM_BACKEND, mergeOptions(getGlobalConfig(), options));
}

const PLATFORM_CAPABILITIES = capabilityMap(PLATFORM_BACKEND);

/** Which operations work on this machine: all of them on Windows and macOS, none elsewhere. */
export function capabilities(): CapabilityMap {
    return PLATFORM_CAPABILITIES;
}

/** Whether one operation works on this machine. */
export function supports(operation: BridgeCapability): boolean {
    return PLATFORM_CAPABILITIES[operation] === true;
}

/** The operations under the process-wide settings — what the package exports as plain functions. */
export const defaultBridge: OutlookBridge = createOperations(PLATFORM_BACKEND);
