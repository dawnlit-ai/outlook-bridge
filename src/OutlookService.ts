// Platform dispatcher for Outlook automation: Windows drives Outlook via
// PowerShell + COM, macOS via AppleScript (see each implementation's index for
// its platform notes — New Outlook for Mac supports composing but not mailbox
// reading). Other platforms report no capabilities and every call rejects.
//
// Both implementations are pinned to `OutlookBridge` here rather than being
// destructured straight off the module namespaces. That is the difference
// between a consumer seeing `Promise<CleanUndeliverableResult>` and seeing
// `Promise<CleanUndeliverableResult | { matched: unknown[] }>` — a union of
// whatever the two platforms happened to declare, which is what this used to
// publish.
import * as windowsBridge from './windows';
import * as macBridge from './mac';
import type { BridgeOptions, ResolvedConfig } from './runtime';
import { getGlobalConfig, mergeOptions, withConfig } from './runtime';
import { UnsupportedPlatformError } from './errors';
import type { BridgeCapability, CapabilityMap, OutlookBridge } from './types';

/** Every operation name, taken from the Windows bridge — the complete one. */
const ALL_CAPABILITIES = Object.keys(windowsBridge.bridge) as BridgeCapability[];

/**
 * The backend for an OS with no Outlook automation at all.
 *
 * A third implementation rather than a fallback to the PowerShell one: making
 * Windows the default for Linux and everything else meant the dispatcher held a
 * bridge whose capability map had to be second-guessed afterwards, and whose
 * functions each had to remember to refuse. Answering "not here" is itself an
 * implementation of the contract, so it is written as one, once.
 */
const unsupportedBridge = {
    bridge: Object.fromEntries(ALL_CAPABILITIES.map(name => [
        name,
        () => Promise.reject(new UnsupportedPlatformError(process.platform)),
    ])) as unknown as OutlookBridge,
    capabilities: Object.freeze(
        Object.fromEntries(ALL_CAPABILITIES.map(name => [name, false])),
    ) as CapabilityMap,
};

// The platform decision, made once. Everything below reads `impl` and
// `implCapabilities` without asking which OS this is.
const { bridge: impl, capabilities: implCapabilities } =
    process.platform === 'win32' ? windowsBridge
        : process.platform === 'darwin' ? macBridge
            : unsupportedBridge;

/** Which operations work on this machine. See `CapabilityMap`. */
export function capabilities(): CapabilityMap {
    return implCapabilities;
}

/** Whether one operation works here — `capabilities()` for a single name. */
export function supports(operation: BridgeCapability): boolean {
    return implCapabilities[operation] === true;
}

/**
 * An Outlook bridge carrying its own settings, plus capability discovery.
 *
 * The free functions exported below are the same automation against the
 * process-wide config; this is what you want when the process holds more than
 * one caller, or when a call needs its own timeout or AbortSignal.
 */
export interface OutlookBridgeInstance extends OutlookBridge {
    /** The settings this instance runs with. */
    readonly options: Readonly<ResolvedConfig>;

    /** Which operations work on this machine. */
    capabilities(): CapabilityMap;

    /** Whether one operation works here. */
    supports(operation: BridgeCapability): boolean;

    /**
     * A bridge like this one with some settings changed — the per-call escape
     * hatch, and how a cancellable call is made:
     * `bridge.withOptions({ signal }).searchInboxByFilter(...)`.
     */
    withOptions(options: BridgeOptions): OutlookBridgeInstance;
}

type AnyFn = (...args: never[]) => unknown;

/**
 * Build a bridge whose every call runs with `config` in force.
 *
 * The wrapper is what makes per-instance settings work at all: the platform
 * implementations read their config from a module-level accessor, so the config
 * has to be established around the call rather than passed into it. `withConfig`
 * puts it in an AsyncLocalStorage that the whole async subtree inherits, which is
 * why two bridges with different timeouts can run concurrently without seeing
 * each other's.
 */
function build(config: ResolvedConfig): OutlookBridgeInstance {
    const instance = {
        capabilities,
        supports,
        withOptions: (options: BridgeOptions) => build(mergeOptions(config, options)),
        options: Object.freeze({ ...config }),
    } as OutlookBridgeInstance;

    const source = impl as unknown as Record<string, AnyFn>;
    for (const name of ALL_CAPABILITIES) {
        const fn = source[name];
        (instance as unknown as Record<string, unknown>)[name] = (...args: never[]) => {
            // Errors thrown synchronously (an argument check ahead of the first
            // await) become rejections here, so a caller has one failure channel.
            try {
                return withConfig(config, () => fn(...args));
            } catch (error) {
                return Promise.reject(error);
            }
        };
    }
    return instance;
}

/**
 * A bridge with its own settings, isolated from `configure()` and from every
 * other instance.
 *
 * ```ts
 * const bridge = createOutlookBridge({ timeoutMs: 30_000 });
 * if (bridge.supports('searchInboxByFilter')) {
 *     await bridge.withOptions({ signal }).searchInboxByFilter(account, filter);
 * }
 * ```
 *
 * Options left out fall back to the process-wide values as they stand when this
 * is called.
 */
export function createOutlookBridge(options: BridgeOptions = {}): OutlookBridgeInstance {
    return build(mergeOptions(getGlobalConfig(), options));
}

export const {
    getOutlookAccounts,
    sendOutlookEmail,
    replyOutlookEmail,
    sendAllDrafts,
    readInboxEmails,
    searchInboxByFilter,
    readSelectedEmail,
    readEmailBody,
    openOutlookEmail,
    listInboxFolders,
    moveOutlookEmails,
    listOutlookDrafts,
    deleteOutlookDrafts,
    deleteOutlookEmails,
    purgeDeletedItems,
    saveEmailAttachment,
    saveEmailAttachmentDetailed,
    saveEmailAttachments,
    cleanUndeliverableEmails,
    collectBouncedRecipients,
    readSentRecipientGroups,
    listOutlookSignatures,
    readOutlookSignatureHtml,
    readTemplateEmails,
    saveTemplateEmail,
    editEmailTemplate,
} = impl;
