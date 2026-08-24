// Turning a platform's bridge object into its capability map.
//
// The maps used to be hand-typed literals — 26 names and 26 `true`s per platform,
// restating the keys of the bridge object directly above them. They were the
// fourth place a new operation's name had to be typed, and the one place a
// stale `false` could sit unnoticed beside a function that works.
//
// Availability is derived from the bridge instead, with gaps named explicitly:
// a platform declares what it has NOT ported, which is the short list and the
// one worth reading. `CapabilityMap`'s `false` still means what `types.ts` says
// it means — "asking is pointless here" — including for an operation that
// answers emptily rather than throwing.
import type { BridgeCapability, CapabilityMap, OutlookBridge } from '../types';

/** A platform's capability map: every operation it implements, minus `notPorted`. */
export function capabilityMap(
    bridge: OutlookBridge,
    notPorted: readonly BridgeCapability[] = [],
): CapabilityMap {
    const names = Object.keys(bridge) as BridgeCapability[];
    return Object.freeze(
        Object.fromEntries(names.map(name => [name, !notPorted.includes(name)])),
    ) as CapabilityMap;
}
