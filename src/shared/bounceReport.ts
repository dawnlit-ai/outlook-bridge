// Measuring what was sent against what bounced. Both reads are platform work
// (collectBouncedRecipients, readSentRecipientGroups); putting them side by side
// is the same everywhere, so it is done once here.
import type { BouncedSend, BounceReport, SentRecipientGroup } from '../types';

function distinctLowerCase(addresses: readonly string[]): string[] {
    return [...new Set(addresses.map(address => address.trim().toLowerCase()).filter(Boolean))];
}

/**
 * Every sent message with a bounced recipient, newest first, and the bounced
 * addresses no sent message accounts for.
 *
 * Kept per message rather than per address or domain: a message's recipient
 * set is the one grouping a mailbox actually records, and whether a failure
 * means anything — one colleague gone, or nobody reached at all — depends on
 * what else that same message went to.
 */
export function buildBounceReport(
    account: string,
    scannedDays: number,
    bouncedAddresses: readonly string[],
    sent: readonly SentRecipientGroup[],
): BounceReport {
    const bounced = distinctLowerCase(bouncedAddresses);
    const bouncedSet = new Set(bounced);
    const matched = new Set<string>();
    const sends: BouncedSend[] = [];

    const newestFirst = [...sent].sort((a, b) => b.sentOn.localeCompare(a.sentOn));
    for (const message of newestFirst) {
        const recipients = distinctLowerCase(message.recipients);
        const failedRecipients = recipients.filter(address => bouncedSet.has(address));
        if (failedRecipients.length === 0) continue;
        failedRecipients.forEach(address => matched.add(address));
        sends.push({
            entryId: message.entryId,
            subject: message.subject,
            sentOn: message.sentOn,
            recipients,
            failedRecipients,
            allFailed: failedRecipients.length === recipients.length,
        });
    }

    return {
        account,
        scannedDays,
        bouncedAddresses: bounced,
        sends,
        unmatchedAddresses: bounced.filter(address => !matched.has(address)).sort(),
    };
}
