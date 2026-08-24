# @dawnlit/outlook-bridge

Drive a real, locally installed Outlook client from Node — read the inbox, send and reply to mail, manage drafts, work
with signatures and template emails, file and delete messages. No Graph API, no app registration, no cloud permissions:
this automates the desktop client itself, the same way a person would.

- **Windows** — via PowerShell + Outlook's COM object model.
- **macOS** — via AppleScript. The whole surface is ported, with two differences no capability flag can express (see
  [Platform differences](#platform-differences)). Requires **legacy** Outlook for Mac: New Outlook implements only a
  slice of the AppleScript dictionary — accounts don't enumerate and the inbox reports 0 messages — so calls fail
  loudly with `ACCOUNT_NOT_FOUND` there rather than quietly doing the wrong thing.
- Any other platform — `capabilities()` reports everything false and every call rejects with `UNSUPPORTED_PLATFORM`.

## Install

Install straight off GitHub `main`, rather than a versioned release:

```bash
npm install github:dawnlit-ai/outlook-bridge#main
```

`prepare` runs `npm run build` automatically on install, so `dist/` is built from that checked-out
source with no separate step. npm resolves the branch to a commit and pins it in `package-lock.json`,
so a later plain `npm install` won't pick up new commits on its own — re-run the command above
(naming the package and `#main` explicitly) whenever you want to advance to the current tip.

Requires Node 24+ and a real Outlook installation on the machine running it — this is desktop automation, not a hosted
API client.

## Usage

```ts
import { getOutlookAccounts, readInboxEmails, sendOutlookEmail } from '@dawnlit/outlook-bridge';

const accounts = await getOutlookAccounts();

const emails = await readInboxEmails(accounts[0]);

await sendOutlookEmail({
    emailAccount: accounts[0],
    to: 'someone@example.com',          // comma- or semicolon-separated
    subject: 'Hello from @dawnlit/outlook-bridge',
    htmlBody: '<p>Sent by driving the Outlook client directly.</p>',
    sendImmediately: false,             // false leaves it in Drafts
});
```

Every function is exposed under one signature regardless of platform — the dispatcher pins both implementations to the
same `OutlookBridge` interface, so nothing here returns a per-platform union you have to narrow.

## Platform differences

Both platforms implement the whole contract, so `capabilities()` is all-true on Windows and on macOS. It still exists,
and is still worth asking, because it answers for the machine you are actually on — on any other OS every entry is
false and every call rejects:

```ts
import { createOutlookBridge } from '@dawnlit/outlook-bridge';

const bridge = createOutlookBridge();

if (bridge.supports('searchInboxByFilter')) { /* … */ }

bridge.capabilities(); // { readInboxEmails: true, replyOutlookEmail: true, … }
```

The map is derived from `keyof OutlookBridge`, so a function added to the contract fails both platform maps at
compile time rather than quietly defaulting to supported.

Two things differ between the platforms that no flag can express, and code moving between them has to know both:

**Ids are not portable.** Windows returns MAPI `EntryID` strings; macOS returns Outlook for Mac's small integer message
ids (`"1263"`). Each platform rejects the other's with `INVALID_REQUEST` rather than failing later as a confusing "not
found". `storeId` has no macOS equivalent: it is returned as `''` and accepted-and-ignored wherever it is a parameter.
Don't persist an id from one platform and look it up on the other.

**`searchInboxByFilter` filters in-process on macOS.** Windows hands `subjectLike` and the date window to Outlook's
server-side `Items.Restrict`; AppleScript has no counterpart, so the macOS walk indexes each folder with bulk property
reads and applies the same tests here. Same results, different cost — a `daysBack` of 0 over a large mailbox tree is
considerably heavier on macOS.

Two smaller notes: `replyOutlookEmail` reports `to` as the reply's resolved recipient addresses on macOS rather than the
display-name string Windows reports; and `editEmailTemplate` on macOS files the edited draft when you answer Outlook's
own "save this message?" prompt on close, where Windows asks for Ctrl+S.

## Configuration

Each call generates a script and shells out to an interpreter. That leaves four things otherwise out of your reach:
how long a run may take, how much it may print, how to call it off, and what the script actually said.

`createOutlookBridge()` gives you those on an instance no other caller in the process can disturb:

```ts
import { createOutlookBridge } from '@dawnlit/outlook-bridge';

const bridge = createOutlookBridge({
    timeoutMs: 120_000,               // default per-run ceiling; 0 disables
    maxBufferBytes: 32 * 1024 * 1024, // raise before a scan that returns thousands of bodies
    tempDir: '/var/tmp/outlook',      // scratch scripts and default attachment destination
    debug: (e) => log(e.script),      // every generated script, its duration, and any error
});

await bridge.readInboxEmails(account);
```

`bridge.withOptions({ … })` derives a bridge with some settings changed — how one call gets its own budget, or its
own cancellation:

```ts
const controller = new AbortController();
cancelButton.onclick = () => controller.abort();

await bridge.withOptions({ signal: controller.signal }).searchInboxByFilter(account, filter);
// → rejects with an AbortedError (code 'ABORTED'), interpreter killed
```

The same functions are also exported directly, running against a process-wide config that `configure()` sets. That is
the simpler thing for a program that owns its process; prefer `createOutlookBridge` in anything sharing a process
with code you do not own.

```ts
import { configure, readInboxEmails } from '@dawnlit/outlook-bridge';

configure({ timeoutMs: 60_000 });
await readInboxEmails(account);
```

`OUTLOOK_BRIDGE_DEBUG=1` in the environment turns on script logging to stderr without a code change.

Calls that carry a genuinely different budget (a full-mailbox walk, a purge) keep their own timeout and ignore
`timeoutMs`. `searchInboxByFilter` is the one most likely to push against `maxBufferBytes`, since it returns whole
message bodies.

## Errors

Every deliberate failure is an `OutlookError` carrying a stable `code`. **The codes are the API; the messages are
not** — branch on `code`, never on message text.

```ts
import { OutlookError } from '@dawnlit/outlook-bridge';

try {
    await bridge.replyOutlookEmail(params);
} catch (error) {
    if (!(error instanceof OutlookError)) throw error; // a real bug
    switch (error.code) {
        case 'ACCOUNT_NOT_FOUND': return promptForAccount(error.account);
        case 'NOT_IMPLEMENTED':   return degrade(error.operation);
        case 'ABORTED':           return;
        case 'SCRIPT_FAILED':     return report(error.script, error.stderr);
    }
}
```

| code | class | meaning |
| --- | --- | --- |
| `UNSUPPORTED_PLATFORM` | `UnsupportedPlatformError` | No Outlook automation exists on this OS. |
| `NOT_IMPLEMENTED` | `NotImplementedError` | The platform could, but this corner isn't written. Carries `operation`. Now only reachable for a folder root macOS has no term for (Journal and the like). |
| `ACCOUNT_NOT_FOUND` | `AccountNotFoundError` | No configured account matches. Carries `account`. |
| `NOT_FOUND` | `NotFoundError` | A folder, template, signature or item did not resolve. Carries `kind`. |
| `INVALID_REQUEST` | `InvalidRequestError` | The arguments cannot produce a call; nothing was attempted. |
| `SCRIPT_FAILED` | `ScriptError` | The interpreter ran and failed. Carries `script`, `stderr`, `runner`, `durationMs`. |
| `TIMEOUT` | `TimeoutError` | The run exceeded its budget and was killed. Carries `timeoutMs`. |
| `ABORTED` | `AbortedError` | The caller's `AbortSignal` fired. |

`ScriptError` carries the generated script verbatim — the thing you actually want when one fires, and otherwise
reachable only if you wired up a `debug` hook in advance.

## Module format

Published as CommonJS. ESM consumers can use named imports normally (`import { sendOutlookEmail } from …`); Node
resolves them off the CJS build.

There is deliberately no dual ESM/CJS build. This package identifies its errors with `instanceof`, and a dual-format
package can load two copies of itself into one process — at which point `error instanceof OutlookError` returns
`false` for errors the package itself threw. One format avoids that failure entirely.

## API

**Accounts & inbox** — `getOutlookAccounts`, `readInboxEmails`, `readEmailBody`,
`listInboxFolders`, `moveOutlookEmails`, `openOutlookEmail`.

**Search & selection** — `searchInboxByFilter` (walks every folder under the Inbox for one
account, filtering by subject pattern/date window/reply-exclusion/attachment presence — the
building block for "find the emails matching X" without knowing which subfolder holds them),
`readSelectedEmail` (the email currently selected or open in Outlook).

**Sending & replying** — `sendOutlookEmail`, `replyOutlookEmail` (answers the sender; pass `replyAll: true` to answer
every recipient), `sendAllDrafts`.

**Drafts** — `listOutlookDrafts`, `deleteOutlookDrafts`, `sendDrafts` (sends a chosen subset by
EntryID — the review-and-select alternative to `sendAllDrafts`'s account-wide sweep).

**Attachments** — `saveEmailAttachment`, `saveEmailAttachmentDetailed`, `saveEmailAttachments`
(the same, batched into one COM round trip for several attachments off one email). All take an optional `destDir`,
created if absent. Without one, each call saves into a fresh private directory — attachments keep the name the sender
gave them, so a shared folder means two emails carrying `invoice.pdf` overwrite each other.

**Signatures & templates** — `listOutlookSignatures`, `readOutlookSignatureHtml`,
`readTemplateEmails`, `saveTemplateEmail`, `editEmailTemplate` (opens the template in a real Outlook compose window,
waits for the user to close it, and returns the saved HTML — so it has no timeout by default; give it an
`AbortSignal` if you need a way to give up).

**Bounce handling** — `cleanUndeliverableEmails`, `collectBouncedRecipients`,
`readSentRecipientGroups`.

**Mailbox cleanup** — `deleteOutlookEmails`, `purgeDeletedItems`.

**Template-section parsing** — a small pure-string toolkit for templates that carry several reply variants between
`[[SECTION]] ... [[/SECTION]]` markers around one shared greeting/signature, with `{{PLACEHOLDER}}` tokens filled in by
the caller: `findTemplateMarkers`, `composeTemplateBody`, `findTokens`,
`findUnfilledTokens`, `replaceToken`, `removeTokenLine`.

**Platform-neutral helpers** — usable with no Outlook session at all: `mailFolderRef` (parse any folder string a user
might type into a well-known root plus segments), `splitQuotedOriginal` (separate a reply's own text from the thread
quoted below it), `WELL_KNOWN_FOLDERS`.

**Instances, capabilities and errors** — `createOutlookBridge` (a bridge with its own settings; see
[Configuration](#configuration)), `capabilities` / `supports`, and the `OutlookError` family
(see [Errors](#errors)).

Every exported type (`InboxEmail`, `ReplyEmailParams`, `MailFolderRef`, `OutlookBridge`, `CapabilityMap`,
`OutlookErrorCode`, etc.) is exported alongside its function.

## Development

```bash
npm run build   # tsc → dist/
npm test        # build, then node:test over the platform-independent parts
```

The tests cover the parts that need no Outlook session — folder-string parsing, quote splitting, template
composition, scratch-file handling, the error taxonomy and its run classifier, capability maps, config scoping, and
the shape of the PowerShell the folder walk emits — so they run anywhere, with or without Outlook installed.

On a **Mac with Outlook installed** they additionally compile every AppleScript the macOS implementation generates.
The runner is stubbed, each operation is driven with plausible arguments, and the captured scripts go through
`osacompile` — which resolves terminology against the real dictionary without opening a session, sending anything or
touching a message. This is worth the trouble because one bad dictionary term fails the WHOLE script at compile time
rather than at the offending line, so a live failure never names the thing that caused it. Those tests skip
automatically elsewhere.

Two guards over the same captured scripts run **everywhere**, with no Outlook needed, because they cover the
failures compiling cannot catch — both of which shipped once and only surfaced against a live mailbox: assigning to a
reserved AppleScript word (`set rest to …` parses, then fails at run time), and reading a record field through a
nested accessor (`address of (sender of m)` does not coerce, and inside a `try` that means a silently empty sender).

What no test can cover is the automation itself: driving a real Outlook client is the whole point of the package, and
behaviour against a live mailbox — that a `move` really filed the mail, that a reply threaded correctly — is verified
by running it, not in CI.

## License

Apache-2.0 © Dawnlit
