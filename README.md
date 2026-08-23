# @dawnlit/outlook-bridge

Drive a real, locally installed Outlook client from Node — read the inbox, send and reply to mail, manage drafts, work
with signatures and template emails, file and delete messages. No Graph API, no app registration, no cloud permissions:
this automates the desktop client itself, the same way a person would.

- **Windows** — via PowerShell + Outlook's COM object model.
- **macOS** — via AppleScript. New Outlook for Mac can compose and send, but cannot read the mailbox the way classic
  Outlook can; functions that need mailbox access will throw a clear error on that combination. A number of functions
  are not yet ported and throw `not implemented for Outlook on macOS yet`.
- Any other platform — every call rejects with a clear "not supported" error.

## Install

```bash
npm install @dawnlit/outlook-bridge
```

Requires Node 20+ and a real Outlook installation on the machine running it — this is desktop automation, not a hosted
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

## Configuration

Each call generates a script and shells out to an interpreter. `configure()` covers the three things that leaves
otherwise out of your reach — how long a run may take, how much it may print, and what the script actually said.

```ts
import { configure } from '@dawnlit/outlook-bridge';

configure({
    timeoutMs: 120_000,               // default per-run ceiling; 0 disables
    maxBufferBytes: 32 * 1024 * 1024, // raise before a scan that returns thousands of bodies
    tempDir: '/var/tmp/outlook',      // scratch scripts and default attachment destination
    debug: (e) => log(e.script),      // every generated script, its duration, and any error
});
```

`OUTLOOK_BRIDGE_DEBUG=1` in the environment turns on script logging to stderr without a code change.

Calls that carry a genuinely different budget (a full-mailbox walk, a purge) keep their own timeout and ignore
`timeoutMs`. `searchInboxByFilter` is the one most likely to push against `maxBufferBytes`, since it returns whole
message bodies.

## API

**Accounts & inbox** — `getOutlookAccounts`, `readInboxEmails`, `readEmailBody`,
`listInboxFolders`, `moveOutlookEmails`, `openOutlookEmail`.

**Search & selection** — `searchInboxByFilter` (walks every folder under the Inbox for one
account, filtering by subject pattern/date window/reply-exclusion/attachment presence — the
building block for "find the emails matching X" without knowing which subfolder holds them),
`readSelectedEmail` (the email currently selected or open in Outlook).

**Sending & replying** — `sendOutlookEmail`, `replyOutlookEmail` (answers the sender; pass `replyAll: true` to answer
every recipient), `sendAllDrafts`.

**Drafts** — `listOutlookDrafts`, `deleteOutlookDrafts`.

**Attachments** — `saveEmailAttachment`, `saveEmailAttachmentDetailed`, `saveEmailAttachments`
(the same, batched into one COM round trip for several attachments off one email). All take an optional `destDir`,
created if absent. Without one, each call saves into a fresh private directory — attachments keep the name the sender
gave them, so a shared folder means two emails carrying `invoice.pdf` overwrite each other.

**Signatures & templates** — `listOutlookSignatures`, `readOutlookSignatureHtml`,
`readTemplateEmails`, `saveTemplateEmail`, `editEmailTemplate` (Windows only — opens the template in a real
Outlook compose window and returns the saved HTML; throws on other platforms).

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

Every exported type (`InboxEmail`, `ReplyEmailParams`, `MailFolderRef`, `OutlookBridge`, etc.) is exported alongside its
function.

## Development

```bash
npm run build   # tsc → dist/
npm test        # build, then node:test over the platform-independent parts
```

The tests cover the pure logic — folder-string parsing, quote splitting, template composition, scratch-file handling —
so they run anywhere, with or without Outlook installed.

## License

Apache-2.0 © Dawnlit
