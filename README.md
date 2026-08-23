# @dawnlit/outlook-bridge

Drive a real, locally installed Outlook client from Node — read the inbox, send and
reply to mail, manage drafts, work with signatures and template emails, file and
delete messages. No Graph API, no app registration, no cloud permissions: this
automates the desktop client itself, the same way a person would.

- **Windows** — via PowerShell + Outlook's COM object model.
- **macOS** — via AppleScript. New Outlook for Mac can compose and send, but
  cannot read the mailbox the way classic Outlook can; functions that need
  mailbox access will throw a clear error on that combination.
- Any other platform — every call rejects with a clear "not supported" error.

## Install

```bash
npm install @dawnlit/outlook-bridge
```

Requires Node 20+ and a real Outlook installation on the machine running it —
this is desktop automation, not a hosted API client.

## Usage

```ts
import { getOutlookAccounts, readInboxEmails, sendOutlookEmail } from '@dawnlit/outlook-bridge';

const accounts = await getOutlookAccounts();

const { emails } = await readInboxEmails(accounts[0]);

await sendOutlookEmail({
    account: accounts[0],
    to: ['someone@example.com'],
    subject: 'Hello from @dawnlit/outlook-bridge',
    bodyHtml: '<p>Sent by driving the Outlook client directly.</p>',
});
```

## API

**Accounts & inbox** — `getOutlookAccounts`, `readInboxEmails`, `readEmailBody`,
`listInboxFolders`, `moveOutlookEmails`, `openOutlookEmail`.

**Sending & replying** — `sendOutlookEmail`, `replyOutlookEmail`, `sendAllDrafts`,
`sendReceivedConfirmation`.

**Drafts** — `listOutlookDrafts`, `deleteOutlookDrafts`.

**Attachments** — `saveEmailAttachment`, `saveEmailAttachmentDetailed`.

**Signatures & templates** — `listOutlookSignatures`, `readOutlookSignatureHtml`,
`readTemplateEmails`, `saveTemplateEmail`, `editEmailTemplate` (desktop-only on
macOS — see below).

**Bounce handling** — `cleanUndeliverableEmails`, `collectBouncedRecipients`,
`readSentRecipientGroups`.

**Mailbox cleanup** — `deleteOutlookEmails`, `purgeDeletedItems`.

**Template-section parsing** — a small pure-string toolkit for templates that
carry several reply variants between `[[SECTION]] ... [[/SECTION]]` markers
around one shared greeting/signature, with `{{PLACEHOLDER}}` tokens filled in
by the caller: `findTemplateMarkers`, `composeTemplateBody`, `findTokens`,
`findUnfilledTokens`, `replaceToken`, `removeTokenLine`.

Every exported type (`InboxEmail`, `ReplyEmailParams`, `MailFolderRef`, etc.)
is exported alongside its function.

## Electron

`editEmailTemplate` on macOS opens an in-app WYSIWYG editor window, and does
so by lazily `require('electron')` — so it only works when this package runs
inside an Electron app, and only on that one function. `electron` is declared
as an **optional** peer dependency: every other export is plain Node
(`child_process`, `fs`, `os`) and works in any Node process, Electron or not.

## License

Apache-2.0 © Dawnlit
