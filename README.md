# @dawnlit/outlook-bridge

Drive a real, locally installed Outlook client from Node — read and search mail, send and reply, manage drafts, save
attachments, work with signatures and template emails, file and delete messages. No Graph API, no app registration, no
cloud permissions: this automates the desktop client itself, the way a person would.

- **Windows** — PowerShell and Outlook's COM object model.
- **macOS** — AppleScript. Requires **legacy** Outlook for Mac: New Outlook implements only a slice of the AppleScript
  dictionary (accounts don't enumerate, the inbox reports no messages), so every call fails with `ACCOUNT_NOT_FOUND`
  there rather than quietly doing the wrong thing.
- **Anywhere else** — `capabilities()` reports nothing supported and every call rejects with `UNSUPPORTED_PLATFORM`.

## Install

```bash
npm install @dawnlit/outlook-bridge
```

Requires Node 24+ and a real Outlook on the machine running it — this is desktop automation, not a hosted API client.

## Usage

```ts
import { getOutlookAccounts, readInboxEmails, readEmailBody, replyOutlookEmail } from '@dawnlit/outlook-bridge';

const [account] = await getOutlookAccounts();

const recent = await readInboxEmails(account, { daysBack: 7, limit: 20 });

// A listing row can be passed straight back in as the email to act on.
const email = recent[0];
const { body } = await readEmailBody(email);

await replyOutlookEmail({
    emailAccount: account,
    entryId: email.entryId,
    storeId: email.storeId,
    htmlBody: '<p>Thanks — received.</p>',
    openDraftWindow: false, // file the reply in Drafts without opening a window
});
```

Every operation takes its required arguments positionally and everything optional as one options object, with the
same defaults and the same validation on both platforms.

### Operations

| Area                   | Operations                                                                                                                                                                                                                                                                                          |
|------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Accounts               | `getOutlookAccounts`                                                                                                                                                                                                                                                                                |
| Reading                | `readInboxEmails` (the Inbox root, or any folder by name or path), `searchInboxByFilter` (walks the whole Inbox tree by subject glob or pattern, date window, attachments, folders), `readEmailBody` (the sender's own text, split from the quoted thread), `readSelectedEmail`, `openOutlookEmail` |
| Folders                | `listInboxFolders`, `moveOutlookEmails` (creates a missing destination chain on request)                                                                                                                                                                                                            |
| Composing              | `sendOutlookEmail` (to/cc/bcc, several attachments, send or draft), `replyOutlookEmail` (reply or reply-all, from HTML or a saved template)                                                                                                                                                         |
| Drafts                 | `listOutlookDrafts`, `sendDrafts` (a chosen subset), `sendAllDrafts`, `deleteOutlookDrafts`                                                                                                                                                                                                         |
| Attachments            | `saveEmailAttachment`, `saveEmailAttachments`                                                                                                                                                                                                                                                       |
| Cleanup                | `deleteOutlookEmails` (refuses received and sent mail unless allowed; dry-run first), `purgeDeletedItems`                                                                                                                                                                                           |
| Bounces                | `cleanUndeliverableEmails`, `collectBouncedRecipients`, `readSentRecipientGroups`                                                                                                                                                                                                                   |
| Signatures & templates | `listOutlookSignatures`, `readOutlookSignatureHtml`, `readTemplateEmails`, `saveTemplateEmail`, `editEmailTemplate`                                                                                                                                                                                 |

Each is documented where it is declared, in `OutlookBridge` (see `dist/types.d.ts`), with every option and default.

### Identifying an email

Listings return an `entryId` and a `storeId`. Operations on one email accept an `EmailRef`: the entry id alone, or any
object carrying `{ entryId, storeId? }` — so the row itself can be passed back. Passing the row is also the reliable way
to keep a `storeId` with its own `entryId`; the store id is what lets an email resolve in a mailbox other than the
default one.

**Ids are not portable between platforms.** Windows returns MAPI EntryIDs; macOS returns Outlook for macOS small integer
message ids, and has no store id (it reports `''` and ignores one passed in). Each platform rejects the other's ids with
`INVALID_REQUEST`. **Moving an email rewrites its id** on both, so ids saved before a move are stale afterward.

Timestamps are local time as `yyyy-MM-dd HH:mm`, which sorts correctly as a string.

### Templates

A template email is an ordinary email saved in a mailbox folder (default `Templates`) and edited in Outlook. One template
can hold several variants of a reply between `[[SECTION]] ... [[/SECTION]]` markers around a shared greeting and
closing, with `{{PLACEHOLDER}}` tokens:

```ts
await replyOutlookEmail({
    emailAccount: account,
    entryId: email.entryId,
    storeId: email.storeId,
    templateSubject: 'Standard reply',
    templateSection: 'ACCEPT',
    templatePlaceholders: { NOTE: 'We will confirm by Friday.' },
    signatureName: 'Work', // fills {{SIGNATURE}}
});
```

The template, the section, the placeholders and the signature are all resolved inside the package, so a large
Word-generated template never has to pass through the caller. Composition refuses rather than guesses: a missing section,
an unknown placeholder, or a marker left unresolved fails with `INVALID_REQUEST` before anything is created.

The same toolkit works on any HTML string: `composeTemplateBody`, `findTemplateMarkers`, and for `{Token}`-style fills,
`replaceToken`, `removeTokenLine` (removes a token together with its table row or paragraph), `findTokens` and
`findUnfilledTokens`.

## Configuration

Every call generates a script and runs it under an interpreter. The settings are how long a run may take, how much it may
print, how to call it off, and what the script said:

```ts
import { createOutlookBridge } from '@dawnlit/outlook-bridge';

const outlook = createOutlookBridge({
    timeoutMs: 120_000,               // ceiling on every run (default: a budget sized for each operation)
    maxBufferBytes: 32 * 1024 * 1024, // output cap; searches returning bodies push against it (default 8 MiB)
    tempDir: '/var/tmp/outlook',      // generated scripts, and attachments saved without a destination
    debug: event => log(event.script),// every generated script, its duration, and any error
});
```

- **Timeouts.** Unless `timeoutMs` is set, each run gets a budget sized for its work — a minute for a lookup, two for
  composing or reading a folder, five for a mailbox walk or a batch, ten for purging. `timeoutMs` replaces all of them;
  `0` disables timeouts; `null` goes back to the budgets. `editEmailTemplate` waits on a person and is never timed —
  cancel it with a signal.
- **Instances.** A bridge's settings apply to its own calls only, even when several bridges run at once.
  `outlook.withOptions({ ... })` derives a bridge with some settings changed — which is how a single call gets its own
  budget or cancellation:

  ```ts
  const controller = new AbortController();
  cancelButton.onclick = () => controller.abort();
  await outlook.withOptions({ signal: controller.signal }).searchInboxByFilter(account, { daysBack: 365 });
  // → rejects with ABORTED, interpreter killed
  ```

- **Process-wide.** The exported functions run under process-wide settings that `configure()` changes. That is the
  simpler choice for a program that owns its process; prefer `createOutlookBridge` in anything sharing one.
- **Debugging.** `OUTLOOK_BRIDGE_DEBUG=1` logs every script to stderr without a code change.

Capability discovery answers for the machine you are on: `capabilities()` returns a map of every operation to
`true`/`false`, and `supports('searchInboxByFilter')` asks about one.

## Errors

Every deliberate failure is an `OutlookError` with a stable `code`. **The codes are the API; the messages are not** —
branch on `code`, never on message text.

```ts
import { OutlookError } from '@dawnlit/outlook-bridge';

try {
    await outlook.readInboxEmails(account, { folder: 'Invoices\\2026' });
} catch (error) {
    if (!(error instanceof OutlookError)) throw error; // a bug
    switch (error.code) {
        case 'ACCOUNT_NOT_FOUND': return askForAccount(error.account);
        case 'NOT_FOUND':         return report(`${error.kind} missing: ${error.message}`);
        case 'OUTPUT_TOO_LARGE':  return retryWithNarrowerWindow();
        case 'SCRIPT_FAILED':     return logFailure(error.script, error.line, error.stderr);
    }
}
```

| code                   | class                      | meaning                                                                                                                                                        |
|------------------------|----------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `UNSUPPORTED_PLATFORM` | `UnsupportedPlatformError` | No Outlook automation exists on this OS.                                                                                                                       |
| `NOT_IMPLEMENTED`      | `NotImplementedError`      | The platform could do this, but this corner isn't written (a folder root macOS has no term for).                                                               |
| `ACCOUNT_NOT_FOUND`    | `AccountNotFoundError`     | No mailbox in the profile matches the address. Carries `account`.                                                                                              |
| `NOT_FOUND`            | `NotFoundError`            | A `folder`, `email`, `attachment`, `template`, `signature` or `file` didn't resolve. Carries `kind`.                                                           |
| `INVALID_REQUEST`      | `InvalidRequestError`      | The arguments can't produce a call — a wrong type, a missing value, a template edited into an unusable state, or a mailbox that can be read but not sent from. |
| `SCRIPT_FAILED`        | `ScriptError`              | The script ran and failed. Carries `script`, `line`, `stderr`, `runner`, `durationMs`.                                                                         |
| `OUTPUT_TOO_LARGE`     | `OutputTooLargeError`      | The script printed more than `maxBufferBytes`.                                                                                                                 |
| `TIMEOUT`              | `TimeoutError`             | The run exceeded its time budget and was killed. Carries `timeoutMs`.                                                                                          |
| `ABORTED`              | `AbortedError`             | The caller's `AbortSignal` fired.                                                                                                                              |

Batch operations (moving, deleting, sending drafts, cleaning bounces) don't throw for one bad item: they report each
failure in `failed` beside the items that worked.

## Security

Caller text ends up inside generated scripts, so the package treats every argument as untrusted:

- Arguments are checked for type at run time — a count must be a finite number, a flag a boolean — before any script
  is generated, whatever TypeScript believed at compile time.
- On Windows, caller text only ever appears in single-quoted PowerShell literals, with every quote character PowerShell
  accepts (including the typographic `‘ ’ ‚ ‛`) escaped. Text that has to appear in a message is referenced through a
  variable, never interpolated into a double-quoted string where `$(...)` would run. Scripts run from UTF-8 files with a
  byte-order mark, so non-ASCII text arrives intact.
- Attachment names come from the sender, so a saved file's name is sanitized — no path separators, reserved device
  names or characters the file system rejects — and never overwrites an existing file.

## Module format

Published as CommonJS; ESM consumers can use named imports normally. There is deliberately no dual build: the package
identifies its errors with `instanceof`, and a dual-format package can load two copies of itself into one process.

## Development

```bash
npm run build   # tsc → dist/
npm test        # build, then node:test
```

The tests need no Outlook. They cover every operation's validation and defaults against a fake backend, the pure
helpers, the error taxonomy, and the generated scripts themselves:

- **On Windows**, every PowerShell script every operation generates is put through PowerShell's own parser, with
  arguments built to break out of a badly escaped literal. The test fails on a syntax error, and on any caller text
  inside a double-quoted string.
- **On a Mac with Outlook installed**, every AppleScript is compiled with `osacompile`, which resolves dictionary terms
  without opening a session. Elsewhere, those tests skip; two guards that need no Outlook — no assignment to a reserved
  AppleScript word, no record field read through a nested accessor — run everywhere.

What no test covers is the automation itself: that a move really filed the mail, or a reply threaded correctly, is
verified against a live mailbox.

## License

Apache-2.0 © Dawnlit
