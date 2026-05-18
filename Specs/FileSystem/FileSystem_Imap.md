# IMAP Virtual File System Plugin

## Overview

RedSalamander ships a built-in IMAP virtual file system implemented by `Plugins/FileSystemCurl/FileSystemCurl.dll`.

- Display name: **IMAP**
- `PluginMetaData.id`: `builtin/file-system-imap`
- `PluginMetaData.shortId`: `imap`

The IMAP implementation maps IMAP mailboxes and messages into a directory/file model to enable browsing, reading, and deleting email.

## Navigation (URI Syntax)

Supported navigation forms:

- Authority form:
  - `imap://[user[:password]@]host[:port][/mailbox][/message.eml]`
- Defaults form (requires `defaultHost`):
  - `imap:/path`

Notes:

- Mailboxes are exposed as directories, using `/` as the UI path separator.
  - The plugin reconstructs mailbox hierarchy using the **hierarchy delimiter** returned by IMAP `LIST`.
  - Example: if the server delimiter is `.` and the server reports `A`, `A.B`, `A.C`, `A.C.E`, the UI shows `A/B`, `A/C`, `A/C/E`.
- Messages are exposed as `.eml` files inside selectable mailboxes.
- User/password in the URI override defaults (passwords in URIs are discouraged; prefer Connection Manager).

Examples:

- `imap://user@example.com/INBOX/`
- Stable direct access by UID: `imap://user@example.com/INBOX/12345.eml`
- With defaults (`defaultHost = imap.example.com`): `imap:/INBOX/`

## Connection Manager Integration

The plugin supports host-reserved navigation:

- `/@conn:<connectionName>/...`

The host resolves `<connectionName>` to a saved `ConnectionProfile` and the plugin requests credentials via `IHostConnections`.

Secret acquisition (high level):

- If the profile requires a secret, the plugin calls:
  - `IHostConnections::GetConnectionSecret` first (no prompt).
  - If it returns `ERROR_NOT_FOUND`, `IHostConnections::PromptForConnectionSecret` prompts the user and caches the secret for the current app run (session-only; not persisted).
  - If the prompt returns `S_FALSE`, treat it as user cancellation (`ERROR_CANCELLED`).

## Virtual File System Mapping

### Directories (mailboxes)

- Root (`imap:/`) lists mailboxes (or the subtree under `defaultBasePath`).
- Entering a mailbox directory lists:
  - sub-mailboxes as directories, and
  - messages as `.eml` files (when the mailbox is selectable).

### Files (messages)

- Directory listing names are formatted as `<subject> [<uid>].eml`.
  - Example: `Quarterly report [12345].eml`.
  - `<subject>` is best-effort decoded from RFC2047 encoded-words (Q/B) into UTF-16; missing/empty becomes `(no subject)`.
  - RFC2047 decoding supports UTF-8, US-ASCII, ISO-8859 families, Windows code pages, and common Japanese, Chinese, Korean, Cyrillic, Thai, and Macintosh MIME charset aliases through Windows code-page conversion. Unknown charsets fall back to UTF-8, the active Windows code page, then Windows-1252.
  - Malformed encoded-word fragments already using the sanitized `=_charset_q_...` shape are decoded as a recovery path before filename sanitization. U+00AD soft hyphen in decoded subjects is normalized to a visible `-` for pane and property display.
  - Subject text is sanitized for Windows filenames and capped at 96 UTF-16 code units. If it exceeds the cap, the first 93 code units are kept and ASCII `...` is appended before the ` [<uid>].eml` suffix.
  - The UID in the final bracketed suffix is the stable operation key. Subjects are display text only and are never used for message identity.
- Directory enumeration retrieves message metadata using:
  - `UID FETCH <uid-set> (UID FLAGS INTERNALDATE RFC822.SIZE ENVELOPE)`
- When a bulk summary fetch fails or returns incomplete metadata for requested UIDs, directory enumeration MUST retry missing summaries in bounded smaller UID batches and then single UID fetches before falling back to numeric `<uid>.eml`, zero size, or missing timestamps in the pane.
- The plugin also accepts `<uid>.eml` for direct navigation/bookmarks even if the directory listing shows the decorated name.
- Ambiguous names such as `Quarterly report 12345.eml` are not valid IMAP message keys; use the bracketed suffix or direct numeric UID form.
- Opening a message file downloads the RFC822 message content using `UID FETCH <uid> BODY.PEEK[]` and exposes it as a read-only file.

#### FileInfo field mapping (messages)

The IMAP plugin populates `FileInfo` fields for message entries as follows:

- `FileIndex`: IMAP UID (message id).
- `CreationTime`: message sent time (parsed from the RFC5322 `Date:` header when available).
- `ChangeTime`: message received time (from IMAP `INTERNALDATE`).
- `LastWriteTime`: same as `ChangeTime` (received time).
- `EndOfFile` and `AllocationSize`: message size in bytes (from IMAP `RFC822.SIZE`; includes attachments as part of the RFC822 message).
- `FileAttributes`:
  - `FILE_ATTRIBUTE_DIRECTORY` for mailbox directories.
  - `0x02000000` when the message has `\\Flagged` (marked).
  - `0x04000000` when the message is **unread** (does not have `\\Seen`).
  - `0x08000000` when the message has `\\Deleted`.

#### Properties metadata

- `GetItemProperties` exposes general item data, remote/display paths, connection metadata, and IMAP metadata when available.
- Mailbox directories commonly do not have meaningful filesystem timestamps; unavailable timestamp fields MUST be omitted rather than shown as `0`.
- Message properties may include sent time, received/internal time, flags, UID, sender, subject, and size. Standard timestamp fields are included only when non-zero.
- Message properties MUST fetch metadata for the selected UID only; they must not enumerate and summarize the full mailbox just to show one message's Properties.
- Mailbox folder properties include the plugin mailbox path, server mailbox path, hierarchy delimiter, root/mailbox identity, connection metadata, and `STATUS` counts (`MESSAGES`, `RECENT`, `UIDNEXT`, `UIDVALIDITY`, `UNSEEN`) when the server provides them.
- Passwords and raw secrets are never emitted in properties; boolean presence fields such as `hasPassword` are sufficient.

## Pane Display Contract

For a message with subject `Quarterly report`, UID `12345`, received time `2026-05-13 21:33`, and RFC822 size `45 KB`:

| Pane mode | Display |
|-----------|---------|
| Brief | `Quarterly report [12345].eml` |
| Detailed | Name line plus `2026-05-13 21:33 • 45 KB • -` on the details line |
| Extra Detailed | Same name and details line unless a later metadata provider adds a third line |
| Thumbnails | File icon/thumbnail slot with the same name and details text |
| Status bar | `45 KB • 2026-05-13 21:33 • -` for a focused or singly selected message |

The full subject and sender belong in Properties. The filename may be truncated for display and layout, but the bracketed UID remains visible at the end of the logical leaf name.

## Performance Metrics

The IMAP implementation emits these metric names when diagnostics or perf JSONL capture is enabled:

- `filesystem.imap.read_directory_us`
- `filesystem.imap.list_mailboxes_us`
- `filesystem.imap.list_uids_us`
- `filesystem.imap.fetch_summaries_us`
- `filesystem.imap.summary_response_bytes`
- `filesystem.imap.summary_parse_us`
- `filesystem.imap.summary_missing_count`
- `filesystem.imap.summary_repair_count`
- `filesystem.imap.build_fileinfo_us`
- `filesystem.imap.properties_message_us`
- `filesystem.imap.properties_mailbox_us`
- `filesystem.imap.status_mailbox_us`

Message Properties performance is protected by a command-count contract: opening Properties for one message must stay constant with respect to mailbox size. The selected UID path uses targeted UID metadata fetches rather than a full mailbox summary refetch.

Mailbox listing summary repair is bounded per listing. The repair path may batch missing UID summaries first and then fall back to singleton fetches, but it MUST cap repair fetch attempts with a per-listing budget, log one warning when the budget is exhausted, and proceed with any summaries already recovered instead of issuing unbounded singleton fetches against flaky servers.

## Operations

### Browse

- `ReadDirectoryInfo()`:
  - Root lists mailboxes.
  - Mailbox directories list sub-mailboxes and message `.eml` files.

### Read

- `CreateFileReader()` downloads the message into a temporary delete-on-close file and returns an `IFileReader`.

### Delete

- Deleting a message file:
  - `UID STORE <uid> +FLAGS.SILENT (\Deleted)`
  - `UID EXPUNGE <uid>` (falls back to `EXPUNGE` if unsupported)
- Deleting a mailbox directory:
  - `DELETE "<mailbox>"`
- Recursive directory deletion deletes contained messages and sub-mailboxes before deleting the mailbox.

## Settings (Plugin Configuration)

The plugin exposes its configuration schema via `IInformations::GetConfigurationSchema()`:

- `defaultHost` (string)
- `defaultPort` (integer, `0` = protocol default, typically 143)
- `ignoreSslTrust` (bool, default `false`): when `true`, TLS certificate validation is skipped (allows self-signed certificates; not recommended).
- `defaultUser` (string)
- `defaultPassword` (string, stored as plain text in settings)
- `defaultBasePath` (string, mailbox prefix; `/` lists all mailboxes)
- `connectTimeoutMs` (integer)
- `operationTimeoutMs` (integer)

## TLS Behavior

- Port `993` uses implicit TLS by using the `imaps` scheme internally.
- Other ports use IMAP with STARTTLS when available (best-effort).
- By default, certificate validation uses the Windows trust store (libcurl native CA store).
- To connect to servers with self-signed/untrusted certificates, set `ignoreSslTrust = true` (either via plugin configuration or per-connection profile `extra.ignoreSslTrust` when using Connection Manager).
