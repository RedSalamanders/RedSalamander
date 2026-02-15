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

- Directory listing names are formatted as `<subject>｜<from>｜<uid>.eml`.
  - Note: `｜` is **U+FF5C FULLWIDTH VERTICAL LINE** (Windows does not allow ASCII `|` in filenames).
  - `<subject>` is best-effort decoded from RFC2047 encoded-words (Q/B) into UTF-16; missing/empty becomes `(no subject)`.
  - `<from>` is the sender email address (addr-spec) extracted from the RFC5322 `From:` header; missing/empty becomes `(unknown sender)`.
  - The UID is always the trailing digits before `.eml`; the plugin parses the UID from the leaf name.
- Directory enumeration retrieves message metadata using:
  - `UID FETCH <uid-set> (UID FLAGS INTERNALDATE RFC822.SIZE ENVELOPE)`
- The plugin also accepts `<uid>.eml` for direct navigation/bookmarks even if the directory listing shows the decorated name.
- Opening a message file downloads the RFC822 message content using `UID FETCH <uid> BODY.PEEK[]` and exposes it as a read-only file.

#### FileInfo field mapping (messages)

The IMAP plugin populates `FileInfo` fields for message entries as follows:

- `FileIndex`: IMAP UID (message id).
- `CreationTime`: message sent time (parsed from the RFC5322 `Date:` header when available).
- `ChangeTime`: message received time (from IMAP `INTERNALDATE`).
- `LastWriteTime`: same as `ChangeTime` (received time).
- `EndOfFile`: message size in bytes (from IMAP `RFC822.SIZE`; includes attachments as part of the RFC822 message).
- `FileAttributes`:
  - `FILE_ATTRIBUTE_DIRECTORY` for mailbox directories.
  - `0x02000000` when the message has `\\Flagged` (marked).
  - `0x04000000` when the message is **unread** (does not have `\\Seen`).
  - `0x08000000` when the message has `\\Deleted`.

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
