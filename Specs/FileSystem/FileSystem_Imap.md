# IMAP Virtual File System Plugin

## Overview

RedSalamander ships a built-in IMAP virtual file system implemented by `Plugins/FileSystemCurl/FileSystemCurl.dll`.

IMAP remains export-only. Capability v2 identifies a message by `UIDVALIDITY` plus UID when the
transport can bind it; absent a conditional exact delete, Move uses Copy-only and reports
`Copied; source kept`. It never expunges a message selected only by mutable pathname text.

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
- Epoch-qualified message access: `imap://user@example.com/INBOX/Quarterly%20report%20[777-12345].eml`
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

Directory-size queries retain the IMAP mailbox/UID listing adapter, rather than
interpreting mail responses as FTP LIST rows. Their outer traversal is iterative
and uses the shared depth/retained-frame ceilings; unknown message sizes produce
partial totals, not successful zero bytes. The shared size callback adapter preserves
progress/cancel ordering and primary errors. IMAP custom requests consume current
per-call operation control, and callback aborts retain their cancellation/error
verdict without logging the request as a server failure.

The mailbox/UID/summary facade still buffers internally. Charging its returned
entries does not bound its transient allocations or qualify real-server cancellation;
the `imapBuffered` size metrics make this distinction explicit. FTP loopback tests
are not IMAP qualification. IMAP containment capabilities are unchanged by this
size/callback repair.

### Directories (mailboxes)

- Root (`imap:/`) lists mailboxes (or the subtree under `defaultBasePath`).
- Entering a mailbox directory lists:
  - sub-mailboxes as directories, and
  - messages as `.eml` files (when the mailbox is selectable).

### Files (messages)

- Directory listing names are formatted as `<subject> [<uidValidity>-<uid>].eml`.
  - Example: `Quarterly report [777-12345].eml`.
  - `<subject>` is best-effort decoded from RFC2047 encoded-words (Q/B) into UTF-16; missing/empty becomes `(no subject)`.
  - RFC2047 decoding supports UTF-8, US-ASCII, ISO-8859 families, Windows code pages, and common Japanese, Chinese, Korean, Cyrillic, Thai, and Macintosh MIME charset aliases through Windows code-page conversion. Unknown charsets fall back to UTF-8, the active Windows code page, then Windows-1252.
  - Malformed encoded-word fragments already using the sanitized `=_charset_q_...` shape are decoded as a recovery path before filename sanitization. U+00AD soft hyphen in decoded subjects is normalized to a visible `-` for pane and property display.
  - Subject text is sanitized for Windows filenames and capped at 96 UTF-16 code units. If it exceeds the cap, the first 93 code units are kept and ASCII `...` is appended before the epoch-qualified suffix.
  - The final bracketed suffix is the stable operation key: mailbox path plus UIDVALIDITY plus UID identify the message. Subjects are display text only and are never used for message identity.
- Directory enumeration retrieves message metadata using:
  - `UID FETCH <uid-set> (UID FLAGS INTERNALDATE RFC822.SIZE ENVELOPE)`
  - Every summary set is listing-shaped for libcurl (`a,b` or `uid:uid`). A bare
    single UID is rewritten to the one-element range so libcurl does not enter
    its literal-download path, which truncates a row after the first `{N}` body.
  - Summary fetches capture libcurl's wire-line (`HEADERFUNCTION`) channel, not
    the body callback. On the pinned libcurl 8.21.0, a custom `UID FETCH` whose
    set contains `,` or `:` writes no body bytes (`imap.c` `is_custom_fetch_listing`).
    SEARCH/LIST/CAPABILITY and RFC822 body downloads keep the body channel.
  - A damaged FETCH row is `ERROR_BAD_NET_RESP`, never `ERROR_FILE_NOT_FOUND`.
    Absence is reported only when a well-formed reply simply lacks that UID.
- Before fetching metadata, `UID SEARCH ALL` must provide at least one complete, case-insensitive SEARCH response. Each UID is a complete nonzero unsigned-32-bit token; malformed tokens, missing replies and truncated framing fail the listing with `ERROR_BAD_NET_RESP`, without publishing partial entries or starting metadata repair. A valid empty SEARCH remains a successful empty mailbox, even if an earlier STATUS count differed. Multiple complete SEARCH rows are combined and duplicate UIDs are removed before metadata requests and descending-UID projection.
- Unsolicited data responses use borrowed quoted/literal/list framing, so message text that resembles a SEARCH line cannot establish identity. Status text remains free-form and does not use literal framing. The current parser does not implement ESEARCH correlation or UID-range expansion: that response returns `ERROR_NOT_SUPPORTED`, never a successful empty mailbox. Enabling or qualifying IMAP4rev2 remains separate from these rev1 response guarantees.
- UID parsing checks current operation control after the request, between response frames, every 256 parsed UIDs and before completion. Control failures remain failures, including a canceled empty result; no FETCH or partial successful listing follows them. These checks do not establish a whole-buffer memory bound or real transport cancellation guarantee.
- When a bulk summary fetch has a recoverable metadata failure, omits requested UIDs, or leaves required metadata incomplete, directory enumeration repairs those summaries in bounded smaller UID batches and then single UID fetches before falling back to `message [<uidValidity>-<uid>].eml`, unknown size, or unavailable timestamps. Completeness requires known size, flags, valid internal date and parsed envelope (or a complete root `BODY[HEADER.FIELDS (SUBJECT FROM DATE)]` response); valid NIL/empty envelope values are known empty, not missing. Cancellation, operation-control failure and expired request/operation deadlines stop the listing immediately, without further repair requests or publication of partial entries as a successful listing. A request-local latch preserves the first operation-control failure; it is not reconstructed by invoking the callback again.
- Summary workspace is reused for each group of at most 200 requested UIDs, then converted to pane entries in descending UID order. Unsolicited FETCH rows for UIDs outside the current request are ignored; FLAGS-only updates for a requested UID replace flags without erasing already acquired size, subject or timestamps. This bounds summary-map entry count, not raw response bytes, mailbox/UID vectors, output entries or opaque transport memory.
- FETCH metadata comes only from complete top-level attributes, independently of order and ASCII case. Quoted/literal message text, nested extension values and opaque BODY payloads cannot establish UID, size, flags or envelope fields. Duplicate recognized metadata and malformed/truncated framing cannot establish a summary. A framing failure stops parsing that response: apparent subsequent FETCH lines might be literal message text, so missing metadata is repaired through a separate bounded request.
- Missing, malformed, overflowing or signed `RFC822.SIZE` values remain unknown in both directory listings and targeted lookups. Only a complete valid unsigned token within the RFC 9051 unsigned-63-bit range establishes size; explicit `0` is a known empty message. FETCH UID and message sequence numbers must be nonzero unsigned-32-bit values. Literal lengths use checked protocol-range parsing and remaining-byte subtraction before offset addition. Framing is iterative with constant stack storage and borrowed payload views; this does not bound the already buffered response itself. Unknown sizes cannot be used as zero-byte completion or source-stability proof.
- Legacy UID-only names such as `<uid>.eml` and `<subject> [<uid>].eml` are parseable only for compatibility diagnostics; version-sensitive fetch/delete operations reject them. Refresh the mailbox to obtain an epoch-qualified identity.
- Ambiguous names such as `Quarterly report 12345.eml` are not valid IMAP message keys.
- Before opening a message, the plugin re-reads mailbox STATUS UIDVALIDITY. A missing epoch fails closed; a changed epoch returns `ERROR_REVISION_MISMATCH` so the stale path cannot fetch a different message with a reused UID. Only then does `UID FETCH <uid> BODY.PEEK[]` expose the RFC822 content as a read-only file.

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

Unavailable or malformed FLAGS do not imply unread, marked or deleted. Listing
repair attempts to obtain them within the same metadata budget; after exhaustion,
those attributes remain unavailable rather than being fabricated.

- `GetItemProperties` exposes general item data, remote/display paths, connection metadata, and IMAP metadata when available.
- Mailbox directories commonly do not have meaningful filesystem timestamps; unavailable timestamp fields MUST be omitted rather than shown as `0`.
- Message properties may include sent time, received/internal time, flags, UID, UIDVALIDITY-qualified path identity, sender, subject, and size. Standard timestamp fields are included only when non-zero.
- Message size metadata is omitted when its value is unknown; no missing size is synthesized as zero.
- Unknown flags and unavailable subject/from metadata are omitted, as are zero/unavailable IMAP timestamp fields. A successfully parsed empty envelope may still expose empty subject/from values.
- Message properties MUST fetch metadata for the selected UID only; they must not enumerate and summarize the full mailbox just to show one message's Properties.
- Mailbox folder properties include the plugin mailbox path, server mailbox path, hierarchy delimiter, root/mailbox identity, connection metadata, and `STATUS` counts (`MESSAGES`, `RECENT`, `UIDNEXT`, `UIDVALIDITY`, `UNSEEN`) when the server provides them.
- Passwords and raw secrets are never emitted in properties; boolean presence fields such as `hasPassword` are sufficient.

## Pane Display Contract

For a message with subject `Quarterly report`, UID `12345`, received time `2026-05-13 21:33`, and RFC822 size `45 KB`:

| Pane mode | Display |
|-----------|---------|
| Brief | `Quarterly report [777-12345].eml` |
| Detailed | Name line plus `2026-05-13 21:33 • 45 KB • -` on the details line |
| Extra Detailed | Same name and details line unless a later metadata provider adds a third line |
| Thumbnails | File icon/thumbnail slot with the same name and details text |
| Status bar | `45 KB • 2026-05-13 21:33 • -` for a focused or singly selected message |

The full subject and sender belong in Properties. The filename may be truncated for display and layout, but the bracketed UIDVALIDITY/UID identity remains visible at the end of the logical leaf name.

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
- `filesystem.imap.summary_workspace_peak_entries`
- `filesystem.imap.build_fileinfo_us`
- `filesystem.imap.properties_message_us`
- `filesystem.imap.properties_mailbox_us`
- `filesystem.imap.status_mailbox_us`
- `filesystem.imap.capability_us`

Message Properties performance is protected by a command-count contract: opening Properties for one message must stay constant with respect to mailbox size. The selected UID path uses targeted UID metadata fetches rather than a full mailbox summary refetch.

Mailbox listing summary repair is bounded per listing. The repair path may batch missing or incomplete summaries first and then fall back to singleton fetches, but it MUST cap repair fetch attempts with a per-listing budget, log one warning when the budget is exhausted, and proceed with any summaries already recovered instead of issuing unbounded singleton fetches against flaky servers. Stop/control failures are not recoverable metadata failures and do not consume this fallback. Parse-anomaly diagnostics retain counts, not raw response prefixes, subjects, addresses or literal payloads. The connection-owned `C0_CurlImapListingTruth` fixture exercises the real parser/repair owner with 1/201/4,097 messages, unknown sizes, unsolicited updates, structural attributes, incomplete metadata and stop failures. It is not live transport or TLS proof.

UIDVALIDITY/UIDPLUS security validation has a constant command-count contract independent of mailbox size: one
STATUS command per mailbox listing or message fetch/delete, plus one CAPABILITY command per delete. STATUS latency is
measured by `filesystem.imap.status_mailbox_us`; CAPABILITY latency, response bytes, UIDPLUS availability, and HRESULT
are measured by `filesystem.imap.capability_us`. Credential-free Release evidence is archived under
`Specs/TestRuns/4cb089111a23/FileSystemCurl/2026-07-16_2028_observatory_track2_imap_safety/`; live network latency
remains explicitly unclaimed until a deterministic network-capable IMAP fixture is available.

## Operations

### Browse

- `ReadDirectoryInfo()`:
  - Root lists mailboxes.
  - Mailbox directories list sub-mailboxes and message `.eml` files.

### Read

- `CreateFileReader()` downloads the message into a temporary delete-on-close file and returns an `IFileReader`.
- IMAP inspection remains available, but the exact capability response uses
  `routeClass:"uncontained"` and `providerWatchdogTimeoutMs:0`; a live libcurl operation or reader
  is not task-cancellable. File Operations therefore rejects mutation before task creation.

### Delete

- Deleting a message file:
  - Re-read STATUS and require the enumerated UIDVALIDITY; return `ERROR_REVISION_MISMATCH` when it changed.
  - Issue CAPABILITY and require the exact `UIDPLUS` token. A server without UIDPLUS is rejected with
    `ERROR_NOT_SUPPORTED` before the target is marked deleted.
  - `UID STORE <uid> +FLAGS.SILENT (\Deleted)`.
  - `UID EXPUNGE <uid>`.
  - If safe expunge fails, preserve that original failure and attempt
    `UID STORE <uid> -FLAGS.SILENT (\Deleted)` as compensation. Cleanup failure is retained in diagnostic context.
  - Never issue mailbox-wide `EXPUNGE` as a single-message fallback; unrelated pre-existing `\Deleted` messages
    must remain untouched.
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
