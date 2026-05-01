# FTP / SFTP / SCP / IMAP Virtual File System Plugins

## Overview

RedSalamander ships a built-in virtual file system implementation in `Plugins/FileSystemCurl/FileSystemCurl.dll` that exposes **four UI-visible file system plugins**:

- **FTP** (`ftp:`)
- **SFTP** (`sftp:`)
- **SCP** (`scp:`)
- **IMAP** (`imap:`)

All four are implemented by the same DLL and share as much code as possible. The transfer backend is an implementation detail (not surfaced in UI).

## Plugin Identities

| Display name | `PluginMetaData.id` | `PluginMetaData.shortId` |
|-------------|----------------------|--------------------------|
| FTP         | `builtin/file-system-ftp`  | `ftp`  |
| SFTP        | `builtin/file-system-sftp` | `sftp` |
| SCP         | `builtin/file-system-scp`  | `scp`  |
| IMAP        | `builtin/file-system-imap` | `imap` |

## Navigation (URI Syntax)

### FTP

Supported navigation forms:
- `ftp://[user-info@]host[:port][/path]`
- `ftp:/path` (uses configured defaults; requires `defaultHost`)

Notes:
- If no user is provided, the plugin uses `anonymous`.
- `user-info` is treated as `user[:password]` (legacy RFC1738-style user info).
- In Connection Manager, anonymous login is opt-in (default auth mode is `password`).

Examples:
- `ftp://example.com/`
- `ftp://user:password@example.com:21/pub/`
- With defaults (`defaultHost = example.com`): `ftp:/pub/`

### SFTP

Supported navigation forms:
- `sftp://[user[:password]@]host[:port][/path]`
- `sftp:/path` (uses configured defaults; requires `defaultHost`)

Notes:
- `user` and `password` in the URI override defaults.
- For key-based auth, use the SSH settings (below).

Examples:
- `sftp://user@example.com/home/user/`
- `sftp://user:password@example.com:2222/var/log/`
- With defaults (`defaultHost = example.com`): `sftp:/`

### SCP

Supported navigation forms:
- `scp://[user@]host[:port][/path]`
- `scp:/path` (uses configured defaults; requires `defaultHost`)

Notes:
- Passwords in the URI are **not allowed** for SCP (`scp://user:password@host/...` fails).
- Use settings (`defaultPassword`) and/or SSH key settings for authentication.

Examples:
- `scp://user@example.com/etc/`
- `scp://example.com/` (uses `defaultUser` if set)

### IMAP

See `Specs/FileSystem/FileSystem_Imap.md`.

### Path normalization rules

- `\` is treated as `/`.
- Multiple slashes are collapsed **except** the leading `//` after the scheme (authority prefix).
- `ftp://host/path` is treated as the authority form; `ftp:/path` is treated as “use defaults”.

## Configuration (Per Plugin)

Each of the four plugins exposes its own configuration schema via `IInformations::GetConfigurationSchema()`. The following keys are used:

### Common keys (FTP / SFTP / SCP)

- `defaultHost` (string): host used for `scheme:/...`
- `defaultPort` (integer, `0` = protocol default)
- `defaultUser` (string)
- `defaultPassword` (string, stored as plain text)
- `defaultBasePath` (string): remote base folder for `scheme:/...`
- `connectTimeoutMs` (integer)
- `operationTimeoutMs` (integer, `0` = no timeout)
- `copyMoveMaxConcurrency` (integer, default `4`, min `1`, max `8`): maximum concurrency for copy/move (top-level items and files within a single folder copy). Higher values may open multiple connections; some servers may throttle or reject parallel transfers.
- `deleteMaxConcurrency` (integer, default `4`, min `1`, max `8`): maximum concurrency for delete (top-level items and deletes within a single recursive directory delete). Higher values may open multiple connections; some servers may throttle or reject parallel deletes.

URI override behavior:
- If the navigated URI specifies `user`, `password`, or `port`, those values override the defaults for that navigation path.

Security note:

- `defaultPassword` / `sshKeyPassphrase` are stored as plain text in plugin settings. For secure storage (WinCred + optional Windows Hello), prefer Connection Manager (`/@conn:<connectionName>/...`).

### FTP-only keys

- `ftpUseEpsv` (bool): toggles EPSV usage (recommended for most servers)

### SSH keys (SFTP / SCP)

- `sshPrivateKey` (string): file path to the private key
- `sshPublicKey` (string): file path to the public key (optional)
- `sshKeyPassphrase` (string, stored as plain text; empty means “no passphrase”)
- `sshKnownHosts` (string): path to `known_hosts` (empty disables strict host key checking)

## Operations and Behavior Notes

- File copy/move across different endpoints (host/user/port) within the same protocol plugin is implemented as **download → upload** via a local temporary file.
- Directory operations support recursion when `FILESYSTEM_FLAG_RECURSIVE` is provided.
- Directory listing is parsed incrementally (streamed) to avoid buffering the full LIST output in memory, improving large-directory memory behavior.
- For a **single** top-level directory copy/move or delete, contained file work SHOULD still parallelize (within-folder parallelism) subject to the effective concurrency limits.
- SCP has protocol limitations; directory listing and command-style operations require the server to support SFTP over SSH.

## Properties Metadata

- `GetItemProperties` returns structured sections for general item data, remote/display paths, connection metadata, and any available protocol-specific metadata.
- FTP, SFTP, and SCP directory-listing timestamps are best-effort:
  - Unix-style and DOS-style listing timestamps populate `lastWriteTime` when parseable.
  - The plugin MUST NOT synthesize creation, access, or change timestamps when the remote listing does not provide them.
  - Unavailable timestamp fields MUST be omitted from the properties JSON instead of emitted as `0`.
- SCP uses the SFTP command/listing path for directory metadata, so the same timestamp rules apply to SCP directory entries.
- IMAP-specific property rules are documented in `Specs/FileSystem/FileSystem_Imap.md`.

## Connection Manager Integration

All four protocols support host-reserved navigation:

- `/@conn:<connectionName>/...` (see `Specs/Core/Core_ConnectionManager.md`)

Concurrency overrides (Connection Manager profiles):

- Connection profiles may specify host-level concurrency caps in `ConnectionProfile.extra`:
  - `copyMoveMaxConcurrency` (integer): `0`/missing = inherit the plugin setting; otherwise clamp `1..8`.
  - `deleteMaxConcurrency` (integer): `0`/missing = inherit the plugin setting; otherwise clamp `1..8`.
- These overrides are global per connection profile across the app:
  - the host uses them to clamp per-task file-op concurrency when `@conn` paths are involved,
  - and the plugin enforces a global per-profile limiter so multiple simultaneous tasks do not exceed the cap.
- Note: IMAP is not currently used for file operations (copy/move/delete), so these keys have no practical effect for IMAP profiles.

Resolution + secret acquisition (high level):

- The plugin resolves `/@conn:...` by requesting the non-secret profile JSON from the host (`IHostConnections::GetConnectionJsonUtf8`).
- If the resolved profile requires a secret:
  - Call `IHostConnections::GetConnectionSecret` first (no prompt).
  - If it returns `ERROR_NOT_FOUND`, call `IHostConnections::PromptForConnectionSecret` to prompt the user and cache the secret for the current app run.
  - If the prompt returns `S_FALSE`, treat it as user cancellation (`ERROR_CANCELLED`).
  - SSH key passphrase may be empty to indicate “no passphrase”.

Authentication failure retry policy (Connection Manager profiles only):

- FTP anonymous rejection:
  - If a server rejects anonymous login for a Connection Manager profile with `authMode = anonymous`, the plugin asks the host to upgrade it (`IHostConnections::UpgradeFtpAnonymousToPassword`) and retries once.
- Session-only secrets (`savePassword = false`):
  - On login/auth failures, the plugin clears the session cache (`IHostConnections::ClearCachedConnectionSecret`) and retries once (which triggers a reprompt on the next resolve).

