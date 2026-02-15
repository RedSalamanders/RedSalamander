# Connection Manager Specification (Host-Owned, Themed, Secure Credentials)

## Overview

Some virtual file system plugins (FTP / FTPS / SFTP / SCP / IMAP / S3 / S3 Table and future protocols) need to ask the user for:

- a **connection target** (host/port/path),
- optional **authentication material** (user/password, SSH key + passphrase, known_hosts),
- and optional **persistence** of those secrets.

This spec defines a **host-owned Connection Manager** that:

- presents a consistent, themed dialog (host UI owns rendering),
- stores non-secret connection attributes in the Settings Store,
- stores secrets in Windows Credential Manager and gates access with **Windows Hello** (when available),
- provides a stable navigation contract so plugins do **not** require secrets in URIs.

## Goals

- One connection list shared by all protocols/plugins.
- The Connection Manager dialog may be opened with a `pluginId` filter (to show only connections relevant to that plugin), but storage remains global.
- Themed dialog consistent with Preferences and plugin configuration dialogs.
- Secure secrets storage:
  - secrets never written to JSON settings,
  - secrets stored as WinCred generic credentials,
  - optional Windows Hello verification before secrets are returned to plugins.
- Host-controlled navigation:
  - typing `nav:<connectionName>`, `nav://<connectionName>`, or `@conn:<connectionName>` navigates to the resolved endpoint,
  - navigating to `ftp:` / `sftp:` / `scp:` / `imap:` / `s3:` / `s3table:` with no host opens Connection Manager (filtered to that protocol),
  - main menu / command palette entry `Connections...` opens the dialog,
  - optional shorthand: `<scheme>://@conn/<connectionName>/...` routes to the named profile (authority `@conn`, e.g. `ftp://@conn/...`, `s3://@conn/...`).
- Extensible plugin integration (future protocols can participate without UI rewrite).

## Non-goals (initial milestone)

- Full SSH host-key management UI.
- Connection discovery / import/export.
- Multi-factor auth flows beyond Windows Hello gating (OTP, PKI, etc.).

## Data Model

### ConnectionProfile (persisted)

Stored in Settings Store (non-secret fields only):

- `id` (string, GUID): stable internal identifier.
- `name` (string): user-visible name, unique (case-insensitive), trimmed, and safe for `/@conn:<name>` (no `/` or `\\`).
- `pluginId` (string): target filesystem plugin long id (e.g. `builtin/file-system-sftp`).
- `host` (string): trimmed (no leading/trailing whitespace).
  - Semantics are plugin-defined. For S3 / S3 Table it stores the AWS region and may be empty to indicate “auto region”.
- `port` (uint32, `0` = protocol default)
- `initialPath` (string): remote initial folder (plugin path, typically `/`).
- `userName` (string)
- `authMode` (enum):
  - `anonymous` (FTP only),
  - `password`,
  - `sshKey` (SFTP/SCP).
  - FTP: when `authMode = password`, `userName` must be non-empty (anonymous is opt-in).
- `savePassword` (bool): whether a password/passphrase is stored in WinCred.
  - When `false`, the host may still prompt for a secret at connect time and cache it for the current app run (session-only; not persisted).
- `requireWindowsHello` (bool, default `true`): hidden expert setting; when `true`, Windows Hello verification is required before releasing secrets.
- `extra` (JSON object): plugin-specific non-secret fields (UTF-8 JSON object; schema is plugin-defined).
  - Example keys:
    - SFTP/SCP: `sshPrivateKey`, `sshKnownHosts`.
    - IMAP: `ignoreSslTrust` (bool): skip TLS certificate validation (allows self-signed certificates; not recommended).
    - S3/S3 Table: `endpointOverride`, `useHttps`, `verifyTls` (and `useVirtualAddressing` for S3 only).

When serialized to JSON, default-valued fields may be omitted (e.g. `authMode=password`, `savePassword=false`, `requireWindowsHello=true`, `initialPath=/`; for S3/S3 Table, `host` may be omitted when empty to represent “auto region”).

For built-in protocols, known default-valued `extra` keys may also be omitted (e.g. S3 `useHttps=true`, `verifyTls=true`, empty `endpointOverride`; SFTP/SCP empty `sshPrivateKey`).

### Quick Connect (session-only, not persisted)

The Connection Manager always exposes a synthetic first entry:

- Display label: `<Quick Connect>` (localized string)
- Internal `ConnectionProfile.name`: `@quick` (reserved; not user-editable)
- Internal `ConnectionProfile.id`: `00000000-0000-0000-0000-000000000001` (reserved)

Behavior:

- Stored in memory for the current app run and shown at the top of the list on every open.
- Never serialized to Settings Store JSON on disk.
- Not renameable and not removable.
- Secrets are stored in memory (not WinCred).

### ConnectionsSettings (persisted)

Stored in Settings Store under `connections`:

- `items` (array of ConnectionProfile)
- Global Windows Hello settings (Preferences → Advanced → Windows Hello for Connections):
  - `bypassWindowsHello` (bool, default `false`): when `true`, Windows Hello verification is skipped even if a profile requires it.
  - `windowsHelloReauthTimeoutMinute` (uint32, default `10`): how long a successful Windows Hello verification is reused for a given connection id (in minutes).
    - `0` means re-ask Windows Hello on every secret access.

### Secret Storage (not persisted in JSON)

Secrets are stored in Windows Credential Manager as **generic credentials**.

- Target name format: `RedSalamander/Connections/<connectionId>/<secretKind>` (internal stable id = `ConnectionProfile.id`, not the user name).
- Username is stored in the credential record as a convenience; the host remains the source of truth for `ConnectionProfile.userName`.
- Secret blob is UTF-16 NUL-terminated text.

Secret kinds (v1):
- `password`
- `sshKeyPassphrase`

Notes:
- `password` is also used for protocols that need a single secret value (e.g. S3 secret access key).

## Navigation Contract

### `nav:<connectionName>`

In NavigationView edit mode, the user can type:

- `nav:<connectionName>`
- `nav://<connectionName>`
- `@conn:<connectionName>`

The host resolves `<connectionName>` to a `ConnectionProfile` and navigates to:

- `<pluginShortId>:/@conn:<connectionName><initialPath>`

### Reserved plugin path prefix: `/@conn:<connectionName>/...`

When the host navigates to a connection, it passes a plugin path beginning with:

- `/@conn:<connectionName>/...`

Where `<connectionName>` is `ConnectionProfile.name` (unique case-insensitive name).

Plugins that support Connection Manager resolution must:

1. Detect `/@conn:` prefix (host-reserved).
2. Parse `connectionName` and the remaining remote path.
3. Query host services (`IHostConnections`) to obtain:
   - non-secret connection attributes,
   - secrets (password/passphrase) as needed.

This ensures secrets are never embedded in URIs.

## Host Services API (ABI)

The Connection Manager is exposed to plugins via a host service queried from `IHost`:

- `IHostConnections` (COM, `QueryInterface` from `IHost`)

### Methods (v1)

- `ShowConnectionManager(...)`: opens the host dialog and returns a selected connection name (`S_OK`) or cancel (`S_FALSE`).
- `GetConnectionJsonUtf8(...)`: returns the non-secret ConnectionProfile JSON (UTF-8).
  - Includes an `extra` object containing the full `ConnectionProfile.extra` payload (plugin-specific non-secret fields).
- `GetConnectionSecret(...)`: returns a secret if available (WinCred when `savePassword == true`, or session cache); does **not** prompt.
- `PromptForConnectionSecret(...)`: prompts and stores a session-only cached secret; does not persist to WinCred.
- `ClearCachedConnectionSecret(...)`: clears a session-cached secret (does not modify WinCred).
- `UpgradeFtpAnonymousToPassword(...)`: FTP-only: prompts for credentials, persistently flips `authMode` to `password`, and stages a session-only password.

### Behavioral contract

- Threading: plugins may call from any thread; host marshals UI work internally.
- Lifetime: host copies all strings before returning.
- Security:
  - secrets are never embedded in URIs and are only returned via host APIs,
  - if `savePassword == true`, the host loads secrets from WinCred (generic credentials),
  - if `savePassword == false`, the host never loads secrets from WinCred; instead it may:
    - return a **per-session cached** secret (from a prior prompt), or
    - prompt the user to enter a secret and cache it **in memory only** for the current app run,
  - if `requireWindowsHello == true`, the host performs Windows Hello verification prior to returning a secret from WinCred (host policy),
  - session-cached secrets are cleared on exit (and may expire after a host-defined TTL).

### Secret retrieval (prompting + session cache)

Plugins may request secrets in two phases:

- `GetConnectionSecret(...)`:
  - returns a secret if it is available (WinCred when `savePassword == true`, or a session-cached secret),
  - does **not** prompt,
  - returns `ERROR_NOT_FOUND` when no secret is available.
- `PromptForConnectionSecret(...)`:
  - shows a themed prompt for a secret (password/passphrase),
  - stores the entered secret in an in-memory per-session cache keyed by `(connectionId, secretKind)`,
  - returns `S_FALSE` if the user cancels the prompt.
  - SSH key passphrase may be empty to indicate “no passphrase”.

Prompt UX notes:

- The prompt displays both the user-visible connection name and a resolved target string (e.g. `sftp://user@host:22`) so users can confirm what they are authorizing.
- The host must never log secrets (including on failures).

### FTP: anonymous rejected

If an FTP server rejects anonymous login, the plugin may ask the host to prompt for credentials and persistently flip the connection profile to `authMode = password` (while keeping the password session-only unless explicitly saved via Connection Manager).

## UI: Connection Manager Dialog

Layout (using RedSalamander theming):

- Left pane: connection list
  - `New…`, `Rename…`, `Remove…`
- Right pane: connection editor
  - Scrollable when the editor content exceeds the available height (bottom buttons remain pinned and content does not overlap them).
  - Address (host), Port, Initial path
  - User name, Password/Passphrase (masked, with Show/Hide control)
    - If a secret is already stored, the field shows a random-length masked placeholder to indicate “a password is saved” without leaking the real length.
  - `Save password`
    - When unchecked, the connection is still usable: the host prompts for the password/passphrase at connect time and keeps it in memory only for the current app run.
  - IMAP: `Ignore trust for SSL` (optional)
  - S3 / S3 Table:
    - `Region` uses an editable dropdown populated with known AWS region names/codes; selecting an entry inserts the region code into the field.
    - `Endpoint override` (optional; for S3-compatible endpoints)
    - `Use HTTPS`, `Verify TLS certificate`
    - S3 only: `Use virtual-hosted style addressing`
  - Hidden expert setting: `requireWindowsHello` is not exposed in UI, but can be set in Settings Store JSON per profile.
  - Protocol-specific fields (SFTP/SCP): SSH key paths, known_hosts path
- New connection defaults:
  - When creating a new connection, the host may prefill host/port/path/user from the target plugin's configured defaults (e.g. `defaultHost`, `defaultPort`, `defaultBasePath`, `defaultUser`).
  - FTP: anonymous login is off by default.
- Bottom buttons:
  - `Connect` (OK): saves config, resolves selected/edited connection, and closes dialog
  - `Close`: saves config and closes dialog (no navigation)
  - `Cancel`: closes without saving changes (no navigation)

Connect-time secret behavior:

- If `authMode = password` and `savePassword = true` but no stored password exists yet, the host prompts for the password and saves it to WinCred before closing.

## FTP / SFTP / SCP / IMAP plugin behavior (initial)

- Navigating to `ftp:` / `sftp:` / `scp:` / `imap:` with no host opens Connection Manager (host-side behavior).
- The FileSystemCurl plugin supports the `/@conn:<connectionName>/...` prefix and uses host services to resolve connection + credentials.
