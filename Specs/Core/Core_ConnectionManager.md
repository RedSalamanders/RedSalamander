# Connection Manager Specification (Host-Managed, Themed, Secure Credentials)

## Implementation status

The Connection Manager is implemented as a **single-canvas DxUi top-level window** (`RedSalamander/ConnectionManagerWindow.{h,cpp}`). It registers one `WS_OVERLAPPEDWINDOW` HWND, attaches a single `DxUi::WindowHost`, and renders the entire list/editor/footer chrome through one DxUi widget tree (no per-widget host HWNDs, no Win32 dialog template). The synchronous `ShowConnectionManagerDialog` entry point is a façade that runs a private nested message pump over the same modeless window with the owner disabled, so plugin-host callers (`HostServices.cpp:1419`) keep their `S_OK`/`S_FALSE` ABI.

Migration record: `Specs/Plans/Done/UI_ConnectionManagerSingleCanvasPlan.md`.

## Single-Canvas Connection Manager Window Contract

- The live implementation is `RedSalamander/ConnectionManagerWindow.{h,cpp}`. `ConnectionManagerDialog.h` and `ConnectionManagerDialog.cpp` are retired and must not be reintroduced.
- `ShowConnectionManagerWindow(...)` is the normal application entry point. It is modeless, single-instance, and posts `WndMsg::kConnectionManagerConnect` to the owner with a copied connection-name payload and the requested target pane after Connect validates and saves.
- `ShowConnectionManagerDialog(...)` is a synchronous facade over the same single-canvas window for host-service callers that require the existing `S_OK` / `S_FALSE` result contract.
- The top-level window class MUST register `CS_DBLCLKS` so native double-click gestures reach DxUi text fields. Text fields in the real Connection Manager window must support bridge-backed character input, word selection on double-click, and masked secret editing without exposing secret text through clipboard copy/cut.
- The Connection Manager window and connection credential prompts MUST apply the persisted `ui.windowBackdrop` setting through the shared window chrome/backdrop helper path with tool-window target semantics. High contrast or unsupported OS state MUST resolve to no system backdrop, and activation handling MUST keep title-bar active/inactive state correct without changing the persisted backdrop policy.
- Profile names are trimmed before save and must be non-empty, unique case-insensitively, and not reserved for Quick Connect. A saved name must resolve to exactly one persisted profile.
- The settings hot-reload contract matches Preferences: clean windows reload from disk automatically, dirty windows prompt before reloading, and stale save-producing actions prompt before overwriting disk state.
- Connection credential prompts center over their valid owner without changing size, z-order, or activation.
  The prompt uses the shared owner-centering geometry contract; an invalid owner or failed geometry query leaves
  the existing position unchanged.

## Overview

Some virtual file system plugins (FTP / FTPS / SFTP / SCP / IMAP / Google Drive / S3 / S3 Table / OneDrive Personal / OneDrive Business / SharePoint and future protocols) need to ask the user for:

- a **connection target** (host/port/path),
- optional **authentication material** (user/password, SSH key + passphrase, known_hosts),
- and optional **persistence** of those secrets.

This spec defines a **host-managed Connection Manager** that:

- presents a consistent, themed dialog (host UI owns rendering),
- opens its long-lived tool window using the independent top-level contract from `Specs/UI/UI_TopLevelToolWindows.md`,
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
- navigating to `ftp:` / `sftp:` / `scp:` / `imap:` / `gdrive:` / `s3:` / `s3table:` / `onedrive:` / `onedrive-pro:` / `sharepoint:` with no host opens Connection Manager (filtered to that protocol),
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

- `id` (string, GUID): stable internal identifier in lowercase canonical `8-4-4-4-12` form.
  - Persisted IDs must be unique after case folding; the reserved Quick Connect ID is forbidden in persisted profiles.
  - Strict reload, Connection Manager commit, and Settings Store save reject invalid, non-canonical, reserved, or duplicate IDs.
  - Startup recovery gives every member of an existing collision group (and every other invalid-ID profile) a fresh unique ID.
    Because an old WinCred target cannot be associated safely once identities collided, recovery sets `savePassword=false`,
    leaves old WinCred entries untouched, persists the repaired settings through source CAS, and tells the user to re-enter
    the affected secret. Recovery never copies or guesses an ambiguous credential.
- `name` (string): user-visible name, unique (case-insensitive), trimmed, and safe for `/@conn:<name>` (no `/` or `\\`).
- `pluginId` (string): target filesystem plugin long id (e.g. `builtin/file-system-sftp`).
- `host` (string): trimmed (no leading/trailing whitespace).
  - Semantics are plugin-defined.
  - S3 / S3 Table: stores the AWS region and may be empty to indicate `auto region`.
  - Google Drive: unused by the built-in plugin and normally empty.
  - OneDrive Personal / OneDrive Business: unused by the built-in plugin and normally empty.
  - SharePoint: stores the tenant host name with an optional site path suffix (for example `contoso.sharepoint.com/sites/Team`).
- `port` (uint32, `0` = protocol default)
- `initialPath` (string): remote initial folder (plugin path, typically `/`).
- `userName` (string)
- `authMode` (enum):
  - `anonymous` (FTP only),
  - `password`,
  - `sshKey` (SFTP/SCP),
  - `oauth2Pkce` (Google Drive / OneDrive Personal / OneDrive Business / SharePoint).
  - FTP: when `authMode = password`, `userName` must be non-empty (anonymous is opt-in).
- `savePassword` (bool): whether a password/passphrase is stored in WinCred.
  - When `false`, the host may still prompt for a secret at connect time and cache it for the current app run (session-only; not persisted).
  - For `authMode = oauth2Pkce`, this means `remember sign-in`: the host persists the OAuth refresh token when checked and keeps it session-only when unchecked.
- `requireWindowsHello` (bool, default `true`): hidden expert setting; when `true`, Windows Hello verification is required before releasing secrets.
- `extra` (JSON object): plugin-specific non-secret fields (UTF-8 JSON object; schema is plugin-defined).
  - Example keys:
    - SFTP/SCP: `sshPrivateKey`, `sshKnownHosts`.
    - IMAP: `ignoreSslTrust` (bool): skip TLS certificate validation (allows self-signed certificates; not recommended).
    - Google Drive:
      - `rootKind` (string): `myDrive` or `sharedDrive`.
      - `sharedDriveId` (string): required when `rootKind = sharedDrive`.
      - `googleDocsMode` (string): export strategy for native Google Docs items.
      - `readOnly` (bool): disables writes once write support exists.
      - `useDefaultClientId` (bool): uses the plugin-level default OAuth client id when `true`.
      - `clientId` (string): optional per-connection OAuth client id override when `useDefaultClientId = false`.
    - S3/S3 Table: `endpointOverride`, `useHttps`, `verifyTls` (and `useVirtualAddressing` for S3 only).
    - OneDrive Personal / OneDrive Business / SharePoint:
      - `tenantAuthority` (string): optional Microsoft Entra authority override such as `consumers`, `organizations`, or a tenant id.
      - `driveId` (string): optional Graph drive id. The built-in SharePoint plugin uses it to pin a specific document library instead of the site's default drive.
    - Global file operations (all file-op capable protocols):
      - `copyMoveMaxConcurrency` (uint32): `0 = inherit plugin setting`, else clamp to `1..8`.
      - `deleteMaxConcurrency` (uint32): `0 = inherit plugin setting`, else clamp to `1..64`.
      - These overrides apply **globally per connection profile across the whole app** (concurrent tasks share the cap).

When serialized to JSON, default-valued fields may be omitted (e.g. `authMode=password`, `savePassword=false`, `requireWindowsHello=true`, `initialPath=/`; for S3/S3 Table and Google Drive, `host` may be omitted when empty to represent a hostless profile).

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
  - `bypassWindowsHello` (bool, default `false`): when `true`, Windows Hello verification is skipped even if a profile requires it (intended for automation).
  - `windowsHelloReauthTimeoutMinute` (uint32, default `10`): how long recent interactive authentication is reused for a given connection id and secret kind (in minutes) for interactive prompts/UX.
    - Interactive authentication includes Windows Hello verification and manual secret entry (password/passphrase).
    - `0` means re-ask Windows Hello on every secret access.
    - The same successful interactive action creates an explicit app-run grant for background plugin access to that exact
      `(connectionId, secretKind)`, so long-running copy/compare work does not prompt mid-operation.
    - Background grants never satisfy an interactive reveal after its timed grant expires.

### Secret Storage (not persisted in JSON)

Secrets are stored in Windows Credential Manager as **generic credentials**.

- Target name format: `RedSalamander/Connections/<connectionId>/<secretKind>` (the canonical internal stable id = `ConnectionProfile.id`, not the user name). Target construction rejects non-canonical IDs.
- Username is stored in the credential record as a convenience; the host remains the source of truth for `ConnectionProfile.userName`.
- Secret blob is UTF-16 NUL-terminated text.

Secret kinds (v1):
- `password`
- `sshKeyPassphrase`
- `refreshToken`

Notes:
- `password` is also used for protocols that need a single secret value (e.g. S3 secret access key).
- `refreshToken` is used by OAuth 2.0 PKCE connections to persist the long-lived refresh token; access tokens stay in memory only.

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
- `GetConnectionSecret(...)`: returns a secret if available (WinCred when `savePassword == true`, or session cache); does **not** prompt for secret entry.
- `PromptForConnectionSecret(...)`: prompts and stores a session-only cached secret; does not persist to WinCred.
- `SetConnectionSecret(...)`: writes a secret into the session cache and optionally persists it to WinCred. The session value is updated before durable persistence; a WinCred failure is returned to the caller, but the current process continues using the new session value instead of falling back to the stale durable secret.
- `DeleteConnectionSecret(...)`: removes the secret from the session cache and optionally from WinCred. The host records a session tombstone before durable deletion; a WinCred failure is returned to the caller, and the tombstone prevents a later read in the same process from resurrecting the stale durable secret.
- `ClearCachedConnectionSecret(...)`: clears a session-cached secret (does not modify WinCred).
- `UpgradeFtpAnonymousToPassword(...)`: FTP-only: prompts for credentials, persistently flips `authMode` to `password`, and stages a session-only password.

### Behavioral contract

- Threading: plugins may call from any thread; host marshals UI work internally. If the host window/owner state is unavailable, connection APIs that require host-owned UI/settings/secret state MUST return `HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE)` instead of executing UI-thread bodies on the caller thread.
- Lifetime: host copies all strings before returning.
- Security:
  - secrets are never embedded in URIs and are only returned via host APIs,
  - if `savePassword == true`, the host loads secrets from WinCred (generic credentials),
  - if `savePassword == false`, the host never loads secrets from WinCred; instead it may:
    - return a **per-session cached** secret (from a prior prompt), or
    - prompt the user to enter a secret via `PromptForConnectionSecret(...)` and cache it **in memory only** for the current app run,
  - OAuth 2.0 PKCE refresh tokens follow the same persistence rule, but they are written through `SetConnectionSecret(...)` after token acquisition or refresh,
  - if `requireWindowsHello == true` (and `bypassWindowsHello == false`), the host may perform Windows Hello verification prior to returning a secret from WinCred (host policy),
    - The host MUST reuse the explicit app-run background grant created by successful interactive authentication for the
      same `(connectionId, secretKind)`; this grant does not bypass the timed policy for an interactive reveal.
  - authorization state is keyed by canonical connection ID, secret kind, and access purpose. It is cleared for the affected
    secret on replacement/deletion and cleared globally on workstation lock/disconnect/logoff, session end, and process shutdown.
  - session-cached secrets are cleared on workstation lock/disconnect/logoff, session end, and process exit (and may expire after a host-defined TTL).

### Secret retrieval (prompting + session cache)

Plugins may request secrets in two phases:

- `GetConnectionSecret(...)`:
  - returns a secret if it is available (WinCred when `savePassword == true`, or a session-cached secret),
  - does **not** prompt for secret entry,
  - may require Windows Hello verification prior to returning a secret from WinCred (host policy),
  - returns `ERROR_NOT_FOUND` when no secret is available.
- `PromptForConnectionSecret(...)`:
  - shows a themed prompt for a secret (password/passphrase),
  - stores the entered secret in an in-memory per-session cache keyed by `(connectionId, secretKind)`,
  - returns `S_FALSE` if the user cancels the prompt.
  - SSH key passphrase may be empty to indicate “no passphrase”.
- OAuth refresh tokens are not manually prompted; Microsoft Drive plugins acquire them via browser sign-in, while Google Drive currently receives them from host-owned OAuth plumbing or tests and persists them with `SetConnectionSecret(...)`.

Prompt UX notes:

- The prompt displays both the user-visible connection name and a resolved target string (e.g. `sftp://user@host:22`) so users can confirm what they are authorizing.
- The host must never log secrets (including on failures).

### FTP: anonymous rejected

If an FTP server rejects anonymous login, the plugin may ask the host to prompt for credentials and persistently flip the connection profile to `authMode = password` (while keeping the password session-only unless explicitly saved via Connection Manager).

## UI: Connection Manager Dialog

Layout (using RedSalamander theming):

- Windowing: the command-surface Connection Manager window is modeless and independent from the main window. It may use the invoking window only for initial placement or context, not ownership.

- Left pane: connection list
  - `<Quick Connect>` is always the first row.
  - All other visible connections are sorted alphabetically by display name (case-insensitive).
  - The single-column profile list must not expose a horizontal scrollbar.
  - `New…`, `Rename…`, `Remove…`
  - The list and command-button row consume the full left pane height; the gap below the buttons should visually match the top window/list gap.
- Right pane: connection editor
  - Scrollable when the editor content exceeds the available height (bottom buttons remain pinned and content does not overlap them).
  - Fields are grouped into cards under section titles, matching Preferences: section titles sit above their card, cards must not touch or overlap vertically, and card content must keep a visible gap from the right edge/scrollbar.
  - Address (host), Port, Initial path
  - Copy/Move max concurrency (0 = default), Delete max concurrency (0 = default)
  - User name, Password/Passphrase (masked, with Show/Hide control)
    - If a secret is already stored, the field shows a random-length masked placeholder to indicate “a password is saved” without leaking the real length.
    - Revealing a stored placeholder loads the secret only after the user clicks Show; when `Require Windows Hello` is enabled and not bypassed, the reveal must pass Windows Hello before the real secret is placed in the editable field.
  - `Save password`
    - When unchecked, the connection is still usable: the host prompts for the password/passphrase at connect time and keeps it in memory only for the current app run.
    - For OAuth 2.0 PKCE profiles, the label means `Remember sign-in` and controls whether the refresh token is persisted.
  - IMAP: `Ignore trust for SSL` (optional)
  - S3 / S3 Table:
    - `Region` uses an editable dropdown populated with known AWS region names/codes; selecting an entry inserts the region code into the field.
    - `Endpoint override` (optional; for S3-compatible endpoints) is visually grouped inside the S3 card.
    - `Use HTTPS`, `Verify TLS certificate`
    - S3 only: `Use virtual-hosted style addressing`
  - Google Drive:
    - `authMode` is fixed to `oauth2Pkce`.
    - `Host` and `Port` are hidden.
    - The password editor is hidden.
    - The `Save password` label changes to `Remember sign-in` and controls whether the refresh token is persisted.
    - New profiles default `extra.rootKind = myDrive`, `extra.useDefaultClientId = true`, and `extra.googleDocsMode = export`.
  - OneDrive Personal / OneDrive Business / SharePoint:
    - `authMode` is fixed to `oauth2Pkce`.
    - The password editor is hidden; sign-in happens in the system browser when the plugin needs a token.
    - OneDrive Personal / OneDrive Business hide host and port.
    - SharePoint uses `host` for the tenant host with optional site path, and may expose a `driveId` field in advanced settings to pin a specific document library.
    - `userName` is a login hint only; it is not a secret.
  - Hidden expert setting: `requireWindowsHello` is not exposed in UI, but can be set in Settings Store JSON per profile.
  - Protocol-specific fields (SFTP/SCP): SSH key paths, known_hosts path
- New connection defaults:
  - When creating a new connection, the host may prefill host/port/path/user from the target plugin's configured defaults (e.g. `defaultHost`, `defaultPort`, `defaultBasePath`, `defaultUser`).
  - FTP: anonymous login is off by default.
- Bottom buttons:
  - `Connect` (OK): saves config, resolves selected/edited connection, and closes dialog
  - `Close`: saves config and closes dialog (no navigation)
  - `Cancel`: closes without saving changes (no navigation)
- `Connect` and `Close` resolve the active profile from the editor's model selection, not only from transient grid-selection chrome. After an in-place rename, the saved selected connection name and the name returned for navigation MUST match the edited model entry even if the visible grid selection was temporarily cleared or rebuilt.
- `FileSystemPluginManager` remains a UI-thread-owned registry. Connection target refresh prepares an immutable browse request on the UI thread by copying the logical plugin id/path, loading a worker-owned module pin, and resolving `RedSalamanderBrowseConnectionTargets` before work is queued. The picker worker invokes only that prepared export and never accesses the manager. Dialog teardown cancels or waits for every owned picker work item before draining posted results and destroying the receiving window.
- System close/title-bar close (`WM_CLOSE`) follows `Cancel` only after confirming when the editor is dirty: choosing OK discards unsaved changes and closes without navigation; choosing Cancel keeps the Connection Manager open.

Connect-time secret behavior:

- If `authMode = password` and `savePassword = true` but no stored password exists yet, the host prompts for the password and saves it to WinCred before closing.
- If `authMode = oauth2Pkce`, Connection Manager does not prompt for a secret. Microsoft Drive launches browser sign-in on demand; Google Drive currently expects a refresh token supplied by host-owned OAuth plumbing or test infrastructure, then stores it according to `savePassword`.

## External Settings Reload Handling

The Connection Manager dialog participates in the settings-editor registry used by `RedSalamander.exe` live settings reload.

Rules:
- The persisted `connections` section on disk is authoritative for non-secret profile data.
- Secrets remain outside the settings JSON (WinCred / session cache) and are not replaced by this reload path.
- Clean dialog/editor state reloads immediately from the newly applied live settings and refreshes its theme.
- Dirty dialog/editor state prompts `Reload from disk` / `Keep editing`.
- Choosing `Keep editing` marks the dialog stale.
- Any later save-producing action (`Connect`, `Close`, explicit save path) on a stale dialog must prompt `Overwrite disk` / `Reload from disk` / `Cancel`.

## FTP / SFTP / SCP / IMAP plugin behavior (initial)

- Navigating to `ftp:` / `sftp:` / `scp:` / `imap:` with no host opens Connection Manager (host-side behavior).
- The FileSystemCurl plugin supports the `/@conn:<connectionName>/...` prefix and uses host services to resolve connection + credentials.

## Google Drive plugin behavior

- Navigating to `gdrive:` with no host opens Connection Manager filtered to Google Drive.
- The built-in Google Drive plugin requires `authMode = oauth2Pkce` and resolves host-owned connection profiles through `/@conn:<connectionName>/...`.
- Google Drive profiles are hostless (`host` empty) and default `extra.rootKind = myDrive`, `extra.useDefaultClientId = true`, and `extra.googleDocsMode = export`.
- Refresh tokens are stored through `IHostConnections::SetConnectionSecret(HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, ...)`, which maps to the `refreshToken` credential suffix in WinCred/session cache.
- The current Google Drive milestone does not own an interactive browser sign-in UI yet; it consumes refresh tokens supplied by host-side OAuth plumbing or tests.

## Microsoft Drive plugin behavior

- Navigating to `onedrive:` / `onedrive-pro:` / `sharepoint:` with no host opens Connection Manager filtered to the matching Microsoft plugin.
- The built-in Microsoft Drive plugin requires `authMode = oauth2Pkce` and uses delegated Microsoft Graph OAuth 2.0 authorization-code flow with PKCE.
- The built-in Microsoft Drive plugin ships with a default plugin-level `clientId` (`90cdea53-7c21-48b0-959e-b4024209027b`). Plugin configuration may override or clear it.
- Refresh tokens are stored through `IHostConnections::SetConnectionSecret(HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, ...)`.
- Connection mapping:
  - OneDrive Personal: `pluginId = builtin/file-system-onedrive-personal`, default `tenantAuthority = consumers`.
  - OneDrive Business: `pluginId = builtin/file-system-onedrive-business`, default `tenantAuthority = organizations`.
  - SharePoint: `pluginId = builtin/file-system-sharepoint`, `host = <tenantHost>[/sitePath]`, default `tenantAuthority = organizations`, optional `extra.driveId` pins a document library.
- All three plugins resolve the selected profile from `/@conn:<connectionName>/...` and use `initialPath` as the starting path inside the resolved drive or library.
