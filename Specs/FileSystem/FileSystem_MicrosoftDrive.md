# Microsoft Drive Virtual File System Plugin

## Overview

RedSalamander ships a built-in virtual file system implementation in `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.dll` that exposes three UI-visible filesystem plugins backed by Microsoft Graph:

- OneDrive Personal
- OneDrive Business
- SharePoint

The plugin uses WinHTTP for HTTPS transport, `yyjson` for JSON parsing, and delegated OAuth 2.0 authorization-code flow with PKCE for user sign-in.

## Plugin identities

| Display name | `PluginMetaData.id` | `PluginMetaData.shortId` |
|-------------|----------------------|--------------------------|
| OneDrive Personal | `builtin/file-system-onedrive-personal` | `onedrive` |
| OneDrive Business | `builtin/file-system-onedrive-business` | `onedrive-pro` |
| SharePoint | `builtin/file-system-sharepoint` | `sharepoint` |

## Navigation

Host-side behavior:

- `onedrive:`
- `onedrive-pro:`
- `sharepoint:`

with no host opens Connection Manager filtered to the matching plugin.

Runtime navigation uses the host-reserved connection prefix:

- `/@conn:<connectionName>/...`

Examples:

- `onedrive:/@conn:Personal Account/`
- `onedrive-pro:/@conn:Work Account/Documents/Reports`
- `sharepoint:/@conn:Team Site/Shared Documents/Q1`

The plugin treats `initialPath` from the selected connection as the starting path inside the resolved drive or document library.

## Transport and authentication

- HTTPS transport: WinHTTP
- API surface: Microsoft Graph `v1.0`
- User agent: `RedSalamander Microsoft Drive/0.1`
- OAuth flow: authorization code + PKCE
- Browser UX: system browser
- Redirect UX: localhost loopback listener on `http://localhost:{ephemeral}/redsalamander/oauth2`

Supported delegated scopes:

- OneDrive Personal: `offline_access Files.ReadWrite User.Read openid profile`
- OneDrive Business: `offline_access Files.ReadWrite User.Read openid profile`
- SharePoint: `offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read openid profile`

Refresh tokens are stored through the host secret API as `refreshToken` when the connection enables `savePassword`. Access tokens are cached in memory only.

### Credential and URL trust boundaries

- A request carrying the Microsoft Graph bearer token must use HTTPS, the default HTTPS port, the exact configured
  Graph origin (`graph.microsoft.com` for the current global-cloud implementation), and the `/v1.0` API path.
  Userinfo, fragments, plaintext schemes, malformed/non-default ports, foreign origins, and paths outside that API
  root are rejected before the request seam or network connection is reached.
- `@odata.nextLink` is treated as an opaque full URL for paging, but it must pass the same Graph-origin validation
  before authorization is attached. A listing tracks already-followed continuation URLs and fails closed before a
  repeated continuation can issue another request.
- Graph-authorized requests disable automatic redirects. Supporting a sovereign Microsoft Graph cloud requires an
  explicit configured-origin policy and matching deterministic tests; it must not be enabled through a broad suffix
  match.
- An upload-session `uploadUrl` is a distinct preauthenticated credential. The current global-cloud policy accepts
  only HTTPS/default-port URLs on `*.up.1drv.com` or `*.sharepoint.com`, rejects userinfo and fragments, disables
  automatic redirects, and never attaches the Graph bearer token. The opaque URL must be sent unchanged after
  validation.
- Diagnostics never serialize request-header blocks or raw request URLs. They emit the method, a redacted target
  class (`graph-api`, `oauth-authority`, `preauthenticated-upload`, or external), status/request ID when available,
  byte counts, and HRESULT. Query values from Graph continuations and upload-session URLs are never logged.
- Transient access/refresh tokens, authorization-header storage, parsed opaque URLs, upload-session response bodies,
  and retained continuation strings are securely cleared where their current in-memory representation permits.

## Plugin configuration

Each logical plugin exposes the same configuration schema through `IInformations::GetConfigurationSchema()`.

Keys:

- `clientId` (string, default `90cdea53-7c21-48b0-959e-b4024209027b`): advanced JSON-only Microsoft Entra application client id override used for PKCE sign-in. It is intentionally hidden from Preferences / Plugin Manager UI and should only be changed in settings JSON when a different app registration is required.
- `connectTimeoutMs` (integer, `1..600000`, default `10000`)
- `requestTimeoutMs` (integer, `1..600000`, default `60000`)
- `pageSize` (integer, `1..999`, default `200`): Graph page size used for `children` listings.
- `uploadChunkMiB` (integer, `1..32`, default `8`): upload-session chunk size in MiB. Chunks are rounded down to Graph's 320 KiB alignment requirement.

## Connection Manager integration

All three plugins are designed to work with host-owned Connection Manager profiles from `Specs/Core/Core_ConnectionManager.md`.

Common requirements:

- `authMode` must be `oauth2Pkce`.
- `userName` is treated as a login hint only.
- `savePassword = true` means `remember sign-in` and persists the refresh token in WinCred.
- `savePassword = false` keeps the refresh token in the host's session cache only.

Shared `extra` keys:

- `tenantAuthority` (string, optional): overrides the Microsoft Entra authority segment used for sign-in.
- `driveId` (string, optional): used by SharePoint to pin a specific document library.

Per-plugin profile mapping:

- OneDrive Personal:
  - `pluginId = builtin/file-system-onedrive-personal`
  - default authority: `consumers`
  - `host` is unused and normally empty
- OneDrive Business:
  - `pluginId = builtin/file-system-onedrive-business`
  - default authority: `organizations`
  - `host` is unused and normally empty
- SharePoint:
  - `pluginId = builtin/file-system-sharepoint`
  - default authority: `organizations`
  - `host` is required and stores the tenant host with an optional site path, for example:
    - `contoso.sharepoint.com`
    - `contoso.sharepoint.com/sites/Team`
  - when `extra.driveId` is omitted, the plugin uses the site's default drive

## Operations

Implemented operations:

- `ReadDirectoryInfo`
- `GetAttributes`
- `GetFileBasicInformation`
- `GetItemProperties`
- `CreateDirectory`
- `CreateFileReader`
- `CreateFileWriter`
- `DeleteItem` / `DeleteItems`
- `MoveItem` / `MoveItems`
- `RenameItem` / `RenameItems`
- `GetDirectorySize`

Behavior notes:

- Directory listings use Graph `children` paging with `$top=pageSize` and `@odata.nextLink`.
  Paging is bounded by the shared page/item/retained-byte/deadline policy and rejects an empty continuation when
  the provider says more data exists or any continuation URL already followed in the same operation.
- Metadata requests use Graph item lookup with a small `$select` set.
- `GetItemProperties` returns structured properties for both files and folders using all Microsoft Graph item metadata the plugin has available, including general identity/path/type data, drive/remote identifiers, timestamps, facets, hashes, and size for files. Missing Graph fields are omitted rather than reported as placeholder values.
- File reads resolve `@microsoft.graph.downloadUrl` and then use ranged HTTP reads against that download URL.
- Writes stage data into a local temporary file first.
  - Up to 250 MiB: simple upload (`/content`)
  - Above 250 MiB: upload session with chunked PUTs
- Every upload-session `202 Accepted` response must contain valid `nextExpectedRanges`. The writer resumes from
  the lowest server-acknowledged missing offset, rejects non-progress/out-of-range/contradictory ranges, and
  accepts completion only through a final `200`/`201` after all declared bytes are acknowledged.
- Same-drive move and rename use Graph `PATCH`.
- Moving a folder onto an existing folder moves children individually but retains the drained source folder and
  reports partial cleanup. Graph item deletion is recursive and the plugin has no atomic "delete only if still
  empty" primitive, so deleting the folder after a re-list would still risk erasing a concurrently-added child.
- Overwrite move results distinguish the committed primary mutation from backup cleanup and rollback. Once the
  requested move has committed, failure to delete the recoverable overwrite backup is logged as cleanup debt and
  does not turn the move into an ordinary failure that callers might destructively retry.
- Cross-drive move returns `ERROR_NOT_SUPPORTED`.
- Server-side copy is not implemented in this version. `CopyItem` and `CopyItems` return `ERROR_NOT_SUPPORTED`, allowing the host bridge to fall back to read/write copy paths where applicable.

## Drive resolution

- OneDrive Personal / OneDrive Business resolve the active drive with `GET /me/drive`.
- SharePoint resolves:
  1. the site from `host`
  2. the drive from `extra.driveId` if present, otherwise the site's default drive

The plugin caches resolved drive metadata per connection name for the current app run.

## Limitations

- Delegated user auth only. App-only / client-credentials auth is out of scope.
- Microsoft sovereign-cloud Graph and upload origins are not currently configured; the credential boundary supports
  only the documented global-cloud origins above.
- The built-in default `clientId` is intended for RedSalamander's shipped Microsoft app registration. Installations that need a different Microsoft Entra app must override it in plugin configuration.
- SharePoint browsing is rooted to the configured site (or tenant root if only the tenant host is provided). There is no tenant-wide search UI.
- `SetFileBasicInformation` is not implemented.
- `IFileSystemSearch` is not implemented.

## Test coverage

Optional selftest coverage is wired through Connection Manager profiles:

- `--fileops-selftest`
  - secret retrieval and sandbox validation for OneDrive Personal, OneDrive Business, and SharePoint
- `--compare-selftest`
  - remote compare smoke
  - remote directory-size callback contract

Profile names and environment-variable overrides are documented in `Specs/Testing/Testing_SelfTestRemoteCredentials.md`.

Credential-boundary coverage is deterministic and does not require live credentials:

- Debug `PluginContractTests` runs the Microsoft Drive exported selftests for strict URL validation, approved
  same-origin pagination, foreign/plaintext continuation rejection before a second request, repeated-link rejection
  before a third request, server-acknowledged upload offsets, concurrent-child-safe merge cleanup, committed-move
  cleanup debt, and injected Graph/upload transport failures whose captured diagnostics contain no bearer or
  opaque-query sentinels.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` preserves the validated URL types, redirect suppression,
  separated/secure-cleared authorization header, continuation guard, and raw-URL/header logging bans.
- Closeout evidence is archived under
  `Specs/TestRuns/4cb089111a23/FileSystemMicrosoftDrive/2026-07-16_2046_observatory_track3_credential_boundary/`.
