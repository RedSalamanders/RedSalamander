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
| OneDrive Personal | `builtin/file-system-onedrive-personal` | `onedrivep` |
| OneDrive Business | `builtin/file-system-onedrive-business` | `onedriveb` |
| SharePoint | `builtin/file-system-sharepoint` | `sharepoint` |

## Navigation

Host-side behavior:

- `onedrivep:`
- `onedriveb:`
- `sharepoint:`

with no host opens Connection Manager filtered to the matching plugin.

Runtime navigation uses the host-reserved connection prefix:

- `/@conn:<connectionName>/...`

Examples:

- `onedrivep:/@conn:Personal Account/`
- `onedriveb:/@conn:Work Account/Documents/Reports`
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

## Plugin configuration

Each logical plugin exposes the same configuration schema through `IInformations::GetConfigurationSchema()`.

Keys:

- `clientId` (string, default `90cdea53-7c21-48b0-959e-b4024209027b`): Microsoft Entra application client id used for PKCE sign-in. Override it only when you want to use a different app registration.
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
- Metadata requests use Graph item lookup with a small `$select` set.
- File reads resolve `@microsoft.graph.downloadUrl` and then use ranged HTTP reads against that download URL.
- Writes stage data into a local temporary file first.
  - Up to 250 MiB: simple upload (`/content`)
  - Above 250 MiB: upload session with chunked PUTs
- Same-drive move and rename use Graph `PATCH`.
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
