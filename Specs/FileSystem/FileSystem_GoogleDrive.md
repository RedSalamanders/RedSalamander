# Google Drive Virtual File System Plugin

## Overview

RedSalamander ships a built-in Google Drive virtual file system implementation in `Plugins/FileSystemGoogleDrive/FileSystemGoogleDrive.dll`.

- Long plugin ID: `builtin/file-system-gdrive`
- Short ID / path prefix: `gdrive`
- Transport stack: `curl` + `yyjson`
- Dependency change relative to the existing curl-based plugins: `curl[http2]`

The current milestone is intentionally narrow:

- multiple Google Drive connection profiles are supported through Connection Manager,
- each profile may target a different Google identity / refresh token / Drive root,
- directory enumeration and drive-info reporting are implemented,
- file reads, uploads, deletes, renames, and moves are not implemented yet.

## Paths

### Root

- `gdrive:/`

Behavior:

- returns an empty directory listing,
- exposes NavigationView menu items for Google Drive,
- exposes generic drive info (`displayName = "gdrive:/"`, volume label = `Google Drive`).

### Connection-scoped paths

Google Drive content is reached through Connection Manager profiles:

- `gdrive:/@conn:<connectionName>/`
- `gdrive:/@conn:<connectionName>/<path>/...`
- `gdrive://@conn/<connectionName>/...` (shorthand authority form accepted by the host and normalized before the plugin sees it)

Each connection name resolves to a `ConnectionProfile` owned by the host. This allows multiple Google accounts or multiple roots for the same Google account without embedding credentials in the URI.

### Duplicate sibling names

Google Drive permits duplicate names in the same folder. The plugin resolves duplicates as follows:

- if a folder has a unique name, the visible name is the raw Drive name,
- if multiple siblings share the same visible name, the plugin appends a synthetic suffix:
  - `<name>[id:<driveFileId>]`
- navigating back into an ambiguous item without the synthetic suffix returns `ERROR_DUP_NAME`.

## Plugin Configuration

`plugins.configurationByPluginId["builtin/file-system-gdrive"]` is a plugin-defined JSON object with these keys:

- `defaultClientId` (string, default `""`): desktop OAuth client ID used when a connection profile opts into the plugin default.
- `connectTimeoutMs` (integer, default `10000`, range `1..600000`): TCP connect timeout in milliseconds.
- `requestTimeoutMs` (integer, default `30000`, range `1..600000`): total libcurl request timeout in milliseconds.
- `pageSize` (integer, default `200`, range `1..1000`): `files.list` page size.

## Connection Manager Contract

Google Drive profiles use Connection Manager with these rules:

- `pluginId` must be `builtin/file-system-gdrive`
- `authMode` must be `oauth2Pkce`
- `host` is intentionally empty and may be omitted from persisted JSON
- `initialPath` is a plugin path and typically `/`

### Profile `extra` fields

Non-secret Google Drive settings are stored in `ConnectionProfile.extra`:

- `rootKind` (string, default `myDrive`)
  - `myDrive`
  - `sharedDrive`
- `sharedDriveId` (string, required when `rootKind = sharedDrive`)
- `googleDocsMode` (string, default `export`)
  - current milestone stores the value for future document-export behavior but does not yet expose file IO
- `readOnly` (bool, default `false`)
  - current milestone is effectively read-only regardless of this flag
- `useDefaultClientId` (bool, default `true`)
- `clientId` (string)
  - used only when `useDefaultClientId = false`

### Secrets

The current implementation does not start an interactive browser sign-in flow.

Instead, the plugin requires a refresh token supplied by the host under secret kind `refreshToken`:

- persisted refresh token: stored in WinCred
- session-only refresh token: stored in the host session cache

If no refresh token is available, connection-scoped enumeration fails with `ERROR_NOT_FOUND`.

If no effective OAuth client ID is available, connection-scoped enumeration fails with `ERROR_NOT_SUPPORTED`.

## Navigation Menu

`INavigationMenu` returns these entries:

- header: `Google Drive`
- separator
- command: `Open Connection...`
- path item: `/`

Selecting `Open Connection...` asks the host to open Connection Manager filtered to `builtin/file-system-gdrive`. If the user chooses a profile, the plugin requests navigation to:

- `/@conn:<connectionName>/`

## Drive Info

### Root (`gdrive:/`)

- display name: `gdrive:/`
- volume label: `Google Drive`
- file system: `Google Drive`

### Connection-scoped paths

For a resolved Google Drive connection, `IDriveInfo` reports:

- display name: `gdrive://<connectionName>` or `gdrive://<connectionName><path>`
- volume label:
  - shared drive name when `rootKind = sharedDrive`
  - `My Drive` otherwise
- file system: `Google Drive`
- storage quota:
  - available for My Drive when Google returns `storageQuota`
  - omitted for shared drives in the current milestone

## HTTP / API Shape

Current API usage is intentionally minimal:

- OAuth token exchange: `POST https://oauth2.googleapis.com/token`
- directory listing: `GET https://www.googleapis.com/drive/v3/files`
- about / quota: `GET https://www.googleapis.com/drive/v3/about`
- shared-drive label lookup: `GET https://www.googleapis.com/drive/v3/drives/<id>`

The implementation always requests partial responses (`fields=...`) and uses paging with `pageToken`.

## Capabilities

The current `GetCapabilities()` payload reports all mutating operations as unsupported:

- `copy = false`
- `move = false`
- `delete = false`
- `rename = false`
- `properties = false`
- `read = false`
- `write = false`

This matches the current milestone: enumeration and drive metadata exist, but no `IFileSystemIO` / file-transfer surface is exposed yet.

## Current Limitations

- No built-in browser / PKCE authorization UX yet.
- No file download stream or upload implementation yet.
- No Google Docs export/open pipeline yet.
- Shortcuts are surfaced as entries, but shortcut dereference/open semantics are not implemented yet.
- Request timeout currently uses libcurl total timeout semantics rather than low-speed timeout semantics.
