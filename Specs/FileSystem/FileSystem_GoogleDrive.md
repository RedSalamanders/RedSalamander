# Google Drive Virtual File System Plugin

## Overview

RedSalamander ships a built-in Google Drive virtual file system implementation in `Plugins/FileSystemGoogleDrive/FileSystemGoogleDrive.dll`.

The provider implements mandatory path-scoped capability v2 and, since R0f-GDrive (2026-09-03),
is a full File Operations destination: copy, move (native parent change), rename, delete, recycle
(trash), create directory, properties, read, and write are all advertised and implemented. A
future `md5Checksum` verification claim is valid only when its algorithm, byte scope, and
revision binding are all declared and contract-tested; until then `verification.providerProof`
stays `none`.

The route is `providerWatchdog` with a provider-owned bound (see "Containment" below). A
read-only connection profile (`extra.readOnly = true`) keeps its route read-only: every mutation
capability is `false` for its paths and every mutation returns `ERROR_ACCESS_DENIED`.

- Long plugin ID: `builtin/file-system-gdrive`
- Short ID / path prefix: `gdrive`
- Transport stack: `curl` + `yyjson`
- Dependency change relative to the existing curl-based plugins: `curl[http2]`

Scope:

- multiple Google Drive connection profiles are supported through Connection Manager,
- each profile may target a different Google identity / refresh token / Drive root,
- directory enumeration and drive-info reporting are implemented,
- file reads (ranged download), uploads (resumable), server-side copies, native moves, renames,
  permanent deletes, trash (recycle), folder creation, and timestamp updates are implemented
  through `IFileSystem`, `IFileSystemIO`, `IFileSystemDirectoryOperations`, and
  `IFileSystemAtomicWriter`.

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

- sibling-name comparison and opaque Drive IDs are ordinal and case-sensitive,
- if a folder has a unique name that cannot be mistaken for a synthetic decoration, the visible name is the raw
  Drive name,
- if multiple siblings have the exact same name, the plugin appends a synthetic suffix:
  - `<name> [id:<percent-encoded-driveFileId>]`
- a literal raw name already shaped like `[id:...]` receives another identity suffix, so it cannot alias a
  duplicate item's generated display name,
- path resolution regenerates the exact exposed names from the current listing instead of case-folding or treating
  every suffix-like literal as syntax. Ambiguous exposed names return `ERROR_DUP_NAME`.

## Plugin Configuration

`plugins.configurationByPluginId["builtin/file-system-gdrive"]` is a plugin-defined JSON object with these keys:

- `defaultClientId` (string, default `""`): desktop OAuth client ID used when a connection profile opts into the plugin default.
- `connectTimeoutMs` (integer, default `10000`, range `1..600000`): TCP connect timeout in milliseconds.
- `requestTimeoutMs` (integer, default `30000`, range `1..600000`): hard total libcurl request deadline and
  no-progress timeout in milliseconds. An authorized logical GET has a separately bounded retry window.
- `pageSize` (integer, default `200`, range `1..1000`): `files.list` page size.

Configuration updates are transactional: malformed or non-object JSON returns `ERROR_INVALID_DATA`, unknown
members round-trip, and a rejected candidate leaves the active settings and token cache unchanged.

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
  - stored for the future document-export behavior; native Google documents (`application/vnd.google-apps.*`
    other than folders and shortcuts) have no byte stream of their own, so `CreateFileReader` on one returns
    `ERROR_NOT_SUPPORTED` and they are never copied out
- `readOnly` (bool, default `false`)
  - honored: the profile's paths advertise no mutation capability and every mutation returns `ERROR_ACCESS_DENIED`
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

Refresh-token and access-token copies are ordinary strings only for the shortest operation/cache lifetime that
requires them. Superseded values, temporary UTF-8 conversions, resolved-connection values, and expired/evicted
cache entries are securely cleared. Neither token is emitted through plugin configuration.

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

API usage:

- OAuth token exchange: `POST https://oauth2.googleapis.com/token`
- directory listing: `GET https://www.googleapis.com/drive/v3/files`
- metadata by id: `GET https://www.googleapis.com/drive/v3/files/<id>` (`fields=id,name,mimeType,modifiedTime,size,trashed,version`)
- content: `GET https://www.googleapis.com/drive/v3/files/<id>?alt=media` with a `Range` header (4 MiB chunks)
- folder creation: `POST https://www.googleapis.com/drive/v3/files` (`mimeType = application/vnd.google-apps.folder`)
- upload: `POST https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable` (new object) or
  `PATCH .../upload/drive/v3/files/<id>?uploadType=resumable` (replace content), then `PUT` chunks of 8 MiB
  with `Content-Range` to the session URL (`308` per intermediate chunk, the resource on the last one; an empty
  object completes with `Content-Range: bytes */0`)
- rename / move / trash / timestamps: `PATCH https://www.googleapis.com/drive/v3/files/<id>` with `name`,
  `addParents`/`removeParents`, `trashed`, or `modifiedTime`
- permanent delete: `DELETE https://www.googleapis.com/drive/v3/files/<id>` (a folder deletes its subtree)
- copy: `POST https://www.googleapis.com/drive/v3/files/<id>/copy` (files; folders are recreated and copied per child)
- about / quota: `GET https://www.googleapis.com/drive/v3/about`
- shared-drive label lookup: `GET https://www.googleapis.com/drive/v3/drives/<id>`

Every request carries `supportsAllDrives=true`. The implementation always requests partial responses
(`fields=...`) and uses paging with `pageToken`.

Transport and paging are fail-closed:

- JSON response bodies are capped at 16 MiB before append,
- each request has a hard total deadline in addition to the low-speed guard,
- read-only authorized requests retry `429` and `5xx` at most three times with capped
  `Retry-After`/exponential delay; each attempt is bounded by the connect timeout and the hard request timeout,
  and the retry sequence as a whole by a separate logical-operation window,
- mutation requests never retry a `429` or `5xx`: the remote service may already have committed the POST,
  PATCH, PUT, or DELETE, so the provider returns the failure with an indeterminate mutation outcome instead of
  risking a duplicate commit,
- one `401` may invalidate only the token that actually failed and trigger one refresh; this is permitted for
  every verb because the authorization rejection means the request was not accepted for execution,
- concurrent callers for the same profile/client share one refresh; a failed refresh leaves the prior cache entry
  intact but unusable until policy permits another refresh,
- directory paging uses the shared page/item/retained-byte/deadline guard and rejects repeated tokens before the
  next request.

Google Drive and FileSystemCurl share the process-wide libcurl lifetime through
`Common::CurlRuntime::ProcessLease`. Google Drive stops new instances, releases all instances/curl handles, and
releases its lease at the explicit plugin quiet point. It never performs isolated DLL-local global cleanup while
FileSystemCurl can still issue work.

## Containment (R0f-GDrive)

Every mutation entry point (`CopyItem`, `MoveItem`, `DeleteItem`, `RenameItem`, and the four batch
variants) keeps the owning call's `FileSystemOptions` on the current thread
(`DriveOperationOptionsScope`). Every libcurl transfer issued under it polls the host's operation
control from its progress callback (libcurl invokes it at least once per second even while a request
waits for the first byte) and returns the host's own verdict (`ERROR_CANCELLED` / `ERROR_TIMEOUT`);
throttle backoffs are quiet points polled every 100 ms. Every attempt is also bounded by the connect
timeout and the hard request timeout, so a silent server returns on its own.

Route facts: `routeClass = providerWatchdog`, `cancellation.deadline = true`,
`providerWatchdogTimeoutMs = DriveProviderWatchdogTimeoutMs(connectTimeoutMs, requestTimeoutMs)` =
2 x (connect + request) + 1 s (one token refresh plus one request; 81 s with the default timeouts).
Cancel on a streaming or silent request returns within a couple of progress ticks (measured below
3 s in the plugin witness). Readers and writers used by the host bridge outside a plugin entry point
carry no operation control on the thread; their transfers rely on the same timeouts.

Every item completion carries a `FileSystemItemMutationResult` receipt: success and definitive
refusals (not found, access denied, directory not empty, name conflicts, invalid names) are known
outcomes; a transport failure, cancel, or deadline leaves the outcome unknown (no receipt) and the
host reports it as indeterminate.

`C0_GDriveCommittedCopyResponseLost` drives the real host/provider Copy route against
the loopback Drive fixture. The backend publishes a file before returning 503 and
counts requests independently from committed objects. It requires exactly one of each,
the source and destination to exist, no returned-error conflict prompt, an immutable
indeterminate/unknown-publication result, and the localized unknown-outcome task summary.
This complements the HTTP retry-policy unit witness; an error returned before any
backend change does not cover committed-response loss.
`C0_GDriveCommittedDeleteResponseLost` applies the same witness to provider-native
Permanent Delete: one DELETE request and one backend removal, no returned-error
decision, and an indeterminate result with source disposition Unknown. A later 404
must not be used to turn the lost response into a second mutation attempt.

## Capabilities

`GetPathCapabilities()` and the typed route facts advertise, for a writable profile:

- `copy = true` (server-side `files.copy`; folders are recreated and copied per child by an iterative frame
  stack with no arbitrary depth cutoff)
- `move = true`, `nativeMove = true` (one `PATCH` re-homes a file or a whole folder; a folder onto an
  existing folder merges per child)
- `delete = true` (permanent; a folder without `FILESYSTEM_FLAG_RECURSIVE` must be empty, otherwise
  `ERROR_DIR_NOT_EMPTY`)
- `recycle = true` (`trashed = true`; recoverable from Drive's trash)
- `rename = true` (case-only renames are supported: names compare ordinal case-sensitive)
- `createDirectory = true`, `properties = true`, `read = true`, `write = true`
- recursive directory sizing uses an explicit pending-folder stack and returns exact descendant counts and
  bytes; it never reports `S_OK` after silently truncating a deep tree
- concurrency: copy/move 4, delete 4, recycle 1
- `identity.revision = version`: the ranged reader re-reads the object's `version` before every chunk and
  fails with `ERROR_REVISION_MISMATCH` if the file changed mid-read
- `publication.committedSize = true` and `IFileSystemAtomicWriter` = supported: a writer stages bytes in a
  delete-on-close temporary file and publishes them with one resumable upload at `Commit`; the name appears
  only when Drive acknowledges the last chunk, and the committed size is proven from the returned resource
- `names.pathTextStableIdentity = true`: exposed names are unique per folder (duplicates carry the
  `[id:...]` decoration) and an ambiguous text fails closed with `ERROR_DUP_NAME`

Conflict rules:

- a destination file that already exists is refused with `ERROR_ALREADY_EXISTS` unless overwrite was granted;
  with overwrite the new object is created first and the replaced file is then trashed (never permanently
  deleted by an overwrite),
- a destination folder is never replaced by a file; a folder onto an existing folder merges per child, skipping
  colliding files without overwrite, and the item ends as `ERROR_PARTIAL_COPY` with every skipped source preserved,
- the writer replaces only the object that occupied the name when it opened, at the version the user saw
  (R3-1, `IFileWriterExpectedReplacement`): a different or changed occupant is `ERROR_REVISION_MISMATCH`, an
  object that appeared later is `ERROR_ALREADY_EXISTS`, and nothing is published in either case; an
  expectation or an occupant without a last-write time carries no identity to compare and is refused with
  `ERROR_REVISION_MISMATCH` before any upload (R0-RC4),
- new names must be non-empty, at most 255 UTF-16 units, contain no `/`, and must not look like the `[id:...]`
  decoration (`ERROR_INVALID_NAME`).

## Current Limitations

- No built-in browser / PKCE authorization UX yet.
- No Google Docs export/open pipeline yet (native documents cannot be read as bytes).
- Shortcuts are surfaced as entries and copied/moved as objects; dereference/open semantics are not implemented.
- Folder-onto-folder merges skip colliding children without a per-child conflict prompt (`ERROR_PARTIAL_COPY`).
- Content proof (R3-2) uses Drive's `sha256Checksum` of the published resource (`IFileWriterContentProof`,
  `verification.providerProof: "writer-digest"`): the host hashes the streamed bytes and compares the digest Drive
  returns when the resumable upload completes, so Verified Copy needs no read-back and Managed Move into Drive
  deletes its source only after the proof matches. Native documents carry no checksum and stay `Unavailable`.

## Deterministic test coverage

Debug `PluginContractTests` runs the Google Drive exported selftests. They cover response caps, trickle/deadline
abort, eight concurrent callers sharing one refresh, bounded `429` retry, repeated page-token rejection before a
third request, case-distinct opaque IDs, case-distinct names, literal suffix-like name round-trips, transactional
configuration rollback, and eight alternating Curl/Google physical-unload cycles in which the surviving plugin
still creates a curl easy handle. These tests use an injected local transport and require no Google credentials.

R0f-GDrive adds a deterministic loopback Drive fixture (`FileSystemGoogleDrive.SelfTest.FakeDrive.cpp`,
test-enabled builds only): the plugin's Drive and token endpoints are redirected to it through
`FileSystemGoogleDriveSelfTest::ConfigureFakeDrive` and `/@conn:google-drive-selftest/...` resolves to a
synthetic connection. The plugin witness (`PluginContractTests --google-drive-r0f-containment-selftests`,
also part of the exported debug self-tests) proves that Cancel returns a streaming request and a silent
request within a few progress ticks with `ERROR_CANCELLED` and deletes nothing, and that an un-canceled
silent request returns on its own within the declared bound. The host self-test
`R0fGDrive_FakeDriveReadWriteCreateMoveDelete` (family `FileOpsFamily_R0fGDriveContainment`) drives the
File Operations matrix through the fixture: cross-provider copy both ways (byte-exact), provider
`CreateDirectory`, `RenameItem`, native `MoveItem`, and a permanent Delete task, with the route qualified
as `providerWatchdog` with a nonzero bound.

## Identity-pinned Permanent Delete (C10)

`FileSystemGoogleDrive` implements `IFileSystemIdentityDelete`. `ResolveDeleteIdentity` resolves
the item the path names and returns its Drive file id (folders included). `DeleteIfIdentity`
resolves the name again and refuses with `ERROR_REVISION_MISMATCH`, deleting nothing, when the name
now refers to a different file id; otherwise it deletes that id permanently as `DeleteItem` does.
The host pins every selected root this way while a Permanent Delete prepares. Witness: the last
phase of `R0fGDrive_FakeDriveReadWriteCreateMoveDelete` resolves the identity, deletes and recreates
the file through the plugin (a new file id under the same name), and requires the refusal with the
new file intact.
