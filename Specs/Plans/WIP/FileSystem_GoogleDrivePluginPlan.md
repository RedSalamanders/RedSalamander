# Google Drive Filesystem Plugin Plan

Last updated: 2026-07-02

Status: WIP - read-only directory-listing milestone landed; file IO, write operations, and remote validation remain.

> **2026-07-02 folder review:** ~1/3 done (enumeration, drive info, `[id:...]` duplicate-name suffixes, token refresh — landed 123ce33a6). Everything mutating/streaming still stubbed: capabilities JSON reports read/write/copy/move/delete/rename ALL false; Copy/Move/Delete/Rename entry points return `ERROR_NOT_SUPPORTED` (`FileSystemGoogleDrive.cpp:1247-1323`). No Drive feature commit since 2026-04-25. Next action = Phase 2 remainder: interactive OAuth PKCE sign-in + read-stream/download support (without in-product sign-in nothing else is usable).

## Closeout audit (`2026-04-25`)

Current code and spec evidence show the plugin project, OAuth/Connection Manager groundwork, directory enumeration, Drive metadata, and the authoritative spec in `Specs/FileSystem/FileSystem_GoogleDrive.md` are present. The plan cannot move to Done because the current plugin still reports mutating capabilities as unsupported, and `CopyItems`, `MoveItems`, `DeleteItems`, file streams, and upload/export paths return `ERROR_NOT_SUPPORTED`.

Remaining closeout checklist:

- [x] Plugin project and solution wiring exist under `Plugins/FileSystemGoogleDrive/`.
- [x] OAuth2/PKCE connection profile groundwork and refresh-token secret retrieval exist. **Caveat (2026-07-02 folder review):** groundwork only — `OAuth2Pkce` auth mode, `HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN`, and `SetConnectionSecret`/`DeleteConnectionSecret` all exist in `HostServices.cpp`, but the interactive PKCE sign-in flow (browser launch + 127.0.0.1 loopback listener + initial token exchange) exists nowhere for Google (grep pkce/code_verifier/loopback = 0 hits; contrast `FileSystemMicrosoftDrive.cpp:2454` `LaunchInteractiveAuth`). A profile cannot acquire its first refresh token in-product; the plugin only consumes a pre-stored refresh token (`FileSystemGoogleDrive.cpp:1614-1712`).
- [x] Directory enumeration and drive-info milestone are documented in `Specs/FileSystem/FileSystem_GoogleDrive.md`.
- [ ] Implement or explicitly defer file download/read streams, including native Google Docs export behavior.
- [ ] Implement or explicitly defer upload/overwrite/create-directory/rename/move/delete/server-side copy.
- [ ] Finish shortcut and duplicate-name path semantics for the host's segment-based virtual filesystem contract.
- [ ] Add deterministic local selftests plus gated remote Drive validation and archive the evidence.

## Goal

Add a new remote filesystem plugin, `FileSystemGoogleDrive`, that:

- fits the existing RedSalamander virtual filesystem model (`/@conn:<name>/...`)
- supports multiple Google connection profiles with separate identities and credentials
- stays dependency-light and Windows-native
- is fast enough for large listings and high-latency WAN links without pulling in a heavy SDK stack

## Recommended Shape

Settled and shipped: dedicated plugin DLL under `Plugins/FileSystemGoogleDrive/` reusing `curl` (HTTPS), `yyjson` (JSON), WIL, the existing `IHostConnections`/Connection Manager model, and the bounded-worker/serialized-callback patterns from `FileSystemCurl`/`FileSystemS3`.

## Why This Approach

Settled: Drive is plain REST/HTTPS with no first-party C++ client, and the real performance wins (fewer round trips, trimmed `fields`, server-side copy, resumable uploads, bounded parallelism with cancellation) come from request shape rather than another client library.

## Dependency Decision

Settled — `curl` + `yyjson` shipped (plus WIL/Win32 already in repo; `curl[http2]` remains an optional low-risk upgrade); `google-cloud-cpp` rejected (wrong abstraction level, heavy DLL surface, poor fit for repo plugin patterns) and `cpprestsdk` rejected (duplicates `curl`+`yyjson` with no performance win).

## Product Constraints From Google Drive

These should be treated as first-order design inputs, not implementation details:

- Full Google Drive filesystem access requires restricted Drive scopes. Shipping a built-in client ID is a product/compliance decision.
- Shared drives require explicit API flags and drive-scoped queries.
- Push notifications require webhook endpoints, which is a poor fit for a desktop plugin. Directory watch should not be a v1 goal.
- Native Google Workspace documents are not regular byte-addressable files and require export semantics on download.
- Drive allows duplicate sibling names, while the current host/plugin contract is path-segment based.

## Recommended v1 Scope

### In scope

- Multiple Google Drive connection profiles via Connection Manager
- Separate credentials per profile
- My Drive root
- One configured Shared Drive root per profile
- Directory enumeration
- Binary file download / upload / overwrite
- Rename
- Move within the same account/root
- Server-side file copy where Drive supports it
- Delete
- Drive quota / account info for `IDriveInfo`
- Read-only mode per profile

### Deferred from v1

- Real-time directory watch
- Offline sync
- Service-account auth
- Domain admin features
- Full fidelity editing of native Google Docs/Sheets/Slides
- Cross-account server-side copy

## Required Host-Side Changes

2026-07-02 folder review: items 1-3 below have landed (see Phase 1 in Phased Delivery); kept here as the authoritative description of the shipped shape.

The current Connection Manager and host secret API are password/SSH oriented. That is not enough for OAuth refresh-token storage.

### 1. Auth model extension

Update `Common/SettingsStore.h` and related host code to add:

- `ConnectionAuthMode::OAuth2Pkce`

Recommended behavior:

- keep `savePassword` semantics internally for now, but expose OAuth profiles in UI as `Remember sign-in`
- keep Windows Hello gating for persisted refresh tokens

### 2. Secret kind extension

Update `Common/PlugInterfaces/Host.h` to add:

- `HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN`

### 3. Host secret persistence API

Add a generic host API so a plugin can persist or delete a secret after completing OAuth in the browser.

Recommended shape:

- `SetConnectionSecret(...)`
- `DeleteConnectionSecret(...)`

Without this, the plugin can obtain an auth code and refresh token but cannot store it safely for future sessions.

### 4. Connection Manager UI / schema

Update the Connection Manager so `Google Drive` profiles can persist non-secret configuration in `ConnectionProfile.extra`, including:

- `clientId`
- `useDefaultClientId`
- `rootKind` = `myDrive` or `sharedDrive`
- `sharedDriveId`
- `readOnly`
- `googleDocsMode`

The OAuth browser flow does not need to live inside the dialog in v1. The plugin can trigger it lazily on first use if the refresh token is missing or revoked.

## Recommended Profile Model

Each Connection Manager profile is one logical Google Drive endpoint:

- one refresh token
- one selected account identity
- one logical root (`My Drive` or a specific shared drive)
- one policy set (`readOnly`, optional custom client ID)

This naturally supports:

- multiple Google accounts
- multiple profiles against the same Google account
- side-by-side personal / shared drive roots

Recommended identity behavior:

- after first successful auth, fetch `about.get` and persist the account email into `userName`
- keep the refresh token as the only persisted secret
- never persist access tokens

## Recommended Plugin Architecture

Create `Plugins/FileSystemGoogleDrive/` with these major components:

### `GoogleDrivePlugin`

- `IFileSystem`
- `IFileSystemIO`
- optional `IDriveInfo`
- no `IFileSystemDirectoryWatch` in v1

### `GoogleConnectionContext`

- resolved connection profile
- refresh token access
- access-token cache
- per-connection HTTP handle pool
- per-connection rate limiting / backoff state

### `GoogleOAuthPkce`

- PKCE verifier/challenge generation
- browser launch via Win32 shell APIs
- loopback callback listener on `127.0.0.1:<ephemeral>`
- token exchange / refresh

Use a tiny local loopback listener implemented with WinSock, not a new web framework or embedded browser.

### `GoogleDriveHttpClient`

- authenticated REST requests
- retry policy for `401`, `429`, and retryable `5xx`
- `fields` parameter helpers
- pagination helpers

### `GoogleDrivePathResolver`

- path segment -> Drive file ID resolution
- per-parent child cache
- duplicate-name disambiguation
- shortcut resolution

### `GoogleDriveTransfers`

- streaming downloads
- multipart upload for small files
- resumable upload for large files
- server-side copy / move helpers

## Path Model And Duplicate Names

This is the hardest contract issue. The host is path-based; Drive is ID-based and allows duplicate names.

### Recommended v1 contract

Keep the host ABI unchanged and encode ambiguity only when needed.

- Unique sibling name: expose the plain display name.
- Duplicate sibling name: expose a stable synthetic suffix, for example `name [id:1A2B3C4D]`.
- When resolving a path, parse an explicit ID suffix first. If there is no suffix, resolve by child name under the parent ID.

Why this is the right v1 compromise:

- no `FileInfo` ABI change
- no new opaque-path host contract
- normal folders stay readable
- duplicate-name folders remain navigable and operable

Required implementation rule:

- the synthetic suffix is a transport detail, not part of the stored Drive file name

## Native Google Docs Handling

Drive-native documents are not normal byte streams. Treat them explicitly.

### Recommended v1 behavior

- List them in directory views.
- Allow metadata operations that Drive supports: rename, move, delete.
- Allow download-to-local through Drive export APIs using a deterministic default export format per MIME type.
- Do not support overwrite-in-place back into a native Google document in v1.

Recommended default export mapping:

- Docs -> `.docx`
- Sheets -> `.xlsx`
- Slides -> `.pptx`
- Drawings -> `.png`

If product wants a stricter first milestone, a valid fallback is:

- list native Google docs as read-only items
- skip overwrite/copy-back support entirely

## Shared Drives And Shortcuts

### Shared drives

Per profile, support either:

- My Drive root
- one specific Shared Drive root

Always set the Drive flags needed for shared drive enumeration and operations.

### Shortcuts

Recommended v1 behavior:

- resolve shortcuts to folders for navigation
- resolve shortcuts to files for read/download
- mark unresolved shortcuts as read-only items and log once

## Performance Plan

### Directory listing

- use `files.list` with aggressive `fields` trimming
- use the largest practical page size
- avoid N+1 metadata fetches
- cache only small parent/child lookup tables, not full trees

### Read path

- stream downloads directly to `IFileReader`
- export native Google docs only on demand
- avoid temp files unless an interface boundary requires them

### Write path

- use multipart upload for small files
- use resumable upload by default for large files and overwrite paths
- refresh access tokens lazily and only once per failure burst

### Intra-Drive operations

- rename: metadata update only
- move: parent update only
- file copy within the same account/root: server-side `files.copy`
- recursive folder copy: host/plugin-managed traversal using server-side file copy for leaf files

### Parallelism

Stay aligned with repo norms:

- default copy/move concurrency: 4
- default delete concurrency: 8
- serialized host callbacks
- cancellation checked at every request boundary and upload/download progress checkpoint

### Optional transport upgrade

If profiling shows metadata latency dominates, add a small internal `curl_multi` dispatcher later. Do not make that a prerequisite for v1.

## Error Handling Rules

Follow existing repo policy:

- Win32 failures: `Debug::ErrorWithLastError(...)`
- unexpected API failures: `Debug::Error(...)`
- retryable remote failures: `Debug::Warning(...)` only when action is taken
- do not log normal control flow such as cancellation, zero-byte folders, or token refresh success

Handle these API cases explicitly:

- `401` / invalid token -> single refresh path, then retry once
- `403` / insufficient scope -> fail clearly, do not spin
- `429` / rate limit -> exponential backoff with jitter
- `404` on stale ID -> invalidate cache and re-resolve once

## Code Touch Points

### New

- `Plugins/FileSystemGoogleDrive/`
- project file and solution wiring
- plugin spec: `Specs/FileSystem/FileSystem_GoogleDrive.md`

### Modify

- `vcpkg.json`
- `Common/SettingsStore.h`
- `Common/PlugInterfaces/Host.h`
- `RedSalamander/ConnectionManagerDialog.cpp`
- `RedSalamander/ConnectionProfileUtils.h`
- `RedSalamander/ConnectionProfileUtils.cpp`
- `RedSalamander/HostServices.cpp`
- `Specs/Core/Core_ConnectionManager.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md`

## Test Plan

### Pure local tests

Add deterministic tests around logic that does not require Google access:

- PKCE code verifier / challenge generation
- duplicate-name path suffix formatting and parsing
- Drive JSON -> `FileInfo` mapping
- shortcut resolution rules
- export MIME selection rules

These can live in plugin-local test helpers or debug self-test code, but they should not require network access.

### Remote self-tests

Add gated self-tests following the existing S3 / Curl pattern.

Recommended environment variables:

- `REDSALAMANDER_SELFTEST_CONN_GOOGLEDRIVE`
- `REDSALAMANDER_SELFTEST_CONN_GOOGLEDRIVE_SHARED`

Recommended cases:

- root listing smoke
- pagination over a large folder
- binary file round-trip upload/download/delete
- rename
- move within one root
- same-drive file copy
- duplicate sibling names
- revoked token refresh / reconnect smoke
- two profiles with different identities active in one run

If native Google docs are in v1:

- export-to-local smoke for one Docs/Sheets/Slides item

### Acceptance criteria

- Two Google profiles with different accounts can be opened concurrently.
- Refresh-token persistence survives restart and does not store access tokens.
- Same-drive rename/move do not re-stream file content.
- Same-drive file copy uses the Drive copy API, not download+upload.
- Release packaging does not add a heavyweight SDK family.

## Spec Updates Required

Add:

- `Specs/FileSystem/FileSystem_GoogleDrive.md`

Update:

- `Specs/Core/Core_ConnectionManager.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md`

Document at minimum:

- profile schema
- OAuth secret-storage behavior
- duplicate-name path contract
- native Google docs export limitations
- unsupported v1 features

## Phased Delivery

### Phase 1: Host auth groundwork — DONE (2026-07-02 folder review)

- [x] add OAuth auth mode and secret kind (`ConnectionAuthMode::OAuth2Pkce`, `HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN`)
- [x] add generic secret save/delete APIs (`SetConnectionSecret`/`DeleteConnectionSecret` in `HostServices.cpp`)
- [x] add Connection Manager protocol entry and schema support

### Phase 2: Read-only Drive core — PARTIAL (2026-07-02 folder review)

- [x] plugin skeleton
- [ ] OAuth PKCE flow — interactive sign-in (browser launch + loopback listener + initial token exchange) does not exist for Google; only refresh of a pre-stored token is implemented (see closeout-checklist caveat above)
- [x] directory listing
- [ ] file download
- [x] quota / account info
- [x] shared drive root support

### Phase 3: Write operations

- folder create
- upload / overwrite
- rename
- move
- delete
- server-side file copy

### Phase 4: Hard cases

- duplicate-name disambiguation
- shortcuts
- native Google docs export behavior
- retry / rate-limit hardening

### Phase 5: Self-tests and spec completion

- add local deterministic tests
- add gated remote self-tests
- finish specs and baseline runs

## Open Product Decisions

These should be decided before implementation starts in earnest:

1. Should RedSalamander ship a built-in Google OAuth client ID, or require users to provide their own?
2. Is v1 allowed to expose native Google docs as export-only read paths?
3. Is the proposed duplicate-name suffix acceptable UX, or do we want to pay for a deeper host ABI change later?

## Recommendation

Proceed with a dedicated `FileSystemGoogleDrive` plugin based on `curl` + `yyjson`, add only `curl[http2]`, and extend the host so OAuth refresh tokens can be stored securely per connection profile.

That is the best fit for this repo's current architecture and the fastest path to a high-performance implementation without taking on an oversized dependency stack.

## References

(Trimmed 2026-07-02 folder review; the full Drive REST guide set is discoverable from these entry points.)

- OAuth 2.0 for native apps (PKCE + loopback): <https://developers.google.com/identity/protocols/oauth2/native-app>
- Drive `files.list` / REST v3 reference: <https://developers.google.com/workspace/drive/api/reference/rest/v3/files/list>
- Drive downloads / export: <https://developers.google.com/workspace/drive/api/guides/manage-downloads>
