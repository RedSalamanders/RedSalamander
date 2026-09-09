# Microsoft Drive Virtual File System Plugin

## Overview

RedSalamander ships a built-in virtual file system implementation in `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.dll` that exposes three UI-visible filesystem plugins backed by Microsoft Graph:

Capability v2 advertises Graph same-drive item Move as native Move. Managed cleanup outside that
primitive requires bound drive/item identity plus a conditional operation; otherwise the host keeps
the source. `quickXorHash` can satisfy verification only with declared byte scope and revision binding.
Hydration-requiring reads are reported so the host can obtain task-scoped consent before recall.

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
- Writes stage data into a local temporary file first. The provider uses the Common delete-on-close reservation helper with the `rsm` prefix; a failed reopen deletes the reservation, and normal close removes the staged pathname through the WIL-owned handle.
  - Up to 250 MiB: simple upload (`/content`)
  - Above 250 MiB: upload session with chunked PUTs
- Every upload-session `202 Accepted` response must contain valid `nextExpectedRanges`. The writer resumes from
  the lowest server-acknowledged missing offset, rejects non-progress/out-of-range/contradictory ranges, and
  accepts completion only through a final `200`/`201` after all declared bytes are acknowledged.
- Same-drive move and rename use Graph `PATCH`. `FILESYSTEM_MOVE_NATIVE_ONLY` selects this
  item-ID-preserving provider route for both single and batch Move. It performs no payload
  copy/delete fallback, and an unknown `moveMode` fails with `E_INVALIDARG` before Graph I/O.
- A failed PATCH transport is ambiguous because Graph may have committed before the response was lost. The provider
  performs one item-ID metadata GET only on that failed-transport path and accepts the mutation as committed only when
  the returned item preserves the requested ID, has the exact requested name, and, for a parent-changing move, has the
  exact requested parent ID. A missing, mismatched, or inconclusive reconciliation preserves the original PATCH
  failure. Completed HTTP conflicts such as `409 nameAlreadyExists` retain their existing error mapping and do not pay
  a reconciliation GET.
- Item-ID DELETE is idempotent. `204` and direct `404/itemNotFound` both satisfy the requested postcondition. After a
  failed DELETE transport, one by-ID GET that reports not-found proves the commit and returns success; a still-present
  item or inconclusive probe preserves the original transport failure. Ordinary successful PATCH and DELETE paths add
  no reconciliation request.
- The public Delete entry point is the ordinary Graph Recycle route. It requires
  `FILESYSTEM_FLAG_USE_RECYCLE_BIN` and rejects a call without that flag before Graph I/O.
  The provider resolves the accepted path to a stable item ID, deletes and reconciles only that ID,
  and sends no root ETag/`If-Match` condition. A replacement installed later at the same path is
  not the accepted object and must survive. Capability v2 therefore advertises
  `operations.recycle: true` and `operations.delete: false`; since R0f-Graph `operations.rename`
  is `true` (the R0f-Graph section owns its typed mutation receipt). The separate Graph
  permanent-delete endpoint remains unavailable until it provides its own exact authority and
  typed mutation receipt.
- Moving a folder onto an existing folder moves children individually but retains the drained source folder and
  reports partial cleanup. Graph item deletion is recursive and the plugin has no atomic "delete only if still
  empty" primitive, so deleting the folder after a re-list would still risk erasing a concurrently-added child.
  Nested folder merges use an explicit frame stack with no arbitrary depth cutoff; the stack retains only the
  open source/destination paths and one directory listing per active frame.
- Each merge child receives one just-in-time destination metadata probe before its first conflict decision. The
  iterative merge threads that explicit `Present` or `Missing` result into the nested move instead of repeating the
  same destination-path GET. Direct moves still probe normally. Retry always re-probes and replaces the hint; a
  destination created after a `Missing` probe must fail closed at the Graph PATCH collision without deleting or
  overwriting either identity. A `Present` result authorizes only the conflict decision; it is never mutation
  authority. Before an existing file is renamed to the overwrite backup, the provider probes the generated rollback
  name, re-resolves the destination path, and requires the exact same nonempty item ID, parent ID, name, and eTag as
  the conflict observation. The backup PATCH carries `If-Match` for that eTag. A missing/replaced/moved/revised
  destination, an inconclusive validation read, or Graph `412 Precondition Failed` maps to the normal conflict result
  and performs no source move or backup DELETE. The paired path validation and conditional PATCH are the
  mutation-time boundary; a fresh GET alone is not sufficient.
- Source-item, source-parent, and destination-parent by-path metadata reads are intentional freshness checks and are
  not reused from directory enumeration. Removing them requires a separate conditional-mutation/path-identity design.
- Merge perf capture emits one aggregate set per operation: `FileOps.MicrosoftDrive.Merge.MetadataGetCount`,
  `FileOps.MicrosoftDrive.Merge.MetadataGetUs`, and `FileOps.MicrosoftDrive.Merge.ChildCount`. Test-enabled
  same-binary proof additionally emits `FileOps.MicrosoftDrive.Merge.Baseline`, `.Candidate`, and `.Improvement`.
  Per-request perf rows are forbidden.
- Overwrite-validation perf capture emits one aggregate set per public move operation only when an overwrite reaches
  the mutation-time check: `FileOps.MicrosoftDrive.OverwriteValidation.MetadataGetCount`,
  `FileOps.MicrosoftDrive.OverwriteValidation.MetadataGetUs`, and
  `FileOps.MicrosoftDrive.OverwriteValidation.ConflictCount`. Test-enabled same-binary proof additionally emits
  `FileOps.MicrosoftDrive.OverwriteValidation.Baseline`, `.Candidate`, and `.Overhead`. The safe path adds exactly one
  metadata GET per overwritten child; missing/no-conflict moves add none, and the existing no-conflict merge
  `83/84 -> 67/68` logical/total GET reduction remains unchanged. The deterministic 16-child fixed-latency advisory
  ceiling is candidate metadata time no greater than 130% of the same-binary pre-validation baseline.
- Graph `$batch`/proactive pacing is deferred for merge. After duplicate-GET removal, the remaining per-child probe is
  the just-in-time conflict/freshness boundary; batching it would widen the stale-snapshot window, while batching
  mutations would alter prompt order, cancellation, throttling, rollback, and partial-result semantics. Reconsider
  only with measured live throttling or large-merge latency evidence and a separate deterministic mixed-result proof.
- Overwrite move results distinguish the committed primary mutation from backup cleanup and rollback. Once the
  requested move has committed, failure to delete the recoverable overwrite backup is logged as cleanup debt and
  does not turn the move into an ordinary failure that callers might destructively retry.
- Cross-drive move returns `ERROR_NOT_SUPPORTED`.
- Server-side copy is not implemented in this version. `CopyItem` and `CopyItems` return `ERROR_NOT_SUPPORTED`, allowing the host bridge to fall back to read/write copy paths where applicable.

### Ranged reader integrity

Microsoft Drive readers issue bounded ranged downloads. The shared HTTP transport has a
64 MiB default response ceiling and each ranged reader tightens it to the requested
chunk size before appending response bytes. A `206` response is accepted only when
`Content-Length`, body length, and a parsed `Content-Range` exactly match the requested
start/end and the known total object size. A `200` response is accepted only for an
exact full-object read starting at zero and without `Content-Range`. Missing, shifted,
short, oversized, malformed, or inconsistent responses fail before any byte is exposed
or cached.

Reader seek arithmetic operates in unsigned file-position space. Positive overflow,
negative results, and the `INT64_MIN` magnitude edge fail explicitly; positions above
`INT64_MAX` remain representable when the requested movement is valid.

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

The deterministic R0d contracts reject non-Recycle public Delete before Graph I/O, prove that
Recycle removes only the path-resolved stable item ID while preserving a same-path replacement,
and require every DELETE request in that route to omit root ETag/`If-Match`. Generic provider
contracts also enforce the `delete:false`, `recycle:true`, `rename:true` capability tuple: Graph
Delete is Recycle by stable item ID (the item goes to the drive's recycle bin; permanent Delete is
honestly absent), and Rename is the PATCH rename/move path with its If-Match revision guard. Every
Recycle and Rename completion carries the provider's `FileSystemItemMutationResult` (success =
known and committed with the original gone; a definitive refusal = known and not committed;
anything else carries no record), so the host never treats a successful Graph mutation as
indeterminate. The plugin advertises `IFileSystemAtomicWriter`: a simple upload is one PUT (with
`If-None-Match: *` unless overwrite was granted) and an upload session's item appears only when the
session completes under the requested conflict behavior, so the host bridge publishes new names
into Graph destinations through the provider writer. A replacement granted through the host
bridge (R3-1) is conditional: the writer receives the occupant the user saw
(`IFileWriterExpectedReplacement`), resolves the item again before publication, refuses a vanished
or changed item with `ERROR_REVISION_MISMATCH`, and sends `If-Match` with the item's eTag on the
content PUT and on upload-session creation (`412` -> `ERROR_REVISION_MISMATCH`, nothing replaced).
The host's token is the occupant's last-write time: when that token or the live item's timestamp is
zero there is no identity to compare, and the writer refuses with `ERROR_REVISION_MISMATCH` before
any upload instead of pinning `If-Match` to whichever item is live (R0-RC4, the same rule as the
Dummy and Curl writers).

Content proof (R3-2). The item Graph returns when an upload completes (simple PUT, upload session,
streaming session) carries `file.hashes`: `sha256Hash` and `quickXorHash` on OneDrive for Business
and SharePoint, `sha1Hash` and `quickXorHash` on personal OneDrive. The writer declares all three
algorithms through `IFileWriterContentProof`, reports whichever the drive returned, and the route
advertises `verification.providerProof: "writer-digest"`; the host hashes the streamed bytes with
the declared algorithms and compares the reported one, so Verified Copy needs no readback and
Managed Move into a Graph drive deletes its source only after the proof matches. An item without
hashes leaves verification `Unavailable` and keeps the Move source.

Cancellation containment (R0f-Graph). The route advertises `routeClass:"providerWatchdog"`,
`deadline:true`, `abort:false`, and `providerWatchdogTimeoutMs` =
`GraphProviderWatchdogTimeoutMs(connectTimeoutMs, requestTimeoutMs)` = 2 x connect + 2 x request
(the WinHTTP resolve, connect, send, and receive timeouts of one attempt; 140 s with the default
timeouts; transport failures are never retried, only 429/5xx throttle replies are). Mechanism: the
mutation entry points (`CopyItem`, `MoveItem`, `RenameItem`, `DeleteItem`, `CopyItems`,
`MoveItems`, `DeleteItems`, `RenameItems`) keep the owning call's `FileSystemOptions` on the
current thread (`GraphOperationOptionsScope`); `SendHttpRequest` polls the host's operation
control before each attempt, between request-body chunks, between response chunks, and while
sleeping before a throttle retry, and reports the host's own verdict (`ERROR_CANCELLED` /
`ERROR_TIMEOUT`) for any transport failure observed after Cancel or a passed deadline
(`GraphTransportFailure`). A server that stops answering fires no chunk, so such a request returns
through the timeouts above. Readers and writers opened by the host bridge run without a task
option context and are bounded the same way. File Operations admits same- and cross-provider
mutation on Graph routes. Witnesses: the Microsoft Drive debug self-test
`RunGraphStalledRequestCancelSelfTests` against the in-tree fake Graph endpoint
(`FileSystemMicrosoftDrive.SelfTest.FakeGraph.cpp` over `Common/LoopbackHttpFixture.h`: item
metadata by path and id, children with `$top` paging, folder creation, PATCH with If-Match,
DELETE, simple content upload, upload sessions with Content-Range chunks, ranged download; the
plugin is redirected to it through the test-only `ConfigureFakeGraph` seam, which also bypasses
the OAuth token and uses the synthetic `/@conn:microsoft-drive-selftest` context) proves that a
canceled Recycle whose children listing is still arriving returns `ERROR_CANCELLED` within a few
read chunks and deletes nothing, and that a request the server never answers returns within the
declared bound with the host's verdict after Cancel and as a transport failure without it; FileOps
`R0fGraph_FakeGraphReadWriteCreateRenameRecycle` proves cross-provider Copy both ways
(byte-exact), provider CreateDirectory and Rename, and Recycle through a task on the same fixture.
The gated `REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_*`/`SHAREPOINT` profiles remain the live proof.
OneDrive and SharePoint are full file-manager destinations.

Optional selftest coverage is wired through Connection Manager profiles:

- `--fileops-selftest`
  - secret retrieval and sandbox validation for OneDrive Personal, OneDrive Business, and SharePoint
- `--compare-selftest`
  - remote compare smoke
  - remote directory-size callback contract

Profile names and environment-variable overrides are documented in `Specs/Testing/Testing_SelfTestRemoteCredentials.md`.

Credential-boundary and merge coverage is deterministic and does not require live credentials:

- Test-enabled Debug and Release `PluginContractTests` run the Microsoft Drive exported selftests for strict URL validation, approved
  same-origin pagination, foreign/plaintext continuation rejection before a second request, repeated-link rejection
  before a third request, server-acknowledged upload offsets, concurrent-child-safe merge cleanup, committed-move
  cleanup debt, and injected Graph/upload transport failures whose captured diagnostics contain no bearer or opaque-query
  sentinels. The merge-metadata contract compares the legacy and candidate paths in one binary with fixed GET latency,
  requires the exact 16-child request reduction from 83 to 67 logical metadata GETs (84 to 68 total Graph GETs) while
  retaining 16 PATCHes, and locks overwrite, Retry-to-missing, nested-folder, and late-destination collision behavior.
  Overwrite-mutation coverage moves an observed destination elsewhere, replaces it at the same path with a new ID,
  and rewrites it in place after validation so the conditional PATCH receives `412`. It covers `MoveItem` and
  `MoveItems`, pre-authorized and prompt-granted overwrite, sequential and parallel public entry points, exact
  identity/content/eTag preservation, MOVE-source retention, and zero backup DELETEs on conflict. The 16-child
  same-binary perf contract requires the exact `83/84 -> 99/100` logical/total GET change, unchanged 32 PATCHes and
  16 cleanup DELETEs, one validation GET per overwritten child, and the 130% advisory ceiling above.
  Mutation-reconciliation coverage injects pre-commit and committed-then-transport-failure outcomes for overwrite
  backup PATCH, primary PATCH, and DELETE. It proves exact name/parent matching, no incorrect rollback or orphaned
  backup, direct already-gone DELETE success, a first DELETE that commits but returns `503` followed by an automatic
  retry that receives `404/itemNotFound`, original-failure preservation when the item remains, exactly one by-ID GET
  for each ambiguous transport, zero reconciliation GETs for ordinary success/direct-or-retried DELETE not-found, and
  no extra GET for a definitive late-destination `409` conflict.
  Ranged-reader coverage accepts one exact `206` control and rejects missing/malformed/
  shifted totals, short/oversized bodies, inconsistent lengths, invalid full-body `200`
  responses, response-cap overflow, positive seek overflow, negative seek, and the
  `INT64_MIN` edge before publishing bytes.
  The deep-merge witness moves a leaf through 70 existing folder pairs and reports only the deliberately retained
  source-folder cleanup as partial, proving that the former depth-64 recursion stop is gone.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` preserves the validated URL types, redirect suppression,
  separated/secure-cleared authorization header, continuation guard, raw-URL/header logging bans, Release-capable
  test gates, merge hint state/replacement, exact request formulas, and aggregate metric wiring.
- Closeout evidence is archived under
  `Specs/TestRuns/4cb089111a23/FileSystemMicrosoftDrive/2026-07-16_2046_observatory_track3_credential_boundary/`.
- Merge destination-metadata reuse evidence is archived under
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-21_131329_msdrive_merge_metadata/`; the detailed implementation
  history and validation matrix are in
  `Specs/Plans/Done/Operation_MicrosoftDrive_MergeDestinationMetadataReuseOptimization_2026-07-21.md`.
- Overwrite mutation-time validation evidence is archived under
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_1600_msdrive_overwrite_validation/`; the failure-first checkpoint is
  under `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_1545_msdrive_mutation_red/`, and the implementation history
  remains in archived EG-MSD-01 of
  `Specs/Plans/Done/Operation_Emberglass_48Hour_RepositoryDeepReview_2026-07-22.md`.
