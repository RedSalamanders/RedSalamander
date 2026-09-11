# S3 / S3 Table Virtual File System Plugin

## Overview

RedSalamander ships a built-in virtual file system implementation in `Plugins/FileSystemS3/FileSystemS3.dll` that exposes **two UI-visible file system plugins**:

Capability v2 is bucket/prefix scoped. `FILESYSTEM_MOVE_NATIVE_ONLY` selects the provider-owned
pinned-revision route: conditioned server-side copy or conditioned bounded relay, followed only by
exact conditional source deletion. The relay is an internal qualified alternative, not a host
pathname copy/delete fallback. Unknown Move modes fail before S3 I/O. Host-managed Move may remove
a source only after binding an exact `VersionId` and performing a conditional delete; otherwise it
reports Copy-only. ETag alone is not a revision id. ETag or CRC64NVME may replace destination
read-back only when algorithm, byte scope, and revision binding are all declared and contract-tested.

- **S3** (`s3:`)
- **S3 Table** (`s3table:`)

Both are implemented using the AWS SDK for C++ (`aws-sdk-cpp`), with the S3 CRT client for S3 operations and the S3 Tables client for S3 Table operations.

## Plugin Identities

| Display name | `PluginMetaData.id` | `PluginMetaData.shortId` |
|-------------|----------------------|--------------------------|
| S3          | `builtin/file-system-s3`      | `s3`      |
| S3 Table    | `builtin/file-system-s3table` | `s3table` |

## Navigation (URI Syntax)

### Path normalization rules

- `\` is treated as `/`.
- Multiple slashes are collapsed **except** the leading `//` after the scheme (authority prefix).

### S3

Root behavior:

- `s3:/` lists buckets (as directories).
  - If the active Connection Manager profile has a non-empty region, the list is filtered to buckets in that region.
  - If the region is empty, the list is unfiltered and bucket regions are resolved automatically when entering a bucket.

Within a bucket:

- `s3:/<bucket>/` lists “folders” (common prefixes) and objects under the current prefix.
- Folder traversal is modeled via `ListObjectsV2` with delimiter `/` and is **paginated** (directory listings are complete; not capped at 1000 items).

Authority form:

- `s3://<bucket>/<path>` is accepted and is normalized internally to `/<bucket>/<path>`.

Examples:

- `s3:/`
- `s3:/my-bucket/`
- `s3://my-bucket/logs/2026/`

### S3 Table

Root behavior:

- `s3table:/` lists S3 Table buckets (as directories).

Within a table bucket:

- `s3table:/<tableBucket>/` lists namespaces (as directories).
- `s3table:/<tableBucket>/<namespace>/` lists tables as files named `"<table>.table.json"`.

Opening a `*.table.json` “file” returns a JSON document derived from `GetTable` for the selected table.

Examples:

- `s3table:/`
- `s3table:/my-table-bucket/default/`
- `s3table:/my-table-bucket/default/my_table.table.json`

## Configuration (Per Plugin)

Each plugin exposes its own configuration schema via `IInformations::GetConfigurationSchema()` and is configured in Preferences.

### S3 keys

- `defaultRegion` (string, default `us-east-1`)
- `defaultEndpointOverride` (string, default empty)
- `useHttps` (bool, default `true`)
- `verifyTls` (bool, default `true`)
- `connectTimeoutMs` (integer, `1..600000`, default `10000`): TCP connect timeout used for S3 requests (ms).
- `requestTimeoutMs` (integer, `1..600000`, default `30000`): network stall / socket-read timeout for S3 requests (ms). Does not cap total transfer time if data keeps flowing.
- `useVirtualAddressing` (bool, default `true`)
- `maxKeys` (integer, `1..1000`, default `1000`): page size (max keys requested per `ListObjectsV2` request).

### S3 Table keys

- `defaultRegion` (string, default `us-east-1`)
- `defaultEndpointOverride` (string, default empty)
- `useHttps` (bool, default `true`)
- `verifyTls` (bool, default `true`)
- `connectTimeoutMs` (integer, `1..600000`, default `10000`): TCP connect timeout used for S3 requests (ms).
- `requestTimeoutMs` (integer, `1..600000`, default `30000`): network stall / socket-read timeout for S3 requests (ms). Does not cap total transfer time if data keeps flowing.
- `maxTableResults` (integer, `1..1000`, default `1000`): page size (max results requested per `ListNamespaces` / `ListTables` request).

Security note:

- AWS secrets are **not** stored in plugin settings. For secure storage (WinCred + optional Windows Hello), use Connection Manager (`/@conn:<connectionName>/...`).

## Connection Manager Integration

Both plugins support host-reserved navigation:

- `/@conn:<connectionName>/...` (see `Specs/Core/Core_ConnectionManager.md`)
- Optional shorthand authority form: `// @conn / <connectionName> / ...` (e.g. `s3://@conn/<connectionName>/...`)

Profile mapping:

- `ConnectionProfile.pluginId`:
  - S3: `builtin/file-system-s3`
  - S3 Table: `builtin/file-system-s3table`
- `ConnectionProfile.host`: AWS region (e.g. `us-east-1`).
  - If empty: treated as “auto” (list all buckets, then resolve per-bucket region on demand).
  - If non-empty: used as the signing/endpoint region and as a filter for bucket listing.
- `ConnectionProfile.userName`: Access key id (optional; when omitted the AWS default credential chain is used)
- Secret access key:
  - stored as the connection `password` secret (WinCred generic credential),
  - retrieved via `IHostConnections::GetConnectionSecret(HOST_CONNECTION_SECRET_PASSWORD)` and may be prompted via `PromptForConnectionSecret(...)` when needed.

Supported `extra` keys (from `ConnectionProfile.extra` and surfaced to plugins via `GetConnectionJsonUtf8().extra`):

- `endpointOverride` (string)
- `useHttps` (bool)
- `verifyTls` (bool)
- `useVirtualAddressing` (bool, S3 only)

Host-level extra keys:

- `copyMoveMaxConcurrency` / `deleteMaxConcurrency` (integer): global per-connection overrides used by the host to clamp file-op concurrency (see `Specs/Core/Core_ConnectionManager.md`).

## Operations and Behavior Notes

- **Read** is supported. File reads download the remote object to a local delete-on-close temporary file before streaming it to the host. S3 handle-backed staging uses the Common reservation helper with the `rs3` prefix; reservation ownership remains RAII-held until a valid delete-on-close WIL handle exists, including reopen-failure paths.
- **Write** is supported (uploads/overwrites objects). A handle-backed upload is successful only when the SDK consumes exactly the declared `ContentLength`; a short source is `ERROR_PARTIAL_COPY`, even if the SDK reports success.
- **Delete** is supported for objects and for recursive prefixes when the host passes `FILESYSTEM_FLAG_RECURSIVE`. Ordinary object Delete consumes the ETag returned by its current-object probe, sends `DeleteObject` with that ETag as `If-Match`, and omits `VersionId`. Recursive flat-prefix Delete retains one canonical account/profile/bucket/trailing-`/` prefix, lists current members with their ETags, issues one serialized ETag-conditioned `DeleteObject` per observation, and re-lists until one complete listing is empty. The active SDK's multi-delete request cannot express a distinct `If-Match` for every member, so unconditioned `DeleteObjects` batching is not an ordinary-delete authority. A known condition mismatch mutates nothing and may be rebound only from a later bounded listing; any other request failure, unsupported/missing ETag, cancellation, or deadline stops immediately without retrying an unknown request result. At most 64 mutation passes run before one final residual listing returns success only if empty, otherwise `ERROR_RETRY`. Non-recursive delete of a prefix is denied unless the path resolves to a concrete object key.
- **Copy/Move and internal Rename paths** are implemented within S3 using provider-side object operations. The implementation uses single-object copy for small objects, multipart copy for large objects, conflict prompts, `.rs-bak` backup objects, and rollback/cleanup paths before deleting move sources. Once the requested destination mutation commits, backup-cleanup failure is recorded and logged as cleanup debt; it does not turn the committed primary operation into an ordinary failure. Capability v2 keeps `operations.rename: false` until the central Rename route can bind and consume the required exact authority and typed receipt. This does not disable the R0b-conditioned provider-owned ordinary object/prefix Delete route, which remains `operations.delete: true`; Recycle remains false.
- **Cross-filesystem copy/move** uses the host bridge via `read`/`write`.
- **S3 Table** is read-only: it supports enumeration/properties/read and advertises no write/import capability.
- Cancellation containment (R0f-S3). S3 advertises `routeClass:"providerWatchdog"`,
  `deadline:true`, `abort:false`, and `providerWatchdogTimeoutMs` =
  `S3ProviderWatchdogTimeoutMs(connectTimeoutMs, requestTimeoutMs)`: per attempt the TCP connect
  timeout plus the CRT stall monitor, which closes a connection that moved no bytes for
  max(3 s, `requestTimeoutMs`), evaluates once per second, and in practice fires after about two
  intervals (measured 12-16 s for two attempts at a 3 s interval); the client makes at most two
  attempts with exponential backoff capped at 2 s (144 s with the default timeouts). Cancel on a
  silent server returns after the current attempt (one stall cycle); Cancel while bytes flow
  returns within one callback. Mechanism: the mutation entry points (`CopyItem`,
  `MoveItem`, `RenameItem`, `DeleteItem`, `CopyItems`, `MoveItems`, `DeleteItems`, `RenameItems`)
  keep the owning call's `FileSystemOptions` on the current thread (`S3OperationOptionsScope`);
  every AWS request issued under them is armed with a continue handler (`ArmS3RequestControl`)
  that the CRT polls from its header, body, and progress callbacks and that cancels the meta
  request as soon as the host's operation control fails; a request that fails after Cancel or a
  passed deadline reports the host's own verdict (`ERROR_CANCELLED` / `ERROR_TIMEOUT`), never a
  transport error. A server that stops answering fires no callback, so such a request returns
  through the provider-owned bound above. The ranged reader opened through `CreateFileReader`
  implements `IFileReaderOperationControl` (R0f-S3, C8): the host hands the task's options right
  after creation, `Read` establishes the same per-thread scope on the reader's own thread and checks
  the control before each request, so a cancel ends an in-flight `GetObject` within the bound
  instead of only between two Reads. Writers opened by the host bridge still run without a task
  option context and are bounded by the transport policy alone. File Operations admits same- and
  cross-provider mutation on S3 routes; permanent Delete runs through the host's native authority
  (`DeleteItem`, R0b exact generation), and every Delete completion carries the provider's
  `FileSystemItemMutationResult` (success = known and committed with the original gone; a
  definitive refusal = known and not committed; anything else carries no record). `anonymous` (plugin configuration or a connection
  profile's `extra`) sends unsigned requests for public buckets and the loopback fixture. S3 Table
  stays read-only and `uncontained` (no mutation is admitted, so no containment claim is needed).
  Witnesses: the S3 debug self-test `RunS3StalledRequestCancelSelfTests` against the in-tree fake
  S3 endpoint (`FileSystemS3.SelfTest.FakeS3.cpp`, a loopback path-style HTTP/1.1 S3 with objects,
  `list-type=2`, server-side copy, multipart, conditional headers and ranges) proves that a
  canceled request whose listing body is still arriving returns `ERROR_CANCELLED` within a few
  body callbacks, that a request the server never answers returns within the declared bound, and
  that a ranged-reader `GET` the server never answers returns `ERROR_CANCELLED` within that bound
  once the host has handed the reader its control (scenario 4, C8),
  with the host's verdict after Cancel and as a transport failure without it; FileOps
  `R0fS3_FakeS3ReadWriteCreateDelete` proves cross-provider Copy both ways (byte-exact), provider
  CreateDirectory and Rename, and Delete through a task on the same fixture. The gated
  `REDSALAMANDER_SELFTEST_CONN_S3` profile remains the live proof. S3 is a full file-manager
  destination. Residual: `operations.rename` stays `false` until the central Rename route consumes
  S3's typed receipt; the provider `RenameItem` works and is exercised by the fixture case.
- Directory-size progress/cancellation failures are authoritative, including the final completion callback; the plugin must not overwrite that callback status with `S_OK`.
- A failed multipart abort transfers the upload session and its owning filesystem instance to a bounded background reconciliation queue. Each pass attempts at most 16 sessions once, failed sessions cool down for at least one second, and pending/active cleanup keeps runtime unload closed. The queue logs the path, attempt count, and status so an orphan risk is observable rather than silently discarded by the writer destructor.
- `ListObjectsV2`, recursive-transfer planning, directory-size enumeration, and S3 Table namespace/table paging
  use the shared page/item/retained-token-byte/deadline guard. Truncated S3 responses must provide a non-empty
  next token, and repeated tokens fail before another provider request instead of looping.
- `concurrency.deleteMax` is `8` so the host can parallelize deletes when many objects are selected.

### Directory models and durable empty directories

S3 capability v2 is bucket/path scoped and returns one explicit directory model:

| Profile | `directories.model` | Directory representation |
|---|---|---|
| Ordinary/general-purpose bucket | `flatPrefix` | A directory is a key prefix; an empty directory is durably represented by a zero-byte object whose key ends in `/` |
| Future directory bucket/profile with provider-native directory semantics | `native` | The provider uses the service's native directory operation and does not create marker objects; the current in-tree provider does not advertise this profile |
| S3 Table | `providerVirtual` | Read-only virtual hierarchy; Create Directory and directory transfer publication are unsupported |

The provider must resolve the bucket profile for the requested path. It must not reuse one
instance-wide directory model for a different bucket, treat S3 Table as a directory bucket, or infer
native semantics merely from a slash in the key.

For `flatPrefix`:

- `CreateDirectory(".../photos")` normalizes the key to `photos/` and conditionally publishes a
  zero-byte marker. Success is reported only after the committed marker is queryable under the
  provider's publication contract; an in-memory synthetic prefix is not durable success.
- `GetAttributes` never infers existence merely because path text ends in `/`. It reports a flat
  directory only after an exact marker or at least one current object under the normalized prefix
  is queryable; an absent trailing-slash path is `ERROR_FILE_NOT_FOUND`.
- Copy/Move of an empty source prefix publishes the destination marker. A non-empty tree may also
  carry an explicit source marker; transferring its children does not silently discard that marker
  when it is required to preserve the source directory itself.
- The marker is an engine-owned publication object with the same exact revision/conditional
  publication and cleanup rules as another S3 object. A pathname or trailing-`/` pattern alone is
  never rollback/delete authority.
- An existing marker means directory-on-directory merge. A non-marker object at the would-be
  directory key or an object/directory ancestor mismatch is a typed conflict; it is never silently
  overwritten to manufacture a folder.
- Deleting a non-empty prefix remains recursive and bounded. Deleting an empty directory removes
  the exact bound marker; failure or unknown outcome preserves it and reports the typed result.
- Ordinary current-key Delete never supplies `VersionId`. In a versioned bucket its matching
  ETag-conditioned request creates a current delete marker and preserves historical versions.
  Supplying an exact historical `VersionId` remains a separate fixed-set cleanup/rollback operation;
  it is never inferred from ordinary visible-folder deletion. The prefix-name object (`prefix`),
  siblings such as `prefix2/`, other buckets, and other profiles are outside a selected `prefix/`.

For a future `native` profile, Create Directory and empty-directory Copy/Move use the native service
operation and bind its returned object/revision identity. They never create `name/` marker objects.
The current in-tree S3 provider exposes only `s3-flat-prefix` and read-only S3 Table, so it must not
advertise `directories.model: native`. If a provider cannot prove which model applies, directory
creation and empty-directory publication fail closed rather than returning a session-only folder.

The deterministic S3 graph tests cover absent trailing-slash attribute lookup, flat marker creation,
post-publication attribute lookup, conditional create collision, and
empty-prefix Copy/Move publication/cleanup. Capability-contract and host-admission tests require
`createDirectory: true` only for `s3-flat-prefix` and false for S3 Table. Native directory-bucket
behavior remains a future executable-matrix row and may not be advertised until its operation and
no-marker behavior are implemented and contract-tested.

### Atomic final-key publication

The cross-filesystem bridge uses S3's final-key writer because `PutObject` and multipart completion are the
publication boundaries; the earlier destination probe is conflict-discovery only. When
`FILESYSTEM_FLAG_ALLOW_OVERWRITE` was not granted, both the small-object `PutObject` request and
`CompleteMultipartUpload` MUST carry `If-None-Match: *`. A conditional `412 Precondition Failed` or the relevant
concurrent-write `409 Conflict` maps to `HRESULT_FROM_WIN32(ERROR_FILE_EXISTS)`, leaves the object that won the race
untouched, and prevents MOVE source cleanup. A replacement granted through the host bridge (R3-1) is conditional
too: the writer records the occupant's ETag, size, and last-write time when it is created with the overwrite flag,
validates the host's expectation (`IFileWriterExpectedReplacement`) against that occupant, and carries
`If-Match: <ETag>` on `PutObject` or `CompleteMultipartUpload`, so a changed object is `412` ->
`ERROR_REVISION_MISMATCH` and nothing is replaced. A host expectation or an occupant without a last-write time
carries no identity to validate, so `SetExpectedReplacement` refuses it with `ERROR_REVISION_MISMATCH` before
any upload (R0-RC4). Only a provider-native overwrite without a host expectation stays
unconditional. An S3-compatible endpoint that rejects or cannot honor conditional publication therefore fails the
no-overwrite operation closed; the implementation MUST NOT substitute another racy `HEAD` request.

Content proof (R3-2). The writer declares the full-object CRC-64/NVME on `CreateMultipartUpload` (`ChecksumType
FULL_OBJECT`), on every `UploadPart`, and on single-request `PutObject`; S3 validates each request and returns the
digest of the stored object (`x-amz-checksum-crc64nvme`, or `ChecksumCRC64NVME` in the `CompleteMultipartUpload`
result). The writer reports that digest through `IFileWriterContentProof`, the route advertises
`verification.providerProof: "writer-digest"`, and the host compares it with the bytes it streamed: Verified Copy
needs no readback and Managed Move into S3 deletes its source only after the proof matches. An endpoint that
returns no digest leaves verification `Unavailable` and keeps the Move source.

Within-provider directory COPY/MOVE uses the same publication boundary. After a just-in-time refresh observes a
destination missing, a no-overwrite small-object `CopyObject` and multipart `CompleteMultipartUpload` MUST carry
`If-None-Match: *`. If server-side copy falls back to a download/upload relay, the relay's `PutObject` or multipart
completion MUST retain the same condition. A late `409`/`412` maps to `ERROR_FILE_EXISTS`: an interactive transfer
refreshes and re-enters the ordinary conflict decision path, while a transfer without a conflict callback fails
closed. MOVE source deletion occurs only after conditional destination publication succeeds. A compatible relay is
allowed when the server-side copy form rejects the condition; if the relay also rejects it, the operation fails
closed without replacing the destination or deleting the source.

The conditional header adds no provider round trip: small writers retain one `PutObject`, small direct copies retain
one `CopyObject`, and multipart writers/copies retain one completion request after their ordered part uploads.
Deterministic provider coverage injects a destination after the preflight boundary for every publication form,
exercises the `412` and `409` mappings, verifies the injected bytes are preserved, keeps MOVE sources on failure,
and retains separate overwrite-authorized replacement controls.

### Exact source revision for COPY and MOVE

Within-provider COPY/MOVE MUST bind every source read and final MOVE deletion to one revision captured during
planning. `S3ObjectRevision` carries the ETag and, when the bucket returns one, the VersionId alongside the planned
key and byte count. Recursive planning performs one `HeadObject` per listed object, requires a usable identity,
requires size and list/HEAD ETag agreement, and fails with revision mismatch before destination mutation if the
listing became stale. A single-object transfer uses the identity returned by its existing object probe.

#### CopySource has one escaping boundary

`CopyObject` and `UploadPartCopy` pass the logical UTF-8 `bucket/key[?versionId=value]` value to the AWS SDK. The
provider MUST NOT percent-encode that value first. The pinned aws-sdk-cpp 1.11.880 request serializers own the sole
wire boundary: both `GetRequestSpecificHeaders()` implementations apply `URI::URLEncodePath` while creating the
`x-amz-copy-source` header. Pre-encoding in the provider would therefore encode `%` a second time, corrupt reserved,
Unicode, nested, and versioned keys, and can turn a valid direct copy into an avoidable relay fallback.

An SDK version change requires re-auditing both generated serializers before changing this ownership rule. Direct
tests MUST inspect the SDK-produced header bytes for ordinary, space, `%`, `+`, `/`, Unicode, and VersionId inputs;
request-object getters alone are insufficient. The deterministic transfer graph must additionally prove reserved and
versioned keys preserve bytes and select direct single-request and multipart routes with no relay.
The manifest floor and exact override therefore remain `aws-sdk-cpp` 1.11.880
until the next audit is performed. The 1.11.840 -> 1.11.880 move was audited against the generated
`aws-cpp-sdk-s3-crt` sources: `CopyObjectRequest` and `UploadPartCopyRequest` each still emit exactly one
`x-amz-copy-source` header through `URI::URLEncodePath`, and that helper is byte-identical between the two
releases, so provider-side ownership is unchanged. A floor alone is not a pin: a newer baseline may
select a later default. The governed vcpkg installer must reject either constraint
when its exact entry is absent from the selected `builtin-baseline` registry before
dependency mutation.

Test-enabled
Release x64 evidence is archived at
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-09_152411_s3_copy_source/` (157/157 assertions). It proves the request
count route but makes no live endpoint latency claim because reviewed credentials were unavailable; existing
`PublicationUs`, `PublicationRequestCount`, and `RelayFallbackCount` metrics remain the live timing contract.

- Versioned `CopyObject` and every `UploadPartCopy` use the captured VersionId in `CopySource` and validate the
  returned source VersionId. The captured ETag is also supplied as the source condition.
- Unversioned server-side copy and multipart parts use source `If-Match` with the captured ETag. A failed condition
  is `ERROR_REVISION_MISMATCH`; it is never eligible for an unpinned relay retry.
- Relay `GetObject` uses the captured VersionId when available and otherwise `If-Match`. The response identity,
  advertised length, and actual downloaded byte count MUST all match the plan before upload begins.
- Final publication MUST return the created destination ETag and, when versioning is enabled, its VersionId. The
  transfer journal stores that immutable publication identity together with each exact backup revision. Rollback
  reads only the recorded publication/backup revision and removes only an operation-owned revision. A concurrent
  destination version is never restored from, overwritten, or deleted. If a successful endpoint response does not
  provide a usable publication identity, the operation preserves every potentially authoritative object and returns
  partial failure.
- Versioned MOVE hides the source key with `DeleteObject` **without** a VersionId and with `If-Match` set to the
  captured current ETag. The created delete-marker VersionId is recorded so rollback can remove that exact marker
  and reveal the unchanged captured source version without copying or deleting any historical version. Unversioned
  MOVE uses the same conditional current-key deletion and restores only from the exact recorded destination
  publication when rollback is required. A missing/mismatched condition preserves the source and rolls back only
  operation-owned destination revisions; deletion never falls back to a check-then-unconditional request.

A newer source version therefore makes MOVE fail closed instead of deleting the older captured VersionId and
exposing history. COPY may still read the explicitly captured immutable version while a newer version remains
current. Endpoints that cannot honor the required source copy/read/delete condition or cannot identify the created
publication/delete marker are incompatible with this operation and MUST also fail closed or return partial failure
while preserving ambiguous objects.
Destination no-overwrite conditions remain independently mandatory. When a small direct copy returns the ambiguous
`412` while both source and destination conditions were present, one conditioned source `HeadObject` distinguishes a
still-valid source plus destination race from a source mismatch; that exceptional probe is included in source
identity metrics.

### Directory-transfer destination probing

S3 directory COPY/MOVE conflict discovery is per object. A normal interactive transfer, or a transfer that already has `FILESYSTEM_FLAG_ALLOW_OVERWRITE`, MUST NOT perform a whole-plan destination metadata preflight. It refreshes each destination object just in time before the first conflict decision so a destination created after planning is not silently overwritten.

The whole-plan preflight is retained only when overwrite is not authorized and no conflict callback exists. That mode MUST return `ERROR_ALREADY_EXISTS` before initial progress or mutation when any planned destination object or object-as-directory ancestor is blocked. A Retry answer MUST refresh again before another decision. After an ancestor blocker is backed up and removed, the transfer MUST refresh again because a deeper stacked blocker may remain. A planned destination ancestor created by the same transfer is never removed to make room for its descendant; the descendant remains a partial/skipped result.

When performance capture is enabled, a completed transfer emits aggregate rows rather than per-request JSONL traffic:

- `FileOps.S3.DirectoryTransfer.DestinationProbeCount`: underlying destination/ancestor summary probes; `value0` is probe count and `value1` is planned object count.
- `FileOps.S3.DirectoryTransfer.DestinationProbeUs`: accumulated summary-probe time with the same count/object values.
- `FileOps.S3.DirectoryTransfer.DestinationRefreshCount`: destination-state refresh count in `value0` and planned object count in `value1`.
- `FileOps.S3.DirectoryTransfer.PublicationRequestCount`: direct plus relay publication requests in `value0`; `value1` is the conditional subset.
- `FileOps.S3.DirectoryTransfer.PublicationUs`: accumulated publication time with the same request/conditional-request values.
- `FileOps.S3.DirectoryTransfer.RelayFallbackCount`: relay publications in `value0` and planned object count in `value1`.
- `FileOps.S3.DirectoryTransfer.PublicationConflictCount`: conditional publication conflicts in `value0` and planned object count in `value1`.
- `FileOps.S3.DirectoryTransfer.SourceIdentityProbeCount`: revision-establishing or ambiguity-resolving source HEAD requests in `value0`; `value1` is planned object count.
- `FileOps.S3.DirectoryTransfer.SourceIdentityProbeUs`: accumulated source identity-probe time with the same count/object values.
- `FileOps.S3.DirectoryTransfer.SourceConditionalReadCount`: conditioned `CopyObject`, `UploadPartCopy`, and relay `GetObject` requests in `value0`; `value1` is planned object count.
- `FileOps.S3.DirectoryTransfer.SourceConditionalDeleteCount`: conditioned MOVE source deletes in `value0`; `value1` is planned object count.
- `FileOps.S3.DirectoryTransfer.SourceRevisionMismatchCount`: source condition/identity mismatches in `value0`; `value1` is planned object count.
- `FileOps.S3.DirectoryTransfer.ImmutablePublicationCount`: final requested publications whose response supplied a usable ETag/VersionId in `value0`; `value1` is planned object count.
- `FileOps.S3.DirectoryTransfer.VersionedDeleteMarkerCount`: versioned MOVE source-hiding requests that returned an exact delete-marker identity in `value0`; `value1` is planned object count.
- `FileOps.S3.DirectoryTransfer.RollbackSourceRestoreCount`: rollback source identities restored by exact marker removal or exact publication copy in `value0`; `value1` is planned object count.
- `FileOps.S3.DirectoryTransfer.ExactOwnedDeleteCount`: exact operation-owned publication, marker, and backup revisions removed during rollback/cleanup in `value0`; `value1` is planned object count.

Ordinary recursive virtual-folder Delete emits one aggregate terminal set per selected prefix:

- `FileOps.S3.VirtualFolderDelete.ElapsedUs`: complete operation time; `value0` is total observed members.
- `FileOps.S3.VirtualFolderDelete.PassCount`: complete listing passes, including the final empty/residual verification; `value1` is total observed members.
- `FileOps.S3.VirtualFolderDelete.ObservedObjectCount`: observations across all listings; `value1` is pass count.
- `FileOps.S3.VirtualFolderDelete.ConditionalRequestCount`: ETag-conditioned current-key delete requests; `value1` is total observed members.
- `FileOps.S3.VirtualFolderDelete.RevisionMismatchCount`: condition mismatches left for fresh re-list/rebind; `value1` is total observed members.
- `FileOps.S3.VirtualFolderDelete.ResidualObjectCount`: members in the last complete listing or still unconsumed from a failed pass; `value1` is pass count.

The test-enabled Release contract `PluginContractTests.exe --s3-r0b-delete-selftests` deletes 4,096
listed 4-KiB fake objects and emits `FileOps.S3.VirtualFolderDelete.SelfTestUs`. It requires exactly
4,096 conditional requests, two listings (work plus empty verification), zero residuals, one live
request at a time, and less than 2,000,000 us per sample. Same-machine five-process evidence is
archived under
`Specs/TestRuns/4cb089111a23/FileOpsS3/2026-08-31_r0b_exact_generation_baseline_release/` and
`Specs/TestRuns/4cb089111a23/FileOpsS3/2026-08-31_r0b_exact_generation_candidate_release/`.
Baseline p95 was 6,334 us with zero conditional requests; candidate p95 was 6,613 us with all 4,096
requests conditioned, a 279-us / 4.40% increase within the 56,334-us safety-stabilization ceiling.
This fake-graph result protects bounded CPU/memory/request shape and makes no live-AWS latency or
throughput claim.

Deterministic test-enabled validation uses the S3 fake graph, exact request counts, and fixed metadata latency. The protected 16-object no-conflict interactive scenario requires one refresh per object; the same-binary legacy comparison requires two. Retry, pre-authorized overwrite, no-callback fail-closed behavior, stacked ancestors, byte preservation, and COPY source preservation are part of the same provider contract. Five-sample Release evidence is archived at `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-21_111500_s3_directory_probe/`: refreshes `32 -> 16`, probes `64 -> 32`, and median fixed-latency probe time `1,011,056 us -> 504,449 us` (-50.11%). This is deterministic request-count evidence, not a live S3 throughput claim.

Direct-publication validation covers small, forced-multipart, and relay paths for COPY and MOVE; an endpoint that
explicitly rejects relay conditions; and late-conflict Skip/Overwrite decisions. The protected 16-object
no-conflict COPY uses fixed publication latency and a same-binary unsafe switch: both baseline and candidate issue
exactly 16 publication requests, while the conditional subset changes `0 -> 16`. The archived test-enabled Release
run at `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_1740_s3_direct_publication_safe/` records
`251,041 us -> 253,786 us` (+2,745 us, 1.09%), within the 30% timing-noise gate. This is deterministic publication
request-shape evidence, not a live endpoint throughput claim.

Exact-source validation covers source replacement before small copy, between multipart parts, before relay read,
and before MOVE deletion; each path covers COPY/MOVE, versioned/unversioned behavior, and same-size/changed-size
replacement. Cancellation and later conditional-delete failure cover rollback. The older test-enabled Release matrix
was 131 / 0 for each of five repeats. Its same-binary 16-object scenario changed identity probes and conditional reads
from `0 / 0` to `16 / 16`; median duration was `536,572 us -> 786,180 us` (+46.52%), within the 175% fixed-latency
gate. That bounded evidence is archived at
`Specs/TestRuns/7d3a1247382a/FileOps/2026-08-04_1024_s3_source_revision_pinning/`; its claim is limited to source
revision pinning and conditional reads/deletes, not immutable destination or delete-marker rollback identity.

Immutable rollback-identity validation extends the matrix to 139 / 0 for each of five test-enabled Release repeats.
It asserts exact version sets for source-history hiding, source and destination replacement, cancellation, and later
delete failure. Protected runs capture `16 / 16` final publication identities and the versioned MOVE scenario captures
`1 / 1` source delete-marker identity. The same-binary source-pinning pair remains within the 175% gate at
`527,169 us -> 775,035 us` (+46.98%); the candidate median is 1.42% below the older archive's candidate median.
Bounded evidence is archived at
`Specs/TestRuns/4cb089111a23/FileOps/20260809_105500_s3_immutable_move_identities_release/`. Both archives are
deterministic request-shape and fake metadata-latency evidence, not live endpoint throughput claims.

## AWS runtime and module lifetime

AWS initialization is a fallible state transition. `FileSystemS3` and ranged-reader creation acquire a runtime reference only after `Aws::InitAPI` succeeds; a named initialization exception leaves a failed, non-unloadable state and the factory returns the failure instead of exposing an unusable object.

Runtime refresh uses the optional plugin quiet-point exports. `RedSalamanderPluginShutdown()` closes new runtime acquisition and schedules pending multipart cleanup. `RedSalamanderPluginCanUnloadNow()` remains false while a filesystem/I/O owner, cleanup item, cleanup callback, initialization/shutdown transition, or AWS runtime reference remains. It becomes true only after the final owner releases and `Aws::ShutdownAPI` returns. Cleanup callbacks carry a module pin through the Windows threadpool return boundary.

Pending multipart aborts are self-converging. A failed abort records a bounded retry
deadline and arms one module-pinned thread-pool timer; the timer resubmits the single
cleanup worker when due. No caller poll or later plugin API call is required to wake the
queue. Timer allocation/submission failure falls back to immediate scheduled work.
`NoSuchUpload`/not-found is terminal success, shutdown stops new producers, and unload
remains denied until worker and timer state plus the pending queue are drained.

The multipart writer exposes `IFileWriterExpectedSize`. The host supplies the known
source size before the first write, and the writer plans against the fixed 5 MiB part
size and S3's 10,000-part limit. A size above the representable maximum returns
`ERROR_FILE_TOO_LARGE` before session creation, allocation, or network activity. Writes
cannot exceed the accepted size and Commit requires the exact expected byte count.
Before acquiring its first assembly buffer, a writer with a known expected size reserves
only `min(partSize, expectedSize - position)`. Four concurrent known-size 4-KiB writers
therefore retain at most 16 KiB of assembly-buffer capacity; unknown-size writers retain
the full-part reservation because their next boundary is not known. The multipart debug
contract measures this exact four-writer gate while the shared payload budget remains four
parts and 256 MiB.
Adaptive part sizing remains a separately reviewed future design; this early fail-closed
boundary prevents hundreds of GiB of doomed transfer work.

Process shutdown still honors `RedSalamanderPluginRetainModuleUntilProcessExit()`: the normal quiet point and AWS shutdown run, while the module/dependency chain remains mapped for OS teardown as an additional defense against late third-party CRT activity.

## Identity-pinned Permanent Delete (C10)

`FileSystemS3` implements `IFileSystemIdentityDelete`. `ResolveDeleteIdentity` resolves and probes
the path and returns the object's version id when the bucket has one, otherwise its ETag (a prefix
returns a directory identity without a revision). `DeleteIfIdentity` probes the live object again
and refuses with `ERROR_REVISION_MISMATCH`, deleting nothing, when the name no longer carries that
identity; otherwise it runs the same conditional delete as `DeleteItem` (R0b per-observation
`If-Match` / version id), so the server itself re-checks the revision. The host pins every selected
root this way while a Permanent Delete prepares and re-checks it after the card's answer.
Witnesses: the final phases of `R0fS3_FakeS3ReadWriteCreateDelete` prove the contract directly
(resolve, replace through the plugin, refuse) and the host path (the pinned task is held before its
post-confirmation re-check while the object is replaced; the task ends `ERROR_REVISION_MISMATCH`
and the new object stays).
