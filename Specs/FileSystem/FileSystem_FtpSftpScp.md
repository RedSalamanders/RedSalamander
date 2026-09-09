# FTP / SFTP / SCP / IMAP Virtual File System Plugins

## Overview

RedSalamander ships a built-in virtual file system implementation in `Plugins/FileSystemCurl/FileSystemCurl.dll` that exposes **four UI-visible file system plugins**:

Every instance implements mandatory capability v2. Copy, native Move, Delete, and Rename remain
advertised `false` until the corresponding endpoint-specific operation and contract tests are
truthful; method-body existence is not admission. Identity, revision, read-back, digest algorithm,
byte scope, and revision binding are reported per path and may conservatively be `false`/unknown.

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

- The non-advertised, explicitly overwrite-authorized direct Copy/Move implementation
  can transfer across endpoints as **download → upload** via a local temporary file.
  Curl temporary files use the Common delete-on-close reservation helper with the
  `rsc` prefix and preserve `FILE_ATTRIBUTE_TEMPORARY |
  FILE_FLAG_SEQUENTIAL_SCAN`; no provider-local `GetTempFileNameW`/reopen helper is
  permitted. Normal host planning does not expose this path while conditional
  publication is unavailable.
- Directory operations support recursion when `FILESYSTEM_FLAG_RECURSIVE` is provided.
- Native `DeleteItem` and `DeleteItems` refuse the resolved server root `/` with
  `ERROR_ACCESS_DENIED` before listing or mutating any descendant, independently of
  recursive flags and configured concurrency. The internal iterative Delete walker
  enforces the same boundary before traversal (including native
  Move source cleanup). Ordinary explicitly selected subdirectories remain deletable;
  a root refusal must not disable the provider or affect neighboring objects.
- Single-item Delete installs the owning `FileSystemOptions` scope before admission
  I/O, just like bulk Delete. Cancellation/deadline control reaches its first listing,
  not only later delete commands. `C0_CurlNativeDeleteGuards` verifies root refusal,
  ordinary selected-subtree deletion, sibling preservation, and stalled-list cancellation
  against the actual native entry points on loopback FTP. This is not evidence that
  all native walkers are iterative or that real FTP/SFTP/SCP endpoints were exercised.
- Directory listing is parsed incrementally (streamed) to avoid buffering the full LIST output in memory, improving large-directory memory behavior.
- Native recursive Delete (including native Move source cleanup) uses an explicit
  ancestry stack and `CurlDirectoryCursor`, not whole-directory/tree entry vectors.
  Root depth is zero; depth 128 is the shared maximum. Each active cursor owns one
  bounded transport chunk and one line (`CURL_MAX_WRITE_SIZE` each), with a reserved
  libcurl paused chunk. Delete batches contain one item at concurrency 1 or at most
  the shared discovery target at higher concurrency; batches drain before descent
  and directory removal. Queue/path pressure drains work rather than imposing a
  limit on aggregate tree size. Retained path text and metadata reservations use
  the shared 16-MiB/8-MiB policy; opaque TLS/socket allocations are not metadata.
  A depth, single-record, or irreducible retained-path limit fails explicitly.
- Native directory Copy/merge uses one iterative `CurlDirectoryCursor` ancestry
  walk for both serial and scheduler-owned file batches. The shared depth-128,
  16-MiB retained-path and 8-MiB frame/batch metadata limits apply; batching uses
  one item serially or the shared discovery target in parallel. Queue pressure
  drains work, never rejects a tree merely for exceeding 4,096 total files.
  Each listing finishes (or fails closed) before a child cursor is opened, so a
  paused parent transport is never held idle across subtree work. File batches
  still drain before a completed parent is popped. Validate a first entry or
  complete empty listing before creating that directory's destination. Invalid first discovery creates no children; later failures preserve
  earlier publications. Root failures retain their exact HRESULT. Cancellation,
  authentication and hard storage/depth limits are terminal even with Continue on
  error. Other child failures may continue independently and finish PARTIAL_COPY.
  Move retains its complete preflight membership/size proof as 64-bit digests of the
  normalized member paths (16 bytes per file with its committed size, 8 per directory,
  in vectors finalized once after preflight), so a tree of N files retains about 16 N
  bytes whatever its path lengths; the `FileOps.Curl.MovePreflight.RetainedDigestBytes`
  row reports that whole-operation figure. Real server quotas remain unqualified.
  The subsequent source-delete pass may delete only names observed in that
  preflight, and it first proves the source root itself is a member: if the two walkers
  ever spelled paths differently, the pass refuses with `ERROR_INVALID_DATA` before any
  deletion instead of silently keeping the whole source as late writers (the digest drops
  one trailing slash, never the root, because the preflight walks a directory spelled with
  it while the delete pass names it without). A file or
  directory created after preflight is retained, the parent `RMD` is refused, and the
  Move reports `ERROR_PARTIAL_COPY`. A late writer whose name collides on 64 bits (and, for
  a file, matches the committed size) would pass as a member; that probability is about
  n^2 / 2^65 and is the accepted residual of the bound.
- A cursor emits only validated Unix/DOS LIST entries. Empty lines, a valid Unix
  `total` header, and exact dot entries are ignored; malformed/unknown records,
  invalid UTF-8, embedded NUL, and separator-bearing child names fail, never become
  a names-only destructive fallback. Unix `c`/`b`/`s`/`p`/`D` rows are device
  entries (`FILE_ATTRIBUTE_DEVICE`), never DOS phantoms named from leftover
  columns. DOS rows require a valid date/time token pair in addition to a
  complete integer size or `<DIR>`. Symlink targets are not traversed; a regular
  file containing ` -> ` retains its literal name. End-of-list success requires a
  successful completed transfer and consumption of every valid row. Emitted rows
  are never replayed after a transport failure.
- Repeated row parsing may reuse filename allocation through the canonical strict
  UTF conversion helper, but must reset all prior row metadata. DOS size tokens
  must parse as a complete unsigned integer; numeric prefixes such as `4junk`
  or `4.5` are malformed, not four-byte files. Allocation reuse does not relax
  complete-listing admission, timestamp policy, or final-reply validation.
- Cursor framing scans bounded chunks for LF and copies each contiguous fragment
  as a block, preserving the same one-chunk/one-line reservation. Fragmentation
  inside UTF-8 or CRLF does not change parsed names, sizes, timestamps or errors.
  LF terminates but is not part of the record limit; a preceding CR occupies one
  byte of that limit. A record exactly at `CURL_MAX_WRITE_SIZE` is accepted when
  otherwise valid, including at successful EOF; one extra non-LF byte fails with
  `ERROR_BUFFER_OVERFLOW`. The test-only `ParseChunksForSelfTest` seam exercises
  the production callback and chunk consumer without I/O; it is not evidence of
  successful transport termination, cancellation or remote mutations.
- Connection-failure diagnostics for the native cursor, targeted size probe and
  namespace quote dispatch capture Curl/OS codes and numeric connection metadata
  before resetting their borrowed handle. They retain no URLs, profiles, credentials
  or reply/error text, preserve the existing HRESULT and retry policy, and grant
  no mutation/replay authority. An unavailable OS error remains explicitly unavailable.
- Cursor I/O is pumped on its owning worker and checked for cancellation/deadline
  between bounded polls. Walkers do not open a child cursor while the parent
  listing is still paused; finishing the parent listing releases its pool borrow
  before descent. An active `Next` uses the provider's existing watchdog bound
  while the operation's absolute deadline remains authoritative. On failure,
  undispatched work in that directory is abandoned; already drained mutations
  retain their truth.
  Continue-on-error can visit independent siblings but does not remove a parent
  known to contain failed children. A failed selected-root listing preserves its
  exact error; continuation of other selected roots belongs to the bulk caller.
  Cursor teardown detaches its easy handle before
  releasing the multi handle and returns the borrowed handle only afterward.
- Namespace command lists (`MKD`, `DELE`, `RMD`, `RNFR`/`RNTO` and their SFTP
  equivalents) use literal UTF-8 paths from the captured base path, not URL-escaped
  path text. FTP keeps the complete pathname after the verb; SFTP/SCP namespace
  operands use libcurl double-quote/backslash escaping, not shell escaping. Spaces,
  quotes, Unicode and literal percent sequences keep their identity. CR/LF, embedded
  NUL and invalid Unicode are rejected before a command is constructed.
- Namespace command lists (`MKD`, `DELE`, `RMD`, `RNFR`/`RNTO` and their SFTP
  equivalents) execute once. Transport failure after dispatch cannot prove no
  commit, so the quote helper must not replay the list against a potentially new
  occupant. Bounded read/probe retries are separate and do not authorize mutation.
  Quote-only requests suppress the unrelated body/directory transfer using
  [`CURLOPT_NOBODY`](https://curl.se/libcurl/c/CURLOPT_NOBODY.html); a subsequent
  listing is not part of the mutation's success contract. The supplied commands
  run before any transfer per [`CURLOPT_QUOTE`](https://curl.se/libcurl/c/CURLOPT_QUOTE.html).
  A lost Delete/Rename reply remains uncertain, with no safe-Retry proof inferred
  from a timeout, disconnect, or a path now naming a replacement.
- `C0_CurlCommittedMutationResponseLost` exercises actual single/bulk Delete/Rename
  against loopback FTP: commit, atomically plant a different source generation,
  drop acknowledgement, then verify one mutation request, exact replacement/sibling
  bytes, and one uncertain completion. Successful controls verify that ordinary
  operations still complete and no incidental NLST is requested.
  `C0_CurlHostCommittedDeleteResponseLost` additionally checks the host's immutable
  result, localized uncertain summary, and absence of Retry/conflict prompts.
  This does not qualify real SSH servers or every multi-command partial-tree failure.
- Destructive completion never infers no-commit from a final HRESULT. A later
  listing/access refusal can follow an earlier child or rollback mutation. Native
  Delete retains one selected-item mutation-attempt flag through recursive and
  parallel child work; it is set before a remote mutation call, never cleared, and
  read only after child workers drain. A successfully observed source with no
  mutation attempt may produce retained-source/no-commit proof. Other failed
  completions remain unknown unless their caller supplies equivalent explicit
  evidence. Native Move/Rename collision preflight retains that proof, but clears
  it before rollback rename, final rename, or Copy-based staging/publication.
  The host must still revalidate authority before any explicit Retry.
- `C0_CurlPartialTreeFailure` covers single/bulk partial deletion at concurrency 1/4,
  untouched selected-root refusal, the shared completion proof boundary, and native
  Move/Rename collision controls. `C0_CurlHostPartialTreeFailure` checks immutable
  uncertain outcome, localized unknown summary, preserved remaining/sibling bytes,
  and no Retry/conflict prompt after an earlier child was deleted. These fixtures
  do not qualify the remaining iterative-walker migration or live SSH endpoints.
- Copy/Move/Delete/Rename progress joins the ABI `operationControl`/deadline checkpoint before
  each bounded libcurl progress/retry unit. In addition (R0f-Curl), every libcurl transfer started
  under a File Operations mutation or listing call, control commands (`SIZE`, `LIST`, `DELE`, `MKD`,
  `RNFR`/`RNTO`, SFTP stat/rename/remove) included, carries a progress callback that polls the
  owning call's operation control. The streaming `IFileReader` returned by `CreateFileReader` (the
  Copy-from-FTP source path) has no options of its own; since R0f-Curl-OR1 (2026-09-04) it implements
  `IFileReaderOperationControl`, the host hands it the task's control right after creation, and its
  progress callback polls that control on every libcurl call, so a cancel ends an in-flight `Read`
  with `ERROR_CANCELLED` within the poll period. Every transfer therefore polls the
  operation control: the mutation entry points and the shared scheduler keep the call's options on
  the current thread, `ApplyCommonCurlOptions` installs the callback, and libcurl invokes it at
  least once per second even while a command waits for the server. A canceled or deadline-expired
  call therefore returns within about a second with the host's own verdict (`ERROR_CANCELLED` or
  `ERROR_TIMEOUT`), never a fabricated transport error. The FTP/SFTP/SCP route advertises
  `routeClass:"providerWatchdog"`, `deadline:true`, `abort:false` (there is no out-of-band abort),
  and `providerWatchdogTimeoutMs` = the larger of the connect timeout and the low-speed abort
  (60 s by default, or the configured operation timeout): the provider-owned bound after which a
  server that stops answering returns on its own. The typed facts and the capability JSON
  advertise `copy`, `move`, `nativeMove`, `delete`, `rename`, `createDirectory`, `read`, and
  `write` for FTP/SFTP/SCP; File Operations admits same- and cross-provider mutation on these
  routes. IMAP mailboxes stay read-only and `uncontained` (no mutation is admitted, so no
  containment claim is needed). The plugin unload drain is a module-lifetime boundary, not a host
  cancellation deadline. FTP, SFTP, and SCP are full file-manager destinations.
- Writer publication (R0f-Curl): the temp-file writer stages the upload as a unique remote
  sibling, verifies its exact size, and renames it onto the requested path in one server-side
  step, so the plugin advertises `IFileSystemAtomicWriter` and the host bridge publishes new
  names through it. A no-overwrite Commit probes the destination first and fails with
  `ERROR_FILE_EXISTS` when it is present (the host then raises its ordinary overwrite conflict);
  with overwrite granted the existing object is moved to a rollback sibling before the rename.
  FTP `RNTO` and libcurl's SFTP rename expose no atomic no-replace primitive: a creator that
  appears between the probe and the rename is overwritten. That residual window is the same one
  every FTP client has and is documented rather than used as a reason to refuse the destination.
  Direct same-provider Copy/Move/Rename without `ALLOW_OVERWRITE` follow the same rule: the
  destination is probed, `ERROR_FILE_EXISTS` is reported when it is present, and the staged
  publication otherwise renames onto the requested name. Cross-provider replacement through the
  host bridge (R3-1): the writer carries the occupant the user saw (`IFileWriterExpectedReplacement`,
  last-write time from the same listing dialect); before staging it probes the destination again
  and refuses with `ERROR_REVISION_MISMATCH` when the occupant vanished or its last-write time
  changed, uploading nothing. The probe-to-rename window of the transport remains the documented
  residual.
- Content proof (R3-2): FTP, SFTP, and SCP report no digest of stored content, so the route keeps
  `verification.providerProof: "none"`. Verify On stays `Unavailable` on these destinations and a
  Move into them remains Copy-only with the source kept, until a transport with a checksum command
  is adopted.
- Delete (R0f-Curl): FTP/SFTP/SCP bind no host objects, so a permanent Delete task runs the
  provider's native `DeleteItem` under the host's native authority (`FileSystem_FileOperations.md`,
  Delete and Recycle): admission records no ingress snapshot, and the plugin's per-item
  `FileSystemItemMutationResult` (emitted for `DELETE`, `MOVE`, and `RENAME` completions: success
  is known and committed; failed no-commit requires explicit retained-source and
  no-mutation evidence, never an error-code guess; absent evidence carries no record
  and stays `Indeterminate`) decides the outcome
  the host shows. Recycle is not offered on these routes.
- For a **single** top-level directory copy/move or delete, contained file work SHOULD still parallelize (within-folder parallelism) subject to the effective concurrency limits.
- Parallel work borrows one libcurl easy handle per thread from the plugin's handle pool; connection reuse is per handle and therefore thread-confined. The process-wide libcurl share covers DNS and TLS sessions only, never the connection pool (R0f-Curl-OR2: sharing the pool across threads crashed inside libcurl's pooled-connection matching during parallel copies).
- Returning an easy handle resets all per-call options before idle publication or
  shutdown/overflow disposal. Idle handles must not retain caller callback/data
  pointers: [libcurl cleanup may invoke FTP/IMAP callbacks](https://curl.se/libcurl/c/curl_easy_cleanup.html).
  [Reset preserves connection, DNS, and TLS-session reuse](https://curl.se/libcurl/c/curl_easy_reset.html).
  Directory-size feedback is terminal after the final snapshot, including during
  transport teardown.
- SCP has protocol limitations; directory listing and command-style operations require the server to support SFTP over SSH.

### Directory-size estimates

`GetDirectorySize` uses the same bounded FTP/SSH cursor as native Delete, including
the parent listing needed to identify a file or directory root. Root lookup must
finish the listing, reject duplicate selected names and malformed records, and
match exact case-sensitive provider names. A parent listing is finished before a child directory is opened, so a paused LIST
does not hold the idle FTP session across a subtree. Sibling directories observed
during that listing are counted before descent; a later global callback abort
inside the first child does not size later siblings' files. Tree traversal is
iterative, with at most 129 active frames including the depth-zero root; admitted
frame/cursor storage uses the shared 16-MiB path and 8-MiB metadata ceilings. A wide tree streams through
the cursor and has no aggregate 4,096-file cutoff. Refused structure/depth/storage
and arithmetic overflow remain explicit failures, never complete totals.

Known numeric listing sizes are read-only estimates, not mutation commitments.
Unknown file sizes use the existing targeted size probe. An unavailable selected
file-root size returns `ERROR_NOT_SUPPORTED`; unavailable child sizes preserve the
observed file count and known byte lower bound with `ERROR_PARTIAL_COPY`. Eligible
denied/vanished/transiently unavailable descendants do not stop later siblings.
Root failures, malformed structure, authentication, callback errors, and cancellation
are global and retain their primary error even if the final callback also fails.
Accumulated skipped-child partial totals do not mask a final callback error or
cancellation; only a prior global failure takes precedence over final feedback.

Size callback control is installed before root resolution and network probing.
Progress/cancel pairs are coalesced at 100 newly observed entries or 200 ms; the
initial and final snapshots are unconditional. Transport pumps poll this same
sticky callback state. No callback is concurrent with another or survives return.
`FileOps.Curl.DirectorySize.Traversal` records duration, peak admitted frames and
final observed files; `.Retained` records peak admitted frame/cursor path storage
and metadata reservations. These are not total process-memory measurements: caller
inputs, resolution/settings state, temporary parsing/joining, and opaque Curl/TLS/
socket allocations are separate. One live transport per ancestry frame still needs
real FTP/SSH session-limit qualification.

### Move-mode admission

- Nonempty single/bulk operation requests validate `FileSystemOptions.moveMode`
  before endpoint resolution, provider I/O or item callbacks. Unknown values,
  and any nondefault value on Copy/Delete/Rename, return `E_INVALIDARG`.
  Existing zero-item bulk no-op semantics are unchanged.
- `FILESYSTEM_MOVE_NATIVE_ONLY` admits only the existing same-endpoint server
  rename route, for files or directory subtrees. Capture that restriction once;
  mutable progress options cannot widen it. A cross-endpoint request returns
  `ERROR_NOT_SUPPORTED` before payload download, destination mutation or source
  deletion. It must never fall through to the legacy copy/delete relay.
- If read-only admission observed the source and no mutation was attempted, the
  completion retains known no-commit/source-present truth. Failed source lookup
  does not fabricate that observation. Mixed bulk requests with Continue on error
  keep independent item outcomes and aggregate partial status: a supported native
  item can succeed while a cross-endpoint item is refused with its source retained.
- Default Move keeps its existing routes and source-size rules below. Native-only
  admission is not proof of conditional cleanup or an overwrite race repair.
  `C0_CurlMovePreflightTruth` covers all operation APIs, directory descendants and
  prefix siblings, plus mixed native-success/refusal with configured concurrency 1/4.

### Native source-size commitments and read integrity

- Provider-native FTP/SFTP/SCP Copy/Move content-transfer routes and `IFileSystemIO::CreateFileReader`
  MUST use the same per-file source-size commitment. The provider first attempts
  a targeted FTP `SIZE` or SFTP/SCP stat; an exact targeted result supersedes
  parent-directory listing metadata, including a stale numeric listing size.
- Only explicit targeted-primitive unavailability may fall back to an exact
  numeric listing size or the caller's unknown-size policy. For FTP, replies
  `500`, `502`, and `504` classify `SIZE` as unavailable. Authentication
  failure (`530`) maps to `ERROR_LOGON_FAILURE`, disappearance (`550`) maps to
  `ERROR_FILE_NOT_FOUND`, and cancellation, other `4xx`/`5xx`, transport, and
  protocol failures MUST propagate instead of being hidden by listing metadata.
- FTP control-reply inspection for this classification is status-only and
  scoped to the current targeted probe. It may retain numeric reply classes but
  MUST NOT retain or log commands, payloads, credentials, or the control stream.
  A successful exact size is authoritative even if an earlier non-fatal control
  reply was observed during the same libcurl operation.
- A known commitment, including zero, is not satisfied merely because libcurl
  reports transport success. The download MUST observe the data-stream
  terminator and exact byte count. A short body fails with
  `ERROR_PARTIAL_COPY`; an overlong body fails with `ERROR_INVALID_DATA`; no
  destination may be published and no Move source may be deleted after either
  mismatch. A zero-byte commitment still performs and validates the transfer so
  a contradictory nonempty body cannot be reported as successful EOF.
- Unknown-size native Copy may materialize the complete download into its local
  delete-on-close stage and use that exact staged byte count for upload and
  publication verification because Copy retains the source. Native relay Move MUST
  fail with `ERROR_PARTIAL_COPY` when any source-file commitment remains unknown,
  before file-content download, destination upload/tree creation, or source
  deletion. Recursive relay Move MUST resolve every descendant file commitment in a
  read-only preflight before the first destination mutation.
- Native recursive Move preflight uses the strict incremental FTP/SSH cursor and
  an explicit ancestry stack, not buffered listing fallback or C++ recursion.
  Malformed/out-of-root/overlong records and depth beyond 128 fail before payload
  download or any destination/source mutation. Every preflight listing/probe error
  is global: there is no partial preflight success. At most 129 cursor frames are
  admitted, charging their reserved slots/metadata and retained path/URL storage
  against the shared 8/16-MiB ceilings. Wide trees stream without an aggregate
  4,096-file cutoff. Per-file targeted commitment resolution is retained unchanged; the
  caller's commitment map and membership set are the digest containers above, whose
  aggregate memory is the reported `RetainedDigestBytes` row. Real transport session
  quotas remain unqualified. Frame-reservation metrics must not be presented as a
  bound on total preflight memory; the digest row is that bound.
- A known-size `IFileReader` validates the remaining committed range after every
  seek. FTP sends `REST <offset>` as a pre-transfer FTP command while continuing
  to ignore an advertised body length for terminator validation; SFTP/SCP use
  libcurl's resume offset. Offsets beyond the commitment fail, offsets exactly at
  the committed end still validate a zero-byte range, and a failed transfer
  generation MUST be discarded before a same-position restart exposes bytes
  from a new generation.
- Reader teardown MUST publish its stopping predicate while holding the same
  mutex used by the readable/writable condition-variable waiters, then request
  stop and notify before joining the worker. This ordering prevents a worker
  from observing the old predicate and sleeping after teardown's only
  notification. Every readable wait uses a predicate that observes stopping,
  EOF, worker failure, and the required buffered-data/committed-end condition;
  every full-buffer writer wait observes stopping, generation/stop changes, and
  newly available capacity. Stopping is terminal for `Read` and returns
  `ERROR_CANCELLED`. The Curl contract surface includes a 32-iteration loopback
  reader-shutdown stress guard plus behavioral cases that stop a consumer blocked
  in `Read` on an empty buffer and a producer blocked on a full ring buffer.
- The deterministic contract surface is
  `.build\x64\<Configuration>\PluginContractTests.exe --curl-selftests`.
  It covers direct readers plus single, batch, and recursive Copy/Move for exact,
  short, overlong, zero, unknown, stale-listing, authentication-failure, and
  disappearance cases; exact FTP `REST` behavior; pre-mutation command counts;
  and the plugin unload quiet point. Performance-sensitive changes to this path
  emit and archive `FileOps.Curl.SourceCommitmentGuard` under `Specs/TestRuns/`.
- Five-repeat test-enabled Release evidence is archived at
  `Specs/TestRuns/4cb089111a23/FileOps/20260816_085809_curl_source_commitment_release/`.
  Every run passed 199/199 assertions and the unload quiet point. The paired
  known-short Copy/Move guard took 4,000–6,000 us (6,000-us median), used exactly
  22 source commands, delivered 4,098 deliberately short bytes, and returned
  `ERROR_PARTIAL_COPY` before destination upload or source deletion. This is a
  same-machine loopback safety-cost baseline with millisecond timing granularity,
  not a live-endpoint throughput result; no apples-to-apples historical metric
  exists, and the transactional-writer archive is not used as a comparison.

### Conditional publication and identity-safe recovery

- FTP `RNFR`/`RNTO` does not expose an atomic publish-if-absent operation or an
  immutable object/revision token through this plugin. With curl `8.21.0`, the
  SFTP quote `rename` path supplies libssh2 `OVERWRITE | ATOMIC | NATIVE` and
  does not expose no-replace control to the plugin; SCP uses the same current
  non-FTP command surface. None of the three surfaces exposes conditional
  delete/restore or identity-safe recursive rollback.
- Before R0f-Curl the three surfaces advertised `operations.copy`, `operations.move`, and
  `operations.rename` as `false`. Since R0f-Curl (2026-09-03) they advertise all three as `true`
  (IMAP stays `false`), with `caseOnlyRename = notApplicable`; the containment, staged publication,
  conditional replacement (R3-1), and Delete rules of the R0f-Curl section govern them. Version-1
  capability booleans are sufficient; no COM or JSON schema extension is used. The host MUST honor
  the typed route facts, including the general File Operations Rename gate and Batch Rename, before
  task creation or endpoint work.
- Direct singular/bulk Copy, Move, and Rename validate the options header before
  endpoint work; malformed `FileSystemOptions::sizeBytes` returns `E_INVALIDARG`.
  Since R0f-Curl, ordinary no-overwrite calls are supported through the destination
  probe described above, returning `ERROR_FILE_EXISTS` for an observed occupant.
  The probe is not atomic no-replace authority: the documented probe-to-rename
  race remains. Do not restore the superseded blanket `ERROR_NOT_SUPPORTED` rule
  merely because the provider lacks an atomic no-replace primitive.
- Explicitly overwrite-authorized calls may move an existing final to a
  randomized rollback sibling and publish the new object. Randomized path text is
  not immutable identity: after successful publication, the rollback sibling is
  retained as cleanup debt rather than deleted by pathname.
- If publication fails after a rollback sibling exists, recovery preserves the
  current final, rollback sibling, and any remaining staging object and returns
  `ERROR_PARTIAL_COPY`; it never deletes the current final or renames the backup
  over a path that a concurrent actor may own. If Move source deletion fails
  after publication, the source, published/current destination, and rollback
  sibling are preserved and the result is `ERROR_PARTIAL_COPY`. Failed recursive
  Copy/Move preserves the uncertain destination tree and independently added
  children; it never recursively deletes that tree by pathname.
- Publication/recovery state is represented internally as one structured result
  carrying primary-mutation, cleanup, and source-deletion HRESULTs; whether the
  primary publication committed; immutable rollback eligibility; preserved-artifact
  disposition; and a content-free cleanup-debt kind mask and saturating count.
  Recursive and bulk workers merge that state thread-safely without collapsing the
  first relevant failure or counting one retained artifact more than once.
- A failed primary mutation retains its primary HRESULT even when operation-owned
  staging cleanup also fails. A source-deletion failure after publication returns
  `ERROR_PARTIAL_COPY`. A committed destination with cleanup-only debt remains
  `S_OK`: its item callback, committed size, destination bytes, and normal directory
  notification all describe the successful publication, while the retained artifact
  is reported separately. Upload, verification, prepare, and promotion failures MUST
  never be converted to committed success.
- Safe pre-publication cleanup may delete only the operation-owned staging sibling.
  If that deletion fails, the primary failure is preserved and the retained staging
  object becomes observable cleanup debt. Rollback siblings, uncertain finals, and
  preserved directory trees remain ineligible for pathname cleanup without immutable
  identity, including after a successful publication.
- One public Copy/Move/Rename operation, or the first writer `Commit()`, emits at most
  one aggregate cleanup-debt observation: one content-free diagnostic, one
  `FileOps.Curl.CleanupDebt` metric, and one localized application-scoped modeless
  warning when `IHostAlerts` is available. The warning contains only the retained-item
  count and manual-review guidance; the diagnostic/metric/warning MUST NOT include
  URLs, paths, users, credentials, payloads, or randomized sibling names. Repeating a
  successful writer `Commit()` is idempotent and MUST NOT upload, notify, warn, or
  alert again.
- The deterministic loopback matrix injects a late destination before RNTO, a
  replacement after publication, a writer final before failed promotion, a source
  delete failure, and a child after destination-directory creation. It covers
  single/bulk Copy, Move, and Rename plus file/tree recovery. Performance capture
  emits `FileOps.Curl.ConditionalPublicationGuard` and
  `FileOps.Curl.IdentitySafeRecoveryGuard`.
- Five-repeat test-enabled Release evidence is archived at
  `Specs/TestRuns/4cb089111a23/FileOps/20260816_123000_curl_conditional_publication_release/`.
  Every process passed 211/211 assertions and the unload quiet point. The six
  no-overwrite shapes issued zero endpoint commands and zero race injections,
  returning `ERROR_NOT_SUPPORTED` before mutation. The three combined recovery
  scenarios took 18,000–22,000 us (20,000-us median), used exactly 103
  deterministic loopback commands, preserved three final/backup/tree artifacts,
  and returned `ERROR_PARTIAL_COPY`. This same-machine archive establishes a
  bounded safety-cost baseline; it is not a live-endpoint throughput result and
  makes no speedup claim.
- Structured cleanup-debt Release evidence is archived at
  `Specs/TestRuns/4cb089111a23/FileOps/20260816_202500_curl_cleanup_debt_release/`.
  Five independent processes passed 221/221 assertions and the unload quiet
  point. `FileOps.Curl.CleanupDebtGuard` covered six committed overwrite shapes
  in 59,000–62,000 us (60,000-us median), with exactly 320 deterministic
  loopback commands and nine retained rollback items per run. The aggregate
  `FileOps.Curl.CleanupDebt` stream had 19 content-free observations and 22
  retained-item counts per process. This is a same-machine bounded correctness
  and safety-cost baseline; no apples-to-apples historical cleanup-debt metric
  exists, and no live-endpoint or speedup claim is made.

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
- A mixed-connection delete batch admits work independently per resolved connection key.
  The outer scheduler MUST NOT collapse the whole batch to the smallest endpoint cap;
  each task acquires the existing keyed limiter for its own connection, so a serial
  endpoint cannot throttle unrelated profiles while every profile still respects its
  configured `deleteMaxConcurrency`.
- File writers retain a strong COM reference to their creating `FileSystemCurl` until
  Commit/cancel/destruction completes. Releasing the factory-facing filesystem pointer
  cannot invalidate a returned writer or allow plugin unload while that writer still
  needs shared state.
- FTP/SFTP/SCP writers stage bytes in a local delete-on-close file. `CreateFileWriter`
  and `Write` never probe, delete, or upload to the final remote path. `Commit` uploads
  to a unique remote sibling, proves its exact size, moves any overwrite target to a
  rollback sibling, and only then promotes the staged object. A pre-publication
  upload/size failure may remove the operation-owned staging sibling. Once a rollback
  sibling exists, failed promotion preserves the current final and rollback artifact
  and returns `ERROR_PARTIAL_COPY`; successful promotion retains the rollback artifact
  as cleanup debt because deletion by pathname is not identity safe. Releasing a
  writer before `Commit` performs no endpoint request.
- The current libcurl FTP/SFTP rename surfaces do not provide a portable atomic
  no-replace primitive. Curl writer requests without `FILESYSTEM_FLAG_ALLOW_OVERWRITE`
  therefore fail with `ERROR_NOT_SUPPORTED` before endpoint access instead of using a
  racy existence-probe/rename sequence. IMAP writers also fail closed because message
  append/delete has no recoverable rename transaction. The same no-overwrite rule now
  applies to provider-native FTP/SFTP/SCP Copy/Move/Rename entry points.
- For the same reason, FTP/SFTP/SCP advertise empty cross-filesystem import Copy
  and Move lists. The host rejects local/foreign-to-Curl bridge planning before
  task creation while retaining Curl-to-other-provider export. Same-provider
  Copy/Move/Rename are also disabled until a separately reviewed endpoint primitive
  can prove atomic fail-if-exists publication; an existence probe followed by
  upload or rename is not sufficient.
- `ALLOW_REPLACE_READONLY` has no standalone meaning. `CreateFileWriter` returns
  `E_INVALIDARG` when it is present without `ALLOW_OVERWRITE`, before the generic
  no-overwrite/unsupported result and before endpoint access.
- Staged upload publication requires an exact proved remote size. A targeted FTP `SIZE`
  or SFTP stat is preferred; a listing is acceptable only when it carries an exact
  numeric size. Unknown size is not zero and fails closed: the staged sibling is removed,
  the existing destination and MOVE source are preserved, and no transport-success
  assumption substitutes for durable remote-byte proof.
- Writer validation emits `FileOps.Curl.Writer.CommitUs`, `.StagedBytes`,
  `.RequestCount`, `.CleanupAttemptCount`, and `.CancellationCount`. Request count is
  the bounded logical staging/upload/verification/promotion count; writer metric
  details are fixed content-free labels rather than plugin paths. Aggregate retained
  recovery items emit `FileOps.Curl.CleanupDebt`, while the six committed overwrite
  shapes emit `FileOps.Curl.CleanupDebtGuard`; the loopback FTP evidence also records
  the exact protocol command inventory.
- Historical five-repeat writer evidence is archived at
  `Specs/TestRuns/4cb089111a23/FileOps/20260809_111300_curl_transactional_writer_release/`:
  transactional overwrite is 27 commands with a 5,000-us median, short-success cleanup
  is 15 commands / 3,000 us, promotion rollback is 35 commands / 5,000 us, and
  abandonment is zero endpoint commands. That archive predates identity-safe backup
  preservation and is retained only as historical timing evidence; its cleanup/restore
  semantics are not the current contract. These are bounded fake-loopback results, not
  live FTP/SFTP throughput claims.
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


## Permanent Delete by name (C10)

FTP, SFTP and SCP name no object identity, so the plugin does not implement
`IFileSystemIdentityDelete` and a Permanent Delete through File Operations deletes whatever the
name refers to when the mutation runs. The host's card states this before the answer
(`IDS_FILEOPS_CONSENT_PERMANENT_DELETE_BY_NAME`). Witness: the last phase of
`R0fCurl_FakeFtpReadWriteCreateDelete` starts a Permanent Delete with the card left unanswered,
requires the by-name sentence in the card's consent detail, cancels, and requires the object intact.
