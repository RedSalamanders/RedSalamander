# Cross-File-System Bridge Specification

## Purpose

The bridge streams a qualified source object into a qualified destination when no tested native
provider strategy can fulfill the request. It implements Copy, managed Move, or Copy-only according
to the path-scoped capability-v2 and binding contract in
`Specs/Plugins/Plugins_VirtualFileSystem.md`.

The bridge never follows a link, treats a hash as identity, or performs pathname source cleanup.

## Admission

Both endpoints must return valid `IFileSystemPathCapabilities2` documents for the concrete paths
and operation. Copy requires a readable source and writable destination with compatible transfer
policy. Managed Move additionally requires:

- no-follow `IFileSystemObjectBinding` for the source;
- exact object/revision snapshot and a reader opened from that bound object;
- destination exclusive stage ownership and conditional publication, or (R3-2) a destination
  route advertising `FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST` whose atomic-final writer stages
  privately, publishes conditionally, and proves the published content;
- exact source `DeleteIfUnchanged`;
- every required consent and requested verification.

If any destructive prerequisite is absent, the bridge selects Copy-only before source cleanup and
reports `Copied; source kept`.

Strategy admission and per-item cleanup are separate safety gates. Admission may select Managed
only when the concrete Move pair, Copy pair, capability claims, and the source binding interface
are present, together with either the destination binding interface or the writer proof route
(R3-2; on that route only regular files carry a proof, so links and directories retain their
source per item). Each item then binds its exact source no-follow with read and delete authority. Local
Managed Move retains that handle with read-only sharing so a concurrent writer, rename, or delete
cannot enter between the bound read and conditional cleanup. If this per-item authority cannot be
acquired, the item is copied through the already-qualified Copy route and the source is retained.
The downgrade is terminal and does not re-enter admission or try an unbound Delete.

For Local Move, equal volume-qualified endpoints are necessary but not sufficient for Native.
Admission probes only the selected top-level shape. Regular files may use same-volume Native;
directories and link objects use Managed because Local does not claim Native semantic-link
equivalence. A regular directory targeting an existing regular directory remains Managed folder
merge so child collisions stay owned by the central conflict engine. No recursive qualification
walk is permitted.

## One-file flow

1. Bind the source no-follow and validate a regular-file/object kind.
2. Bind/classify the destination root, ancestors, and existing final object no-follow.
3. Create a cryptographically unpredictable sibling stage exclusively and retain the returned
   ownership token.
4. Open the source reader from the binding and stream into the stage writer with bounded buffers.
5. Reject provider success with null outputs, oversized read/write counts, early EOF, byte-count
   mismatch, failed Commit, or unknown stage outcome.
6. Publish through the stage token. Absence requires `expectedDestination == nullptr`; replacement
   requires the exact rebound destination token.
7. If verification was selected, verify the exact published object with BLAKE3 or an equivalent
   bound provider proof; on a route without object binding whose writer proves content
   (`IFileWriterContentProof`, R3-2), compare the digest the provider reports for the published
   object with the streamed bytes.
8. Apply metadata and any required deferred-consent decision.
9. For managed Move only, call `DeleteIfUnchanged` on the same exact bound source retained from the
   read. That provider boundary performs the identity/revision revalidation and mutation; the host
   does not first reopen or compare the source pathname. A known non-commit retains the source and
   reports `Copied; source kept`; changed or otherwise unavailable proof also retains it. Unknown
   mutation outcome is indeterminate and is not retried.

One transaction record (R3-3). The bridge keeps every fact of this flow for a regular file in one
`PublicationTransaction` record (admission grants and expectations, the source authority and
metadata snapshot, the Move cleanup record, the reader and the sizes the transfer is held to, the
writer route, the stage and writer handles with their proofs, the pump counters, and the
post-commit proofs) with an explicit state: `Admitted`, `SourceBound`, `Staged`, `Written`,
`Committed`, `Published`, `Verified`. The flow runs as ten phase methods on that record
(`AdmitDestination`, `BindSource`, `RouteWriter`, `PrepareStage`, `TransferPreContentMetadata`,
`Pump`, `CommitAndProve`, `Publish`, `VerifyPublished`, `FinishCleanup`); each early return is the
status the item reports, and the receipt (publication, verification, source disposition, failure
phase) follows from the state reached. The three payloads share the one owner of the exact
`PublishAs` outcome (`PublishOwnedStageAs`): the file payload publishes its owned stage, the link
payload records a partial post-publication cleanup as a stage-commit failure, and the directory
payload publishes without payload diagnostics. Provider-native atomic operations (Native Move,
provider copy) stay adapters outside this record.

The bound source plus retry count forms one item-sized cleanup record. A retryable known non-commit
enters the shared conflict surface. `Retry` calls only `DeleteIfUnchanged` again while retaining the
same authority; it never repeats stage creation, transfer, publication, verification, or metadata.
`Skip` makes source retention terminal. Cancel releases the record without another delete. The
record is released only after Removed, retained-source Skip, Cancel, or indeterminate/contract-fail
classification, so an outer item retry cannot lose or accidentally reconstruct its authority.
Sharing, access denied, path too long, and disk full are source-cleanup buckets. Each permits at most
one Retry of `DeleteIfUnchanged`; Skip retains the source. Destination metadata is not consulted to
classify those post-publication failures.

There is no mandatory destination reread and no mandatory digest when verification is off.
Verification is independent of delete authority. Publication or cleanup with unknown outcome is
not retried until reconciliation proves identity.

Verify On deliberately routes every non-Native copied regular payload through this exact-authority
bridge, including same-endpoint Local Copy that would otherwise call provider `CopyItem`. This is one
semantic engine cutover, not a provider copy followed by a pathname hash: publication, link policy,
metadata, typed conflict callbacks, and proof/readback all remain attached to the returned authority.

When verification is requested, the source digest is updated from the same ordered bytes accepted
by the stage writer. After publication the bridge first queries the returned exact authority for
`IFileSystemBoundContentProof`; a valid BLAKE3/size proof completes verification without an extra
destination read. Otherwise an advertised host-readback route opens the exact bound publication and
streams BLAKE3 through the same bounded buffer, pause/cancel checkpoints, and task bandwidth
limiter. It never reopens the destination pathname. Digest mismatch is `Failed`; no usable exact
proof/readback is `Unavailable`; cancellation is `Canceled`. All three release the cleanup record
source-retained and prohibit `DeleteIfUnchanged`. They do not recopy the item or enter an ordinary
Exists/Retry conflict.

Writer proof (R3-2). A destination route that advertises `FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST`
binds no published object; its atomic-final writer instead reports the digest the service holds for
the object it published. Before the first byte the bridge asks the writer which digests it can
prove (`GetContentProofAlgorithms`) and hashes the streamed bytes with each of them alongside the
BLAKE3 source digest; after a successful Commit it asks for the proof (`GetCommittedContentProof`)
and compares size and digest. That comparison is the item's verification result when verification
was requested, and it is the content proof Managed Move requires on such a route: the cleanup
record stays armed through publication, a matching proof allows `DeleteIfUnchanged` on the bound
source, and a missing or different proof retains the source (`Copied; source kept`). The writer
proof never substitutes for bound authority in conflict decisions: overwrite on those routes stays
the R3-1 conditional replacement, and nothing is deleted by name.

The bridge reserves one verification transfer lane per task. Discovery-ahead may continue, but the
next file's reader/writer pump and publication cannot overlap the current file's provider proof or
host readback. Directory workers and top-level item scheduling both obey this bound; verification
progress is therefore sequential with transfer even when ordinary Copy concurrency is greater than
one.

An opaque provider proof advances logical verification progress once for the proven content size and
is charged to the task bandwidth limiter as one completed unit after the provider returns. It does
not fabricate host-readback bytes, read calls, a verification graph sample, or a readback rate. Exact
host readback is charged incrementally per buffer and is the only source of the verification graph
series. Each selected item verifies through its own bridge context, and the task's verification progress
is the sum over its items: each report adds the bytes that item verified since its previous report.

## Directory traversal

The bridge uses the single-pass discovery scheduler from
`Specs/FileSystem/FileSystem_FileOperations.md`.

- There is no separate pre-calculation traversal or Calculating phase.
- Discovery-ahead and execution share one bounded queue.
- Skip discovery stops ahead-work and releases its reservation; just-in-time safety discovery
  remains mandatory.
- Directory-on-directory merges. Conflicts are raised for colliding children.
- Recursive Move deletes each proven source file after its own destination contract, then removes
  directories non-recursively and deepest-first.
- A new/changed, skipped, or retained source descendant blocks parent removal without opening a
  second cleanup prompt. That retention propagates through only the affected ancestor chain.
- Directory metadata retains O(depth + directories with in-flight children) state. Metadata is
  restored post-order only for a destination directory created or replaced by the current operation;
  a pre-existing merge target keeps its own metadata. The bridge holds no handle on a created directory
  while its children copy (I35: a Samba server refuses every rename inside a directory another handle
  holds with DELETE access): it releases the published directory at once and keeps its identity
  snapshot; when the children have finished it binds the path again, no-follow, and writes the metadata
  only through a bound directory whose snapshot equals the recorded one. A destination that cannot bind
  it again records the one-per-task note that it keeps no timestamps; a renamed, replaced or removed
  directory is never written (Warning `metadata.basic.lost`).
- Closure at listing end (I35 S2). The parallel producer counts each listing whole when it reads it
  (a name the destination refuses is not counted; the walk records it as a failure when it reaches
  it), so a root's discovery closes once the root and every directory its listings counted are listed;
  a directory the producer will not list (a link it does not follow, a skipped or failed directory)
  stops holding the scope open when that is decided. The producer moves on as soon as a directory is
  listed: each queued file holds a count on its directory's completion record, and whoever finishes
  last (the producer or a worker) restores the created directory's metadata, children first. A
  directory whose children are still in flight counts against the producer's queue lead, so parked
  directories stay bounded; after the workers join, a stop releases the queued work it abandoned so
  every created directory is still restored.
  The sequential walker counts each listing when it reads it and lists pending plain directories ahead
  of the walk within the lookahead budget of the File Operations spec, so a root's discovery closes at
  its last listing; a rename merge lists only what it walks into.
- `Common/FileOperationTraversalPolicy.h` is shared with the Local provider for the quantitative
  ceilings plus queue-target and worker-reservation calculations; bridge queue/error/I/O ownership
  remains local to the host.
- The bridge walkers (sequential, parallel producer, and link traversal) keep their directory
  frames on an explicit stack, not on the C++ stack: a frame holds one directory's enumeration,
  child cursor, retained-byte accounting, created-directory metadata, and Managed cleanup record,
  and pops post-order with exactly the metadata restore and cleanup decision the recursive walker
  performed (R4-T1). Depth is therefore reported (`bridge.traversal.*` counters), not terminal,
  and so is each open directory's retained listing (the provider's `IFilesInformation` buffer plus
  the frame's index of child-name views, released when the frame pops; R4-T3). Retained memory is
  therefore proportional to the widest listing on the currently open path, not to the tree, and is
  no longer capped by a fixed budget: a duplicate child
  name, or a case-only variant when the destination folds case, is still refused per directory,
  but no aggregate metadata budget stops a large directory. The hard ceilings that remain are
  4,096 retained work entries and 16 MiB retained UTF-16 queued-path text over the task-lifetime
  parallel work queue. Reaching one of those stops discovery for the task even under Continue on error, emits the
  exact limit diagnostic, and leaves prior completed publications intact. The self-test depth override
  (`SetFileOpsBridgeTraversalDepthLimitForSelfTest`) is the only way `bridge.traversal.depthLimit`
  is still produced. The Local provider's permanent-delete walk keeps its frames on an explicit
  stack as well (R4-T2); only its native copy walkers, reachable through the provider's own
  `CopyItem`, keep a depth ceiling.

Links are handled as objects. Skip records the link outcome without opening its target. Preserve
binds and classifies the source no-follow, calls `ReadBoundLink` on that exact authority, creates a
cryptographically named link stage through `CreateExclusiveLink`, and publishes only through that
owned stage's `PublishAs`. Provider success with a null link token or a published non-Link snapshot
is a contract violation. The bridge never materializes, hydrates, or mutates a target object. For
Managed Move, `DeleteIfUnchanged` on the retained link authority becomes reachable only after
successful publication; any unsupported/changed/indeterminate publication reports source-kept.

Preserve is literal (see the File Operations spec's link policy): the payload is the stored target
text, kind, and relative flag, and the provider reports it as outside-root with no source-relative
component; any other mapping is a provider-contract failure. Each link publishes in provider order
as soon as it is read, under both the sequential walk and parallel directory traversal. The bridge
keeps no sparse component map, failed-prefix records, deferred link queue, dependency
classification, or held directory cleanups, and it never builds a complete source tree, target
graph, or pathname cleanup manifest. A link into the moved tree therefore keeps naming the source
location; the spec states that consequence and the engine does not repair it.

The sequential walker validates one provider directory buffer, then keeps only that active
directory's bounded entry/name views plus its O(depth) ancestor frames and the listings its lookahead
made ahead of the walk (at most 128 listings and 16 MiB).

## Buffering and scheduling

- Active admitted file pumps: at most 16; queued/active work is also charged to the traversal entry
  and path-text ceilings until its terminal item result releases the reservation.
- Aggregate host pump buffers: at most 256 MiB.
- Buffer size is clamped from both endpoint hints and measured provider behavior.
- A scheduling checkpoint occurs at no more than 4 MiB or 50 ms.
- Enumeration shares provider concurrency but not byte-bandwidth accounting.
- Verification bytes consume the configured bandwidth limit and are reported separately from pump
  throughput.
- Pause, Cancel, deadline, and live bandwidth changes are observed at checkpoints.

Progress phases (I35 S6, no ABI involved): the bridge names each file's phase on the host progress
boundary: Opening before its writer and stage exist (that report honors a cancellation; on a
connection with a copy/move limit it comes before the permit wait, so the wait is the file's Opening
rather than the previous file's publication), Transferring for pump chunks, Finalizing from the
always-sent report of the file's last bytes (commit-and-prove follows it without a report of its own;
a file of unknown size reports Finalizing before commit-and-prove), Publishing before publication
and for link records, and Verifying for exact readback. The Publishing boundary is sent only when the
stream's last report is at least one progress period old, so a stream of small files keeps the
ordinary report cadence. The Finalizing and Publishing boundaries are observation points only: they
never consume a cancellation, so a cancel is met at the next ordinary checkpoint and the cleanup it
triggers is unchanged.

Current numeric queue targets are perf parameters, not Preferences or ABI. Tuning must preserve
boundedness, non-starvation, and immediate Skip-discovery release.

The FBP-10 architecture decision is to retain the current overlapped bridge pipeline: a disabled-
pipeline baseline and enabled-pipeline candidate must be measured in the same test-enabled Release
process on the same machine before changing its worker or buffer shape. The accepted 2026-08-27
evidence used four 32-MiB Dummy files with 30-ms chunk latency and an 8-MiB resolved buffer; the
disabled baseline completed in 130.484 seconds and the enabled candidate in 0.250 seconds. Therefore
no additional thread layer, per-file buffer owner, or larger aggregate budget is authorized. Future
tuning must preserve the 16-pump/256-MiB ceilings, checkpoint cancellation, discovery non-starvation,
and correctness while demonstrating a lower measured cost without a throughput regression.

## Conflict and cleanup

Stage creation is exclusive. Overwrite/replace-read-only consent applies only to final conditional
publication and never authorizes a stage pathname. Keep Both is finalized by create-exclusive
publication. Collision classification reads source and destination metadata no-follow before the
prompt. Folder-on-folder merges recurse and prompt only on colliding children; file/directory kind
mismatches never gain Overwrite. An exact final destination link is a typed Replace link / Keep
Both / Skip conflict, while any linked destination ancestor remains a hard stop.

Before a destructive collision action is exposed, the host binds the current final destination
no-follow and retains that exact authority across the prompt. Overwrite, Replace read-only, and
Replace link return it with the decision; the provider consumes it once through `PublishAs` or
`RenameIfUnchanged`. If the provider cannot bind that authority, Replace link is withheld, and
Overwrite/Replace read-only are offered only when the destination's atomic-final writer accepts
the overwrite flag (R3-1): the bridge reads the occupant no-follow before the prompt, creates the
writer with the granted flags, passes that occupant through `IFileWriterExpectedReplacement`
before the first byte, and the provider replaces only that occupant (`ERROR_REVISION_MISMATCH`
otherwise, nothing published). A route that can do neither withholds every destructive action.
Replacing a directory link uses an atomically created, identity-owned directory stage;
the resulting receipt never becomes a recursive overwrite grant. Its metadata is restored like any
created directory's: through the path bound again and accepted only with the identity recorded at
publication, never by a blind pathname write.

An operation-wide compatibility overwrite flag is not a destination receipt. When the final name is
absent, both Local and cross-provider publication remain create-exclusive and retain the newly
published authority; the flag cannot force the replacement path or require a nonexistent bound
destination. When exact published authority is unavailable, metadata restoration is skipped with a
typed authority-unavailable diagnostic. Copy keeps the provider defaults; managed Move applies the
metadata/source-retention gate before exact source cleanup.

Keep Both belongs to the reported collision depth. The bridge chooses the sibling path in the
colliding item's parent and records only that changed component in the sparse semantic-link map.
It never retries a nested collision by renaming the selected top-level directory. Apply-to-all is
an explicit exact-scope grant. Skip with the checkbox checked records the same scoped grant as
Skip All; unchecked Skip is one-shot. Skip All is unavailable for Recycle or unresolved semantic-link
target retention.

Cleanup may call `AbortOwnedObject` only for the exact stage token returned by exclusive creation.
A name pattern such as `.rs_tmp_` is never cleanup authority. Publication failure, cancellation,
or provider uncertainty preserves any object whose ownership/outcome cannot be proven and reports
an artifact/result axis.

The current bridge chooses exactly one publication route per file: a provider-confirmed atomic
final writer, or an identity-owned exclusive stage. Absence of both fails before source bytes are
read. The atomic-final route carries a granted replacement only together with the occupant
expectation (R3-1); a writer that cannot carry it fails the item before any byte is read
(`bridge.replace.expectationUnavailable`), and a refused replacement is reported as
`bridge.replace.occupantChanged` with nothing published; the bridge then re-raises the collision on
the current occupant, at most twice per item, before `ERROR_REVISION_MISMATCH` reaches the item
executor's ordinary retryable conflict (Retry, Skip, Cancel). The identity-owned route binds an existing destination no-follow after overwrite consent,
commits the stage, calls `PublishAs` with that exact expected destination, validates the returned
known/committed state and non-null published authority, and performs size verification and metadata
write through that authority. Content verification duplicates the retained content-capable handle;
it does not reopen the published pathname, so a concurrent replacement is inspected as the new
pathname owner but is never substituted for or deleted as the operation-owned publication. The
bridge never promotes or cleans a stage by pathname. `S_OK` with a null
source reader, writer, expected binding, stage authority, or committed publication authority is a
provider-contract violation and fails closed.

The bridge never escalates privileges and never pathname-deletes after a sharing/access failure.

## Results

Every item reports strategy plus `PublicationState`, `VerificationState`,
`SourceDisposition`, `ItemCompletion`, exact HRESULT/provider status, byte/timing counters,
link action, metadata loss, and artifact references.

- Copy-only with publication is not a failed Copy; it is a partial/with-notes Move result.
- A source is Removed only after exact conditional deletion reports a known committed outcome.
- Unknown publication or source disposition maps to `ERROR_IO_INCOMPLETE`.
- Security-significant metadata-loss consent is resolved against the exact owned stage before
  publication. Cancel reports `NotPublished + Retained + Canceled`; a later cancellation preserves
  the exact publication state already recorded by the bridge.
- Cancel after prior completed items is Partial/Canceled according to aggregate mapping.

## Provider validation

All reader/writer/bound-object methods validate required output pointers. The host validates:

- capability-v2 path profile and required sections;
- QI claims and non-null returned interfaces;
- `sizeBytes` on every public struct;
- read/write counts and monotonic positions;
- exact Commit/publication output;
- conditional-mutation `outcomeKnown` on every return path;
- Abort/deadline and quiet-point behavior advertised by the profile.

An advertised true/supported claim has an executable provider contract test. Contract violation
disables the affected strategy and is diagnosed once per provider instance/profile.

## Metrics and tests

Metrics record endpoint profiles, selected strategy, source read/destination write/source hash/
exact published-object readback bytes, proof/readback call counts,
pump/Commit/publication/verification/metadata/delete/quiet-point duration,
buffer and admitted-file high-water, traversal depth/retained entries/path bytes/metadata bytes and
limit hits, semantic-link sparse-map/deferred-queue high-water and immediate/resolved/retained counts,
discovery mode/queue/starvation, and typed result axes. The
identity-owned route additionally records aggregate stage-creation duration/count, publication
duration/count, and retained-stage count; it never emits a row per descendant.

Deterministic tests cover short/over-read/write, early EOF, null-success outputs, stage collision,
publication replacement race, source replacement race, conditional delete conflict, cancellation
at every phase, backward links, forward links with a nested Keep Both rename, link cycles, failed
dependency source retention, directory mutation during traversal, verification Off zero-work,
provider proof, bound host-readback, zero-byte content, mismatch, unavailable proof, cancellation
before cleanup, Native NotApplicable, deep long paths, wide-tree
retention high-water, exact ceiling diagnostics with prior-success preservation, post-order directory
metadata, bounded memory, and Copy-only downgrade.

## Logical-file progress at bridge mutation boundaries

The bridge owns stable logical-leaf receipts across pump/publication retries. It reports active
Finalizing after payload pumping, Publishing before destination mutation, and Verifying before
proof. Successful owned-stage mutation reports Published in the committed branch before any
missing-authority, provider HRESULT or later proof failure; direct atomic-final Commit reports
Published immediately after the successful commit before metadata consent/proof can fail.
Directory-stage publication has no file receipt and never increments the copied-file counter.
Terminal preserves the last observed written bytes, including partial failed/canceled attempts.

Known file sizes survive phase-only observations, including files at most one MiB; no 0/0 reset
stands in for finalization. A selected single file can adopt its actual reader size as the whole
total. Sequential directory children cannot: discovery owns that folder denominator, and their
cumulative written bytes must not clamp to the first child's size. Listing metadata with omitted
sizes does not certify a partial total. Complete totals use independently summed reader sizes
only after every discovered file-shaped leaf has a known size and traversal closes; unsupported
size stays unavailable. Existing fail-closed source-size/proof policy is not relaxed.

Host-owned nested stage/promote/cleanup calls suppress discovery and file-work observation.
Receipt storage remains bounded to admitted regular workers plus producer-only semantic-link
work, and adds no per-chunk observer traffic. Operation control remains the only pause/cancel
checkpoint. The test-only committed-error seam binds one exact owned virtual source and sandboxed
physical destination, consumes once only after real committed publication with returned authority,
and clears at the worker quiet point; it never synthesizes a progress event or selects a renderer.

The observer independently certifies ordinary-directory completion. A sequential directory emits
`ObjectCompleted` only after its successful epilogue; a parallel directory does so only after its
last required child and metadata complete. Failed, skipped, canceled and abandoned subtrees do not
become copied folders. These reports use the existing depth/in-flight directory state rather than
an unbounded path set, do not increment copied-file count and do not add per-chunk callbacks.
Replay without identity-preserving deduplication withdraws directory-count certainty. The task
retains copied file/folder counts through completion for the single post-Skip work summary.

Deferred metadata/hydration consent carries the admitted destination endpoint's qualified root ID
separately from the selected destination item path. Repeated choices therefore share the same
task/risk/profile/root across selected roots while each item still performs its own identity and
mutation checks. Retargeting cannot carry a grant to another endpoint. EFS Move consent explicitly
states that continuing the move removes the encrypted original after successful transfer; its
alternative is Copy and keep source. EFS Copy has no redundant retain-source choice.
