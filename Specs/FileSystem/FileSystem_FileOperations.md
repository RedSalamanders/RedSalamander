# File Operations Specification

## Authority and scope

This is the authoritative product and engine contract for RedSalamander Copy, Move, Delete,
Recycle, inline Rename, Batch Rename, and archive cleanup. Provider mechanics are owned by
`Specs/Plugins/Plugins_VirtualFileSystem.md`; the cross-provider pump is owned by
`Specs/Core/Core_FileSystemBridge.md`; UI presentation is owned by
`Specs/UI/UI_FileOperationsPopup.md`.

The implementation may change public and host-internal payloads while this program is in
development. The host, every shipped provider, Dummy, and all tests use one ABI version. There is
no capability-v1 adapter or dual parser.

Create Directory remains a provider command outside `StartOperation`, but not outside admission.
The host queries capability v2 for the concrete parent and provider-joined candidate with
`FILESYSTEM_CREATE_DIRECTORY`, requires `operations.createDirectory: true`, matching non-empty
path-scoped `rootId`/`pathProfile`, one feasible provider-native component, and exact immediate-child
containment. It never reinterprets a provider path as Win32 because it looks absolute. A native
`CreateDirectoryW` fallback is allowed only for an explicitly qualified builtin Local path after
device/`GLOBALROOT`/NT namespace rejection; every other provider uses its directory interface.
S3 flat-prefix success requires a committed zero-byte trailing-`/` marker, S3 Table is unsupported,
and a future native-directory profile may advertise success only when it creates a real native
directory object.

## Product invariants

1. Path text is an address, not object identity.
2. Every authoritative bind is no-follow. A link is an object; the engine never follows its target
   to copy, overwrite, rename, delete, verify, or clean up the target.
3. A hash proves content only. No hash, size, timestamp, filename, or pathname authorizes Skip,
   overwrite, artifact cleanup, or source deletion.
4. Move is native Move (one provider rename per selected object, or a per-child rename merge when a
   directory meets an existing directory on the same endpoint), managed copy-then-delete across
   endpoints, or Copy-only with the source retained.
5. Managed Move removes a source only after identity-owned destination publication, every required
   consent, requested verification, and the retained source binding's exact `DeleteIfUnchanged`
   revalidation/mutation boundary. The host never reopens the source pathname as destructive
   revalidation or deletion authority. A rename merge removes an emptied source directory only
   through that same exact cleanup boundary or, on a profile without object binding, through the
   provider's own empty-directory removal, which fails source-kept on `ERROR_DIR_NOT_EMPTY`.
6. Any uncertain destructive outcome fails closed. Copy-only is a valid result and is displayed as
   `Copied; source kept`.
7. Folder-on-folder is merge. It is never folder overwrite.
8. Each file publishes independently. Cancellation or crash may leave both copies; it must not
   leave a published partial file presented as complete.
9. Task intent and consent are immutable after confirmation. Workers do not reread Preferences,
   provider configuration, or the clipboard to make data choices.
10. Recursive traversal, queues, path storage, buffers, and recovery evidence are
    bounded and observable.
11. Pane, Find, Compare, watch, and search consumers act from typed result axes, never pathname
    inference.
12. No File Operations path silently relaunches with elevation or kills a provider worker.
13. A Move never applies the link policy. `Preserve links` / `Skip links` is a Copy option; every
    Move relocates every link object literally, by native rename on one endpoint or by literal
    publication across endpoints.

## Identity vocabulary

| Term | Meaning |
|---|---|
| Qualified endpoint | Plugin ID, provider instance ID, path profile ID, and root identity |
| Path identity | The provider path/name requested inside a qualified endpoint |
| Object identity | Stable no-follow local file ID or provider object ID |
| Revision identity | Immutable provider generation, VersionId, or ETag where the path profile guarantees it |
| Content identity | BLAKE3, an equivalent exact provider content proof, or the digest the destination service reports for the object its writer published (R3-2) |
| Bound object | Open handle or provider token retaining identity and authorized operations |
| Complete item | Fully committed destination object with required publication semantics |

The qualified endpoint `rootId` is the provider's opaque, path-scoped namespace-root identity and
is interpreted only inside the `(pluginId, instanceId, pathProfileId)` tuple. Capability v2 must
return it before admission. Providers with multiple roots in one instance distinguish them (for
example local volume, S3 bucket, MTP storage, or cloud drive); a provider with one configured root
may return one stable provider-root token. Display paths and host-manufactured hashes are not root
identity. Missing `rootId` fails admission closed; exact object mutation authority is still rebound
and revalidated immediately before mutation.

For the builtin Local provider only, `/` is the pre-navigation provider-root metadata sentinel used
to discover route and name semantics before a pane has a concrete volume path. That sentinel is not
mutation admission. Every Local mutation must re-query its concrete provider path, and the provider
must reject relative, incomplete UNC, Win32 device, `GLOBALROOT`, and NT object-manager namespaces
before child-name validation, joining, binding, or mutation can authorize the request.

`rootId` qualifies strategy topology; it does not define the provider's object-ID comparison
domain. When plugin ID, provider instance ID, and path profile ID match, retained object IDs are
still compared across different root spellings (for example a local drive path and an
administrative-share UNC alias). Different `rootId` values prohibit assuming one native strategy
route, but they do not suppress an exact-object or ancestor-identity stop.

The host canonicalizes pane/provider instance context into a non-empty plan identifier. An empty
context becomes the reserved `host/default` instance; a present context is namespaced as
`host/context/<opaque-context>`. These forms prevent an absent context from colliding with a provider
context whose literal text happens to be `default`. They are comparison-only qualification and are
never display labels or mutation authority.

When an ingress supplies an explicit filesystem object that is not the pane filesystem, admission
queries that object for `IInformations` and derives the plugin ID/short ID from its metadata. It MUST
NOT label the explicit endpoint with the pane provider's identity. Missing or malformed explicit
endpoint metadata fails admission before queue publication.

Local exact identity is volume identity plus `FILE_ID_INFO` from a no-follow handle. The legacy
64-bit `nFileIndex`, FNV-1a, basic metadata, and path spelling are insufficient destructive
authority. S3 destructive cleanup requires VersionId/generation or an equivalent conditional token.

The host represents binding as a typed result: Bound, Unsupported, Missing, Indeterminate, or
ProviderContractViolation. It copies opaque object/revision bytes immediately and retains the COM
bound object; it never retains provider snapshot pointers. Revalidation rebinds the current path
no-follow and produces Same, Changed, Missing, Unsupported, Indeterminate, or
ProviderContractViolation by asking the original bound object to compare with the current binding
and cross-checking the copied object/revision bytes. Provider comparison and immutable object-ID
equality must agree; disagreement is a provider-contract violation. A matching object with a changed
revision or kind is Changed. Only Same may preserve exact-object authority. Local `OpenReader`
reopens the retained content-capable handle as an independently positioned overlapped handle and
verifies that handle's `FILE_ID_INFO`; it does not reopen by pathname or require the object to remain
reachable by its former name. A local
reparse object is a Link only when `IsReparseTagNameSurrogate` is true (symlink,
junction/mount, or equivalent name surrogate). WOF, cloud-placeholder, dedup, and other
non-name-surrogate reparses retain their regular-file/directory kind and are not followed merely to
classify them. The same classifier owns Local Verify-Off Copy at the selected root, serial child
walk, parallel child queue, and destination-directory ensure: only a name surrogate enters
Preserve/Skip semantic-link handling; a non-name-surrogate reparse is copied/merged according to its
regular file/directory kind.
The parallel cross-provider producer must enqueue known regular-file reparses through
its ordinary bounded file-admission queue. It releases its own copy-buffer allocation
before workers run, so it must never copy placeholder payload on the producer. Workers
retain the normal source-binding, hydration-consent, cancellation and publication contracts.

Every capability, name-policy, and concurrency gate queries the concrete source or destination
provider path it governs. `L"/"` is valid only when it is itself that concrete provider path; it is
never an instance-wide placeholder substituted for a selected item or destination. Before task
publication, the plan factory enforces the operation Boolean for every selected source and the
destination when both endpoints match. For different endpoints it instead enforces stable path
identity, source read and destination write, and both sides of the matching export/import pair. A
Move intent may use the proven **Copy** export/import pair as `CopyOnly`; that strategy never
requires or invokes source Delete and reports `Copied; source kept`. The Move export/import pair
and exact source-delete proof are required only for a managed destructive Move. An empty Move pair
therefore prohibits managed copy-delete; it does not prohibit the safe Copy-only downgrade.
Capturing profile/root IDs without enforcing those fields is not admission.

Native Move qualification queries the concrete Move profiles and does not depend on a Copy
capability document. When Native is unavailable, the engine queries and enforces the concrete
source/destination Copy profiles and Copy export/import pair. Failure of that Copy route rejects
the request. The engine then qualifies the separate Move pair plus source bound/conditional Delete,
destination exclusive-stage/conditional-publication, and binding-interface facts: all facts select
Managed; any absent destructive fact selects `CopyOnly`. This prevents a valid native route from
being rejected by an unrelated Copy profile while making the safe publication route mandatory for
both non-native outcomes.

Copy and Move execute only through the identity-guarded per-item executor. Transfer admission
defaults to per-item execution and any bulk Copy/Move request fails before provider I/O; provider
bulk execution remains an explicit Delete-only choice in this slice. Serial and parallel scheduling
call one qualified per-item policy that owns mutation dispatch, Native race requalification,
Keep Both, conflict decisions, Recycle escalation, prompt clearing, traversal-limit/cancel terminal
classification, and typed result publication. Concurrency-one execution stays on the operation
worker so provider thread/session affinity does not change; higher concurrency uses the shared
scheduler. The common policy permits Copy/CopyOnly publication,
provider-native Move, or Managed Move through an exact bound-source reader and identity-owned
destination publication. It rejects Native Move when a destination bridge is present. CopyOnly
never arms the legacy bridge source-cleanup manifest or a source-delete hook. A Managed item that
cannot acquire exact source authority before reading uses the same Copy route, retains the source,
and reports `Copied; source kept`; it never falls through to pathname cleanup.

For a regular-file bridge, the host requests no-follow read-content and metadata authority together.
A successful exact binding is the only reader authority used: the host passes the task's existing
`FileSystemOptions`, including `operationControl`, to `IFileSystemBoundObject::OpenReader` and does
not fall back to a pathname reader when that exact open fails. Legacy `CreateFileReader` is retained
only for providers that explicitly do not support object binding. A missing, indeterminate, or
contract-invalid binding fails closed before content transfer.

Permanent Delete on a route without object binding (C10): while the task prepares, the host asks
the provider for the identity of every selected root through `IFileSystemIdentityDelete` (an object
key with its version or ETag on S3, a file id on Google Drive), re-resolves it after the card's
answer, and deletes through `DeleteIfIdentity`, which the provider refuses with
`ERROR_REVISION_MISMATCH` when the live identity differs; nothing is deleted then and the item is a
definitive non-commit. A provider that cannot name the object (FTP/SFTP/SCP) keeps the native path
delete, and the card's consent text adds that this location deletes whatever the name refers to when
the deletion runs (`IDS_FILEOPS_CONSENT_PERMANENT_DELETE_BY_NAME`). A directory identity names the
directory object; descendants keep the provider's own per-observation rules.

## Typed operation contract

`RedSalamander/FolderWindow.FileOperationsInternal.h` owns the executable host types:
`QualifiedEndpoint`, `ProviderIdentitySnapshot`, `QualifiedSourceItem`,
`TransferDestinationMapping`, `DestructiveConsentReceipt`, `OperationOptions`,
`TransferPlan`, `RenamePlan`, `DeletePlan`, `FileOperationPlan`, and `FileOperationPlanGroup`.

- Copy/Move intent exists only in `TransferPlan`.
- Recycle is `DeletePlan::mode == Recycle`, not another operation variant.
- `RenameOrigin::Unspecified` is invalid. `RenameOrigin::InlineRename`,
  `RenameOrigin::BatchRename`, and `RenameOrigin::ChangeCase` are explicit and distinct even for
  one step, so an omitted origin never inherits another command's Queue/presentation contract.
- An ingress identity snapshot is comparison evidence only. Execution binds again before mutation.
- Permanent Delete requires a matching consent receipt. Archive cleanup receipts cover only the
  qualified selected items captured by the successful archive command.
- Ordinary transfer has no explicit mapping. Compare supplies exactly one unique, in-range,
  non-escaping destination mapping per selected source.
- Plans are immutable after successful qualification. Ordinary typed operations qualify before task
  publication; scheduled Rename (`BatchRename` or `ChangeCase`) publishes its ordinary gated task
  first and atomically installs the immutable plan during common Preparing before any acceptance,
  clipboard consumption, breadcrumb publication, queue entry, or mutation.
- One child plan has exactly one source endpoint. If a single user request spans source roots, the
  host partitions it into one child plan per `(pluginId, instanceId, pathProfileId, rootId)` and
  admits the immutable child-plan group as one task. The original relative ordering inside each
  root is retained; explicit mappings are rebased to child-local indexes. Confirmation, task ID,
  result card, cancellation, and clipboard Move consumption belong to the group and occur once.
  A malformed child rejects the complete group before publication; a later execution failure in
  one child is ordinary per-item partial-result behavior and does not make another root's identity
  applicable to it.

### Mutation ingress table

Every mutation enters through the typed plan factory. The current production call site is migration
evidence, not a permanent exemption.

| Ingress | Required plan and confirmation owner |
|---|---|
| F5/F6 Copy or Move to the other pane | `TransferPlan`; routine accepted-default admission has no generic confirmation, while explicit/material routes retain their owning decision |
| Pane Delete/Recycle | `DeletePlan`; Delete confirmation and Recycle rules |
| Clipboard paste, including same-pane duplicate | `TransferPlan`; captured clipboard sequence for Move |
| Internal drag/drop | `TransferPlan`; retain the qualified source endpoint |
| External OLE / `CF_HDROP` drop | Qualify the local source, then `TransferPlan`; report Copy to OLE until asynchronous Move is proven |
| Find Files Copy/Move/Delete | `TransferPlan` or `DeletePlan`; retain resolved provider/item correlation |
| Compare Directories synchronization | `TransferPlan` with complete `explicitMappings`; Compare summary is the ingress confirmation |
| Pack cleanup | `DeletePlan(Permanent, PackCleanup)` after successful/verified pack; the archive prompt supplies the one captured `ArchiveDeleteAfter` consent |
| Unpack cleanup | `DeletePlan(Permanent, UnpackCleanup)` after successful extraction; the archive prompt supplies the one captured `ArchiveDeleteAfter` consent and no private `std::filesystem::remove` is permitted |
| Inline F2 | One-step `RenamePlan(InlineRename)`; inline editor is the only initial confirmation |
| Batch Rename | `RenamePlan(BatchRename)`; preview and Run are the initial confirmation |
| Change Case | Discovery emits immutable mappings into `RenamePlan(ChangeCase)`; the Change Case dialog is the initial confirmation |
| Command/plugin extension | The same plan factory; direct provider mutation is forbidden unless named here |

One physical source object may occur in at most one scheduled `RenamePlan` step. Repeated path
text is a blocking preview error; admission also rejects different path aliases or hard links whose
no-follow ingress snapshots identify the same object. The engine does not serialize multiple final
names for one physical object.

## Admission and confirmation

Before queue admission, validate only the immutable request envelope: non-empty qualified endpoints
and items, structurally valid paths, correct plan variant, valid Compare mappings, valid receipt
scope, and a complete captured cut list. Reject local device namespaces such as
`\\?\GLOBALROOT`, `\\.\`, and `\??\`.

Do not recursively enumerate or preflight all top-level state before the first mutation. Object
kind, existence, identity, containment, capability, metadata, space, and descendants are discovered
and revalidated when reached. A later failure preserves prior successful items and yields a partial
result.

Batch Rename is the narrow preview exception because its explicit preview already materializes a
bounded, selection-proportional mapping rather than a recursive tree. Its UI boundary publishes an ordinary
gated File Operations task after copying only immutable pane/provider identity, mappings, execution
choice, and completion/progress endpoints. The worker owns path capabilities, provider-parent and
destination derivation, no-follow ingress binding, duplicate-object rejection, one identity-domain
proof, dependency/cycle construction, validation, and atomic immutable-plan publication. Failure or
cancellation during this phase completes the task with no plan and no mutation; live pane navigation
cannot retarget the captured request.

After plan publication and before the first rename mutation, the worker rebinds every changed source
and checks every unrelated final namespace slot once. A missing/replaced source, an occupied unrelated
target, or an indeterminate binding fails before any row mutates. A target occupied by another planned
source is a declared dependency/cycle edge and is not rejected by this check. This pass is comparison-
only: it does not retain destination replacement authority, enumerate descendants, or replace the per-
hop exact revalidation required at each mutation boundary.

Routine accepted-default Copy/Move, including F5/F6, destination-picker transfers, clipboard paste,
and internal drag/drop, does not invoke a generic confirmation. The stable
`cmd/pane/copyToOtherPaneWithOptions` (`Shift+F5`) and
`cmd/pane/moveToOtherPaneWithOptions` (`Shift+F6`) commands are the deliberate pane-transfer
ingress to the same single pre-consumption options surface; their acceptance publishes the same
plan as the corresponding routine command with the selected task-local options, and Cancel publishes
no task and mutates nothing. A Copy/Move confirmation remains mandatory only when the ingress sets
`requireConfirmation`, a Move is already known to contain a Copy-only strategy, or an existing exact
artifact/risk gate requires a decision. These causes share one confirmation and never stack generic
dialogs. Permanent Delete retains its separate consent rule; preview/editor/drop decisions that
already own initial consent remain authoritative. Find-result and Compare synchronization commands
retain their specialized behavior on the original routine command IDs and do not inherit the pane
with-options commands.

When an explicit/material Copy/Move confirmation is required, it captures:

- operation, qualified endpoints, source summary, and destination;
- for Copy only, Links: `Preserve links` or `Skip links`; a Move confirmation shows and captures no
  link policy;
- Verify copied file contents, defaulting from `fileOperations.verifyAfterCopy == false`;
- Queue/Parallel and current bandwidth controls;
- any already-known EFS, sparse-inflation, or hydration facts.

Clipboard Move also states: “Starting this Move consumes the cut list. If an item cannot be removed,
its source will remain and the cut list will not be restored.”

The Verify control carries a path-qualified availability state in
`HostFileOperationPromptOptions::verificationAvailability`: Supported, Unsupported, Check during
operation, or Not applicable. Unsupported and Native-only selections show a disabled Off value with
an explicit reason. A capability lookup that exceeded the UI budget keeps the choice enabled and
records Check during operation; later absence of both exact proof and bound readback reports
`VerificationState::Unavailable` without clearing the accepted intent. The host validates this enum
at the ABI boundary and writes accepted values back only on the affirmative prompt result. The
mandatory capability-v2 document must still complete and validate before admission; the elapsed
budget changes only the Verify presentation state and never authorizes mutation from partial,
timed-out, or stale capability data.

The central boundary emits one aggregate metric set per attempted plan, never per discovered child:
`fileops.plan.construct_us`, `fileops.plan.admit_us`, `fileops.plan.source_count`,
`fileops.plan.payload_bytes`, and `fileops.plan.rejection_bucket`. Timing rows carry source count and
retained plan bytes so admission cost can be compared by selection size without logging path text.

Batch Rename additionally emits `batchrename.admission.ui_thread_us` for the complete UI
capture/publication path, `batchrename.admission.wall_us` for end-to-end qualification, and
`batchrename.admission.worker_wall_us` for worker-only qualification. The
1,024-row reference-profile UI budget is 50,000 us, with no row-proportional provider calls permitted on
that thread. Capability/bind counts, zero retained admission bindings, and cancellation latency are
archived as bounded values; no provider path or identity payload is logged.

The legacy link setting migrates silently: Follow becomes Skip, CopyReparse becomes Preserve, and
Skip remains Skip. Preserve is the default for absent/new settings. The old Follow numeric value is
reserved and rejected at configuration and ABI boundaries.

Admission reads the provider-owned default once and seeds the immutable task options. When an
explicit/material Copy/Move confirmation is present, it may replace Links, Verify, Queue/Parallel,
and bandwidth for this task. On acceptance the host copies the complete effective option set into
every child plan, revalidates the group, and only then publishes one immutable plan group. Every
provider call receives the captured Links value in `FileSystemOptions`; it overrides the provider's
current configuration for that call. Workers and providers MUST NOT reread live settings after
admission. Task-local confirmation choices never write provider configuration or Preferences.

### Common Preparing lifecycle

Every admitted ordinary, clipboard, Inline Rename, Batch Rename, and Change Case task has one
engine-owned monotonic lifecycle:

`Preparing -> AwaitingAcceptance/Ready -> Waiting/Running -> Stopping -> Terminal`.

`AwaitingAcceptance` is used while the immutable preparation snapshot is available and the common
pre-consumption decision is unresolved. `Ready` means common preparation and every required
acceptance/clipboard receipt have succeeded. Pause, conflict, verification, and discovery are
orthogonal facts; they neither replace the lifecycle nor authorize a skipped transition.

Preparing is selected-root/top-level proportional. The worker qualifies a scheduled Rename plan
when applicable, validates the immutable plan group, binds the selected top-level mutation/interlock
scopes, publishes one immutable authority-free `PreparationSnapshot`, runs the one
`preConsumptionDecisionGate`, and observes cancel/stop. The snapshot contains task/operation,
selected-root count, copied endpoint/role/path and scope facts, known strategy counts including
Copy-only, and preparation timing/status. It retains no executable mutation authority and its
observer cannot mutate the task. Preparing performs no recursive enumeration, descendant conflict
scan, content read/hash, totals walk, mutation, clipboard consumption, breadcrumb write, or history
write.

Task publication arms the fixed 500-ms `kTaskCardRevealDelayMs` presentation deadline without
constructing or showing the popup synchronously. A queued task or an actionable conflict, failure,
partial, canceled, or indeterminate outcome reveals immediately. A task still Preparing, waiting,
or running at the deadline reveals the same stable card and replaces its status in place; it never
creates a separate calculation surface or moves focus. Clean accepted-default work that completes
before the deadline records its ordinary completed history but does not open the popup. Explicit
**Show File Operations** reveals any still-hidden live task immediately.

Failure or cancellation before Ready produces dense terminal truth for every selected root:
`NotAttempted` plus source `Retained`, with the exact Failed or Canceled outcome. It consumes no cut
sequence, writes no Move breadcrumb, and performs no mutation. A clipboard Move clears its captured
sequence exactly once only after successful common preparation and the decision gate. An admitted
Move creates its breadcrumb only after preparation and every required acceptance/clipboard receipt,
then before queue/execution admission. A known Copy-only Move therefore cannot mutate without the
explicit acceptance captured before task publication; a runtime-only downgrade remains governed by
the per-item execution gate and truthful source-retained result.

## Capability and strategy routing

Every shipped provider exposes the separate mandatory
`IFileSystemRouteCapabilities` interface. `GetRouteFacts` returns the complete
typed route, cancellation, namespace, path-identity, operation, proof,
publication/binding/delete/link, concurrency, and full provider/profile/root facts
for the requested path and operation. Missing QI, malformed facts, an identity
mismatch, an uncontained route, or a failed executable claim disables the affected
strategy. `IFileSystemPathCapabilities2` JSON remains diagnostics/extensions only:
positive JSON cannot enable a route, and missing/false/malformed JSON cannot disable
or repair valid typed facts. Optional
`IFileSystemObjectBinding`/`IFileSystemBoundObject` support is used only when queried
successfully.

Delete admission is flag-aware. A request carrying `FILESYSTEM_FLAG_USE_RECYCLE_BIN` requires
`operations.recycle: true`; a request without that flag requires `operations.delete: true`.
One bit never implies the other. Permanent Delete, Recycle, and Rename are admitted
only when the typed path profile advertises the exact operation, has a contained
cancellation route, exposes stable path semantics, and binds the executable authority
required by central planning. A true typed capability is descriptive feasibility,
not evidence that a same-named provider method or exact mutation authority exists.

Transfer admission (Copy and Move, plain and explicitly mapped) admits every destination
leaf through the destination provider's `ValidateChildName` before any path is joined.
That query is a provider call, so it runs in Preparing on the task thread under the
synchronous-I/O cancel watch (C1): the UI thread keeps only the leaf-shape validation and
publishes the task, and a leaf the destination cannot hold fails the task on its card
before any mutation, with the provider's reason (`ERROR_INVALID_NAME` by default), the
`PlanRejectionBucket::InvalidDestinationName` telemetry row and a
`preparation.destinationName` diagnostic naming the leaf, instead of failing, or being
silently altered, at stage creation; a provider without the contract keeps the plain join. The
remaining host path joins (`JoinFileSystemPath` in transfer destination derivation, native
Move shaping, artifact projection, and the Batch Rename engine) therefore only join leaves
the destination contract, or the rename admission, has already accepted (R0-RC3 step 9).

Cross-provider transfer requires strict directional agreement from
`IsTransferPeerAllowed`: source Export and destination Import both receive the full
opposite plugin ID. `ValidateChildName`, `GetChildNameCollisionKey`, and `JoinPath`
are the only provider/path-scoped name-feasibility surfaces. Validation applies to
proposed names only: `GetChildNameCollisionKey` and `JoinPath` enforce just the shared
shape rules (non-empty, not `.`/`..`, no accepted separator), so a collision key and a
join exist for every name the provider can already enumerate, including legacy names
that break the provider's own proposed-name rules. Inline Rename, Batch
Rename preview and worker admission, Change Case, and Create Directory consume one
composite result from a retained typed route generation for each proposed name and
only the collision key or join for existing names. The provider validation
status and exact failure, provider-owned joined path, and provider collision key are
copied before return. Missing, unsupported, malformed, or changed facts fail before
mutation; provider Invalid preserves its exact HRESULT, and the shared provider base
distinguishes the length bound (`ERROR_FILENAME_EXCED_RANGE`) and reserved device
names (`ERROR_BAD_DEVICE`, including `COM0`-`COM9`, `LPT0`-`LPT9`, their superscript
forms, `CONIN$`, and `CONOUT$`) from every other invalid proposed name
(`ERROR_INVALID_NAME`). Batch Rename checks shape before the provider call
(`name_empty`, `name_dot`, `name_separator`), maps provider Invalid to
`name_too_long`, `name_reserved_device`, or `name_invalid_character`, combines the
typed parent path key with the provider collision key for duplicate, existing-target,
and acyclic-schedule decisions, lists each parent through `ReadDirectoryInfo`, and
revalidates each step immediately before mutation. An existing child the provider
cannot key is skipped by that listing, by Create Directory's suggestion, and by Change
Case discovery; it cannot collide with a provider-valid target, and execution still
binds every destination, so the command stays available and fail-closed. Change Case
validates every changed leaf before its artifact guard or first rename and revalidates
each depth batch. Create Directory uses provider collision keys for initial and
suffixed names. Generic Win32 name, separator, length, normalization, or case rules
are never a fallback for a virtual provider. The Local provider caches its resolved
volume facts briefly so per-name queries cost one `GetVolumePathNameW`.

For copied regular objects, the engine also queries the bound source for optional
`IFileSystemBoundMetadata`. A successful snapshot is an exact-object claim, not a pathname hint.
The provider applies sparse and EFS state to the exclusive destination stage before the first
content byte, then transfers named streams, MOTW, extended attributes, ordinary metadata, and
security-change evidence to that same stage after content commit. NTFS compression is never
transferred: the destination inherits its parent folder's compression state, exactly as File
Explorer and `CopyFileExW` behave, so a compressed source is neither attempted nor reported lost.
Every attempted class is reported independently as preserved, lost, or changed. Missing QI is
supported degradation: Copy continues
with exact loss reporting, while Managed Move must pass every applicable deferred-consent gate
before source cleanup. Typed route facts and diagnostic JSON never substitute for
this per-object result.

For host-managed publication, an admitted file must have either a provider-confirmed atomic-final
writer for the exact final path/grants or an exclusive writer plus retained bound-stage authority.
The latter route binds any expected destination no-follow, publishes conditionally through
`PublishAs`, verifies/sets basic metadata through the returned exact published authority, and aborts
only the exact stage token. There is no unowned temp-path promotion or pathname cleanup fallback.
The Local provider implements this bound route; FileSystemDummy implements the atomic-final route.
A granted replacement on an atomic-final route without object binding publishes through that writer
with the occupant the user saw as its expectation (R3-1, Conflict model): the writer revalidates
that occupant with its strongest identity (item id, ETag, version, or last-write time) immediately
before publication, and a changed or vanished occupant is a definitive non-commit, never an unknown
publication.
Any required `S_OK` null output, unknown mutation outcome, or identity mismatch fails closed and
retains every object whose ownership cannot be proven.

For ordinary built-in Local Copy of a regular file to an absent destination name, the provider uses
the visible direct-final publication shape: it creates the final leaf exclusively, retains the exact
bound object returned with that writer, transfers sparse and EFS state before content,
streams content through a bounded 1-MiB buffer, commits, and transfers named streams, MOTW, extended
attributes, ordinary metadata, and security-change evidence to that same exact object. No sibling
stage is created. Success reports final publication `Published` and owned-object disposition
`Published`. Metadata classes the destination volume cannot hold are best effort exactly as with
`CopyFileExW`: a lost sparse state before content, or a lost named stream, MOTW,
extended attribute, or ordinary attribute after commit, is recorded as
`FileOps.Local.DirectFinalMetadataLost` and never fails the Copy or removes the published content.
The one exception is an EFS source that cannot stay encrypted at the destination: it fails closed
with `ERROR_ENCRYPTION_FAILED` before the first content byte, because plaintext is never published
for an encrypted source. Any pre-commit failure closes the writer and calls `AbortOwnedObject` on
the retained object; it never deletes or reopens the current pathname occupant. That abort uses the
same bounded cleanup snapshot as sibling-stage compensation (`MakeOwnedStageCleanupOptions`: no
primary `operationControl`, a fresh deadline at most two seconds ahead). A successful exact abort reports
`Removed`, an unknown abort outcome reports `Unknown`, and a known surviving partial final leaf
reports `RetainedIncomplete` with `ERROR_IO_INCOMPLETE`. Overwrite, directories, links, Move, and
non-Local providers do not use this direct-final route.

Local sibling-stage compensation for files, links, and directories is also exact-handle-only, but
uses a separate bounded cleanup snapshot after primary work stops. Immediately before
`AbortOwnedObject`, both the host bridge and Local provider call the shared
`MakeOwnedStageCleanupOptions` helper, which removes the primary `operationControl` and cookie and
sets a fresh absolute deadline no more than two seconds in the future. Primary Cancel therefore
cannot revoke already-proved authority to remove that exact owned stage; timeout or uncertain abort
remains visible cleanup debt. Writer-allocation failure after exclusive stage creation uses the same
cleanup snapshot. No compensation path reopens or deletes the current pathname occupant.

Capability v2 separates `operations.move` from `operations.nativeMove`. `move` admits the provider
Move entry point but never selects a strategy. `nativeMove` is strategy evidence only after the
engine has qualified the exact source/destination profile pair. Local advertises
`nativeMove: true` because its provider Move entry point performs one native rename only, for a
regular file, a directory tree, or a link object alike; folder merge is the host rename merge
below, and the provider contains no copy/delete or per-child fallback of its own. The host still
passes `FileSystemOptions::moveMode == FILESYSTEM_MOVE_NATIVE_ONLY` as the explicit strategy
assertion.
MTP remains `nativeMove: false` until its tested same-device leaf-preserving WPD route is similarly
isolated.

Just before mutation, a stable capability-v2 path profile is also structural authority for exact
path equivalence and strict destination-inside-source rejection within one qualified root. This
check remains mandatory when optional object binding is unavailable. It is not object identity:
different path text is never proof that two objects differ, and it never authorizes source cleanup.

Every provider that advertises `nativeMove: true` accepts `FILESYSTEM_MOVE_NATIVE_ONLY` at both its
single and batch Move boundaries and rejects every unknown mode before provider I/O. Dummy proves
an in-memory node relocation/merge, Microsoft Drive proves an item-ID-preserving Graph `PATCH`, and
S3 proves its provider-owned pinned-revision publication plus exact conditional source deletion.
None may silently enter a generic pathname copy/delete fallback.

Local `MoveItem`/`MoveItems` never perform cross-volume, folder-merge, or semantic-copy fallback,
including when a legacy/test caller omits `FILESYSTEM_MOVE_NATIVE_ONLY`. A Local Native call is one
provider mutation only. Cross-volume returns `ERROR_NOT_SAME_DEVICE`; a regular destination
directory that appears after admission returns a known destination-exists non-commit; and a
destination link returns the typed link conflict. Every such failure retains the complete source
without moving any child. Managed Move is the sole Local copy/publication/delete route and exists
only across endpoints; folder merge on one endpoint is the rename merge.

| Strategy | Admission | Terminal meaning |
|---|---|---|
| Native Move | Source and destination share one qualified endpoint (provider, instance, path profile, root) whose provider advertises `operations.move` and `operations.nativeMove`; the host requests `FILESYSTEM_MOVE_NATIVE_ONLY`. No link proof and no delete proof is required | Destination and source disposition reported from provider proof; inability to use the native route fails before fallback |
| Rename merge | A Native directory Move whose destination is an existing regular directory on the same endpoint | Each non-colliding child relocated by one provider Native rename; collisions through the typed conflict surface; emptied source directories removed by exact cleanup; no bytes transferred, no stage, no proof |
| Managed Move | Bound source, exact conditional source delete, and either identity-owned stage/publication with destination proof or (R3-2) a destination writer that proves the published content | Destination published and proved, then the exact unchanged source removed |
| Copy-only | The Copy operation/pair works but any destructive prerequisite or Move pair is absent | Destination published; source retained with reason; no source Delete call is reachable |

Host qualification includes a non-empty same-connection Curl directory rename against
an owned FTP fixture. Success must relocate nested files and empty directories with
one RNFR/RNTO pair and preserve siblings. Server refusal must never trigger a relay
copy, pathname cleanup or automatic retry. Curl currently reports an attempted failed
rename without an authoritative non-commit receipt conservatively as Indeterminate;
fixture-observed source retention does not grant the host stronger mutation authority.

Native admission does not depend on item shape. A regular file, a directory tree, or a link object
on one qualified endpoint is Native whenever the provider advertises `operations.nativeMove`; the
pane listing, admission, and Preparing perform no top-level shape probe, no provisional split, and
no recursive preflight. A rename relocates every descendant and every link object as stored, which
is exactly the Move link contract, so no separate link proof exists. The Native call remains one
provider mutation: a regular directory that meets an existing regular directory returns the
provider's known destination-exists non-commit (`ERROR_ALREADY_EXISTS` or `ERROR_FILE_EXISTS`)
before any child is touched, and the same task continues that item as a rename merge without a
second task, a repeated confirmation, consumed clipboard state, or a replayed Native call. Every
other Native non-commit (a destination link, a type mismatch, a read-only or sharing failure)
enters the ordinary typed conflict surface; Native never merges through the link namespace, and a
native sharing failure never falls through to pathname deletion or to a copy. A case-only directory
rename names the exact source object and is not a merge shape. Local capability `rootId` remains
volume-specific, so a cross-volume pair is never Native and uses Managed only when every
destructive proof exists, otherwise Copy-only. A native rename can succeed on an open file when
managed copy/delete cannot.

Rename merge. The host enumerates the source directory with the bounded walker it already uses for
Managed Move, in provider order, and for each child: an absent destination name is one provider
Native rename of that child (a file, a subtree, or a link object, never entered); a regular
directory on both sides recurses; every other collision (file, link, type mismatch, read-only)
enters the typed conflict surface, where Overwrite is the provider's Native replace route already
used by Native file Move (`RenameIfUnchanged` on a bound profile, the provider's own replace
sequence otherwise), Keep Both renames to the collision-safe sibling name, and Skip retains the
child. After its children, an emptied source directory is removed through the exact cleanup
boundary described under Publication and managed Move; `ERROR_DIR_NOT_EMPTY` retains it and its
ancestors as `Moved; source folder kept` without a prompt, a recopy, or a second delete. The merge
transfers no bytes, creates no stage, computes no proof, reports verification `NotApplicable`, and
counts items for progress. A profile without object binding removes the emptied directory through
its own delete route by name, the authority its Permanent Delete already has.

Managed Move retains one item-sized cleanup record containing the exact bound source authority from
reader creation through terminal source disposition. A known non-commit in a retryable sharing,
access, path, or transient-provider bucket uses the central `Retry / Skip / Cancel` prompt. Retry
reissues only `DeleteIfUnchanged` against that same record; it never republishes or recopies the
item. Skip releases the record and reports `Copied; source kept`. An unknown conditional-mutation
outcome or contradictory provider result is terminal/indeterminate: no Retry is offered, no second
Delete is issued, and the published destination is preserved for reconciliation.

When exact conditional cleanup of a bound source directory returns a known non-commit
`ERROR_DIR_NOT_EMPTY` after its selected children were processed, the directory membership changed
after discovery or an unselected child remains. The engine automatically reports `Copied; source
kept` (`Moved; source folder kept` for a rename merge), releases that directory's cleanup record, and propagates retained-source disposition through
its ancestors. It does not offer Retry/Skip, issue a second Delete, copy the unselected child, or
classify this as a destination collision. Already-published selected children remain committed.

Cleanup classification is source-owned. In particular, `ACCESS_DENIED` from
`DeleteIfUnchanged` is an AccessDenied source-cleanup conflict; destination attributes cannot
reinterpret it as a read-only replacement conflict after publication. The cleanup call itself is
the retained binding's identity/revision revalidation and mutation boundary. The host does not
reopen the source pathname for a redundant pre-delete comparison and never reconstructs cleanup
authority from a pathname, digest, destination reread, or whole-tree manifest.

Provider baseline:

| Provider/profile | Copy | Move |
|---|---|---|
| Local Win32 | Native/managed | Native for every shape on one volume; rename merge onto an existing directory; managed cross-volume only with exact binding/publication/delete |
| 7z | Export | Copy-only |
| Curl FTP/SFTP/SCP | Same-connection stream copy; bridge export/import Copy | Native same-connection rename for files and directories; rename merge onto an existing directory; Copy-only across connections |
| Curl IMAP | Export message content | Copy-only; UID identity never authorizes unbound cleanup |
| S3 flat prefix | Stream/server-side copy | Conditional VersionId/generation strategy or Copy-only |
| S3 directory bucket | Path-profile provider semantics | Path-profile provider semantics |
| Microsoft Drive | Cross-endpoint bridge import/export Copy; same-qualified-endpoint Copy unsupported | Native same-drive Graph Move by item ID with ETag/reconciliation |
| MTP | Stream copy | Copy-only/source kept; neither native nor host-managed destructive Move is qualified |
| Google Drive | Server-side `files.copy` per file; folders recreated per child | Native parent/name change by item ID (`files.update` addParents/removeParents); folder-onto-folder merges per child |
| Dummy | Executable test profiles | Exercises supported/unsupported matrices |

## Streaming discovery and scheduling

The separate recursive pre-calculation model is retired. There are no
`fileOperations.preCalcEnabled`, `fileOperations.preCalcMaxWorkers`, or standalone
`Calculating` state.

One traversal discovers descendants and feeds the bounded execution queue. It never walks a tree
solely for totals and then walks it again. Discovery-ahead is enabled by default and does not delay
the first safe mutation. Counts reported while discovery remains open are provisional: a task whose own traversal is
still open renders indeterminate progress and publishes no ETA denominator; the footer aggregate
follows the popup's Known-work rule (`../UI/UI_FileOperationsPopup.md`, `FO-DISCOVERY-01`): the
cohort of included tasks with closed totals renders determinate `Known work: N%` while other tasks
still discover, and the Windows taskbar stays indeterminate while any included total is open.

The host records observation-only discovery-service facts from the first completion it can prove
while traversal remains open: elapsed microseconds from discovery start to that observation,
completed-byte delta, and completed mutation/item delta. Provider progress contributes real deltas.
Exact Local permanent Delete has no provider progress callback, so its typed `Removed` selected-root
result contributes that root's discovered bytes and exactly one mutation before the root's discovery
scope closes. This is not a synthetic callback, descendant mutation count, progress publication, or
scheduler input; it cannot alter queue targets, worker limits, admission, or mutation ordering.

Each provider reports the selected root even when it is a single leaf and no recursive enumeration
is needed. A leaf's discovery record contains its exact file byte size and closes that item's
traversal before publication starts; recursive roots let their existing walker own the root and
descendant cumulative totals. A provider must not leave leaf totals at zero merely because the copy
path itself does not enumerate a directory.

The traversal retains queued file/reparse work plus O(depth) directory/ancestor state; it never
retains a completed-tree-sized directory list or pathname manifest. `Common/FileOperationTraversalPolicy.h`
is the single implementation owner consumed by the host bridge and Local Copy walker for the
shared quantitative ceilings and reservation calculations. The host bridge walkers keep their
directory frames on an explicit stack (R4-T1), so tree depth is not a bridge ceiling, and each open
directory's retained listing (the provider's buffer plus the frame's index of child-name views) is
reported through `FileOps.Bridge.TraversalMaxMetadataBytes`, not capped (R4-T3): memory grows with
the widest listing on the currently open path, never with the whole tree; a duplicate or,
on a case-folding destination, case-only-variant child name is still refused per directory. The
hard ceilings are 4,096 retained work entries and 16 MiB of retained UTF-16 queued-path text over
the task-lifetime link and cleanup records. The bridge additionally admits at most 16 file pumps and
the aggregate host pump-buffer budget remains 256 MiB. Reaching any ceiling is task-terminal even
when Continue on error is selected: discovery stops at that item, the exact ceiling is diagnosed,
and already published prior items remain published and are reported as partial prior success. The Local
provider's permanent-delete walk keeps its frames on an explicit stack as well (R4-T2); only the
Local provider's native copy walkers, reachable through its own `CopyItem`, still recurse with a
128-level ceiling. The `R4T_DeepTreeCopyCompletesIteratively` self-test guards the bridge walk and
`R4T2_DeepTreeLocalDeleteCompletes` the Local permanent-delete walk, each on a 300-level tree.

Permanent recursive Delete uses that same discovery/execution rule. The Local provider opens each
directory without following reparses, retains at most 256 child records for the next mutation batch,
then closes the enumeration handle and deletes that batch before discovering more. It never flattens
the selected tree before the first delete. Its walk keeps one frame per directory on an explicit
stack (the directory path, its terminal-failure record, the current batch and cursor), so depth is
reported through `FileOps.DeleteTraversal.MaxDepth` and is not a ceiling; a parallel batch still
hands each entry to a worker once, and the walk below that entry runs through the same frame loop
with a serial budget. Continue-on-error retains at most 4,096 terminal child
failures and 16 MiB of their UTF-16 path text per active walk so a repeatedly failing child cannot
cause an unbounded rediscovery loop. Recycle remains shell-owned and can report only the selected
top-level objects unless that provider exposes descendant discovery.

Curl native Copy/merge uses strict incremental discovery and an explicit ancestry
stack for both serial and parallel batches. Previously published files survive a
later discovery failure; no pathname-based recursive rollback is attempted. The
first invalid directory row is a failure, not a successful empty copy. Hard
depth/storage limits stop work even with Continue on error; other eligible child
failures may continue independently and result in partial completion. See the
provider spec for quantitative reservations and the distinct Move preflight proof.

Directory metadata is captured only when the operation creates or replaces that destination
directory, retained on its O(depth) traversal frame, and restored after all admitted children have
finished. Folder-on-folder merge never overwrites the metadata of the pre-existing destination
directory. Restoration uses the exact newly created destination authority, deepest-first. A
publication rename may clear staging-only Hidden/Temporary state; ordinary source attributes and
timestamps are therefore applied once more through the exact published-object authority after
publication, never by reopening the destination pathname.

The popup exposes one-way **Skip discovery** while discovery-ahead is active. It:

- stops new discovery-ahead admissions at the next scheduling checkpoint;
- keeps work already discovered;
- releases discovery reservations/throttling by the next checkpoint, target at most 50 ms excluding
  an active provider call;
- switches to just-in-time discovery and permits the full configured operation concurrency and
  bandwidth;
- does not skip enumeration, identity, containment, capability, name, conflict, or safety checks;
- remains irreversible for that task.

Enumeration shares provider concurrency but does not consume byte-bandwidth accounting. Scheduler
parameters are internal rather than Preferences or provider ABI. The current Local and host-bridge
queues use low-water 32, target `max(128, effectiveConcurrency * 16)`, and a hard discovery-ahead
target of 256 work records. Just-in-time mode admits at most the effective active concurrency. While
discovery-ahead remains open, transfer keeps one effective slot available for discovery; below
low-water it uses at most half the effective transfer slots so metadata traversal receives enough
provider and I/O bandwidth to refill the queue. Skip discovery removes both restrictions at the
next scheduling checkpoint. Boundedness, no starvation, and checkpoints at no more than 4 MiB or
50 ms are normative. Changing 32/128/256 or the half-rate rule requires same-machine discovery/
transfer evidence and an update to this section. HDD local roots remain serial where the storage
profile requires it; Queue/Parallel, Start now/reorder, live bandwidth limits, and
storage-adaptive concurrency survive.

Discovery reservation is task- and endpoint-scoped. It must not create a host-global gate between
independently authorized fixed volumes. A provider that requires one serialized device session may
keep maximum backend concurrency at one while the host task remains live; the host must not infer a
general device-class throttle from that narrow contract. A route whose typed cancellation facts are
uncontained is rejected before task publication, provider access, or mutation and is not replaced
by a live remote benchmark; R0f closes that gap per route (Local UNC and mapped shares are closed by
R0f-SMB, below).

The host supplies `IFileSystemOperationControl` on every governed provider call. Recursive
providers query its current discovery mode at bounded scheduling checkpoints and synchronously
report cumulative discovered bytes/files/directories, current retained queue depth, and traversal
closure. The host converts each top-level item's cumulative values to task deltas, clamps overflow,
and treats callback failure as operation failure. A provider without a separate run-ahead queue may
discover just in time in its mutation traversal; it still must not perform a second totals walk.

## Publication and managed Move

For every copied regular item:

1. bind the source no-follow and capture exact identity/revision;
2. classify destination ancestors and final entry no-follow;
3. create an unpredictable stage exclusively and retain its identity/token;
4. stream with bounded buffers and exact byte-count validation;
5. commit, then publish only if the expected destination identity still matches;
6. optionally verify the exact published object;
7. apply metadata and required consent;
8. for managed Move only, revalidate and conditionally remove the exact original source.

A rename merge performs none of these steps. The directory cleanup rules below apply to it
unchanged.

An exclusive stage created by the Local provider has one of two exact authority modes. A nonzero
`FILE_ID_INFO` remains the persistent authority on filesystems that support it. When a remote
redirector cannot supply `FileIdInfo`, the provider may instead retain the exact creation handle and
attach a cryptographically random lifetime nonce to that handle-bound stage. That nonce is an
opaque capability, not a filesystem identity: it cannot be reconstructed from a pathname, used for
pre-existing-object binding, cached across the handle lifetime, or replaced by
`BY_HANDLE_FILE_INFORMATION` volume/index fields. A successful all-zero `FILE_ID_128` is treated as
Unsupported. If neither authority mode is available, exclusive-stage creation fails closed before
publication.

Failure to obtain cryptographically secure stage-name entropy fails the item before writer creation.
There is no PID, thread-ID, clock, counter, or other predictable-name fallback. Copy and Copy-only
Move use the same immutable verification option: byte-count and committed-size validation remain
mandatory, while the optional BLAKE3 content pass runs only when requested by the plan. Verification
never changes Copy-only source retention into source-delete authority.

Success with a null reader, writer, capability document, directory result, bound object, or
published object is a provider-contract failure. Unknown publication or mutation outcome forbids
retry, stage cleanup, or source deletion until exact reconciliation.

Directory cleanup is non-recursive and deepest-first. A changed/new descendant prevents parent
removal. A skipped link/collision or any descendant whose exact cleanup retains the source also
retains each containing source directory without another directory-not-empty prompt. Retention is
propagated only through the affected ancestry; a clean sibling subtree remains eligible for exact
post-order cleanup. Sharing, access, path-too-long, and disk-full errors never trigger pathname
cleanup or UAC.

## Folder and same-folder semantics

- Directory onto directory merges and recurses. Only colliding children conflict.
- Regular file collisions, type mismatches, and destination link objects use the conflict model.
- An existing regular-file destination always enters the typed conflict model, including when its
  current bytes equal the source. Size, timestamp, or digest equality never silently selects Skip.
- Hidden, system, and filter-hidden descendants are included when they are descendants of a
  selected directory; pane filters do not redefine the selected tree.
- Same-pane `Ctrl+C` then `Ctrl+V` copies into a collision-safe Keep Both sibling name.
- Same-folder Move is rejected before mutation.
- Same-folder comparison and explicit destination-mapping containment use the admitted path-profile
  identity contract, including its separator and case rules; host `std::filesystem` lexical rules
  are not provider authority.
- The immutable transfer plan is also the single destination-path owner. For an ordinary transfer,
  derive the leaf with the admitted source path profile and join it to the admitted destination
  folder with the destination path profile. For an explicit mapping, use the mapped provider path
  exactly. Execution, completion results, focus/refresh, and cache notifications all consume that
  same resolution; none reconstructs a provider path with host `std::filesystem` rules.
- For local Copy/Move, execution binds source, destination, and required ancestors no-follow when
  each top-level item is reached, then revalidates the retained authorities immediately before the
  provider call. Same object, destination-inside-source, destination-link/ancestor-link, source or
  destination replacement, missing authority, unsupported binding, and indeterminate comparison
  stop before mutation. A direct same-path Copy alone is converted to Keep Both; an alias such as a
  hard link is not.
- Copying a hard-linked file creates an independent destination file. Preserving an entire
  hard-link group is not implied.

## Link policy: literal Preserve

Preserve means preserve the link object as it is stored: its kind (file symlink, directory symlink,
or junction), its target text with the original relative or absolute spelling, and its relative
flag. The engine never follows, opens, resolves, or rewrites a target. Skip reports and skips the
link. Providers that cannot preserve the required link kind report unsupported and do not
materialize its target.

The policy applies to Copy. A Move never applies it: a Move relocates every link object literally,
by native rename on one endpoint or by literal publication across endpoints, and the Move
confirmation shows no link choice.

The host Preserve route binds the source link no-follow, reads its bounded literal payload through
`IFileSystemObjectBinding::ReadBoundLink`, creates an unpredictable sibling stage exclusively
through `CreateExclusiveLink`, and publishes only through that exact stage's `PublishAs`. Success
with a null stage token or a published non-Link snapshot is a provider-contract failure. Managed
Move retains the same source-link authority and may call `DeleteIfUnchanged` only after link
publication; unsupported, changed, failed, or indeterminate publication leaves the source and never
falls back to pathname cleanup or target materialization.

Literal copying has visible consequences that the engine states rather than repairs:

- an absolute or root-relative link into the copied tree keeps naming the source location, and
  after a Move it dangles until the user retargets it;
- a relative link keeps its relative text, so it resolves against the destination link's parent and
  finds the copied sibling when the tree shape is unchanged, and dangles when Keep Both renamed the
  component it names;
- a missing or external target is copied as stored; existence is never checked;
- a link cycle is preserved once per link object; traversal never enters a link.

Every link publishes in provider order like any other item, with no forward/backward
classification, dependency records, deferred queue, link ceiling, or whole-tree map. A colliding
link enters the ordinary typed conflict surface (Type mismatch, Replace link, Keep Both, Skip). Keep
Both on a nested target renames only that target; the links that named it are not rewritten. The
built-in Local recursive Copy route obeys the same rule: it reconstructs file symlinks, directory
symlinks, and junctions from no-follow payloads with the stored target text and publishes each
through the exclusive-stage route, under both its sequential child walk and its parallel item
queue.

`ReadBoundLink` reports every payload as `FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT` with no
source-relative component; the transform's roots and component mappings are validated for ABI
hygiene and otherwise ignored. The remaining mapping values and an optional retarget transform
belong to a later plan.

A native provider Move relocates a tree or a link object as stored. Because a Move never applies
the link policy, no link proof gates Native admission and `nativeMove: true` is sufficient. The
pane listing, admission, and Preparing make no shape decision for Native Move.

Executable cases include:

| Source case | Required result |
|---|---|
| `S/docs/link -> ../assets/a.png`, no rename | Destination link text is `..\assets\a.png`; it resolves to `D/assets/a.png` |
| Absolute `S/docs/link -> S/assets/a.png` | Destination link text is unchanged; it still names `S/assets/a.png` |
| Top-level `assets` becomes `assets (2)` | Link text is unchanged; a relative link into `assets` now dangles, by design |
| Nested target directory receives Keep Both | Only that directory is renamed; links naming it are not rewritten |
| Target appears after link in traversal | Link publishes in provider order; nothing is deferred |
| Missing or external target | Stored target text is copied; nothing is resolved |
| Symlink/junction cycle | Link object is preserved once; traversal never follows it |

## Conflict model

Every prompt is created from a typed conflict record containing source/destination kind and
identity, endpoint profile, and one immutable engine-owned action policy. That policy contains the
complete ordered action IDs, the exact primary/`More...` placement, the default and Escape actions,
Apply-to-all and Skip All eligibility, metadata-loading state, and whether decision buttons are
publishable. The engine is the only production owner that may add, remove, order, place, or choose a
default for conflict actions. A presentation surface may copy and render these fields but MUST NOT
derive actions or placement from `ConflictBucket`.
The task becomes visible immediately as **Needs attention — Reading details** before the host begins
the no-follow collision-metadata read. This loading state publishes no decision actions and accepts
no decision: its ordered, primary, and overflow action counts are zero; default/Escape are `None`;
scope eligibility and button publishability are false. After metadata resolves, the host replaces it
in place with one stable, complete typed action set; buttons are never added to or removed from an
already actionable prompt. A
name-surrogate reparse is a link; non-name-surrogate cloud/WOF/dedup-style reparses retain their
file/directory kind. An
exact final destination link is classified here and is not mistaken for a linked ancestor; linked
ancestors remain hard stops. Revalidate immediately before mutation. Overwrite, Replace read-only,
and Replace link are shown when the mutation provider can consume the exact retained destination
authority returned with that decision. On a destination route without object binding (FTP/SFTP/SCP,
S3, Microsoft Drive/SharePoint, Google Drive, MTP, Dummy) Overwrite and Replace read-only are also shown
when the destination's atomic-final writer accepts the overwrite flag (R3-1): the host reads the
occupant no-follow before the prompt, hands it to the writer through
`IFileWriterExpectedReplacement`, and the writer replaces only that occupant, failing with
`ERROR_REVISION_MISMATCH` and publishing nothing when it changed or vanished; the host then raises
the collision again on the current occupant (at most two more decisions for the same item) and
finally reports that status through the ordinary retryable conflict (Retry decides again on the
current occupant, Skip keeps it). Replace link stays bound-only. In both forms the receipt applies to
that object once; it is not an operation-wide flag and cannot authorize a colliding child.

Legacy operation flags remain input compatibility only. `ALLOW_OVERWRITE` does not turn an absent
name into a replacement: an absent destination always uses create-exclusive publication and retains
the exact new authority. A destructive receipt exists only after a no-follow destination bind and a
typed user decision for that object.

| Conflict | Actions | Default | Apply-to-all |
|---|---|---|---|
| Regular file exists | Overwrite / Keep Both / Skip / Cancel | Cancel | Same regular-file/profile class |
| Read-only regular file | Replace read-only / Keep Both / Skip / Cancel | Cancel | Read-only regular files only |
| Exists plus read-only | One combined prompt | Cancel | Same combined class |
| Type mismatch | Keep Both / Skip / Cancel | Cancel | Same kind pair |
| Exact destination link | Replace link / Keep Both / Skip / Cancel | Cancel | Same link kind |
| In-tree link target cannot be mapped/published | Retain the source link and dependent source ancestry; report target conflict | Source retained | Never infer another mutation grant |
| Name infeasible | Keep Both if a valid name exists; otherwise Skip / Cancel | Cancel | Same feasibility class |
| Sharing/access/transient provider error | Retry / Skip / Cancel | Cancel | Exact error class |
| Recycle failed | Cancel / Delete permanently / Skip | Cancel | Never |

Decisions are one-shot unless the user explicitly selects an eligible scope. **Skip All** is Skip
plus explicit Apply-to-all for that row and is exposed through **More…**. It is never available for
Recycle escalation. A digest match never auto-Skips. Keep Both preview names are advisory until
create-exclusive publication succeeds. The task cache key is the complete typed scope: conflict
class, source kind, destination kind, source path profile, destination path profile, and—when the
destination is a link—the exact known link kind. Unknown link kinds remain one-shot. Plain Skip is
never cached; only explicit Skip All records the corresponding scoped Skip grant.

The safe `Cancel` action is directly visible in the primary row and is both the published default
and Escape action; a three-button limit MUST move a non-Cancel action into `More...` instead of hiding
Cancel. `Enter` invokes only the currently focused action or the published default. `Escape` invokes
the published Cancel action for the current decision and never Cancel All. `WM_CLOSE`/window Close
hides the popup without resolving, accepting, rejecting, retrying, skipping, or canceling the
decision. Loading prompts publish no default or Escape decision, so Escape outside an actionable
decision only hides the popup.

Caption Close applies to the File Operations popup as a whole and remains a presentation-only
hide operation; it is never shorthand for task Cancel or Cancel All. The decision surface MUST
remain reachable: **View > File Operations**, the canonical `cmd/app/showFileOperations` command,
and its default application shortcut `Ctrl+Shift+J` show the existing popup without changing any
task or decision. If the popup is hidden while a conflict or deferred-consent prompt is still
loading, publication of the immutable actionable decision MUST automatically show that same popup
without activating a decision. A later user Close wins over an already queued show request; only a
new actionable publication or an explicit Show File Operations invocation may restore it.

The engine builds the policy once after typed metadata refinement and uses that same instance both to
validate an eligible cached decision and to publish the prompt. Applying a cached action MUST NOT
re-derive an action set in the popup or through a second engine builder.

Retry is proof-gated, not counted. A prompt offers Retry only for an attempt the canonical receipt
classifier proved a non-commit (or one that failed before any provider mutation); unknown or
contradictory mutation outcome never reaches the prompt as a retryable conflict. Each click of
Retry requests exactly one further attempt: the per-item loop re-enters its own revalidation
(exact authority, source stability, scope) before that attempt, a conflict that changed comes back
to the same arbiter as a new prompt, and nothing loops, caches Retry, or turns a later Retry into a
silent Skip. The prompt says how many attempts have failed ("Attempt N failed.") and keeps offering
Retry while the failure stays a proved non-commit; Skip and Cancel remain. Every attempt reuses the
retained exact authority and may not rebind a pathname or widen the selected object set. The exact
source-cleanup retry of a Managed Move stays bounded compensation, not a restart of primary work.

After a dispatched provider-native Copy, Move, Recycle, or Permanent Delete returns failure, the per-item
executor classifies that attempt's matching completion receipt with
`FileSystemRouteContract::ClassifyFailedMutation` before any returned-error conflict,
Keep Both, native-race requalification, or escalation. Missing, unknown, or contradictory
proof ends as `ERROR_IO_INCOMPLETE` with an indeterminate item and the ordinary localized
unknown-outcome summary, never a replay/Skip prompt. Known committed or retained-stage
failures also cannot restart primary work. A later completion without a receipt clears
earlier child/attempt evidence, and every new attempt starts with fresh receipt state.
A user or host cancel is the exception at the task level: a canceled item without a receipt
keeps its `Unknown` publication or source axis (shown on the item, and it still blocks
Retry), but that axis does not make the task indeterminate; with no failed item the task
ends `ERROR_CANCELLED`, never `ERROR_IO_INCOMPLETE`, because the cancel is its terminal cause.
Valid independent facts survive uncertainty: known final non-publication and retained
source remain known when only owned-stage cleanup is unknown. The directory-merge
contract's `ERROR_PARTIAL_COPY` with an observed child Skip is terminal partial work,
not a returned-error retry opportunity; an explicit uncertain or invalid receipt still
takes precedence over that partial marker. This does not infer non-commit from Skip.
Local Copy also checks its tracked publication receipt before its provider-internal
returned-error prompt: a post-publication attributes failure cannot replay the copy.
This returned-error boundary does not replace provider-internal pre-mutation conflict
handling or the bridge's separate publication/cleanup proof.

Keep Both is resolved at the colliding item, not by renaming the selected top-level root. The Local
provider's recursive Copy callback accepts the typed KeepBoth action and chooses the collision-safe
sibling in that child's parent. This is required for nested link/directory/file collisions.
Providers that do not implement child-level Keep Both do
not receive that callback action; the host may use its plan-owned top-level retry only when the
reported collision is the selected top-level destination itself.

The built-in Local Copy and Native Move paths consume destructive receipts through
`PublishAs`/`RenameIfUnchanged`. Replacing a selected-root destination link and then encountering a
regular child collision therefore creates a second typed decision; the link receipt cannot silently
overwrite that child. Folder-on-file and file-on-folder are both Type mismatch and never gain
Overwrite.

If a provider returns a collision to the outer item executor without consuming the shared issue
callback, that status carries no exact destination authority. The host may still offer Retry, Keep
Both, Skip, Skip All, or Cancel when otherwise valid, but MUST withhold Overwrite, Replace read-only,
and Replace link. A destructive button is never offered merely because an HRESULT resembles Exists.

## Delete and Recycle

Recycle and Permanent are distinct `DeletePlan` modes. Permanent Delete requires fresh explicit
consent and confirmation-to-mutation identity continuity. Admission makes no provider call for it:
it decides native authority from the provider's binding interface alone and gathers the
confirmation's words from the pane listing. When exact host binding exists, Preparing pins each
selected root no-follow on the task thread (the retained delete authority), then asks the
confirmation on the task card in the words the modal used (Cancel first and default, Permanently
delete); after the answer it binds each confirmed pathname again and requires the object/revision
identity to match the pinned root before recording the consent receipt, and execution calls
`DeleteIfUnchanged` on the retained authority. It never confirms one object and later deletes a
pathname replacement: the object the user confirms is the pinned one, and a pathname that no longer
names it refuses with `ERROR_REVISION_MISMATCH` before any mutation. The pinned roots stay pinned
(no-follow, write and delete sharing denied) while the card waits for the answer, so on Local a
pathname swap in that window is itself refused with a sharing violation and the re-check guards
only a provider whose pin is weaker; a cancelled confirmation ends the task as canceled with nothing
mutated and releases them. A provider
without host binding supports Permanent Delete through its contract-tested native mutation (plan
`nativeAuthority`): admission finds no bound object (`Unsupported`), records no ingress snapshot and
no exact-delete interlock authority (the interlock's confirmation-to-mutation continuity check applies
only to bound authority), and execution calls the provider's `DeleteItem` for the
selected path, keeping the receipt that mutation returns (S3 consumes its exact generation, MTP its
object identity; FTP/SFTP/SCP have no object identity and delete the selected path, so a same-named
replacement created between confirmation and deletion is deleted there, the documented residual every
FTP client shares). A missing or indeterminate native receipt stays `Indeterminate`; the mode is never
guessed from the HRESULT. Every provider completion may include a versioned `FileSystemItemMutationResult` containing
the provider's per-item `outcomeKnown`, `mutationCommitted`, and `originalStillPresent` truth. The
host copies that record during the callback. A missing, malformed, indeterminate, or contradictory
record can never authorize destructive recovery.

Concurrency, bandwidth, connection-profile, and scheduler settings constrain only an operation that
has already passed typed capability and identity admission. They never grant a missing mutation capability:
for example, `deleteMaxConcurrency` cannot make Permanent Delete available on an unbound provider.

The canonical receipt classifier combines typed route support with HRESULT and
`FileSystemItemMutationResult` truth. Its only states are `Unsupported`,
`RetryableNoCommit`, `FailedKnown`, `Indeterminate`, and `ContractViolation`.
Retry is legal only for a proved non-commit with the source still present and owned
stage cleanup known. Missing/malformed/unknown destructive receipts are
`Indeterminate`; a retained owned stage remains explicit cleanup debt and is never
collapsed into Retry.

For synchronous Local Permanent Delete, a failure before removal of the selected root is reported as
a known non-commit with that root retained. This includes a locked leaf and a recursive delete that
removed some admitted children but retained the selected directory because another child failed.
Such an unexpected partial mutation remains `Failed`/`Retained` and aggregates as
`ERROR_PARTIAL_COPY`; only an explicit user Skip is `Skipped` and eligible for `S_FALSE`. Recycle
does not inherit this shortcut: shell-owned work remains unknown unless its per-item completion sink
provides terminal mutation truth.

For a selected regular file, the retained exact-delete binding may report its committed size through
the same operation-control discovery callback used by the mutation walk. Those bytes become completed
delete progress only after `DeleteIfUnchanged` reports a committed removal; a known non-commit leaves
them discovered but not completed. The host does not start a second pathname size walk.

Recycle failure may create a one-item `RecycleEscalation` receipt only when all of these conditions
hold:

1. the Recycle provider reports a known non-commit with the original still present;
2. the task still owns the exact no-follow pre-operation source authority from its mutation
   interlock;
3. a fresh `READ_METADATA | DELETE | NO_FOLLOW` bind matches that retained object identity,
   revision (when present), and kind;
4. the user chooses **Delete permanently** for that exact rebound object.

The fresh delete-bound authority remains alive through the prompt. It denies concurrent write/delete
sharing and `DeleteIfUnchanged` mutates that retained object directly, so the host does not reopen the
pathname after consent. A source replacement before the fresh bind fails before the prompt; an
indeterminate Recycle result, unavailable binding, identity mismatch, cancellation, or unknown delete
outcome keeps the source and never retries permanent deletion. Default is Cancel; Skip keeps the item;
there is no Apply-to-all and no task-wide escalation. The old behavior of clearing the Recycle flag
and retrying provider pathname deletion is prohibited.

When the exact permanent-delete escalation succeeds, that result replaces the earlier Recycle failure
for item completion and diagnostics. The final item is successful; the stale Recycle HRESULT is retained
only as historical prompt context and cannot make the completed escalation appear Failed. A skipped or
canceled escalation similarly reports its explicit terminal action without weakening the original
known-noncommit/source-retained mutation truth.

Pack and Unpack delete-after-success submit ordinary exact `DeletePlan` tasks after the archive
contract succeeds. No archive component performs a private pathname delete.

## Verification and metadata

Verification defaults off and applies after publication to copied regular content. Selecting Verify
On is an explicit engine cutover for non-Native copied payloads, including same-endpoint Local Copy:
the host exact-authority bridge owns transfer, publication, proof/readback, link handling, metadata,
and conflict callbacks. The host does not run provider `CopyItem` and then reopen a final pathname to
hash it; therefore Verify On changes the engine path but not the observable Copy/Move semantics.
Verify Off keeps the provider `CopyItem` route for same-endpoint Local Copy, so SMB server-side copy
and block-clone offloads stay available.

- The host hashes the source bytes already flowing through the bounded transfer pipeline with the
  canonical `Common::Crypto::Blake3Hasher`; verification Off constructs no hasher and performs no
  proof/readback calls.
- The preferred destination proof is `IFileSystemBoundContentProof::GetContentProof` on the exact
  published authority with `FILESYSTEM_CONTENT_PROOF_BLAKE3_256`. A provider proof substitutes only
  when its size, algorithm, committed content size, and bound object/revision are valid.
- If the qualified destination profile advertises `verification.hostReadback`, the host may instead
  duplicate/open the exact published binding and stream BLAKE3 readback. It never reopens the final
  pathname.
- Writer proof (R3-2). A destination profile that advertises `verification.providerProof:
  "writer-digest"` (`FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST`) binds no published object, but its
  atomic-final writer reports the digest the service holds for the object it published
  (`IFileWriterContentProof`: `sha256Checksum` on Google Drive, `file.hashes` SHA-256 / SHA-1 /
  QuickXorHash on Microsoft Drive, the full-object CRC-64/NVME on S3, SHA-256 on the Dummy store).
  Before the first byte the host asks which digests the writer can prove and hashes the streamed
  bytes with those algorithms; after Commit it compares the provider's proof of the same size. A
  match is `Verified` without any destination read, a different digest is `Failed`, and a missing
  or unusable proof is `Unavailable`. The proof never authorizes overwrite, skip, or deletion by
  itself; it is the content proof Managed Move requires on such routes (Move semantics below).
- Zero-byte verification proves object kind, existence, committed size zero, and content proof.
- The states are `NotRequested`, `NotApplicable`, `Verified`, `Failed`, `Unavailable`, and
  `Canceled`. Digest mismatch is Failed. Inability to obtain an exact advertised proof/readback is
  Unavailable and does not enter the ordinary copy/conflict retry loop.
- Directories and Native Move are NotApplicable regardless of the saved preference; Native Move
  performs zero digest, proof, and readback work.
- Verification owns one per-task transfer lane in this slice. Discovery-ahead may continue within
  its bounded reservation, but transfer/publication of the next top-level or nested regular file
  does not start until the current item's proof/readback reaches a terminal state. This applies even
  when Queue/Parallel or provider concurrency would otherwise permit multiple file pumps.
- Failure, cancellation, or unavailability on Move retains the source.
- Verification never authorizes Skip, overwrite, cleanup, or deletion.

Metadata consequences:

| Metadata | Copy | Managed Move |
|---|---|---|
| MOTW, ADS, extended attributes | Preserve or report exact loss | Keep source unless task-scoped “Move anyway; metadata lost” |
| ACL/owner | Destination inheritance accepted; warn on observed change | Warning alone does not block cleanup |
| EFS | Never silently decrypt | Gate before plaintext; default Keep source/Stop |
| Sparse allocation | Preserve or state logical inflation | Gate before inflated write |
| Compression | Never transferred; the destination inherits its folder's state (File Explorer rule) | Never a loss, never a gate |
| Timestamps/ordinary attributes | Best effort with warnings | Does not block cleanup after required gates |
| Directory timestamps | Restore deepest-first with bound authority | Warning on failure |

The Local exact metadata implementation uses the retained source handle and exact exclusive stage.
`BackupRead`/`BackupWrite` copies MOTW, non-default ADS, and EA records without copying the default
data stream; sparse and EFS controls are attempted before content, compression is left to
destination-folder inheritance, and destination ACL and owner inheritance is compared and reported
rather than silently cloned. Stage publication may clear
staging-only attributes, so basic attributes/timestamps are finally re-applied through the returned
published authority. Metadata results never authorize an identical-file Skip or source deletion.

Security-significant metadata-loss consent is evaluated on the exact identity-owned stage before
`PublishAs`. Cancel discards/cleans that unpublished stage where its outcome is known and reports
`NotPublished + Retained + Canceled`; no final destination object is claimed or refreshed. Move
anyway and Keep source first publish the stage, then respectively permit the already-qualified
conditional cleanup or retain the source. A later cancellation point must preserve whatever exact
publication state its mutation record reports rather than collapsing every Cancel to `NotAttempted`.

The shared deferred-consent gate owns insufficient/unknown space, hydration, sparse inflation, EFS
plaintext, metadata-loss Move, and Recycle escalation. It serializes the prompt through the task
arbiter while unrelated workers may remain paused at their governed boundary. The prompt carries
known item/byte facts plus source/destination identity when available; unknown facts stay unknown.
Cancel is the safe default and is normally displayed last. **Recycle escalation is the explicit
exception:** Cancel is displayed first and receives initial/default focus, followed by Delete
permanently and Skip, so Enter cannot authorize permanent deletion. A submitted action that was not
in the published layout is coerced to Cancel.

| Deferred risk | Actions before Cancel | Eligible scope |
|---|---|---|
| Insufficient destination space | Continue | Same task/risk/destination root when explicitly selected |
| Destination space unknown | Continue | Same task/risk/destination root when explicitly selected |
| EFS plaintext destination | Continue as plaintext / Keep source | Same task/risk/destination root when explicitly selected |
| Sparse allocation inflation | Continue | Same task/risk/destination root when explicitly selected |
| Placeholder hydration/recall | Download and continue | Same task/risk/destination root when explicitly selected |
| Security-significant metadata loss | Move anyway / Keep source | Same task/risk/destination root when explicitly selected |
| Recycle escalation | Delete permanently / Skip | Exact item only; never Apply-to-all |

Each accepted choice creates a bounded receipt containing risk, task ID, optional item index,
destination root, optional source/destination identity, and decision. The task retains at most 64
receipts; exhaustion fails closed. Receipts are task-local, never session-sticky, never Preferences,
and never transferable to another risk/root/task. Recycle escalation is always item-scoped even if a
caller requests broader scope. An explicit eligible Apply-to-all acceptance creates one receipt for
that exact task/risk/destination-root grant; later matching items consume the cached scope and do not
pretend to have separate per-object receipts. Recycle never uses this scoped cache.

## Overlapping-operation interlock

Every governed scope carries one access role:

- `ReadSource` for Copy and Copy-only Move sources;
- `WriteSource` for Delete, the same-parent Rename envelope, and Native or managed Move sources;
- `PublishDestination` for every resolved Copy/Move final destination leaf. Rename's destination leaf is inside its
  mandatory same-parent `WriteSource` envelope, so a second weaker scope is unnecessary.

Two overlapping `ReadSource` scopes may run concurrently. Every other overlapping role pair is a
same-host overlap that the user decides (R4-A02-1, `FO-OVERLAP-01`): during Preparing, once the
task's scopes exist, the host publishes that immutable scope vector under the queue lock before it
compares with every other published-prepared or active task using this same predicate (no
enumeration, no provider call). Activation transfers the registration from prepared to active under
the same lock. Prepared peers are compared in an acyclic order: a task names every older prepared
peer, and a newer prepared peer once that peer has decided its own advisory and does not already
yield to this task, so an older task that publishes after a newer peer prepared still learns about
it and queues after it, while no two tasks can queue after each other. Only a peer whose advisory is
still pending at that instant stays unnamed; the interlock below still serializes that pair. A
Queue-after edge holds in both admission modes: the parallel path waits on it, and queue mode admits
the first queued task that no still-published predecessor blocks rather than the queue head alone, so
a task that queues after a peer behind it lets that peer go first instead of holding the queue; such a
task also takes a fresh queue place behind every task admitted so far, so the visible order matches. The host
raises one warning on the task card naming the
concrete problem and the first overlapping task (this task may remove what task N reads; both may
create or replace the same names; both may remove or rename the same members; this task may remove
what task N publishes; this task may replace what task N reads). **Queue after these tasks** (the
default) keeps the admission wait below as the transient edge, released without a second warning once
the named tasks terminate; **Run at the same time** admits the task concurrently only when the full
disclosed relation set fits the fixed 64-task bound; **Don't start** (Escape) ends the task as a
Preparing cancel with no mutation, clipboard consumption, or breadcrumb. Run stores only those exact
task IDs and bypasses the root interlock only while those tasks are concurrently live. An undisclosed
overlap still waits or raises its own warning, and the relation is discarded with the task rather
than reconstructed from completed history. A task queued by an overlap edge offers no `Start now`. Inline
F2 (`RenamePlan(InlineRename)`) keeps honoring the interlock without a prompt, as its completion
exception states. The warning names only overlaps the paths prove: a scope in a conservative
identity domain (a provider without bound objects, where the interlock treats every
conflicting-access pair as overlapping) keeps the silent wait it always had and raises none. The
non-Run admission wait itself is unchanged: a destructive or publishing task does not
rename, remove, replace, or publish through a
tree while another task is reading it. Endpoint/path-profile containment and retained no-follow
root-to-ancestor identity remain the overlap authority. Role-aware admission must not increase
retained authority beyond the existing bound of unique governed provider paths and path depth.

An absent final leaf is represented by its nearest exact existing ancestor plus the normalized
relative suffix from that anchor. Equal or ancestor/descendant suffixes under the same exact anchor
overlap; distinct sibling suffixes do not. This comparison is alias-safe across alternate path
spellings because the anchor, not the pathname spelling, supplies authority. Binding
Indeterminate or provider-contract violation rejects the task before mutation. Unsupported identity
uses a conservative provider/profile identity-domain scope and serializes potentially aliased
writers rather than becoming a path-only scope.

Admission precomputes normalized suffix keys outside the queue lock and indexes exact retained-anchor
identity domains during the overlap check. Within one exact anchor group, a sorted ancestor-stack scan
detects equal and ancestor/descendant suffixes while preserving role compatibility: read/read remains
admissible and every overlapping writer/publisher pair serializes. Conservative, legacy, or otherwise
unkeyable scopes fall back to the exact Cartesian overlap predicate; the index never turns incomplete
identity into path authority.

Successful Copy, Move, Rename, and directory-shell publication records the actual provider path in
a host-local live-publication index while the publishing task remains active. Each entry pairs that
path with the task ID and the immutable planned scope that contains it: a transfer's
`PublishDestination` scope, or for a rename plan, which holds only its parents as `WriteSource`
scopes, that parent scope. Exact duplicates are ignored, the table holds at most 256 entries, and all
entries for a task are removed atomically with its active interlock registration at terminal state.
A task that has already fallen back to overflow, or a full table, records nothing further and takes
no lock for later publications, so a large operation pays nothing per item beyond the first 256.
This is transient scheduling state, not deletion authority, provenance history, or a
cross-process/host guarantee.

Before Delete, recycle escalation, source-destructive Move, Rename, or an overwrite/replacement
invalidates an item, the existing per-item mutation seam compares that exact provider path with live
published paths. Same-root comparison uses component-aware equality or ancestor/descendant
containment, so distinct siblings do not conflict. Alternate spellings use only the frozen retained
anchors and scopes as conservative positive evidence. The check performs no descendant enumeration
and no provider, network, or device call solely for correlation. A still-live exact Run relation
covering the publisher allows the effect to continue through the ordinary authority and conflict
rules. Without such a relation, the item parks at **Queue until the other task finishes** (default),
**Skip**, the operation-specific destructive action, or **Cancel**. Queue waits for only the named
publisher, then repeats the guard and resumes the ordinary current-membership/identity
revalidation. Skip is reported as an explicit skipped item. Destructive consent covers only that
parked effect, is never cached or Apply-to-all, and does not weaken exact-object, publication,
verification, Retry, cleanup, or immutable-result rules.

If a publication cannot be associated with a planned scope or the 256-entry table fills, overflow
stays sticky while any affected publisher remains active. Correlation then falls back to the already
frozen active `PublishDestination` scopes, and for a rename publisher its `WriteSource` parent
scopes, and deliberately over-protects possible aliases; the
destructive action is withheld, while Queue, Skip, and Cancel remain available. Terminal cleanup
removes the affected task's overflow contribution and wakes exact-publisher waiters.

## Inline F2, Batch Rename, and Change Case

Inline F2 submits `RenamePlan(InlineRename)` to File Operations workers. It never calls provider
`RenameItem`, nor the provider's child-name contract, on the UI thread and shows no second initial
confirmation: admission proves the leaf is not empty and spells no accepted separator (a lexical
check that needs no provider) and that the source has a parent, then publishes the plan with its
provider join and collision key pending (the operation-artifact touch guard
classifies that pending destination by the host join of the parent and leaf, as it does a
transfer's implicit destination); Preparing queries the contract on the task thread (the same step
that admits transfer leaves), fills both from the answer, or fails the task on its card with the
provider's reason and a `preparation.destinationName` diagnostic before any interlock or mutation;
execution keeps its own re-query against that admitted join.

If the host rejects the request before publishing a task, FolderView shows a localized pane-scoped
operation error because no task card exists. After successful task publication, failure presentation
belongs only to the central task UI; FolderView does not add a duplicate pane alert.

- Inline F2 bypasses global Queue waiting but still honors overlapping-root interlocks.
- Before mutation, the worker retains exact no-follow authority for every governed root and its
  provider ancestors. Roots overlap when their qualified endpoint/path profiles overlap lexically or
  when retained authority proves that an alias reaches the same root or ancestor object. Common
  root-to-ancestor chains are interned and shared within a task, so retained authority grows with
  unique governed provider paths rather than selected sibling count or recursive tree size.
- Cancellation or task teardown wakes an interlock waiter; a waiter canceled before ownership never
  calls the provider mutation boundary.
- Waiting, conflict, failure, partial, cancel, or indeterminate reveals its task card immediately.
- A clean-running card is delayed by the 500-ms product constant
  `kInlineRenameCardRevealDelayMs`; it is not a Preference.
- Clean completion before reveal creates no card/group row and does not activate the popup. It still
  emits telemetry, refreshes the pane, and retains focus by identity.
- Once revealed, ordinary card and auto-dismiss rules apply.
- Tests use a fake clock or zero delay and never sleep.

When Inline Rename encounters an existing destination, any destructive action is offered only with
an exact no-follow destination receipt that includes the provider's rename/publication authority.
Metadata-only binding is insufficient: accepting Overwrite/Replace and then receiving the same stable
conflict again is a contract failure, not a reason to reprompt indefinitely. The worker either consumes
that receipt in one conditional `RenameIfUnchanged` attempt, obtains a fresh receipt after a proven
replacement race, or terminates without mutating either object.

Case-only Local rename uses an exclusive cryptographically unpredictable temporary sibling. Entropy
failure or a raced existing temporary name stops before the source is moved; the temporary hop never
uses replacement semantics against an unowned object.

Batch Rename, including one-step plans and F2 that opens Batch Rename, uses ordinary Queue/Parallel
and card behavior. Preview defines final mappings; File Operations owns execution, conflicts,
dependency scheduling, cancellation, and result axes.

Change Case keeps only iterative, provider-owned discovery outside the mutation engine. The discovery
worker binds each selected root no-follow to distinguish files from directories, completes the full
recursive traversal before any mutation, and returns the exact provider HRESULT when a selected or
discovered directory cannot be enumerated. It never treats a failed selected-directory read as a
file/directory discriminator and never omits a subtree as success. Planning validates provider joins
and collision keys and emits immutable `BatchRenameExecutionOp` mappings; production then submits
those mappings as `RenamePlan(ChangeCase)` through the same scheduled-rename admission, identity
binding, acyclic schedule, artifact guard, conflict policy, conditional mutation, and typed terminal
result owner used by Batch Rename. `FileSystemRenameBatch` is not a production Change Case mutation
authority.

The Change Case discovery card is informational only while the command-specific traversal is active.
Once mappings are published, the central File Operations task owns Queue/Parallel behavior, popup
presentation, cancellation, partial results, and completion refresh. Successful completion restores
focus only when the pane still shows the captured folder; every terminal central result refreshes the
origin pane. Change Case introduces no recovery journal, Resume/Roll back surface, command-local
name policy, implicit Keep Both choice, or alternate success definition.

Every Batch Rename row is path-qualified by the worker after ordinary task publication and before any
mutation, and all rows in one plan must share one endpoint/root/profile. Worker admission stores
comparison-only ingress identity for each changed source and atomically publishes an immutable
acyclic dependency schedule. Execution cross-checks that identity when it acquires the exact
no-follow source authority. A cycle is rejected with `ERROR_CIRCULAR_DEPENDENCY` before the first
provider mutation; the host does not invent a temporary hop.
The dependency scheduler is pure: it orders engine-owned conditional rename calls but cannot invoke
provider `RenameItems`/`RenameItem` directly. Window close/cancel signals the published task and wakes
qualification as well as execution waits, while a
late generation-scoped completion may update the ordinary task card but cannot resurrect a closed or
superseded Batch Rename window.

Every scheduler callback is an exact mutation boundary. Before each final rename, the task revalidates
the retained source object and its retained parent scope. A
successful provider call must return the bound authority for the renamed object; the task cross-checks
that immutable object id before allowing a later dependency step. A successful namespace mutation
without valid returned authority is `Indeterminate` and is never retried. A user Skip is
`NotPublished / Retained / Skipped`, not a successful rename.

Rows admitted to one runnable dependency layer are independent. Unless cancellation is requested, the
engine attempts every row in that current layer even when an earlier sibling fails or is skipped, and
every result remains truthful. A failed, skipped, canceled, or indeterminate row never vacates its
source namespace and therefore never unlocks a later dependency layer that targets that
namespace. The scheduler stops before those dependent mutations; it does not turn one row's failure
into permission to overwrite or rename through a retained source.

Batch Rename has no durable journal or recovery executor. It never creates, reads, enumerates,
migrates, acknowledges, or deletes Batch Rename journal files, including bytes left by older builds.
The File Operations UI exposes no Batch Rename Resume or Roll back action. Its copyable undo-plan
report is informational only and grants no mutation authority.

## Cancellation and lifetime

Task cancellation or stop exits every host pause/conflict wait even while Queue mode
or a peer-owned conflict remains active. A cancelled queue-paused transfer drains
independently of its predecessor: the predecessor may remain user-paused, and the
user need not resume it or change Queue/Parallel mode. Waking a condition variable
without exiting on cancellation is not a quiet-point guarantee. Stop/cancel wakeups
must synchronize with each wait predicate's mutex after publishing the stop state,
so a worker between a false predicate and sleep cannot lose the notification.

File Operations admits mutation only after the exact path-and-operation capability response
classifies its complete route as `bounded` or `providerWatchdog`. `bounded` carries a zero provider
watchdog timeout and means every provider call and reader/writer used by the admitted route has a
deterministic tested quiet point. `providerWatchdog` carries a nonzero provider-owned timeout and
requires tested backend cancel, quarantine, and unload gating. `uncontained` carries a zero timeout
and is rejected before task/card/worker creation; this disables only mutation and does not disable
browsing or inspection. Same-provider work and both endpoints of cross-provider work pass the same
gate. Operation booleans and transfer import/export lists cannot override it.

Cancel changes an admitted task to Stopping, admits no new work, and signals the route's proved
Abort/deadline/provider-watchdog mechanism. Owned stage cleanup uses retained identity only. The host
never synthesizes a join deadline, abandons or terminates an in-process worker, or unloads a module
under an active callback. Local fixed-volume and deterministic Dummy/7-Zip routes are `bounded`;
Local UNC/mapped-remote paths are the distinct `local-win32-smb` profile and are `bounded` as well
through the synchronous-I/O cancel watch below (R0f-SMB); MTP is
`providerWatchdog` with its provider-owned 30-second watchdog; FTP/SFTP/SCP are `providerWatchdog`
through libcurl's progress callback polling the operation control plus the transport timeouts
(R0f-Curl; IMAP stays read-only and uncontained); S3 is `providerWatchdog` through the CRT
continue handler polling the operation control plus its bounded connect, stall-monitor, and retry
policy (R0f-S3; S3 Table stays read-only and uncontained); Microsoft Drive/SharePoint is
`providerWatchdog` through operation-control polling around every WinHTTP step plus the
resolve/connect/send/receive timeouts (R0f-Graph; Graph Delete is Recycle by stable item ID);
Google Drive is `providerWatchdog` through the libcurl progress callback polling the operation
control on every transfer plus the connect and hard request timeouts (R0f-GDrive). Every network
route R0e left `uncontained` now carries a route-specific proof; R5 process isolation stays a
fallback only.

`uncontained` is a transitional classification, never a supported end state. As of R0f every
supported destination above reads, writes, creates, deletes, and renames through File Operations;
the routes that remain `uncontained` are read-only by design (IMAP mailboxes and S3 Tables). Should
a route regress to `uncontained`, the host refuses its mutation before task creation and tells the
user why, browsing and inspection remain available, and the refusal is a defect against R0f, not a
capability decision.

The admitted fixed-volume Local regular-file source route has a tested quiet point inside a pending
kernel read. Its exact bound reader uses an independent overlapped request, polls the task operation
control while that request is pending, and on Cancel/deadline calls `CancelIoEx` for that exact
request and drains its completion before returning the original `ERROR_CANCELLED`/`ERROR_TIMEOUT`.
No bytes from the canceled request are reported. Normal EOF remains `S_OK` with zero bytes, and
reader seek state is private to that reader. This closes the Local source-read part of the current
UI-thread shutdown join; it does not claim writer cancellation, generic provider cancellation, a
host join deadline, worker abandonment, or containment for any route still classified
`uncontained`.

Application exit drains cancellation off the UI thread. After the cancel-all confirmation the host
cancels every task and defers the close instead of destroying the window: the window keeps pumping,
the popup shows each task Stopping, and the close resumes from the completion path once the last
task has been reaped there. An Exit with nothing active closes at once. The `WM_DESTROY` join
remains the final ownership boundary and finds nothing to wait for on that path; it still runs
synchronously when teardown is forced (end of session, a direct destroy). No worker is detached,
killed, or abandoned to make a close faster.

The Local `local-win32-smb` route (UNC paths and mapped remote drives) is `bounded` by the
synchronous-I/O cancel watch (R0f-SMB). Every File Operations worker thread and every Local provider
scheduler worker registers itself with a per-module watch for the duration of its provider work. The
watch polls the task's cancel/stop state (host workers) or the unit's operation control (provider
workers) every 50 ms and, once that state has requested cancellation for 500 ms without the call
returning, cancels the thread's pending synchronous Win32 call with `CancelSynchronousIo` on every
poll until the call returns. The grace lets cooperative checkpoints and owned-stage cleanup finish
untouched on a live share; a call still wedged after it (a dead share holding an open, enumerate,
write, rename, delete, flush, or close) fails with `ERROR_OPERATION_ABORTED`, which a canceled task
reports as the cancellation itself and never as a provider fault: the durable cancel request, not the
status the execution path happened to return around it, decides that at every result seal, so such
an item ends `Canceled` and the task ends `ERROR_CANCELLED`. The overlapped source reader keeps
its `CancelIoEx` quiet point. Nothing is abandoned, terminated, or unloaded: the wedged call returns
to its caller and the item carries the same receipt truth as any other canceled item
(published/retained/removed per root, owned stage removed or retained). Fixed-volume Local routes
carry the same watch; their calls complete long before the grace, so their behavior is unchanged.

A UNC share root (`\\server\share`) is the topmost container of its route: the host's provider
parent derivation reports no parent for it (mutation-interlock ancestors, rename and create
admission stop there instead of binding `\\server`), and the Local provider opens a share root with
the trailing separator Win32 requires (`\\?\UNC\server\share\`); deeper paths are untouched.

Once application exit is committed, main-window command dispatch is fenced before auxiliary-window
teardown or the final main-window destroy is queued; no keyboard, function-bar, menu, or posted
command may start another operation behind that fence. A synchronous admission/confirmation prompt
that dispatches the accepted close MUST stop its nested message drain as soon as the prompt completes,
its overlay becomes hidden, or application prompt shutdown begins; it then returns cancellation to the
admission caller and unwinds that caller before dispatching any queued owner-destroy message. Admission also fails with
`ERROR_SHUTDOWN_IN_PROGRESS` if the File Operations state has begun shutdown, including immediately
after every synchronous confirmation boundary. Destroying `FileOperationState` underneath an active
admission stack is forbidden.

Every custom File Operations nested prompt follows the same shutdown fence. In particular, the
custom speed-limit prompt checks application prompt shutdown before entering and after every
dispatched message, clears its pending result, and unwinds; it cannot keep draining arbitrary thread
messages after owner teardown begins. The speed-limit and both Cancel-All confirmation paths also
enter the File Operations prompt-dispatch lifetime scope before pumping messages. Shutdown may cancel
and join current work, but destruction of `FileOperationState` is deferred until the outermost scoped
prompt frame unwinds. A speed-limit prompt captures only the task ID and initial scalar value; after
the pump it re-finds the task, and a missing/shut-down task makes submission a no-op. Holding a raw
task pointer across any nested prompt pump is forbidden.

The retired `quietPointTimeoutMs` capability is invalid. The host does not copy a decorative timeout
onto an endpoint or claim it can bound a provider call. Only a provider-watchdog route owns timeout
handling: its timeout yields an indeterminate result and provider quarantine while preserving
possible artifacts and sources. Verifying and Stopping count as live I/O for app-exit policy.

### Completion posting and lifetime

Completion applies and removes tasks only on the owning UI thread. The worker
first posts a registered `TaskCompletedPayload`. If that fails, it queues the
payload under the fallback mutex and posts a payload-less wake carrying the
task ID; the UI adopts the matching queued completion. Neither `Task::ThreadMain`
nor a threadpool worker may call UI completion/removal or synchronously send
completion to a UI thread that may be waiting for that worker.

If both posts fail, an admitted reaper takes ownership of the exact task's
`jthread`, joins it, then posts only that task's UI wake. Handle transfer,
callback admission, and failure restoration share the transfer mutex with
shutdown's task-worker scan. Failed admission restores the `jthread` to `Task`
before the final payload-less wake, so UI removal retains the required join
ownership. The fallback completion remains queued until UI adoption or shutdown.

Shutdown closes admission and waits for the outstanding reaper predicate before
destroying `FileOperationState`. A callback decrements and notifies under that
predicate's mutex as its final access to the state; observing zero after
reacquiring the mutex is the lifetime barrier. This preserves the forced-destroy
ownership boundary and the normal off-UI cancellation drain described above.

## Results and consumers

Per-item executable axes are `PublicationState`, `VerificationState`,
`SourceDisposition`, `OwnedStageDisposition`, and `ItemCompletion`.

Mutation commitment determines publication/source/stage truth, not the user-visible completion axis.
If publication committed and a later phase fails, the item remains `Published` but is `Failed`; if the
user cancels after a partial commit, it remains `Published` but is `Canceled`. Conversely, a committed
mutation cannot by itself erase a terminal failure or cancellation HRESULT. Consumers use the typed
axes together and never equate `mutationCommitted` with `ItemCompletion::Completed`.

`OwnedStageDisposition` is independent of final-destination publication and is one of
`NotApplicable`, `NotCreated`, `Owned`, `Removed`, `Published`, `Retained`, `Unknown`, or
`RetainedIncomplete`. For example, a failed pre-publication copy can be exactly `NotPublished /
source Retained` while the owned stage is `Unknown`; a direct-final Local Copy whose exact abort is
known not to have completed is `NotPublished / source Retained / RetainedIncomplete`.
`Retained`, `Unknown`, or `RetainedIncomplete` stage truth blocks automatic retry and Move source
deletion and never rewrites exact final publication to Unknown. `Unknown` and `RetainedIncomplete`
map the item to Indeterminate/996 and surface a Possible artifact diagnostic. Exact `Retained` is
independent cleanup debt: a committed primary remains Completed/Published and receives one localized
`Completed; cleanup item retained.` warning, while a failed or canceled primary keeps that terminal
status. `RetainedIncomplete` is reserved for a known surviving provider-created object that does not
contain the complete requested payload; ordinary known cleanup debt remains `Retained`.

| Example axes | Summary | Presentation |
|---|---|---|
| Published + Verified + Removed | Moved and verified | Success plus verified mark |
| Published + NotRequested + Removed | Moved | Success |
| Published + verification non-success + Retained | Copied; source kept + detail | Warning/partial |
| Published + Retained for planned Copy | Copied | Success |
| Published + Retained for Move downgrade | Copied; source kept + reason | Warning/partial |
| Published + committed primary + owned stage Retained | Completed; cleanup item retained. | Warning with primary success preserved |
| Partially published recursive destination + Retained + Failed | Copy/move incomplete; source retained | Error/partial |
| NotPublished + Retained + Skipped | Skipped + reason | Neutral |
| NotPublished + Retained + Canceled | Canceled before publication | Neutral/canceled |
| NotPublished + RetainedIncomplete + Indeterminate | Incomplete visible destination retained | Error/Possible artifact |
| Publication or source Unknown | Outcome unknown | Indeterminate |

Aggregate mapping:

| Condition | HRESULT / task state |
|---|---|
| Every item fulfilled requested intent | `S_OK` / Completed |
| Only intentional skips or preplanned Copy-only outcomes | `S_FALSE` / Partial or Completed-with-notes |
| Publication succeeded but requested verification is unavailable | `S_FALSE` / Completed-with-notes |
| User cancel without indeterminate state | `ERROR_CANCELLED` / Canceled or Partial |
| Mixed success and unexpected failure/retention | `ERROR_PARTIAL_COPY` / Partial |
| No mutation and one exact failure | Exact HRESULT / Failed |
| Any indeterminate publication/source | `ERROR_IO_INCOMPLETE` / Indeterminate or Partial |

Source rows are removed or ghosted only for `SourceDisposition::Removed`. Retained remains and
refreshes. Unknown refreshes without optimistic removal. Destination refreshes for Published or
Unknown. The same mapping applies to FolderView, Find, Compare, search index, and watcher consumers.

The engine stores one typed result for every selected top-level source before completion is posted.
Provider mutation truth and host-owned Managed publication/cleanup truth feed that record; a missing
destructive result is `Unknown`, never inferred from an aggregate success code. The aggregate
`HRESULT` also preserves known host-bridge truth: when some recursive children were published and a
later child failed, the selected item is `Published` / `Retained` / `Failed` with
`ERROR_PARTIAL_COPY`, not an indeterminate outcome.
The Local provider reports a typed receipt on every Copy/Move completion: a failure
that never created a destination stage is known no-commit (`NotCreated`); a merge
that already published a child is known committed; a refused native rename is known
no-commit and remains Retry-eligible. User Cancel with a proved no-commit receipt
stays `Canceled` and does not become `ERROR_IO_INCOMPLETE`. A cancellation with no
receipt at all is still `Canceled`, with publication `Unknown` rather than
`NotAttempted`. An explicit unknown receipt remains indeterminate.
A successful provider-native Delete or Native Move without a valid mutation receipt is classified
before the generic success branch as `ItemCompletion::Indeterminate` with
`ERROR_IO_INCOMPLETE`. Delete records `PublicationState::NotAttempted` and an unknown source
disposition; Native Move records unknown publication and source disposition. Copy and a
host-proved Managed/Copy-only retained-source result do not inherit this destructive fallback.
Terminal storage is single-assignment. Retry-loop cancellation, Keep Both path failure,
requalification failure, traversal limit, Skip, Continue-on-error, verification failure, and stage
cleanup failure all flow through the same final store. Neither HRESULT, transferred-byte count, nor
the absence of a provider receipt can overwrite explicit publication or stage truth.
The aggregate `HRESULT`, completed-task counts, Issues entry, telemetry, popup result text, cache notifications,
removal focus, Find row removal, and Compare invalidation are all reduced from those records. There
is no legacy `S_OK`-means-remove fallback. Localized canonical result text is `Copied`, `Moved`,
`Copied; source kept`, `Skipped`, or `Outcome unknown`. Verification is an independent localized
detail/badge: `Verified`, `Verification failed`, `Verification unavailable`, or `Verification
canceled`. A Move verification problem also keeps the source-retained badge/detail; choosing the
verification wording as the primary summary must not hide that source disposition.

### Clipboard consumption

For clipboard Move, capture the full selected list and Windows clipboard sequence. After the whole
request is accepted, initialize one `Preparing` result builder per selected root, publish the task,
and construct its worker. `StartOperation` never waits for provider readiness on the UI thread. The
worker binds every selected-root mutation/interlock scope and completes the unresolved Preparing-gate
hook, then either fails closed or posts one registered readiness payload to the Folder window and
waits. Only successful readiness permits the UI-owned exact-sequence consumer to clear the cut
clipboard. That UI dispatch is teardown-fenced, holds no raw task across clipboard/OLE work, and
re-finds the task by ID before publishing the result. Successful `S_OK` consumption is a one-shot
receipt and releases the worker toward queue/discovery/mutation. Readiness or consumption
failure wakes the worker only into terminal reduction: every selected root is `NotAttempted +
Retained` with the exact failed/canceled status, and no discovery or provider mutation may follow.
Thread/queue admission failure rolls back the unpublished task and sequence receipt without touching
the clipboard. The host records accepted sequence numbers in a bounded ledger atomically with task
publication, so the same cut sequence cannot admit a second Move even when clipboard clearing fails.
Never restore a successfully cleared cut list after failure, cancel, Copy-only, partial, or
indeterminate result. Confirmation cancel, queue failure, readiness failure, or clipboard-consumption
failure leaves the current clipboard content unchanged. Copy clipboard content remains repeatable.

Each selected-root result builder moves monotonically from `Preparing` to `MutationPossible` and then
to one immutable terminal value. Before `MutationPossible`, missing callback truth is known no-
mutation truth (`NotAttempted + Retained`), never `E_PENDING` or Unknown. At and after
`MutationPossible`, a missing receipt is `Indeterminate` with `ERROR_IO_INCOMPLETE` even when the
task status succeeded (Copy keeps `SourceDisposition::Retained`); the one exception is a root whose
status is `S_FALSE`, which is a conflict the user answered Skip before any mutation and is reported
`Skipped + NotAttempted + Retained`. Likewise a Native Move or Delete root whose provider returned
`ERROR_PARTIAL_COPY` after the host answered Skip at its conflict prompt is a user decision, not a
destructive success without proof: it is reported `Skipped` with the source retained and the task
ends `S_FALSE`. Terminal storage is compare-and-store:
invalid indexes and duplicate writes are diagnosed and rejected without changing the first terminal
truth. Every worker completion path performs the same dense final reduction.

The task acknowledges **“Move queued — cut list cleared.”** A consumption failure says that the Move
did not start, the source remains unchanged, and the same accepted sequence cannot be submitted
again. If sources remain after successful consumption, show **“Some sources
remain. The cut list was consumed when the Move started and was not restored.”** with Open source,
Open destination, and Select retained items where known. “Cut retained items again” is an explicit
new action available only when every retained source is known exactly. Select/Cut-again rebind every
retained pathname no-follow at click time and require the captured object identity; missing or
replacement objects abort the whole action. If any source disposition is Unknown, say that its state
is unknown, offer Open source/destination where paths are known, and offer no Select/Cut-again action.

## Progress and popup states

There is no Calculating phase. The popup consumes the explicit engine lifecycle and presents
Preparing/AwaitingAcceptance as **Preparing**; it never infers that state from absent progress.
Ready is a transition boundary before Waiting or Running rather than a synthetic progress phase.
Task presentation also includes Waiting, Discovering, Running, Paused, Verifying, Stopping,
Completed, Partial, Canceled, Failed, and Indeterminate. Discovery-ahead is secondary activity;
after Skip discovery it reads **Discovering as needed**.

- Open traversal: whole-task total and ETA are indeterminate/Estimating in compact and expanded
  presentation; discovered/completed/current counts remain exact where known, but growing totals do
  not create a determinate compact fraction.
- Closed traversal: freeze the verification denominator from the same discovery totals and compute
  ETA as remaining transfer/publication time **plus** remaining verification time. Do not divide the
  combined bytes by a combined rate because the two phases are sequential in this slice.
- Verification uses a second hatched `fileOps.progressVerify` segment, per-file verification
  mini-bars, and `fileOps.graphVerify`; transfer uses `fileOps.progressTotal`.
- The hosted DxUi `ProgressBar` is the exclusive determinate whole-task owner. With verification On
  it receives transfer and proof values as two equal semantic halves plus the verification theme
  color/hatch state. There is no second painter: when the hosted control tree or the Direct2D
  target cannot be created, the popup shows its native failure surface
  (`../UI/UI_FileOperationsPopup.md`, "Failure Surface") with the explanation and Cancel all, and
  draws no progress. The hosted control exposes the same two-axis state to accessibility.
- The primary graph series and pump throughput contain transfer bytes only; `fileOps.graphVerify`
  is the separate verification series. The current/aggregate bandwidth marker and configured
  bandwidth budget count both.
- Verification off creates no verification reads, graph samples, or artificial phase.

The existing More menu, Pause/Resume, Cancel, Start now/reorder, Issues pane, completed grouping,
filters, auto-dismiss, speed controls, graph history, and task diagnostics survive. UI details are
owned by `Specs/UI/UI_FileOperationsPopup.md`.

## Artifact warnings and retained state

`FileOperationDurableStore` is the app-local mechanics boundary for retained File Operations state,
including the interrupted-Move breadcrumb. It provides configured bounded reads,
no-follow regular-file validation, final-root identity checks, handle-based direct-child enumeration,
atomic `LocalFileTransaction` publication, and exact-handle direct-child removal. It is not a generic
recovery engine: JSON fields and versions, schema ownership, and Move notice projection remain in
their separate owners. Batch Rename is not a durable-store consumer and the helper never enumerates
the retired Batch Rename journal root on its behalf. Concurrent readers are supported by immutable
old/new file publication; each retained state instance still has exactly one schema-authorized writer. Bounded readers open with
read/write/delete sharing, and replacement uses `LocalFileTransaction`'s handle-based POSIX rename
when supported so an already-open reader remains bound to its complete old generation while a new
reader opens the complete replacement. Read and acknowledgement reject a final reparse entry.
Replace publication instead atomically replaces that final directory entry without following it;
a junction/symlink target is never read, written, or removed by the durable-store helper.

Artifact classification is an orthogonal presentation/safety axis. It MUST NOT hide an item from a
pane, Find, Search, Compare, enumeration, or an ordinary user-selected query/filter result.

| Class | Evidence and authority | Presentation |
|---|---|---|
| Possible | Current/legacy generated artifact-name shape only; never an ownership claim | Always visible with `Possible interrupted-operation artifact`; Inspect/Reveal/Open location remain ordinary inspection |
| Ordinary | No artifact provenance evidence | Normal user item |

### Artifact source of truth

Artifact classification is deliberately non-authoritative. The registry creates one empty in-memory
generation and never opens or enumerates the retired Batch Rename journal root. Only the recognized
generated-name shapes `.rs_tmp_`, `.rs_copy_tmp_`, `.~rs-write-`, `.rs_bak_`, and `.rs_ren_` produce
`Possible`; all other names are Ordinary. Neither a name shape, pane row, search-index field, badge,
nor an in-memory task object proves app ownership or grants cleanup/recovery authority.

Folder enumeration and Find load one empty classifier snapshot per request. After Batch Rename
journal retirement there are no production durable claims and therefore no endpoint/parent claim
index: only `HasPossibleArtifactName` can admit a row to the projection probe. Ordinary names do not
perform path-capability, identity, or bind work. A recognized artifact-name candidate is bound
no-follow only to capture its exact current identity for presentation and the later touch guard; that
identity does not upgrade the name-shape classification into ownership. FolderView renders a warning
glyph in every display mode; Compare inherits the same projection through its FolderViews. Find gives
matching result rows the localized Possible warning badge. A snapshot or projection failure never
hides a row and never grants recovery or cleanup. A future durable claim source must define and prove
its own authoritative schema before any keyed candidate index can return.

### Visibility and touch guard

Possible and Ordinary items participate in ordinary enumeration, sort, user filters, name search,
content search, and Compare decisions. Possible classification adds a warning overlay/badge and the
reason `name shape only`; it is never an artifact-specific suppression rule and carries no task,
phase, intended-name, durable-record, cleanup, or recovery metadata. Properties, selection,
read/open/edit/preview, Copy/export, Inspect, Reveal/Open location, configured viewer/editor and user-menu
launches, and Shell-owned surfaces remain available without a prompt solely because of the name.

A RedSalamander-owned mutation of a Possible item is a guarded touch: Move, Rename, Delete/Recycle,
overwrite of an existing object, attribute/content/metadata writes, and use as an existing publication
target. Before admission, show one warning that the name resembles an interrupted-operation artifact
and that continuing may alter incomplete data. The safe default is Cancel. A Continue grant covers
only the exact no-follow identities in that accepted request, is never Apply-to-all or persisted, and
is revalidated immediately before mutation; an identity mismatch invalidates the whole request and
fails closed. Ordinary items are not copied into the receipt. If a Possible item cannot supply a
current exact identity, Continue cannot mint a receipt. Consent never upgrades a name-shape match into
ownership and never authorizes automatic cleanup or recovery.

The central typed `StartOperation` admission is the first executable touch-guard owner. It projects
Move sources and existing publication targets from the immutable child plans; Copy sources are
read-only and are not candidates. It binds only name-shape candidates no-follow, presents one
localized exact-set warning with `Cancel` as the default and `Continue` as the affirmative action, and
stores only an ephemeral task receipt. The
worker rebuilds the empty registry generation and reclassifies/revalidates the complete exact set after queue/interlock
waiting and immediately before operation I/O; any set, endpoint, path, identity, or availability
change fails before mutation. Missing publication names are not artifact objects and remain under
the ordinary exclusive-create/conflict contract. FolderView and Find default Shell Open/Open With,
configured external viewer/editor launches, user-menu programs, the current-directory Shell context
menu, and the Shell Security property page do not infer RedSalamander mutation authority and do not
apply a name-only warning. Internal viewer/VFS inspection is likewise prompt-free.
Non-recursive Change Attributes guards and immediately revalidates its selected exact set after its
options are accepted and before the first synchronous mutation. It retains the accepted receipt and
reopens/rebinds each guarded selected object no-follow immediately before changing that object.
Item Properties performs the same exact warning before deleting a named stream, and Cancel leaves the
stream and base object unchanged.
Change Case discovery never mutates. Its central scheduled-Rename task constructs one exact artifact
request from the complete immutable mapping, marshals the decision to the UI thread, retains the
accepted receipt, and revalidates it before central execution. Each rename step additionally consumes
fresh no-follow source authority and the provider name contract at its mutation boundary. Cancel
starts no mutation.

Every Folder-window File Operations admission or external-touch entrance that can invoke this
artifact warning holds the Folder window's prompt-dispatch lifetime scope until the nested
`HostShowPrompt` message pump has completely unwound. Folder-window shutdown may immediately cancel
and join File Operations work, clear presentation state, and mark destruction pending, but it must
not destroy `FileOperationState` while one of those entrance frames remains on the UI-thread stack.
This applies equally to direct `StartOperation` admission, the posted Batch Rename artifact prompt,
and direct RedSalamander mutation touch guards. If shutdown occurs during the prompt, the
entrance returns `ERROR_SHUTDOWN_IN_PROGRESS`, publishes no task, and performs no mutation. The
outermost prompt-dispatch scope performs the deferred destruction after the prompt stack unwinds.

Recursive Change Attributes and Change Case use their existing iterative one-pass discovery. After discovery/planning and
before the first mutation, the worker synchronously marshals one exact-set warning to the UI thread
and remains paused until the user decides. An accepted receipt stays worker-owned and ephemeral.
Recursive Change Attributes reopens/rebinds each guarded object immediately before that object's
attribute/time/stream mutation; recursive Change Case delegates the accepted mapping and receipt to
the same central scheduled-Rename authority described above. Already-mutated receipt members are not
reopened through their old pathname. A missing object, replacement identity, endpoint/profile mismatch, blocked
identity probe, or canceled warning stops the remaining command before another mutation. Direct
internal mutation commands outside `StartOperation` are never implied covered merely because a
related File Operations ingress is guarded.

### Interrupted Move breadcrumb

Every Move that reaches Ready owns one compact, versioned breadcrumb under the application-local
File Operations state directory. The worker publishes it atomically with fail-if-exists semantics
after common preparation and every required acceptance/clipboard receipt, but immediately before
queue/execution admission, discovery, or provider mutation. Failure to create the initial breadcrumb
terminates the task with dense no-mutation/source-retained truth. A preparation, decision, clipboard,
or cancellation failure creates no breadcrumb. The filename contains the task ID plus a CSPRNG
token; later phase writes atomically replace only that exact record.

The record is bounded independently of selection/tree size. It contains only:

- schema/kind, task ID, creation time, and durable phase (`Admitted`, `Executing`, or `Terminal`);
- admitted top-level Native, Managed, and Copy-only strategy counts;
- the total qualified source-root count and at most 16 fixed source-root samples, with an explicit
  truncation bit; and
- one qualified destination location; and
- presentation-only source/destination pane-role hints for restart navigation.

Each qualified location contains plugin/instance/profile/root identity plus one representative path
for user navigation. Its `instanceId` is the comparison-only canonical File Operations identity:
exactly `host/default`, or `host/context/<opaque-context>` with a nonempty opaque suffix. Restart
navigation decodes those forms to the ordinary pane instance context (empty for `host/default`, the
opaque suffix for `host/context/...`); it never passes the canonical identity text as a navigation
context. Any other form invalidates the record. Pane-role hints choose which pane receives an explicit
Open source/destination action after restart; they are not provider, identity, containment, or
mutation authority. The record
contains no per-item mapping, child enumeration, object identity,
revision, deletion receipt, publication token, conflict grant, or recovery instruction. Its size is
therefore O(1) with respect to recursive descendants and is capped at 128 KiB. A root count/sample/
truncation mismatch, zero admitted strategy count, malformed UTF-8/JSON, unknown schema/phase, or
missing required qualified field invalidates the record.

After entering Running and immediately before discovery or provider I/O, the worker advances the
breadcrumb once to `Executing`. A failed phase write leaves the last durable phase unchanged and does
not become mutation authority. Known terminal completion first persists `Terminal`, then removes the
JSON record. If terminal persistence/removal fails, the stale record remains deliberately and may
produce a conservative restart warning.

At startup, the host enumerates the exact breadcrumb root through a no-follow directory handle, caps
the scan at 4,096 total direct children, and considers only regular non-reparse `*.move.json` children.
A store that exceeds that cap withholds every notice and logs one warning naming the directory.
Each read is capped at 128 KiB before allocation and verifies a stable complete file size. Valid
non-terminal records become interrupted-Move warning cards. Valid terminal records may be removed as
stale completed bookkeeping. Malformed, truncated, invalid-UTF-8, wrong-version, directory, reparse,
or oversized records are rejected and left untouched; one rejected sibling does not discard another
valid record. Loading, projecting, navigating, or dismissing a record never opens, deletes, copies,
moves, or otherwise mutates any source or destination object. Dismissal may remove only the exact
regular non-reparse JSON handle opened as a direct child of the still-identical configured breadcrumb
root.

Restart presentation states that the outcome is unknown and that source and destination objects may
both exist. It exposes qualified source/destination navigation in the original pane roles and
ordinary refresh only. The notice has no synthetic failed-item/Issues rows: an indeterminate outcome
does not by itself expose Failed items, Show log, or Export issues. It never
offers Resume, Roll back, source cleanup, destination cleanup, conflict retry, or automatic recovery.
The breadcrumb is not a Possible name-shape projection and cannot grant deletion authority.

The restarted card allocates a fresh in-session `summary.taskId`; that value alone addresses card
actions and dismissal. The historical breadcrumb `record.taskId` is exposed only as the explicitly
named `interruptedOperationId` and rendered as **Interrupted operation ID**, never as an active task
ID. It identifies conservative durable bookkeeping only and implies no Resume, retry, rollback,
cleanup, or mutation authority. Dismissal resolves the session card to its exact breadcrumb path and
acknowledges only that regular direct child; it never derives a filename or provider action from
either numeric ID.

Batch Rename has no selection-proportional mapping journal and no Resume/Roll back executor. Existing
retired journal bytes are outside discovery and remain untouched. Undo/G2 remains deferred and may
mutate only objects proven created/changed by the original task.

## Settings

Host settings:

- `fileOperations.mode`: Queue or Parallel;
- concurrency and bandwidth controls already exposed;
- `fileOperations.verifyAfterCopy`: Boolean, default false.

There are no pre-calc settings. Skip discovery, conflict grants, deferred consent, clipboard
sequence, and task link choice are never persisted as Preferences. Provider configuration may
persist only the Preserve/Skip link default, which applies to Copy; Follow is removed.

## Performance and observability

Parallel Local Copy/Move bandwidth pacing divides the configured byte budget by the
number of copy/move transfer slots actually held, including workers sleeping inside
the throttle. Callback age is not worker-lifetime evidence: a correctly throttled
worker may sleep longer than its progress-callback cadence. When a transfer releases
its slot, the remaining transfers inherit its share on their next callback.

`FileOps.Progress.MaxCallbackDeltaBytes` records the largest per-stream byte delta
observed at the host progress boundary. The deterministic parallel-bandwidth fairness
selftest reads in-flight rows only under their owning mutex and compares candidate
skew against both the shared-only baseline and a one-callback-quantum observation
floor. This floor covers admission/publication phase at the callback boundary; skew
above it remains a regression rather than sampling tolerance.

Each applicable item records qualified topology/profile, strategy, read/write/hash/reread bytes,
pump/commit/publication/verification/metadata/delete/interlock/consent/end-to-end timing, digest CPU,
typed result axes, link policy/action, metadata loss, discovery mode and queue high-water, path/depth
and deferred-link bounds, buffers/workers, abort/deadline/quiet-point result, and final-phase
completion.

The task boundary emits `fileops.operation.strategy` once per non-empty bounded
intent/strategy/topology bucket. Ordinary Copy uses the explicit `Copy` strategy and emits
`copy.copy.{same-root|cross-root}`; Move emits
`move.{native|managed|copy-only}.{same-root|cross-root}`. `value0` is the number of child plans in
that bucket, and `value1` is their selected top-level item count. It contains no provider path and
MUST NOT become a per-item or per-descendant metric.

Interlock preparation emits `FileOps.Interlock.PrepareUs` once per task and exactly three bounded
`FileOps.Interlock.ScopeCount` rows with `detail` equal to `read-source`, `write-source`, or
`publish-destination`. For each role row, `value0` is that role's scope count and `value1` is the
task's total scope count. These metrics contain no provider path or identity and MUST NOT become
per-item, per-ancestor, or per-discovered-descendant emissions. Validation compares the role totals
with the retained-authority-node count and proves that concurrent read/read work remains admitted
while every overlapping writer/publisher pair waits.

Preparation also emits `FileOps.Interlock.BindCount`. Admission emits
`FileOps.Interlock.CheckUs` and `FileOps.Interlock.ScopeComparisons`, with comparison count and
active-candidate count carried as bounded numeric values and no path/identity payload. The
256-source/256-destination scenario proves shared ancestors are interned, distinct sibling leaves
remain admissible, same/ancestor leaves serialize, and exact retained-anchor indexing avoids a full
Cartesian scan. Conservative/unkeyable input must still exercise the Cartesian fallback. The
same-machine Release trigger and accepted indexed baseline are archived under
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-29_022850/` (65,536 comparisons,
2,600,470 microseconds) and `Specs/TestRuns/4cb089111a23/FileOps/2026-08-29_024402/`
(8,704 indexed probes, 2,689 microseconds), respectively.

The deterministic large-selection retention scenario emits
`FileOps.SelfTest.InterlockLargeSelectionRetainedAuthorityNodes` and
`FileOps.SelfTest.InterlockLargeSelectionMaxAuthorityDepth`. It prepares 256 selected local sibling
files without executing their transfers, requires one read scope and one final-leaf publication
scope per selected top-level item, and requires unique retained nodes to remain within selected items plus a
small shared-ancestor allowance. It never enumerates or retains descendants.

Required quantitative gates:

- bridge: at most 16 admitted files and 256 MiB pump buffers;
- bridge pipeline: retain the current overlapped reader/writer pipeline. Same-machine test-enabled
  Release FBP-10 evidence on 2026-08-27 measured the disabled baseline at 130.484 seconds and the
  enabled candidate at 0.250 seconds for four 32-MiB Dummy files with 30-ms chunk latency and an
  8-MiB resolved buffer. This authorizes no additional worker/buffer restructuring and no increase
  to the aggregate ceilings; any future change requires a same-process before/after result with
  equal correctness, cancellation, discovery, and bounded-memory behavior;
- recursive traversal, which now means only the Local provider's own `CopyItem` walker (host
  Copy/Move walks through the bridge and Local permanent Delete keep their frames on explicit
  stacks and report depth and retained listings without a ceiling; see the Traversal ceilings
  paragraph): at most 128 descendant directory edges below the selected root (the root is depth
  0), 4,096 retained work entries, 16 MiB retained UTF-16 queued-path text, and 8 MiB retained
  ancestor/directory metadata; emitted high-water rows must stay within those limits and name any
  limit hit;
- MTP: the public writer retains no private payload buffer and stages to a delete-on-close file;
  WPD upload/device relay buffers are at most 4 MiB, verification uses at most two 4-MiB buffers,
  and the reader owns one reusable request state of at most 8 MiB. The provider emits
  `mtp.writer.private_payload_buffer_high_water_bytes`, `mtp.writer.memory_high_water_bytes`,
  `mtp.reader.reusable_request_bytes`, `mtp.verify.memory_high_water_bytes`, and
  `mtp.transfer.bridge_buffer_high_water_bytes`;
- S3: four known 4 KiB objects reserve at most 16 KiB; multipart remains at most four payloads and
  256 MiB;
- collision/name index: at most 64 MiB for 65,536 names of 240 UTF-16 units; Batch Rename emits
  `batchrename.preview.collision_index_retained_bytes` and retains only provider-canonical keys,
  never a second raw-name copy;
- Skip-discovery reservation release target at most 50 ms excluding an active provider call;
- fake-provider UI cancel acknowledgement target at most 2 seconds.
- Recycle Bin batching is correctness-gated by exact test-enabled route counts, requested/observed
  items, maximum batch size, failures, and fallbacks. Live `IFileOperation` baseline/candidate wall
  time remains archived evidence but is advisory because Shell broker, Defender, and indexing work
  are outside the provider's scheduler boundary.

No universal throughput number is imposed. Debug and test-enabled Release evidence follows
`Specs/Testing/Testing_PerformanceValidation.md`.

## Deterministic verification inventory

Generate the live source and runner inventory with
`.\Tools\Get-TestInventory.ps1 -Format Json`; its canonical artifact is
`.build/SpecInventory/test-inventory.json`.

The File Operations selftest and provider-contract matrix cover:

- direct/case/hard-link, volume-GUID, 8.3, SUBST, drive/UNC, and junction/mount aliases where the
  environment exposes them, plus destination ancestor/final-link swaps and indeterminate
  revalidation;
- bulk Copy/Move rejection, explicit Copy plus pure Native/Managed/Copy-only qualification,
  mixed Local Native/Managed plan splitting, raced-directory same-task requalification, Native-with-bridge
  rejection, Managed exact file/tree cleanup, unavailable per-item cleanup downgrade, and CopyOnly
  source retention without allocating cleanup authority in serial and parallel schedules;
- the interlock access-role matrix, Copy/Copy-only source plus destination scope classification,
  concurrent Read/Read admission, and writer/publisher exclusion on overlapping authority;
- exclusive stage collision/replacement/publication/crash cases;
- cryptographic stage-entropy failure before any exclusive-writer attempt;
- folder merge, type/read-only/link conflicts, Keep Both, and hidden/system descendants;
- Copy/Move link kinds and every literal-link worked case above;
- same-pane duplicate, same-folder Move reject, discovery-ahead and Skip-discovery release;
- clipboard sequence change, queue failure, cancel, partial, Copy-only, and unknown;
- managed-Move source/destination swaps and conditional-delete failures;
- verification and metadata gates;
- Recycle escalation identity and no-Apply-to-all;
- Recycle Bin batch-size `1` bypass, one 384-item sibling batch at size `500`, and exactly three
  256-item sibling batches for 768 inputs through the pane-equivalent `BulkItems` admission mode,
  with no failed observations or per-item fallback; repeat invocations use unique leaf names so
  existing Recycle Bin entries cannot trigger a conflict prompt;
- Inline F2 worker/Queue/interlock/fake-clock presentation and Batch distinction;
- Possible name-shape artifact warnings, S3 markers, null-success provider outputs, distinct
  Recycle/Permanent Delete admission, and the shipped provider capability matrix;
- cancellation through enumeration/read/write/commit/verify/delete/gates;
- bounded queue/depth/path/memory and exact Pack/Unpack `DeletePlan`.
- a deterministic traversal-ceiling stop that preserves and reports a prior completed sibling, plus
  deep long-path and wide-tree retention high-water coverage and post-order directory metadata.

Every destructive race uses test-owned roots or deterministic fake providers. No test mutates an
ambiguous live cloud or device path.
