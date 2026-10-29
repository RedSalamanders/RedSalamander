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
14. A Move refused before mutation at an existing destination whose object binding is Unsupported
    reports that collision and the unavailable safe replacement authority, rather than only the
    generic `ERROR_NOT_SUPPORTED` system text. A separate successful existence read establishes
    this diagnostic fact only; it authorizes no overwrite, source cleanup, or identity/equality
    claim. Missing destinations, failed probes, and unsupported source/ancestor bindings retain
    their original failure classification. The original status and both objects are preserved.
    The extra read runs only on this failure path, on the operation worker. Existing
    `fileops.identity.guard_us` instrumentation is retained, and
    `FileOps.Conflict.DestinationIdentityUnavailable` counts the specialized diagnostic.

## Identity vocabulary

| Term | Meaning |
|---|---|
| Qualified endpoint | Plugin ID, provider instance ID, path profile ID, and root identity |
| Path identity | The provider path/name requested inside a qualified endpoint |
| Object identity | Stable no-follow local file ID or provider object ID |
| Revision identity | Immutable provider generation, VersionId, or ETag where the path profile guarantees it |
| Content identity | BLAKE3, an equivalent exact provider content proof, or the digest the destination service reports for the object its writer published (R3-2) |
| Bound object | Open handle or provider token retaining identity and authorized operations |
| Identity-less location | A Local path the provider opens but cannot identify: `BindObject` returns `ERROR_NOT_SUPPORTED` (`Unsupported`), as on a redirector without `FileIdInfo` (I35 R7.3, D-7.3). An open that fails (`Indeterminate`) is not identity-less, including an open the location refuses with `ERROR_NOT_SUPPORTED`, which the Local provider reports as `ERROR_OPEN_FAILED` |
| Corroboration facts | What a delete by name compares instead of an identity: kind, the compared attributes, size (a file), last-write time (a file; and a folder, where it moves when a direct child is added, removed or renamed), creation time, and the 64-bit file index with the volume serial only when the location reports a nonzero index (corroboration only, never sole proof); a folder adds its pinned tree |
| Pinned tree | What a folder deleted by name deletes, and nothing else: the no-follow manifest of every entry below it (relative path, kind, compared attributes, size and last-write time of a file or link, creation time, and the listed file index; a folder's last-write time is left out because its parent's listing updates it lazily and every change it reflects is an entry of the tree), taken before its card, again after the answer and again on the handle that deletes it; its BLAKE3 digest and counts are part of the folder's corroboration facts, and only its entries are deleted, each after its own facts are compared (I35 R7.3) |
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
the deletion runs (`IDS_FILEOPS_CONSENT_PERMANENT_DELETE_BY_NAME`), as a paragraph of its own after
the card's question (`consentNote`), never inside the counts text. A directory identity names the
directory object; descendants keep the provider's own per-observation rules.

The Local provider also implements `IFileSystemIdentityDelete`, for identity-less objects only (I35
R7.3, D-7.3): `ResolveDeleteIdentity` opens the path no-follow, attributes-only and fully shared and
returns its corroboration facts as text (`lc1|k=…|a=…[|s=…]|w=…|c=…[|v=…|i=…][|m=…|n=…]`; a folder
carries kind, attributes, last-write and creation time but no size, and its pinned tree). A folder is
listed whole through a no-follow listing handle of its own, still fully shared: a link or junction is a
leaf, never entered; the folders of one depth are listed at the same time on the shared File Operations
scheduler (the provider's delete concurrency, at most 8), each through its own no-follow handle; the listing
uses `FileIdBothDirectoryInfo` (`FileFullDirectoryInfo`, without file index, on a redirector that refuses
it). `m` is the BLAKE3 digest of the canonical manifest (every
entry's relative path, sorted ordinally over its upper-case form, with its kind, compared attributes,
size and last-write time for a file or link, creation time and listed file index; a folder's
last-write time is left out, since its parent's listing updates it lazily) and `n` its file, folder
and link counts, so an entry added, removed, renamed or changed anywhere in the tree is a difference.
The listing is bounded at 200,000 entries, 32 MiB of relative names and `kTraversalMaxDepth` levels;
over a bound it returns `HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA)` with a log record and nothing is
deleted. It returns `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)` for any object the location can identify,
so a delete by name never replaces exact authority, and reports an open or a listing the location
refuses with `ERROR_NOT_SUPPORTED` (the item's, or a listed sub-folder's) as `ERROR_OPEN_FAILED`. `DeleteIfIdentity` deletes exactly the pinned tree. It
opens the selected root once, no-follow, for `DELETE` (and listing, for a folder) while sharing only
reads (no other program can write, rename or delete the item itself), compares the root's own facts on
that handle (a difference refuses before anything is listed), then lists the folder's tree again through
it, reporting each listed folder's contents as discovered as the listing reaches them, and compares the
digest: any difference, an identity the location now reports, or a tree now over the bounds is
`ERROR_REVISION_MISMATCH` with nothing deleted, and an open failure (another program holding the item) is
returned as such, a known non-commit. It then deletes only the listed entries. Every listed folder is
opened no-follow for `DELETE`, sharing only reads, and held from the check of its facts (kind, compared
attributes, creation time, listed index) until its own deletion through that very handle, so no other
program can rename it, turn it into a link or replace it meanwhile, and no path below a held folder can
be redirected; folders are held top-down, a level at a time. Each file and link is opened no-follow
through its held folders, sharing only reads, and deleted through that handle once its facts equal the
listed ones. Files and links go first (a chunk's folders taken in turn, so entries deleted at the same
time rarely share a folder), then folders deepest level first, each pass on the shared File Operations
scheduler with the provider's delete concurrency; at most 1,024 folders are held at once
(plus the chain above them): a larger tree goes in chunks of sibling subtrees that fit, a subtree that
alone does not fit going through the same way under its own held folder (one stack level each; a nearly
exhausted stack stops the delete with `ERROR_STACK_OVERFLOW`, as the recursive Copy walkers do). A pass
or a listing level the scheduler ends before every item ran (its shutdown) stops the delete with
`ERROR_OPERATION_ABORTED`. The root goes last, through the compared handle. A listed entry that changed or is gone stops the delete with `ERROR_REVISION_MISMATCH`;
a folder that holds anything the listing never saw cannot be removed (`ERROR_DIR_NOT_EMPTY`) and stops it
too; nothing new starts after a stop (passes already running finish their current entries), and nothing
is listed again, so an item another task or program adds to the tree after the listing, through any
spelling or route, is never deleted. Two residuals remain: corroboration cannot see a change that keeps
a file's size, dates and attributes (an in-place write that keeps them is deleted with its file), and an
SMB client's directory cache can let the execution listing match a pin although the tree changed since,
so a change made shortly before the delete may be noticed only when the delete reaches it (a partial
stop instead of "refused, nothing deleted"). Until the first deletion every stop is a known non-commit;
after it the receipt is {outcome known, mutation committed, original still present}, the reducer's
partial-tree shape (the item fails with its source retained). The call reports every listed entry plus
the item as its items, so the host can say how many of the listed items were deleted. The provider trace
names the step (open, compare, entry, compare-entry, folder, compare-folder, disposition) with the
deleted and listed counts (`FileOps.Local.DeleteByNameStopped`; `FileOps.Local.DeleteByNameTree` with the
deletion time and a `parallel` or `serial` detail), and the listing its time, entries and name bytes
(`FileOps.Local.DeleteByNameManifest`).

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

Copy/Move and admitted permanent-Delete confirmation locations use the captured endpoint path
identity contracts. The source and destination apply their own provider rules; no slash rewriting,
case folding, Windows `parent_path` interpretation, or drive-prefix guess may change displayed
provider spelling. Common-parent grouping requires one qualified provider/instance/profile/root
domain with the same path contract and uses the canonical parent and path-equivalence helpers.
Case-distinct parents in a case-sensitive namespace, unknown contracts, and mixed domains use the
existing bounded localized Multiple locations summary. A displayed single source, sample file,
parent or destination preserves its raw spelling; a directory marker uses only the captured
contract's accepted/preferred separator. An unknown contract adds no separator. Formatting makes
no additional provider call and leaves operation endpoints and consent defaults unchanged.

Clipboard Move also states: “Starting this Move consumes the cut list. If an item cannot be removed,
its source will remain and the cut list will not be restored.”

The Verify control carries a path-qualified availability state in
`HostFileOperationPromptOptions::verificationAvailability`: Supported, Unsupported, Check during
operation, or Not applicable. Admission computes it once from the plans, whether or not the
confirmation is shown (I35 XP-04): bound readback, an exact provider BLAKE3 proof or a destination
writer that proves its own content digest makes it Supported; a deferred capability check makes it
Check during operation; a destination with none of these is Unsupported, and a request for
verification then copies without it (a confirmation-free start clears the plan's Verify intent the
same way the prompt shows it disabled) instead of refusing to publish. A queued task whose destination
changes asks for verification again when only the old destination's missing route cleared it, and the
new destination's availability is decided the same way. Requested verification never
removes a destination's only publication route: on an identity-less destination without a writer
proof the atomic-final writer still publishes and the item reports `VerificationState::Unavailable`. Unsupported and Native-only selections show a disabled Off value with
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
sequence, writes no Move breadcrumb, and performs no mutation. The completed summary carries the
fact `startedMutation` (I35 R9.4): true when any selected root crossed the mutation gate (its builder
reached `MutationPossible`), an owned stage was created, the cut list was used up, an item has an
unknown outcome, or any item published or removed a source; it is decided from the roots' phase and
receipts, never from byte or item counts, so a Delete root that reached its provider is started even
when it removed nothing. A canceled task with `startedMutation == false` gets the canonical result
line `Not started; nothing was changed`; its retained-source counts and actions are unchanged
(`retainedSourcePaths` still lists every kept root) but the popup shows no Source facts for it. A
cancel during preparation logs the Info line `preparation.canceled`, the same line a **Don't start**
answer produces. A clipboard Move clears its captured
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
volume facts for two seconds in one entry. An existing supported Local path beneath
the cached mount may reuse those facts only after a fresh following handle reports
the same volume GUID and a fresh query of the cached mount still reports that GUID.
The lexical prefix limits this optimization; it never establishes volume identity.
The handle requests no data access, shares read/write/delete, follows links and is
closed within the query. It is not the no-follow object binding used for mutation.
The opened-name GUID output has a `MAX_PATH` stack bound; no path handle is retained.
Missing, inaccessible, unsupported, oversized, remote or expired observations use
the original `GetVolumePathNameW` resolver. A known GUID mismatch prevents reuse of
the old same-spelling root cache entry, including if fallback resolution fails.
The usual short-root result uses a `MAX_PATH` stack buffer.
A failed, unterminated or one-character-short result without the trailing backslash
retries the same input once at the existing 32,768-character bound. Failure or an
incomplete result at that bound stays unknown; lexical drive text never substitutes
for the actual mounted root. Volume GUID output uses the API's documented
50-character stack bound. No GUID result retains the original system-resolved root
token fallback. Per-path/operation typed querying and mutation-boundary revalidation
remain required. Native correctness and paired timing qualification of the guarded
observation remain tracked in [I29 Q2-L](../Plans/WIP/FileOperations_PostMergeQualification_2026-10-05.md)
until its source-bound receipts pass.

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
extended attribute, or ordinary attribute after commit, never fails the Copy or removes the published
content. It is traced as `FileOps.Local.DirectFinalMetadataLost` and reaches the host with the item's
receipt (`FileSystemItemMutationResult` V3 metadata suffix, I35 S1b), so the task reports it exactly as
it reports the same loss on the bridge (route parity, R5.5).
The one exception is an EFS source that cannot stay encrypted at the destination: it fails closed
with `ERROR_ENCRYPTION_FAILED` before the first content byte, because plaintext is never published
for an encrypted source. Any pre-commit failure closes the writer and calls `AbortOwnedObject` on
the retained object; it never deletes or reopens the current pathname occupant. That abort uses the
same bounded cleanup snapshot as sibling-stage compensation (`MakeOwnedStageCleanupOptions`: no
primary `operationControl`, a fresh deadline at most two seconds ahead). A successful exact abort reports
`Removed`, an unknown abort outcome reports `Unknown`, and a known surviving partial final leaf
reports `RetainedIncomplete` with `ERROR_IO_INCOMPLETE`. Overwrite, directories, links, Move, and
non-Local providers do not use this direct-final route. A Local Copy that overwrites (after a granted
replacement) streams into an exclusive sibling stage and publishes it with `PublishAs`; it transfers
the same metadata to that stage (sparse and EFS before content, failing closed with
`ERROR_ENCRYPTION_FAILED` for an EFS source it cannot keep encrypted; named streams, MOTW, EA, ordinary
metadata and the security classification after commit) and reports the outcome through the same
receipt suffix. Before I35 it kept only the data stream and ordinary metadata, silently dropping named
streams, MOTW and EA even on NTFS.

Local sibling-stage compensation for files, links, and directories is also exact-handle-only, but
uses a separate bounded cleanup snapshot after primary work stops. Immediately before
`AbortOwnedObject`, both the host bridge and Local provider call the shared
`MakeOwnedStageCleanupOptions` helper, which removes the primary `operationControl` and cookie and
sets a fresh absolute deadline no more than two seconds in the future. Primary Cancel therefore
does not revoke already-proved authority to remove that exact owned stage within that budget: the abort
runs inside a cleanup window during which the synchronous-I/O cancel watch leaves the thread's I/O
alone (I35 R7.4, "Cancellation and lifetime"). After the budget a wedged cleanup is canceled like any
other call; timeout or uncertain abort remains visible cleanup debt. Writer-allocation failure after exclusive stage creation uses the same
cleanup snapshot. No compensation path reopens or deletes the current pathname occupant.

Local conditional publication rechecks operation control after temporary-attribute preparation,
immediately before the final rename. Cancellation during that preparation is a known non-commit;
the exact owned stage remains eligible for bounded compensation, and any displaced destination is
restored through its retained authority. The primary cancel watch remains active throughout.
`FileOps.Local.PublicationCanceledBeforeRename` records the stopped boundary. An attempted rename
that cannot be reconciled still reports Unknown (`FileOps.Local.RenameReconcileUnknown` preserves
the underlying Win32 failure); cancellation never licenses replay or deletion by pathname.

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
its own delete route by name, the authority its Permanent Delete already has. Each directory,
including an empty one, runs exactly one such cleanup after its children. Removing the emptied
selected source directory is the Move's own mutation: a merge that relocates no child (an empty
source, or one whose only content is empty subfolders already present at the destination) and
removes that root reports a known committed receipt with the original removed and completes as
`Moved`, never `Outcome unknown`.

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
| Curl IMAP | None while its route is `uncontained` (browse/open/properties only, I35 XP-08) | None; UID identity never authorizes unbound cleanup |
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
cohort of included tasks with closed totals renders determinate `N% of the work counted so far` while
other tasks still discover, and the Windows taskbar stays indeterminate while any included total is open.

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
path itself does not enumerate a directory. The host bridge is a traversal owner under this rule for
every route it executes. When its selected root is a known regular file, including a cloud, WOF or
dedup placeholder that binds as regular content, it publishes that exact size and closes the root's
scope at authoritative source binding, before writer, stage or payload work. A selected root that
classifies as a name-surrogate link is one item under both Preserve and Skip: neither follows the
target, so the record is item-based with no target enumeration and no target byte total. Closing a
discovery scope is neither transfer completion nor permission to remove a source; publication,
verification, conflict decisions and exact source cleanup all follow it.

**Fresh selected-file counting (I42).** For Copy/Move from fixed, removable or RAM disk paths in
the Local provider, the task worker queries fresh
metadata for the selected regular files before admitting payload work. A Ctrl+A selection therefore
does not wait for preceding copies to finish before its later files contribute to the total. This
pass uses neither pane sizes nor cached application counts, performs no payload read, and keeps no
source handle between probes. Ordinary files use a fresh Win32 attribute/size query; reparse files
must bind no-follow as regular content to participate. Literal links, directories and inaccessible
or unknown-size roots retain their existing traversal/error ownership. UNC/mapped network source
roots retain the producer's I/O policy so a serial metadata prepass cannot delay their first transfer.
Only the drive class is cached during this probe; file counts/sizes are always fresh. Other providers retain their
own listing/metadata and cancellation/session policy. Folder traversal is still single-pass.

The metadata pass runs off the UI thread inside the task's synchronous-I/O cancellation watch and
checks pause, cancellation and Skip counting at each item. Its state is O(selected roots), allocated
only for eligible Copy/Move tasks. Closing its leaf scopes grants no publication, overwrite, read
content or delete authority: the mutation path still binds and revalidates the exact source. Provider
reports/retries for a counted leaf are redundant and must not add it again or reopen discovery. If
a later authoritative report differs in size/kind, keep the counting snapshot and withdraw certainty
instead of presenting a false fixed denominator. Unknown size remains different from known zero.
Publish fresh byte certainty before the last leaf closes discovery and posts the popup update;
the closed snapshot must not depend on a later payload callback to gain its denominator.

**Equivalent selection shapes (I42).** Copy and managed Move apply the same effective file budget
to one selected directory and the equivalent selected file list. A directory is one outer scheduled
root, but its internal workers receive the task's file budget. Local Auto mode honors the host's
nonzero storage-resolved budget without re-clamping it to the inactive Manual value; Manual mode
continues to take the smaller configured/requested limit. Provider/session limits, explicit one,
verification serialization, discovery reservations, bandwidth and safety checks remain effective.

The shared per-item I/O pool retains four to sixteen transfer workers (processor count clamped
to that range), plus one starvation-reserve worker. The reserve must not reduce a four-file
selection to three active transfers on a four-processor host. Existing fairness caps still reserve
capacity for another task; workers remain bounded, reused and event-driven when idle.

The selection-shape qualification counts the native `Transferring` lifecycle phase
after exact source/destination handles, metadata preparation, buffer allocation and
admission, including zero/full byte-report boundaries. Finalizing and Publishing
still own their permits. A Begin-only or queued root is not an active file transfer.
Partial-byte sampling remains diagnostic; actual concurrency must still equal the
four-worker or explicit-one fixture budget, with all content/counting checks intact.

Managed folder Move uses the parallel producer/worker walker when its effective budget exceeds one.
Each directory completion owns exact source cleanup authority, captured before its listing, through
the producer's count and every queued child/directory count. Only the last finisher may restore
destination metadata and conditionally remove that source directory, after its listing and all
descendant publication/proof/cleanup have completed. A failed, skipped, retained or canceled child
keeps its ancestors. New/unselected source children are never removed; a retained source is explicit,
and an unknown conditional cleanup outcome is not retried or relabeled as known retention. Retained
authority/identity bytes remain charged while directory completions are parked; workers join before
their shared state is destroyed. Concurrency-one and verification keep the sequential walker.

Discovery scope belongs to whichever traversal is actually producing the work. A provider call that
a traversal owner makes inside its own walk is a different discovery scope even though it shares the
task's cancellation, deadline and bandwidth controls: exact source cleanup, a child's Native rename,
and owned-stage creation or removal must not publish child-local totals or closure into the selected
root. The host passes such calls an operation-control cookie that suppresses discovery reporting,
while the ordinary progress, conflict and result callbacks keep the root's own cookie. A provisional
mutation whose own report cannot yet know whether it resolved the root is treated the same way in
one direction only: a Native directory rename that may continue as a host rename merge keeps its
cumulative counts and has its one-way closure withheld until the host knows which traversal owns the
rest of that root. Closure stays one-way and idempotent, and the per-item terminal guard remains the
fallback that closes every root exactly once.

Because closure means the totals are final, a correct task never discovers more work after it. The
host counts any such late report as `FileOps.Discovery.GrowthAfterClose`; the value is observation
only and must be zero.

**Closure at listing end (I35 R6.1-R6.7, D-6.1).** For Copy and Move transfer roots, a selected
directory root's scope closes when the walker has consumed every listing it needs (including the
entries' sizes), never after in-flight transfers, verification, post-order directory metadata or source
cleanup. The producer runs at most `DiscoveryQueueTarget` entries of work ahead of its workers
(128-256 in discovery-ahead mode). The Local provider's walker keeps listing past that lead (D-6.2):
an entry it cannot queue yet waits in its directory as a name record (leaf, attributes and kind, never
a work entry) within one 16 MiB name budget per walk, so a folder closes as soon as it is listed, and
its held entries then go to the queue in listing order. A directory that ends its listing while an
ancestor still lists drains what it holds before the ancestor continues, so a nested tree of busy
folders keeps the queue-lead bound. Beyond the name budget the remaining entries are listed as the
copy frees queue slots and closure follows the last listed entry (D-6.1). The host bridge's walkers
count each directory's whole listing at once; their subdirectories are listed as the copy frees queue
slots (parallel producer) or within the lookahead budget below (sequential walker), and closure follows
the last listing. The Local provider's walker always runs its producer ahead, one worker included (that
worker copies in listing order); `CopyDirectoryInternal` remains the walker of a bulk call's roots at
concurrency 1 (a sequential bulk call keeps every callback on its caller's thread and closes its
discovery once, at the call's end) and the fallback when the shared scheduler has no workers. The host bridge's parallel producer moves on as soon as a directory is
listed. Its sequential walker (Verify On, rename merge and concurrency-1 destinations)
counts each listing when it is read and, while discovery runs ahead, lists pending plain directories
ahead of the walk (up to four per child step, at most 128 listings and 16 MiB held, 4,096 pending
names), handing each listing to the walk so every directory is still listed once; a rename merge lists
only what it walks into, and a link is decided by the walk. A listing is a whole provider buffer, so a
directory is complete before a Managed Move removes any of its children. A subtree the walk then
skips stays in the totals, as any item the walk does not complete does. The discovery-ahead worker
reservation (`DiscoveryWorkerLimit`) ends when the producer's listing ends, for each walk of a
multi-root task as well; it never ends by marking the producer done early. A worker's failure that
stops the walk stays that failure, ahead of the cancellations its stop causes in the other workers, and
the producer never asks about it again; in a bulk call a root that a sibling's failure stopped reports a
cancellation, not the sibling's failure. A retried directory never counts its entries twice, still
counts the subtree of a subdirectory that an earlier attempt counted and held by name but never
walked, and does not walk again a subdirectory that an earlier attempt left incomplete (skipped, or
failed under continue-on-error). A held subdirectory that is skipped or fails before its listing ends
stops holding the scope open at once.

**A producer waiting for queue room (Local walker).** The Local walker's producer waits only for its own
walk's queue. Each time it has waited 50 ms without room it asks the caller for a cancel (the operation
control, else the callback; a bulk call shares the provider's 50 ms cancel-check throttle and serializes it
with its other callbacks), so a cancel reaches a walk whose queued items are all held or not yet
dispatched. A producer that runs on a Local provider scheduler worker (each root of a bulk
`CopyItems` walks on one) copies one of its walk's queued items itself when none of the walk's items is
in flight: every worker of the pool can be such a producer, and the copy steps that make room would
otherwise have no worker left (the calls never returned and no cancel reached them). It runs only its
own walk's items, one file or link at a time, never another call's work, and an item it starts runs to
completion, so a large file holds that walk's listing back even if workers free up meanwhile. A producer
on the caller's thread (`CopyItem`, the host's per-item executor) waits for the workers.

**Shared scheduler dispatch (Local provider).** The provider's shared scheduler uses a bounded
four-to-sixteen transfer pool plus one starvation-reserve worker. Both static and
dynamic jobs clamp their concurrency below the actual created pool size whenever it has at least two
workers, including partial thread-creation failure. One throttled job must leave capacity for another
job's first unit; round-robin dispatch cannot preempt a file already in flight. Per-operation transfer
limits remain at most sixteen. Every global worker stream, including the reserve, owns a distinct
progress/throttle/context slot; the reserve must never alias stream zero.

The scheduler dispatches a dynamic job's
step (a Local walk's, a parallel Delete's) only while the job has a ready unit and fewer steps in flight than
its concurrency. A step that finds nothing it may run (a walk at its active limit while its listing runs
ahead, an empty queue) parks its job until the job is woken: by one of its own steps that ran a unit (such a
step may have made others ready, as a Delete promoting a parent folder does), or by a notification of a
change made outside its steps. A Local walk notifies when an item is queued or finished, when its listing
ends, when its producer first sees Skip counting lift the reservation (it reads the mode as it lists and
queues entries), and when it ends. Each wake advances that job's own wake generation under the scheduler's
lock, so a worker never sleeps through it and another job's changes do not wake it. The scheduler never
dispatches a parked job again without a wake (no busy loop of empty dispatches); only a producer waiting for
queue room runs one of its own walk's steps itself when none is in flight (above). A walk parked at its
limit always has an item in flight, so a limit raised while its producer is held in one system call or a
prompt takes effect at the latest when one of its items finishes. A job whose step can find nothing to run
must make every change its step depends on inside its own steps or notify it.

**Observation signals (I35 R6.10).** The Debug `discovery.closed` record names its closer (`producer`
when the walker reported the last root closed, `per-item guard` when the per-item executor closed a root
whose call returned, `execution end` otherwise) and the bytes completed while the scope was open. A
BulkItems call (the pane Delete) covers every selected root with one scope and one discovery slot, so its
own closure closes the task's discovery; the executor still closes it when the call returns (I35 DISC-7). A Copy
or Move whose root is still open while its transfer completes work, with no discovery report for two
seconds, records one Information note per task, `discovery.silence` ("Listing paused: no new items were
found for {0} seconds."): a walker held back by its queue lead, or a slow listing. Neither signal closes
discovery, renders a denominator or changes the footer's open-total state.

The traversal retains queued file/reparse work plus O(depth + directories with in-flight children)
directory/ancestor state: a directory whose listing ended while its children still copy keeps only a
small completion record (its paths, captured metadata and the created directory's identity, with no
destination handle; a managed Move also retains its exact source authority), charged to the
same queue budget as queued work (Local) or to the traversal metadata high-water (bridge); it never
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
directory, retained on its traversal frame or, once its listing has ended, on its completion record,
and restored after all admitted children have finished, by whichever finishes last (the producer or
a worker); a stop releases abandoned queued work so every created directory is still restored, children
first. Folder-on-folder merge never overwrites the metadata of the pre-existing destination
directory. Restoration is deepest-first and reaches only the directory the operation created. A
publication rename may clear staging-only Hidden/Temporary state; ordinary source attributes and
timestamps are therefore applied once more after publication, through an identity-checked reopen of
the path: the Local provider's walker compares the FileIdInfo identity or file index it recorded, the
host bridge binds the path again and compares the identity snapshot it recorded (both below).

**Local walker: no handle on a created directory (I35, NAS nested Copy, 2026-10-12).** The Local
provider's walker holds no handle on a directory it created while that directory's children copy.
Some SMB servers refuse every rename inside a directory another handle holds with DELETE access
(`ERROR_SHARING_VIOLATION`; the Samba NAS probed on 2026-10-12), and every child's publication is such
a rename, so a nested Copy to such a share stopped at its second level with `0x800703E4`; a handle
per directory for the whole walk buys nothing. Right after publication, while the stage's handle is
still open, the walker records the directory's path and identity: its persistent FileIdInfo identity
(volume serial number and 128-bit file id) or, where the volume has no FileIdInfo (the stage's identity
is then its retained handle), the volume serial number and the nonzero 64-bit file index that open
handle reports. The handle then closes. When the directory's children have finished, the walker opens
the path again for attribute access only (`FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES |
SYNCHRONIZE`, no DELETE, sharing read, write and delete, no-follow), requires a directory that is not
a reparse point and has the recorded identity (FileIdInfo compared exactly when it was recorded, never
replaced by the file index; otherwise the file index and volume serial number), and only then writes
the source's attributes and times through that handle, which it closes at once. A directory whose
identity could not be recorded is never touched: it keeps the times the copy gave it, and the item
reports `FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES` as not kept with `ERROR_NOT_SUPPORTED`, which the
host shows as the one-per-task Information note of a destination that keeps no timestamps (I35 R5.4),
as for a destination without object binding. A directory the path no longer names (renamed away,
replaced by another directory or by a link, removed), or that cannot be opened and checked again, is
never touched either: the item reports the times as lost with that refusal (`ERROR_FILE_INVALID` for
another directory at the path, `ERROR_DIRECTORY` for a link), and the host warns `metadata.basic.lost`,
as for the host bridge below (P1: a refused timestamp write is a Warning). Without FileIdInfo the proof is weaker than a retained
handle: nothing pins the directory between publication and restore, so on a file system that reuses
file numbers a directory created at the same path after the original was removed could carry the
recorded number; the restore then writes only the source's attributes and times to it, never data,
links or security. `R4T5_LocalCopyHoldsNoDirectoryHandle`, `R4T6_LocalDirectoryRestoreChecksIdentity`
and `R0fSmb_NestedTreeCopyToShare` guard this.

**Host bridge: no handle on a created directory (I35, NAS Copy and Managed Move, 2026-10-12).** The
host bridge (a cross-provider Copy, Verify On, a Managed Move) follows the same rule, through the ABI:
the object a created directory's publication returns is released at once, and the bridge keeps only its
identity snapshot (the provider's object id, `FileSystemBoundObjectSnapshot::objectId`). The same Samba
server refused each child's stage publication inside a directory the bridge still held, so a bridge Copy
or a Managed Move to it ended with `bridge.publication.unknown` and a visible stage, even for a flat
folder. When the directory's children have finished, the bridge binds its path again, no-follow, through
the destination's `IFileSystemObjectBinding` (`FILESYSTEM_BIND_PUBLICATION`, held only for the write) and
writes the source's basic information through that object only when it is a directory whose snapshot
equals the recorded one. A destination that cannot bind the directory again (`BindObject` Unsupported:
a Local share without FileIdInfo, such as that NAS) keeps the copy's times, the task records the one
Information note that the destination keeps no timestamps, and the bridge stops binding again for the
rest of the item. A path that no longer names the created directory (renamed away, replaced by another
directory or by a link, removed) is never written, and the task warns `metadata.basic.lost`. A route
without object binding records the note as before. `R4T7_BridgeCopyHoldsNoDirectoryHandle`,
`R4T8_BridgeDirectoryRestoreChecksIdentity`, `R0fSmb_BridgeTreeCopyToShare` and
`R0fSmb_ManagedMoveTreeToShare` guard this.

The popup exposes one-way **Skip discovery** (button label `Skip counting`) while discovery-ahead is active. It:

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
publication. The Local walker's record of a directory it created (directory metadata, above) reads the
volume serial number and file index from that retained handle while it is still open, only to find
the directory again for its attribute and time restore; the record never authorizes publication,
replacement, deletion or ownership.

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
- A refusal before the qualified item reaches a mutating provider or host bridge entry reports
  `NotAttempted` and `Retained`, with the exact resolved source/destination paths. Crossing that
  mutation boundary makes missing receipt truth conservative, including on a later retry refused by an
  identity guard. Live-output consent, breaker admission and retained-source revalidation remain
  before that boundary. Valid provider receipts and stronger managed publication/cleanup proof remain
  authoritative; a later refusal must not erase a previous possible mutation.
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

Two renderers, one decision (I35 S5, policy P4). A published decision is rendered on its task card
and, while a pane of the same Folder window shows the folder it is about, by the decision mirror over
that pane (`../UI/UI_FileOperationsPopup.md#decision-mirror`). Both copy the same published prompt under
the task's arbiter lock, ask the question in the same words (`FolderWindow.FileOperations.PromptText`),
and answer only through `Task::SubmitConflictDecision(generation, action, applyToAll, placement)` with
the prompt generation they rendered: the broker accepts the first answer for that generation, refuses
every later or stale one, and the other renderer disappears at the next presentation pass. Escape on
either surface submits only that prompt's own escape action with the Escape placement. Neither
renderer decides on its own: no timer, focus change, close, navigation or teardown submits anything.
Whether a task-wide decision is still "before any change" comes from the engine fact
`Task::HasMutated()`: true once any selected root's result builder entered `MutationPossible` (a
per-item dispatch or the bulk gate; never reset). A task-wide decision is mirrored only before that
point; afterwards only a per-item decision, about its own item's folders, is.
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
typed user decision for that object. Local Native Move follows the same rule: only a destination
that existed at its probe can reach the identity-bound `RenameIfUnchanged` replacement, and its
unbound rename (an absent destination, or a case-only rename including the second hop of its
temporary-sibling route) never asks the file system to replace an existing name. A different object
that appears at the destination name between the probe and the rename therefore survives; the item
fails with `ERROR_ALREADY_EXISTS` (a known no-commit, back to the conflict prompt), except that a
case-only temporary hop whose revert also collides reports unknown publication with the source
left under its temporary sibling.

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
| Recycle failed | Cancel / Delete permanently / Skip; under **More…** in a multi-item task, Delete all remaining permanently | Cancel | Never the checkbox; only the explicit task-scoped choice (see Delete and Recycle) |

Decisions are one-shot unless the user explicitly selects an eligible scope. **Skip All** is Skip
plus explicit Apply-to-all for that row and is normally exposed through **More…**; Disk full exposes
it directly beside Retry and Cancel when its destination scope is qualified. It is never available
for Recycle escalation. A digest match never auto-Skips. Keep Both preview names are advisory until
create-exclusive publication succeeds. The task cache key is the complete typed scope: conflict
class, source kind, destination kind, source path profile, destination path profile, and—when the
destination is a link—the exact known link kind. Unknown link kinds remain one-shot. Skip with the
Apply-to-all checkbox checked records the same scoped Skip grant as Skip All; unchecked Skip applies
only to the current item. Retry always authorizes one attempt and is never cached, even when the
checkbox is checked. Cached choices remain confined to the current task and matching scope;
different parent folders alone do not require another decision. Skip for AccessDenied,
SharingViolation, PathTooLong and NetworkOffline ignores source/destination/link kind while
preserving exact error class, both profiles, destination root and repeat eligibility. Thus a
checked Skip on an inaccessible file also skips an inaccessible folder in that scope, including
when metadata is unavailable. The recorded diagnostic still carries each item's observed kinds.
All other actions and error classes retain typed matching unless explicitly specified otherwise.

Keep Both also distinguishes selected-root and descendant contexts, without changing ordinary
Skip/Replace matching across parents. A checked descendant choice may cover matching descendants
when executable, but a selected-root grant cannot authorize nested collisions. Metadata consent
uses the admitted destination endpoint's namespace `rootId`, not a selected item's destination
pathname. The repeat cache has 64 slots for distinct keys, not 64 question categories or files;
the current 26 classes can produce many more keys. Existing slots remain updateable, and a new
scope without capacity receives an explicit one-item limitation instead of a silently lost grant.

Recursive parallel Copy/Move does not guarantee sibling completion order. The nested Keep Both
regression waits for the nested-collision handshake, releases its callback pause promptly, and
verifies sibling bytes and source disposition after task completion. It must not wait for a sibling
while holding the provider's serialized issue callback. Serial coverage retains its earlier-sibling
partial-tree witness; both modes must retain exactly two explicit prompts and no renamed/retried root.

DiskFull is an endpoint-storage failure whose only reusable action is Skip. Its repeat scope
normalizes file/folder/link shape while retaining task, class, source/destination profile and
qualified destination root. Bulk uses a common qualified admitted destination; missing/ambiguous
qualification withholds Skip all with an explanation. A qualified prompt offers Retry / Skip all /
Cancel directly, with no checkbox/overflow, and caches only the explicit Skip all choice. Other
conflict classes retain their original typed shape boundaries except the four non-mutating Skip
failure families above. Read-only permanent-delete consent
is a separate, explicit noncached action for the exact retained selected tree; known partial work
can resume under that authority, while unknown mutation and generic retry rules remain unchanged.

Conflict classification controls resolution policy, not whether an HRESULT has an explanation.
An otherwise Unknown conflict with Win32 error 2 or 3 still explains the missing required file or
folder. Common failures use plain resources in the application's selected language; uncatalogued
provider errors retain the available trimmed system description through the shared
`FileOperations::FormatConflictProblemText`. This formatter does not grant additional cached
decisions or infer which source/destination/intermediate item caused the failure. Engine outcome
and proof sentinels retain their own semantic text. The original HRESULT stays in Technical details,
logs and exports instead of being appended to the user's explanation. Identity guards explain
same-item aliases, an already-selected folder, a destination inside the source, unsafe linked
ancestry, source/destination revision changes and unavailable reliable identity using their stable
producer categories. A generic unsupported-operation status alone must not imply a collision or
claim either endpoint was preserved. `FileOps_ConflictMessageText` validates the
shared card/mirror/Issues wording, localized fallbacks and finished failures without retained
diagnostics (including live-language notes, synthetic rows and production popup snapshots). It
records bounded formatter work under
`FileOps.SelfTest.ConflictProblemTextBatchUs`; candidate timings alone establish no speedup.

`FileOps_SafetyRemediationContracts` and the focused `FileOps_RepeatDecisionContracts` exercise the
production issue callback across five AccessDenied metadata shapes in different nested parents,
in both file-first and folder-first order. Checked Skip must publish one same-scope prompt and
four cached applications in one aggregate row; each fresh unchecked task must publish five prompts.
Different source/destination profiles, destination root and error class must still prompt.
`FileOps.SelfTest.AccessDeniedSkipScope` records denial prompt count in `value0`, cached application
count in `value1` (not aggregate row count), and starting kind, total prompt count and cached row
count in detail (including separate profile/root and SharingViolation prompts), with zero
`durationUs`. These are deterministic
work counts, not transfer-throughput or UI-latency measurements. `Phase9_ConflictPrompt_SkipApplyToAll`
also checks ordinary checked Skip on actual colliding files and preserves both destination payloads.

The safe `Cancel` action is directly visible and remains the default and Escape action for ordinary
name conflicts. The I26 placement approved on 2026-09-19 publishes eligible actions in the direct
order Keep Both, Overwrite (displayed as Replace / French Remplacer), Skip, Cancel; it removes the
three-primary-button cap for these collisions. Read-only/link variants retain their specific action
identity, label and authority gates in the replacement slot. Width never moves these choices into
More; the consumer wraps them. Skip All remains a distinct scoped More action. Other conflict classes
retain their published order and safe defaults, including Queue after for same-host overlap.
This engine placement is implemented with focused x64 Debug policy/native integration evidence in
[I26's policy slice](../TestRuns/4cb089111a23/FileOps/2026-09-19_105819-i26-policy-candidate/README.md).
The consumer implements measured whole-action wrapping; V65 x64 qualification and its limits are
recorded in [I26](../Plans/Done/UI_FileOperationsReview_2026-09-16.md). Placement does not widen mutation
authority or change conflict eligibility, default consent, repeat scope or action acceptance.
`Enter` invokes only the currently focused action or the published default. `Escape` invokes
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

Accepted conflict choices are observations, distinct from mutation authority and outcome. The worker
records an offered choice only when consuming an active actionable prompt; pending/loading prompts,
invalid ingress and cancellation without such consumption do not invent a user decision. Each record
carries task-local prompt identity, optional selected-root index, exact prompt paths, scope/action and
user or cached origin. A cached application retains its originating user-choice generation.
The original conflict HRESULT is context, not the accepted action's result. Records use the existing
bounded diagnostic/completed-task retention, survive optional Info logging being disabled, and add
schema-v1 facts to JSONL/task-issue exports without changing warning/error counts. Eviction means
unknown; selected-root indexes and path strings do not establish descendant outcome correlation.
A cached apply-to-all choice is never recorded per item (owner decision Q10). Every application of
one originating prompt generation folds into one cached-origin record, which keeps its first
application's prompt generation, selected-root index and paths and adds `appliedCount`,
`lastPromptGeneration` and `lastSourceItemIndex` (zero/`null` on a user choice). An item it decided
carries that record's prompt reference, so its typed result folds into the same record; the record
has a `conflictResult` only when every applied item stored the same typed result. The record keeps
its first application's time. While the task runs, the retained record is refreshed at most every
250 ms by a later application or item result, and when the task raises its next user prompt; a
count can therefore trail by the applications of the last 250 ms until then. When the task
completes its final count is published and only that final record is written to JSONL, once. A
300k-conflict Copy with Apply to all therefore adds O(1) records. The user's own answer to the
originating prompt keeps its separate record and row (the prompt it answered), so each Apply to all
prompt has two rows: the user's choice and the one automatic aggregate. Observations never push
real problems out of the bounded per-task (128) and global (`maxDiagnosticsInMemory`, default 256)
Issues reservoirs (model#2). Oldest first, a reservoir evicts an entry that projects no Issues row
(an Info/Debug note), then a decision or result record whose row retains no original Warning/Error
problem, and only when every retained record completes a retained problem's row, the oldest
problem together with its row's decision and result records. A retained prompt therefore never
loses the decision and result the user saw.
Original problem and accepted/cached records carry the same explicit `conflictPrompt` reference,
conflict class and selected-root facts. Its task/generation key is scoped to the producing
FileOperationState lifetime; combining historical logs requires separate run provenance.
The UI can associate problems and choices despite enqueue ordering; later results still require
independent typed evidence. JSONL and task-issue export include the reference additively.
Inline Rename carries an observation-only accepted task/prompt reference into its terminal item
result. An exact Batch Rename scheduler callback may carry the same reference only for the
conflict choice actually accepted for that callback; preview/Run is not a conflict acceptance.
The callback captures the returned conditional `RenameIfUnchanged` receipt and the unambiguous
accepted reference before the engine folds callback statuses into a coarse task result. Its typed
terminal item preserves the receipt-derived publication, verification, source disposition,
completion, owned-stage disposition and HRESULT independently. A malformed, unknown or
contradictory conditional receipt after the provider call leaves publication and source disposition
`Unknown`, completion `Indeterminate`, and status `ERROR_IO_INCOMPLETE`; even when the provider
changed the namespace, the host must not infer success or attempt the mutation again.

For per-item Recycle failure, the accepted **Delete permanently** choice is exact-item only, and the
task-scoped **Delete all remaining permanently** choice still mutates each item through its own
exact authority. The fresh delete-bound authority remains owned through that choice and its single
`DeleteIfUnchanged` call. The final item retains the choice's explicit task/prompt reference and
the bound-delete receipt rather than associating a result by source path or the task HRESULT.
Delete has `NotAttempted` destination publication and `NotRequested` content verification;
a valid committed removal is `Removed`/`Completed`. A missing, malformed, unknown or
contradictory conditional-delete receipt is instead `Unknown` source disposition,
`Indeterminate` completion and `ERROR_IO_INCOMPLETE`, with no second destructive attempt.
The original failed Recycle HRESULT remains problem context, never evidence of the accepted
delete's outcome.
An exact **Skip** choice performs no bound deletion and retains the accepted reference with
`NotAttempted / NotRequested / Retained / Skipped / S_FALSE` item truth. The task reducer reports
intentional skips as `S_FALSE` (completed with notes), not an unexpected `ERROR_PARTIAL_COPY` failure.
Summary, popup and export preserve this distinction and do not report a removed source.

After first-truth terminal storage succeeds, diagnostics retain these independent facts as
`conflictResult` schema v1. Rejecting a duplicate terminal store or invalid source index emits
no later result. A foreign/zero reference cannot associate an otherwise valid terminal item.
Logging disabled for ordinary Info records does not discard this bounded observation or turn it
into a warning. The reference changes neither provider ABI nor mutation authority.
Copy, Move, bulk Recycle and other executors without an exact carrier remain uncorrelated;
no selected-root index, path or overall-task inference fills the gap. Missing or contradictory
UI observations remain explicitly unknown.
The prompt closes before optional diagnostic I/O. Diagnostics never authorize mutation or retry.
Conflict prompt observations also retain the already provider-qualified source/destination names and
parent folders as display-only strings. They copy the prompt's own facts (a cached aggregate keeps
its first application's), without retaining its inspection capability. Empty qualification requires exact-path
fallback. These additive JSONL/TSV fields do not identify a later mutation result or authorize navigation.

See [the UI contract](../UI/UI_FileOperationsPopup.md#accepted-decision-observations-i26-implementation)
for projection and the remaining I26 qualification.

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
The completed summary still reports `Outcome unknown` with one warning, and the popup presents
the task as Indeterminate, never as a clean Canceled that auto-dismiss may remove.
A user Skip gets the same task-level exception: a `Skipped` item that consumed a descendant Skip
without a receipt keeps its `Unknown` axes on the item (they still block Retry), but the task
completes with skipped items (`S_FALSE`, or `ERROR_PARTIAL_COPY` beside failed or canceled items);
the completed summary does not count it as indeterminate and never reports `Outcome unknown` for it.
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
When the Local Copy of a selected root takes Keep Both after a failed attempt, its receipt
describes the collision-safe sibling it actually wrote: the earlier attempt's stage at the original
name no longer decides publication or owned-stage disposition, while its unresolved cleanup debt
(retained or unknown stage) still does.

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

Recycle never deletes permanently. Local Recycle runs the Shell's `IFileOperation` without its
confirmation UI, which would otherwise answer the Shell's own "delete permanently?" question when it
cannot recycle an item. The Shell offers each item to the progress sink's `PreDeleteItem` with
`TSF_DELETE_RECYCLE_IF_POSSIBLE` only when it will move it to the Recycle Bin, so the sink refuses
every item offered without that flag before any mutation. This flag behavior is measured for a
UNC/SMB share (the loopback administrative-share alias); for a volume without a Recycle Bin or an
item over the bin quota it is expected, and the backstop below covers a Shell that deletes such an
item anyway. The refused item completes with `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)` and a known
non-commit with the original still present, and the exact-item `RecycleEscalation` below then asks
before any permanent delete. Both executors honour continue-on-error the same way (owner decision,
2026-10-07): with it, a failed Recycle is never escalated; the item keeps its known non-commit, is
reported `Failed` with an `item.continueOnError` diagnostic, and the task continues with its other
items and ends `ERROR_PARTIAL_COPY`. The per-item executor records this before the escalation; the
BulkItems executor (the pane Delete command), which escalates after its one provider call, applies
the same rule to each failed item. Only without continue-on-error does a failed Recycle reach the
escalation prompt below. The BulkItems provider call then stops at its first failed item. Once the
escalation has resolved every failure that call reached (Delete permanently, the task-scoped choice
below, or Skip), the task goes on as the per-item executor would: exactly one further provider call
recycles every item the stopped call never reached, each result mapped back to its own selected item,
and their failures then escalate in turn in selection order. That call alone passes continue-on-error
to the provider (the task's own continue-on-error policy is unchanged), so it attempts every item it
is given and reports each failure once; the number of provider calls never grows with the selection
(at most two), even when every item fails on a location without a Recycle Bin. Because that call
recycles every remaining item before the next prompt, a Cancel at a later prompt, or a later failure
that cannot be escalated (the item changed, or a contradictory receipt), keeps the items it already
recycled. The further call is its own discovery scope: its totals add to the task's totals after the
first call's closure and are never counted as growth after close. The escalation's exact delete of a
file adds no discovery, because its Recycle attempt already reported it; a folder's recursive
escalation adds the descendants it walks, as a scope of its own, so the task's totals stay exact.
Once a Local Recycle call stops (another item's failure without continue-on-error, a fatal provider
error, or cancellation), every entry it has not yet attempted is left unreported, as never reached.
A Recycle that the stop interrupts before the Shell holds its item (a known non-commit, before its
discovery is reported) is also left unreported. Only the caller's cancellation stops a single-item
Shell operation already under way: one that another item's failure or a fatal error stops runs to
its terminal observation and is reported with it (recycled, refused, failed or destroyed), so an
item whose discovery the call already reported is never left for a resume that would count it
again, and a commit the Shell already observed is never reported as an unknown outcome. Such an
operation keeps polling the caller's cancellation (operation control, or the legacy callback) after
the stop, so a user Cancel issued while it runs (a large folder's walk, for example) still stops the
Shell and is recorded as the caller's cancellation. A stop-time cancellation of a Recycle whose
commit the Shell had already observed (canceled just before another item's stop, for example) is
reported as recycled, never as an unknown outcome; an item the Shell destroyed instead stays a
failure. Any other stopped Recycle the Shell may have acted on is
reported with
`HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE)` and an unknown-outcome receipt. Only a cancellation by the
caller is ever reported as canceled, so another item's stop is never mistaken for a user Cancel.
Every batch entry the Shell observed, and every entry it offered without a terminal observation, is
reported even after the call stops; only a report the host refuses ends that reporting. Only a
first call that stopped at the failure of an item it reported, which the escalation resolved, is
resumed; a call that ended with cancellation, a fatal provider error or a failed batch keeps its
own result. An item without a status is one the provider never touched, so only such items are
resumed, and an item whose Recycle the provider may have acted on is never resumed. A Cancel at a
prompt, or an escalation that fails, ends the task without asking about any other failed item. In
the BulkItems executor without continue-on-error, an escalation that fails (the item changed after
its Recycle attempt, its identity is unavailable, its exact delete fails, its outcome is unknown, or
its receipt is contradictory) stops the escalation pass at that item, as the per-item executor stops
at its first failure, and the task ends with that failure: every failed item not yet escalated
keeps its known non-commit and is reported `Failed` and retained, and is never prompted or covered
by a task-scoped grant accepted earlier in the task. With continue-on-error nothing escalates, as
above.
Local keeps advertising Recycle for SMB
and removable routes; this refusal and the escalation, not admission, handle a location without a Recycle Bin. The Shell
decides recyclability for a whole batched `IFileOperation`, and a refusal ends that operation before
it touches another item, so the refused entry and every entry the Shell never offered are replayed one
by one through the single-item route, which recycles or refuses each on its own. An entry the Shell
offered without a terminal observation is not replayed and keeps the unknown outcome: the provider
reports it with `HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE)` and an unknown-outcome receipt before any
observed failure of the same batch can stop the call, so the host records it as attempted, never
escalates it, never covers it with the task-scoped grant and never resumes it. As a backstop,
a Shell success without a Recycle Bin item is reported as a failure
(`COPYENGINE_E_RECYCLE_FORCE_NUKE`) with a committed receipt and the original removed. The host
treats that status as "destroyed, not recycled": the item stays `Failed` with the source removed,
the `delete.recycle.destroyed` diagnostic names it, it is never escalated to a permanent-delete
prompt, and it never counts as a successful Recycle. Any other failed Recycle status whose receipt is
a committed removal is still completed as a removal, because the provider attests that the item left
through the Recycle route.

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
outcome keeps the source and never retries permanent deletion. Default is Cancel; Skip keeps the item
(this item only); Cancel ends the task. There is no Apply-to-all checkbox and no Skip All. The old
behavior of clearing the Recycle flag and retrying provider pathname deletion is prohibited.

On a volume without a Recycle Bin every item of a large Delete fails the same way, so a task with
more than one selected item also offers, under **More…** of a real escalation prompt, the explicit
task-scoped choice **Delete all remaining permanently** (owner decision Q2). Accepting it deletes
the prompted item exactly as **Delete permanently** does and records one task-scoped grant in that
task's decision cache, keyed by the Recycle-failed class, the provider profile, and the delete
endpoint's root (the route/volume). A later item of the same task consumes the grant instead of
showing another prompt only when it meets conditions 1-3 above on that same profile and root; it is
then deleted through its own fresh binding and its own `DeleteIfUnchanged`, exactly like an
individually accepted item, and a folder is deleted recursively as below (the grant covers files
and folders alike). Each such item is counted in the task's one cached-scope accepted record for
the originating prompt (Q10, above), records a `delete.recycle.escalated` diagnostic, and folds its
typed result into that record. An
item whose Recycle outcome is unknown, destroyed or contradictory, whose identity changed or could
not be bound, or that is on another profile or root never consumes the grant: the unknown,
destroyed, contradictory and identity failures fail without a prompt as above, and another root
asks again. The grant is never offered without a real escalation prompt, never shown as the
Apply-to-all checkbox, never stored in Preferences, and ends with the task: a new task asks again.
A one-item task does not offer it. Continue-on-error, as above, keeps every failed item without
offering it.

A selected folder escalates like a file. `DeleteIfUnchanged` receives the plan's flags, as the
plan's permanent Delete would: `FILESYSTEM_FLAG_RECURSIVE` when the Delete is recursive (the pane
Delete always is) and `FILESYSTEM_FLAG_CONTINUE_ON_ERROR` when the task has it, so an accepted
**Delete permanently** removes the folder's contents and then the exact bound folder. Because the
consent covers the whole tree, a recursive folder escalation reports the tree, not the folder
entry: before the prompt the host sizes the bound folder's path recursively (links and mount points
count as leaves, as the delete walk treats them) and reports the folder plus its files and
subfolders as the item count and the files' total as the bytes. The walk stops when the task is
canceled. These counts describe the tree when the prompt opens and grant nothing; a partial or failed
walk reports no counts rather than understating them. A folder the task already holds a matching
task-scoped grant for shows no prompt, so the host does not walk it for counts. A file reports one
item and its committed size. A child the recursive delete cannot remove leaves the folder present: the escalation is a known
non-commit of the selected folder and the item fails without another prompt.

When the exact permanent-delete escalation succeeds, that result replaces the earlier Recycle failure
for item completion and diagnostics. The final item is successful; the stale Recycle HRESULT is retained
only as historical prompt context and cannot make the completed escalation appear Failed. A skipped or
canceled escalation similarly reports its explicit terminal action without weakening the original
known-noncommit/source-retained mutation truth.

### Identity-less locations: delete by name (I35 R7.3, D-7.3)

When a Local permanent Delete root binds `Unsupported` (the location opened it but exposes no file
identity, as the NAS probed on 2026-10-11 does), Preparing pins its corroboration facts through
`ResolveDeleteIdentity` instead of failing, in a conservative interlock scope marked by-name; it never
retains exact delete authority or reuses a `Bound` state for it. A folder's facts include its pinned
tree; a folder over the listing's bounds is refused there, before any card, with
`preparation.deleteRoot.byName`, `ERROR_NOT_ENOUGH_QUOTA` and the plain reason "Too many items to check
by name" (`IDS_FILEOPS_NOTE_BY_NAME_TOO_MANY_ITEMS`), nothing deleted. Only `Unsupported` qualifies: an
`Indeterminate` bind (sharing violation, access denied, delete pending) keeps failing at Preparing and
logs the open and identity results separately (`Debug::Warning`, `fileops.delete.root.pinFailed`, and
`preparation.deleteRoot` with `openHr=… identityHr=…`); it is never offered by name. Such a root is
not retried at Preparing: the task fails with both results and the user starts the delete again once
the other program let go (an accepted UX gap, S4 departure 2; a root another program takes after the
card's answer gets the item's Retry question). Only a task that asks its own consent for such a root
takes this route: the permanent-delete card (`_permanentDeleteConfirmationPending`) or a Recycle's
escalation question. A Pack or Unpack cleanup, whose consent is captured before the task, keeps the
refusal (as `DeleteHandoffForPlan` keeps its leftovers' refusal), and execution refuses a by-name root
whose consent is not the task's own `PermanentDelete` card with `ERROR_NOT_SUPPORTED`. A by-name root
that another live task of this host holds in a prepared or active scope (the same path, an ancestor or
a descendant, at the same provider root) is refused with `ERROR_BUSY` and
`delete.byName.heldByAnotherTask`, at Preparing, again after the answer, and again right before each
`DeleteIfIdentity` (a task admitted after the answer, for example one the user let run at the same
time, is caught there).

The same-host overlap advisory compares a by-name root by path, exactly as it compares an exact delete
root, whether the delete's card waits, it is prepared, waits in the queue or runs: a Permanent Delete
publishes its prepared registration before its card is shown, so a task admitted while the card waits
sees the root. A task admitted around it whose scopes name the root, an ancestor or a descendant at the
same provider root is asked `Another task is working in the same place` with **Queue after these
tasks** as the default and never **Run at the same time** (`CanRunConcurrently()` is false whenever an
overlapping pair involves a delete by name): the question names that delete first and says what it does
to the task (`IDS_FMT_FILEOPS_OVERLAP_DELETED_BY_NAME_WRITTEN` for a task that writes or changes items
there, `IDS_FMT_FILEOPS_OVERLAP_DELETED_BY_NAME_READ` for one that reads there), and queued after it the
task enters only once the delete ended. The question's texts hold for every case: the delete may remove
a file or a folder by name, a Recycle's by-name facts may never be used, and the asked task may write
into, change, read or remove an ancestor of that item ("Task #3 may delete by name an item that this
task writes or changes, or a folder that holds it", or "reads"). A newer task still waiting on that
question does not hold the root for the delete's re-checks: it cannot enter before it answers, and its
answer cannot be to run beside the delete. Queue-after edges never close a cycle: a peer that already
waits on the asking task, directly or through other tasks' Queue-after decisions, is not named, and a
decision published after such a chain formed drops the edge that would close it
(`FileOps.Overlap.CycleEdgeDropped`). A spelling of the same folder through another root of the same
identity domain (another share name or a mapped drive), or a writer outside admission (Pack writing its
archive, Unpack extracting), cannot be compared by path: the conservative interlock serializes running
tasks, and the pinned tree makes the delete refuse a tree that changed before its last listing, or stop
when it reaches a change after it; an item added to the tree is never deleted, with the residuals named
above (a change that keeps a file's size, dates and attributes; a change an SMB directory cache hides
until the delete reaches it).

The card asks one explicit consent in its own class (`PermanentDeleteByNameConfirmation`, title
"Permanent delete by name"): **Delete by name** and **Cancel**, Cancel first, the default and the Escape
action, no Apply to all and not a preference. Its question ends with a paragraph saying, in the C10
wording family (never the engine word *identity*, P5), that the location cannot prove which object a
name refers to, that only what it lists now is deleted by name (each selected item and everything inside
a selected folder), that each item is checked again right before it is deleted (attributes, dates and a
file's size), the delete stopping at an item that changed and never deleting anything added after the
listing, and the residuals: a change that keeps a file's size and dates cannot be noticed, and some
network shares report recent changes late, so the delete may stop after part of a folder is gone
(`IDS_FILEOPS_CONSENT_BY_NAME_CHECKED`), plus the leftover warning
(`IDS_FILEOPS_CONSENT_POSSIBLE_LEFTOVER`) when a selected root has a Possible name. After the answer
each root is re-resolved, a folder's pinned tree included, and refused on any difference
(`delete.confirmation.identityChanged`, `ERROR_REVISION_MISMATCH`, nothing deleted); execution then
calls `DeleteIfIdentity`, which compares the facts and the tree again on the handle that deletes and
deletes only the listed entries. A refused or stopped root records `delete.permanent.byName` with the
step it stopped at and the open and corroboration results (`openHr=… identityHr=…`): `open` (nothing
deleted; an item another program holds without delete sharing gets the item's own Retry question, never
a weaker route), `compare` (a changed item or tree, a definitive non-commit, nothing deleted),
`contents` (a folder checked but stopped before anything was deleted: the folder and its items were
kept), `partial` (part of the pinned tree was deleted before the stop: the receipt is committed with the
folder still present, the item fails with its source retained, and an Error, the same severity as on the
Recycle escalation, says "The delete by name stopped after deleting N of the M items it listed; the
folder and the items still in it were kept" (`IDS_FMT_FILEOPS_NOTE_BY_NAME_PARTIAL`, N and M from the
call's progress counts), never "kept" for everything) or `disposition` (a file checked but not deleted).
A Delete item's committed receipt stays committed: a later call for the same item that reports nothing
deleted never turns it back into "nothing changed". The host tells a stop after the check from a refused open by the discovery the provider
reports only after the facts matched, and a partial delete from a stop before any deletion by the
provider's receipt. A live-output conflict
never offers "Delete it anyway" for a delete by name: it waits or skips. A removal records a Debug
`delete.permanent.byName` per item and one Information note per task ("Deleted by name: this location
cannot prove which object a name refers to.").

A Recycle the location refuses on such a root escalates through the same route behind its own consent
class (`RecycleFailedByName`: **Delete by name**, Skip, Cancel the default; for a multi-item task the
explicit **Delete all remaining by name** under More, whose task-scoped grant is keyed by that class and
never covers an exact-identity escalation, nor the reverse). Preparing pins a Recycle root's facts best
effort for this (a folder's pinned tree included); a root whose pin failed keeps the reason with its
scope (`byNamePinStatus`), so its refused Recycle offers no delete by name and says why: over the
listing's bounds, the plain reason "Too many items to check by name"
(`IDS_FILEOPS_NOTE_BY_NAME_TOO_MANY_ITEMS`, `ERROR_NOT_ENOUGH_QUOTA`), otherwise the pin's own
result. The facts and a
folder's tree are re-checked before the question and after the answer, the provider deletes only the
pinned tree as above, another live task holding the path refuses it (checked right before its delete), a
refusal names its step as above (`delete.recycle.byName`, never "kept" once the check passed; a partial
stop carries the same note), and a removal records the Information note `delete.recycle.byName`. Move, Rename and expected-destination replacement proofs never
use this route, and only the Local provider offers it for a binding provider (D-7.3); providers without
binding keep C10.

Pack and Unpack delete-after-success submit ordinary exact `DeletePlan` tasks after the archive
contract succeeds; on an identity-less location they keep the refusal and never delete by name. No
archive component performs a private pathname delete.

## Verification and metadata

Verification defaults off and applies after publication to copied regular content. Selecting Verify
On is an explicit engine cutover for non-Native copied payloads, including same-endpoint Local Copy:
the host exact-authority bridge owns transfer, publication, proof/readback, link handling, metadata,
and conflict callbacks. The host does not run provider `CopyItem` and then reopen a final pathname to
hash it; therefore Verify On changes the engine path but not the observable Copy/Move semantics.
Verify Off keeps the provider `CopyItem` route for same-endpoint Local Copy, so SMB server-side copy
and block-clone offloads stay available.

**Routes inside one provider instance (I35 XP-09).** When the source and destination panes share the
plugin and instance context, execution runs the provider's own `CopyItem` even when the two endpoints
are different roots of that instance (two Local volumes, a volume and its administrative-share alias,
two S3 buckets); admission's endpoint qualification does not choose that route. A provider whose typed
facts admit its own plugin as an import/export peer therefore MUST accept every source/destination pair
admission can produce inside one instance. Local and S3 do (Local copies across volumes with its
offloads; S3 copies across buckets server-side, with a relay fallback). Google Drive does not yet: its
`CopyItem` refuses a pair across connections or shared drives with `ERROR_NOT_SAME_DEVICE` although its
facts admit it, so that pair fails with that error unless Verify On routes it through the bridge (an open
owner decision: relay inside the provider, or path-aware peer facts). `FileOps_CrossRootLocalCopy` pins
the Local contract.

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
| MOTW, ADS, extended attributes | Preserve or report exact loss (Warning per object) | Keep source unless task-scoped “Move anyway; metadata lost” |
| ACL/owner | Destination inheritance accepted: one Information note per task; a protected source DACL replaced by inheritance, or an unreadable descriptor, is a Warning | Neither blocks cleanup |
| EFS | Never silently decrypt | Gate before plaintext; default Keep source/Stop |
| Sparse allocation | Preserve or state logical inflation | Gate before inflated write |
| Compression | Never transferred; the destination inherits its folder's state (File Explorer rule) | Never a loss, never a gate |
| Timestamps/ordinary attributes | Best effort; a destination that cannot store them (no metadata interface or authority, `ERROR_NOT_SUPPORTED`/`E_NOTIMPL`) is one Information note per task; a real write failure is a Warning | Does not block cleanup after required gates |
| Directory timestamps | Restore deepest-first with bound authority | Same rule as ordinary timestamps |

The Local exact metadata implementation uses the retained source handle and exact exclusive stage.
`BackupRead`/`BackupWrite` copies MOTW, non-default ADS, and EA records without copying the default
data stream; sparse and EFS controls are attempted before content, compression is left to
destination-folder inheritance, and destination ACL and owner inheritance is compared rather than
silently cloned: equal descriptors are preserved; differing descriptors are `informationalFeatures`
(D-5.1: inherited ACEs, explicit ACEs the inheriting destination lacks, owner and primary group are the
destination's expected behavior) unless the source DACL is protected (`SE_DACL_PROTECTED`) and the
destination's is not, which is `changedFeatures`; a descriptor that cannot be read is `lostFeatures`
with `firstFailure`. Stage publication may clear
staging-only attributes, so basic attributes/timestamps are finally re-applied through the returned
published authority. Metadata results never authorize an identical-file Skip or source deletion.

Metadata diagnostics are route-neutral (I35 R5.4/R5.5): the bridge and the provider-native Copy
receipts classify through one host function (`Task::ReportMetadataOutcome`).

| Outcome | Severity | Category | Recorded |
|---|---|---|---|
| MOTW, named streams, EA, sparse, EFS (after consent), placeholder or ordinary attributes not kept | Warning | `metadata.<feature>.lost` | per object, localized sentence per feature |
| Protected source DACL replaced by the destination's inheritance | Warning | `metadata.security.changed` | per object |
| Security descriptor could not be read or compared | Warning | `metadata.security.lost` | per object |
| Permissions differ only by destination inheritance (ACEs, owner, group) | Information note | `metadata.security.inherited` | once per task, with the object count |
| Destination keeps no timestamps or permissions (no metadata interface or authority, `ERROR_NOT_SUPPORTED`, `E_NOTIMPL`), files and created directories | Information note | `metadata.basic.unsupported` | once per task, naming the destination folder |
| Compression | none | none | never a note or a loss |

`bridge.metadata.inspect` and `bridge.metadata.read` (an unexpected inspection failure) and
`item.cleanup.retained` (an artifact remains on disk) stay Warnings. Notes follow the Information-note
rules of `../UI/UI_FileOperationsPopup.md` "User-facing policy": always recorded, never counted.

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
| EFS plaintext destination | Copy/Move without encryption; Copy and keep source on Move | Same task/risk/destination root when explicitly selected |
| Sparse allocation inflation | Continue | Same task/risk/destination root when explicitly selected |
| Placeholder hydration/recall | Download and continue | Same task/risk/destination root when explicitly selected |
| Security-significant metadata loss | Move anyway / Copy and keep source | Same task/risk/destination root when explicitly selected |
| Recycle escalation | Delete permanently / Skip; More: Delete all remaining permanently (multi-item task) | Exact item, or the explicit task-scoped choice for later known non-commits on the same profile/root; never the Apply-to-all checkbox |

Each accepted choice creates a bounded receipt containing risk, task ID, optional item index,
destination root, optional source/destination identity, and decision. The task retains at most 64
receipts; exhaustion fails closed. Receipts are task-local, never session-sticky, never Preferences,
and never transferable to another risk/root/task. Recycle escalation never offers the Apply-to-all
checkbox; its only broader scope is the explicit **Delete all remaining permanently** action. An
explicit eligible Apply-to-all acceptance, or that action, creates one receipt for that exact
task/risk/destination-root grant; later matching items consume the cached scope and do not pretend
to have separate per-object receipts. Recycle uses this scoped cache only for that action's grant.

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
task also takes a fresh queue place behind every task admitted so far, so the visible order matches
("behind" compares queue order keys, so a reordered or late-queued peer is placed where it waits). The host
raises one warning on the task card naming the
concrete problem and the first overlapping task (this task may remove what task N reads; both may
create or replace the same names; both may remove or rename the same members; this task may remove
what task N publishes; this task may replace what task N reads; task N deletes this folder by name and
this task writes or reads there, named first, I35 R7.3). **Queue after these tasks** (the
default) keeps the admission wait below as the transient edge, released without a second warning once
the named tasks terminate; **Run at the same time** admits the task concurrently only when the full
disclosed relation set fits the fixed 64-task bound and no overlapping pair involves a delete by name; **Don't start** (Escape) ends the task as a
Preparing cancel with no mutation, clipboard consumption, or breadcrumb. The question's decision stays
unrecorded; the metric `FileOps.Overlap.DontStart` counts only the user's own **Don't start** answer,
and `FileOps.Overlap.CanceledWhilePending` counts a Cancel all or close that ended the task while the
question waited (I35 R9.6). Run stores only those exact
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

### Nested archive Copy to local and UNC folders

Preparing derives the selected filename through the source endpoint's admitted path identity,
then applies the destination provider's existing child-name validation before any mutation.
For FileSystem7z, `/` and `\` are accepted path separators and `/` is preferred. Selecting
`/folder\file.ext` therefore copies `file.ext` into the chosen local or UNC destination;
the source parent must not become part of the destination filename or an extra output directory.
Do not bypass destination validation to accept a malformed relative path as a child name.
Reserved Windows names, trailing dots and embedded separators still fail before mutation.

`FileOps_ArchiveNestedCopyLocal` and `FileOps_ArchiveNestedCopyUnc` exercise the real reader,
Preparing and host bridge with ZIP, 7z and four-volume RAR fixtures. They cover canonical paths,
Windows FolderView joins and backslash paths, require exact bytes and exactly one output
leaf, and retain negative destination-name checks. UNC execution uses only the existing
guarded loopback alias of the marked sandbox; an unreachable administrative share is a
declared skip, never evidence of an SMB pass. Each successful copy emits
`FileOps.SelfTest.ArchiveNestedCopy` with target, archive format, spelling, elapsed
microseconds, bytes and published-file count. These small correctness fixtures do not
qualify NAS throughput or large multipart RAR integrity.

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
  The exception survives live Queue reapplication and completion-driven pause recalculation:
  inline renames are excluded from the oldest-active selection used to compute queue-induced pauses.
  An explicit user pause and cancellation remain independent and retain their authority.
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
`NotPublished / Retained / Skipped / S_FALSE`, not a successful rename; the engine's own Batch Rename
report keeps its skipped-row wording. An Inline Rename Skip records the same typed `S_FALSE` item.

Rows admitted to one runnable dependency layer are independent. Unless cancellation is requested, the
engine attempts every row in that current layer even when an earlier sibling fails or is skipped, and
every result remains truthful. A failed, skipped, canceled, or indeterminate row never vacates its
source namespace and therefore never unlocks a later dependency layer that targets that
namespace. The scheduler stops before those dependent mutations; it does not turn one row's failure
into permission to overwrite or rename through a retained source.

A row whose mutation callback was never invoked, because the scheduler stopped before its layer or
before its turn in a canceled layer, has no outcome of its own. Its typed result is `NotAttempted /
Retained` with `Failed` and `ERROR_OPERATION_ABORTED`, or `Canceled` and `ERROR_CANCELLED` when a
cancel stopped the run. It never inherits the stopping row's HRESULT or its axes: an earlier row's
unknown receipt does not make an untouched row `Indeterminate`, and an earlier row's access-denied
failure is not reported on a row that was never tried. The task result still names the stopping
row's outcome. A row that was invoked but rejected before its provider call keeps its own failure
status as `NotPublished / Retained / Failed`; every outcome after a provider call is the
callback's own typed result.

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
task has been reaped there. The exit stays pending until that resumed `WM_CLOSE` has been handled:
a message posted ahead of it (a Change Case result, a posted pane Delete command) is refused like
any other new work instead of starting work the close would then ask to cancel again. The close
handler ends the pending exit (`ConfirmCancelAllFileOperations` finds nothing active), and a settle
message posted right behind the `WM_CLOSE` (`WndMsg::kFileOperationsDeferredCloseSettled`) ends it
when the target closed another way or stayed open, so a later close or operation is never blocked.
An Exit with nothing active closes at once. The consent is asked only
while a task can still be canceled (`HasCancelableOperations`); when every remaining task is already
Stopping the close is deferred without a prompt. While an accepted exit is pending (the deferred close
target is armed, `FolderWindow::IsFileOperationsExitPending`), every new file-operation admission is
refused before any prompt, clipboard consumption, breadcrumb or other mutation: pane Copy/Move/Delete
commands, drag and drop, clipboard paste, inline rename, Batch Rename (opening and executing), Change
Case, archive pack/unpack cleanup Deletes, Find Files, Compare Directories sync and pane commands, and
any other caller of `AdmitOperation`, `AdmitInlineRename` or `AdmitScheduledRename`. FolderView asks the
host through its refusal callback before it shows its own rename prompt (context menu Rename, F2) or the
Move folder picker, so a refused rename or Move never asks for a name or a folder first. The pending
exit is application-wide: the Compare Directories window's own `FolderWindow` and engine are linked to
the application `FolderWindow` (`FolderWindow::SetApplicationExitFolderWindow`), report its pending
exit, and show the refusal in the Compare pane. The link also makes the application exit cover every
linked engine: the consent is asked when any engine (main or Compare) has a cancelable task, "Cancel
all and exit" cancels the tasks of every engine, the deferred close resumes only when every engine is
quiet (a linked engine's completion path resumes it too, as does unlinking), and a further close
brings forward the popup of an engine that still has work. Closing a Compare window runs the same
consent on its own `FolderWindow` for its own engine only ("Close Window"): declining keeps the work
running, accepting cancels it and defers that window's close until its engine is quiet, and while
that close is pending the window's admissions are refused with `IDS_FMT_PANE_FILEOP_REFUSED_CLOSE_PENDING`
(same HRESULT) while the main window stays usable. No engine with running workers is destroyed
without that cancel path, except when teardown is forced. A refused Batch Rename execution reports no row as
failed: the window's report is canceled with the refusal text, and the pane alert explains it. `StartOperation` checks again after
its last prompt, because a prompt pumps messages and the exit may be accepted while it is open. A
refused admission returns `FolderWindow::kFileOperationRefusedForPendingExit`
(`HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IS_SCHEDULED)`), publishes no task and shows a localized pane
alert (`IDS_FMT_PANE_FILEOP_REFUSED_EXIT_PENDING`, titled `IDS_CAPTION_FILEOPS_EXIT`) naming the
operation and saying that RedSalamander is closing; a FolderView request receives it as `S_FALSE` so
that alert is not replaced by a generic failure. A further close attempt while the exit is pending
does not prompt again: it brings the File Operations popup forward and keeps the pending exit. Tasks
that were waiting in the queue when Cancel all ran end canceled without ever entering their operation
(`EnterOperation` checks the durable cancel first). The `WM_DESTROY` join
remains the final ownership boundary and finds nothing to wait for on that path; it still runs
synchronously when teardown is forced (end of session, a direct destroy). No worker is detached,
killed, or abandoned to make a close faster.

The Local `local-win32-smb` route (UNC paths and mapped remote drives) is `bounded` by the
synchronous-I/O cancel watch (R0f-SMB). Every File Operations worker thread and every Local provider
scheduler worker registers itself with a per-module watch for the duration of its provider work. The
watch polls the task's cancel/stop state (host workers) or the unit's operation control (provider
workers) every 50 ms and, once that state has requested cancellation for 500 ms without the call
returning, cancels the thread's pending synchronous Win32 call with `CancelSynchronousIo` on every
poll until the call returns. The grace lets cooperative checkpoints finish untouched on a live share.
Owned-stage cleanup has its own budget (I35 R7.4): every `AbortOwnedObject` (the host bridge's
`AbortOwnedStage` and the Local provider's own compensation), the raw `DeleteExactHandle` removal of a
stage or created object the provider just made, and the `PublishAs` rollback and backup cleanup that
protect the user's original run inside a `Common::SynchronousIoCancelWatch::CleanupWindow`. The window
publishes a per-thread cleanup deadline (the cleanup control's deadline, at most
`FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS`, two seconds; nested windows keep the latest and restore the
previous one) that every module's watch registration reads, and the watch issues no
`CancelSynchronousIo` before it. A cleanup still wedged after that budget is canceled like any other
call and stays visible cleanup debt. The watch reads the window when it polls, so a poll that read it
just before it opened can cancel a cleanup call that then starts inside it: a cleanup call that fails
with `ERROR_OPERATION_ABORTED` while its own window has not expired is retried once
(`CleanupWindow::CanceledInsideWindow`; the exact-object removals and restorations are idempotent on
their own handle). The deadline lives in one per-thread slot exported by Common.dll
(`Common::SynchronousIoCancelWatch::CurrentThreadCleanupDeadline`), so the host and the Local plugin,
each with its own watch registry, read the same window (an accepted S4 departure from a plan that
named no shared slot). (Before I35 the 500 ms grace also applied to cleanup, so a slow SMB removal of a
canceled Copy's stage could be cut short and leave a `.rs_tmp_` leftover.) A call
still wedged after the grace (a dead share holding an open, enumerate,
write, rename, delete, flush, or close) fails with `ERROR_OPERATION_ABORTED`, which a canceled task
reports as the cancellation itself and never as a provider fault: the durable cancel request, not the
status the execution path happened to return around it, decides that at every result seal, so such
an item ends `Canceled` and the task ends `ERROR_CANCELLED`. The overlapped source reader keeps
its `CancelIoEx` quiet point. Nothing is abandoned, terminated, or unloaded: the wedged call returns
to its caller and the item carries the same receipt truth as any other canceled item
(published/retained/removed per root, owned stage removed or retained). Fixed-volume Local routes
carry the same watch; their calls complete long before the grace, so their behavior is unchanged.

Cancellation qualification distinguishes a known pre-rename refusal from an interrupted in-flight
rename. The former must remove its exact unpublished stage. The latter, when the provider cannot
reconcile the exact object's location, must preserve Unknown publication, retain the affected source,
and expose uncertain stage debt on the completed card; the absence of a stage name alone cannot prove
publication either. A retained name remains a Possible artifact without recovery authority. Tests
must reject unexplained leftovers and verify surviving contents rather than demanding unsafe replay
or pathname cleanup to obtain an empty destination.

For a selected root with parallel bridge transactions, owned-stage debt is cumulative: Unknown
dominates retained-incomplete and retained debt, and successful publication or removal of a sibling
cannot clear it. The joined root result and completed-card counts must preserve that debt regardless
of worker completion order. Each new bridge attempt starts with its own stage state.

A UNC share root (`\\server\share`) is the topmost container of its route: the host's provider
parent derivation reports no parent for it (mutation-interlock ancestors, rename and create
admission stop there instead of binding `\\server`), and the Local provider opens a share root with
the trailing separator Win32 requires (`\\?\UNC\server\share\`); deeper paths are untouched. The
Local provider treats an extended UNC path (`\\?\UNC\server\share\…`, `UNC` in any case) as its
plain form: it is a UNC path on the `local-win32-smb` route, its storage characteristics and transfer
hints are those of a high-latency network share, and admission gives it the root identity of the
plain share root `\\server\share\`. A Win32 device path (`\\.\…`) is never a UNC path.

Once application exit is committed, main-window command dispatch is fenced before auxiliary-window
teardown or the final main-window destroy is queued; no keyboard, function-bar, menu, or posted
command may start another operation behind that fence. A synchronous admission/confirmation prompt
that dispatches the accepted close MUST stop its nested message drain as soon as the prompt completes,
its overlay becomes hidden, or application prompt shutdown begins; an overlay destroyed with its owner
window counts as hidden even when another alert on the same window was cleared first. It then returns
cancellation to the admission caller and unwinds that caller before dispatching any queued owner-destroy message. Admission also fails with
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

Every worker admission, including Batch Rename/Change Case, MUST synchronously
arm an owner-window recovery heartbeat before releasing that worker. Timer admission
failure rejects the operation before mutation. `StartOperation` arms the heartbeat before
any admission prompt (operation-artifact touch, Copy/Move confirmation, Delete confirmation),
so a timer failure is refused before the user consents, and arms it again at publication
because a modal prompt can let the idle heartbeat stop.

A refused admission is never silent. Pane Copy and Move (F5/F6, with or without options),
Delete (F8) and Permanent delete (Shift+Delete) route every failed admission result, including
a heartbeat failure after an accepted prompt, to the source pane's localized operation alert,
titled with the operation and naming the HRESULT; nothing was started, so no task card
exists. A cancel (`S_FALSE`, `ERROR_CANCELLED`) or an admission refused for shutdown shows
nothing, and neither does an operation-artifact touch the guard blocked: its blocked prompt
(Retry / Open location in File Explorer / Close, see "Visibility and touch guard") already said the
operation was not started, so `StartOperation` and `ConfirmExternalArtifactTouch` return
`FolderWindow::kFileOperationRefusedAlreadyReported` (`HRESULT_FROM_WIN32(ERROR_REQUEST_REFUSED)`)
and the pane adds no second, generic alert (a FolderView request receives it as `S_FALSE`). Change Case reports a refused admission in its own pane alert, and Batch Rename in
its dialog's failure report.

The heartbeat remains armed while any
task is owned, including terminal tasks awaiting completion consumption, and stops
when no work remains. Recovery does not depend on popup creation/visibility or on
successful posting to a full window queue.

A fallback completion is pending until an admitted reaper takes ownership of the exact task's
`jthread` and joins it. Only after join may the matching fallback become eligible
for owner-thread consumption; a terminal flag or transferred/nonjoinable handle is
not a worker-exit proof. The reaper also attempts that task's UI wake. Handle transfer,
callback admission, and failure restoration share the transfer mutex with
shutdown's task-worker scan. Failed admission restores the `jthread` to `Task`
before the final payload-less wake. The owner heartbeat retries off-thread drain
admission without joining a potentially running worker on the UI thread.
If threadpool admission remains unavailable, a zero-time wait returning signaled
on the exact native thread handle is a worker-exit proof; only then may owner
recovery join that already-exited handle and mark the fallback eligible without
allocating a reaper. A nonjoinable/transferred handle or task-finished flag alone
never grants this exception.

Early wakes leave pending entries intact; joined entries are consumed exactly once by
either the normal wake or heartbeat, even if every wake fails. Shutdown consumes
the remaining entries only after joining all workers and reapers. Stale nonzero
payload tokens MUST NOT be treated as payload-less fallback wakes.

Completion dispatch retains state through callback reentrancy, but synchronous
owner shutdown may still destroy individual tasks. After each external completion
callback, owner processing MUST re-find the task and stop if it or the owner was
destroyed; retaining a raw task pointer across the callback is not lifetime proof.

Shutdown closes admission and waits for the outstanding reaper predicate before
destroying `FileOperationState`. A callback decrements and notifies under that
predicate's mutex as its final access to the state; observing zero after
reacquiring the mutex is the lifetime barrier. This preserves the forced-destroy
ownership boundary and the normal off-UI cancellation drain described above.

## Results and consumers

How these results are worded on task cards, prompts and the footer (plain language, counters,
notes, terminal facts) is owned by `../UI/UI_FileOperationsPopup.md` "User-facing policy".

Per-item executable axes are `PublicationState`, `VerificationState`,
`SourceDisposition`, `OwnedStageDisposition`, and `ItemCompletion`.

`SourceDisposition::Missing` (appended numeric value 3; earlier values remain stable) means the
host's source-side no-follow transfer guard positively observed absence before a provider mutation
attempt. It is neither removal by this operation, retained-source existence, nor an unknown
mutation outcome. Preserve the real missing-file/path HRESULT and Failed completion, count the item
separately in completed facts, and exclude it from retained-source paths and recovery actions.
Do not infer this disposition from a generic provider `ERROR_FILE_NOT_FOUND`/`ERROR_PATH_NOT_FOUND`,
which may describe a destination or intermediate path; provider receipt and post-mutation truth
keep their existing contracts. Source-row consumers continue to remove/ghost rows only for Removed,
so Missing never claims a successful Move. See the popup spec for its source-specific explanation.

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
| NotPublished + Retained + Completed for a proved clean Copy no-op | Copied | Success; destination unchanged |
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
| User cancel whose only unknown-axis items are Canceled | `ERROR_CANCELLED` / presented Indeterminate with a warning |
| Skips whose only unknown-axis items are Skipped | `S_FALSE` / Completed with skipped items (axes stay on the item) |
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

This missing-receipt rule applies after the selected root crosses the mutation gate, including a
null final callback that supersedes earlier child/attempt truth and cancellation without a terminal
item callback. Move/Delete source disposition is Unknown until proved; Copy retains its source, and
Delete has no destination publication (`NotAttempted`). A root still in Preparing without a valid
receipt is proved untouched; a malformed receipt cannot invent pre-gate mutation or stage ownership.
Valid receipt truth keeps its authority, including an explicitly unknown outcome.
Recursive Local Copy workers MUST contribute to the selected item's shared destination mutation
facts before the root reports completion. Publication, uncertainty and retained-stage facts survive
worker completion and retries; a merge into an existing directory cannot report no-commit after a
child was published. These facts belong to one selected item, never the bulk call's shared scheduler.
The scheduler joins borrowers before the owning context or its terminal receipt is released.
Local `CopyItems` and `MoveItems` MUST report each completed selected root's typed receipt, just as
their single-item entry points do. Sequential bulk contexts reset item mutation facts between roots,
after borrowers have joined; parallel contexts retain independent item facts. Neither a previous
root's publication nor its cleanup debt may influence another root's receipt.
A valid no-commit receipt remains Retained, and a valid committed receipt retains its precise axes.
Local Native Move reports a known failure with the real error when it issued no mutation. A source
that is missing before any rename (top-level discovery, the rename's own source probe, or the Keep
Both re-probe) and an atomic rename refusal, including a missing destination parent
(`ERROR_PATH_NOT_FOUND`), are known no-commits (`TRUE/FALSE/TRUE`): the item is `Failed` with that
HRESULT, publication `NotPublished`, the source as this item left it (`Retained`), and Retry
remains legal (it fails again if the source is still gone). Only an unproved rename/revert outcome,
or explicit conditional source uncertainty (a `RenameIfUnchanged` non-commit that reports the
original absent), makes the receipt unknown and cannot authorize Retry; known publication remains
authoritative. Sequential bulk contexts reset this source uncertainty between selected roots.
Structural prefix validity is insufficient: a known no-commit receipt claiming an absent original or
a Published owned stage violates the shared mutation contract. After possible mutation it is rejected
as `Indeterminate`/`ERROR_IO_INCOMPLETE`, with unknown mutation and owned-stage axes (Copy retains its
source invariant; Delete has no publication axis). Skip, a cancellation status, and a later task cancel
cannot override this rejection. Before the host mutation gate, an invalid receipt cannot invent work
or stage ownership and the host's untouched-source proof remains authoritative.
The pure item-result reducer owns these provider-evidence rules for generic and qualified routes;
host-owned Managed publication/partial-tree and cleanup proofs retain their stronger explicit truth.
An explicit per-root Skip remains `Skipped` alongside a known receipt's exact publication/source
axes, including partial publication. An explicitly unknown or malformed receipt, or unresolved
owned-stage cleanup, remains `Indeterminate` even when Skip was requested; Skip supplies no missing
storage proof. A descendant Skip with a missing receipt remains terminal `Skipped`, but its
publication/source axes remain unknown after possible mutation (Copy still proves source retention).
Only a whole-root Skip before mutation proves `NotAttempted`/`Retained`; a host-owned bridge proof
retains its precise partial-publication and source axes. Whether a root consumed a descendant Skip
is decided from the executor's own status: a root that returned success after a descendant Skip is
partial work even when a Copy-only or source-kept Move then reports it as `Copied; source kept`
(`S_FALSE`). It stays partial `Skipped` work with its committed facts, and its copied bytes stay in
the task's progress; only an executor `S_FALSE` before mutation is a whole-root Skip. A known no-commit Copy receipt with exact
`S_OK`, a retained original and clean stage (`NotApplicable`, `NotCreated`, or `Removed`) may fulfill
the requested intent without changing the destination: merging an empty directory or an entirely
existing empty directory tree is `Completed`/`NotPublished`/`Retained`. It MUST NOT fabricate publication
or modify destination metadata merely to claim a commit. Skip/cancel causes, contradictory receipts,
and unresolved cleanup retain their existing precedence. A `Failed` item MUST carry a failure HRESULT:
other provider-success/known-non-commit combinations, including Move/Delete with a retained original,
are normalized to `E_UNEXPECTED`, preserving exact storage truth and preventing aggregate success.
A known committed Delete receipt with a retained selected root likewise MUST NOT be `Completed`:
without an explicit Skip/cancel cause, provider success is `Failed`/`NotAttempted`/`Retained` with
`E_UNEXPECTED`. Mutation commitment may describe partially removed children; it does not prove
that the selected root was deleted. A provider's failure HRESULT, explicit Skip/cancel, cleanup
uncertainty and a later task cancellation retain their existing precedence and exact storage axes.
The host-owned bridge's exact publication facts also remain authoritative on success: a successful
empty-directory merge with `NotPublished` proof MUST NOT become `Published` merely because its
completion is successful. Provider-receipt and host-bridge controls exercise these separate routes.
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
removal focus, Find row removal, and Compare invalidation are all reduced from those records. The
partial-Move `task.result` diagnostic likewise states the typed counts of moved, kept and unknown
sources; it never claims "source preserved" for a root whose source was removed. There
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
ends `S_FALSE`. The per-item Copy/Move/Delete executor crosses that gate one selected root at a
time, when it dispatches the root to its executor; a root never dispatched before a user/host
cancel, or before another root's failure stops the task (whether or not that failure requested
cancel), stays `Preparing` and finalizes `NotAttempted + Retained + Canceled` with
`ERROR_CANCELLED`, never Unknown and never the failing root's status, so it keeps its place in the
retained-source actions and cannot make the summary "Outcome unknown". A root that recorded its own
pre-dispatch status keeps it, and a setup failure before any root is dispatched, with no cancel
requested, finalizes every root `NotAttempted + Retained + Failed` with that status. A bulk
provider call and the rename executors can touch every selected root at once and cross the gate
for all of them together.
Terminal storage is compare-and-store:
invalid indexes and duplicate writes are diagnosed and rejected without changing the first terminal
truth. Every worker completion path performs the same dense final reduction.

The task acknowledges **“Move started; the cut list was used up.”** A consumption failure says that the Move
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
Completed, Partial, Canceled, Failed, and Indeterminate. Discovery-ahead is secondary activity.
After Skip counting, enumeration continues internally just in time, but the card withdraws discovery
counters/activity and shows one truthful completed-work summary. It keeps the operation name in
the header, with no unknown `0/?` fraction, footer counting label or provisional still-counting ETA.

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
  color/hatch state. There is no second painter: when the hosted control tree cannot attach or the
  Direct2D target keeps failing (a transient failure is retried), the popup shows its native failure surface
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
| Possible | Current/legacy generated artifact-name shape only; never an ownership claim | Always visible with `Possibly incomplete file (left by an interrupted copy or move)`; Inspect/Reveal and, for a Local path, Open location in File Explorer remain ordinary inspection |
| Ordinary | No artifact provenance evidence | Normal user item |

### Artifact source of truth

Artifact classification is deliberately non-authoritative. The registry creates one empty in-memory
generation and never opens or enumerates the retired Batch Rename journal root. Only the recognized
generated-name shapes `.rs_tmp_`, `.rs_copy_tmp_`, `.~rs-write-`, `.rs_bak_`, `.rs_ren_` and
`.rs_case_tmp_<32 hex>` (the Local case-only rename's temporary name), and the provider staging names
(I35 R7.6) `.rs-bak-<32 hex>[-<leaf>]` (an S3 backup key), `.redsalamander-upload-`, `-writer-` and
`-rollback-<32 hex>` (Curl; `rollback` also Microsoft Drive) and `.<leaf>.rs-mtp-overwrite-` /
`.rs-mtp-create-<32 hex>.tmp` (MTP writer temps) produce `Possible`; all other names are Ordinary. Neither a name shape, pane row, search-index field, badge,
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
reason `name shape only`. The badge reads "Possibly incomplete file (left by an interrupted copy or
move)", except for the shapes that may hold the only complete copy of the user's data, which read
"Possibly an original kept by an interrupted copy or move" (`Projection::mayHoldOriginal`): the backup
(`.rs_bak_`), the rename hop (`.rs_ren_`), the case-only rename's temporary name (`.rs_case_tmp_`),
the S3 backup key (`.rs-bak-`) and the remote rollback sibling (`.redsalamander-rollback-`); it is never an artifact-specific suppression rule and carries no task,
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

A blocked request explains itself (I35 R7.1). Each candidate records why it has or lacks a current
exact identity (`FileOperationArtifacts::IdentityProbe`): the binding ABI returns one HRESULT, so the
host derives the pair from the binding state. An `Unsupported` binding means the provider opened the
object (or has no binding interface) and the identity query failed: the open status is `S_OK` and the
identity status is the binding's result, normally `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)` (a Local
path on a redirector without `FileIdInfo`, such as the NAS probed on 2026-10-11). Any other failure is
the no-follow open's (sharing violation, access denied, delete pending) and the identity query was not
reached. For every blocked item the host logs `Debug::Warning` with the path and both statuses and
emits `fileops.artifact.touch.blocked{openHr, identityHr, outcome}`. The prompt names at most three
items as `<leaf> (0x<HRESULT>): <reason>` (then `…and N more`) with the plain reason for each class,
keeps the fail-closed sentence `The operation was not started.`, and offers only non-granting actions:
`Retry` when an item failed to open and the caller can capture the candidates again (typed admission
and external touches; never a worker-built Batch Rename request), `Open location in File Explorer`
when an item is on the Local provider, and `Close`, which is the default and the Escape action. Retry
captures every candidate again and asks again; Open location is the inspection-class shell action
below and still refuses the request. The closing hint "To check or remove these items, open their
location in File Explorer." appears only when a Local item lacks an identity because the location
cannot provide one; an item that could not be opened is explained by its own Retry line and is never
offered a by-name delete. Closing a blocked prompt is a refusal the prompt already explained:
`StartOperation` and `ConfirmExternalArtifactTouch` return
`FolderWindow::kFileOperationRefusedAlreadyReported`, non-recursive Change Attributes shows no second
alert, and recursive Change Attributes ends as canceled without a failure record. A guard that could
not run at all (capture or revalidation failure) shows the pane alert "…Red Salamander could not check
them, so the operation was not started." `AcceptTouchGuard` mints a receipt only when the aggregate
flag and every item's own `revalidatable` flag agree.

A Delete of a Possible item that its location opened but cannot identify (I35 R7.3/R7.6) is not a
receipt item when that Delete has its own consent for such an item: on a provider without object
binding (Curl's FTP, SFTP and SCP, S3, Google Drive and Microsoft Drive; their C10 by-name card or
native route, an MTP device cannot delete at all) and on the Local by-name route. Handing the Possible
items of providers without object binding to the C10 card, with the leftover warning, goes beyond the
plan's Local-only partition (an accepted S4 departure, flagged to the owner): without it the R7.6
staging names would make their leftovers undeletable. A permanent Delete hands the item to its card, which then carries the leftover warning; a
Recycle keeps the guard's warning (OK/Cancel, and Continue mints no receipt for the item) and a by-name
escalation repeats the warning. Exact-identity items of the same request keep their receipt, and the
worker revalidates only those: `RevalidateTouchGuard` leaves the handed-off items out, and an item that
was handed off but has become identifiable no longer matches the receipt. Every other touch of such an
item (Move, Rename, overwrite, publication target, external actions) and any item whose open failed
stays refused as above. Without this partition, registering the provider staging names would make
the Delete of an FTP, SFTP, SCP, S3, Google Drive or Microsoft Drive leftover impossible. A Pack or
Unpack cleanup never hands an item off (`DeleteHandoffForPlan`).

**Clean Up Leftovers of Interrupted Operations** (I35 R7.5, D-7.2) is a pane command
(`cmd/pane/cleanUpLeftovers`; the Commands menu, and the artifact submenu of an item menu whenever the
listing has a Possible row). It opens a modeless DxUi-hosted listing of the current folder's Possible
rows (name, size, modified time, kind), optionally including subfolders through a bounded, cancelable
worker walk that never follows a link and lists a folder with a Possible name without entering it; its
bounds (4096 folders, 4096 rows, 32 levels) each report their own line when reached. Cancel is the
default and the Escape action. Shapes that may hold an original are listed as "Possible original" and
never pre-checked; a row the user checked or unchecked keeps that state when the subfolders are
included or left out again. Rows in a folder that a live task of this window works in (a prepared or
running scope on that folder, an ancestor or a descendant: a running Copy's stage is a sibling of its
destination) are left out and counted, before the listing shows and again when Delete Selected is
pressed, so the note can say that no operation running in this window works in their folders; it also
says the items may be left over with incomplete data or still in use by another window or program.
**Delete Selected** submits one ordinary `DeletePlan` (`DeleteOrigin::LeftoverCleanup`) through
`StartOperation` on the provider and plugin identity the listing was opened on (not the pane's current
one): to the Recycle Bin where the provider supports it, as the pane Delete does, else a permanent
Delete behind its card, so the touch guard, the confirmation and the by-name rules apply unchanged. On
a provider that cannot delete (an MTP device) the listing shows its rows unchecked, disables Delete
Selected and says why. Closing the listing stops its walk, re-issuing the I/O cancel for a bounded
200 ms; a walk still blocked keeps the closed listing alive until it returns (the folder window waits
for it only when it is destroyed). The listing deletes nothing itself, never opens by itself or at
startup, has no entry in the pane background menu (only the Commands menu and the artifact submenu, an
accepted S4 departure), and leaves the interrupted-Move breadcrumb untouched.

`Open location in File Explorer` (I35 R7.2) is inspection-class and prompt-free: it opens File Explorer
on the item's parent folder with the item selected (`FileOperationArtifacts::OpenLocationInFileExplorer`,
`SHOpenFolderAndSelectItems`) and grants no authority. The shell call runs on its own
single-threaded-apartment thread-pool callback, so it works when the prompt was answered on an MTA
worker (recursive Change Attributes) and an unreachable share never blocks the UI thread. Only a Win32 path, the built-in Local provider's,
offers it; the Possible projection sets `capabilities.openLocation` only for that provider.

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

The destination captured at admission is fixed for the task's lifetime. Queued destination changes
are unsupported: there is no candidate preparation, scheduler replacement reservation, destination
change consent, refusal notification or admitted-breadcrumb replacement API. Cancel and submit a
new request to choose another destination. Initial admission, worker-side destination-name
validation, overlap decisions, mutation interlocks, clipboard admission and monotonic Move breadcrumb
phase transitions retain their ordinary contracts.

Overlap permissions and Queue-after relations identify an exact admitted generation through a
host-unique admission key. Stable task IDs remain the public UI/deque identity. Admission keys equal
task IDs. Run grants authorize only that admitted operation's scopes, including in the live-output guard. A predecessor relation releases when its exact admission is withdrawn.
Prepared/active registrations own immutable scope and decision snapshots; they never borrow
authority or consent arrays from the current Task. Advisory completion replaces the small decision
snapshot while sharing its immutable scope vector. A retained older view stays valid and unchanged.
Execution readers retain an owner for the full lifetime of any scope reference or iterator; empty
preparation exposes no authority. Withdrawal names one exact admission key. The live-output index
captures the key and borrows scopes only while its active registration owns them under the queue lock.
The Task publishes one immutable selected-admission view containing the destination provider,
folder, pane, provider context, resolved mappings, typed plans, artifact receipt, preparation facts
and interlock snapshot. Aggregate UI, execution, completion and diagnostic readers retain one
view instead of combining independently loaded fields. A null destination provider keeps its
original source-provider fallback meaning. Initial construction and Rename/Delete plan resealing
share heavy destination owners. An empty bootstrap view
contains no plans or authority. Accepted destination inputs cannot change. Execution and completion retain this view and exact
admission key; no reader may reconstruct a selection from separate mutable fields.

Artifact-touch preparation returns an ephemeral receipt for the proposed plans/provider objects.
UI prompt replies carry a monotonically increasing per-task request ID, are accepted at most once,
and cannot install the live task receipt directly. Failure or cancellation preserves the caller's
prior output. Success with no possible artifact clears that private output. The caller installs a
successful receipt only when it selects the corresponding admission. A drained/stale nonzero
message token is ignored; the explicitly payload-less shutdown fallback may only report failure.

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
- `fileOperations.verifyAfterCopy`: Boolean, default false;
- `fileOperations.mirrorDecisionsInPane`: Boolean, default true; presentation only (the decision
  mirror over the pane, see "Conflict model"); it never changes what is asked, the action policy or the
  answer.

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

Private preparation may index duplicate scopes and retained ancestors by the canonical path key
within an identical qualified endpoint/path contract. A key hit MUST still satisfy the original
endpoint and `EquivalentPath` predicates and retain the earliest matching authority. Unkeyable
paths, a different incoming contract at the same root, or an exhausted index budget MUST use the
original pairwise predicate; a missing key is never object authority.
These indexes hold vector ordinals, survive preparation-vector reallocation, and are discarded
before immutable scopes are published. They retain at most 256 qualified-contract buckets,
32,768 scope/ancestor entries and 16 MiB of key capacity, including terminators. Exhausting
the contract-bucket cap disables all indexed lookups, including existing buckets, so a later
same-root contract cannot be omitted from an earlier bucket's query. A bucket that already
requires the original scan adds no further unused keys. Index storage is additional temporary storage,
not a cache of provider authority across tasks. `FileOps.Interlock.PrepareLookupChecks` records
scope/ancestor predicate evaluations, and `FileOps.Interlock.PrepareIndexStorage` records key
capacity bytes and entry count once per preparation. Neither metric emits provider path text.

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
  limit hit. Before each directory level the walkers also check the running thread's remaining stack
  and stop with `ERROR_STACK_OVERFLOW`, like the depth ceiling, when less than 128 KiB is left, so a
  deep caller or larger (Debug) frames end in a reported failure, never in a stack overflow; the
  parallel walker's per-entry work runs outside its recursive frames (I35 S2), and
  `Phase7_CopyItemsSingleFolderRecursiveParallelism` copies 96 levels on the default stack and proves
  the stop from a 256 KiB thread;
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

### Transfer phases, stalls and throughput tuning (I35 S6)

The progress ABI carries no phase. The host derives what each stream is doing from its reports and
attributes every silence to it (R10.1, no ABI change):

- The Local provider announces each non-empty file at zero of its bytes once its size is known,
  before its open, create and metadata preparation (**Opening**; that report honors a cancellation,
  nothing exists yet). Chunk reports are **Transferring**; a report whose item bytes are all written
  is **Finalizing** until the stream names another item; a report without item bytes (0 of 0) names
  no phase (**Unknown**), so a silence after it is attributed to Unknown. Its native pumps add one
  zero-delta report at a phase boundary: after the commit flush on the direct-final route, and after
  the commit and after the finalize metadata on the staged-overwrite route. A boundary report is an
  observation point only: it ignores the callback's result, so a cancellation is met at the item's
  next ordinary check and a boundary never changes what a cancel leaves behind; a boundary within
  50 ms of the previous report is skipped.
- The host bridge names its phases explicitly: Opening before its writer and stage exist (honors a
  cancellation; on a connection with a copy/move limit, before the permit wait), Transferring for pump
  chunks, Finalizing from the always-sent report of the file's last bytes (a file of unknown size
  reports it before commit-and-prove), Publishing before publication and for link records (sent only
  when the stream's last report is a progress period old; a small file keeps its indeterminate item
  sample), Verifying for exact readback. Its Finalizing and Publishing boundaries are observation-only
  like the Local ones.
- Per task the host records, under its stream-statistics lock, the inter-report gaps within one item
  per stream (a histogram of 1 ms buckets below 200 ms, 5 ms below 1 s, 50 ms below 10 s and 1 s
  below 120 s; P95/P99, maximum, gaps over 600 ms by the phase announced before them), the task-level
  silences over 600 ms by the phase the streams had announced, the zero-delta Finalizing reports that
  repeat a stream's item (`finalizingRepeatCount`: the boundaries after a commit are seen), and
  whole-second throughput buckets from the first byte (coefficient of variation). At task end it
  emits `FileOps.Progress.Cadence.GapP95Us`, `.StallUs`, `.SilenceUs` and `.ThroughputCvPermille`.
  The Local pumps time their silent phases as `FileOps.Local.FileOpenUs`, `FileCommitUs`,
  `FileFinalizeMetadataUs` and `FilePublishUs` (`detail` names the route, `value0` the file bytes).
- The phase the streams had announced and the slower side of the transfer
  (`FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY` on exactly one side) name the popup's stall text
  (`Specs/UI/UI_FileOperationsPopup.md`, "Speed and Stalls"). A destination admitted without a
  provider of its own (a bridge transfer between two roots of one provider, such as a Local Managed
  Move from a local volume to a network share) is served by the source's provider, and that provider
  describes the destination's storage: it names the side and sets the destination's half of the
  concurrency budget exactly as a destination provider of its own would
  (`FileOps_SameProviderBridgeDestinationStorage`).

`FileSystemOptions.transferTuning` carries the R10.3 engine candidates as flags, so one build can
measure each against the same corpus. The host fills it per route: a provider's own Copy/Move from
`kAdoptedProviderRouteTransferTuning` (`FILESYSTEM_TRANSFER_TUNING_NONE`) and the host bridge's
provider calls from `kAdoptedHostBridgeTransferTuning` (`WRITE_THROUGH_REMOTE` since 2026-10-13, see
below); a self-test may override both:

- `NO_TEMPORARY_REMOTE` / `NO_TEMPORARY_LOCAL` (candidate e): the stage or direct-final handle is
  created without `FILE_ATTRIBUTE_TEMPORARY` on a remote (UNC or `DRIVE_REMOTE`) / local
  destination;
- `PREALLOCATE` (b): the Local pumps reserve the destination allocation (`FileAllocationInfo`) of a
  non-sparse file larger than one copy buffer; a refusal is ignored and the exact-abort and
  owned-stage cleanup paths remove an incomplete object as before;
- `SKIP_ABSENT_BACKUP_STREAMS` (d): the finalize metadata pass skips the `BackupRead`/`BackupWrite`
  copy when the source has none of the backup-stream features (the inspection passes still run, and
  the Overwrite route's metadata transfer is otherwise unchanged);
- `DESTINATION_CONCURRENCY` (c, host only): the provider's own Copy/Move route budgets
  `min(source, destination)` concurrency, as the bridge always does, and records the destination
  storage kind in `autoConcurrencyDestinationStorageKind`;
- `WRITE_THROUGH_REMOTE` (w): on a remote destination the stage or direct-final handle is created
  with `FILE_FLAG_WRITE_THROUGH` and without `FILE_ATTRIBUTE_TEMPORARY`, so each write returns once
  the server holds it, and the Local pumps write 8 MiB per call there instead of 1 MiB (the
  provider's own transfer hint for a high-latency destination, which the host bridge already uses).

A candidate is adopted only when, on the NAS corpus of `FileOps_SmbCopyThroughputVariance`, its P95
inter-callback gap and its 1 s-bucket variance are lower than the baseline's with equal or higher
wall-clock throughput, and local NVMe-to-NVMe throughput regresses by no more than 5 %. NVMe or
loopback evidence alone never adopts a candidate. On the registered NAS (2026-10-12) none passed: the
long silences are commit flushes (the 1 GiB file's flush alone about 12 s at about 54 MB/s), `e` and `c`
are neutral, `d` halves the finalize-metadata time without lowering the variance, and `b` lengthens the
flushes. A pipelined pump (candidate a) adds outstanding writes but cannot shorten a flush bound by the
server's persistence rate, so it is not built.

Bounding the dirty data a writer may hold was measured next (2026-10-13, owner report: a Managed
Move of a 15 GB file to the NAS read `Waiting for storage…` with a fully hatched graph while the
network was busy). A cached writer on a network share takes data at memory speed until the client
cache's dirty-page threshold (282,198 pages, about 1.1 GiB, on the 61.6 GB reference machine) and is
then held inside one write while the cache drains to the server, so a file larger than the threshold
stalls in the middle, not only at its commit. A probe with the stage's own open flags and 8 MiB
synchronous writes to the NAS held one write for 48.5 s after the first 2.5 GB and 85 of 2,560
writes for over 600 ms (20 GiB, 53 MB/s with its 19 s commit flush); without
`FILE_ATTRIBUTE_TEMPORARY` (candidate e) the longest hold was still 19 s; a flush every 64 MiB was
slower (47 MB/s) with 1.2 s gaps; written through, no write took more than 0.35 s at 52 MB/s with
no commit flush left, and unbuffered writes measured the same but need sector-aligned I/O. The gate,
interleaved base/w runs on the NAS, `Specs/TestRuns/4cb089111a23/FileOps/`:

- host bridge, `large` corpus (4 GiB + 16 x 16 MiB, run `20261013T142330Z-62100-…`): w moved 52.4 and
  52.8 MB/s against 50.3 and 52.6, its P95 report gap fell from 1,900 ms to 1,050 and 930 ms, its 1 s
  coefficient of variation from 1.26 and 1.99 to 0.15, and its longest task silence from 14.7 s to
  0.48 s (no silence over 600 ms against 15 and 35): **adopted for the host bridge**, whose writers
  serve every transfer it carries (a Managed Move such as a Local Move to a network share, a
  cross-provider transfer, a Verify On copy the provider does not prove itself);
- provider's own route, `full` corpus (1 GiB + 64 x 16 MiB at concurrency 8): written through in
  1 MiB writes, w lost 13 % of the rate (45.9 against 52.6 MB/s, run `20261013T141900Z-62100-…`);
  with 8 MiB writes it still lost 7 % (median 50.1 against 53.9 MB/s over three pairs, run
  `20261013T161108Z-63304-…`) and its per-stream P95 gap rose from 1,100 ms to 1,550 ms (eight files
  share each server round), although its longest task silence fell from 12 s to 0.6 s and its
  variation from 1.10 to 0.14: **not adopted** on that route, which keeps its cached 1 MiB pump and
  names its commit flushes `Finalizing file…`. Run `20261013T143758Z-33940-…` measured a degraded
  NAS (33 MB/s in a write probe) and decides nothing;
- local reference (run `20261013T141002Z-70040-…`): unchanged within noise (1,235 and 1,564 MB/s
  base, 1,408 and 1,546 MB/s w), as the flag applies to remote destinations only.

Each written-through write waits for the server, so a high-latency WAN share may write slower than
through the cache; the registered NAS is on 10 GbE.

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
- Recycle escalation identity and no Apply-to-all checkbox;
- identity-less Local locations (I35 R7.3/R7.4/R7.6, `FileOpsFamily_IdentityLess`): on the loopback
  alias with the bind injection, the by-name card (Cancel keeps the item, Delete by name removes it
  with one Information note), a size change on the deleting handle and a last-write change before the
  answer refused with `ERROR_REVISION_MISMATCH` and nothing deleted, a planted leftover handed to the
  card without a touch-guard prompt, single- and two-item Recycle escalation by name with one question,
  and a Recycled leftover warned without a receipt (`Delete_IdentityLessLocationByName`); a reader
  without delete sharing after the pin gets the sharing question with Retry and both results logged, an
  `Indeterminate` root is refused at Preparing with both results, and a path another live task holds is
  refused (`Delete_IdentityLessSharingAndInterlock`); with queue mode on, a Move run in parallel into a
  folder whose delete by name waits in the queue is asked (the delete-by-name question, Queue after these
  tasks the default, no Run at the same time) and, queued after it, loses nothing
  (`Delete_IdentityLessQueuedRootOverlap`); a Move into the folder through its local spelling while the
  delete by name of its loopback spelling waits runs first and the delete then refuses the changed tree
  (`Delete_IdentityLessPinnedTreeOtherSpelling`); an archive Pack writes into the folder while the
  delete waits and the archive survives (`Delete_IdentityLessPinnedTreeArchiveWriter`); a Copy into the
  folder admitted while its card waits is asked about the delete, and a file added deep in the tree
  before the answer refuses it with nothing deleted (`Delete_IdentityLessPinnedTreeCardPending`); a file
  written into the tree after the delete's last listing (a provider pause between the check and the
  first deletion) survives with its folders, the listed files are gone and the stop is reported as a
  partial delete (an Error naming 2 of the 3 listed items, the item's committed and original-present
  receipt; `Delete_IdentityLessPinnedTreeConcurrentWriter`); a folder over a test-only entry bound is
  refused at Preparing with its plain reason and no card (`Delete_IdentityLessPinnedTreeBudget`); a
  listed file changed or removed at that pause stops the delete with nothing deleted and a known
  non-commit receipt (`Delete_IdentityLessPinnedTreeChangedEntry`); a refused Recycle escalated by name
  stops on a file written at that pause and reports the same partial delete
  (`Delete_IdentityLessPinnedTreeEscalationPartial`); a listed folder swapped for a junction at that
  pause, with a listing that reports no file index and a decoy forged with the listed file's facts, never
  redirects the delete (`Delete_IdentityLessPinnedTreeLinkSwap`); a sub-folder the location refuses to
  open with `ERROR_NOT_SUPPORTED` is an open failure (`Delete_IdentityLessPinnedTreeSubfolderOpen`); a
  Recycle over the bounds keeps its reason for its refused escalation
  (`Delete_IdentityLessPinnedTreeRecycleOverBudget`); a decided holder of the folder refuses the delete
  at the re-check after the answer and at the re-check right before it
  (`Delete_IdentityLessPinnedTreeHolders`); and a tree of 2,466 entries is deleted by name on the loopback
  alias, whole and again with a test-only budget of four held folders (so it goes in chunks), and, where
  `path.smb.primary` is configured, on the share, its timings logged (`Delete_IdentityLessPinnedTreeScale`);
  a child added while the card waits or between the answer and the delete refuses with
  `ERROR_REVISION_MISMATCH`, and a task admitted after the answer, before the re-check or while the
  delete is about to run, is asked about the delete (no Run at the same time) and its pending question
  does not refuse it (`Delete_IdentityLessFolderChangeAndHolders`); an open refused with
  `ERROR_NOT_SUPPORTED` is a failed open (no card, the leftover refused by the guard) and Pack and Unpack
  cleanups never delete by name (`Delete_IdentityLessOpenFailureAndArchiveOrigins`); the touch-guard
  partition and the R7.6 name shapes (`FileOps_IdentityLessGuardPartition`); the leftover listing's walk
  bounds, stop and link rule (`FileOps_LeftoverCleanupWalkBounds`); a bridge stage removal delayed one second under a requested
  cancel completes inside the cleanup budget (`FileOps_OwnedStageCleanupBudget`, with the Local cancel
  watch debug self-test's cleanup-window scope); and, only where `path.smb.primary` is configured, the
  registered NAS share binds `Unsupported`, corroborates with its 64-bit file index and deletes a file
  and a small folder (its pinned tree compared entry by entry on the share) by name
  (`Delete_NasFileIdInfoUnsupportedByName`);
- a multi-item Recycle whose items (files and a folder) all fail with a known non-commit asks once;
  the task-scoped Delete all remaining permanently choice, offered only under More, deletes every
  item, each Completed with its accepted decision and escalation recorded; a new task asks again,
  Skip covers one item and Cancel ends the task; an item whose Recycle outcome is unknown is never
  covered by the grant (`Phase10_RecycleEscalationTaskScopedGrant`);
- the BulkItems executor (pane Delete) with continue-on-error keeps every failed Recycle without a
  prompt, reports it Failed and continues; without it, a pane-equivalent non-empty folder escalates
  and an accepted Delete permanently removes the whole tree, and a multi-item task goes on past the
  provider call's first failure: a Skip asks again for the next failing item, and Delete all
  remaining permanently then covers the last one without another prompt
  (`Phase10_RecycleEscalationBulkItemsContinueOnError`);
- a pane-equivalent BulkItems Delete of 1,000 files whose Recycle all fails asks once, and Delete
  all remaining permanently deletes every file with at most two provider calls, exact discovery
  totals and no growth after close (the elapsed time is traced); a Recycle batch entry offered
  without a terminal observation, next to a known failure in the same call, is never replayed,
  prompted, covered by the grant or deleted and stays unknown; a sibling worker's Recycle held in
  flight while another item's failure stops the call is never reported canceled and is resumed after
  the one prompt; the committed entries of a batch observed while another worker's failure stops
  the call are all reported recycled and never resumed; an escalation that fails because its item
  was replaced after the Recycle attempt ends the task there, so the failed items after it are
  never prompted or deleted under the grant and stay, reported retained; and a single-item Recycle
  the Shell holds when another item's failure stops the call, whether before or after the Shell
  recycled it, is recycled and reported once, with exact discovery and no resume or unknown outcome
  (`Phase10_RecycleEscalationBulkItemsResumeBounded`). A folder escalated after its Recycle fails
  adds its descendants to the task's discovery totals without growth after close
  (`Phase10_RecycleEscalationBulkItemsContinueOnError`);
- Recycle on a location without a Recycle Bin (the loopback administrative-share alias of the
  sandbox): a file, a folder, and a sibling batch are each kept as a known non-commit, and the host
  asks for exact-item permanent-delete consent instead of reporting a Recycle
  (`R0fSmb_LoopbackRecycleNeverDestroys`; the per-item verdict is also a FileSystem debug self-test);
- a Recycle the provider reports as destroyed (`COPYENGINE_E_RECYCLE_FORCE_NUKE`, committed removal)
  fails with `delete.recycle.destroyed` in both the per-item and BulkItems executors, without a
  permanent-delete prompt (`Phase10_RecycleDestroyedIsNeverRecycled`);
- a non-empty folder whose Recycle fails with a known non-commit escalates to one consent that
  reports the whole tree's item and byte counts, and an accepted Delete permanently removes the
  tree recursively and completes the folder as removed (`Phase10_RecycleEscalationDirectoryTree`);
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

## Recursive Copy progress units and lifecycle

Folder Copy has four independent facts: discovered child files/folders and their payload sizes,
written bytes, actually published child files, and selected-root outcomes. A selected folder is
one root result but can contain many copied files. Byte completion never fabricates publication
or root success. Empty files count when actually published; Skip and pre-publication cancellation
do not. A published child stays counted if subsequent proof or cleanup fails, while the root's
failure remains visible through the existing typed result model.

The Native provider and host bridge expose the optional logical-file observer described in
[the VFS contract](../Plugins/Plugins_VirtualFileSystem.md#file-and-completed-object-observation). Task-wide
copied-count availability requires support for every selected per-item scope, or the one declared
public Bulk scope. Lost retry identity withdraws certainty permanently. Aggregates update under
the existing row mutex using bounded active receipts and incremental scope/count/size totals;
regular rows remain bounded to sixteen, with separate bounded producer receipts. Late events
are guarded by scope, logical leaf ID and source identity and cannot retire a newer row.

Discovery still uses one bounded traversal and closes at listing completion, independently of
transfer drain. A complete byte total requires a provider's independent enumeration certificate
or every discovered file's independently summed known size after closure. Zero and unavailable
are distinct. Sequential bridge children never adopt the first child's size as the folder's
whole denominator; every child's bytes contribute to written progress. Snapshot publication is
ordered under the existing counter lock, and row/count/availability/size tuples are captured
coherently. Verification weighting and selected-root result reduction remain unchanged.

`DiscoveryScope_FolderLiveProgress` copies real nested files through serial, overwrite, parallel,
bridge and Verify routes and checks exact counts/bytes, independent destination content, actual
advancing meter pixels, delayed finalization and known-size Publishing. Separate controls cover
open traversal, empty roots and publication followed by committed-error/proof failure.
`DiscoveryScope_MultipleFoldersLiveProgress` covers real concurrent selected roots and zero-payload
file links. Its Native `CopyItems` API control uses the production reducer on a non-admitted task;
it does not bypass the product's Bulk Copy admission policy or claim a product Bulk popup route.

## Recursive-delete enumeration and retained parents (2026-10-16)

Local directory enumeration uses `Common::DirectoryRecords::Read` for full/id records before constructing child paths. Malformed records fail `ERROR_INVALID_DATA`. The same decoder serves the manifest, ordinary listing, watch notification, search-index and durable-store consumers with their own layout and error contracts.

Each permanent-delete traversal frame that owns its root retains a no-follow DELETE/read-attributes/write-attributes handle with `FILE_SHARE_READ` through enumeration, all child callbacks and final disposition. Ancestor frames remain alive until their descendants finish. The listing handle may close between batches; it is not deletion authority. Final disposition uses the retained handle. A bound contents-only root uses its caller's existing DELETE/share-read handle, avoiding a conflicting duplicate open. Failure to acquire a pin fails the operation; identity-less servers do not justify an unpinned fallback. Cancellation and errors join callbacks before releasing frame state. Pins require one handle per live depth/frame; `FileOps.DeleteTraversal.MaxDepth` reports observed depth. The permanent-delete stack retains its existing depth behavior; resource exhaustion fails the open and unwinds pins rather than falling back to unpinned deletion. SMB/NAS share-mode compatibility still requires provider qualification.

`FileOps_ReviewRemediationContracts` pauses after a committed child and listing closure, attempts descendant-parent rename and write opens, and checks success/cancelled handle release through 32 nested levels for serial/parallel and bound/unbound routes. The bound contents-only route remains serial; metrics distinguish configured from effective concurrency. It also injects unreachable/refused XP-09 prerequisites through the actual case result path. An unreachable prerequisite is `skipped` with a reason, a sandbox refusal is `failed`, and neither is a cross-root pass.

`FileOps_SmallFileCopyCost` validates an exact unique destination filename set and byte-for-byte source equality on both provider and bridge routes after timing stops. Missing and same-sized corrupt output negative controls must fail validation before performance evidence is emitted. A requested file count is never accepted as an observed result. The Dummy fixture provider must produce `alignof(FileInfo)`-aligned records accepted by the canonical packed cursor; mixed filename lengths are checked independently in PluginContractTests.
