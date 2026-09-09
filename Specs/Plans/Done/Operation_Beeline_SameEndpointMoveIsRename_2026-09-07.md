# Operation Beeline — Same-Endpoint Move Is a Rename

> **NON-NORMATIVE RETIRED WORKING-PLAN RECORD.** Moved from `Specs/Plans/WIP/` on
> 2026-09-08 at closeout. The same-endpoint Move rename, the per-child rename
> merge, the split of the link policy into a Copy-only option, and the removal of
> the `links.nativeMoveSemanticTransform` capability landed with tests, a Fresh
> Full gate, and archived evidence; that behavior is now owned by
> `Specs/FileSystem/FileSystem_FileOperations.md` and
> `Specs/Plugins/Plugins_VirtualFileSystem.md`. **No remaining implementation
> action belongs to this file.** The provider work it routed out and the
> residuals its gates recorded are owned by
> `FileOperations_BeelineResidualsAndProviderRoutes_2026-09-08.md`.

## Status

- **State:** DONE (2026-09-08; implemented and gated on the `c5a0a019` working tree, see History)
- **Priority:** P1 product simplification, owner instruction of 2026-09-07
  ("behavior closer to File Explorer in the simple scenarios; same disk move
  just moves; same remote server move just moves").
- **Planned at:** `c5a0a019aab2e7436fa4d68391aa30772f699fcc`
- **Ownership boundary:** same-endpoint Move strategy for every provider that
  advertises `operations.nativeMove` (Local fixed volumes and SMB shares, Curl
  FTP/SFTP/SCP, S3, Microsoft Drive, Google Drive, Dummy), the same-endpoint
  folder merge, the link policy as it applies to Move, and the
  `links.nativeMoveSemanticTransform` capability. Not owned: cross-endpoint
  Move (Managed and Copy-only stay as specified), Copy semantics and the Copy
  link policy, Verify, Delete/Recycle, inline Rename and Batch Rename, MTP, and
  the Shell Copy/Move experiment held in
  `Operation_FileOperations_LaterFeatures_2026-09-05.md`.
- **Drift check:**
  `git diff c5a0a019aab2e7436fa4d68391aa30772f699fcc..HEAD -- RedSalamander/FolderWindow.FileOperations.cpp RedSalamander/FolderWindow.FileOperationsInternal.h RedSalamander/FolderWindow.FileOperations.State.cpp RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp RedSalamander/FolderWindow.FileOperations.State.Private.h RedSalamander/HostServices.cpp Plugins/FileSystem Plugins/FileSystemCurl Common/PlugInterfaces/FileSystem.h Common/FileSystemRouteContract.cpp Common/FileSystemRouteContract.h Common/FileSystemRouteProviderBase.h Specs/FileSystem/FileSystem_FileOperations.md Specs/Plugins/Plugins_VirtualFileSystem.md Specs/UI/UI_FileOperationsPopup.md Specs/UI/UI_PreferencesDialog.md Specs/Testing/Testing_PerformanceValidation.md`
- **Authoritative specs at closeout:** `Specs/FileSystem/FileSystem_FileOperations.md`,
  `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/UI/UI_FileOperationsPopup.md`,
  `Specs/UI/UI_PreferencesDialog.md`, `Specs/Testing/Testing_PerformanceValidation.md`.

## Why

The providers already do what File Explorer does. The Local provider renames a
directory tree or a link object with one `MoveFileWithProgressW` call
(`Plugins/FileSystem/FileSystem.FileOps.cpp`, the Native `MoveItem` path), and
the Curl provider renames a file or a directory server-side whenever both paths
resolve to one connection (`Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp`,
`MoveItem` under `CanServerSideRename`). The host never asks them to, because of
three spec rules. What the user sees today, verified by reading the planned-at
tree:

| Scenario | File Explorer | RedSalamander at `c5a0a019` |
|---|---|---|
| Same volume, regular file | rename | rename (Native) |
| Same volume, directory to a new name | rename | full copy, verify, exact delete (Managed) |
| Same volume, directory onto an existing directory | per-child rename | full copy of every child, then delete (Managed merge) |
| Same SMB share, directory | rename | every byte streamed down and back up through the client bridge, then delete (Managed) |
| Same FTP/SFTP/SCP connection, file or directory | server rename | download, upload, source kept as `Copied; source kept` (Copy-only) |
| Same Google Drive, any item | parent change by item id | never Native by the host gate (typed facts carry no semantic-transform claim); confirm in Step 0 |
| Same SMB share, Copy with Verify off | server-side copy | server-side copy through provider `CopyFileExW` (unchanged) |
| Inline Rename (F2), any provider | rename | rename (unchanged) |

The three rules that produce the right-hand column:

1. **Native tree/link admission requires a link proof that exists only to honor
   a Skip-links policy on Move.** `FolderWindow.FileOperations.cpp` admits a
   non-local endpoint as Native only when the provider proves
   `links.nativeMoveSemanticTransform`, and splits a local selection so that a
   plain directory is Managed (`(nativeMoveSemanticTransform || localEndpoint)`
   gate and the top-level shape split). Local, Curl, and Google Drive advertise
   the claim as false; the spec rationale is that "an atomic Win32 tree/link
   rename cannot honor a captured Skip policy without enumerating".
2. **A Native call is one provider mutation and Managed is the sole folder-merge
   route.** The same-endpoint merge therefore copies every child through the
   host bridge and deletes the source, even when a per-child rename would do.
3. **Native admission is coupled to delete proofs through the strategy
   selector.** When rule 1 rejects Native, `SelectMoveStrategy` requires a bound
   and conditional source delete for Managed. Curl has neither, so a same-server
   Curl move degrades to Copy-only and retains the source.

The two spec documents already disagree about Curl: the File Operations provider
baseline table says "Copy-only until a safe native subset is advertised and
tested" while the provider spec's `curl-remote-path` row says same-provider
native Move is admitted. This plan reconciles both.

## Decisions

Accepted by owner instruction on 2026-09-07. Each row is a product decision, not
activation paperwork; a reopen trigger is stated so the decision is not
re-litigated by drift.

| ID | Decision | Rationale | Reopen trigger |
|---|---|---|---|
| BEE-D1 | **A Move never applies the link policy.** `Preserve links` / `Skip links` is a Copy option. Every Move relocates every link object literally: by native rename on one endpoint, by literal publication across endpoints. The Move confirmation no longer shows or captures a link choice; the Preferences default is documented as a Copy default. | A rename cannot skip anything without enumerating, and Explorer offers no link choice on move. Once Move always preserves, the "semantic transform" proof has nothing to prove. Literal preservation is already the only Move-safe behavior the engine implements. | A product request for per-link Skip on Move. The fallback design is policy-gated Native (Native only under Preserve), which keeps the shape probe; it is deliberately not chosen here. |
| BEE-D2 | **Native admission is endpoint plus capability only.** A same qualified endpoint (provider, instance, path profile, root) whose provider advertises `operations.move` and `operations.nativeMove` is Native for every item shape: regular file, directory tree, link object. No link proof, no delete proof, no pane-listing shape probe, no provisional split in Preparing. `links.nativeMoveSemanticTransform` is removed from the typed facts, JSON, parsers, and contract tests (ABI lockstep per `ATLAS-DEC-ABI-01`). | The claim was introduced by the retired P3.2c package solely to gate rule 1. With BEE-D1 it is dead weight in every provider, every fixture, and two source-shape contract tests. | A provider whose native Move provably does not relocate link objects as stored. None of the shipped providers is such a provider. |
| BEE-D3 | **Same-endpoint folder merge is a rename merge.** When a Native directory Move meets an existing regular directory, the host enumerates the source with the bounded walker it already owns, relocates each non-colliding child by one provider Native rename (a file, a subtree, or a link object, never entered), recurses into directory-onto-directory pairs, routes every other collision through the existing typed conflict surface, and removes each emptied source directory through the exact cleanup boundary. No stage, no byte transfer, no proof, no verification. Managed copy/publish/delete exists only across endpoints. | This is what Explorer does for a same-volume merge, and it removes the copy, proof, and cleanup triad from the most common directory move. The provider stays one mutation per call; only the host loop changes its leaf action from copy-and-delete to rename. | Evidence that a shipped provider's rename cannot relocate a subtree into an existing parent (none known: Win32 `MoveFileEx`, FTP `RNFR`/`RNTO`, SFTP rename, Graph `PATCH`, Drive `files.update`, and the S3 per-object route all can). |
| BEE-D4 | **Unchanged.** Cross-endpoint Move remains Managed with every proof or Copy-only with the source retained. A same-endpoint pair whose provider has `move` but not `nativeMove` (MTP today) keeps its current strategy. Copy semantics are unchanged; Verify-off same-endpoint Local Copy keeps the provider `CopyItem` route so SMB server-side copy and block-clone offloads keep applying, and Verify-on keeps the host bridge cutover. | The user asked for Explorer behavior in the simple scenarios; cross-volume Explorer also copies then deletes, and the proofs there are cheap and protect data. | None. This row records scope, not a choice. |

## Target behavior

The following is the normative text to land at closeout. Line numbers refer to
the planned-at commit; re-anchor after drift.

### `Specs/FileSystem/FileSystem_FileOperations.md`

**Product invariants.** Replace invariant 4 and append invariant 13:

> 4. Move is native Move (one provider rename per selected object, or a per-child
>    rename merge when a directory meets an existing directory on the same
>    endpoint), managed copy-then-delete across endpoints, or Copy-only with the
>    source retained.
>
> 13. A Move never applies the link policy. `Preserve links` / `Skip links` is a
>     Copy option; every Move relocates every link object literally, by native
>     rename on one endpoint or by literal publication across endpoints.

Invariant 5 gains one sentence: "A rename merge removes an emptied source
directory only through the same exact cleanup boundary, or, on a profile without
object binding, through the provider's own empty-directory removal that fails
source-kept on `ERROR_DIR_NOT_EMPTY`."

**Admission and confirmation** (captured options, line 271): replace
"- Links: `Preserve links` or `Skip links`;" with "- for Copy only, Links:
`Preserve links` or `Skip links`; a Move confirmation shows and captures no link
policy;".

**Capability and strategy routing**, paragraph "Capability v2 separates
`operations.move` from `operations.nativeMove`" (lines 476–484): replace the
Local sentence with:

> Local advertises `nativeMove: true` because its provider Move entry point
> performs one native rename only, for a regular file, a directory tree, or a
> link object alike; folder merge is the host rename merge below, and the
> provider contains no copy/delete or per-child fallback of its own.

Paragraph "Local `MoveItem`/`MoveItems` never perform cross-volume, folder-merge,
or semantic-copy fallback" (lines 497–503): replace the last sentence with
"Managed Move is the sole Local copy/publication/delete route and exists only
across endpoints; folder merge on one endpoint is the rename merge."

Strategy table (lines 505–510): rewrite the Native row's admission column to
"Source and destination share one qualified endpoint (provider, instance, path
profile, root) whose provider advertises `operations.move` and
`operations.nativeMove`; the host requests `FILESYSTEM_MOVE_NATIVE_ONLY`. No link
proof and no delete proof is required." Add the row:

> | Rename merge | A Native directory Move whose destination is an existing regular directory on the same endpoint | Each non-colliding child relocated by one provider Native rename; collisions through the typed conflict surface; emptied source directories removed by exact cleanup; no bytes transferred, no stage, no proof |

Replace lines 511–535 (from "The completed engine uses native Local Move only for
a path-scoped same-volume regular-file item." through "No recursive preflight is
introduced.") with:

> Native admission does not depend on item shape. A regular file, a directory
> tree, or a link object on one qualified endpoint is Native whenever the
> provider advertises `operations.nativeMove`; the pane listing, admission, and
> Preparing perform no top-level shape probe, no provisional split, and no
> recursive preflight. A rename relocates every descendant and every link object
> as stored, which is exactly the Move link contract, so no separate link proof
> exists. The Native call remains one provider mutation: a regular directory
> that meets an existing regular directory returns the provider's known
> destination-exists non-commit (`ERROR_ALREADY_EXISTS` or `ERROR_FILE_EXISTS`)
> before any child is touched, and the same task continues that item as a
> rename merge without a second task, a repeated confirmation, consumed
> clipboard state, or a replayed Native call. Every other Native non-commit (a
> destination link, a type mismatch, a read-only or sharing failure) enters the
> ordinary typed conflict surface; a native sharing failure never falls through
> to pathname deletion or to a copy. Local `rootId` remains volume-specific, so a
> cross-volume pair is never Native and uses Managed or Copy-only exactly as
> before.
>
> Rename merge. The host enumerates the source directory with the bounded walker
> it already uses for Managed Move, in provider order, and for each child: an
> absent destination name is one provider Native rename of that child (a file, a
> subtree, or a link object, never entered); a regular directory on both sides
> recurses; every other collision (file, link, type mismatch, read-only) enters
> the typed conflict surface, where Overwrite is the provider's exact
> `RenameIfUnchanged` replace route already used by Native file Move, Keep Both
> renames to the collision-safe sibling name, and Skip retains the child. After
> its children, an emptied source directory is removed through the exact cleanup
> boundary under Publication and managed Move; `ERROR_DIR_NOT_EMPTY` retains it
> and its ancestors as `Moved; source folder kept` without a prompt, a recopy, or
> a second delete. The merge transfers no bytes, creates no stage, computes no
> proof, reports verification `NotApplicable`, and counts items for progress. A
> profile without object binding removes the emptied directory through its own
> delete route by name, the authority its Permanent Delete already has.

Provider baseline table (lines 562–574): Local Move column becomes "Native for
every shape on one volume; rename merge onto an existing directory; Managed
cross-volume only with exact binding/publication/delete". Curl FTP/SFTP/SCP Copy
column becomes "Same-connection stream copy; bridge export/import Copy" and its
Move column "Native same-connection rename for files and directories; rename
merge onto an existing directory; Copy-only across connections". The Microsoft
Drive, S3, Google Drive, and Dummy rows keep their Move text.

**Publication and managed Move**: after "For every copied regular item:" and its
eight steps, add "A rename merge performs none of these steps. The directory
cleanup rules below apply to it unchanged."

**Link policy: literal Preserve**, first paragraph: append "The policy applies to
Copy. A Move never applies it: a Move relocates every link object literally, by
native rename on one endpoint or by literal publication across endpoints, and
the Move confirmation shows no link choice." Replace the two paragraphs at lines
796–815 ("Native provider Move may substitute for a tree or link object only
when…" and "The Local Native Move shape decision…") with:

> A native provider Move relocates a tree or a link object as stored. Because a
> Move never applies the link policy, no link proof gates Native admission and
> `nativeMove: true` is sufficient. The pane listing, admission, and Preparing
> make no shape decision for Native Move.

**Verification and metadata**, first paragraph: append "Verify Off keeps the
provider `CopyItem` route for same-endpoint Local Copy, so SMB server-side copy
and block-clone offloads stay available."

**Settings** (line 1766): "Provider configuration may persist only the
Preserve/Skip link default, which applies to Copy; Follow is removed."

### `Specs/Plugins/Plugins_VirtualFileSystem.md`

- Remove `"nativeMoveSemanticTransform": false` from the capability JSON sample
  (line 1501) and the definition bullet (lines 1548–1551). Replace the bullet
  with: "`operations.nativeMove: true` admits Native Move for every item shape on
  one qualified endpoint. There is no separate link proof because a Move never
  applies the link policy."
- Line 1136: "`FileSystemOptions` includes `FileSystemLinkPolicy::{Preserve,
  Skip}`, consumed by Copy; a Move always receives Preserve."
- Curl SFTP/SCP and FTP rows (lines 1394–1395): replace "same-provider
  Copy/Move/Rename are disabled because…" with "same-provider Copy, native Move
  (files and directories through one server rename), Rename, and Delete are
  admitted; rename overwrite has no atomic no-replace primitive, so overwrite is
  the rollback-sibling sequence, and a directory-onto-directory destination is a
  known `ERROR_FILE_EXISTS` non-commit that the host continues as a rename
  merge."
- Provider rows (Local 1585, Dummy 1587, Curl 1589, Microsoft Drive 1592, S3
  1595, Google Drive): delete every `nativeMoveSemanticTransform` phrase. Local
  becomes "Same-endpoint Move of a file, tree, or link object is Native;
  directory-onto-directory is the host rename merge; cross-volume uses
  Managed." Curl's "no Move bridge" stays.

### `Specs/UI/UI_FileOperationsPopup.md`

- Line 146: "- for Copy, `Links`, with exactly `Preserve links` and `Skip links`;
  Follow is never displayed or accepted; a Move confirmation does not show the
  row;".
- Add the result wording `Moved; source folder kept` for a rename merge whose
  source directory retained a skipped or conflicting child, alongside the
  existing `Copied; source kept`.

### `Specs/UI/UI_PreferencesDialog.md`

- Lines 132–133: the `Reparse points (symlinks/junctions)` setting is "a
  provider-owned default for Copy".

### `Specs/Testing/Testing_PerformanceValidation.md`

- Lines 889–897: `fileops.plan.admit_us` performs no shape probe and acquires no
  `IFileSystemIO` for shape. A destination directory that appears after
  admission is first attempted through the provider's one-mutation Native
  boundary; its known non-commit emits
  `fileops.execute.native_directory_race_requalified_after_noncommit` and
  continues as a rename merge. `fileops.execute.native_directory_race_fallback_unavailable`
  is removed; there is no missing-fallback outcome.
- New correctness/strategy rows: `FileOps.SelfTest.SameVolumeTreeMoveIsRename`
  (bridge byte counters zero, `FILE_ID_INFO` of every descendant unchanged),
  `FileOps.SelfTest.SameVolumeRenameMerge` (same, plus one file conflict and one
  skipped child retaining the ancestor as `Moved; source folder kept`),
  `FileOps.SelfTest.CurlSameConnectionTreeMoveIsRename` (fixture-gated).
- New throughput row: same-volume directory Move of a 10,000-file tree,
  baseline Managed copy at `c5a0a019` versus candidate rename, archived under
  `Specs/TestRuns/<commit>/FileOps/` with `fileops.operation.strategy` showing
  `move.native.same-root`.

### ABI

`Common/PlugInterfaces/FileSystem.h` drops `BOOL nativeMoveSemanticTransform`
from the typed route facts. Current source-tree lockstep (`ATLAS-DEC-ABI-01`)
permits the atomic layout change; every producer and consumer is in this tree.
If the owner prefers a smaller diff, the alternative is to keep the field as
reserved-must-be-false and drop only its JSON and admission use; the plan
recommends removal.

## Implementation steps

Order matters: Step 1 makes directories Native, Step 2 makes the merge safe
without a copy, Step 3 removes the policy from Move, Step 4 deletes the dead
capability. Build and run the File Operations self-test after each step.

### Step 0 — Confirm two readings before coding

- Google Drive typed facts (`FileSystemGoogleDrive.cpp`, the descriptor around
  line 2918) carry no semantic-transform claim, so a same-drive Move should show
  as Copy-only or Managed in `fileops.operation.strategy` today. Confirm with the
  Google Drive fixture or a debug run and record the observed bucket here. If it
  is Native by some path this plan missed, correct the Why table.
- Curl `MoveItem` against the C11 FTP fixture: a directory to a new name on the
  same connection succeeds through one `RNFR`/`RNTO`; a directory onto an existing
  directory returns `ERROR_FILE_EXISTS` from `PrepareOverwriteTargetForRename`
  before any mutation. Record the server used.

### Step 1 — Host admission (`RedSalamander/FolderWindow.FileOperations.cpp`, `FolderWindow.FileOperationsInternal.h`)

- Delete `LocalNativeMoveItemShape`, `QualifyLocalNativeMoveItemShape` (lines
  96–238), `LocalNativeMoveItemShapeFacts` and
  `QualifyLocalNativeMoveItemShapeForPreparing` (lines 1095–1122), and the
  per-item split loop (lines 2412–2450).
- `providerNativeMoveQualified` becomes `providerNativeMoveAdvertised`
  (lines 2410–2411 and 2468–2469); `nativeMoveQualified` follows. Remove the
  `fileops.plan.native_semantic_route_reclassified` counter.
- Remove `nativeShapeProvisional`, `localNativeShapeQualified`,
  `containsNativeDirectoryRaceCandidate`, and `nativeDirectoryRaceFallback`
  from `SourcePartition` and `TransferPlan`, and their `ValidatePlan` rules
  (lines 312–325). The rename merge needs no prepared fallback plan: it is
  admitted by the same facts as the Native plan it continues.
- `SelectMoveStrategy` is unchanged; `MoveStrategyQualificationFacts.nativeMoveQualified`
  now means "same endpoint and `nativeMove` advertised".
- `FolderWindow.FileOperations.State.cpp` lines 8740–8800: delete the Preparing
  re-partition of provisional plans.

### Step 2 — Rename merge (`RedSalamander/FolderWindow.FileOperations.State.cpp`)

- The Native item executor (around line 18795) keeps `FILESYSTEM_MOVE_NATIVE_ONLY`.
  When the provider returns `ERROR_ALREADY_EXISTS` or `ERROR_FILE_EXISTS` for a
  directory source whose destination is a regular directory (no-follow bind,
  `FILESYSTEM_BOUND_DIRECTORY`), emit
  `fileops.execute.native_directory_race_requalified_after_noncommit` and run the
  rename merge for that item on the same task thread. Replace the fallback
  selection at lines 18630–18640 and the Managed test at line 17259.
- The merge reuses the sequential directory walk (around lines 15790–16035) and
  the parallel child queue (around line 16660) with a new leaf action: for an
  absent destination child call `_fileSystem->MoveItem(child, destinationChild,
  itemFlags, options with NATIVE_ONLY)`; a subtree or a link object is one call
  and is never entered. Directory-onto-directory recurses; every other collision
  goes to the existing conflict arbiter with the existing Overwrite / Keep Both /
  Skip / Cancel receipts. Overwrite reuses the provider's `RenameIfUnchanged`
  route (`FileSystem.FileOps.cpp`, the `destinationExistedBeforeCopy` branch of
  Native `MoveItem`).
- Post-order cleanup of the emptied source directory reuses
  `FinalizeManagedSourceCleanup` / `DeleteIfUnchanged`; a retained child keeps
  the ancestry exactly as `hadRetainedChild` does today, with the new result
  wording `Moved; source folder kept`. For a profile without object binding
  (Curl) call the provider's directory delete by name and treat
  `ERROR_DIR_NOT_EMPTY` as retained.
- The item reports strategy `native`, verification `NotApplicable`, zero bridge
  bytes, and item-counted progress. The `CrossFileSystemBridge` reader/writer
  path is not constructed for a same-endpoint Move.
- The executor no longer refuses a Native item because the task carries an
  explicit destination provider object: the endpoint tuple decides, and the
  source provider executes the rename and the rename merge.

### Step 3 — Move ignores the link policy

- `Task::InitializeFileSystemOptions` (`State.cpp` around line 9855) sets
  `FILESYSTEM_LINK_PRESERVE` for every Move plan; Copy plans keep the captured
  policy.
- `FolderWindow.FileOperations.cpp` lines 2253–2257 read the plugin default for
  Copy only; a Move plan's `options.linkPolicy` is Preserve.
- The Copy/Move confirmation (`State.Runtime.cpp` lines 943 and 992,
  `HostServices.cpp` prompt validation, the popup) hides the Links row for a Move
  and ignores `HostFileOperationPromptOptions::linkPolicy` on a Move; the field
  stays in the host ABI for Copy.
- While in these branches, align the Copy Skip test with Preserve: the bridge
  Skip branches at `State.cpp` lines 16133, 16150, 17000, and 17028 key on the
  raw `FILE_ATTRIBUTE_REPARSE_POINT` bit, while `CopyLink` (around line 15687)
  classifies through `BindObject` and copies cloud, WOF, and dedup placeholders
  as regular files. Skip must skip name-surrogate links only. Add
  `FileOps.SelfTest.CopySkipLinksKeepsPlaceholders` (a WOF-compressed file inside
  a copied tree under Skip is copied, not skipped).

### Step 4 — Remove `nativeMoveSemanticTransform`

Delete the field and every producer, parser, fixture, and assertion:
`Common/PlugInterfaces/FileSystem.h:453`, `Common/FileSystemRouteProviderBase.h`
(lines 56 and 154), `Common/FileSystemRouteContract.{h,cpp}` (lines 58, 162,
375), `RedSalamander/FolderWindow.FileOperations.cpp` (lines 720, 870),
`FolderWindow.FileOperationsInternal.h:397`, `State.Private.h:170`,
`State.cpp:8740`; provider descriptors in `FileSystemDummy.cpp:5512`,
`FileSystemS3.Core.cpp:641`, `FileSystemMicrosoftDrive.cpp:6817`; every
capability JSON string (`FileSystem.cpp:5991`, `FileSystemCurl.Shared.cpp`
lines 5186 and 5234, `FileSystem7z.h:227`, `FileSystemS3.h` lines 360 and 406,
`FileSystemDummy.h:310`, `FileSystemGoogleDrive.cpp:212`,
`FileSystemMtp.Core.cpp:4248`); test fixtures (`ViewerSqliteTests.cpp:90`,
`ViewerPETests.cpp:119`, `FolderViewRefreshDuplicatePathPerfTest.cpp:54`,
`PluginContractTests.cpp` lines 218, 235, 2515, 2536); self-test snapshots
(`FolderWindow.FileOperations.SelfTest.cpp` lines 2723, 2939, 3167, 3225,
`Commands.SelfTest.BatchRename.cpp:916`); and `docs/dev/PluginHostModel.md` if it
names the claim.

### Step 5 — Curl

No provider code change is expected. Update the capability JSON and the two
provider-spec rows. The provider-side proof already exists: the Move preflight
truth export (`RedSalamanderCurlMovePreflightTruthForSelfTest`, run by the
`C0_CurlMovePreflightTruth` case) moves a directory with a nested subtree on one
connection through exactly one `RNFR`/`RNTO` pair and no `RETR`/`STOR`, and
rejects another connection with a known non-commit. The host side is the same
admission rule as Local, so no separate host case is needed.

### Step 6 — Tests

Update or retire the assertions that pin the old rules:

- `Commands.SelfTest.PluginConfig.cpp:5657` (source-shape contract for the
  `(nativeMoveSemanticTransform || localEndpoint)` gate): delete. This is in the
  direction Operation Astrolabe already owns.
- `FolderWindow.FileOperations.SelfTest.Phases05_06.cpp` capability matrix
  (lines 1040–1052, 4017–4029, 4185): drop the transform expectations.
- `Phase12_ReparsePointPolicy`: keep for Copy; add the Move half that proves the
  policy is ignored and links are relocated.
- `FileOps_ReparseDirectoryMergeIntoExistingFolder`,
  `Riptide_ReparseMoveSourceNeverUsesDirectoryMerge`,
  `Riptide_ReparseNativeMoveRelocatesLinkObject`: re-express against the rename
  merge; the link-object expectations stay true.
- `FileOps.SelfTest.FairstreamMoveWindowPreserved` (new file appears in the
  source during a folder Move): the destination-absent shape can no longer race
  because the tree is renamed; keep the merge shape, where the late file makes
  the exact cleanup return `ERROR_DIR_NOT_EMPTY` and the result is `Moved;
  source folder kept`.
- New: `FileOps.SelfTest.SameVolumeTreeMoveIsRename`
  (`Beeline_SameVolumeTreeMoveIsRename`), `FileOps.SelfTest.SameVolumeRenameMerge`
  (`Beeline_RenameMergeSkipKeepsSourceFolder`),
  `FileOps.SelfTest.NativeDirectoryRaceContinuesAsRenameMerge` (inside the
  move-merge qualification and provider-matrix cases), and
  `FileOps.SelfTest.CopySkipLinksKeepsPlaceholders`
  (`Beeline_CopySkipLinksKeepsPlaceholders`, skipped with a reason on a volume
  without the WOF driver). Every new case joins `kFileOpsFamilyDefinitions` in
  `FolderWindow.FileOperations.SelfTest.cpp`, or the Full suite skips it silently.
- `Beeline_LoopbackShareMoveIsRename` (`FileOps.SelfTest.LoopbackShareMoveIsRename`):
  a Move whose source and destination both live on the loopback
  administrative-share alias of the sandbox proves the SMB profile takes the
  same Native route; an unreachable share is a logged skip.
- `Phase10_MetadataPreservationAndSourceRetention` and the Move sub-steps of
  `Phase10_ContentVerification` claim Managed-route behavior (metadata transfer,
  deferred consent, verification before the exact source delete). They target
  the loopback administrative-share alias (`local-win32-smb`), the one second
  qualified endpoint a single-volume machine has, and skip with a log line when
  it is unreachable. `Phase10_TypedResultsAndConsumers` keeps its Local
  destination and becomes a rename-merge proof.
- `FileOps_ResolvedItemsExactDestinations`: a resolved `DirectoryShell` records a
  host receipt (destination folder published, source untouched) instead of
  tripping the Native no-receipt rule; the summary counts it as neither a
  removed nor a kept source (`Specs/Core/Core_CompareDirectories.md`).

### Step 7 — Performance evidence

Follow `.github/skills/perf-validation/SKILL.md`. Scenario: same-volume Move of a
10,000-file, 1 GiB tree to a sibling name, and the same tree onto an existing
copy of itself with 10 % file conflicts answered Skip. Record baseline at
`c5a0a019` and candidate with `fileops.operation.strategy`, task wall time, bridge
byte counters, and peak retained bytes. Archive under
`Specs/TestRuns/<commit>/FileOps/` with the metric dictionary required by
`Specs/Testing/Testing_PerformanceValidation.md`.

### Step 8 — Closeout

Apply the Target behavior text to the five authoritative documents, record the
reversal of the P3.2c rationale in this file's history section, move this plan to
`../Done/`, and update the WIP index. Run
`.\Tools\Get-SpecInventory.ps1 -FailOnFindings`.

## Files in scope

- `RedSalamander/FolderWindow.FileOperations.cpp`, `FolderWindow.FileOperationsInternal.h`
- `RedSalamander/FolderWindow.FileOperations.State.cpp`, `State.Runtime.cpp`, `State.Private.h`
- `RedSalamander/HostServices.cpp` (prompt options for Move)
- `Plugins/FileSystem/FileSystem.cpp` (capability JSON), `FileSystem.FileOps.cpp` (no route change expected)
- `Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp` (capability JSON)
- `Common/PlugInterfaces/FileSystem.h`, `Common/FileSystemRouteContract.*`, `Common/FileSystemRouteProviderBase.h`
- Every provider and test fixture listed in Step 4
- `RedSalamander/SelfTest/FileOperations/*`, `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`
- The five authoritative specs named under Status

## Verification

```powershell
.\build.ps1
.\.build\x64\Debug\RedSalamander.exe --fileops-selftest
.\Tools\Run-AllTests.ps1 -Suite Full
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

The Full suite serializes File Operations runs across worktrees through the
session-global mutex; a concurrent run exits with code 3 and must be re-run, not
reported as a failure.

## Done criteria

- A same-volume Move of a directory tree, a junction, or a symlink is one
  provider rename: zero bridge bytes, unchanged `FILE_ID_INFO` for every
  descendant, `move.native.same-root` in `fileops.operation.strategy`.
- A same-volume directory onto an existing directory relocates every
  non-colliding child by rename, prompts only for colliding children, and
  removes the emptied source directory; a skipped child yields
  `Moved; source folder kept` with no copy.
- A same-connection Curl directory Move is one server rename and leaves no
  source (the existing Move preflight truth export, `C0_CurlMovePreflightTruth`).
- The Move confirmation shows no link choice; a Move with links relocates them
  regardless of the Preferences default; Copy still honors Preserve/Skip, and
  Skip no longer skips placeholder files.
- `nativeMoveSemanticTransform` no longer exists in headers, providers, JSON,
  fixtures, self-tests, or specs.
- The cross-volume Local Move boundary test, the Fairstream literal-link tests,
  and every Managed cross-endpoint test still pass unchanged.
- Fresh Full gate green on the merged tree; performance evidence archived; the
  five specs updated; this plan moved to `../Done/`.

## STOP conditions

- A shipped provider's native rename is found not to relocate a subtree or a
  link object as stored. Stop, record the provider, and reopen BEE-D3 for that
  profile only.
- The rename merge cannot keep invariant 8 (each child publishes independently;
  no child is presented as moved while it is not) under cancel or crash. Stop
  and redesign the leaf action before removing the Managed same-endpoint path.
- The Curl fixture server refuses `RNTO` on a non-empty directory. Record the
  server, keep Curl directory Move Native only where `CanServerSideRename` and
  the fixture prove it, and do not fall back to the bridge silently.
- The owner asks to keep `Skip links` on Move. Stop; the fallback is policy-gated
  Native admission, which is a different plan.
- Fresh Full is not green after the merge. Do not close; classify residuals
  against the pre-existing failure set recorded in `Specs/TestRuns/README.md`.

## Accepted consequences

These are stated, not repaired, exactly as the existing literal-Preserve section
states its consequences:

- An absolute link inside a renamed tree keeps naming the old location and
  dangles until the user retargets it. Explorer behaves the same.
- A renamed tree keeps its source security descriptor; inherited ACEs are not
  recomputed for the new parent. Win32 rename and Explorer behave the same, and
  the former Managed copy applied destination inheritance instead.
- A tree containing an open file may fail the rename with a sharing violation and
  enters the typed conflict surface; there is no copy fallback. A single open
  regular file still renames where a copy would have failed.
- A cross-endpoint Move of a link into a provider that cannot preserve links
  retains that link at the source, as Preserve does today; Skip is no longer
  available to discard it during a Move.

## Out of scope, routed elsewhere

- Microsoft Drive same-drive Copy streams through the client although Graph has
  a server-side `copy` action; the provider baseline table already claims
  "server-side copy". A separate small plan should implement it. Not this one.
- MTP `nativeMove: false` stays until its tested WPD route is isolated.
- The Shell Copy/Move experiment stays on HOLD under
  `Operation_FileOperations_LaterFeatures_2026-09-05.md`; it covers single files
  only and is not a dependency of this plan.

## History

- 2026-09-07: plan written from a code and spec reading at `c5a0a019` after the
  owner asked why a same-disk move does not use the Windows rename. The P3.2c
  package in `../Done/Operation_FileOperations_GlobalBehaviorDecisionReview_2026-08-16.md`
  introduced the semantic-transform gate to honor a Skip-links policy on Move;
  BEE-D1 removes that policy from Move, which is why the gate can go.
- 2026-09-07/08: implemented spec-first on the `c5a0a019` working tree. Two late
  findings from the first complete File Operations run: a resolved
  `DirectoryShell` had no receipt and tripped the Native no-receipt rule (host
  receipt added, summary counts shells as neither removed nor kept), and the
  task-level `strategy.nativeBridgeRejected` rejection plus the executor guard
  refused a Native item whenever the task carried an explicit destination
  provider object (both deleted; the endpoint tuple decides). Managed-route
  proofs that need two Local endpoints on one machine now target the loopback
  administrative-share alias, and `Beeline_LoopbackShareMoveIsRename` proves the
  SMB profile takes the Native route. The Local provider no longer reports
  MOTW/ADS/EA as unsupported when `BackupRead` inspection fails; it returns the
  failure (`fileops.local.metadata.inspect_failed`). The first Full gate
  (2160 / 2096 / 6 / 58) aborted the family that carries the Beeline evidence
  rows through a transient `Phase7_SharedPerItemScheduler` cancel stall, so a
  second gate was run. Fresh Full gate on the finished package (governed runner `20260907T205843Z-46028-f4bfabc0e38e458cad88a4d92654fcac`, Debug x64, build receipt `0fabf2c625e565828cf8905b54aaaea1e935a11c710454f4d0984650761ce11f`): 2160 total / 2096 passed / 3 failed / 61 skipped in 78m47s. Residuals: `R4A19_DiscoveryIndependentVolumes` (environment gate, pre-existing), `Phase7_ParallelCopyMoveKnobs` (parallel Copy in-flight observation; passes in isolation; the preceding gate showed the sibling `Phase7_SharedPerItemScheduler` stall instead, so the Local parallel-Copy concurrency path is recorded as an open follow-up outside this plan), `cmd_pane_navigationView_full_path_popup_edit_route` (UI timing artifact; passes alone). Every Beeline case passed with its evidence rows. Evidence: `Specs/TestRuns/4cb089111a23/FileOps/20260907_205843_beeline_same_endpoint_move_rename/`.
