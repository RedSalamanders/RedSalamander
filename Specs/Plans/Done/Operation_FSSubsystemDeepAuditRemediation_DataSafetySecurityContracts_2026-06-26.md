# Operation FS Subsystem Deep-Audit Remediation - Data Safety, Security, and Contracts - 2026-06-26

## Status

- State: Done as umbrella closeout. Track A and Track C were implemented in this slice; remaining Track B through Track I work was split into explicit WIP follow-up plans.
- Reviewed and narrowed at: `aa89c6013` on 2026-06-26.
- Original audit docs committed at: `aa89c6013`.
- Completed overlapping implementation: `1d1cccc11` (`Finish FS bridge search index data safety`), documented in `Specs/Plans/Done/Operation_FileSystemBridgeSearchIndex_DataSafetyStudyAndFix_2026-06-25.md`.
- This plan now tracks only remaining file-system subsystem audit work. Completed overlap from the older FSBSI plan has been removed from the active worklist.

## Closeout - 2026-06-26

Implemented in this slice:

- Track A same-FS copy overwrite transaction:
  - Built-in local regular-file copy overwrite now stages into a unique sibling temp, verifies the temp byte count against the source, and promotes only after the copy is complete.
  - Built-in local copy-reparse overwrite also stages and promotes the completed link object instead of copying over the live destination.
  - Injected promotion failure preserves original destination bytes/object and readonly attributes, and best-effort temp cleanup removes staged files.
  - Regression: `Floodgate_LocalCopyOverwriteIsStaged`.
- Track C conflict consent bucket isolation:
  - `ERROR_DIR_NOT_EMPTY` maps to `NonEmptyDirectory`, and `ERROR_REPARSE_POINT_ENCOUNTERED` maps to `ReparsePoint` for copy/move conflicts.
  - Prompt text now distinguishes non-empty directory and reparse/link replacement risk classes.
  - Apply-to-all cache entries are separated by the new destructive-risk buckets, so a plain file Overwrite does not authorize a later non-empty directory/reparse replacement.
  - Regressions: `Fairstream_ReparseReplaceNonEmptyDirRequiresConsent`, `Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate`, and adjacent `Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent`.

Durable spec updates:

- `Specs/FileSystem/FileSystem_FileOperations.md` now documents same-FS copy overwrite staging, copy-reparse promotion, readonly failure preservation, and destructive-risk bucket scoping for apply-to-all.

Follow-up WIP splits:

- Track B: `Specs/Plans/WIP/Operation_FSDeepAudit_TrackB_CrossFsDirectoryMoveDeleteSafety_2026-06-26.md`
- Track D: `Specs/Plans/WIP/Operation_FSDeepAudit_TrackD_SearchServiceTrustBoundaryPipeProtocol_2026-06-26.md`
- Track E: `Specs/Plans/WIP/Operation_FSDeepAudit_TrackE_ReparseCanonicalTraversalSafety_2026-06-26.md`
- Track F: `Specs/Plans/WIP/Operation_FSDeepAudit_TrackF_ProviderTransferContracts_2026-06-26.md`
- Track G: `Specs/Plans/WIP/Operation_FSDeepAudit_TrackG_SearchIndexResidualCorrectness_2026-06-26.md`
- Track H: `Specs/Plans/WIP/Operation_FSDeepAudit_TrackH_FileOpsObservabilityPartialResults_2026-06-26.md`
- Track I: `Specs/Plans/WIP/Operation_FSDeepAudit_TrackI_ThreadingSecretsPluginQuietPoints_2026-06-26.md`
- Validation follow-up: `Specs/Plans/WIP/Operation_FileOpsFullAggregateSelfTestExitMinusOne_2026-06-26.md`

Validation evidence archived:

- Build: `.\build.ps1 -ProjectName RedSalamander` passed with 0 warnings / 0 errors; log `.build/logs/msbuild-20260626_161927_056.log`. Final FileSystem plugin rebuild after reparse-temp cleanup also passed with 0 warnings / 0 errors; log `.build/logs/msbuild-20260626_162649_524.log`.
- Track A focused guard: `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_162751/` (`Floodgate_LocalCopyOverwriteIsStaged`, 3 passed).
- Track C focused guards:
  - `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_144814/` (`Fairstream_ReparseReplaceNonEmptyDirRequiresConsent`, 3 passed).
  - `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_144834/` (`Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate`, 3 passed).
  - `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_144856/` (`Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent`, 3 passed).
- Family validation:
  - `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_145255/` (`FileOpsFamily_Fairstream`, 38 passed).
  - `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_162849/` (`FileOpsFamily_Fairstream`, 38 passed after the final FileSystem plugin rebuild).
- Shared selftest helper validation after reducing `WritePatternFileFsIo`'s scratch buffer from 256 KiB to 32 KiB:
  - `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_162424/` (`FileOpsFamily_Phase08_Validation`, 7 passed).

Full aggregate validation note:

- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild` was attempted during closeout. One retry first hit exit code `3` because another stale selftest process from `C:\Users\PVB\.codex\worktrees\cac5\RedSalamander` was holding the selftest mutex (`--compare-selftest --selftest-case=mtp_`); it exited before `Stop-Process` found it.
- With the mutex clear, full FileOps aggregate exited `-1` after about 39 seconds without `fileops_results.json`. The last trace line was `NextStep: Fairstream_SaturationConcurrentCopiesMakeProgress`; no new WER event or crash dump was created after the earlier fixed Phase8 stack crash. The isolated Fairstream family passed immediately afterward at `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_161711/`, passed after the final app build at `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_162522/`, and passed after the final FileSystem plugin rebuild at `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_162849/`.
- This aggregate validation anomaly is split to `Specs/Plans/WIP/Operation_FileOpsFullAggregateSelfTestExitMinusOne_2026-06-26.md`; the Track A/Track C guards and their family validation are green.

## Review Sources

- `Specs/Reviews/FS-Subsystem-Audit.md`
- `Specs/Reviews/FS-Subsystem-Findings.md`
- `Specs/Reviews/_AuditProgress.md`
- `Specs/Plans/Done/Operation_FileSystemBridgeSearchIndex_DataSafetyStudyAndFix_2026-06-25.md`

## Intent

The file-system subsystem audit found data-loss, corruption, security, incomplete-result, and contract gaps across file operations, virtual providers, search/index, and host services. The first remediation pass fixed the overlapping bridge/search/SQLite issues from the earlier FSBSI plan. This umbrella plan remains open for the broader audit findings that are still unresolved.

The priority order remains:

1. Prevent data loss and silent corruption.
2. Make destructive operations transactional or explicitly partial.
3. Make incomplete results visible to callers and users.
4. Close service and provider trust-boundary holes.
5. Encode provider contracts so plugins cannot reintroduce unsafe behavior.

## Completed or De-duplicated - Do Not Reimplement

These items were part of the original umbrella text but are now closed by the FSBSI plan. Do not add a second implementation for them here.

| Area | Current disposition | Proof / test names |
| --- | --- | --- |
| Old Track 0, reconcile the 2026-06-25 WIP | Done. The older plan was completed and moved to `Specs/Plans/Done/`. | Done plan closeout under `Specs/Plans/Done/Operation_FileSystemBridgeSearchIndex_DataSafetyStudyAndFix_2026-06-25.md`. |
| Cross-FS single-file COPY/MOVE with unknown source size | Done for the generic bridge. COPY now refuses unverifiable unknown-size sources; MOVE preserves the source. | `Floodgate_CrossFsCopyGetSizeFailureRefusesCommit`, `Floodgate_CrossFsMoveGetSizeFailurePreservesSource`; spec lines in `Specs/FileSystem/FileSystem_FileOperations.md`. |
| `IFileReader::Read` EOF zero-byte contract | Clarified in `Common/PlugInterfaces/FileSystem.h` for the generic bridge contract. Broader shared `ReadFully` helper/provider body verification remains open under Track F. | Header comments and FSBSI closeout. |
| Built-in local `CreateFileWriter(ALLOW_OVERWRITE)` staging | Done. The local writer stages to a sibling temp and promotes on `Commit()`. Abort deletes only the temp. | `Floodgate_LocalWriterOverwriteIsStaged`; `Win32FileWriter` staged mode in `Plugins/FileSystem/FileSystem.cpp`; file-operation spec update. |
| Search malformed materialization skip handling | Done. Native local search sets `FILESYSTEM_SEARCH_WARNING_OVERFLOW` and continues past malformed materialization failures instead of abandoning a parallel chunk. | `search_low_hardening_smoke`, `local_search_scan_wide_tree_parallel_walk_name_only`; `Specs/Core/Core_Search.md`. |
| SQLite legacy auto-vacuum queueing and redundant maintenance checkpoint cleanup | Done. Legacy non-incremental stores queue maintenance without WAL/freelist pressure, and automatic maintenance keeps a single final checkpoint after metadata dirties. | `search_service_sqlite_legacy_auto_vacuum_queues_idle_maintenance`, `sqlite_index_store_automatic_`, `search_service_sqlite_idle_maintenance_queue_and_completion`, `search_service_sqlite_status_reports_maintenance_history`. |
| FSBSI durable spec updates | Done for the completed scope. | `Specs/Core/Core_Search.md`, `Specs/FileSystem/FileSystem_FileOperations.md`, `Common/PlugInterfaces/FileSystem.h`. |

The old broad Track A/B/H text is therefore misleading where it asks to fix generic unknown-size bridge COPY, local writer overwrite staging, malformed search materialization, or SQLite legacy maintenance again. The remaining unresolved work has been split into the WIP follow-up plans linked above; detailed track text below is historical context for those splits and for the Track A/Track C implementation completed here.

## Old-to-New Track Mapping

| Original track | Current disposition |
| --- | --- |
| Track 0 - Existing WIP reconciliation | Closed. |
| Track A - Transfer completeness | Generic bridge unknown-size source guard is closed; provider short-body/read-proof work remains in Track F. |
| Track B - Atomic overwrite | Closed for this umbrella closeout. Local writer overwrite was closed by the FSBSI plan; same-FS copy/reparse overwrite staging was closed by Track A in this slice. |
| Track C - Cross-FS directory MOVE source delete safety | Still open as Track B. |
| Track D - Conflict consent bucket isolation | Closed by Track C in this slice. |
| Track E - Search service trust boundary and pipe protocol | Still open as Track D. |
| Track F - Reparse/canonicalization/traversal safety | Still open as Track E. |
| Track G - Provider-specific contract fixes | Still open as Track F. |
| Track H - Search/index correctness | Partially closed by malformed-entry and SQLite maintenance work; snapshot/root casing/hardlink/prefilter/buffered API issues remain Track G. |
| Track I - Secondary error surfacing and observability | Still open as Track H. |
| Track J - Threading, secret state, and plugin quiet points | Still open as Track I. |

## Current Remaining Track Map

| Track | Priority | Remaining issue | Current evidence to re-check | Primary targets |
| --- | --- | --- | --- | --- |
| A | P0 | Done in this slice. | Same-FS copy/reparse overwrite stages to sibling temp and promotes only after successful completion. | `Floodgate_LocalCopyOverwriteIsStaged`; `Specs/FileSystem/FileSystem_FileOperations.md`. |
| B | P0 | Split to follow-up WIP. | Cross-FS directory MOVE stale source delete remains planned, not hidden in this umbrella closeout. | `Specs/Plans/WIP/Operation_FSDeepAudit_TrackB_CrossFsDirectoryMoveDeleteSafety_2026-06-26.md`. |
| C | P0/P1 | Done in this slice. | Plain exists, non-empty directory, and reparse/link replacement are distinct conflict buckets for prompt text and apply-to-all cache scope. | `Fairstream_ReparseReplaceNonEmptyDirRequiresConsent`, `Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate`. |
| D | P1 | Split to follow-up WIP. | Search service trust boundary and pipe protocol remain planned, not hidden in this umbrella closeout. | `Specs/Plans/WIP/Operation_FSDeepAudit_TrackD_SearchServiceTrustBoundaryPipeProtocol_2026-06-26.md`. |
| E | P1 | Split to follow-up WIP. | Reparse, canonicalization, and traversal safety remain planned, not hidden in this umbrella closeout. | `Specs/Plans/WIP/Operation_FSDeepAudit_TrackE_ReparseCanonicalTraversalSafety_2026-06-26.md`. |
| F | P1 | Split to follow-up WIP. | Provider transfer, identity, delete-proof, and overwrite contracts remain planned, not hidden in this umbrella closeout. | `Specs/Plans/WIP/Operation_FSDeepAudit_TrackF_ProviderTransferContracts_2026-06-26.md`. |
| G | P2 | Split to follow-up WIP. | Search/index residual correctness remains planned, with P2 priority unless a sub-item proves data loss/security impact. | `Specs/Plans/WIP/Operation_FSDeepAudit_TrackG_SearchIndexResidualCorrectness_2026-06-26.md`. |
| H | P2 | Split to follow-up WIP. | Secondary partial-result observability remains planned, with escalation if a sub-item hides destructive failure. | `Specs/Plans/WIP/Operation_FSDeepAudit_TrackH_FileOpsObservabilityPartialResults_2026-06-26.md`. |
| I | P2 | Split to follow-up WIP. | Threading, secret state, plugin refresh/reload, and quiet-point work remains planned, with escalation for credential corruption or callback-after-unload risk. | `Specs/Plans/WIP/Operation_FSDeepAudit_TrackI_ThreadingSecretsPluginQuietPoints_2026-06-26.md`. |

## Global Engineering Rules

Apply these to every remaining track:

- Use WIL RAII for Win32 handles, COM pointers, GDI objects, and cleanup paths.
- Do not introduce `catch (...)`.
- Do not introduce `sprintf_s` or `swprintf_s`.
- Use non-throwing filesystem APIs with `std::error_code` in operation paths where practical.
- Do not block the UI thread on provider I/O, service I/O, or index maintenance.
- Use posted-payload helpers for cross-thread UI payloads.
- Add deterministic selftests before or with each data-safety/security fix.
- Any hot-path or responsiveness-sensitive change needs perf evidence archived under `Specs/TestRuns/`.
- Merge durable behavior into authoritative specs before closing the plan.

## Drift Check

Run before each implementation slice:

```powershell
git status --short
git rev-parse --short HEAD
git diff -- Common/PlugInterfaces/FileSystem.h `
  RedSalamander/FolderWindow.FileOperations.State.cpp `
  RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp `
  RedSalamander/FolderWindow.FileSystem.cpp `
  RedSalamander/HostServices.cpp `
  Plugins/FileSystem/FileSystem.cpp `
  Plugins/FileSystem/FileSystem.FileOps.cpp `
  Plugins/FileSystem/FileSystem.Path.cpp `
  Plugins/FileSystem/FileSystem.Watch.cpp `
  Plugins/FileSystem/FileSystem.Search.cpp `
  Common/SearchServiceBroker.cpp `
  RedSalamanderSearchService/Main.cpp `
  Common/LocalSearchIndexCore.cpp `
  Common/SqliteIndexStore.cpp `
  Plugins/FileSystem7z `
  Plugins/FileSystemCurl `
  Plugins/FileSystemS3 `
  Plugins/FileSystemMicrosoftDrive `
  Plugins/FileSystemGoogleDrive `
  Plugins/FileSystemDummy `
  Specs/FileSystem/FileSystem_FileOperations.md `
  Specs/Core/Core_Search.md `
  Specs/Plugins/Plugins_VirtualFileSystem.md
```

Stop and update this plan if another active change touches the same paths. Do not revert user changes.

## Track A - Same-FS Copy Overwrite Transaction

### Problem

The FSBSI work fixed `CreateFileWriter(ALLOW_OVERWRITE)`, but same-filesystem COPY still appears to use `CopyFileExW` over the final destination when overwrite is granted. A cancel, disk-full, I/O error, or power loss can destroy the pre-existing destination and leave a partial file.

### Targets

- `Plugins/FileSystem/FileSystem.FileOps.cpp`
- `RedSalamander/SelfTest/FileOperations/*`
- `Specs/FileSystem/FileSystem_FileOperations.md`

### Tasks

1. Replace overwrite-to-final `CopyFileExW` with a transaction:
   - Create a unique sibling temp in the destination directory.
   - Copy into the temp using `CREATE_NEW` semantics.
   - Flush temp data/metadata as appropriate.
   - Verify copied bytes against source size before promotion.
   - Replace the destination using `ReplaceFileW` where possible, with a documented fallback.
2. Apply the same contract to the reparse/symlink copy path where it currently uses `CopyFileExW`.
3. Make readonly handling transactional:
   - Do not restore attributes onto a partial destination.
   - Preserve or deliberately set final attributes only after verified promotion.
4. Preserve existing behavior when no overwrite is allowed.
5. Add tests:
   - Cancel/fail same-FS overwrite leaves original bytes and attributes.
   - Reparse copy overwrite failure leaves original link/object as documented.
   - Readonly overwrite failure restores original attributes.
   - No temp leak after failure.
6. Update `Specs/FileSystem/FileSystem_FileOperations.md` with same-FS copy overwrite transaction semantics.

### Validation

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_LocalWriterOverwrite
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Fairstream_Overwrite
git diff --check
```

Add a dedicated same-FS copy-overwrite failure case and include it in the validation list before closing this track.

## Track B - Cross-FS Directory MOVE Source Delete Safety

### Problem

Cross-filesystem directory MOVE can copy an enumeration snapshot and then recursively delete the live source tree. A file created or modified after enumeration can be deleted without ever being copied.

### Targets

- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `Plugins/FileSystem/FileSystem.FileOps.cpp` for same-FS move verification patterns
- `RedSalamander/SelfTest/FileOperations/*`
- `Specs/FileSystem/FileSystem_FileOperations.md`

### Tasks

1. Define safe cross-FS directory MOVE semantics:
   - Delete only source entries proven copied.
   - Preserve entries created after the copy snapshot.
   - Preserve or report partial for entries modified after copy unless the destination is proven to contain the same version.
2. Replace blanket recursive source delete with per-entry cleanup based on a copied manifest or a re-enumeration/verification pass.
3. Verification metadata should use what the provider can reliably supply:
   - Relative path.
   - Size.
   - Last-write time when stable.
   - Provider identity/version/etag when available.
   - Reparse/directory kind.
4. Report partial cleanup explicitly. Do not return final success if source entries remain because they could not be safely deleted.
5. Add tests:
   - File added after snapshot remains in source.
   - File modified after copy remains or produces `ERROR_PARTIAL_COPY`.
   - Copied file is deleted only after destination verification.
   - Directory with extra child is not removed.
   - Reparse source cleanup does not unexpectedly follow the target.

### Validation

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFs
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Riptide_Bridge
git diff --check
```

Archive a perf run if the copy/delete cleanup changes large-directory throughput.

## Track C - Conflict Consent Bucket Isolation

### Problem

Conflict handling still maps materially different destructive risks into `ConflictBucket::Exists`. A user's "Overwrite + Apply to all" for a plain file can be reused for non-empty directory or reparse destination cases.

### Targets

- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- Resource files for conflict prompt text if new prompts are needed
- `RedSalamander/SelfTest/FileOperations/*`
- `Specs/FileSystem/FileSystem_FileOperations.md`

### Tasks

1. Split or key conflict buckets by destructive risk:
   - Existing file.
   - Existing empty directory.
   - Existing non-empty directory.
   - Destination reparse point.
   - Source/destination type mismatch.
2. Include operation kind, source kind, destination kind, and provider capability where relevant in apply-to-all cache semantics.
3. Update prompt text for non-empty directory and reparse replacement so the user sees the risk class.
4. Add tests:
   - Plain file apply-to-all does not approve reparse replacement.
   - Plain file apply-to-all does not approve non-empty directory replacement.
   - Reparse apply-to-all is scoped only to equivalent reparse-risk cases.

### Validation

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Fairstream_Overwrite
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFs
git diff --check
```

## Track D - Search Service Trust Boundary and Pipe Protocol

### Problem

The search service can run with elevated/service authority while accepting broad pipe clients. Requests carry paths and patterns, and the service must prove that the client is allowed to access the requested root. Pipe framing, cancellation, and shutdown also need consistent overlapped/deadline behavior.

### Targets

- `Common/SearchServiceBroker.cpp`
- `RedSalamanderSearchService/Main.cpp`
- `Specs/Core/Core_Search.md`
- Search/Compare service selftests or a dedicated fake pipe test harness

### Tasks

1. Decide and document the service identity model:
   - Prefer least privilege.
   - If elevated identity remains, authorize path access under the client identity.
2. Tighten pipe access and isolation:
   - Review SDDL.
   - Isolate by session/user where required.
3. Add client authorization:
   - Use `ImpersonateNamedPipeClient` plus RAII revert, or equivalent access checks.
   - Reject roots the client cannot read.
   - Log denied roots as warnings, not successful empty results.
4. Validate request scope:
   - Reject empty roots, malformed extended paths, unsupported device namespaces, and roots outside supported providers.
   - Canonicalize roots before cache/index lookup.
5. Fix pipe I/O semantics:
   - Use overlapped I/O consistently or create synchronous handles.
   - Bound frame size and per-frame read/write time.
   - Treat short read/write as protocol failure.
   - Disconnect failed clients and continue serving later clients.
   - Cancel outstanding I/O on shutdown.
6. Add tests:
   - Unauthorized root denied.
   - Malformed root denied.
   - Partial/short frame fails without hang.
   - Slow client cannot block all service work.
   - Shutdown cancels in-flight request.
7. Update `Specs/Core/Core_Search.md`.

### Validation

```powershell
.\build.ps1 -ProjectName RedSalamanderSearchService
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild
git diff --check
```

Archive service responsiveness evidence if request scheduling or pipe I/O changes.

## Track E - Reparse Point, Canonicalization, and Traversal Safety

### Problem

Some metadata, path, watch, and traversal paths still conflate link-object and link-target behavior. Extended path canonicalization and recursive traversal also need explicit guards against wrong-target operations and cycles.

### Targets

- `Plugins/FileSystem/FileSystem.Path.cpp`
- `Plugins/FileSystem/FileSystem.FileOps.cpp`
- `Plugins/FileSystem/FileSystem.Watch.cpp`
- `Plugins/FileSystem/FileSystem.Search.cpp`
- `Common/LocalSearchIndexCore.cpp`
- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/Core/Core_Search.md`

### Tasks

1. Classify every open by intent:
   - Target contents.
   - Object metadata.
   - Reparse point itself.
   - Recursive traversal.
2. Add `FILE_FLAG_OPEN_REPARSE_POINT` where metadata/watch/delete operations must apply to the link object.
3. Canonicalize `\\?\` and normal Win32 forms consistently:
   - Collapse `.` and `..`.
   - Normalize trailing dot/space policy.
   - Preserve display casing separately from comparison keys.
4. Add traversal guards:
   - Physical visited-set when following local reparse directories.
   - Maximum depth and queued-directory caps for providers without identity.
   - Visible warning when traversal is incomplete due to a guard.
5. Add tests:
   - Metadata copy stamps link object, not link target.
   - `\\?\` path with `..` cannot escape the approved root.
   - Recursive search/live scan does not loop through a junction cycle.
   - Watch setup reports or handles reparse roots according to the documented policy.

### Validation

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild
git diff --check
```

## Track F - Provider-Specific Contract Fixes

### Problem

Generic bridge unknown-size behavior is fixed, but provider implementations still have short-body, delete-proof, identity-proof, and overwrite-contract gaps. These must be fixed with fake transport/provider tests rather than relying on live cloud accounts.

### Targets

- `Plugins/FileSystem7z/*`
- `Plugins/FileSystemCurl/*`
- `Plugins/FileSystemS3/*`
- `Plugins/FileSystemMicrosoftDrive/*`
- `Plugins/FileSystemGoogleDrive/*`
- `Plugins/FileSystemDummy/*`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- Provider selftests/fakes

### Tasks by provider

#### 7z

- Detect short extraction/spool reads and return failure instead of `S_OK` EOF.
- Reject or canonicalize `..`/absolute archive entry keys.
- Detect file/directory key collisions and surface invalid archive data.

#### Curl / HTTP / IMAP

- Verify `Content-Length` when present and fail premature EOF.
- Treat partial upload/server acceptance as failure unless protocol proves completion.
- For IMAP, validate UID/UIDVALIDITY where possible.
- Do not use mailbox-wide EXPUNGE for single-message delete without preserving unrelated `\Deleted` messages.

#### S3

- Verify download body length against object metadata when available.
- For ranged readers, distinguish short range response from genuine EOF.
- Verify destination size/ETag/version before delete-after-copy.
- Treat broken pagination as invalid data, not restart/infinite loop.
- Persist or retry cleanup for orphaned multipart uploads and backup objects.

#### Microsoft Drive

- Verify `Content-Range` for ranged reads.
- Honor upload-session `nextExpectedRanges`.
- Revalidate provider IDs before destructive delete.
- Treat backup cleanup failures separately from already-completed user-visible moves.

#### Google Drive

- Treat IDs as case-sensitive and names as display metadata.
- Avoid synthetic marker collisions and support unnamed objects safely.
- Revalidate object ID before destructive delete or overwrite.

#### Dummy FS

- Make recursive delete honor readonly descendants unless `ALLOW_REPLACE_READONLY` is granted.
- Make directory merge/copy/move report partial without silently emptying source on first collision.
- Use Dummy FS as a deterministic contract test provider for host behavior.

### Validation

```powershell
.\build.ps1
.\Tools\Run-AllTests.ps1 -Suite Full
```

For providers requiring credentials, add fake transport tests first. Do not make CI depend on live S3, Drive, HTTP, FTP, or IMAP accounts.

## Track G - Search and Index Residual Correctness

### Problem

Malformed local search materialization and SQLite legacy maintenance are fixed. Other search/index audit items remain: non-atomic snapshot writes, exact-case volume root lookup, hardlink/path modeling, conservative prefilters, buffered result failure behavior, missing integrity/corruption handling, and delta count contracts.

### Targets

- `Common/LocalSearchIndexCore.cpp`
- `Common/SqliteIndexStore.cpp`
- `Common/SearchServiceBroker.cpp`
- `Specs/Core/Core_Search.md`
- Compare selftests

### Tasks

1. Make snapshot writes atomic:
   - Write temp sibling.
   - Flush.
   - Replace final snapshot.
   - Preserve previous snapshot on failed write.
2. Normalize index root lookup:
   - Local root identity comparison must be case-insensitive.
   - Preserve display casing separately.
3. Decide and document hardlink/multiple-path behavior:
   - Store aliases or explicitly report non-exhaustive alias coverage.
4. Fix conservative prefilters:
   - Literal and extension prefilters may produce false positives, never false negatives.
5. Fix buffered service result semantics:
   - Failed `QueryComplete` clears partial candidates or reports partial status.
   - No-batch overflow returns capped prefix plus warning when contract allows.
6. Add SQLite robustness:
   - Treat busy checkpoint as non-fatal maintenance deferral where safe.
   - Run `quick_check` before destructive `VACUUM` rewrite or document why not.
   - Fix `ApplyJournalDelta` skipped-operation counts.
7. Add tests:
   - Failed snapshot write preserves prior snapshot.
   - Differently-cased roots reuse one index.
   - Literal and multi-part extension prefilters retain valid hits.
   - Hardlink alias behavior matches the spec.
   - Buffered failure/overflow semantics are deterministic.

### Validation

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild
git diff --check
```

Archive index/search evidence if maintenance or query scheduling changes.

## Track H - Secondary Error Surfacing and Observability

### Problem

Several secondary paths still hide partial work or weak diagnostics. They are not the first data-loss cuts, but they keep users and tests from seeing incomplete results.

### Targets

- `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`
- `Plugins/FileSystem/FileSystem.Watch.cpp`
- `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp`
- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp`
- `Common/MonitorIssueStore*`
- Relevant specs

### Tasks

1. `GetDirectorySize`:
   - Use extended paths.
   - Return explicit partial status for denied/vanished children.
   - Preserve best-effort total separately from complete total.
2. Watch overflow/rearm:
   - Surface zero-byte completion as overflow/resync-required.
   - Surface rearm failure instead of silently killing the watch.
3. Diagnostic writes:
   - Check write byte counts.
   - Treat short writes as diagnostic failure.
4. Issue dedup:
   - Keep provider/root/path context so distinct failures do not collapse.
5. Temp/progress cleanup:
   - Add policy for orphaned `.rs_tmp_*` sweep or user-visible cleanup.
   - Make progress underestimation visible instead of freezing silently.
6. Add deterministic tests where possible.

### Validation

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Full
git diff --check
```

## Track I - Threading, Secret State, Plugin Refresh, and Quiet Points

### Problem

Host and provider state paths still need explicit lock/lifetime/persistence guarantees. These interact with filesystem plugins and search hosts, so failures can become wrong-provider operations, credential-state corruption, or callback-after-unload hazards.

### Targets

- `RedSalamander/HostServices.cpp`
- `RedSalamander/FolderWindow.FileSystem.cpp`
- Provider connection/session code identified during drift checks
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `AGENTS.md` only if a repo-wide rule is missing

### Tasks

1. Inventory shared mutable state:
   - Session secrets.
   - Persistent connection secrets.
   - Provider session state.
   - Callback cookies.
   - In-flight request registries.
2. Add or verify locking:
   - No shared non-atomic state read/write without a lock.
   - Do not hold locks while invoking plugin callbacks or doing I/O.
3. Fix host services secret ordering:
   - Persist before mutating session cache when durability is required.
   - On persistence failure, do not expose the new secret as live.
   - If the host window is unavailable, return a clear transient failure instead of running UI-thread bodies on the worker.
4. Fix plugin refresh/reload context:
   - Preserve or explicitly reinitialize mount/connection context.
   - If reload fails, neutralize stale pane state and surface failure.
5. Document plugin unload quiet point:
   - Stop producers.
   - Request shutdown/cancel.
   - Stop posting payload messages.
   - `SetCallback(nullptr, nullptr)`.
   - Release instances.
   - Unload only when callbacks cannot run.
6. Add stress/selftests:
   - Concurrent secret set/get.
   - Persist failure does not mutate live cache.
   - Plugin reload failure leaves panes in a safe state.
   - Provider shutdown with in-flight callback does not use freed callback state.

### Validation

```powershell
.\build.ps1
.\Tools\Run-AllTests.ps1 -Suite Full
git diff --check
```

## Active Cross-Track Test Matrix

| Risk | Required deterministic test |
| --- | --- |
| Same-FS overwrite | Injected/cancelled overwrite leaves original destination bytes and attributes intact. |
| Cross-FS directory MOVE stale delete | File created after copy snapshot remains in source and the item reports partial. |
| Conflict consent leak | Plain-file apply-to-all does not approve reparse or non-empty directory replacement. |
| Search service authorization | Client/root unauthorized under client identity is denied. |
| Pipe protocol failure | Short read/write fails request without hanging the service. |
| Reparse metadata | Link metadata operation does not stamp target unexpectedly. |
| Canonicalization escape | `\\?\` path with `..` cannot bypass root containment. |
| Recursive traversal cycle | Junction cycle is bounded and reports incomplete traversal. |
| Provider short body | Fake 7z/S3/HTTP/Drive provider fails incomplete body. |
| Provider delete proof | Copy-then-delete source happens only after destination identity/size proof. |
| Snapshot failure | Failed index snapshot write preserves last good snapshot. |
| Root casing | Differently-cased local roots do not create duplicate/shadowed index rows. |
| Prefilter false negative | Literal and multi-dot extension searches retain valid hits. |
| Watch overflow | Watch reports resync-required state. |
| Secret concurrency | Concurrent secret access is locked and persist failure does not update live cache. |

## Performance Validation

Performance validation is mandatory for tracks touching copy loops, directory enumeration, provider I/O scheduling, search/index maintenance, watch, service pipe I/O, or queueing.

For each touched perf-sensitive track:

1. Define the scenario before coding.
2. Capture baseline command/output if the scenario already exists.
3. Add or reuse instrumentation.
4. Run deterministic selftests.
5. Archive evidence under `Specs/TestRuns/` with date, commit, command, and summary.

Suggested evidence files:

- `Specs/TestRuns/YYYYMMDD_SameFsOverwriteTransaction.txt`
- `Specs/TestRuns/YYYYMMDD_CrossFsMoveDeleteSafety.txt`
- `Specs/TestRuns/YYYYMMDD_SearchServicePipeSecurity.txt`
- `Specs/TestRuns/YYYYMMDD_ReparseTraversalSafety.txt`
- `Specs/TestRuns/YYYYMMDD_ProviderTransferContracts.txt`
- `Specs/TestRuns/YYYYMMDD_SearchIndexResidualCorrectness.txt`

## Spec Updates Required

Do not close this plan until durable behavior is merged into authoritative specs:

- `Specs/FileSystem/FileSystem_FileOperations.md`
  - Same-FS copy overwrite transaction semantics.
  - Cross-FS directory MOVE manifest/per-entry cleanup.
  - Conflict consent buckets.
  - Reparse/link operation policy.
  - Directory size/watch/diagnostic partial-result status.
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
  - Provider completion proof, identity proof, delete proof, and overwrite behavior.
  - Recursive delete readonly-descendant contract.
  - Directory-change string lifetime if changed.
  - Plugin unload quiet point if not already sufficiently documented.
- `Specs/Core/Core_Search.md`
  - Search service identity and client authorization.
  - Pipe protocol failure behavior.
  - Snapshot atomicity and root canonicalization.
  - Hardlink/path alias semantics and buffered partial-result semantics.
- `AGENTS.md`
  - Update only if implementation discovers a repo-wide rule that belongs outside domain specs.

## Recommended Implementation Sequence

1. Track A: same-FS copy overwrite transaction.
2. Track C: conflict consent isolation.
3. Track B: cross-FS directory MOVE delete safety.
4. Track D: search service trust boundary and pipe protocol.
5. Track E: reparse/canonicalization/traversal.
6. Track F: provider-specific short-transfer and delete-proof contracts, grouped by provider.
7. Track G: search/index residual correctness.
8. Track H: secondary observability.
9. Track I: threading/secret/plugin quiet points.
10. Required spec updates, perf archives, broad validation, and move to Done.

Stop conditions:

- If same-FS overwrite preservation requires ACL/ADS/security descriptor redesign, implement the narrow content-preservation transaction first and split metadata fidelity into a follow-up.
- If search service impersonation cannot be tested without service install/elevation, add a fake pipe/broker harness before changing production security code.
- If provider live credentials are required, add fake transport coverage first and document live validation separately.
- If a fix materially changes user-visible behavior, update resources/specs in the same change set.

## Final Verification

Before moving this plan to `Specs/Plans/Done/`:

```powershell
git diff --check
.\build.ps1
.\Tools\Run-AllTests.ps1 -Suite Full
```

Also verify:

- All required specs are updated.
- Perf evidence files exist for every touched perf-sensitive path.
- `Specs/Reviews/FS-Subsystem-Findings.md` is referenced by closeout notes, with remaining items either implemented or split into follow-up WIP plans.
- No new `catch (...)`, `sprintf_s`, `swprintf_s`, raw owning COM pointer, raw Win32 cleanup, or `DestroyWindow(_hWnd.get())` pattern was introduced.

## Done Criteria

This plan is done when:

- P0 and P1 tracks A through F are implemented or explicitly split into smaller WIP plans with no unplanned data-loss/security item left behind.
- P2 tracks G through I are implemented or explicitly downgraded with rationale in this plan.
- Deterministic tests cover the active cross-track matrix.
- Required specs are updated.
- Perf evidence is archived for every touched perf-sensitive path.
- Full validation is green, or any unrelated pre-existing failure is documented with command, artifact path, and process cleanup notes.
- This file is moved from `Specs/Plans/WIP/` to `Specs/Plans/Done/`.
