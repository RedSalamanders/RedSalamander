# Operation FileSystem Bridge/Search/SQLite Data-Safety Study and Fix - 2026-06-25

> **Executor instructions**: This is a WIP remediation plan, not a completed design.
> Read it fully before editing. Execute the phases in order. Run every verification
> command listed for a phase before moving on. If any STOP condition triggers, stop
> and report with the exact file/line and command output. Do not improvise around
> data-loss semantics.

## Status

- **Priority**: P1 for search silent-drop and bridge unknown-size COPY; P2 for SQLite maintenance and same-FS overwrite; P3 for telemetry-only cleanup.
- **Effort**: L overall. Tracks A and B1 are S/M; tracks C and D require study plus fault-injection tests.
- **Risk**: HIGH where file operations can delete or overwrite user data; MEDIUM for search/index maintenance; LOW for redundant checkpoint cleanup.
- **Depends on**: current dirty working tree from the search/index remediation session.
- **Category**: bug, data-safety, search correctness, file operations, performance, diagnostics, tests, docs.
- **Planned at**: commit `ec403afac`, 2026-06-25, with an intentionally dirty worktree.
- **Plan location**: `Specs/Plans/WIP/Operation_FileSystemBridgeSearchIndex_DataSafetyStudyAndFix_2026-06-25.md`.

## Repository State at Planning Time

- Working directory: `D:\RedSalamander`.
- Branch: `master`.
- HEAD: `ec403afac`.
- Worktree is intentionally dirty from the prior search/index fixes. Do not reset, stash, or discard changes.
- Relevant changed files already in the working tree include:
  - `Common/SqliteIndexStore.cpp`
  - `Common/SqliteIndexStore.h`
  - `Plugins/FileSystem/FileSystem.Search.cpp`
  - `RedSalamander/FindFilesWindow.cpp`
  - `RedSalamander/SearchFallbackEngine.cpp`
  - `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
  - `Specs/Core/Core_Search.md`

## Drift Check - Run First

Run:

```powershell
git status --short
git diff --stat
git diff --stat ec403afac..HEAD -- Common\SqliteIndexStore.cpp Common\SqliteIndexStore.h Common\PlugInterfaces\FileSystem.h Plugins\FileSystem\FileSystem.Search.cpp Plugins\FileSystem\FileSystem.cpp Plugins\FileSystem\FileSystem.FileOps.cpp RedSalamander\FolderWindow.FileOperations.State.cpp RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp RedSalamander\SelfTest\FileOperations RedSalamander\SelfTest\Commands Specs\Core\Core_Search.md Specs\FileSystem\FileSystem_FileOperations.md
```

Expected:

- `git status --short` shows this plan plus the current dirty search/index files.
- `git diff --stat ec403afac..HEAD ...` may be empty if no commits landed after planning; the dirty worktree is still authoritative.
- If the cited files have changed substantially since this plan was written, compare the "Current State Evidence" excerpts below against live code before implementing.

STOP if:

- The working tree contains unrelated user edits in the same files and you cannot separate them from this remediation.
- A cited symbol no longer exists.
- A needed change appears to require deleting or rewriting unrelated file-operation behavior.

## Why This Matters

This plan covers the remaining findings from a focused file-system plugin, bridge, search, and SQLite maintenance review. The common theme is "never silently lose user information": search must not silently omit files, COPY/MOVE must not report success when integrity is unknown, and overwrite paths should not destroy the existing destination before the new bytes are durable. Some findings are in the current search/index diff; others are broader bridge/file-operation gaps that became relevant during the review. The executor must treat data-integrity correctness as more important than preserving optimistic success paths.

## Findings Covered

| ID | Priority | Verdict | Area | Summary |
|----|----------|---------|------|---------|
| FSBSI-A | P1 | Confirmed | Search plugin | `MaterializeEntryPaths` failure now silently skips entries; parallel chunk worker still uses `return`, dropping the rest of a chunk with no warning. |
| FSBSI-B | P2 | Plausible | SQLite service maintenance | Automatic `VACUUM` may fail/requeue under a sustained writer; queueing may not be triggered solely by legacy `auto_vacuum` mode. Study and fix scheduling/idle gating. |
| FSBSI-C | P3 | Confirmed cleanup | SQLite maintenance | Automatic auto-vacuum rewrite path performs redundant checkpoint work. |
| FSBSI-D | P3 | Confirmed diagnostics gap | SQLite maintenance | `AutomaticMaintenanceResult::ranVacuum` is test-visible but not surfaced in service diagnostics/events. |
| FSBSI-E | P1 | Confirmed design gap | Cross-filesystem bridge | COPY with unknown source size lacks the MOVE guard and final size backstop; can report a staged copy as success when integrity was not proven. |
| FSBSI-F | P2 | Confirmed design gap | Local filesystem writer / same-FS overwrite | `CreateFileWriter(...ALLOW_OVERWRITE...)` opens the real destination with `CREATE_ALWAYS`; abort deletes/truncates the pre-existing destination instead of preserving it via staging. |
| FSBSI-G | P2 | Follow-up nuance | SQLite maintenance queue | Legacy non-incremental `auto_vacuum` rewrite works once maintenance runs, but `ShouldQueueAutomaticMaintenance` currently appears to queue only for WAL/free-page thresholds. |

## Refuted Items - Do Not Re-flag

- `ApplyPragmas` masks legacy `auto_vacuum`: refuted empirically. On existing non-empty DBs, `PRAGMA auto_vacuum=INCREMENTAL` does not make `PRAGMA auto_vacuum` report `2` until `VACUUM` persists the mode.
- `StartsWithPreFoldedInvariant` changes prefix semantics: refuted. Current folding is length-preserving simple lowercasing through Windows APIs.
- `FactoryEnumeratePlugins` empty-span guard is unsafe: refuted. The implementation zeros out params before returning `S_OK`.
- `NamePrefilter.upperBound` caching diverges from SQL binding: refuted. The prefix upper bound is computed at the single prefix prefilter construction site and reused consistently.

## Repo Conventions to Follow

- Use WIL RAII wrappers for Windows resources. No raw handle ownership.
- Do not introduce `catch (...)`.
- Do not introduce `sprintf_s` or `swprintf_s` in non-PoC code.
- User-facing text must use `.rc` string resources. Resource format placeholders must be positional.
- For cross-thread payloads, use the existing posted-payload helpers; do not raw-post allocated pointers.
- Follow the existing HRESULT style. Win32 failures should return `HRESULT_FROM_WIN32(GetLastError())` or a specific documented HRESULT.
- For tests, prefer deterministic selftests under `RedSalamander/SelfTest/...`; archive perf evidence under `Specs/TestRuns/` only from test runs, not by manually editing archived outputs.
- Do not search or edit archived `Specs/TestRuns/**` unless explicitly investigating a historical run. The live source/spec/test files are the authority.

## Commands You Will Need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build Debug target | `.\build.ps1 -ProjectName RedSalamander` | exit 0; `Diagnostics: 0 warning(s), 0 error(s)` |
| Full selftest without rebuild | `.\Tools\Run-AllTests.ps1 -SkipBuild` | exit 0, or documented unrelated/environmental failures only |
| Compare targeted search tests | `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_low_hardening_smoke` | passed |
| Compare targeted SQLite tests | `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter sqlite_index_store_automatic_` | all matching cases passed |
| Commands targeted regex status | `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter search_local_plugin_invalid_regex_reports_single_completion` | passed |
| FileOps targeted bridge tests | `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter <new-case-prefix>` | new cases passed |
| Diff hygiene | `git diff --check` | exit 0; line-ending warnings are acceptable, whitespace errors are not |

## Scope

### In Scope

Search skip/warning work:

- `Common/PlugInterfaces/FileSystem.h`
- `Plugins/FileSystem/FileSystem.Search.cpp`
- `RedSalamander/SearchFallbackEngine.cpp` only if warning bit propagation must stay consistent between native and fallback search.
- `RedSalamander/FindFilesWindow.cpp`
- `RedSalamander/RedSalamander.rc`
- `RedSalamander/Lang/*/RedSalamander-*.rc`
- `RedSalamander/Resource.h`
- `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- `Specs/Core/Core_Search.md`

SQLite maintenance work:

- `Common/SqliteIndexStore.cpp`
- `Common/SqliteIndexStore.h`
- `Common/SearchServiceBroker.cpp`
- `Common/LocalSearchIndexCore.cpp`
- `Common/LocalSearchIndexCore.h`
- `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- `Specs/Core/Core_Search.md`

Bridge/file-operation data-safety work:

- `Common/PlugInterfaces/FileSystem.h`
- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `RedSalamander/SelfTest/FileOperations/**`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp` only if a command-level regression is the right harness.
- `Specs/FileSystem/FileSystem_FileOperations.md`

Local writer/overwrite study:

- `Plugins/FileSystem/FileSystem.cpp`
- `Plugins/FileSystem/FileSystem.FileOps.cpp`
- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `RedSalamander/SelfTest/FileOperations/**`
- `Specs/FileSystem/FileSystem_FileOperations.md`

### Out of Scope

- Cloud-provider checksum redesigns beyond what is required to test bridge unknown-size behavior.
- S3/Curl/MSDrive recursive MOVE verification contracts already tracked in other WIP/Done plans, unless a test hook is required for the generic bridge.
- ViewerWeb, folder rendering, menu/input, or unrelated Commands failures.
- Archived `Specs/TestRuns/**` content.
- Any commit, branch creation, push, or PR unless the user explicitly asks.

## Current State Evidence

Line numbers drift; verify with `rg` before editing.

### FSBSI-A: Search `MaterializeEntryPaths` skip paths are silent

`Plugins/FileSystem/FileSystem.Search.cpp` currently has:

```cpp
if (! MaterializeEntryPaths(metadata, directoryFullPath, relativeBase, &runtime.materializedEntryPathPairs))
{
    return;
}
```

inside `AccumulateParallelNameOnlyEntry`. This skips one matched entry but sets no warning.

`BuildParallelDirectoryResult` currently has:

```cpp
if (! MaterializeEntryPaths(metadata, frame.fullPath, frame.relativeBase, &runtime.materializedEntryPathPairs))
{
    continue;
}
```

This is the right control-flow direction for recursive directory enqueue, but it still sets no warning.

`ParallelScanWorkerBody` currently has two malformed-entry exits:

```cpp
if (! MaterializeEntryPaths(entry, chunk->directoryFullPath, chunk->relativeBase, &runtime.materializedEntryPathPairs))
{
    return;
}
```

Those two `return`s abandon the rest of the chunk and set no warning. This is the highest-priority in-diff bug.

`MaterializeEntryPaths` itself fails only when the directory path or display name is empty:

```cpp
if (directoryFullPath.empty() || metadata.displayName.empty())
{
    return false;
}
```

The trigger is narrow, but the result is a clean completion with missing matches.

### FSBSI-B/G: SQLite maintenance rewrite runs, but queue/scheduling needs study

`Common/SqliteIndexStore.cpp` currently computes:

```cpp
const bool shouldRewriteAutoVacuum = ! result.before.incrementalAutoVacuumEnabled;
...
result.maintenanceNeeded = shouldCheckpoint || shouldRewriteAutoVacuum || shouldIncrementalCompact;
```

and runs:

```cpp
PRAGMA auto_vacuum=INCREMENTAL;
VACUUM;
```

This is correct once `RunAutomaticMaintenance` is called.

`Common/SearchServiceBroker.cpp` queueing appears to be based on checkpoint/compaction only:

```cpp
const bool shouldCheckpoint =
    storeInfo.autoCheckpointEnabled && storeInfo.autoCheckpointTargetBytes != 0u && storeInfo.writeAheadLogBytes >= storeInfo.autoCheckpointTargetBytes;
...
if (! storeInfo.autoCompactionEnabled || storeInfo.pageCount == 0u || storeInfo.freelistPageCount == 0u)
{
    return false;
}
```

Study whether a legacy non-incremental auto-vacuum store with low WAL and no freelist pages ever queues maintenance automatically. If not, queue it.

`OpenConnection` applies a busy timeout:

```cpp
sqlite3_busy_timeout(connection.get(), kSqliteBusyTimeoutMs);
```

but `VACUUM` still fails if the writer remains active past that timeout. Study whether the broker loop guarantees no warmup/rebuild/index writer is active while maintenance runs.

### FSBSI-C/D: Redundant checkpoint and limited visibility

The auto-vacuum path does:

1. pre-maintenance `RunWalCheckpoint(..., "automatic WAL checkpoint")`;
2. `VACUUM`;
3. `RunWalCheckpoint(..., "post-auto-vacuum-rewrite WAL checkpoint")`;
4. trailing final checkpoint because `result.ranCheckpoint` is true.

This is probably more I/O than needed. Preserve durability and metadata correctness, but avoid redundant full checkpoint work.

`AutomaticMaintenanceResult::ranVacuum` exists in `Common/SqliteIndexStore.h`, but service status/events do not obviously surface "full VACUUM rewrite happened" versus cheap checkpoint/incremental vacuum. Add diagnostics only if useful and low-risk.

### FSBSI-E: Cross-FS COPY with unknown source size lacks MOVE's guard

`RedSalamander/FolderWindow.FileOperations.State.cpp` currently reads source size:

```cpp
HRESULT hrReaderSize = reader->GetSize(&fileTotalBytes);
if (SUCCEEDED(hrReaderSize))
{
    hasKnownFileTotalBytes = true;
}
else if (task._operation == FILESYSTEM_MOVE)
{
    return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
}
```

COPY proceeds when `GetSize` fails. Later, the only final size mismatch check is gated on `hasKnownFileTotalBytes`:

```cpp
if (hasKnownFileTotalBytes && fileCompletedBytes != fileTotalBytes)
{
    return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
}
```

`IFileReader::Read` is only documented as "Minimal Win32-like file reader"; it does not explicitly say `S_OK` with `bytesRead == 0` can occur only at EOF. If an implementation returns a premature zero-byte read and source size is unknown, the bridge treats it as EOF, commits/promotes the destination, and reports success.

### FSBSI-F: Local writer overwrites in place

`Plugins/FileSystem/FileSystem.cpp` currently uses:

```cpp
const DWORD creationDisposition = allowOverwrite ? CREATE_ALWAYS : CREATE_NEW;
```

for `CreateFileWriter`. `CREATE_ALWAYS` truncates an existing destination before new bytes are durable.

The writer destructor currently aborts by deleting the path:

```cpp
if (! _committed)
{
    _file.reset();
    DeleteFileW(_path.c_str());
}
```

For a pre-existing destination, abort can destroy the original file. This may match raw Win32 copy behavior, but it is weaker than the cross-FS bridge's temp-stage/promote pattern.

## Execution Plan

### Phase 0: Reconfirm Findings and Choose Narrow Test Hooks

1. Re-run focused `rg` reads for the symbols in this plan:

   ```powershell
   rg -n "MaterializeEntryPaths|ParallelScanWorkerBody|ShouldQueueAutomaticMaintenance|RunAutomaticMaintenance|CreateFileWriter|CREATE_ALWAYS|GetSize\\(|ERROR_PARTIAL_COPY" Plugins\FileSystem Common RedSalamander
   ```

   Expected: all cited symbols exist.

2. Identify existing test harnesses:

   ```powershell
   rg -n "RunCase\\(|Bridge|GetSize|FileWriter|CreateFileWriter|search_low_hardening_smoke|sqlite_index_store_automatic_" RedSalamander\SelfTest
   ```

   Expected: Compare search/SQLite cases and FileOps/Commands file-operation cases are discoverable.

3. Decide test placement:
   - Search warning/skip tests: `CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`.
   - SQLite queue/maintenance tests: same Compare selftest file.
   - Bridge unknown-size COPY tests: prefer `RedSalamander/SelfTest/FileOperations/...` if the file-operation selftest harness can exercise a fake bridge reader. If not, use a focused Commands FileOps selftest with a fake `IFileSystemIO`.
   - Same-FS overwrite tests: prefer FileOps selftests; use plugin-level `IFileSystemIO::CreateFileWriter` tests if a direct test is simpler.

**Verify**: no code changes yet. `git diff --stat` should show only the pre-existing dirty worktree plus this plan.

STOP if no deterministic harness can exercise the bridge without real data loss; write a small fake plugin/IO harness first rather than using live user files.

### Phase A: Fix Search Malformed-Entry Skips

Goal: malformed entries are skipped at the smallest safe granularity and the user gets a visible warning that results may be incomplete.

1. Define the warning contract.
   - Preferred: add a new bit such as `FILESYSTEM_SEARCH_WARNING_ENTRIES_SKIPPED = 0x40` in `Common/PlugInterfaces/FileSystem.h`.
   - Add localized Find status text in:
     - `RedSalamander/Resource.h`
     - `RedSalamander/RedSalamander.rc`
     - `RedSalamander/Lang/cs-CZ/RedSalamander-cs-CZ.rc`
     - `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`
     - `RedSalamander/Lang/ja-JP/RedSalamander-ja-JP.rc`
     - `RedSalamander/Lang/sk-SK/RedSalamander-sk-SK.rc`
   - Plumb it through `FindFilesWindow::BuildWarningSummary`.
   - If adding a new ABI-visible bit is not acceptable, STOP and ask whether to reuse `FILESYSTEM_SEARCH_WARNING_OVERFLOW`. Reusing overflow is lower fidelity but avoids expanding the enum.

2. Add a helper in `Plugins/FileSystem/FileSystem.Search.cpp` to flag malformed skipped entries consistently.
   - Keep it small and local, for example `FlagMalformedSearchEntrySkipped(...)`.
   - It must update the correct warning accumulator for each path:
     - serial/native runtime: `runtime.warningFlags |= ...`;
     - parallel directory walk result/state: `result.warningFlags |= ...` and/or `state.warningFlags.fetch_or(...)`;
     - parallel chunk: `result.warningFlags |= ...`.
   - Avoid unsynchronized writes to `runtime.warningFlags` from worker threads.

3. Replace the two `return` statements in `ParallelScanWorkerBody` after `MaterializeEntryPaths` failure with warning + `continue`.
   - The worker should continue to later entries in the chunk.
   - Do not increment candidate/content counters after a skipped entry.
   - Existing cancellation and real I/O error `return`s must remain unchanged.

4. Add warnings to the existing `continue`/`return S_OK`/bare `return` skip sites:
   - `AccumulateParallelNameOnlyEntry`
   - `BuildParallelDirectoryResult`
   - `EvaluateEntry`
   - `ParallelScanWorkerBody`
   - `SearchDirectoryTree`
   - any equivalent fallback path if applicable.

5. Update source guards in `search_low_hardening_smoke`.
   - Assert the parallel worker malformed-entry block contains `continue;`, not bare `return;`.
   - Assert the warning bit is referenced in malformed-entry skip paths.
   - Assert `FindFilesWindow::BuildWarningSummary` includes the new warning resource.

6. Add or adapt a deterministic behavioral test.
   - Prefer a fake `IFilesInformation` entry with empty `FileNameSize` followed by a valid matching sibling.
   - Expected: valid sibling is still returned; final progress warning includes the malformed-entry warning bit.
   - If creating a real malformed plugin buffer is too invasive, keep the source guard and add a lower-level helper test for the skip helper.

**Verify**:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_low_hardening_smoke
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter local_search_scan_wide_tree_parallel_walk_name_only
git diff --check
```

Expected: build 0 warnings/0 errors; targeted Compare tests pass.

### Phase B: Fix SQLite Maintenance Queueing, Busy Handling, Checkpoint Efficiency, and Diagnostics

Goal: legacy `auto_vacuum` rewrite is queued when needed, maintenance does not spin wastefully under active writers, and checkpoints/telemetry are coherent.

1. Study scheduler ownership.
   - Read `Common/SearchServiceBroker.cpp` around `RefreshAutomaticMaintenanceQueue`, `ShouldQueueAutomaticMaintenance`, `RunQueuedAutomaticMaintenance`, and broker main loop.
   - Read `Common/LocalSearchIndexCore.cpp` warmup/rebuild writer paths that can hold SQLite writer transactions.
   - Determine whether automatic maintenance can overlap startup warmup/rebuild or only runs when broker is idle.

2. Queue legacy auto-vacuum rewrite explicitly.
   - Extend `ShouldQueueAutomaticMaintenance(...)` so a SQLite store with `inspectionSucceeded == true` and `incrementalAutoVacuumEnabled == false` queues maintenance even when WAL/free pages are small.
   - Ensure this does not queue for invalid/missing stores.
   - Add a Compare selftest that creates a legacy non-incremental store with no meaningful freelist/WAL pressure and verifies service/broker maintenance queue state or direct scheduler predicate if exposed.
   - If `PersistentStoreInfo` does not expose enough auto-vacuum state, add the smallest field needed and populate it from existing inspection data.

3. Avoid maintenance livelock under active writers.
   - If there is already a central repository/indexer state that says warmup/rebuild/write is active, make `ShouldRunAutomaticMaintenanceNow` or `RunQueuedAutomaticMaintenance` defer until idle.
   - If no such signal exists, add a conservative lock/probe:
     - maintenance may attempt once;
     - on `SQLITE_BUSY`/busy-derived HRESULT, log a concise warning, keep maintenance queued with backoff, and do not repeatedly checkpoint before retrying.
   - Preserve fail-open query behavior. Queries should still degrade/fallback rather than block on maintenance.
   - STOP if the fix requires a large redesign of service threading; document the needed design instead.

4. Reduce redundant checkpoints.
   - For auto-vacuum rewrite:
     - avoid a pre-checkpoint if it is immediately invalidated by `VACUUM`, unless SQLite requires it for correctness in WAL mode;
     - keep one final checkpoint after metadata updates so WAL shape is stable;
     - keep timestamps accurate.
   - For incremental vacuum:
     - keep the final checkpoint after metadata updates;
     - avoid duplicate checkpoint if no metadata was dirtied.
   - Use comments sparingly to explain why each remaining checkpoint exists.

5. Surface `ranVacuum` diagnostics if useful.
   - Consider service events/status payload: "maintenance completed: checkpoint, incremental vacuum, full vacuum rewrite".
   - At minimum, log a Debug info/warning event when full VACUUM rewrite runs.
   - Do not expand wire protocol unless tests and compatibility are straightforward.

6. Tests.
   - Existing passing cases to preserve:
     - `sqlite_index_store_automatic_checkpoint_truncates_wal`
     - `sqlite_index_store_automatic_compaction_is_bounded`
     - `sqlite_index_store_automatic_rewrites_legacy_auto_vacuum`
   - Add:
     - "legacy auto_vacuum queues maintenance without freelist/WAL pressure";
     - "maintenance busy defers/requeues without clearing queued state" if a deterministic writer lock can be held;
     - "auto-vacuum path performs expected bounded checkpoint count" as a source guard or instrumentation test only if not brittle.

**Verify**:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter sqlite_index_store_automatic_
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_idle_maintenance_queue_and_completion
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_status_reports_maintenance_history
git diff --check
```

Expected: all targeted tests pass. If service-status tests fail with existing environmental service pipe issues, capture the failure and rerun any direct non-service SQLite tests to isolate code correctness.

### Phase C: Fix Cross-FS COPY with Unknown Source Size

Goal: cross-filesystem COPY should not report success when source size is unknown and destination completeness cannot be verified.

1. Clarify `IFileReader::Read` contract.
   - Update `Common/PlugInterfaces/FileSystem.h` to state whether `S_OK` + `*bytesRead == 0` is EOF only.
   - If all plugin implementations already obey EOF-only zero-byte reads, document it and add a contract test.
   - If not guaranteed, bridge code must not trust zero-byte EOF without size verification.

2. Choose policy for unknown-size COPY.
   - Safer default: mirror MOVE behavior. If `reader->GetSize` fails for bridge COPY, return `ERROR_PARTIAL_COPY`, do not commit/promote the temp destination, and preserve source.
   - Alternative: allow COPY only if a post-promote re-stat/readback can prove size/hash. This is more complex and should be a separate design if chosen.
   - Recommended for this remediation: refuse unknown-size COPY in the bridge until a verified-copy contract exists.

3. Implement the guard in `RedSalamander/FolderWindow.FileOperations.State.cpp`.
   - Replace `else if (task._operation == FILESYSTEM_MOVE)` with a branch that rejects both COPY and MOVE when `GetSize` fails, unless the operation has a documented safe verification path.
   - Log diagnostic code should distinguish COPY and MOVE text:
     - MOVE: "preserving source";
     - COPY: "not committing destination because source size is unknown".
   - Ensure temp cleanup deletes uncommitted temp destination.
   - Ensure source remains untouched for MOVE.

4. Add a deterministic selftest.
   - Use the existing `ConsumeBridgeFailNextSourceGetSizeForSelfTest` hook if it can be driven for COPY as well as MOVE.
   - Test name suggestion: `fileops_bridge_copy_unknown_source_size_refuses_commit`.
   - Fixture:
     - create source with known bytes;
     - configure bridge source `GetSize` failure for COPY;
     - run cross-filesystem COPY to destination;
     - expected operation result is `ERROR_PARTIAL_COPY` or a documented partial-copy issue;
     - destination final path does not exist, or if the system creates a placeholder, it is not reported as success and does not contain a truncated committed file;
     - source still exists with original bytes.
   - Add MOVE parity test if not already present:
     - `fileops_bridge_move_unknown_source_size_preserves_source`.

5. Update file-operation spec.
   - `Specs/FileSystem/FileSystem_FileOperations.md` must state:
     - cross-FS COPY and MOVE require a known source size or a stronger post-copy verification;
     - if verification is unavailable, operation must fail partial and preserve source/temp cleanup;
     - users must not see a successful copy for unverified unknown-size input.

**Verify**:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter bridge
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter fileOps
git diff --check
```

Expected: new bridge unknown-size COPY/MOVE tests pass. If broad `fileOps` filter is too large or flaky, run the exact new case names and document why broader coverage was deferred.

### Phase D: Study and Improve Same-FS Overwrite Crash/Abort Safety

Goal: decide and implement whether local same-filesystem overwrite should preserve the pre-existing destination until new bytes are durable, matching the bridge's staging philosophy.

This track has two valid outcomes:

1. Implement staged overwrite with atomic replace; or
2. Document Explorer-compatible truncate-in-place as intentional and add user-visible warning/tests.

The preferred outcome for this plan is staged overwrite unless study finds a blocking compatibility constraint.

1. Study call paths.
   - `Plugins/FileSystem/FileSystem.cpp::CreateFileWriter`
   - `Plugins/FileSystem/FileSystem.FileOps.cpp` local `CopyFileExW` paths
   - `RedSalamander/FolderWindow.FileOperations.State.cpp` bridge staged temp/promotion helpers
   - any rollback helper for pre-existing destinations.

2. Define target semantics.
   - Existing destination must remain intact until replacement bytes are fully written, flushed, and ready to promote.
   - On write error/cancel/source read error, original destination remains intact.
   - On successful overwrite, metadata behavior must be explicit:
     - content replaced;
     - timestamps/attributes either source-like or existing behavior preserved;
     - ACL/ADS behavior documented.
   - Promotion should use `ReplaceFileW` when possible for same-volume file replacement; fallback to `MoveFileExW(...REPLACE_EXISTING...)` only if semantics are acceptable and tested.

3. Pick implementation layer.
   - Option A: change `IFileSystemIO::CreateFileWriter` local implementation so `ALLOW_OVERWRITE` creates a temp sibling and commits with replace.
     - Pro: all host writer users benefit.
     - Con: changes shell-new/template writer semantics if they ever use overwrite.
   - Option B: keep `CreateFileWriter` raw and change host file-operation same-FS overwrite path to stage.
     - Pro: narrower behavior change.
     - Con: direct plugin writer still has footgun.
   - Preferred: Option A only if all current writer callers tolerate temp/commit semantics; otherwise Option B plus documented `IFileWriter` warning.

4. Implement with WIL RAII.
   - Temp file name must be hidden/sibling, collision-resistant, length-safe.
   - Temp cleanup must not delete the original destination.
   - Commit must flush temp before replacement.
   - Abort destructor must delete only the temp path, never a pre-existing destination.
   - Preserve `CREATE_NEW` behavior when overwrite flag is absent.
   - Handle read-only replacement consistently with `FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY`.

5. Tests.
   - Direct writer abort test:
     - pre-existing destination contains `original`;
     - open overwrite writer;
     - write `partial`;
     - release without `Commit`;
     - expected destination still contains `original`.
   - Commit replacement test:
     - same setup;
     - write `new`;
     - `Commit`;
     - expected destination contains `new`.
   - Failure injection test:
     - simulate write error or commit failure if harness supports it;
     - expected original preserved.
   - File-operation test:
     - copy over existing same-FS file and cancel/fail mid-stream;
     - expected original preserved and no temp leak.

6. Spec update.
   - `Specs/FileSystem/FileSystem_FileOperations.md` must state same-FS overwrite safety semantics.
   - If choosing Explorer-compatible truncate-in-place, the spec must explicitly say this is by design and why; otherwise future reviews will keep flagging it.

**Verify**:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter overwrite
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter copy
git diff --check
```

Expected: new overwrite tests pass and no temp files remain in test directories.

STOP if:

- Preserving ACLs/ADS/security descriptors requires a large metadata-copy subsystem. In that case, implement a narrower host-level fail-safe first or split a new plan.
- `ReplaceFileW` fails for common supported targets in a way that cannot be made deterministic.
- The change would break documented plugin writer expectations outside file operations.

### Phase E: Closeout, Specs, and Full Validation

1. Update authoritative specs:
   - `Specs/Core/Core_Search.md` for malformed-entry warning and skip granularity.
   - `Specs/FileSystem/FileSystem_FileOperations.md` for unknown-size bridge COPY/MOVE and same-FS overwrite semantics.
   - `Common/PlugInterfaces/FileSystem.h` comments if `IFileReader::Read` EOF semantics are clarified.

2. Add source guards only where behavioral tests cannot reach the path.
   - Source guards are acceptable for "do not accidentally reintroduce bare return in worker" and "do not remove warning summary mapping".
   - Prefer behavioral tests for file-operation integrity.

3. Run focused validations:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_low_hardening_smoke
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter sqlite_index_store_automatic_
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter bridge
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter overwrite
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter search_local_plugin_invalid_regex_reports_single_completion
git diff --check
```

4. Run broad validation when focused gates are green:

```powershell
.\Tools\Run-AllTests.ps1 -SkipBuild -TimeoutMultiplier 2
```

Expected:

- Full run exits 0, or any failures are unrelated/environmental and documented with archived result paths.
- No lingering `RedSalamander.exe`, `MSBuild`, `cl`, or `link` processes after validation.

5. Closeout rules:
   - Move this WIP plan to `Specs/Plans/Done/` only after all accepted tracks are implemented or intentionally split into follow-up WIP plans.
   - Durable behavior discovered during implementation must be merged into authoritative specs before closeout.
   - Do not leave normative requirements only in this WIP plan.

## Test Matrix

| Area | Required case | Expected proof |
|------|---------------|----------------|
| Search malformed entry | Empty display name followed by valid sibling in same chunk | valid sibling still found; warning bit set; no clean completion without warning |
| Search source guard | `ParallelScanWorkerBody` malformed block | warning + `continue`, not warning-less `return` |
| Find warning UI | malformed-entry warning bit | localized warning appears in `BuildWarningSummary` |
| SQLite legacy queue | non-incremental auto_vacuum, low WAL/freelist | maintenance queues and rewrites mode |
| SQLite busy writer | held writer lock during maintenance attempt | maintenance defers/requeues; no tight loop; no query blocking |
| SQLite checkpoints | auto-vacuum rewrite path | no redundant pre/post/final checkpoint sequence beyond documented required calls |
| Bridge unknown-size COPY | source `GetSize` fails | COPY fails partial; destination not committed; source preserved |
| Bridge unknown-size MOVE | source `GetSize` fails | MOVE fails partial; source preserved |
| Same-FS writer abort | overwrite writer released without commit | original destination remains intact |
| Same-FS writer commit | overwrite writer commits | destination contains new bytes |
| Same-FS operation failure | copy overwrite fails mid-stream | original destination remains intact; temp cleaned |

## Done Criteria

All must be true:

- [x] Search malformed-entry skip paths set a warning and never abandon the rest of a parallel chunk.
- [x] Find Files surfaces the new/incomplete-results warning with localized text.
- [x] SQLite legacy auto-vacuum stores queue maintenance even without unrelated WAL/free-page pressure, or a documented reason exists for manual-only rewrite.
- [x] Automatic maintenance avoids redundant checkpoint work while preserving timestamp and WAL correctness.
- [x] Full VACUUM rewrite is visible in logs/status/tests enough for operators to distinguish it from cheap maintenance.
- [x] Cross-FS COPY does not report success when source size is unknown and no stronger verification was performed.
- [x] Same-FS overwrite behavior is either crash/abort safe or explicitly documented as Explorer-compatible truncate-in-place with tests proving the chosen contract.
- [x] `Specs/Core/Core_Search.md` and `Specs/FileSystem/FileSystem_FileOperations.md` contain the lasting contracts.
- [x] Targeted Compare/FileOps/Commands tests listed above pass.
- [x] `.\build.ps1 -ProjectName RedSalamander` passes with 0 warnings and 0 errors.
- [x] `git diff --check` passes.

## Closeout - 2026-06-26

Implemented and validated:

- Search malformed-entry materialization failures now set `FILESYSTEM_SEARCH_WARNING_OVERFLOW` and parallel chunk workers continue to later entries.
- SQLite persistent store info exposes legacy `auto_vacuum` state; the service queues idle maintenance for non-incremental stores without needing WAL/freelist pressure; automatic maintenance uses one final truncate checkpoint after metadata updates.
- Automatic full VACUUM rewrite is visible through `AutomaticMaintenanceResult::ranVacuum`, service logs, and Compare selftests.
- Cross-filesystem Copy and Move now fail `ERROR_PARTIAL_COPY` when source size is unknown and no stronger verification exists; Copy does not commit the staged destination and Move preserves the source.
- Built-in local overwrite writers stage to a sibling temp file and only replace the destination on `Commit()`; abort deletes only the staged temp file.
- Durable contracts were merged into `Specs/Core/Core_Search.md`, `Specs/FileSystem/FileSystem_FileOperations.md`, and `Common/PlugInterfaces/FileSystem.h`.

Validation:

- `.\build.ps1 -ProjectName RedSalamander` passed with 0 warnings, 0 errors.
- `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_low_hardening_smoke` passed.
- `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter sqlite_index_store_automatic_` passed.
- `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_legacy_auto_vacuum_queues_idle_maintenance` passed.
- `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_idle_maintenance_queue_and_completion` passed.
- `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_status_reports_maintenance_history` passed.
- `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter local_search_scan_wide_tree_parallel_walk_name_only` passed.
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFsCopyGetSizeFailureRefusesCommit` passed.
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFsMoveGetSizeFailurePreservesSource` passed.
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_LocalWriterOverwriteIsStaged` passed.
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Riptide_Bridge` passed.
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Fairstream_Bridge` passed.
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFs` passed.
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Fairstream_Overwrite` passed.
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_LocalWriterOverwrite` passed.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter search_local_plugin_invalid_regex_reports_single_completion` passed.
- Literal `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter bridge` was not a valid harness filter because FileOps filters match exact names/prefixes rather than substrings; the bridge prefix subsets above were used instead.
- Broad `.\Tools\Run-AllTests.ps1 -SkipBuild -TimeoutMultiplier 2` was attempted after focused validation. It exceeded the 20-minute tool timeout while running the unrelated broad Commands UI suite at `cmd_preferences_dialog_hot_paths_live_dx_interaction`; the trace had not advanced since 2026-06-26 11:59:53. Stuck process `RedSalamander.exe --commands-selftest --selftest-timeout-multiplier=2` PID 61220 was recorded and stopped, leaving no RedSalamander/MSBuild/compiler/linker processes running.

Remaining review:

- No product or spec-contract work remains for this FSBSI remediation scope.
- The only open caveat is rerunning broad all-suite validation in a separate pass and investigating the unrelated Commands UI stall if it reproduces. Focused gates for the changed search, SQLite, bridge, overwrite, and regex paths passed.
- Pre-existing untracked audit/review files under `Specs/Reviews/` and the separate `Operation_FSSubsystemDeepAuditRemediation_DataSafetySecurityContracts_2026-06-26.md` plan are outside this closeout and were not included in this spec commit.

## STOP Conditions

Stop and report instead of improvising if:

- A proposed file-operation fix could delete or overwrite a pre-existing destination during a failure path before a test proves preservation.
- A fake test harness cannot reproduce bridge behavior without touching real user files.
- Adding a new search warning bit would break ABI or protocol compatibility and no existing warning bit is acceptable.
- SQLite maintenance serialization requires changing service threading architecture beyond `SearchServiceBroker`/`LocalSearchIndexCore` boundaries.
- Same-FS overwrite preservation requires metadata/ACL/ADS semantics that cannot be preserved in a small change.
- Any validation command hangs past its timeout and leaves child processes running. Stop the child processes only after recording command line, PID, and current test artifact path.

## Maintenance Notes for Reviewers

- Review all file-operation changes from the perspective of failure ordering: source preservation, destination preservation, temp cleanup, and user-visible result must be correct for cancellation, read failure, write failure, commit failure, and crash-like abort.
- Be suspicious of any path that returns `S_OK` after skipping or failing verification.
- For search, warning propagation must use the right accumulator for serial versus parallel paths; do not write shared non-atomic runtime state from worker threads.
- For SQLite, preserve fail-open search behavior. Maintenance failure should not make queries fail if live-scan fallback is available.
- If same-FS overwrite staging lands, future plugin writers should not assume `Commit()` simply marks an already-open destination as committed; it may perform a replacement.
