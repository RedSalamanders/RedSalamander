# DxUi and FileSystem Remaining Improvement Plan

**Last updated:** 2026-04-26

**Status:** Complete - ready for `Specs/Plans/Done/`

**Scope:** Keep only unresolved closeout work for DxUi/FileSystem architecture, validation, and performance. Completed historical work belongs in authoritative specs or Done plans, not in this WIP file.

## Closeout Checklist

### 1. FileSystem copy/move recursive parallelism

- [x] Root-cause the old single-folder bypass and prove `CopyItems(..., count == 1, ...)` used the serial `CopyPathInternal(...)` path before the fix.
- [x] Replace the "top-level item only" model with `CopyPathInternalWithDirectoryParallelism(...)` for `CopyItem(...)`, `CopyItems(...)`, and cross-volume move copy-delete fallback.
- [x] Make recursive workers enumerate nested directories progressively and enqueue file/directory/reparse work under one bounded operation state.
- [x] Add a shared transfer gate so recursive subtree worker allowances can borrow spare scheduler capacity without exceeding the operation's effective `copyMoveMaxConcurrency` file-transfer limit.
- [x] Preserve existing progress streams, issue handling, cancellation, throttling, reparse handling, and `continueOnError` behavior.
- [x] Cover the deterministic matrix:
  - [x] single selected folder through `CopyItems(count == 1)`,
  - [x] direct `CopyItem(...)`,
  - [x] one-child deep folder,
  - [x] three selected folders with one dominant subtree,
  - [x] many shallow files,
  - [x] several child directories with one dominant subtree,
  - [x] mixed folder/file selections,
  - [x] forced `ERROR_NOT_SAME_DEVICE` move fallback,
  - [x] existing reparse skip/copy/follow policy coverage,
  - [x] cancellation while recursive workers are active,
  - [x] partial-copy/`continueOnError` for a locked recursive child.
- [x] Archive focused red/green and matrix evidence in `Specs/TestRuns/4cb089111a23/FileOps/`.
- [x] Update `Plugins/FileSystem/FileSystem.h` and `Specs/FileSystem/FileSystem_FileOperations.md` so the durable contract no longer says `copyMoveMaxConcurrency` is only top-level item concurrency.

### 2. Split remaining broader work to owner plans

- [x] FileSystem stability/search/perf follow-ups moved to `Specs/Plans/WIP/FileSystem_StabilitySearchPerfCloseoutPlan.md`.
- [x] DxUi residual migration/accessibility/perf closeout is completed in `Specs/Plans/Done/UI_DxUiRemainingMigrationCloseoutPlan.md`; historical umbrella detail lives in `Specs/Plans/Done/UI_DxUiSharedGridPlan.md` and `Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md`.
- [x] Callback-drain and shutdown/threading hardening stays owned by `Specs/Plans/WIP/Core_HardeningImprovementPlan_2026-03-19.md`.

### 3. Closeout gates

- [x] Authoritative FileSystem spec updated for copy/move concurrency, cancellation, partial-copy, and fallback behavior.
- [x] Validation evidence linked below.
- [x] No remaining work is stranded only in this umbrella plan.
- [x] Move this file to `Specs/Plans/Done/`.

## Current Code Review: Copy/Move Parallelism Gap

The focused single-folder gap has been fixed in code, tests, specs, and perf telemetry.

- `FileSystem::CopyItem(...)`, `FileSystem::CopyItems(...)`, and cross-volume move fallback now route recursive local-directory copy work through `CopyPathInternalWithDirectoryParallelism(...)`.
- `CopyItems(..., count == 1, ...)` and direct `CopyItem(...)` now use the same recursive directory-parallel path when `copyMoveMaxConcurrency > 1`.
- `CopyDirectoryChildrenParallel(...)` now queues nested directory/file/reparse work so a root with one child directory and many files underneath can still fan out across worker streams.
- `CopyItems(...)` / `MoveItems(...)` now assign recursive subtrees a spare-budget worker allowance and gate actual file transfers through shared operation state so 2 or 3 selected folders can keep heavy subtrees busy without exceeding `copyMoveMaxConcurrency`.
- `Phase7_CopyRecursiveParallelismMatrix` now covers shallow-file folders, dominant child-directory subtrees, mixed folder/file selections, forced move copy-delete fallback, `continueOnError`, and cancellation while recursive workers are active.
- The focused regression `Phase7_CopyItemsSingleFolderRecursiveParallelism` passed and archived evidence under `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_131759/`, then passed again after the shared transfer gate under `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_134413/`.
- The multi-root regression `Phase7_CopyItemsMultiRootUnevenRecursiveParallelism` first failed with one dominant stream and now passes with multiple dominant streams; evidence is archived under `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_133510/` and `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_134219/`.
- The matrix regression first failed on the forced move fallback seam, then passed after adding the Debug-only forced fallback hook; evidence is archived under `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_140013/`, `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_140220/`, and final recheck `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_140937/`.

No remaining work is owned by this plan. Broader follow-ups were split:

- FileSystem stability/search/perf follow-ups: `Specs/Plans/WIP/FileSystem_StabilitySearchPerfCloseoutPlan.md`.
- DxUi migration/accessibility/perf closeout: `Specs/Plans/Done/UI_DxUiRemainingMigrationCloseoutPlan.md`, with historical umbrella detail in `Specs/Plans/Done/UI_DxUiSharedGridPlan.md` and `Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md`.
- Callback-drain and shutdown/threading hardening: `Specs/Plans/WIP/Core_HardeningImprovementPlan_2026-03-19.md`.

## Implementation Shape Landed

- `CopyPathInternalWithDirectoryParallelism(...)` is the shared recursive copy entry used by single-item copy, batch copy, and cross-volume move fallback.
- Recursive copy work is modeled as directory/file/reparse work items under one `ParallelOperationState`.
- Workers enumerate nested directories and enqueue discovered children, while file copies continue to use the existing copy/throttle/progress path.
- Batch copy/move uses `CalculateNestedCopyMoveConcurrency(...)` to keep selected roots active while lending spare worker allowance to recursive subtrees; actual file copies pass through a shared transfer gate tied to the task's effective `copyMoveMaxConcurrency`.
- Debug selftests can force the move copy-delete fallback through `REDSALAMANDER_FILEOPS_FORCE_MOVE_COPY_FALLBACK`; Release builds only enter that fallback after the OS reports `ERROR_NOT_SAME_DEVICE`.
- Issue handling, cancellation, progress stream identity, reparse policy, and `continueOnError` stay on the existing operation context surface.
- Serial fallback telemetry currently covers max concurrency `1` and scheduler unavailable. Add more fallback counters only in the new owner plan if later stress/perf work needs them.

## Validation Plan

- Build:
  - `.\build.ps1 -ProjectName FileSystem`
  - `.\build.ps1 -ProjectName RedSalamander`
- Focused selftest now landed:
  - `Phase7_CopyItemsSingleFolderRecursiveParallelism`
  - Run with: `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase7_CopyItemsSingleFolderRecursiveParallelism --selftest-fail-fast --selftest-timeout-multiplier=2.0`
  - `Phase7_CopyItemsMultiRootUnevenRecursiveParallelism`
  - Run with: `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase7_CopyItemsMultiRootUnevenRecursiveParallelism --selftest-fail-fast --selftest-timeout-multiplier=2.0`
  - `Phase7_CopyRecursiveParallelismMatrix`
  - Run with: `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase7_CopyRecursiveParallelismMatrix --selftest-fail-fast --selftest-timeout-multiplier=2.0`
- Perf archives:
  - Focused red baseline archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_125833/`.
  - Focused green candidate archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_131759/`.
  - Focused post-transfer-gate recheck archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_134413/`.
  - Multi-root red baseline archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_133510/`.
  - Multi-root green candidate archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_134219/`.
  - Matrix red baseline archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_140013/`.
  - Matrix green candidate archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_140220/`.
  - Matrix final recheck archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-04-26_140937/`.

## References

- `Plugins/FileSystem/FileSystem.FileOps.cpp`
  - `CopyDirectoryChildrenParallel(...)`
  - `FileSystem::CopyItem(...)`
  - `FileSystem::CopyItems(...)`
  - `MovePathInternal(...)`
  - `FileSystem::MoveItems(...)`
- `Plugins/FileSystem/FileSystem.h`
  - `copyMoveMaxConcurrency`
  - `FileSystemConcurrencyMode`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/TestRuns/`
