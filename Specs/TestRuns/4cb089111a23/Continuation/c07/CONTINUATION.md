# FolderView WarpDrive Continuation Baton - 2026-07-03 08:35

## Current State

- Branch: `codex/folderview-warpdrive`
- HEAD: `275c04034`
- Origin sync: `25 0` from `git rev-list --left-right --count HEAD...origin/master`; no rebase was needed.
- Active WIP plan: `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- This is still Task 11 closeout work. Do not move the plan to Done until Full is green and final Debug plus test-enabled Release FolderView perf matrices are rerun and cited from the current HEAD.

## What Changed In This Slice

Only the current focused Compare red-gate source file was edited in this slice:

- `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`

Changes:

- Made `search_source_allocation_and_folding_guard` source-text invariants whitespace-insensitive for the `CanAppendClientBufferedCandidates(...)` before `outCandidates.insert(...)` ordering check.
- Made `search_low_hardening_smoke` source-text invariant whitespace-insensitive for `result.ranVacuum = true`.
- Isolated `search_service_query_reports_live_progress` by giving the foreground service test a per-case `--storage-root` and `--store-backend=snapshot`.
- Raised that service test's request allowance from `2` to `3` because readiness/status, query, and final status each consume a request.
- Added a final `SearchServiceBroker::GetStatus(...)` before waiting for the foreground service to exit, so shutdown is deterministic instead of racing a pipe instance boundary.

The focused patch is archived at:

- `patches\compare-red-gates-fix.patch`

## Verification Performed

Debug build after source guard hardening:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug *> .build\codex-runs\msbuild_after_compare_source_guard_compact_20260703.log
```

Result: exit 0, build succeeded.

Initial focused Compare repro:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'search_source_allocation_and_folding_guard,search_low_hardening_smoke,search_service_sqlite_cold_start_stale_root_refreshes_before_query,search_service_query_reports_live_progress' -TimeoutMultiplier 2 *> .build\codex-runs\compare_red_gates_repro_20260703.log
```

Result: exit 1, `1 passed / 3 failed / 0 skipped`; stale-root was already green.

After source guard hardening:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'search_source_allocation_and_folding_guard,search_low_hardening_smoke,search_service_sqlite_cold_start_stale_root_refreshes_before_query,search_service_query_reports_live_progress' -TimeoutMultiplier 2 *> .build\codex-runs\compare_red_gates_after_source_guard_compact_20260703.log
```

Result: exit 1, `3 passed / 1 failed`; only `search_service_query_reports_live_progress` remained red.

Focused live-progress repro:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'search_service_query_reports_live_progress' -TimeoutMultiplier 2 *> .build\codex-runs\compare_live_progress_focused_repro_20260703.log
```

Result: exit 1, `0 passed / 1 failed`; failure varied between `0x80070002` and `0x800700E7`, consistent with service instance / pipe readiness race and machine-global storage use.

Debug build after isolating live-progress:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug *> .build\codex-runs\msbuild_after_live_progress_isolated_service_20260703.log
```

Result: exit 0, build succeeded.

Focused live-progress after fix:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'search_service_query_reports_live_progress' -TimeoutMultiplier 2 *> .build\codex-runs\compare_live_progress_after_isolated_service_20260703.log
```

Result: exit 0, `1 passed / 0 failed / 0 skipped`.

Original four Compare red gates after all fixes:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'search_source_allocation_and_folding_guard,search_low_hardening_smoke,search_service_sqlite_cold_start_stale_root_refreshes_before_query,search_service_query_reports_live_progress' -TimeoutMultiplier 2 *> .build\codex-runs\compare_red_gates_green_after_live_progress_20260703.log
```

Result: exit 0, `4 passed / 0 failed / 0 skipped`.

ViewerPE direct repro:

```powershell
.\.build\x64\Debug\ViewerPETests.exe *> .build\codex-runs\viewerpe_direct_repro_20260703.log
```

Result: exit 1, about 47s.

Important ViewerPE observation:

- The old Full-reported ViewerVLC HUD/focus concern did not reproduce in the direct run.
- The log repeatedly reports `ViewerPETests passed.` inside fresh child runs.
- The tail shows `TestViewerTextDiffModesAndPlaceholders` fresh harness process exits, reports exit code, and passes.
- Immediately after that, the parent still prints `ViewerPETests failed.`
- This looks like a parent aggregation / recursive fresh-process harness issue rather than a failing viewer case. Start by inspecting `RunFullSuiteInFreshProcesses()`, `RunViewerTestInFreshProcess(...)`, and the outer `success` aggregation in `Tests\ViewerPETests\ViewerPETests.cpp`.

Archived summary files:

- `viewerpe-direct-summary.txt`
- `viewerpe-direct-tail.txt`
- `run-artifacts\viewerpe_direct_repro_20260703.log`

## Archived Artifacts

Run logs copied from `.build\codex-runs`:

- `run-artifacts\compare_red_gates_repro_20260703.log`
- `run-artifacts\msbuild_after_compare_source_guard_compact_20260703.log`
- `run-artifacts\compare_red_gates_after_source_guard_compact_20260703.log`
- `run-artifacts\compare_live_progress_focused_repro_20260703.log`
- `run-artifacts\msbuild_after_live_progress_isolated_service_20260703.log`
- `run-artifacts\compare_live_progress_after_isolated_service_20260703.log`
- `run-artifacts\compare_red_gates_green_after_live_progress_20260703.log`
- `run-artifacts\viewerpe_direct_repro_20260703.log`

State snapshots:

- `git-status-short.txt`
- `git-head.txt`
- `origin-master-ahead-behind.txt`
- `git-diff-name-status-no-testruns.txt`
- `git-diff-stat-no-testruns.txt`
- `source-snapshots\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- `source-snapshots\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.snapshot.md`
- `inventory\debug-test-exes.txt`

## Remaining Gates

Continue in this order:

1. Fix or explain the `ViewerPETests.exe` parent failure where all visible inner/fresh runs pass but the parent exits 1.
2. Resolve the missing `.build\x64\Debug\RedConfigureTests.exe`. Current Debug test exe inventory contains `CrashHandlingTests.exe`, `DxUiTests.exe`, `FileSystemCurlTests.exe`, `LocalizationTests.exe`, `PluginContractTests.exe`, `SettingsSchemaTests.exe`, `ViewerPETests.exe`, and `ViewerSqliteTests.exe`; no `RedConfigureTests.exe`.
3. Re-run or isolate `RedSalamanderMonitorEtwLatency` from the Full red baton.
4. Re-run or isolate `ToolsPesterTests` from the Full red baton.
5. Re-run `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2`.
6. If Full is green, rerun the final FolderView perf matrix in Debug and test-enabled Release from current HEAD.
7. Update durable specs, cite final archives, and only then move the WIP plan to `Specs\Plans\Done`.

## Exact Resume Commands

ViewerPE next:

```powershell
rg -n "RunFullSuiteInFreshProcesses|RunViewerTestInFreshProcess|fresh harness run passes|TestViewerTextDiffModesAndPlaceholders|ViewerPETests failed" Tests\ViewerPETests\ViewerPETests.cpp
.\.build\x64\Debug\ViewerPETests.exe *> .build\codex-runs\viewerpe_direct_repro_retry_20260703.log
```

RedConfigure inventory next:

```powershell
rg -n "RedConfigureTests|RedConfigure.*Tests|Tests.exe" . -g "*.sln" -g "*.vcxproj" -g "*.ps1" -g "*.md"
Get-ChildItem .build\x64\Debug -Filter "*Tests.exe" | Sort-Object Name
```

Full gate after focused gates are green:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2 *> .build\codex-runs\full_skipbuild_after_compare_viewerpe_red_gates_20260703.log
```

