# FolderView Warp Drive Continuation Baton

Created: 2026-07-03 23:58 local time
Archive: `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_235801_folderview_warpdrive_full_red_compare_commands_baton\`

## Current State

This supersedes the `2026-07-03_225627_folderview_warpdrive_full_compare_ntfs_baton`
archive. The new Full run used the CompareDirectories status-diagnostic build and
exited red with two Full-order failures.

Latest Full evidence:

- `run-artifacts\full_after_compare_status_diag_20260704.log`
- `run-artifacts\full_after_compare_status_diag_20260704.exit.txt`
- `last-run-top-level\run-all-tests-results.json`

Full result:

- Overall: 1098 passed / 2 failed / 52 skipped / 1152 total.
- CompareDirectories: 207 passed / 1 failed / 30 skipped.
- Commands: 775 passed / 1 failed / 2 skipped.
- FileOperations, DxUiTests, ViewerPETests, Tools Pester, and downstream native/script suites passed in this Full run.

## Failure 1: CompareDirectories

Case:

`search_service_sqlite_ntfs_traversal_seed_stays_degraded`

Reason, now with the added diagnostic:

`Forced NTFS traversal-seed service did not finish startup warmup. Last status:
pid=101408 storeKind=2 warmupEnabled=true running=false completed=0 failed=1
total=1 currentRoot='' indexedVolumes=0 indexedEntries=0 ready=true
inspected=true ... storeState=2 syncPhase=2 ... discoveredRoots=1
lastFailure=true failureHr=0x80070002 ... lastGetStatusHr=0x80070002.`

Interpretation to prove next:

- The service did not merely run slowly. Startup warmup ended with one failed
  root, zero completed roots, and `ERROR_FILE_NOT_FOUND`.
- Focused and broad CompareDirectories still pass with the same diagnostic
  build, so this is Full-order/load/context sensitive.

Recent green evidence:

- Focused target after diagnostic build:
  `run-artifacts\runall_compare_ntfs_seed_focused_status_diag_20260704.log`
  (1 passed / 0 failed / 0 skipped).
- Broad CompareDirectories after diagnostic build:
  `run-artifacts\runall_compare_broad_status_diag_20260704.log`
  and `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-03_231441`
  (208 passed / 0 failed / 30 skipped).
- The two-case suspected predecessor recipe also passed:
  `run-artifacts\compare_coldstart_plus_ntfs_seed_after_full_red_20260704.log`
  and `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-03_225949`.

Next Compare move:

Inspect the `ERROR_FILE_NOT_FOUND` path in startup warmup. Start with
`Common\SearchServiceBroker.cpp`, `LocalSearchIndexCore::Repository::EnsureReadyForRoot`,
and the forced `REDSALAMANDER_TEST_FORCE_NTFS_TRAVERSAL_SEED` path. Do not
replace the wait with a longer timeout; the diagnostic proves the warmup failed,
not that it was still running.

## Failure 2: Commands

Case:

`cmd_pane_navigation_context_menu_keeps_navigation_shell_stable`

Reason:

`Pane Context Menu should close back to the baseline navigation shell without
churn; focusTarget=0 ... focusedFolderView=0x0, expectedFolderView=0xA2114C0,
focusHwnd=0x0, workerStarted=yes, sessionObserved=yes, popupOwnerValid=yes,
sessionClosed=yes.`

Interpretation to prove next:

- Path, item count, selected count, focused item, worker state, popup owner, and
  context-menu session close all match expectations.
- The remaining failure is focus restoration: after the context menu closes,
  Win32 focus is null instead of the expected folder view.

Next Commands move:

Run the case focused and in the likely nearby Navigation prefix before patching.
Start in `RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp`
around the `pane_context_menu_nav_shell` helper and the product path that closes
pane context menus. Prove whether this is selftest focus setup, menu-close focus
handoff, or another Full-order focus leak.

## Included Snapshots

- Current staged and unstaged no-TestRuns patches under `patches\`.
- Source snapshots for:
  - `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
  - `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
  - `RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp`
  - `Common\SearchServiceBroker.cpp`
  - `Tools\Run-AllTests.ps1`
  - `Tools\TestRunPlan.ps1`
  - `Tests\DxUiTests\DxUiTests.Menu.cpp`
- Latest `SelfTest\last_run` top-level files.
- Latest Compare/Commands suite archives copied under `test-archive-summaries\`.
- Crash sidecar and WER dump manifests under `crash-context\`.
- Process snapshot records a live Release `RedSalamander.exe` window at archive
  time; it was not killed.

## Not Done Yet

Do not move the WIP plan to Done until both Full-order red cases are fixed or
otherwise proven resolved by a green proper Full run, then the final Debug plus
test-enabled Release FolderView perf matrices are rerun from current HEAD and
cited.
