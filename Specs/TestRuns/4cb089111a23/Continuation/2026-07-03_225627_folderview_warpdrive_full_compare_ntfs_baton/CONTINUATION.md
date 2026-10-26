# FolderView Warp Drive Continuation Baton

Created: 2026-07-03 22:56 local time
Archive: `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_225627_folderview_warpdrive_full_compare_ntfs_baton\`

## Current State

DxUi and ViewerPE are no longer the active blockers. The latest proper Full run
after those fixes exited 1 only because CompareDirectories failed
`search_service_sqlite_ntfs_traversal_seed_stays_degraded`.

Green evidence captured here:

- `run-artifacts\msbuild_dxuitests_menu_arm_pointer_probe_20260704.log`: DxUiTests build after the pointer-root guard, exit 0.
- `run-artifacts\dxui_menu_suite_after_arm_pointer_probe_20260704.log`: focused Menu suite, exit 0.
- `run-artifacts\msbuild_dxuitests_menu_cleanup_20260704.log`: rebuild after removing temporary menu diagnostics, exit 0.
- `run-artifacts\dxui_menu_suite_after_cleanup_20260704.log`: focused Menu suite after cleanup, exit 0.
- `run-artifacts\dxui_direct_broad_after_menu_fix_20260704.log`: direct broad DxUiTests, exit 0.
- `run-artifacts\viewerpe_vlc_tab_focus_after_dxui_green_20260704.log`: focused VLC HUD test rerun, green archive-backed output.
- `run-artifacts\viewerpe_direct_broad_after_dxui_green_20260704.log`: direct broad ViewerPETests, exit 0.

Latest Full evidence captured here:

- `run-artifacts\full_after_dxui_viewerpe_green_20260704.out.log`
- `run-artifacts\full_after_dxui_viewerpe_green_20260704.err.log`
- `run-artifacts\full_after_dxui_viewerpe_green_20260704.exit.txt`
- `last-run-top-level\run-all-tests-results.json`

Full result: 1099 passed / 1 failed / 52 skipped / 1152 total. The single
failing suite is CompareDirectories, and the failing case is
`search_service_sqlite_ntfs_traversal_seed_stays_degraded` with reason
`Forced NTFS traversal-seed service did not finish startup warmup.`

Focused rerun after the Full red:

- Archive: `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-03_225200`
- Result: passed 1 / failed 0 / skipped 0, duration about 10.3s.

## Working Root Cause Thread

The failing case is in
`RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
around `search_service_sqlite_ntfs_traversal_seed_stays_degraded`.

It waits up to `SelfTest::ScaleTimeout(15'000)` for:

- SQLite persistent store.
- At least one startup-warmup completed root.
- At least one indexed volume.

In Full with timeout multiplier 2 this still timed out at about 30.3s; focused,
the same case passed in about 10.3s.

The immediately prior cold-start case sets
`REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS=5000` and restores it by
scope exit. The next investigation should prove whether the Full-only red is an
order/load timeout, stale environment interaction, service status-poll limit, or
a real warmup starvation.

## Resume Plan

1. Inspect selftest filtering and `ForegroundSearchServiceProcess` request semantics.
2. Try to reproduce the prior-case-plus-target sequence if multi-case filtering supports it; otherwise rerun broad CompareDirectories before changing code.
3. If still red, add focused diagnostics to the target failure message/status capture: persistent store kind, completed/total/failed warmup roots, warmup running/current root, indexed volume count, indexed entry count, last `GetStatus` HRESULT, sync phase, ready-for-query cutover, and relevant environment values.
4. If evidence says this is only a Full-order slow warmup, make the target wait robust with a better timeout/diagnostic gate and rerun focused Compare, broad CompareDirectories, then Full.
5. Do not move the plan to Done until Full is green and the final Debug plus test-enabled Release FolderView perf matrices are rerun from current HEAD and cited.

## Included Snapshots

- `patches\staged-code-and-spec-no-testruns.patch`
- `patches\unstaged-code-and-spec-no-testruns.patch`
- `source-snapshots\Specs_Plans_WIP_Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- `source-snapshots\Tests_DxUiTests_DxUiTests.Menu.cpp`
- `source-snapshots\RedSalamander_SelfTest_CompareDirectories_CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- `source-snapshots\Common_SearchServiceBroker.cpp`
- `source-snapshots\Tools_Run-AllTests.ps1`

## Why This Is Not Done Yet

The implementation is converging, but the plan cannot move to Done until:

- The CompareDirectories Full-order startup-warmup red is either fixed or proven
  transient with a green broad CompareDirectories plus proper Full rerun.
- `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` is green.
- The final Debug and test-enabled Release FolderView perf matrices are rerun
  from current HEAD and cited in the WIP plan.
