# FolderView WarpDrive Continuation Baton - 2026-07-03 09:38

## Current State

- Branch: `codex/folderview-warpdrive`
- HEAD: `275c04034`
- Origin sync: `0 25` from `git rev-list --left-right --count origin/master...HEAD`; this worktree is 25 commits ahead of `origin/master`.
- Active WIP plan: `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- Goal is not done. Do not move the plan to Done until Full is green and the final Debug plus test-enabled Release FolderView perf matrices are rerun and cited.

## What Changed Since The Previous Baton

Previous baton:

- `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_084000_folderview_warpdrive_gate_convergence_baton\CONTINUATION.md`

New confirmations:

- Full `Tools\Tests` Pester was rerun directly after inventory count updates and is green: `Passed: 115 Failed: 0 Skipped: 0`.
- Fresh Debug build is green: `.\build.ps1 -Configuration Debug`, 0 warnings / 0 errors.
- Full suite was rerun from that build and is narrower but still red: `1151 total / 1092 passed / 7 failed / 52 skipped`.
- Old blockers now green inside Full: FileOperations, RedConfigureTests, MonitorTest, PerformanceTests2, and the previous Compare live-progress/source-guard slice except one remaining Compare case.

## Latest Full Run

Command:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

Primary log:

- `.build\codex-runs\full_skipbuild_after_red_gates_20260703.out.log`

Structured results:

- `C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run\run-all-tests-results.json`
- Archived copy: `last-run\run-all-tests-results.json`

Summary:

- CompareDirectories: `206 passed / 1 failed / 30 skipped`
- Commands: `773 passed / 3 failed / 2 skipped`
- FileOperations: green
- ViewerPETests: red
- RedSalamanderMonitorEtwLatency: red
- ToolsPesterTests: red
- PerformanceTests2: green

## Exact Remaining Failures

Compare:

- `local_search_service_single_request_uses_query`
- Reason: `Service should still accept one more request after the search. hr=0x80070002`
- Source anchor: `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`

Commands:

- `cmd_pane_find_dialog_restored_combined_view_state_grid_doubleClick_activates_selection_after_sort_cycles_and_search_rerun`
- Reason: reopened Find window did not re-expose the expected selected combined results-grid state after sort cycles and search rerun.
- Source anchor: `RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp`

- `cmd_pane_navigation_menu_keeps_navigation_shell_stable`
- Reason: `Pane Menu did not activate a live DxUI main-menu session.`
- Source anchor: `RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp`

- `cmd_app_menuBar_persistent_mouseleave_clears_hover_without_live_cursor_switch`
- Reason: `WM_MOUSELEAVE must not switch the active top-level menu from the later live cursor (selectedBeforeHover=5, selectedAfterWait=4, expectedView=5, plugins=4, suppressedPoint=(549,253), suppressedHit=5, pointerSwitches=0).`
- Source anchor: `RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp`

ViewerPETests:

- Suite exits 1.
- Repeated observed failure: `ViewerVLC exposes a visible HUD child that can take focus`.
- The earlier direct ViewerPE run passed after targeting `RedSalamander.ViewerVLC.Video`, so this Full failure likely needs a second readiness/focus race investigation around the HUD child in shell churn.
- Source anchor: `Tests\ViewerPETests\ViewerPETests.cpp`

RedSalamanderMonitorEtwLatency:

- Full runner reports exit 1 with empty output log.
- Earlier exact direct command was green with empty output; rerun the exact command after this Full run before changing code.

ToolsPesterTests:

- Full Pester failed 2 resource-localization tests.
- Both failures are `DirectoryNotFoundException` while `Tools\Tests\ResourceLocalizationContracts.Tests.ps1` scans under `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_091418_commands_pack_prompt_crash\last_run\commands\w\...`.
- Direct full Tools Pester was green earlier in this slice, so reproduce this exact file before patching. Likely fix: constrain resource discovery to source/resource roots and exclude archived `Specs\TestRuns`.

## Archived Artifacts

State and diffs:

- `git-status-short.txt`
- `git-branch.txt`
- `git-head.txt`
- `origin-master-ahead-behind.txt`
- `git-diff-name-status-no-testruns.txt`
- `git-diff-stat-no-testruns.txt`
- `git-diff-cached-name-status-no-testruns.txt`
- `git-diff-cached-stat-no-testruns.txt`
- `patches\unstaged-code-and-spec-no-testruns.patch`
- `patches\staged-code-and-spec-no-testruns.patch`

Run logs:

- `run-artifacts\tools_pester_after_inventory_counts_20260703.log`
- `run-artifacts\msbuild_debug_after_viewerpe_redconfigure_pester_20260703.log`
- `run-artifacts\full_skipbuild_after_red_gates_20260703.out.log`
- `run-artifacts\full_skipbuild_after_red_gates_20260703.err.log`
- plus all earlier `20260703` focused logs listed in `run-artifacts\codex-runs-20260703-list.txt`.

Last-run essentials:

- `last-run\run-all-tests-results.json`
- `last-run\commands\results.json`
- `last-run\commands\trace.txt`
- `last-run\compare\results.json`
- `last-run\compare\trace.txt`
- `last-run\fileops\results.json`
- `last-run\fileops\trace.txt`
- `last-run\perf\*.json`
- top-level suite output logs from `last_run`

Source snapshots:

- `source-snapshots\Specs_Plans_WIP_Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- `source-snapshots\RedSalamander_SelfTest_CompareDirectories_CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- `source-snapshots\RedSalamander_SelfTest_Commands_Commands.SelfTest.Search.cpp`
- `source-snapshots\RedSalamander_SelfTest_Commands_Commands.SelfTest.Navigation.cpp`
- `source-snapshots\RedSalamander_SelfTest_Commands_Commands.SelfTest.ViewCommands.cpp`
- `source-snapshots\Tests_ViewerPETests_ViewerPETests.cpp`
- `source-snapshots\Tools_Tests_ResourceLocalizationContracts.Tests.ps1`
- `source-snapshots\Tools_Tests_TestInventory.Tests.ps1`
- `source-snapshots\Specs_Testing_Testing_TestCoverage.md`
- `source-snapshots\Specs_Testing_Testing_PerformanceValidation.md`
- `source-snapshots\Tests_README.md`
- `source-snapshots\RedConfigure_Workspace_WorkspaceDiscovery.cpp`
- `source-snapshots\.agents_skills_red-salamander-crash-forensics_SKILL.md`

Crash/debug context:

- `crash-sidecars\red-salamander-crashes-latest.txt`
- latest RedSalamander crash sidecar `.txt` files from `C:\Users\eric\AppData\Local\RedSalamander\Crashes`

## Continue From Here

1. Reproduce each remaining red case focused before patching.
2. Compare: inspect the single-request search-service helper and prove why the final request returns `ERROR_FILE_NOT_FOUND`.
3. Commands: isolate Find restore state, Pane Menu session activation, and persistent menu `WM_MOUSELEAVE` as separate focused runs.
4. ViewerPE: repeat `TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly` under shell churn/full ordering; inspect HUD visibility/enabled/focusability diagnostics before changing the test.
5. Monitor ETW: rerun the exact Full command directly after the red Full run and capture exit code.
6. Tools Pester: reproduce `ResourceLocalizationContracts.Tests.ps1`; likely constrain file discovery away from `Specs\TestRuns`.
7. Rebuild Debug, rerun focused gates, then rerun Full.
8. Only after Full is green, rerun final Debug and test-enabled Release FolderView perf matrices, update the WIP plan with archive paths, merge durable contracts into authoritative docs, and move the plan to Done.

