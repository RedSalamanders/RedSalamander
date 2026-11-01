# FolderView WarpDrive Continuation Baton

Created: 2026-07-03 10:00 local time
Branch: codex/folderview-warpdrive
Head: 275c04034
Archive: Specs/TestRuns/4cb089111a23/Continuation/2026-07-03_100049_folderview_warpdrive_current_baton
Previous broad-suite archive: Specs/TestRuns/4cb089111a23/Continuation/2026-07-03_093804_folderview_warpdrive_full_red_after_convergence_baton

## Current Broad Truth

The latest authoritative broad run is still the Full skip-build run from 2026-07-03 08:41-09:33 UTC, archived in the previous baton. It is red:

- Full Debug: 1151 total / 1092 passed / 7 failed / 52 skipped.
- Failing suites: CompareDirectories, Commands, ViewerPETests, RedSalamanderMonitorEtwLatency, ToolsPesterTests.
- Do not move `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` to Done yet.
- Do not mark the active goal complete until Full is green and the final Debug plus test-enabled Release FolderView perf matrices are rerun and cited.

## Fixed Or Advanced After The Red Full Run

1. ToolsPester / ResourceLocalizationContracts:
   - Root cause: `Tools/Tests/ResourceLocalizationContracts.Tests.ps1` searched broad top-level repo directories and let archived `Specs/TestRuns` perturb resource validation.
   - Patch: `Get-RSResourceFiles` now scopes resource validation to actual `.rc` roots: `Common`, `Plugins`, `PoC`, `RedConfigure`, `RedSalamander`, `RedSalamanderMonitor`, and `Tests`, plus root `.rc` files.
   - Focused resource localization: `Passed: 4 Failed: 0`.
   - Full Tools Pester direct: `Passed: 115 Failed: 0 Skipped: 0` in `.build/codex-runs/tools_pester_after_resource_root_scope_20260703.log`.

2. CompareDirectories `local_search_service_single_request_uses_query`:
   - Red Full reason: `Service should still accept one more request after the search. hr=0x80070002`.
   - Root cause lead: the service request budget was too tight for the modern service-backed search path/status/query/status traffic.
   - Patch: `service.Start(pipeName, 6u, ...)` changed to `12u` for this selftest in `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`.
   - Rebuild: `.build/codex-runs/msbuild_redsalamander_after_compare_budget_20260703.log`, diagnostics `0 warning(s), 0 error(s)`.
   - Focused rerun: `C:/Users/eric/AppData/Local/RedSalamander/SelfTest/last_run/compare/results.json` showed `passed=1 failed=0`, duration 1759ms.

3. Commands focused evidence:
   - The latest volatile Commands `last_run` is a focused pass for `cmd_pane_navigation_menu_keeps_navigation_shell_stable`: `passed=1 failed=0`, duration 480ms, case filter recorded correctly.
   - This does not clear broad Commands. It only proves that one Full failure can pass in isolation.

## Remaining Gates To Converge

1. Commands broad-suite failures from the last Full run must be reproduced/fixed or proven flaky with stronger guards:
   - `cmd_pane_find_dialog_restored_combined_view_state_grid_doubleClick_activates_selection_after_sort_cycles_and_search_rerun`
   - `cmd_pane_navigation_menu_keeps_navigation_shell_stable`
   - `cmd_app_menuBar_persistent_mouseleave_clears_hover_without_live_cursor_switch`

2. ViewerPETests still needs the Full-ordering HUD focus/readiness fix:
   - Last Full: `ViewerVLC exposes a visible HUD child that can take focus`; shell churn also fails because the nested ViewerVLC gate fails.
   - Earlier direct/focused ViewerPE evidence had been green, so reproduce under Full-like ordering before patching.

3. RedSalamanderMonitorEtwLatency still needs a post-current rerun:
   - Last Full failed with exit code 1 and little output.
   - Earlier direct command had been green before later changes; rerun the exact command before patching.

4. Full suite must be rerun after the above:
   - `./Tools/Run-AllTests.ps1 -Suite Full -SkipBuild`
   - Use only one Full run at a time; it overwrites `SelfTest/last_run`.

5. Final FolderView closeout still requires:
   - Debug FolderView perf matrix.
   - Test-enabled Release FolderView perf matrix.
   - `Tools/Show-PerfRuns.ps1 -FailOnQuality -ShowBuildFlavor` quality pass.
   - Update durable specs under `Specs/Testing/*` / FolderView domain specs as needed.
   - Move the WIP plan to `Specs/Plans/Done/` only after the above is green and cited.

## Crash Evidence Preserved

RedSalamander crash sidecars are under `C:/Users/eric/AppData/Local/RedSalamander/Crashes`; this archive copied recent `.txt` sidecars and the latest crash manifest into `crash-sidecars/`.

Important latest stack already preserved:

- `RedSalamander-20260702-100018-p73832.txt`
- Exception: `0xC0000005`
- Stack: `WideCharToMultiByte -> Debug::detail::ConvertUtf8 -> WritePerfJsonl -> Debug::Perf::Scope::~Scope -> FolderView::ProcessThumbnailLoadQueue -> FolderView::EnumerationWorker`
- Root-cause lead from the crash skill: suspect `Debug::Perf::Scope` holding a `std::wstring_view` to a temporary metric/detail string that dies before JSONL write. Check calls like `perf.SetDetail(path.parent_path().native())` or `perf.SetDetail(std::format(...))`; the durable fix is to make the perf scope own strings when sinks are enabled or extend producer lifetime through scope destruction.

## Suggested Next Commands

Focused Commands repros, one at a time:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_find_dialog_restored_combined_view_state_grid_doubleClick_activates_selection_after_sort_cycles_and_search_rerun --selftest-timeout-multiplier=2 *> .build\codex-runs\commands_find_grid_rerun_repro_20260703.log
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigation_menu_keeps_navigation_shell_stable --selftest-timeout-multiplier=2 *> .build\codex-runs\commands_navigation_menu_repro2_20260703.log
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_app_menuBar_persistent_mouseleave_clears_hover_without_live_cursor_switch --selftest-timeout-multiplier=2 *> .build\codex-runs\commands_menubar_mouseleave_repro_20260703.log
```

Viewer/monitor rechecks:

```powershell
.\.build\x64\Debug\ViewerPETests.exe *> .build\codex-runs\viewerpe_direct_post_current_20260703.log
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --wait-instance --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260 *> .build\codex-runs\monitor_etw_latency_post_current_20260703.log
```

After focused gates are green, rerun Full:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild *> .build\codex-runs\full_skipbuild_after_current_gates_20260703.log
```

## Contents

- `git-status-short.txt`, branch/head, ahead/behind, diff stats, name-status snapshots.
- `patches/staged-code-and-spec-no-testruns.patch` and `patches/unstaged-code-and-spec-no-testruns.patch`.
- `source-snapshots/` for the active plan, crash skill, resource Pester fix, compare selftest fix, runner/filter sources, ViewerPE test, and testing docs.
- `run-artifacts/` for current build/focused logs plus the red Full log.
- `last-run/` small current results/traces only; locked `.dmp` and massive scratch `work/` payloads are intentionally excluded.
- `crash-sidecars/` recent RedSalamander crash `.txt` call stacks and manifest.
