# FolderView WarpDrive Continuation Baton

Created: 2026-07-03 13:23 Europe/Paris
Worktree: `Z:\src\RedSalamander`
Branch: `codex/folderview-warpdrive`
HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`

## Current State

The FolderView WarpDrive plan is converging, but it is not ready to move to Done.

Commands is broad-green after the latest selftest hardening:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`
  - `cmd_preferences_dialog_category_tree_boundary_navigation_from_scrolled_state` now waits for the category tree focus/selection/title snapshot after boundary navigation.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`
  - `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable` now waits for the status bar/focused-item snapshot and refocuses the folder view before sending navigation keys.
- WIP plan evidence was updated in `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`.

Latest green gates:

- Debug RedSalamander build: `.build/codex-runs/msbuild_redsalamander_after_nav_pref_waits_fix_20260703.log`
- Focused Preferences boundary: `.build/codex-runs/runall_commands_pref_boundary_focused_after_nav_pref_waits_20260703.log`
- Focused Navigation status bar: `.build/codex-runs/runall_commands_nav_status_focused_after_nav_pref_waits_20260703.log`
- Navigation 11-case window: `.build/codex-runs/runall_commands_nav_window11_after_nav_pref_waits_20260703.log`
- Preferences 11-case window: `.build/codex-runs/runall_commands_pref_boundary_window_after_nav_pref_waits_20260703.log`
- Broad Commands: `.build/codex-runs/commands_broad_after_nav_pref_waits_20260703.log`
  - Archived run: `Specs/TestRuns/4cb089111a23/Commands/2026-07-03_123938`
  - Result: 776 passed, 0 failed, 2 skipped.
- ViewerPE direct: `.build/codex-runs/viewerpe_direct_after_commands_green_20260703.log`
- Monitor ETW direct: `.build/codex-runs/monitor_etw_latency_after_commands_green_20260703.log`
- Full Debug build before Full: `.build/codex-runs/msbuild_all_debug_before_full_20260703.log`

Latest Full run:

- Command: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2`
- Log: `.build/codex-runs/full_skipbuild_after_commands_viewerpe_monitor_green_20260703.log`
- Result: failed, 1097 passed, 2 failed, 52 skipped.

## Remaining Failures

1. CompareDirectories failed in Full:

- Archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-03_124734`
- Case: `decision_cache_eviction_budget_pending_wide_tree`
- Reason: `pending budget: failed to drain pending subtree updates.`
- Results: 206 passed, 1 failed, 30 skipped.
- Important: this may be order/load-sensitive; do not assume the case is broken until focused and predecessor-window runs are checked.

2. Monitor ETW latency failed only under the Full runner:

- Archive: `Specs/TestRuns/4cb089111a23/Monitor/2026-07-03_131854`
- Summary: `Monitor DxUI chrome selftest requires ENABLE_TESTS.`
- Missing metrics include `monitor.frame.append_to_visible_us`, `monitor.etw.batch_drain_us`, `monitor.etw.selftest_burst_drain_us`, and queue metrics.
- Direct Monitor ETW latency command passed earlier, so the likely issue is the Full runner/build/test-hook contract, not the monitor scenario itself.

## Next Commands

Start with narrow evidence, then widen only after the focused result is understood:

```powershell
rg -n "decision_cache_eviction_budget_pending_wide_tree|pending budget: failed to drain pending subtree updates|failed to drain pending subtree" RedSalamander Tests Tools
.\Tools\Run-AllTests.ps1 -Suite CompareDirectories -SkipBuild -CaseFilter decision_cache_eviction_budget_pending_wide_tree -TimeoutMultiplier 2 *> .build\codex-runs\compare_pending_wide_tree_focused_after_full_red_20260703.log
```

If focused passes, run the nearby CompareDirectories predecessor window around the decision-cache eviction budget cases before editing product code.

For the Monitor Full-run failure:

```powershell
rg -n "RedSalamanderMonitorEtwLatency|monitor-etw|chrome-selftest|RSBuildEnableTests|ENABLE_TESTS" Tools RedSalamander RedSalamanderMonitor* *.vcxproj
```

Check why the Full runner produced `requires ENABLE_TESTS` while the direct Monitor ETW latency command passed. Preserve the direct command contract if the fix is in `Tools/Run-AllTests.ps1`.

## Closeout Criteria

Do not move the WIP plan to Done yet. Remaining gates:

1. Fix or explain the CompareDirectories Full failure with focused and window evidence.
2. Fix the Monitor ETW Full-run `ENABLE_TESTS`/metric contract.
3. Rebuild Debug all.
4. Rerun direct ViewerPE and direct Monitor ETW gates if touched.
5. Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2` and get green.
6. Run the final Debug and test-enabled Release FolderView perf matrices required by the WarpDrive plan.
7. Run `Tools/Show-PerfRuns.ps1 -FailOnQuality -ShowBuildFlavor`.
8. Update the WIP plan with final archive paths and only then move it to `Specs/Plans/Done/`.
9. Migrate durable behavior/contracts into authoritative `Specs/*` or repo guidance before closeout.

## Archive Contents

This folder contains:

- `run-artifacts/` with the latest focused, broad, direct, build, Full, CompareDirectories, and Monitor result artifacts.
- `source-snapshots/` with the current WIP plan, latest touched selftests, the CompareDirectories case file, `Tools/Run-AllTests.ps1`, and the RedSalamander crash-forensics skill.
- `patches/` with staged and unstaged no-TestRuns diffs.
- Git status, branch/head, origin ahead/behind, latest codex-runs list, and latest crash sidecar list.

