# FolderView WarpDrive Continuation Baton

Created: 2026-07-03 15:05 local time

Branch: `codex/folderview-warpdrive`
HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
Local relation to `origin/master`: `25` ahead / `4` behind at archive time.

## Current State

The previous CompareDirectories crash (`RedSalamander-20260703-142049-p98100`) is root-caused and patched in the dirty worktree. The crash was not a simple missing lock around `_contentCompareCache.find`; the cache lookup already happened under `_mutex`.

Root cause:

- `CompareDirectoriesSession::~CompareDirectoriesSession()` requested stop and notified workers, but relied on implicit member destruction to join `std::jthread`s.
- In `CompareDirectoriesEngine.h`, `_scanWorkers` is declared before `_contentCompareCache` and content-compare state.
- Reverse member destruction therefore joined/destroyed content workers/state before scan workers.
- A scan worker could still run `ComputeDecisionForFolder()` and access `_contentCompareCache` after that state was destroyed.
- This matches the crash sidecar stack exactly: scan worker -> `ComputeDecisionForFolder()` -> `_contentCompareCache.find`.

Patch currently in worktree:

- `RedSalamander/CompareDirectoriesEngine.cpp`
  - Destructor now wakes scan/content queues and explicitly joins `_scanWorkers` before `_contentCompareWorkers`, avoiding dependence on member destruction order.
- `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp`
  - Registers `worker_shutdown_joins_before_state_teardown`.
- `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp`
  - Adds the runtime/structural regression guard for active shutdown while content compare is in flight.

No new RedSalamander crash-handler dump appeared after the worker-join fix. Newest crash artifacts remain the preserved `p98100` pair from 2026-07-03 14:20:49.

## Verification Captured

Build:

- `.build/codex-runs/msbuild_redsalamander_after_compare_worker_join_20260703_rerun.log`
- Result: Debug `RedSalamander` build passed, `0 warning(s), 0 error(s)`.

Compare focused gates:

- `.build/codex-runs/compare_worker_shutdown_join_guard_20260703.log`
  - `worker_shutdown_joins_before_state_teardown`: 1 passed / 0 failed / 0 skipped.
  - Stable archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-03_143301`.
- `.build/codex-runs/compare_pending_wide_tree_after_worker_join_20260703.log`
  - `decision_cache_eviction_budget_pending_wide_tree`: 1 passed / 0 failed / 0 skipped.
  - Stable archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-03_143335`.

Full CompareDirectories:

- `.build/codex-runs/compare_full_after_worker_join_20260703_rerun.log`
- Stable archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-03_144451`
- Result: 208 passed / 0 failed / 30 skipped.

Commands focused/window gate:

- `.build/codex-runs/commands_clipboard_nav_edit_window_after_worker_join_20260703.log`
- Stable archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-03_144529`
- Result: 5 passed / 0 failed / 0 skipped.

Broad Commands:

- `.build/codex-runs/commands_broad_after_worker_join_20260703.log`
- Stable archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-03_150248`
- Result: 775 passed / 1 failed / 2 skipped.
- Failure: `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`
- Message: `Navigation shell did not stay quiet during Show Status Bar; focusTarget=0, editMode=no, historyVisible=no, suggestVisible=no, popupVisible=no, childWindows=0, currentPath='', historyCount=0, refreshCount=22, itemCount=3, selectedCount=0, focusedItem='b.log', statusVisible=no, statusDirectWrite=no, statusClass=''.`

## Archive Contents

- `CONTINUATION.md`: this restart note.
- `archive-manifest.txt`: complete file list for this baton.
- `git-status-short.txt`, `git-head.txt`, `git-branch.txt`, `origin-master-ahead-behind.txt`.
- `git-diff-*-no-testruns.txt`: tracked change summaries excluding `Specs/TestRuns`.
- `patches/staged-code-and-spec-no-testruns.patch`.
- `patches/unstaged-code-and-spec-no-testruns.patch`.
- `run-artifacts/`: copied wrapper logs and stable run result/trace JSONs for the build, Compare gates, Commands window, and broad Commands red.
- `source-snapshots/`: snapshots of the Compare destructor/test files, current Navigation selftest file, the WIP spec, and the crash-forensics skill.
- `crash-artifacts/`: latest RedSalamander crash list plus the preserved `RedSalamander-20260703-142049-p98100.txt/.dmp` pair and `last_crash.txt`.

## Next Steps

1. Check current branch relation to `origin/master` before final gates; this archive recorded `25` ahead / `4` behind.
2. Run the focused Commands red:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_pane_navigation_status_bar_keeps_navigation_shell_stable" -TimeoutMultiplier 2 *> .build\codex-runs\commands_status_bar_nav_shell_focused_after_worker_join_20260703.log; exit $LASTEXITCODE
```

3. If focused passes, run the navigation/status predecessor window:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_pane_navigation_refresh_keeps_navigation_shell_stable,cmd_pane_navigation_directory_impact_preserves_selection,cmd_pane_navigation_directory_impact_preserves_selection_across_chained_renames,cmd_pane_navigation_zoom_panel_keeps_navigation_shell_stable,cmd_pane_navigation_display_modes_keep_navigation_shell_stable,cmd_pane_navigation_status_bar_keeps_navigation_shell_stable,cmd_pane_navigation_hot_paths_keeps_navigation_shell_stable" -TimeoutMultiplier 2 *> .build\codex-runs\commands_status_bar_nav_shell_window_after_worker_join_20260703.log; exit $LASTEXITCODE
```

4. If the failure repeats, inspect `TestStatusBarKeepsNavigationShellStable(...)` in `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`. Current suspicion is a selftest wait/setup race around returning focus to the left pane before toggling the status bar; do not assume product behavior is wrong without focused/window evidence.
5. After Commands is green again, rerun proper Full with build, then final Debug and test-enabled Release `folderView_perf_` matrices from the current rebased HEAD.
6. Only move the WIP plan to `Specs/Plans/Done/` after Full and final perf evidence are green and durable spec migrations are complete.

Do not run multiple `Tools/Run-AllTests.ps1` invocations in parallel; they share volatile selftest state.
