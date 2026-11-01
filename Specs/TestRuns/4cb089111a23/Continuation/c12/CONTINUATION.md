# FolderView WarpDrive Continuation Baton

Created: 2026-07-03 15:55:47 Europe/Paris
Workspace: `Z:\src\RedSalamander`
Branch: `codex/folderview-warpdrive`
Branch relation at archive time: `25` ahead / `4` behind `origin/master`

## Current State

This is a pause archive for the active WIP plan:

```text
Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md
```

Do not move the plan to Done yet. The FolderView perf implementation is broadly through its original product/perf work, but closeout is still gated by broad-suite Commands reliability, then origin/master resync, then final full-suite and perf matrix evidence.

The latest red edge is not the earlier crash and not the status-bar/menu/double-click failures. The current reproducible blocker is:

```text
cmd_pane_navigationView_full_path_popup_edit_route
```

It failed only in the 17-case NavigationView prefix window, then passed focused:

```text
.build\codex-runs\commands_nav_view_prefix_after_path_doubleclick_surface_ready_20260703.log
Specs\TestRuns\4cb089111a23\Commands\2026-07-03_155328

.build\codex-runs\commands_nav_full_path_popup_edit_route_focused_after_window_red_20260703.log
Specs\TestRuns\4cb089111a23\Commands\2026-07-03_155401
```

Failure text from the red prefix run:

```text
Navigation view full-path popup did not reflow across the synthetic DPI transition.
popupVisible=yes popupEditMode=no dpi=96 client=0x0 baseline=3177x60 currentPath=''.
```

Initial interpretation: the popup is still visible and not in edit mode, but the synthetic DPI transition did not settle into the expected debug snapshot in this broad order. `dpi=96`, `client=0x0`, and `currentPath=''` suggest the test is reading either a stale/unready NavigationView snapshot, a non-applied synthetic DPI state, or a transient popup recreation/layout state. Do not patch this by weakening the assertion until the product/test lifecycle is understood.

## What Changed Since The Previous Baton

Compare crash root cause was fixed and verified:

- `CompareDirectoriesSession::~CompareDirectoriesSession()` now wakes scan/content workers and explicitly joins `_scanWorkers` before `_contentCompareWorkers`.
- Added `worker_shutdown_joins_before_state_teardown`.
- Re-ran Compare targeted and full suite successfully:
  - `.build\codex-runs\compare_worker_shutdown_join_guard_20260703.log` => 1 passed / 0 failed / 0 skipped.
  - `.build\codex-runs\compare_pending_wide_tree_after_worker_join_20260703.log` => 1 passed / 0 failed / 0 skipped.
  - `.build\codex-runs\compare_full_after_worker_join_20260703_rerun.log` => 208 passed / 0 failed / 30 skipped.

Commands status-bar broad failure was investigated and no longer reproduces in focused/window runs:

- Focused pass: `.build\codex-runs\commands_status_bar_nav_shell_focused_after_worker_join_20260703.log`.
- 7-case window pass: `.build\codex-runs\commands_status_bar_nav_shell_window_after_worker_join_20260703.log`.

Commands menu-region broad failure was patched in selftest code:

- Added `WaitForNavigationViewKeyboardMenuFocusReturn(...)`.
- Replaced raw single-sample final focus checks in menu-region and disk-info keyboard activation cases with a passive condition-based wait and diagnostics.
- Build passed: `.build\codex-runs\msbuild_redsalamander_after_nav_menu_focus_wait_20260703.log`.
- Focus pair passed: `.build\codex-runs\commands_nav_keyboard_menu_focus_pair_after_wait_20260703.log`.
- 10-case window passed: `.build\codex-runs\commands_nav_view_keyboard_menu_window_after_focus_wait_20260703.log`.

Commands path double-click-after-tab broad failure was patched in selftest code:

- Added `pathEditSurfaceReady(...)` for `TestPaneNavigationViewPathDoubleClickEntersEditMode(...)`.
- The edit surface can now be considered ready when edit mode, text, host/input, selection, and folder state are correct even if `focusTarget` is still `None` during a transient null-focus handoff.
- The later tab-handoff loop now uses the same readiness predicate and can focus the edit input when the edit surface is already valid.
- Build passed: `.build\codex-runs\msbuild_redsalamander_after_nav_path_doubleclick_surface_ready_20260703.log`.
- Focused triplet passed: `.build\codex-runs\commands_nav_path_doubleclick_triplet_after_surface_ready_20260703.log`.

The next broader NavigationView window found the current red edge:

```text
.build\codex-runs\commands_nav_view_prefix_after_path_doubleclick_surface_ready_20260703.log
16 passed / 1 failed / 0 skipped
```

## Recent Crash State

No new crash dump was observed after the earlier Compare crash while investigating the Commands failures.

Latest copied RedSalamander crash artifacts:

```text
C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260703-142049-p98100.dmp
C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260703-142049-p98100.txt
```

Both are copied under:

```text
crash-artifacts\
```

Remember: RedSalamander crash dumps normally live in:

```text
C:\Users\eric\AppData\Local\RedSalamander\Crashes
```

and each dump should have a nearby `.txt` sidecar with the call stack.

## Archive Contents

This archive contains:

- `CONTINUATION.md` - this baton.
- `archive-manifest.txt` - file list.
- Git refs/status/ahead-behind and no-TestRuns diff snapshots.
- `patches\staged-code-and-spec-no-testruns.patch`.
- `patches\unstaged-code-and-spec-no-testruns.patch`.
- Focused, window, broad Commands logs from this turn.
- Compare worker-join verification logs from this turn.
- Recent Commands/Compare `Specs\TestRuns` result and trace files.
- Current source snapshots for `Commands.SelfTest.ViewCommands.cpp`, Compare worker-join files, `NavigationView*`, and the WIP spec.
- Latest RedSalamander crash sidecar artifacts.

## Resume Plan

1. Continue with systematic debugging. Read `NavigationView.cpp` `OnDpiChanged(...)`, `UpdateFullPathPopupWindow(...)`, debug snapshot population, and `TestPaneNavigationViewFullPathPopupEditRoute(...)`.
2. Compare the similar synthetic-DPI edit-suggest test around `Commands.SelfTest.ViewCommands.cpp` lines near `6605`; it may already have a better wait/predicate pattern.
3. Run reduced predecessor slices before patching:
   - `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane,cmd_pane_navigationView_full_path_popup_edit_route`
   - `cmd_pane_navigationView_path_ancestor_click_navigates_to_ancestor,cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane,cmd_pane_navigationView_full_path_popup_edit_route`
   - If needed, extend toward the exact 17-case red prefix.
4. Root-cause whether the failure is product state, stale HWND/snapshot state, or a test wait that assumes synthetic `WM_DPICHANGED` changes `GetDpiForWindow(...)` synchronously in all orders.
5. Patch narrowly, rebuild `RedSalamander` Debug, rerun focused, predecessor/window, then broad Commands.
6. After broad Commands is green, re-check branch relation (`25` ahead / `4` behind at this archive). Resync/rebase with origin/master only after preserving/staging current work and rerun relevant tests after the rebase.
7. Final closeout still needs broad Commands, FileOps/ViewerPE/Tools Pester as applicable, Full suite, and final Debug plus test-enabled Release FolderView perf matrix evidence cited in the WIP plan.

