# FolderView WarpDrive Continuation Baton - 2026-07-02 11:09

## Current State

- Branch: `codex/folderview-warpdrive`
- HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
- Against `origin/master`: `0 25` (`0` behind / `25` ahead at archive time)
- Active plan: `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- This baton supersedes `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_103700_folderview_warpdrive_prefs_root_baton\CONTINUATION.md`.

## What Changed Since The Previous Baton

- `RedSalamander\Preferences.Dialog.cpp`
  - `PreferencesDialog::SelectCategory(...)` now clears `pluginsRetainedSelectedPluginId` in the same root-category path that already clears `pluginsSelectedPlugin`, `pluginsSelectedPluginId`, and `pluginsDetailsActive`.
  - Rationale: programmatic category-root selection via `DebugSelectPreferencesCategory(kPrefCategoryPlugins)` must match the real category-tree delegate path; otherwise stale retained plugin IDs can restore a Plugins details page after the caller asked for the Plugins root.
- `RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`
  - `cmd_preferences_dialog_viewers_live_search_dx_interaction` setup now uses deterministic `DebugSelectPreferencesCategory(kPrefCategoryViewers)`.
  - It clears retained Viewers search text before the live-search assertions and reports diagnostic snapshot details through `state.Require(...)` instead of failing bare.
  - Rationale: the prior Preferences test can intentionally leave a zero-row Viewers search active; the next test must establish its own unfiltered baseline.

## Verification Completed

- Debug build passed:
  - `.build\logs\msbuild-20260702_104317_912.log`
  - 0 warnings / 0 errors.
- Focused Viewers live-search passed:
  - `.build\codex-runs\prefs_viewers_live_search_after_root_state_fix_20260702.log`
- Plugins four-case root/detail probe passed:
  - `.build\codex-runs\prefs_plugins_fourcase_after_root_state_fix_20260702.log`
- Full Preferences-dialog family passed:
  - `.build\codex-runs\prefs_dialog_family_after_root_state_fix_20260702.log`
- Broad Commands first rerun got past Preferences and failed later at credential prompt theme:
  - `.build\codex-runs\commands_broad_after_prefs_root_state_fix_20260702.log`
  - Result: 209 passed / 1 failed / 568 skipped.
  - Failure: `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible`
  - Reason: `Credential prompt live secret length changed unexpectedly after the rainbow theme update; expected 18, saw 19.`
- Credential prompt focused and immediate predecessor cluster reruns passed:
  - `.build\codex-runs\credential_prompt_theme_focused_after_prefs_root_state_fix_20260702.log`
  - `.build\codex-runs\credential_prompt_theme_predecessor16_after_prefs_root_state_fix_20260702.log`
- Broad Commands second rerun got past credential prompt and failed later at Compare Directories options tab traversal:
  - `.build\codex-runs\commands_broad_rerun_after_credential_cluster_green_20260702.log`
  - Result: 522 passed / 1 failed / 255 skipped.
  - Failure: `cmd_compare_directories_options_tab_traversal_live_dx_interaction`
  - Reason: `Compare Directories options Compare attributes toggle focus target not reached during tab traversal. actualFocus=0 visibleLegacyEdits=0 visibleLegacyToggles=0 visibleDxBodyCards=10 optionsVisible=true usesDxStatics=true usesDxToggles=true usesDxEdits=true routedTo=0x9E1002C class='RedSalamander.CompareOptions.DxHost' beforeFocus=3 scroll=93/978.`

## Active Red Point

The current stopping point is the Compare Directories options tab traversal failure from broad Commands.

Trace excerpt:

```text
Case: cmd_compare_directories_options_tab_traversal_live_dx_interaction
compare_options_tab: label='Compare size toggle' reverse=false target=0x9E1002C class='RedSalamander.CompareOptions.DxHost' beforeFocus=1 beforeScroll=0/978
compare_options_tab: label='Compare size toggle' afterFocus=2 afterScroll=0/978
compare_options_tab: label='Compare date/time toggle' reverse=false target=0x9E1002C class='RedSalamander.CompareOptions.DxHost' beforeFocus=2 beforeScroll=0/978
compare_options_tab: label='Compare date/time toggle' afterFocus=3 afterScroll=0/978
compare_options_tab: label='Compare attributes toggle' reverse=false target=0x9E1002C class='RedSalamander.CompareOptions.DxHost' beforeFocus=3 beforeScroll=0/978
compare_options_tab: label='Compare attributes toggle' afterFocus=0 afterScroll=93/978
```

The immediate predecessor window was archived in:

```text
run-artifacts\commands-broad-rerun-case-window-around-compare-tab.txt
```

Cases 517..521 passed immediately before the failure:

- `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics`
- `cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons`
- `cmd_compare_directories_options_long_run_open_close_stays_stable`
- `cmd_compare_directories_options_live_dx_body_interaction`
- `cmd_compare_directories_options_pointer_click_toggles_live_dx_interaction`

## Archive Contents

- `run-artifacts\`
  - Build log plus the focused/family/broad Commands wrapper logs.
  - Case-window and trace-excerpt files for the active Compare Options failure.
- `last-run\`
  - Copied from `C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run`.
  - Includes `run-all-tests-results.json`, `results.json`, `trace.txt`, `commands\results.json`, and `commands\trace.txt` from the latest broad Commands failure.
- `patches\`
  - `staged-code-and-spec-no-testruns.patch`: staged code/spec patch excluding `Specs\TestRuns\**` so the baton does not balloon with archived logs.
  - `unstaged-diff.patch`: current unstaged tracked changes at archive time.
  - `unstaged-prefs-root-state-fix.patch`: focused patch for the two files changed in this slice.
- `source-snapshots\`
  - `Preferences.Dialog.cpp`
  - `Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`
  - `Commands.SelfTest.CompareOptions.cpp`
  - `CompareDirectoriesWindow.cpp`
  - `CompareDirectoriesWindow.Internal.h`
- Git/status files:
  - `git-status-short.txt`
  - `git-diff-stat.txt`
  - `git-diff-name-status.txt`
  - `git-diff-cached-stat.txt`
  - `git-diff-cached-name-status.txt`
  - `origin-master-ahead-behind.txt`
  - `archive-manifest.txt`

## Resume Plan

1. Focused-run the active red case:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter 'cmd_compare_directories_options_tab_traversal_live_dx_interaction' -TimeoutMultiplier 2.0 *> .build\codex-runs\compare_options_tab_focused_after_prefs_root_state_fix_20260702.log
```

2. If focused passes, rerun the immediate Compare Options predecessor cluster:

```powershell
$filter = 'cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics,cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons,cmd_compare_directories_options_long_run_open_close_stays_stable,cmd_compare_directories_options_live_dx_body_interaction,cmd_compare_directories_options_pointer_click_toggles_live_dx_interaction,cmd_compare_directories_options_tab_traversal_live_dx_interaction'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 2.0 *> .build\codex-runs\compare_options_tab_predecessor6_after_prefs_root_state_fix_20260702.log
```

3. If still green, widen with the archived case window around indices 506..522. Use `run-artifacts\commands-broad-rerun-case-window-around-compare-tab.txt` as the exact source of names.

4. If the red reproduces, inspect before patching:

```powershell
rg -n "cmd_compare_directories_options_tab_traversal_live_dx_interaction|Compare attributes toggle focus target|compare_options_tab|Compare attributes" RedSalamander\SelfTest\Commands RedSalamander
```

Likely starting files:

- `RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp`
- `RedSalamander\CompareDirectoriesWindow.cpp`
- `RedSalamander\CompareDirectoriesWindow.Internal.h`

Working hypothesis: the DX host scrolls during the `Compare attributes` Tab step (`scroll 0 -> 93`) and focus bookkeeping returns `None` even though routing stayed on `RedSalamander.CompareOptions.DxHost`. Prove whether this is a test expectation race, a scroll-into-view focus transition, or a product tab-routing bug before changing behavior.

Do not run multiple `Run-AllTests.ps1` invocations in parallel; they share the volatile `SelfTest\last_run` directory.

## Still Required Before Done

- Resolve the current broad Commands red point and rerun broad Commands.
- Rerun or complete the remaining closeout gates from the WIP plan:
  - Full suite.
  - Final Debug FolderView perf matrix.
  - Final test-enabled Release FolderView perf matrix.
  - Durable spec migration.
  - Move the WIP plan to `Specs\Plans\Done\` only after the above are green and cited.
