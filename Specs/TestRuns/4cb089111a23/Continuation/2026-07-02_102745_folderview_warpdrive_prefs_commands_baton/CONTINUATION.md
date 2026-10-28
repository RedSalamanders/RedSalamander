# FolderView WarpDrive Continuation Baton - Preferences/Commands

Date: 2026-07-02 10:27 local
Branch: `codex/folderview-warpdrive`
HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
Origin/master distance: `0 25` (`origin/master...HEAD`)

## Archive Purpose

This archive captures the work after
`Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_101000_folderview_warpdrive_commands_av_baton\CONTINUATION.md`.

The previous baton ended with a broad Commands access violation at
`cmd_pane_selection_invert`. This follow-up proves that the selection-invert
crash does not reproduce in focused or short predecessor slices, then records
the next broad-suite red points around Connection Manager and Preferences
Plugins state.

## Current State

- Do not move
  `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
  to Done yet.
- No product/source fix was made in this archive window.
- Current closeout remains blocked by broad Commands stability, then Full suite
  and final Debug plus test-enabled Release FolderView perf matrix reruns.
- The latest actionable investigation point is Preferences Plugins root/detail
  state and custom-path scrolling in the Commands selftests.

## What Was Verified

Selection-invert AV reduction:

- Focused `cmd_pane_selection_invert` passed.
- Last-10 predecessor filter passed.
- Last-20 predecessor filter passed.
- The huge prefix-to-invert exact filter attempt is not valid evidence: the
  command line/filter is too large for the harness route used here and did not
  produce a trustworthy exit/result artifact.

Connection Manager reduction:

- Broad Commands after the selection reductions failed normally at
  `cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable`.
- Focused `cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable`
  passed.
- The immediate 5-case Connection Manager cluster passed.
- The plugin-config plus Connection Manager 15-case sequence passed.
- The last-35 suffix from the failing trace passed.

Preferences Plugins reduction:

- Broad Commands after the Connection reductions got past that Connection
  Manager case and later failed at
  `cmd_preferences_dialog_plugins_page_exposes_live_uia_grid_selection`.
- Focused
  `cmd_preferences_dialog_plugins_page_exposes_live_uia_grid_selection` passed.
- The 14-case immediate Plugins predecessor sequence passed.
- `cmd_preferences_dialog_` family fail-fast then reproduced a Preferences-only
  failure, but at a later/different case:
  `cmd_preferences_dialog_plugins_custom_paths_long_run_list_scrolling_stays_bounded`.
- Focused
  `cmd_preferences_dialog_plugins_custom_paths_long_run_list_scrolling_stays_bounded`
  passed.

## Active Failure Signatures

### Broad Commands Connection Manager red, not reduced

Preserved in:

```text
run-artifacts\commands_broad_failfast_connection_churn_red_20260702\
```

Result:

- `run-all-tests-results.json`: 181 passed / 1 failed / 596 skipped / total 778.
- Failure:
  `cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable`.
- Reason:
  `Connection Manager window did not open for the focused protocol-churn test.`
- Focused and predecessor reruns passed, so treat this as intermittent or
  order-dependent unless it recurs.

### Broad Commands Preferences Plugins grid red, not reduced locally

Preserved primarily in:

```text
run-artifacts\commands_broad_failfast_after_connection_reductions_20260702.log
```

Failure:

```text
cmd_preferences_dialog_plugins_page_exposes_live_uia_grid_selection
```

Reason:

```text
Preferences Plugins page did not expose its DX grid surface for UIA selection
validation; category=7 title='S3' pluginItemSelected=true
pluginsPaneVisible=true detailsActive=true rows=23 visibleRows=12
renderedHosts=1 childWindows=1 paneWindows=0 resizeFailures=0 renderTotal=10.
```

Interpretation:

- The test expects the Plugins category root/main grid:
  `!pluginItemSelected`.
- The snapshot is on a plugin detail page (`title='S3'`,
  `pluginItemSelected=true`, `detailsActive=true`).
- This points to retained plugin child/detail selection leaking into tests that
  call `DebugSelectPreferencesCategory(kPrefCategoryPlugins)` and assume it
  lands on the category root.

### Preferences family custom-path scroll red

Preserved in:

```text
run-artifacts\prefs_dialog_family_failfast_after_plugins_uia_red_20260702.log
```

Failure:

```text
cmd_preferences_dialog_plugins_custom_paths_long_run_list_scrolling_stays_bounded
```

Reason:

```text
Preferences long-run Plugins custom-paths scrolling chunk 0 changed the active
category unexpectedly.
```

Focused rerun passed. The family run sequence immediately before failure was:

```text
cmd_preferences_dialog_plugins_page_exposes_live_uia_grid_selection
cmd_preferences_dialog_plugins_custom_paths_page_exposes_live_uia_grid_selection
cmd_preferences_dialog_plugins_long_run_list_scrolling_stays_bounded
cmd_preferences_dialog_plugins_custom_paths_long_run_list_scrolling_stays_bounded
```

## Files To Inspect Next

Start with these product/debug paths:

```text
RedSalamander\Preferences.cpp
RedSalamander\Preferences.Dialog.cpp
RedSalamander\Preferences.Plugins.cpp
RedSalamander\Preferences.h
```

Search targets:

```powershell
rg -n "DebugSelectPreferencesCategory|pluginsDetailsActive|pluginItemSelected|DebugScrollPreferencesPluginsCustomPathsListByWheelDetents|pluginsCustomPaths" RedSalamander\Preferences* RedSalamander\SelfTest\Commands -g "*.cpp" -g "*.h"
```

Useful selftest files:

```text
RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ChromeAndPlugins.cpp
RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp
RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp
```

## Next Implementation Move

1. Inspect whether `DebugSelectPreferencesCategory(kPrefCategoryPlugins)` is
   intended to select the Plugins root node or only the category host.
2. Inspect how plugin detail selection is retained/cleared:
   `pluginsDetailsActive`, selected plugin row, selected custom-path row, and
   snapshot fields.
3. Inspect the custom-path wheel helper to verify it targets the custom-paths
   list, not the category tree or retained plugin detail surface.
4. Do not patch yet unless the root/detail contract is clear. The likely safe
   direction is a test/debug helper that explicitly selects the Plugins root
   before tests that assert `!pluginItemSelected`.
5. After a fix or diagnostic hardening:
   - Build RedSalamander Debug.
   - Rerun focused plugins grid.
   - Rerun focused custom-paths scroll.
   - Rerun `cmd_preferences_dialog_` fail-fast.
   - Rerun broad Commands fail-fast.

Suggested commands:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter 'cmd_preferences_dialog_plugins_page_exposes_live_uia_grid_selection' -TimeoutMultiplier 2.0
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter 'cmd_preferences_dialog_plugins_custom_paths_long_run_list_scrolling_stays_bounded' -TimeoutMultiplier 2.0
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter 'cmd_preferences_dialog_' -TimeoutMultiplier 2.0
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2.0
```

Do not run multiple `Run-AllTests.ps1` invocations in parallel; they share the
volatile `last_run` directory.

## Archive Contents

Root snapshots:

```text
git-status-short.txt
git-branch.txt
git-head.txt
origin-master-ahead-behind.txt
git-diff-stat.txt
git-diff-cached-stat.txt
git-diff-name-status.txt
git-diff-cached-name-status.txt
patches\unstaged-diff.patch
patches\staged-diff.patch
Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.snapshot.md
process-snapshot-redsalamander.txt
recent-crash-dumps.txt
codex-runs-recent-files.txt
```

Copied run artifacts include:

```text
run-artifacts\focused_selection_invert_after_av_baton_20260702.log
run-artifacts\selection_av_last10_after_baton_20260702.*
run-artifacts\selection_av_last20_after_baton_20260702.*
run-artifacts\selection_av_prefix_to_invert_after_baton_20260702.*
run-artifacts\commands_broad_failfast_after_selection_reductions_20260702.*
run-artifacts\commands_broad_failfast_connection_churn_red_20260702\
run-artifacts\focused_connection_protocol_churn_after_broad_red_20260702.log
run-artifacts\connection_protocol_churn_immediate5_20260702.*
run-artifacts\plugin_config_plus_connection_churn_20260702.*
run-artifacts\connection_churn_red_last35_20260702.*
run-artifacts\commands_broad_failfast_after_connection_reductions_20260702.*
run-artifacts\focused_prefs_plugins_uia_grid_after_broad_red_20260702.log
run-artifacts\prefs_plugins_uia_grid_predecessor14_20260702.*
run-artifacts\prefs_dialog_family_failfast_after_plugins_uia_red_20260702.log
run-artifacts\focused_prefs_plugins_custom_paths_scroll_after_family_red_20260702.log
```

## Remaining Before Done

- Resolve or explicitly waive broad Commands blockers.
- Run FileOps, ViewerPE/ViewerVLC HUD checks, Tools Pester, and Full suite.
- Rerun final Debug and test-enabled Release FolderView perf matrices from the
  current HEAD.
- Cite final archives in the WIP plan.
- Move the WIP plan to `Specs\Plans\Done\` only after all closeout gates pass.
