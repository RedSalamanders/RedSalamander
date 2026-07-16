# DxUi UIA Continuation Baton - 2026-06-29

## Purpose

Preserve the exact continuation state for the S3 plugin configuration modal hang fix and the broader DxUi/UIA reliability work that followed it.

This is a continuation note, not a completed plan. Do not move it to `Specs/Plans/Done/` until the remaining broad verification is green or explicitly superseded by a narrower follow-up plan.

## Current branch and anchor commit

- Repo: `Z:\src\RedSalamander`
- Branch: `codex/folderview-warpdrive`
- Last committed fix: `886b3fb71 Fix plugin config UIA modal selftest hang`

The committed fix addressed the user-visible S3 plugin configuration modal selftest hang by removing the two-dialog race in the live DxUi plugin configuration test and making the UIA root aggregation handle modal ordering correctly.

## Current worktree status

As of this note, the worktree is intentionally dirty with broader DxUi/UIA resilience changes. Do not reset these changes.

Modified areas:

- `Common/DxUi/*`
  - Accessibility snapshot refresh plumbing.
  - UIA provider/path resilience for virtualized and rebuilt DxUi controls.
  - `WindowHost::RefreshAccessibilitySnapshot() noexcept`.
- `RedSalamander/FolderWindow.ItemProperties.cpp`
  - Refreshes DxUi accessibility snapshots after async item-properties document rebuilds.
- `RedSalamander/ConnectionManagerWindow.cpp`
  - Handles `WM_NCDESTROY` before DxUi host message dispatch to avoid stale host access.
- `RedSalamander/SelfTest/Commands/*`
  - More robust UIA helpers for single-canvas DxUi windows.
  - Named edit lookup for Connection Manager `Name` edits.
  - Additional diagnostics for Connection Manager validation failures.
  - Removed the temporary Item Properties stream-remove scroll workaround after the snapshot refresh fix.
- `Plugins/FileSystemDummy/*`
  - Dummy plugin lifetime/behavior support for the surrounding selftests.
- `Tests/DxUiTests/DxUiTests.Accessibility.cpp`
  - Accessibility regression coverage for the broader DxUi changes.

## Verified passing evidence

Debug build:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
```

Result: passed. Latest known successful log:

```text
Z:\src\RedSalamander\.build\logs\msbuild-20260629_194323_230.log
```

Focused S3 plugin configuration modal regression:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_plugin_configuration_dialog_live_dx_interaction
```

Result: passed.

Archived run:

```text
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_162726
```

Focused Connection Manager duplicate-name validation:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_connection_manager_window_rejects_duplicate_profile_name_case_insensitive
```

Result: passed after targeting the edit named `Name`.

Archived run:

```text
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_162706
```

Focused Connection Manager blank-name validation:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_connection_manager_window_rejects_blank_profile_name
```

Result: passed after the broader diagnostic work.

Archived run:

```text
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_162931
```

Focused Item Properties stream-remove regression:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_itemProperties_window_streams_can_remove
```

Result: passed after refreshing the DxUi accessibility snapshot when the async properties document is rebuilt.

Focused screenshot follow-up checks from the later "S3 dialog stuck" investigation:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_navigation_zoom_panel_keeps_navigation_shell_stable,cmd_pane_navigation_display_modes_keep_navigation_shell_stable,cmd_pane_navigation_status_bar_keeps_navigation_shell_stable,cmd_pane_navigation_hot_paths_keeps_navigation_shell_stable,cmd_pane_navigation_find_window_keeps_navigation_shell_stable,cmd_pane_navigation_connection_manager_keeps_navigation_shell_stable,cmd_pane_navigation_compare_directories_keeps_navigation_shell_stable,cmd_pane_navigation_item_properties_window_keeps_navigation_shell_stable,cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_compare_directories_options_live_dx_body_interaction
```

Results: passed. Archives:

```text
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_171526
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_171605
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_172931
```

The short Compare Directories predecessor sequence also passed after the diagnostic snapshot-message fix:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics,cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons,cmd_compare_directories_options_long_run_open_close_stays_stable,cmd_compare_directories_options_live_dx_body_interaction
```

Archive:

```text
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_173620
```

The Preferences retained main-selection case that failed in a broad run passed by itself, and its immediate 17-case predecessor cluster also passed:

```text
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_174054
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_174214
```

Later focused and broad gates after the Preferences/plugin-config/Search helper fixes:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_plugin_configuration_dialog_live_dx_interaction
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_preferences_dialog_plugins_main_checkbox_click_toggles_live_dx_interaction
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_find_dialog_destination_navigation_stale_edit_host_hit_testing,cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

Results: build passed with 0 warnings / 0 errors; focused plugin config, focused Preferences checkbox-click, focused Find shortcut, and paired Find predecessor+shortcut checks passed; broad Commands passed with 774 passed / 0 failed / 2 opt-in ViewerSpace skips. Archives:

```text
Z:\src\RedSalamander\.build\logs\msbuild-20260629_194323_230.log
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_192805
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_191626
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_194555
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_194621
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_200556
```

Continuation sanity checks:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe
git diff --check
```

Results: DxUiTests build passed with 0 warnings / 0 errors; `DxUiTests.exe` passed with `All DxUi tests passed.`; whitespace sanity passed with line-ending warnings only. Evidence:

```text
Z:\src\RedSalamander\.build\logs\msbuild-20260629_200837_714.log
```

## Known failing or incomplete evidence

Broad Commands verification is now green on the current dirty worktree. Suite Full and the FolderView final perf matrix are still incomplete, so this baton remains WIP and must not be used as a closeout certificate.

Earlier broad runs failed before the current fixes:

- `Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_161424`
  - Failed in `cmd_connection_manager_window_live_dx_interaction`.
- `Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_161836`
  - Failed in `cmd_connection_manager_window_rejects_duplicate_profile_name_case_insensitive`.
- `Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_162859`
  - Failed in `cmd_connection_manager_window_rejects_blank_profile_name`.

Those focused cases later passed, and the full Commands suite subsequently passed at `2026-06-29_200556`.

The user-supplied screenshot of the S3 plugin configuration modal was not a live hang. It was a visible waypoint in a broad Commands run:

```text
Z:\src\RedSalamander\.build\codex-runs\commands_20260629_165937
```

That run moved past S3 into Connection Manager, Preferences, Find, and Compare Directories, then crashed later in `textinputframework.dll` with exception `0xc0000005`. No dump was recoverable from WER temp storage, but the trace, stdout/stderr, exit code, WER report, and event evidence were archived:

```text
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_171144_commands_textinputframework_crash
```

A follow-up broad Commands run with LocalDumps enabled did not crash, but it failed an order-sensitive Compare Directories assertion:

```text
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_172815
```

Failure:

```text
cmd_compare_directories_options_live_dx_body_interaction
Compare Directories options should keep its one-host DX body intact after live DX toggle/edit interaction.
```

Important: the archived failure message is unreliable because the source formatted `snapshot` in the same `state.Require(...)` call that mutates it via `waitForOptionsSnapshot(...)`. C++ does not guarantee the wait argument is evaluated before the message argument, so the printed snapshot can be stale and look passing. The current dirty worktree now has a diagnostic-only fix in `Commands.SelfTest.CompareOptions.cpp`: split the wait result into `keptStableAfterLiveInteraction` before formatting and expand `describeOptionsSnapshot(...)`.

The latest broad Commands attempt after that diagnostic fix also did not hang or crash; it failed earlier in Preferences:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

Archive:

```text
Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-29_173947
```

Result:

```text
Passed 239, failed 1, skipped 536.
Failing case: cmd_preferences_dialog_plugins_main_selection_survives_legacy_list_clear
Failure: Failed to focus the Preferences category host for Plugins retained main-selection validation.
```

That failure is also order-sensitive: the single failing case passed alone, and the immediate 17-case predecessor sequence from the archived trace passed. The active problem is therefore broad-suite UIA/focus nondeterminism, not the S3 modal.

The next broad run then failed in `cmd_plugin_configuration_dialog_live_dx_interaction` (`2026-06-29_191924`) because the S3 configuration dialog exposes multiple toggles and newer DxUi accessibility can name a toggle from its associated label rather than the state text. The selftest now scopes TogglePattern collection/invocation to `DebugGetPluginConfigurationDialogFirstVisibleToggleHostAndClientRect()` and only requires a visible name change for state-label toggles.

The next broad run then failed in `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions` (`2026-06-29_193758`) after 404 passed cases. Focused and immediate-predecessor repros passed, so the selftest helper was hardened for broad-suite residue: the keyboard context-menu route now reacquires the Find window, foreground, DxUi results-grid focus, selected-row snapshot, and valid row rect before sending `VK_APPS`. Focused and paired reruns passed at `2026-06-29_194555` and `2026-06-29_194621`; broad Commands passed at `2026-06-29_200556`.

## Next steps

1. Review the dirty tree and decide the commit split. The S3 modal fix is already committed; the current uncommitted changes are broader DxUi/UIA reliability work plus test diagnostics and should be committed separately. Keep the current broad Commands archive `2026-06-29_200556` with that split.
2. Continue FolderView plan closeout: Suite Full plus the final Debug/test-enabled Release FolderView perf matrix are still required before moving `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` to Done.

## Do not lose

- The S3 modal screenshot bug is fixed at the focused-regression level and committed in `886b3fb71`.
- The S3 modal is fixed at focused-regression level and the latest broad Commands gate is green (`2026-06-29_200556`).
- The earlier full Commands attempts moved past S3; one crashed later in `textinputframework.dll`, one failed Compare Directories with a stale diagnostic message, one failed a Preferences category-host focus assertion, one failed plugin config toggle scoping, and one failed Find keyboard context-menu focus settle. Each has either focused green evidence or a helper/test fix with broad green evidence now.
- Do not revert the current dirty worktree unless the user explicitly asks. It contains the current continuation state.

## 2026-06-30 11:26 Update

Latest continuation archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_112600_folderview_warpdrive_pause\
```

The active broad-suite blocker has moved past the earlier S3 modal and the stale
clipboard/navigation-edit lead. Current failure:

```text
cmd_connection_manager_window_tab_traversal_live_dx_interaction
```

The trace shows the Connection Manager traversal reaching the `Name` edit, then
entering the `Protocol` combo step with `preModifiers=0x4`; the routed Tab lands
back on the `Remove` command button and clears modifiers afterward. Resume by
proving whether that modifier state is causal before changing DxUi or test code.

## 2026-06-30 12:05 Update

Latest continuation archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_120509_folderview_warpdrive_pause\
```

The Connection Manager traversal failure did not reproduce in the follow-up
focused, paired, block, plugin-config-plus-connection-manager, and prefix runs.
The current broad-suite red point is now Compare Directories options:

```text
cmd_compare_directories_options_live_dx_body_interaction
Reason: Compare Directories options DX edit 'Ignore files' did not restore its original value.
```

The focused Compare Options case and immediate Compare Options cluster pass. The
best current repro is the 7-case Preferences Compare Directories plus Compare
Options sequence in:

```text
.build\codex-runs\commands_prefs_compare_then_compare_options_20260630_resume3
```

Next debugging should instrument the `Ignore files` edit restore path before
changing behavior, then prove whether this is Preferences deferred focus
restoration, retained settings state, or stale/wrong UIA edit matching.

## 2026-06-30 12:36 Update

Latest continuation archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_123644_folderview_warpdrive_pause\
```

The Compare Options `Ignore files` blocker did not reproduce after
diagnostic-only instrumentation:

```text
.build\codex-runs\commands_prefs_compare_then_compare_options_diag_20260630
Result: 7 passed / 0 failed / 0 skipped
```

The current broad-suite red point has moved to the navigation-shell selection
case:

```text
cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable
Broad run: .build\codex-runs\commands_broad_after_compare_diag_20260630
Result: 565 passed / 1 failed / 210 skipped
Reason: Navigation shell did not stay quiet during Unselect dialog cancel; currentPath='', refreshCount=21, selectedCount=3, focusedItem='b.log'.
```

Focused and immediate-cluster reruns are green:

```text
.build\codex-runs\commands_selection_unselect_navshell_focused_20260630
.build\codex-runs\commands_selection_navshell_cluster_553_565_20260630
```

Next debugging should be diagnostic-first in
`RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp`: print the
baseline and current navigation-shell state, then inspect
`WaitForNavigationViewSnapshot` / `DebugGetNavigationViewSnapshot` before
changing DxUi or product behavior.

## 2026-06-30 13:01 Update

Latest continuation archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_130129_folderview_warpdrive_pause\
```

The selection/unselect navigation-shell blocker no longer reproduces after
diagnostic-only context was added:

```text
.build\codex-runs\commands_selection_unselect_navshell_focused_diag_20260630
.build\codex-runs\commands_selection_navshell_cluster_553_565_diag_20260630
.build\codex-runs\commands_fileops_issues_to_selection_diag_20260630
.build\codex-runs\commands_shortcuts_compare_fileops_selection_diag_20260630
```

The current broad-suite red point has moved to Preferences category-tree reverse
keyboard navigation:

```text
cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation
Broad run: .build\codex-runs\commands_broad_after_selection_diag_20260630
Result: 381 passed / 1 failed / 394 skipped
Reason: Preferences category host lost keyboard focus after the second reverse-navigation VK_UP.
```

Focused and predecessor repro attempts are green:

```text
.build\codex-runs\commands_prefs_category_reverse_focused_20260630
.build\codex-runs\commands_prefs_compare_to_reverse_cluster_20260630
.build\codex-runs\commands_prefs_350_to_reverse_cluster_20260630
```

Next debugging should add narrow diagnostics to
`TestPreferencesDialogCategoryTreeReverseKeyboardNavigation` in
`RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`
before changing product or DxUi behavior. Capture snapshot state, raw focus HWND,
focused control id/class/title, and category/page fields after each key step.
