# Operation Crosscut Compare Directories Remediation - Handoff 2026-06-24

## Closed — 2026-07-02 folder review

Mission fully accomplished; this handoff is archived.

- Parent plan is in `Specs/Plans/Done` (`Operation_Crosscut_CompareDirectoriesRemediation_SyncDataSafetyAndOptionsSimplification_2026-06-15.md`, commit `7613b115c`).
- Every resume step below and the 11 failing Commands cases were owned and completed by `Operation_CommandsSelfTestInputIsolation_2026-06-24.md` — broad Commands green `762 passed / 0 failed / 2 skipped` on 2026-06-28, re-confirmed `776 passed / 0 failed / 2 skipped` at `Specs/TestRuns/4cb089111a23/Commands/2026-07-02_121623`.
- The Full-suite gate is owned by `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md`.
- The Repository State section describes a dirty `D:\RedSalamander` tree that was committed in `7613b115c` and no longer exists.
- The new CompareDirectories bug (sticky Invert Differences Selection, found 2026-07-01) is TW-8, owned by Operation_Tailwind — it must NOT reopen this handoff or the Done Crosscut plan.

Everything below is retained as historical record only.

## Repository State

- Working directory: `D:\RedSalamander`
- Branch: `master`
- Local `HEAD`: `2fc4f817fcf4`
- `origin/master`: `2fc4f817fcf4`
- Worktree status: intentionally dirty; do not reset or stash blindly.
- Active WIP plan: `Specs/Plans/WIP/Operation_CommandsSelfTestInputIsolation_2026-06-24.md`
- Completed Compare plan: `Specs/Plans/Done/Operation_Crosscut_CompareDirectoriesRemediation_SyncDataSafetyAndOptionsSimplification_2026-06-15.md`
- Authoritative Compare spec: `Specs/Core/Core_CompareDirectories.md`

## Last User Request

Add a WIP spec for the broad Commands errors and move the CompareDirectories plan to Done if no Compare work remains. Do not commit unless the user asks again.

## What Is Saved

The CompareDirectories plan is closed and moved to Done. The remaining broad Commands failures are now tracked by a dedicated WIP plan:

- Done: `Specs/Plans/Done/Operation_Crosscut_CompareDirectoriesRemediation_SyncDataSafetyAndOptionsSimplification_2026-06-15.md`
- WIP: `Specs/Plans/WIP/Operation_CommandsSelfTestInputIsolation_2026-06-24.md`
- Compare Directories implementation slices are complete and durable behavior has been merged into `Specs/Core/Core_CompareDirectories.md`.
- Current blocker is Commands order/input sensitivity, not CompareDirectories correctness.

Source edits and spec edits are saved in the working tree. No staging or commit was performed for this handoff.

## Latest Build Evidence

Latest clean test-enabled Debug build:

```powershell
$env:RSBuildEnableTests='true'
.\build.ps1 -Configuration Debug -Platform x64 -ProjectName RedSalamander
```

Logs:

```text
.build\logs\msbuild-20260624_142759_244.log
.build\logs\msbuild-20260624_142801_196.log
```

Result: `0 warning(s), 0 error(s)`.

## Latest Commands Evidence

Focused menu mouse-open blocker is green:

```text
.build\logs\run-commands-menu-mouse-open-20260624_143208_163.log
```

Result: `cmd_app_menuBar_mouse_open_keeps_popup_selection_clear`, `1 passed / 0 failed`.

Latest broad Commands closeout run:

```text
.build\logs\run-commands-closeout-20260624_143239.out.log
Specs\TestRuns\7d3a1247382a\Commands\2026-06-24_150328\commands_results.json
```

Result: `751 passed / 11 failed / 2 skipped`.

The exact 11 failed cases passed when rerun in focused clusters:

```text
.build\logs\run-commands-failing-credential-20260624_150453_924.log      2 passed / 0 failed
.build\logs\run-commands-failing-preferences-20260624_150512_968.log     7 passed / 0 failed
.build\logs\run-commands-failing-menu-20260624_150555_203.log            2 passed / 0 failed
```

A wider plugin/connection-manager/credential replay still reproduced UIA/value-entry instability:

```text
.build\logs\run-commands-plugin-connection-credential-20260624_150743_323.log
```

Result: `41 passed / 5 failed`. Important clue: one case expected the localized default name `New connection` and observed `rce`, which points to stale keyboard/focus contamination or an equivalent value-entry ordering defect.

## Current Broad Commands Failures

- `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible`
- `cmd_connection_credential_prompt_live_dx_interaction`
- `cmd_preferences_dialog_plugins_theme_cycle_keeps_surface_legible`
- `cmd_preferences_dialog_viewers_theme_cycle_keeps_surface_legible`
- `cmd_preferences_dialog_viewers_live_search_dx_interaction`
- `cmd_preferences_dialog_viewers_add_update_live_dx_interaction`
- `cmd_preferences_dialog_viewers_selection_survives_legacy_list_clear`
- `cmd_preferences_dialog_plugin_tree_selection_keeps_single_visible_pane`
- `cmd_preferences_dialog_plugins_main_selection_survives_legacy_list_clear`
- `cmd_app_menuBar_arrow_switches_top_level_popup`
- `cmd_app_menuBar_submenu_placement_matches_spec`

## Input-Safety Rule

Minimize direct mouse/keyboard paths. Use UIA/message-pumped helpers where possible. When a test genuinely requires real cursor movement, display a large on-screen warning with this exact text before the movement and dismiss it immediately afterward:

```text
don't touch the mouse
```

Current code already does this for the direct-hover/mouse path that requires cursor movement.

## Next Resume Steps

1. Treat the remaining Commands blocker as input/focus isolation until disproven.
2. Harden the shared UIA value-entry/cancel-reopen helpers used by credential, plugin configuration, connection manager, and Preferences tests:
   - verify the active window and focused owner before mutating controls,
   - avoid direct keyboard when UIA ValuePattern/message routes are sufficient,
   - drain stale input/messages after modal teardown,
   - keep real cursor movement guarded by the warning above.
3. Rebuild with tests enabled:

   ```powershell
   $env:RSBuildEnableTests='true'
   .\build.ps1 -Configuration Debug -Platform x64 -ProjectName RedSalamander
   ```

4. Rerun the wider replay:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_plugin_configuration_dialog_uses_dxui_command_buttons,cmd_plugin_configuration_dialog_uses_dxui_form_surface,cmd_plugin_configuration_dialog_long_run_scrolling_keeps_dx_surface_stable,cmd_plugin_configuration_dialog_long_run_open_close_stays_stable,cmd_plugin_configuration_dialog_theme_cycle_keeps_surface_legible,cmd_plugin_configuration_dialog_live_dx_interaction,cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction,cmd_plugin_configuration_dialog_enter_and_escape_route_default_cancel,cmd_plugin_configuration_dialog_access_keys_route_expected_actions,cmd_plugin_configuration_dialog_pointer_click_toggles_visible_dx_toggle,cmd_connection_manager_window_uses_dxui_command_buttons,cmd_connection_manager_window_uses_dxui_form_inputs,cmd_connection_manager_window_layout_keeps_cards_and_fields_clean,cmd_connection_manager_window_uses_dxui_form_action_buttons,cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable,cmd_connection_manager_window_modeless_connect_posts_left_navigation,cmd_connection_manager_window_modeless_connect_posts_right_navigation,cmd_connection_manager_window_rejects_blank_profile_name,cmd_connection_manager_window_rejects_duplicate_profile_name_case_insensitive,cmd_connection_manager_window_rejects_reserved_quick_profile_name,cmd_connection_manager_window_trims_profile_name_before_save,cmd_connection_manager_window_clean_external_reload_refreshes_list,cmd_connection_manager_window_dirty_external_reload_prompts_and_keeps_editing,cmd_connection_manager_window_stale_save_prompts_before_overwrite,cmd_connection_manager_window_uses_localized_strings_for_dynamic_labels,cmd_connection_manager_window_list_no_horizontal_scroll_and_saved_secret_reveals,cmd_connection_manager_window_retired_dialog_files_absent,cmd_connection_manager_window_live_dx_interaction,cmd_connection_manager_window_textfield_doubleclick_selects_word,cmd_connection_manager_window_masked_secret_accepts_native_chars,cmd_connection_manager_window_close_persists_new_profile,cmd_connection_manager_window_wm_close_discards_new_profile,cmd_connection_manager_window_long_run_list_scrolling_stays_bounded,cmd_connection_manager_window_theme_cycle_keeps_form_and_selection_legible,cmd_connection_manager_window_applies_selected_tool_backdrop,cmd_connection_manager_window_long_run_open_close_stays_stable,cmd_connection_manager_window_tab_traversal_live_dx_interaction,cmd_connection_manager_window_enter_from_dx_input_routes_default_connect,cmd_connection_manager_window_escape_from_dx_input_closes_cancel,cmd_connection_manager_window_access_keys_focus_expected_controls,cmd_connection_manager_window_pointer_click_toggles_visible_dx_toggle,cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible,cmd_connection_credential_prompt_dxui_validation_and_accept,cmd_connection_credential_prompt_escape_cancels_secret_only,cmd_connection_credential_prompt_long_run_open_close_stays_stable,cmd_connection_credential_prompt_live_dx_interaction" -TimeoutMultiplier 2
   ```

5. Rerun full Commands:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
   ```

6. Only after Commands is stable, rerun the full closeout:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2
   ```

## Commit Readiness

Not ready to commit yet. The code builds and the Compare plan/spec are closed, but the active WIP Commands input-isolation plan still owns a red broad Commands gate.
