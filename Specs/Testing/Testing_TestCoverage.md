# Test Coverage Specification

## Overview

This document provides a comprehensive inventory of every declared test case across all
RedSalamander test suites. It serves as the authoritative reference for test coverage.

Total declared cases: **520+ self-test cases** across 3 suites, plus PoC and performance tests.

Related documents:
- `Specs/Testing/Testing_SelfTests.md` — result contract
- `Specs/Testing/Testing_PerformanceValidation.md` — perf validation requirements
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md` — remote credential setup
- `Tests/README.md` — test infrastructure index

Recent UI-retirement evidence:

- `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_111643/` — Connection Manager battle-test family after dialog shim retirement; `cmd_connection_manager_window_` passed 25 passed / 0 failed / 0 skipped and archived `perf/perf_metrics.jsonl`.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_012857_remaining_win32_ui_dependency_post_closeout_recheck/` — post-closeout broad HFONT/GDI/native-control audit recheck. Audit perf: 3.446 seconds; zero unallowlisted findings.
- `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_012857_remaining_win32_ui_dependency_post_closeout_menu/` — post-closeout `.build\x64\Debug\DxUiTests.exe --filter Menu` run, exited 0 in 15.373 seconds.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_012017_remaining_win32_ui_dependency_final_recheck/` — final wrapper broad HFONT/GDI/native-control audit recheck. Audit perf: 3.315 seconds; zero unallowlisted findings.
- `.build/logs/msbuild-20260426_010339_544.log` and `.build/logs/msbuild-20260426_010518_794.log` — wrapper-pass Debug/Release builds after the Function Bar paint-DC stabilization and final closeout documentation refresh.
- `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_011821_remaining_win32_ui_dependency_final_full_dxui/` — full `.build\x64\Debug\DxUiTests.exe` run, exited 0 in 110.547 seconds.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_010657/`, `2026-04-26_010717/`, `2026-04-26_010746/`, `2026-04-26_011225_connection_manager_final_recheck/`, `2026-04-26_011255_plugin_configuration_final_recheck/`, and `2026-04-26_011808_preferences_final_recheck/` — wrapper focused command recheck for UI chrome, status bar shell stability, Compare Options, Connection Manager, Plugin Configuration, and Preferences; all exited 0 with zero failures.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_005800_remaining_win32_ui_dependency_final_recheck/` — final broad HFONT/GDI/native-control audit recheck. Audit perf: 3.352 seconds; zero unallowlisted findings.
- `.build/logs/msbuild-20260426_005323_200.log` and `.build/logs/msbuild-20260426_005751_693.log` — final Debug/Release `RedSalamander` builds after the Function Bar layout/material-tolerant pixel guard.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_005455/`, `2026-04-26_005540/`, and `2026-04-26_005709/` — final focused command recheck for UI chrome, status bar shell stability, and Compare Options; all passed with zero failures. The failed `2026-04-26_005614/` Compare Options attempt is non-gating because the same failed case passed in isolation at `2026-04-26_005634/` and the full family passed on rerun.
- `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_005914_remaining_win32_ui_dependency_final_menu/` — final `DxUiTests.exe --filter Menu` recheck, exited 0 in 14.632 seconds.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_001034_remaining_win32_ui_dependency_recheck/` — current-day broad HFONT/GDI/native-control audit recheck. Audit perf: 3.339 seconds; zero unallowlisted findings.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_001212/`, `2026-04-26_001216/`, and `2026-04-26_001234/` — current-day focused command recheck for UI chrome, status bar shell stability, and Compare Options; all passed with zero failures. The failed `2026-04-26_001047/` and `2026-04-26_001156/` `cmd_app_toggleUiChrome` attempts are non-gating because they launched the visibility-sensitive GUI selftest with a hidden top-level window.
- `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_001241_remaining_win32_ui_dependency_recheck_menu/` — current-day `DxUiTests.exe --filter Menu` recheck, exited 0 in 15.259 seconds.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_235110_remaining_win32_ui_dependency_closeout/` — broad HFONT/GDI/native-control audit after the remaining app-owned HDC paint/selection seams moved behind `D2DHdcPaint::Session` and the last Compare Options legacy `STATIC` fallback was replaced. Audit perf: 3.219 seconds; zero unallowlisted findings.
- `.build/logs/msbuild-20260425_234924_528.log` — Debug build after the remaining paint/static closeout.
- `.build/logs/msbuild-20260425_235255_302.log` — Release build after the remaining paint/static closeout.
- `.build/logs/msbuild-20260425_235401_128.log` and `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_000537_remaining_win32_ui_dependency_closeout/` — DxUiTests build plus full DxUiTests run, exited 0.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_235651/`, `2026-04-25_235730/`, `2026-04-25_235809/`, `2026-04-25_235945/`, `2026-04-26_000025/`, and `2026-04-26_000514/` — closeout focused command refresh for UI chrome, status bar, Compare Options, Connection Manager, Plugin Configuration, and Preferences; all passed with zero failures.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_232753_remaining_win32_ui_native_hosts_reduced_candidate/` — broad HFONT/GDI/native-control audit after shared native-font helper removal and DxUi host-window class cleanup. Audit perf: 3.433 seconds; remaining closeout blockers are 22 HDC text/selection bridges and one Compare Options legacy `STATIC` fallback.
- `.build/logs/msbuild-20260425_232017_534.log` — Debug build after plugin viewer unused native font handles were removed.
- `.build/logs/msbuild-20260425_232609_620.log` — Debug build after native host cleanup.
- `.build/logs/msbuild-20260425_233112_306.log` — Release build after native host cleanup.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_232948/` — Connection Manager focused family, 13 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233010/` — Plugin Configuration focused family, 10 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233026/` — Compare Options focused family, 9 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233028/`, `2026-04-25_233032/`, and `2026-04-25_233037/` — Preferences shell/general/panes refresh after shell font propagation removal.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233054/` — Navigation edit-suggest keyboard routing direct focused rerun passed after hidden-launch archive `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233044/` failed transiently.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_224400_remaining_win32_ui_bridge_allowlist_uimetrics_viewerweb_candidate/` — broad HFONT/GDI/native-control audit after hidden text bridge allowlist, `Win32UiHelpers` deletion/`UiMetrics` split, dead HFONT measurement bridge removal, dialog base-`LOGFONT` cloning removal, and ViewerWeb DirectWrite status drawing. Audit perf: 3.331 seconds.
- `.build/logs/msbuild-20260425_222825_786.log` — Debug build after the helper split/native-font bridge cleanup.
- `.build/logs/msbuild-20260425_224509_797.log` — Release build after the helper split/native-font bridge cleanup.
- `.build\x64\Debug\DxUiTests.exe --filter TextInputBridge` — focused hidden text-service bridge regression, exited 0.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224155/` — Connection Manager focused family, 13 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224242/` — Plugin Configuration focused family, 10 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224312/`, `2026-04-25_224320/`, `2026-04-25_224331/`, and `2026-04-25_224338/` — Preferences shell/page-host/general/panes focused guards.
- Non-closeout caveat: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224026/` was an accidental broad unfiltered Commands run and failed 4 Preferences Viewers reorder/search cases; do not use it as green closeout evidence for `Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md`.

---

## 1. Commands Suite (`--commands-selftest`)

**Source:** `RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp` orchestrator + 12 included `.cpp` family files (70K+ lines, 360+ cases)

The Commands suite is split into logical `.cpp` family files included from the main orchestrator:
- `SelfTest\Commands\Commands.SelfTest.Settings.cpp` — Settings hot-reload, store, registry, preview/file-action guards, shortcut defaults (13+ cases)
- `SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp` — Plugin configuration and file-system plugin (13 cases)
- `SelfTest\Commands\Commands.SelfTest.Connections.cpp` — Connection manager and credentials (36 cases)
- `SelfTest\Commands\Commands.SelfTest.Preferences.cpp` coordinator + 7 included chunk files — Preferences dialog automation (129 cases)
- `SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp` — Compare directories options, chrome, and progress (11 cases)
- `SelfTest\Commands\Commands.SelfTest.Search.cpp` — Find dialog, local search index, quick search/filter (52 cases)
- `SelfTest\Commands\Commands.SelfTest.Shortcuts.cpp` — Shortcuts window (31 cases)
- `SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp` — View commands, selection, sort, pane, tabs (29+ cases)
- `SelfTest\Commands\Commands.SelfTest.FileOps.cpp` — File operations issues pane, speed limit prompt (21 cases)
- `SelfTest\Commands\Commands.SelfTest.Navigation.cpp` — Navigation location, GoTo, navigation/drive menu shell stability, Escape focus reclaim to the active FolderView, navigation-menu `Go to >` placement before drive rows, nonstandard file-system `Common Folders` submenu coverage, and directory-impact selection preservation
- `SelfTest\Commands\Commands.SelfTest.ShellCommands.cpp` — Shell-integrated pane commands including Change Attributes attributes/date-time/stream reports and recursive progress
- `SelfTest\Commands\Commands.SelfTest.Dialogs.cpp` — About, fatal error, splash, change case, filter, rename, including long initial rename-selection clipping, etc. (48 cases)

The suitetests UI automation, dialog interactions, preferences, shortcuts,
themes, navigation, and command dispatch. All cases run inside the live application
window using UIAutomation and direct Win32 message simulation.

### 1.1 Application-Level Commands (17 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_app_about_access_key_routes_ok` | About dialog keyboard access keys |
| `cmd_app_about_enter_and_escape_route_ok` | About dialog Enter/Escape handling |
| `cmd_app_about_live_dx_interaction` | About dialog DxUi surface interaction |
| `cmd_app_about_long_run_open_close_stays_stable` | About dialog stability under repeated open/close |
| `cmd_app_about_uses_dxui_surface` | About dialog uses DxUi rendering |
| `cmd_app_fatal_error_access_key_routes_ok` | Fatal error dialog access keys |
| `cmd_app_fatal_error_enter_and_escape_route_ok` | Fatal error Enter/Escape |
| `cmd_app_fatal_error_live_dx_interaction` | Fatal error DxUi interaction |
| `cmd_app_fatal_error_long_run_open_close_stays_stable` | Fatal error stability |
| `cmd_app_fatal_error_long_run_scrolling_stays_bounded` | Fatal error scroll bounds |
| `cmd_app_fatal_error_theme_matrix_keeps_message_legible` | Fatal error theme legibility |
| `cmd_app_fatal_error_uses_dxui_surface` | Fatal error DxUi surface |
| `cmd_app_fullScreen` | Full-screen toggle |
| `cmd_app_menuBar_mouse_open_keeps_popup_selection_clear` | Mouse-opened menu popups keep item selection empty until pointer movement or keyboard navigation |
| `cmd_app_menuBar_submenu_placement_matches_spec` | Menu submenu placement and parent-hover retention |
| `cmd_app_swapPanes` | Pane swap command |
| `cmd_app_toggleUiChrome` | UI chrome visibility toggle, function bar layout, painted themed pixels, DirectWrite text metrics, and splitter arrow affordances |

### 1.2 Prompt and Splash (5 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_app_prompt_access_keys_route_expected_actions` | Prompt dialog access keys |
| `cmd_app_prompt_long_run_open_close_stays_stable` | Prompt dialog stability |
| `cmd_app_prompt_uses_alert_overlay_window` | Prompt uses alert overlay |
| `cmd_app_splash_live_dx_text_update` | Splash screen DxUi text update |
| `cmd_app_splash_long_run_open_close_stays_stable` | Splash screen stability |
| `cmd_app_splash_uses_dxui_surface` | Splash screen DxUi surface |
| `cmd_app_viewWidth` | View width command |

### 1.3 Shortcuts Window (31 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_app_shortcuts_column_reorder_survives_search_roundtrip` | Column reorder persists through search |
| `cmd_app_shortcuts_column_reorder_survives_sort_cycles` | Column reorder persists through sort |
| `cmd_app_shortcuts_copy_follows_reordered_columns` | Copy respects column order |
| `cmd_app_shortcuts_escape_closes_from_search_and_grid` | Escape key behavior |
| `cmd_app_shortcuts_grid_doubleClick_activates_selected_command` | Double-click activation |
| `cmd_app_shortcuts_grid_enter_activates_selected_command` | Enter key activation |
| `cmd_app_shortcuts_group_collapse_persists_through_search` | Group collapse state preservation |
| `cmd_app_shortcuts_group_collapse_persists_through_sort` | Group collapse through sort |
| `cmd_app_shortcuts_header_drag_reorders_columns_without_sort` | Header drag reorder |
| `cmd_app_shortcuts_header_resize_changes_visible_width` | Header resize |
| `cmd_app_shortcuts_header_resize_survives_search_roundtrip` | Resize survives search |
| `cmd_app_shortcuts_key_column_uses_natural_key_order` | Key column semantic sort order |
| `cmd_app_shortcuts_left_right_collapse_expand_group` | Left/Right group expand/collapse |
| `cmd_app_shortcuts_live_dx_search_interaction` | DxUi search interaction |
| `cmd_app_shortcuts_long_run_open_close_stays_stable` | Stability |
| `cmd_app_shortcuts_long_run_scrolling_stays_bounded` | Scroll bounds |
| `cmd_app_shortcuts_restores_collapsed_group_state` | Collapsed state restoration |
| `cmd_app_shortcuts_restores_combined_view_state` | Combined view state restore |
| `cmd_app_shortcuts_restores_persisted_grid_layout` | Grid layout persistence |
| `cmd_app_shortcuts_restores_persisted_sort_order` | Sort order persistence |
| `cmd_app_shortcuts_restores_reordered_sorted_grid_layout` | Reordered+sorted layout |
| `cmd_app_shortcuts_row_tooltip_tracks_hovered_cell` | Row tooltip tracking |
| `cmd_app_shortcuts_search_preserves_selection_and_group_semantics` | Search selection semantics |
| `cmd_app_shortcuts_tab_traversal_matches_expected_order` | Tab order |
| `cmd_app_shortcuts_theme_cycle_keeps_grid_legible` | Theme legibility |
| `cmd_app_shortcuts_uses_dxui_surface` | DxUi surface |
| `cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_search_roundtrip` | Reordered+resized copy after search |
| `cmd_app_shortcuts_reordered_resized_columns_survive_sort_cycles` | Reordered+resized layout through sort |
| `cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_sort_cycles` | Reordered+resized copy after sort |
| `cmd_app_shortcuts_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip` | Reordered+resized layout through sort and search |
| `cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_sort_cycles_and_search_roundtrip` | Reordered+resized copy after sort and search |

### 1.4 Connection Manager (29 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_connection_manager_window_access_keys_focus_expected_controls` | Access key focus |
| `cmd_connection_manager_window_clean_external_reload_refreshes_list` | Clean settings hot-reload refresh |
| `cmd_connection_manager_window_close_persists_new_profile` | Close command persistence for newly created profiles |
| `cmd_connection_manager_window_dirty_external_reload_prompts_and_keeps_editing` | Dirty settings hot-reload prompt |
| `cmd_connection_manager_window_escape_from_dx_input_closes_cancel` | Escape from DxUi input |
| `cmd_connection_manager_window_enter_from_dx_input_routes_default_connect` | Enter from DxUi input routes the default Connect command |
| `cmd_connection_manager_window_applies_selected_tool_backdrop` | Shared tool-window backdrop application |
| `cmd_connection_manager_window_live_dx_interaction` | DxUi interaction |
| `cmd_connection_manager_window_masked_secret_accepts_bridge_chars` | Masked secret field accepts WM_CHAR through the DxUi bridge |
| `cmd_connection_manager_window_textfield_doubleclick_selects_word` | Real Connection Manager text field double-click selects a word |
| `cmd_connection_manager_window_long_run_list_scrolling_stays_bounded` | Scroll bounds |
| `cmd_connection_manager_window_long_run_open_close_stays_stable` | Stability |
| `cmd_connection_manager_window_modeless_connect_posts_left_navigation` | Modeless Connect posts left-pane navigation |
| `cmd_connection_manager_window_modeless_connect_posts_right_navigation` | Modeless Connect posts right-pane navigation |
| `cmd_connection_manager_window_pointer_click_toggles_visible_dx_toggle` | Toggle click |
| `cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable` | Protocol churn form/UIA stability |
| `cmd_connection_manager_window_rejects_blank_profile_name` | Blank profile-name validation |
| `cmd_connection_manager_window_rejects_duplicate_profile_name_case_insensitive` | Case-insensitive duplicate profile-name validation |
| `cmd_connection_manager_window_rejects_reserved_quick_profile_name` | Reserved Quick Connect name validation |
| `cmd_connection_manager_window_retired_dialog_files_absent` | Retired dialog file/source/project guard |
| `cmd_connection_manager_window_stale_save_prompts_before_overwrite` | Stale settings save overwrite prompt |
| `cmd_connection_manager_window_tab_traversal_live_dx_interaction` | Tab traversal |
| `cmd_connection_manager_window_theme_cycle_keeps_form_and_selection_legible` | Theme cycle form and selection legibility |
| `cmd_connection_manager_window_trims_profile_name_before_save` | Profile-name trim before save |
| `cmd_connection_manager_window_uses_dxui_command_buttons` | DxUi command buttons |
| `cmd_connection_manager_window_uses_dxui_form_action_buttons` | Form action buttons |
| `cmd_connection_manager_window_uses_dxui_form_inputs` | Form inputs |
| `cmd_connection_manager_window_uses_localized_strings_for_dynamic_labels` | Resource-backed dynamic labels |
| `cmd_connection_manager_window_wm_close_discards_new_profile` | System close/WM_CLOSE prompts before dirty discard |

Connection Manager closeout requires:

- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_credential_prompt_" -TimeoutMultiplier 2`
- Debug and Release `.\build.ps1 -ProjectName RedSalamander`
- A `Specs/TestRuns/.../Commands/...` archive for the full Connection Manager window family with zero failures.

### 1.5 Connection Credential Prompt (8 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_connection_credential_prompt_access_keys_focus_expected_controls` | Access keys |
| `cmd_connection_credential_prompt_dxui_validation_and_accept` | Validation and accept |
| `cmd_connection_credential_prompt_enter_and_escape_route_default_cancel` | Enter/Escape |
| `cmd_connection_credential_prompt_escape_cancels_secret_only` | Secret-only cancel |
| `cmd_connection_credential_prompt_live_dx_interaction` | DxUi interaction |
| `cmd_connection_credential_prompt_long_run_open_close_stays_stable` | Stability |
| `cmd_connection_credential_prompt_pointer_click_toggles_secret_visibility` | Secret visibility toggle |
| `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible` | Theme cycle surface legibility and shared tool-window backdrop application |

### 1.6 Pane Commands (24+ cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_pane_changeAttributes_applies_attributes_removes_streams_and_reports` | Change Attributes selected-item scope, attribute set/clear, stream removal, and report contents |
| `cmd_pane_changeAttributes_options_dialog_uses_dxui_not_win32_template` | Change Attributes DxUi dialog, tri-state cycle back to leave unchanged, date/time rows, Include subdirectories disabled for file-only selection, UIA toggle surface |
| `cmd_pane_changeAttributes_recurse_applies_datetime_with_progress` | Recursive Change Attributes applies date/time to selected folder descendants and creates a File Operations progress task |
| `cmd_pane_changeCase` | Change case command |
| `cmd_pane_copy_text` | Copy text to clipboard |
| `cmd_pane_focusAddressBar_tab_traversal` | Address bar tab traversal |
| `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation` | Pack/Unpack ZIP round trip, sorted entries, empty-directory preservation, overwrite validation, invalid destination handling, unsupported-provider feedback, and archive perf artifact output |
| `cmd_pane_listOpenedFiles_shows_sources_prunes_closed_editors_and_focuses_items` | List Opened Files dialog empty state, viewer/editor/preview source rows, closed external-editor pruning, focus navigation, and perf artifact output |
| `cmd_pane_navigationView_full_path_popup_edit_route` | Navigation full path editing |
| `cmd_pane_navigationView_history_dropdown_keyboard_navigation` | History dropdown navigation |
| `cmd_pane_navigationView_path_doubleClick_enters_edit_mode` | Path double-click edit |
| `cmd_pane_navigationView_region_tab_traversal` | Navigation region tab order |
| `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane` | NavigationView actions in an unfocused pane activate that pane, navigate the clicked target, and return keyboard focus to that pane's FolderView |
| `cmd_pane_navigation_ambient_escape_returns_focus_to_active_folder_view` | Escape from main-window chrome restores active pane FolderView focus without clearing selection |
| `cmd_pane_navigation_menu_escape_returns_focus_to_active_folder_view` | Escape from the keyboard-owned pane menu restores active pane FolderView focus without clearing selection |
| `cmd_pane_navigation_nonstandard_menu_common_folders` | Nonstandard file-system NavigationView menus expose a `Common Folders` submenu with local known-folder rows and stock icons |
| `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable` | Status bar focused-item detail updates, typography fit/no-clipping guard, and inactive-pane dimming |
| `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu` | Owned pane status-bar surface, no retained native font assignment, DxUI sort popup opening, readable popup debug state, and above-status-bar popup placement |
| `cmd_pane_navigation_directory_impact_preserves_selection` | Directory-impact refresh preserves surviving selection, drops deleted names, follows same-folder rename chains of depth 3+, and keeps renamed unselected originals unselected |
| `cmd_pane_refresh` | Refresh command |
| `cmd_pane_selection_goto_selected_name` | Go to selected name |
| `cmd_pane_selection_hide_names` | Hide selected names |
| `cmd_pane_selection_invert` | Invert selection |
| `cmd_pane_shares_shows_synthetic_rows_opens_paths_and_reports_access_denied` | Shared Directories sorted synthetic rows, open-path navigation, access-denied empty/error state, shortcut dispatch, and perf artifact output |

### 1.7 Find Files Dialog (46 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_pane_find_dialog_access_keys_focus_expected_fields` | Access key focus |
| `cmd_pane_find_dialog_action_buttons_activate_expected_commands` | Action button dispatch |
| `cmd_pane_find_dialog_command_enablement_matches_idle_running_and_selection_states` | Command state management |
| `cmd_pane_find_dialog_copy_follows_reordered_columns` | Copy follows column order |
| `cmd_pane_find_dialog_directory_activation_navigates_into_selection` | Directory navigation |
| `cmd_pane_find_dialog_enter_from_checkbox_invokes_default_search` | Checkbox Enter behavior |
| `cmd_pane_find_dialog_escape_closes_popup_before_cancel` | Escape priority |
| `cmd_pane_find_dialog_escape_from_dx_control_closes_cancel` | DxUi Escape handling |
| `cmd_pane_find_dialog_exposes_live_uia_selection_and_inputs` | UIA accessibility |
| `cmd_pane_find_dialog_failure_status_is_readable` | Failure status display |
| `cmd_pane_find_dialog_grid_doubleClick_activates_selection` | Double-click activation |
| `cmd_pane_find_dialog_grid_enter_activates_selection` | Enter activation |
| `cmd_pane_find_dialog_header_click_sorts_results` | Header click sort |
| `cmd_pane_find_dialog_header_drag_reorders_columns_without_sort` | Header drag reorder |
| `cmd_pane_find_dialog_header_resize_changes_visible_width` | Header resize |
| `cmd_pane_find_dialog_large_local_search_uses_incremental_updates` | Incremental search updates |
| `cmd_pane_find_dialog_local_root_overrides_stale_context` | Local root override |
| `cmd_pane_find_dialog_long_run_open_close_stays_stable` | Stability |
| `cmd_pane_find_dialog_long_run_scrolling_stays_bounded` | Scroll bounds |
| `cmd_pane_find_dialog_mode_typeahead_updates_selection_and_dependencies` | Typeahead mode |
| `cmd_pane_find_dialog_open_parent_keeps_directory_focused_in_parent` | Open parent focus |
| `cmd_pane_find_dialog_pointer_click_toggles_recursive_checkbox` | Recursive toggle |
| `cmd_pane_find_dialog_reordered_columns_survive_search_rerun` | Column order persistence |
| `cmd_pane_find_dialog_reordered_columns_survive_sort_cycles` | Column order through sort |
| `cmd_pane_find_dialog_reordered_resized_columns_survive_search_rerun` | Reorder+resize persistence |
| `cmd_pane_find_dialog_reordered_resized_columns_survive_sort_cycles` | Reorder+resize through sort |
| `cmd_pane_find_dialog_resized_columns_survive_search_rerun` | Resize persistence |
| `cmd_pane_find_dialog_resized_columns_survive_sort_cycles` | Resize through sort |
| `cmd_pane_find_dialog_restored_combined_view_state_*` | (10 sub-cases) Combined view state restoration |
| `cmd_pane_find_dialog_restores_combined_view_state` | Combined view state |
| `cmd_pane_find_dialog_restores_persisted_grid_layout` | Grid layout persistence |
| `cmd_pane_find_dialog_restores_persisted_sort_order` | Sort order persistence |
| `cmd_pane_find_dialog_restores_reordered_grid_layout` | Reordered layout |
| `cmd_pane_find_dialog_restores_reordered_sorted_grid_layout` | Reordered+sorted layout |
| `cmd_pane_find_dialog_restores_resized_grid_layout` | Resized layout |
| `cmd_pane_find_dialog_running_status_shows_phase_and_path` | Running status display |
| `cmd_pane_find_dialog_search_ops` | Search operations |
| `cmd_pane_find_dialog_service_status_shows_backend_diagnostics` | Service status |
| `cmd_pane_find_dialog_service_unavailable_warning_is_distinct` | Service warning |
| `cmd_pane_find_dialog_tab_traversal_matches_expected_order` | Tab order |
| `cmd_pane_find_dialog_theme_cycle_keeps_grid_legible` | Theme legibility |
| `cmd_pane_find_dialog_uses_dxui_host_without_visible_child_controls` | DxUi host |

### 1.8 FileOps Issues Pane and Speed Limit Prompt (21 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_pane_fileops_issues_pane_copy_follows_reordered_columns` | Copy follows column order |
| `cmd_pane_fileops_issues_pane_diff_refresh_preserves_selection` | Selection preservation |
| `cmd_pane_fileops_issues_pane_exposes_live_uia_selection` | UIA accessibility |
| `cmd_pane_fileops_issues_pane_header_click_sorts_results` | Header sort |
| `cmd_pane_fileops_issues_pane_header_drag_reorders_columns_without_sort` | Column reorder |
| `cmd_pane_fileops_issues_pane_header_resize_changes_visible_width` | Column resize |
| `cmd_pane_fileops_issues_pane_long_run_open_close_stays_stable` | Stability |
| `cmd_pane_fileops_issues_pane_long_run_scrolling_stays_bounded` | Scroll bounds |
| `cmd_pane_fileops_issues_pane_reordered_columns_survive_sort_cycles` | Column order through sort |
| `cmd_pane_fileops_issues_pane_reordered_copy_follows_visible_columns_after_sort_cycles` | Reordered copy after sort |
| `cmd_pane_fileops_issues_pane_reordered_resized_columns_survive_sort_cycles` | Reorder+resize through sort |
| `cmd_pane_fileops_issues_pane_reordered_resized_copy_follows_visible_columns_after_sort_cycles` | Reorder+resize copy after sort |
| `cmd_pane_fileops_issues_pane_restores_combined_view_state_after_recreate` | Combined view state restore |
| `cmd_pane_fileops_issues_pane_resized_columns_survive_sort_cycles` | Resize through sort |
| `cmd_pane_fileops_issues_pane_tab_keeps_grid_focus` | Tab focus |
| `cmd_pane_fileops_issues_pane_theme_cycle_keeps_grid_legible` | Theme legibility |
| `cmd_pane_fileops_issues_pane_uses_dxui_host_without_visible_child_controls` | DxUi host and shared tool-window backdrop application |
| `cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface` | Speed-limit prompt DxUi surface, progress popup shared tool-window backdrop application, and file-operations popup caption glyph DirectWrite guard |
| `cmd_pane_fileops_speedLimit_prompt_live_dx_interaction` | Speed-limit prompt live input |
| `cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable` | Speed-limit prompt stability |
| `cmd_pane_fileops_speedLimit_prompt_keeps_navigation_shell_stable` | Speed-limit prompt navigation-shell stability |

### 1.9 Preferences Dialog (~100 cases)

Covers every preference page with DxUi interaction, tab traversal, roundtrip restore,
and accessibility tests. Pages: General, Panes, Viewers, Editors, Mouse, Keyboard,
Themes, Plugins, Compare Directories, Hot Paths, File Operations, Advanced.

Key coverage patterns per page:
- `*_live_dx_interaction` — DxUi surface responds to input
- `*_tab_traversal_live_dx_interaction` — Tab order correct
- `*_roundtrip_restores_dxui_surface` — Settings persist through dialog reopen
- `*_page_uses_dxui_*` — Correct DxUi control types used
- `cmd_preferences_dialog_general_window_backdrop_apply_updates_supported_windows` — Preferences HWND applies the selected shared tool-window backdrop without adding a main-window acceptance assertion
- `cmd_preferences_dialog_general_page_uses_dxui_toggle_cards` — General page uses DxUi toggle/card hosts and guards HFONT-free layout through `generalUsesDxUiTypographyContext` and `generalUsesDxUiTypographyMetrics`
- `cmd_preferences_dialog_panes_page_uses_dxui_statics_and_toggles` — Panes page uses DxUi cards/toggles/combos and guards HFONT-free layout through `panesUsesDxUiTypographyContext` and `panesUsesDxUiTypographyMetrics`
- `cmd_preferences_dialog_viewers_page_uses_dxui_combo_and_button_chrome` — Viewers page uses DxUi combo/button/grid chrome and guards HFONT-free layout through `viewersUsesDxUiTypographyContext` and `viewersUsesDxUiTypographyMetrics`
- `cmd_preferences_dialog_opens_with_french_satellite_resources` — Preferences opens from the `fr-FR` satellite dialog template while still creating executable-owned custom dialog child windows
- `*_long_run_*_stays_bounded` — No resource leaks under repeated use
- `*_page_exposes_live_uia_*` — UIAutomation accessibility
- `*_search_*` — Search/filter functionality
- `*_pointer_click_*` — Mouse interaction

### 1.10 Plugin Configuration Dialog And Settings (13 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_plugin_configuration_dialog_access_keys_route_expected_actions` | Access keys |
| `cmd_plugin_configuration_dialog_enter_and_escape_route_default_cancel` | Enter/Escape |
| `cmd_plugin_configuration_dialog_live_dx_interaction` | DxUi interaction |
| `cmd_plugin_configuration_dialog_long_run_open_close_stays_stable` | Stability |
| `cmd_plugin_configuration_dialog_long_run_scrolling_keeps_dx_surface_stable` | Scroll stability |
| `cmd_plugin_configuration_dialog_pointer_click_toggles_visible_dx_toggle` | Toggle click |
| `cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction` | Tab traversal |
| `cmd_plugin_configuration_dialog_uses_dxui_command_buttons` | DxUi buttons |
| `cmd_plugin_configuration_dialog_uses_dxui_form_surface` | DxUi form |
| `settings_file_system_plugin_roundtrip` | File-system plugin settings roundtrip |
| `settings_viewer_text_plugin_roundtrip` | ViewerText plugin settings roundtrip, including rejection of transient diff-UI persistence keys |
| `viewer_text_diff_perf` | ViewerText diff perf baseline including theme-driven semantic row paint, clickable hidden-banner reveal, built-in rainbow-mode theme-switch repaint, parsed hunk-jump latency, bounded viewport growth, unresolved placeholder bands, and backtrack cache reuse |
| `viewer_text_hex_byte_color_perf` | ViewerText hex byte color perf baseline |

### 1.11 Compare Directories Options And Progress (11 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons` | DxUi menu bar, banner buttons, banner title/progress text, no visible legacy banner text, no native font state |
| `cmd_compare_directories_options_access_keys_focus_expected_controls` | Access keys |
| `cmd_compare_directories_options_enter_and_escape_route_default_cancel` | Enter/Escape |
| `cmd_compare_directories_options_live_dx_body_interaction` | DxUi body |
| `cmd_compare_directories_options_long_run_open_close_stays_stable` | Stability |
| `cmd_compare_directories_options_pointer_click_toggles_live_dx_interaction` | Toggle click |
| `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable` | Scroll stability |
| `cmd_compare_directories_options_tab_traversal_live_dx_interaction` | Tab traversal |
| `cmd_compare_directories_options_theme_cycle_keeps_surface_legible` | Theme legibility |
| `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics` | DxUi labels, zero visible native body/footer controls, and DirectWrite options typography metrics |
| `cmd_compare_directories_progress_perf` | Compare progress performance |

### 1.12 Settings and Infrastructure (26+ cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `folderView_empty_folder_state` | Empty folder centered state plus row-sized focused `Go to parent` placeholder item |
| `folderView_filter_watermark_empty_state` | Filter watermark display |
| `folderView_perf_large_folder_baseline` | Large folder performance baseline |
| `file_action_resolution_v16_action_ids_are_case_insensitive` | File-action resolver matches action IDs case-insensitively, preserves action-definition casing, and collapses case-only references |
| `help_menu_links_external_documentation` | Help menu external documentation command placement and registry binding |
| `icon_bitmap_alpha_normalization` | Icon alpha normalization including premultiply and AND-mask transparency semantics |
| `mask_syntax_wildcards` | Wildcard mask syntax parsing |
| `menu_copy_text_group_contract` | Menu copy text contract |
| `menu_load_selection_links_restore` | Selection links restoration |
| `modeless_window_ownership` | Modeless window lifecycle |
| `navigation_location_edit_input_expands_environment_variables` | Environment variable expansion |
| `red_salamander_help_lists_diagnostics_options` | Help text documents Release diagnostics ETW and perf JSONL switches |
| `registry_integrity` | Registry settings integrity; every command has a non-empty short function-bar label with guarded examples (`MakeDir`, `UsrMenu`, `ByTime`) |
| `resource_hresult_details_format_is_valid` | Localized HRESULT details resource uses valid positional `std::format` placeholders |
| `resource_invalid_format_string_returns_raw_fallback` | Runtime resource formatting logs the failing resource ID/detail and degrades to the raw localized resource text instead of throwing or returning blank text when a localized string has invalid `std::format` syntax |
| `resource_format_placeholders_are_positional` | Product `.rc` resources reject bare `{}` and unindexed `std::format` specs while allowing documented literal file-action macros |
| `search_local_index_stream_stop_after_first` | Search stream stop semantics |
| `settings_file_operations_precalc_roundtrip` | Pre-calc settings roundtrip |
| `settings_file_system_plugin_roundtrip` | Plugin settings roundtrip |
| `settings_hot_reload_*` | (4 cases) Hot reload merge/suppression |
| `settings_shortcuts_*` | (2 cases) Shortcut settings roundtrip |
| `settings_store_search_roundtrip` | Search settings roundtrip |
| `pane_view_options_toggle_preview_pane_tabs_and_selection` | Preview pane tab-strip pointer clicks switch Folder/Preview, selected/hovered Preview close-glyph visibility, delayed Folder tab path tooltip, Preview close glyph closes preview mode, old embedded text content is cleared before rendering the next focused item, and source-pane focus is preserved |
| `pane_view_options_preview_uses_configured_embedded_viewer_and_preserves_focus` | Preview uses configured embedded viewers, keeps the same embedded instance and HWND across same-plugin image/media focus changes, keeps media-to-media and media-to-image switches responsive while VLC stop/release is slow, verifies VLC video child parenting after video-to-video preview navigation, forwards wheel seek from VLC child surfaces, preserves source-pane focus, and persists VLC preview volume/mute state |
| `pane_view_options_preview_uses_builtin_embedded_viewer_with_empty_associations` | Preview consults built-in embedded viewer defaults when saved viewer associations are empty before falling back to ViewerText |
| `shortcut_defaults_mapping` | Default shortcut mappings |
| `shortcut_functionbar_dispatch_refresh` | Function bar dispatch and refresh using command short labels |

---

## 2. CompareDirectories Suite (`--compare-selftest`)

**Source:** `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp` coordinator + 3 included case files (141 cases)

Tests the Compare Directories engine, search backends, SQLite index store,
crash quarantine, OAuth, and remote storage comparisons.

### 2.1 Core Compare Engine (35 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `accessors` | Engine accessor methods |
| `attributes` | File attribute comparison |
| `baseInterfaces` | Base interface contract |
| `cancel_completes_bounded` | Cancellation bounded completion |
| `content` | Content comparison |
| `content short reads` | Short-read content handling |
| `content_dual_io` | Dual I/O content paths |
| `content_inflight_stamp_guards_restart` | In-flight stamp restart guard |
| `content_no_io_disables_compareContent` | No-I/O content disable |
| `content_pending_elided` | Pending content elision |
| `content_queue_bounded_hi_lo` | Content queue bounds |
| `content_size_mismatch_no_pending` | Size mismatch handling |
| `contentCacheHit` | Content cache hit path |
| `concurrent_get_or_compute_decision` | Concurrent decision computation |
| `decision_cache_eviction_budget_pins_visible` | Cache eviction budget |
| `decisionUpdatedCallback` | Decision update callback |
| `deep_tree` | Deep tree traversal |
| `dircache_not_polluted_by_compare_scan` | Directory cache isolation |
| `dummy_content` | Dummy plugin content compare |
| `empty_directories` | Empty directory handling |
| `ignore` | Ignore pattern matching |
| `ignore_multiple_patterns` | Multiple ignore patterns |
| `ignore_pattern_count_cap` | Pattern count cap |
| `ignore_pattern_length_cap` | Pattern length cap |
| `ignore_wildcard_pathology_runtime_bound` | Wildcard pathology guard |
| `invalid_directory_entry_buffer` | Invalid entry buffer handling |
| `invalidate` | Invalidation |
| `invalidateForPath` | Path-specific invalidation |
| `missing folder` | Missing folder handling |
| `no_sync_deep_scan` | No-sync deep scan |
| `reparse` | Reparse point handling |
| `scan_inflight_stamp_guards_restart` | Scan restart guard |
| `setCompareEnabled` | Compare enable/disable |
| `setSettingsInvalidates` | Settings invalidation |
| `showIdentical` | Identical file display |
| `size` | Size comparison |
| `subdir pending` | Subdirectory pending state |
| `subdirattrs` | Subdirectory attributes |
| `subdirs` | Subdirectory comparison |
| `time` | Timestamp comparison |
| `typemismatch` | Type mismatch detection |
| `uiVersion` | UI version tracking |
| `unicode_filenames` | Unicode filename support |
| `unique` | Unique file detection |
| `zero_vs_nonzero_content` | Zero/non-zero content |
| `zeroByteContent` | Zero-byte file content |

### 2.2 Directory Size Callbacks (3 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `directory_size_local_callback_contract` | Local plugin directory size callback |
| `directory_size_dummy_callback_contract` | Dummy plugin directory size callback |
| `directory_size_7z_callback_contract` | 7z plugin directory size callback |

### 2.3 Search — Local Native (12 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `local_search_callback_contract` | Search callback contract |
| `local_search_content_literal` | Literal content search |
| `local_search_invalid_query_rejected` | Invalid query rejection |
| `local_search_name_and_content_and_semantics` | Name+content search |
| `local_search_name_wildcard_recursive` | Wildcard recursive search |
| `local_search_name_windows_filesystem_case_parity` | Case parity with Windows |
| `local_search_native_matches_host_fallback` | Native vs fallback parity |
| `local_search_native_unicode_long_path_matches_host_fallback` | Unicode long path parity |
| `local_search_qi_and_capabilities` | QI and capabilities |
| `local_search_scan_follow_symlink_loop_guard` | Symlink loop guard |
| `local_search_scan_wide_tree_parallel_walk_name_only` | Parallel walk |
| `local_search_backend_preferences_roundtrip` | Backend preferences |

### 2.4 Search — Host Fallback (5 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `host_fallback_search_7z_name_only` | 7z fallback search |
| `host_fallback_search_access_denied_warning` | Access denied handling |
| `host_fallback_search_content_degraded_without_io` | Content degradation |
| `host_fallback_search_dummy_name_only` | Dummy fallback search |
| `host_fallback_search_local_plugin_path_root` | Local plugin path root |
| `host_fallback_search_short_read_and_cancel` | Short read and cancel |

### 2.5 Search — Service and Indexed (30+ cases)

Covers the search service binary CLI, SQLite bootstrap, query/status roundtrip,
multi-client scenarios, rebuild control, cold start, stale root refresh, prefilter,
journal replay, snapshot reload, corruption rebuild, maintenance, and more.

### 2.6 Search Text Helpers (2 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `search_text_helpers_chunk_overlap_literal_and_regex` | Chunk overlap matching |
| `search_text_helpers_decoding_and_binary` | Text decoding and binary detection |

### 2.7 SQLite Index Store (6 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `sqlite_index_store_automatic_checkpoint_truncates_wal` | WAL checkpoint |
| `sqlite_index_store_automatic_compaction_is_bounded` | Compaction bounds |
| `sqlite_index_store_bootstrap_creates_schema` | Schema bootstrap |
| `sqlite_index_store_load_and_apply_journal_delta` | Journal delta |
| `sqlite_index_store_manual_compaction_reclaims_space` | Manual compaction |
| `sqlite_index_store_upgrade_paths` | Schema upgrade |

### 2.8 Infrastructure and Security (8 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `crash_quarantine_synthetic_marker` | Crash quarantine |
| `oauth_authmode_roundtrip` | OAuth auth mode roundtrip |
| `oauth_refresh_token_storage` | OAuth token storage |
| `plugin_path_math` | Plugin path mathematics |
| `try_make_relative_outside_root` | Path relativization |
| `windows_hello_cache` | Windows Hello cache |
| `connection_display_url` | Connection display URL |
| `google_drive_*` | (4 cases) Google Drive plugin contract |
| `onedrive_personal_cleared_client_id_requires_configuration` | OneDrive client ID |

### 2.9 Remote Smoke Tests (Conditional)

These cases skip when connection profiles or secrets are absent.
See `Specs/Testing/Testing_SelfTestRemoteCredentials.md`.

| Case Name | Coverage Area |
|-----------|---------------|
| `remote_file_s3` | S3 compare smoke |
| `remote_file_ftp` | FTP compare smoke |
| `remote_file_onedrive_personal` | OneDrive Personal compare |
| `remote_file_onedrive_business` | OneDrive Business compare |
| `remote_file_sharepoint` | SharePoint compare |
| `remote_s3_pagination` | S3 pagination |
| `remote_*_directory_size_callback_contract` | Remote directory size |
| `remote_ftp_continue_on_error_partial` | FTP error continuation |
| `remote_s3_metadata_smoke` | S3 metadata |
| `remote_s3_delete_missing` | S3 delete missing |

---

## 3. FileOperations Suite (`--fileops-selftest`)

**Source:** `RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp` coordinator + 4 included phase files (68 phases)

Tests file operations (copy, move, delete, rename) using a tick-driven async state machine.
Each phase represents a test case that exercises one aspect of the file operations pipeline.

### 3.1 Phase 5 — Pre-Calculation and Queue (7 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase5_PreCalcSettingsApplied` | Pre-calc settings application |
| `Phase5_PreCalcCancelReleasesSlot` | Cancel releases pre-calc slot |
| `Phase5_PreCalcCancelLatencyLocal` | Cancel latency measurement |
| `Phase5_PreCalcSkipContinues` | Skip continues operation |
| `Phase5_CancelQueuedTask` | Queued task cancellation |
| `Phase5_SwitchParallelToWaitDuringPreCalc` | Parallel to wait mode switch |
| `Phase5_SwitchWaitToParallelResume` | Wait to parallel resume |

### 3.2 Phase 6 — Popup and Bandwidth (4 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase6_PopupSmokeResizeAndPause` | Popup resize and pause UI |
| `Phase6_DeleteBytesMeaningful` | Delete bytes tracking |
| `Phase6_LocalBandwidthThrottle` | Local bandwidth throttling |
| `Phase6_ParallelBandwidthThrottleFairness` | Parallel bandwidth fairness |

### 3.3 Phase 7 — Watcher, Concurrency, and Batching (16 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase7_WatcherChurn` | Directory watcher churn |
| `Phase7_CacheBorrowNoWatchInvalidation` | Cache borrow without watch invalidation |
| `Phase7_CrossPaneVisibleRefreshLocal` | Cross-pane refresh (local) |
| `Phase7_CrossPaneVisibleRefreshDummy` | Cross-pane refresh (dummy) |
| `Phase7_CrossPaneRelocateLocal` | Cross-pane relocate |
| `Phase7_LargeDirectoryEnumeration` | Large directory enumeration |
| `Phase7_ParallelCopyMoveKnobs` | Parallel copy/move knobs |
| `Phase7_CopyRecursiveParallelismMatrix` | Recursive copy/move matrix including copied reparse items, nested concurrency 1, forced and optional real cross-volume move fallback, partial error, and active-worker cancellation |
| `Phase7_CopyMoveConcurrency16Perf` | 16-thread concurrency performance |
| `Phase7_AutoConcurrencyHints` | Auto concurrency hints |
| `Phase7_PerItemDirectoryCopyInFlightLines` | Per-item directory copy in-flight |
| `Phase7_SharedPerItemScheduler` | Shared per-item scheduler |
| `Phase7_ParallelDeleteKnobs` | Parallel delete knobs |
| `Phase7_RecycleBinBatchDelete` | Recycle bin batch delete |
| `Phase7_RecycleBinBatchDeleteMultiBatch` | Multi-batch recycle bin delete |

### 3.4 Phase 8 — Defaults and Validation (5 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase8_DefaultBandwidthLimitFromSettings` | Default bandwidth limit |
| `Phase8_TightDefaults_NoOverwrite` | Tight defaults without overwrite |
| `Phase8_InvalidDestinationRejected` | Invalid destination rejection |
| `Phase8_InvalidSizeBytesRejected` | Invalid size rejection |
| `Phase8_PerItemOrchestration` | Per-item orchestration |

### 3.5 Phase 9 — Conflict Prompts (7 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase9_ConflictPrompt_OverwriteReplaceReadonly` | Overwrite readonly |
| `Phase9_ConflictPrompt_ApplyToAllUiCache` | Apply-to-all UI cache |
| `Phase9_ConflictPrompt_OverwriteAutoCap` | Overwrite auto cap |
| `Phase9_ConflictPrompt_SkipAll` | Skip all |
| `Phase9_ConflictPrompt_RetryCap` | Retry cap |
| `Phase9_ConflictPrompt_SkipContinuesDirectoryCopy` | Skip continues directory copy |
| `Phase9_PerItemConcurrency` | Per-item concurrency |

### 3.6 Phase 10–15 — Advanced Operations (12 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase10_PermanentDelete` | Permanent delete confirmation, cancellation guard, and confirmed execution |
| `Phase11_CrossFileSystemBridge` | Cross-filesystem bridge |
| `Phase11_BridgeSingleFolderParallelCopyInFlightLines` | Bridge single-folder parallel |
| `Phase11_BridgeMultiFolderParallelCopyInFlightLines` | Bridge multi-folder parallel |
| `Phase11_BridgePipelineDummyToDummyPerf` | Bridge pipeline performance |
| `Phase11_ConnectionOverridePrecedence` | Connection override precedence |
| `Phase11_ConnectionOverrideGlobalGate` | Connection override global gate |
| `Phase11_ConnectionOverrideClamp` | Connection override clamp |
| `Phase12_ReparsePointPolicy` | Reparse point policy |
| `Phase13_PostMortemDiagnostics` | Post-mortem diagnostics |
| `Phase14_PopupHostLifetimeGuard` | Popup host lifetime and reentrant visibility/placement guard |
| `Phase15_FileSystem7zReadSeekSmoke` | 7z read/seek smoke |
| `Phase15_FileSystem7zMountPathImpact` | 7z mount path impact |

### 3.7 Phase 16 — Remote Storage (Conditional, 16 cases)

These cases skip when connection profiles or secrets are absent.

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase16_RemoteWatchContractExposure` | Remote watch contract |
| `Phase16_RemoteFtpSecret` | FTP secret validation |
| `Phase16_RemoteFtpSandbox` | FTP sandbox operations |
| `Phase16_RemoteSftpSecret` | SFTP secret validation |
| `Phase16_RemoteSftpSandbox` | SFTP sandbox operations |
| `Phase16_RemoteScpSecret` | SCP secret validation |
| `Phase16_RemoteScpSandbox` | SCP sandbox operations |
| `Phase16_RemoteImapSecret` | IMAP secret validation |
| `Phase16_RemoteImapSandbox` | IMAP sandbox operations |
| `Phase16_RemoteS3Secret` | S3 secret validation |
| `Phase16_RemoteS3Sandbox` | S3 sandbox operations |
| `Phase16_RemoteS3FileOps` | S3 file operations |
| `Phase16_RemoteOneDrivePersonalSecret` | OneDrive Personal secret |
| `Phase16_RemoteOneDrivePersonalSandbox` | OneDrive Personal sandbox |
| `Phase16_RemoteOneDrivePersonalFileOps` | OneDrive Personal file ops |
| `Phase16_RemoteOneDriveBusinessSecret` | OneDrive Business secret |
| `Phase16_RemoteOneDriveBusinessSandbox` | OneDrive Business sandbox |
| `Phase16_RemoteSharePointSecret` | SharePoint secret |
| `Phase16_RemoteSharePointSandbox` | SharePoint sandbox |

### 3.8 Cleanup (1 case)

| Case Name | Coverage Area |
|-----------|---------------|
| `Cleanup_RestorePluginConfig` | Restore plugin configuration after test |

---

## 4. Performance Tests (`PerformanceTests2/`)

**Framework:** Microsoft CppUnitTest (DLL)

| Test Class / File | Cases | Coverage Area |
|-------------------|-------|---------------|
| `FolderIconEnumerationPerfTest` | 1+ | Icon enumeration caching under load |
| `FolderIconEnumerationDuplicatePathPerfTest` | 1+ | Duplicate path icon edge cases |
| `FolderViewRefreshDuplicatePathPerfTest` | 1+ | FolderView refresh with duplicate paths |

---

## 5. PoC Tests

| Project | Cases | Coverage Area |
|---------|-------|---------------|
| **DxUiTests** | ~50+ | DxUi color parsing, theme rendering, control creation, HSL/RGB conversion, submenu cascade hover timing, single-line text selection clipping, selected TextField emoji color-font rendering, text-input bridge keyboard forwarding, FolderView incremental-search helper behavior, inactive-pane visual-state helpers, and empty-folder placeholder layout metrics |
| **LocalizationTests** | ~5 | Resource owner registration, satellite string/menu/dialog lookup, localized dialog templates with executable-owned custom child classes, fallback to embedded resources, and persisted `ui.language` roundtrips |
| **ViewerPETests** | ~20+ | PE viewer plugin, image viewer, text viewer including AppTheme-driven parsed diff semantic colors with runtime theme switching, explicit rainbow-mode coverage, and high-contrast coverage, base-background unchanged rows plus dim diff-marker metadata, clickable hidden-context banners in hunks-only diff mode, non-anchor parsed hunk presentation, pane-local side-by-side visual-layout metadata and visible split-row counters for parsed diff viewports, split-row top-visible text snapshot coverage after hunk navigation, parsed hunk count and active-hunk snapshot metadata, next/previous hunk navigation, diff presentation, lazy referenced-file expansion, parsed-document reuse across diff variants, active-section-only unchanged-text hydration with on-demand section jumps, viewport-windowed unchanged-row rehydration on scroll, range-bounded referenced-file reads for viewport-nearby unchanged context, cached reuse when revisiting already hydrated viewport ranges, referenced-file content reuse across expanded layouts, hatched placeholder-gap metadata for unresolved rows, diff section navigation, horizontal scrolling in both parsed diff layouts after wrap is disabled including side-by-side to inline presentation switches, shared viewer combo-host popup expansion/collapse and clean close behavior across `ViewerPE`, `ViewerWeb`, `ViewerImgRaw`, and `ViewerText`, larger fully buffered diff parsing including promotion beyond the normal text buffer size for both extension-recognized and header-sniffed diffs, raw streamed multi-file section indexing/navigation beyond the fully buffered parse cap, and parser-fallback coverage, space viewer, VLC viewer |
| **ViewerSqliteTests** | ~15+ | SQLite engine, query execution, schema inspection, UI rendering |
| **MonitorTest** | ~5 | ETW TraceLogging provider emit/receive validation plus compile-time/runtime guards that invalid rectangle visualization remains opt-in, normal Debug and Release monitor self diagnostics suppress Info/Perf output unless the runtime ETW flag is enabled, and normal monitor display rejects self-originated ETW events |

---

## 6. Tooling Script Tests

| Test File | Coverage Area |
|-----------|---------------|
| `Tools\Tests\VcpkgInstallSafety.Tests.ps1` | vcpkg triplet leaf-name validation and staging/install child path containment |

---

## Coverage Matrix

| Area | Commands | Compare | FileOps | Perf | PoC |
|------|----------|---------|---------|------|-----|
| **UI Dialogs** | ✅ 200+ | — | — | — | — |
| **DxUi Framework** | ✅ | — | — | — | ✅ |
| **Preferences** | ✅ ~100 | — | — | — | — |
| **Shortcuts** | ✅ 30 | — | — | — | — |
| **Navigation** | ✅ 15 | — | — | — | — |
| **Find/Search UI** | ✅ 46 | — | — | — | — |
| **Compare Engine** | — | ✅ 35+ | — | — | — |
| **Search Backends** | — | ✅ 50+ | — | — | — |
| **SQLite Index** | — | ✅ 6 | — | — | — |
| **Search Service** | — | ✅ 30+ | — | — | — |
| **File Copy/Move** | — | — | ✅ 30+ | — | — |
| **File Delete** | — | — | ✅ 10+ | — | — |
| **Conflict Prompts** | — | — | ✅ 7 | — | — |
| **Cross-FS Bridge** | — | — | ✅ 8 | — | — |
| **Remote Storage** | — | ✅ 10 | ✅ 19 | — | — |
| **Bandwidth Throttle** | — | — | ✅ 4 | — | — |
| **Icon Cache** | — | — | — | ✅ | — |
| **Folder View Perf** | — | — | — | ✅ | — |
| **ETW Tracing** | — | — | — | — | ✅ |
| **Viewer Plugins** | — | — | — | — | ✅ |
| **Settings Roundtrip** | ✅ 10 | ✅ 2 | — | — | — |
| **Theme System** | ✅ 20+ | — | — | — | ✅ |
| **OAuth/Credentials** | — | ✅ 5 | — | — | — |
