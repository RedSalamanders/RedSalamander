# Visible Typography Audit

Last updated: 2026-04-26

This document tracks the remaining visible typography and visible GDI-text drift in app-owned surfaces. Refresh the inventory with `Tools/Get-VisibleTypographyAudit.ps1`.

## Shared Contract

- Windows 11 runtime families:
  - `Segoe UI Variable Small` for caption/header-scale text
  - `Segoe UI Variable Text` for normal body and control text
  - `Segoe UI Variable Display` for large display text
- Exceptions:
  - `Segoe Fluent Icons` for icon glyphs
  - `Segoe UI Emoji` for emoji glyphs
  - `Consolas` for monospace
- App-owned visible surfaces must route typography through `Common/DxUi/DxUi.Typography.h`.

## Converted Or Retired In This Pass

- `Common/DxUi/DxUi.WindowHost.cpp`
- `Common/DxUi/DxUi.TextInput.cpp`
- `RedSalamander/FolderView.Rendering.cpp`
- `RedSalamander/NavigationView.Rendering.cpp`
- `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- `RedSalamander/Ui/AlertOverlay.h`
- `RedSalamander/NavigationView.cpp`
- `RedSalamander/FunctionBar.cpp`
- `RedSalamander/FolderWindow.StatusBar.cpp`
- `RedSalamander/Preferences.Dialog.cpp`
- `RedSalamander/ConnectionManagerDialog.cpp`
- `RedSalamander/ManagePluginsDialog.cpp`
- `RedSalamander/CompareDirectoriesWindow.cpp`
- `RedSalamander/CompareDirectoriesWindow.Options.cpp`
- `RedSalamander/FolderWindow.cpp`
- `RedSalamander/FolderWindow.FileSystem.cpp`
- `RedSalamander/Preferences.Plugins.cpp`
- `RedSalamander/ThemedControls.cpp` (retired; the interim `Win32UiHelpers` staging module was later retired by `UI_RemainingWin32UiDependencyRetirementPlan.md`)
- `Plugins/ViewerPE/ViewerPE.cpp`
- `Plugins/ViewerVLC/ViewerVLC.cpp`
- `Plugins/ViewerWeb/ViewerWeb.cpp`
- `Plugins/ViewerSpace/ViewerSpace.cpp`
- `Plugins/ViewerImgRaw/ViewerImgRaw.cpp`
- `Plugins/ViewerText/ViewerText.cpp`

## Remaining Visible Drift Buckets

### Menu and owner-draw measurement status

The fresh audit no longer reports active owner-draw menu measurement in `Plugins/ViewerImgRaw/ViewerImgRaw.cpp` or `Plugins/ViewerSpace/ViewerSpace.cpp`. The hidden main-menu owner-draw path in `RedSalamander/RedSalamander.cpp` and the dead viewer owner-draw menu-theme code in `Plugins/ViewerWeb/ViewerWeb.cpp` / `Plugins/ViewerText/ViewerText.MenuTheme.cpp` are also retired.

The strict active-menu source search archived on 2026-04-23 records zero product/test hits for `NONCLIENTMETRICS`, `lfMenuFont`, `GetTextExtentPoint32W`, `GetTextMetricsW`, and `CreateMenuFontForDpi` across `RedSalamander`, `Plugins`, `Common`, and `Tests`. The archived proof lives in `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_171000_segoe_variable_gdi_retirement_final/`.

For `RedSalamander/CompareDirectoriesWindow.Options.cpp`, the hidden owner-draw footer buttons and hidden owner-draw toggle controls behind the visible DxUi options surface are now retired. The full `cmd_compare_directories_options_` family is green after the deterministic scroll-stability fix, so that area is no longer blocked on a visible owner-draw or scroll-settling seam.

For `RedSalamander/ConnectionManagerDialog.cpp`, the hidden legacy owner-draw command buttons, form-action buttons, and toggle controls behind the visible DxUi shell are now retired. The refreshed audit no longer reports that file, and the focused `cmd_connection_manager_window_` family is green after the red/green rerun pair archived on 2026-04-22.

For `RedSalamander/ManagePluginsDialog.cpp`, the hidden legacy owner-draw command buttons and hidden owner-draw toggle control behind the visible DxUi shell are now retired. The refreshed audit no longer reports that file, and the full `cmd_plugin_configuration_dialog_` family is green after the archived rerun on 2026-04-22.

### Shared-font GDI measurement seams

The refreshed 2026-04-23 audit reports only the shared transitional helper allowlist in `Common/DxUi/DxUi.Typography.h`: two `CreateFontIndirectW` helper call sites and the fallback `DEFAULT_GUI_FONT` lookup used to derive a LOGFONT when Windows cannot provide a message font. No visible product file is currently listed by `Tools/Get-VisibleTypographyAudit.ps1`.

The broader manual regex audit archived with the final 2026-04-23 evidence still lists expected allowlist/exception paths such as `Common/DxUi/DxUi.Typography.h`, the then-active hidden text-input bridge diagnostic, Fluent icon font creation, and the NavigationView font-related perf scope. The strict menu/HFONT measurement search in the same archive is clean; the hidden text-input bridge diagnostic has since been retired by the native text-input closeout.

The refreshed closeout audit archived at `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_202647_segoe_variable_gdi_retirement_closeout/` confirms the same final state after the late command/spec refresh: the visible audit still reports only the shared helper allowlist in `Common/DxUi/DxUi.Typography.h`, and the strict product/test source search still returns no active `ThemedControls`, menu-HFONT, or GDI-measurement callers.

`RedSalamander/ThemedControls.h/.cpp` no longer exist. After the owner-draw button/toggle, modern combo, combo sizing, list-view theming, and list-header subclass APIs were retired, the files only held generic HWND helper code. That residual code first moved through the temporary `RedSalamander/Win32UiHelpers.h/.cpp` staging module, then the later remaining-Win32-UI closeout deleted that module. The surviving pure color/DPI helpers now live in `RedSalamander/UiMetrics.h/.cpp`; the remaining explicitly allowlisted HDC paint/interop seams are centralized behind `RedSalamander/D2DHdcPaint.h/.cpp` and documented in `Specs/UI/UI_VisibleNativeAudit.md`.

The focused source search archived on 2026-04-23 records zero references to `RedSalamander/ThemedControls.h`, `RedSalamander/ThemedControls.cpp`, `ThemedControls`, `EnableOwnerDrawButton`, `CreateModernComboBox`, `ApplyThemeToComboBox`, `ApplyThemeToListView`, `EnsureListViewHeaderThemed`, `DrawThemedPushButton`, `DrawThemedSwitchToggle`, and the retired modern-combo helpers across `RedSalamander`, `Plugins`, `Common`, and `Tests`.

`RedSalamander/FolderWindow.FileSystem.cpp` no longer carries the dead fenced change-case and selection-mask dialog implementations that previously kept extra owner-draw and combo/input-frame support code beside the live prompt-window path. The only live route for those commands is now the owned DxUi prompt windows already covered by `cmd_pane_changeCase_*` and `cmd_pane_selection_mask_dialogs`, and the focused source-audit refresh on 2026-04-22 records zero remaining references to the retired legacy dialog symbols.

`RedSalamander/NavigationViewInternal.h` is no longer in this bucket. The dead shared-font layout helper has been removed, the fresh visible-typography audit now reports only `Common/DxUi/DxUi.Typography.h` transitional helper hits, and the full `cmd_pane_navigationView_` family is green again after the focused selftests switched to the `NavigationViewDebugSnapshot` edit-host contract instead of enumerating descendant native `Edit` windows.

`RedSalamander/FunctionBar.cpp` is no longer in this bucket. Its visible key-glyph, label, and modifier text path now renders through a shared DirectWrite/D2D function-bar surface, the focused `shortcut_functionbar_dispatch_refresh`, `cmd_app_toggleUiChrome`, and `cmd_app_toggleUiChrome_keeps_navigation_shell_stable` command cases are green in fresh same-machine archives on 2026-04-22, and the focused audit archive on 2026-04-22 records zero plain `DrawTextW` hits remaining in that file.

`RedSalamander/FolderWindow.StatusBar.cpp` is no longer in this bucket. Its text paint and security-part sizing now stay on the shared `DxUi.Typography` / DirectWrite path, the focused `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu` and `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable` command cases are green in fresh same-machine archives on 2026-04-22, and the refreshed audit archive on 2026-04-22 no longer reports that file.

`Plugins/ViewerText/ViewerText.cpp` is no longer in this bucket. The root shell chrome now paints through DirectWrite/D2D even when the shared HWND render target is unavailable, the stale visible GDI/HFONT fallback and dead notify path are retired, the focused `TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo` runtime contract now asserts zero visible legacy GDI/HFONT text surfaces in the active shell, and the focused audit archive on 2026-04-22 records `0 legacy hits` remaining in that file.

### Remaining font-derivation holdouts

None in the refreshed visible audit. The only reported font creation entries are the documented transitional helper sites in `Common/DxUi/DxUi.Typography.h`.

## Native Text-Input Status

DxUi text input now runs through the host-HWND native retained session for `TextField` and editable `ComboBox`. The former hidden `DxUiTextInputBridgeWindow`, hidden edit/RichEdit fallback, bridge-only font diagnostic, and bridge LOGFONT contract are retired. NavigationView edit mode runs through a visible DxUi-backed text host instead of a native visible Win32 edit surface.

## Verification Snapshot

Fresh evidence for this audit refresh:

- `.\build.ps1 -Configuration Debug -ProjectName DxUiTests`
- `.\build.ps1 -Configuration Debug -ProjectName RedSalamander`
- `.\.build\x64\Debug\DxUiTests.exe --suite=Grid --perf-jsonl=...`
- `.\.build\x64\Debug\DxUiTests.exe --suite=Tree --perf-jsonl=...`
- `.\.build\x64\Debug\DxUiTests.exe --suite=Tooltip`
- archived focused row/tooltip runs: `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_162341_grid/`, `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_162341_tree/`, and `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_162341_tooltip/`
- archived final audit bundle: `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_171000_segoe_variable_gdi_retirement_final/`
- archived closeout audit bundle: `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_202647_segoe_variable_gdi_retirement_closeout/`
- archived focused command reruns: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_162951/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_163057/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_163705/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_163728/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_163738/`, and `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_170933/`
- archived late closeout command reruns: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_185757/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_193330/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_195938/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_195940/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_195954/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_200627/`, and `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_201922/`
- `.\build.ps1 -Configuration Debug -ProjectName ViewerImgRaw`
- `.\build.ps1 -Configuration Debug -ProjectName ViewerSpace`
- `.\build.ps1 -Configuration Debug -ProjectName ViewerWeb`
- `.\build.ps1 -Configuration Debug -ProjectName ViewerText`
- `.\build.ps1 -Configuration Debug -ProjectName ViewerPETests`
- `.\build.ps1 -Configuration Debug -ProjectName RedSalamander`
- `powershell -ExecutionPolicy Bypass -File .\Tools\Get-VisibleTypographyAudit.ps1`
- `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost`
- `.\.build\x64\Debug\DxUiTests.exe --suite=NativeTextInput`
- archived native text-input bridge-removal rerun: `Specs/TestRuns/local_scratch/dxui_native_textinput_after_production_bridge_removal_20260517_1958.jsonl`
- `.\.build\x64\Debug\ViewerPETests.exe`
- `.\.build\x64\Debug\ViewerPETests.exe TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo`
- `.\.build\x64\Debug\ViewerPETests.exe TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses`
- `.\.build\x64\Debug\ViewerPETests.exe TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo`
- `.\.build\x64\Debug\ViewerPETests.exe TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo`
- `.\.build\x64\Debug\ViewerPETests.exe TestViewerTextDiffModesAndPlaceholders`
- `.\.build\x64\Debug\ViewerSqliteTests.exe`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=viewer_text_diff_perf`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigationView_menu_region_keyboard_activation_opens_menu`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigationView_history_dropdown_escape_returns_focus_to_folder_view`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigationView_disk_info_region_keyboard_activation_opens_menu`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigationView_history_dropdown_keyboard_navigation`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigationView_path_doubleClick_enters_edit_mode`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigationView_full_path_popup_edit_route`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigationView_edit_suggest_keyboard_routing`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigationView_`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_compare_directories_options_`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_app_openDriveMenus_keeps_navigation_shell_stable`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_compare_directories_options_long_run_open_close_stays_stable`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_compare_directories_options_access_keys_focus_expected_controls`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_compare_directories_options_`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_connection_manager_window_`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_changeCase_`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_selection_mask_dialogs`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_plugins_`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_plugin_configuration_dialog_`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_connection_manager_window_`
- archived ViewerText shell-contract rerun: `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_204713_viewer_text_shell_contract/`
- archived ViewerText diff-modes rerun: `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_204720_viewer_text_diff_modes_green/`
- archived `viewer_text_diff_perf` baseline/candidate pair: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_203653/` and `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_204739/`
- archived ViewerText shell-audit refresh: `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_204845_viewer_text_shell_directwrite/`
- archived `cmd_pane_changeCase_` family rerun: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_215346/`
- archived exact `cmd_pane_selection_mask_dialogs` rerun with copied `last_run` trace and exit code `0`: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_215513_selection_mask_dead_dialog_retirement/`
- archived FolderWindow source cleanup audit: `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_215513_folderwindow_dead_dialog_source_cleanup/`
- archived focused Preferences plugin checkbox red/green pair around the DxUi Grid double-click toggle fix: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_143442/` and `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_143955/`
- archived focused Preferences shell/page/viewers/general/panes/plugin reruns: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144026/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144030/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144034/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144037/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144040/`, and `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144136/`
- archived focused plugin configuration and Connection Manager reruns: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144214/` and `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144824/`
- archived ThemedControls visible-surface source/audit refresh: `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_144844_themedcontrols_visible_surface_cleanup/`
- archived ThemedControls file-retirement command reruns: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152530/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152611/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152627/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152701/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_153313/`, and `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_153329/`
- archived ThemedControls file-retirement source/audit refresh: `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_153351_themedcontrols_file_retirement/`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_plugin_configuration_dialog_`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=shortcut_functionbar_dispatch_refresh`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_app_toggleUiChrome`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_app_toggleUiChrome_keeps_navigation_shell_stable`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`

Archived runs:

- `Specs/TestRuns/4cb089111a23/DxUi/2026-04-22_125500_typography_phase2/`
- `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_131500_typography_phase2b/`
- `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_145615_viewer_web_menu_model_green/`
- `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_145615_viewer_text_menu_model_green/`
- `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_151245_viewer_imgraw_menu_model_green/`
- `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_151252_viewer_space_menu_model_green/`
- `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_151515_viewerpe_full_suite/`
- `Specs/TestRuns/4cb089111a23/ViewerSqliteTests/2026-04-22_131600_typography_phase2b/`
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_151300_visible_typography_audit/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_140844/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_140852/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_140932/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_140936/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_140944/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_153114/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_153404/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_154109/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_161055/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_161617/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_163431/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_163502/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_184838/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_185143/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_191302/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201333/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201346/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201347/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201749/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201750/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201751/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_195812/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_195813/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_212811/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_212817/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_212835/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_212845/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_213733/`
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_195820_statusbar_typography_audit/`
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_201900_functionbar_textpath/`
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_212930_navigationviewinternal_contract/`
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_2138_themedcontrols_compare_options_helper_cleanup/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_143442/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_143955/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144026/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144030/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144034/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144037/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144040/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144136/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144214/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144824/`
- `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_144844_themedcontrols_visible_surface_cleanup/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152530/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152611/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152627/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152701/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_153313/`
- `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_153329/`
- `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_153351_themedcontrols_file_retirement/`

Focused NavigationView command selftests are green again after the stale descendant-`Edit` assumptions were retired. The refreshed `cmd_pane_navigationView_path_doubleClick_enters_edit_mode`, `cmd_pane_navigationView_full_path_popup_edit_route`, `cmd_pane_navigationView_edit_suggest_keyboard_routing`, and full `cmd_pane_navigationView_` family archives from 2026-04-22 now validate the DxUi host/bridge contract through `NavigationViewDebugSnapshot` rather than walking for a visible native edit child.

Focused NavigationView command selftests stay green after the main-menu and navigation owner-draw cleanup, and the refreshed visible-typography audit no longer reports `Plugins/FileSystem/FileSystem.Menu.cpp`, `Plugins/FileSystemS3/FileSystemS3.Menu.cpp`, `RedSalamander/ConnectionManagerDialog.cpp`, `RedSalamander/ManagePluginsDialog.cpp`, `RedSalamander/FunctionBar.cpp`, `RedSalamander/FolderWindow.StatusBar.cpp`, `RedSalamander/ThemedControls.cpp`, or `Plugins/ViewerText/ViewerText.cpp`. `ViewerPETests` now assert that `ViewerWeb`, `ViewerText`, `ViewerImgRaw`, and `ViewerSpace` keep a plain hidden `HMENU` model with zero owner-draw items and that the active ViewerText shell exposes zero visible legacy GDI/HFONT text surfaces, the native DxUi text-input replacement is green on focused `NativeTextInput` and adjacent suites after retiring the hidden bridge, the full Compare Directories options family is green after the hidden owner-draw retirement plus the deterministic scroll-stability resize fix and still green after retiring the stale `ThemedControls::CenterEditTextVertically()` Compare Options hook, the focused `cmd_preferences_dialog_plugins_`, `cmd_connection_manager_window_`, and `cmd_plugin_configuration_dialog_` families are green after retiring the remaining visible ThemedControls caller set, the late 2026-04-23 reruns keep the file-system plugin configuration, connection credentials, drive-menu shell stability, pane status-bar sort, change-case prompts, and Shortcuts shell/live-search coverage green, the Function Bar slice is green on focused command coverage with a fresh same-machine paint metric drop from `1953.2 us` average to `250.0 us` average in `cmd_app_toggleUiChrome_keeps_navigation_shell_stable`, and the pane status-bar slice is green with a fresh audit proving that file no longer sits in the visible GDI seam bucket. The fresh ViewerText `viewer_text_diff_perf` rerun adds `viewer.chrome.paint_us`; open-to-first-visible stays effectively flat or better across the archived baseline/candidate pair, while the large side-by-side theme-switch and scroll repaint timings are higher in the same-machine candidate (`14708 -> 73924 us` and `5839 -> 8918 us`), so no perf win is claimed for that slice. The 2026-04-23 strict menu-HFONT audits are clean, and the Segoe UI Variable / visible GDI retirement plan is closed. The former red broad `cmd_app_shortcuts_*` grouped-collapse / reorder / persisted-layout / search-state cases were tracked separately and closed in `Specs/Plans/Done/DxUI_MigrationRoadmap.md`; they do not represent a visible typography, menu-HFONT, or `ThemedControls` regression.

The final `ThemedControls.h/.cpp` retirement is also archived: the old files are deleted, the temporary `Win32Ui::...` call sites from `RedSalamander/Win32UiHelpers.h/.cpp` were subsequently removed, and the focused source search has zero `ThemedControls`, `Win32Ui::`, `using Win32Ui`, or `Win32UiHelpers` references across product and test source. This is a module-retirement and regression-proofing slice, so no performance win is claimed.

The FolderWindow filesystem prompt slice is green after dead-code retirement as well. `FolderWindow.FileSystem.cpp` no longer carries the dead fenced change-case and selection-mask dialog implementations or their unused hook helpers, the focused `cmd_pane_changeCase_` family rerun is archived `4/4` green, the exact `cmd_pane_selection_mask_dialogs` rerun is archived with exit code `0` plus copied `last_run` trace, and the focused source-audit archive records zero remaining references to the retired legacy dialog symbols. No perf win is claimed for this slice because it removes disabled dead code rather than changing an active hot path.
