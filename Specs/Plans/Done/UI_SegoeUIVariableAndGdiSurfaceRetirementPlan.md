# App Typography Unification and Remaining GDI Surface Retirement Plan

Last updated: 2026-04-26

Status: Done

Authoritative specs:

- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/UI/UI_FolderView.md`
- `Specs/UI/UI_NavigationView.md`
- `Specs/UI/UI_VisibleNativeAudit.md`
- `Specs/UI/UI_VisibleComctlAudit.md`

Related plans:

- `Specs/Plans/Done/UI_DxUiWinUIDesignAlignmentPlan.md`
- `Specs/Plans/Done/UI_PreferencesWin32Removal.md`
- `Specs/Plans/Done/DxUI_MigrationRoadmap.md`
- `Specs/Plans/Done/UI_DxUiRemainingMigrationCloseoutPlan.md`

## 0. Top Checklist

Use this section as the execution control list for the whole migration. Mark completed work with `[x]`, keep pending work as `[ ]`, and add newly discovered required items under the relevant bucket before implementation continues.

### Shared Typography Foundation

- [x] Add one shared typography module for app-owned UI text and fonts, with one source of truth for family, weight, size, and fallback policy.
- [x] Move `Common/DxUi/DxUi.WindowHost.cpp` `FontSpec` / `GetFontSpec(...)` onto that shared module.
- [x] Move `Common/DxUi/DxUi.TextInput.cpp` hidden-bridge font creation onto the same shared module for the transition period.
- [x] Add companion HFONT / LOGFONT creation helpers for temporary HWND-backed surfaces so no visible app code calls `DEFAULT_GUI_FONT`, `CreateFontW`, or `CreateFontIndirectW` directly.
- [x] Update `Specs/UI/UI_DxUiWinUIDesign.md` to encode the Windows 11 Segoe UI Variable contract plus the icon, emoji, and monospace exceptions.
- [x] Add a visible-typography / visible-GDI audit script under `Tools/` and keep a checked-in audit summary under `Specs/UI/`.

### Application DirectWrite Surfaces

- [x] Convert `RedSalamander/FolderView.Rendering.cpp` to the shared typography helper.
- [x] Convert `RedSalamander/NavigationView.Rendering.cpp` to the shared typography helper.
- [x] Convert `RedSalamander/FolderWindow.FileOperations.Popup.cpp` to the shared typography helper.
- [x] Convert `RedSalamander/Ui/AlertOverlay.h` and any sibling alert-overlay code paths to the shared typography helper.
- [x] Re-tune row heights, truncation, watermarks, overlays, and hit-testing where Segoe UI Variable metrics differ from Segoe UI.
  - [x] Clamp DxUi Grid/Tree interactive text rows to the shared 20 DIP Segoe UI Variable Body line-height minimum and cover the row/hit-test cadence in focused `DxUiTests` suites.
  - [x] Update editable ComboBox popup click selftests to derive item hit points from debug popup geometry instead of stale fixed row coordinates after shared popup-row metric changes.

### HWND / HFONT Application Surfaces

- [x] Replace non-menu `DEFAULT_GUI_FONT` / `CreateFontW` / `CreateFontIndirectW` usage in `RedSalamander/FunctionBar.cpp`.
- [x] Replace non-menu `DEFAULT_GUI_FONT` usage in `RedSalamander/FolderWindow.StatusBar.cpp`.
- [x] Replace the explicit GDI `Segoe UI` path font in `RedSalamander/NavigationView.cpp`.
- [x] Replace shared dialog-font derivation in `RedSalamander/Preferences.Dialog.cpp`, `RedSalamander/ConnectionManagerDialog.cpp`, `RedSalamander/ManagePluginsDialog.cpp`, `RedSalamander/CompareDirectoriesWindow.cpp`, and `RedSalamander/FolderWindow.FileSystem.cpp`.
- [x] Retire the stale `RedSalamander/ThemedControls.cpp` `CenterEditTextVertically()` helper and the inactive Compare Directories legacy-edit hook that used it.
- [x] Delete the dead `#if 0` legacy change-case and selection-mask dialog implementations from `RedSalamander/FolderWindow.FileSystem.cpp` now that the DxUi prompt-window route owns those commands.
- [x] Retire the remaining visible `ThemedControls.cpp` owner-draw button/toggle and combo path from `RedSalamander/Preferences.Dialog.cpp`.
- [x] Retire the remaining visible `ThemedControls.cpp` toggle/combo seam from `RedSalamander/ManagePluginsDialog.cpp`.
- [x] Retire the remaining visible `ThemedControls.cpp` combo/list/header seam from `RedSalamander/ConnectionManagerDialog.cpp`.
- [x] Reduce `RedSalamander/ThemedControls.cpp` to either shared non-visible compatibility seams or fully retired dead code.
- [x] Delete `RedSalamander/ThemedControls.h/.cpp`; the temporary `RedSalamander/Win32UiHelpers.h/.cpp` staging module used during this phase was later retired by `Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md`.

### Menu / Popup / Owner-Draw Retirement

- [x] Retire GDI owner-draw main-menu code in `RedSalamander/RedSalamander.cpp`.
- [x] Retire GDI owner-draw navigation popup code in `RedSalamander/NavigationView.Menus.cpp`.
- [x] Retire any remaining compare-directories command/menu GDI surfaces.
- [x] Retire plugin/file-system owner-draw menu paths in `Plugins/FileSystem/FileSystem.Menu.cpp` and `Plugins/FileSystemS3/FileSystemS3.Menu.cpp`.
- [x] Retire dead viewer owner-draw menu paths in `Plugins/ViewerText/ViewerText.MenuTheme.cpp` and `Plugins/ViewerWeb/ViewerWeb.cpp`.
- [x] Retire the remaining active viewer owner-draw menu paths in `Plugins/ViewerSpace/ViewerSpace.cpp` and `Plugins/ViewerImgRaw/ViewerImgRaw.cpp`.
- [x] Remove `NONCLIENTMETRICS.lfMenuFont`, `GetTextExtentPoint32W`, and menu-HFONT measurement from active menu/popup rendering once those paths are migrated.

### Hidden Bridge and Navigation Closure

- [x] Replace the hidden text-input bridge in `Common/DxUi/DxUi.TextInput.cpp` with a non-visible service-window path that does not keep a hidden edit-control dependency in the product.
- [x] Replace the NavigationView visible native edit path and popup text-entry surfaces with a DxUi-backed text host or a non-visible helper only.
- [x] Keep IME, clipboard, undo/redo, accessibility, and keyboard-routing parity through the service-window replacement.

### Plugin / Viewer Surface Migration

- [x] Migrate the remaining visible GDI/HFONT paths in `Plugins/ViewerText/ViewerText.cpp`, `ViewerText.Text.cpp`, and `ViewerText.Hex.cpp`.
- [x] Migrate the remaining visible GDI/HFONT paths in `Plugins/ViewerWeb/ViewerWeb.cpp`.
- [x] Migrate the remaining visible GDI/HFONT paths in `Plugins/ViewerSpace/ViewerSpace.cpp`.
- [x] Migrate the remaining visible GDI/HFONT paths in `Plugins/ViewerImgRaw/ViewerImgRaw.cpp`.
- [x] Migrate the remaining visible GDI/HFONT paths in `Plugins/ViewerPE/ViewerPE.cpp`.
- [x] Migrate the remaining visible GDI/HFONT paths in `Plugins/ViewerVLC/ViewerVLC.cpp`.
- [x] Verify `Plugins/ViewerSqlite/ViewerSqlite.cpp` and the remaining plugin windows against the final typography and visible-GDI audit even where no active drift was found in the initial inventory.

### Validation and Closeout

- [x] Add or refresh `Tests/DxUiTests` coverage for text roles, rendering, menus, multiline input, and window-host behavior.
- [x] Add or refresh `Tests/ViewerPETests/ViewerPETests.cpp` and any required viewer/plugin harness coverage for migrated plugin windows.
- [x] Refresh the NavigationView commands selftests to validate the `NavigationViewDebugSnapshot` edit-host contract instead of enumerating descendant native `Edit` windows.
- [x] Refresh focused Preferences, plugin configuration, and Connection Manager command coverage for the retired `ThemedControls.cpp` caller cluster, including the DxUi Grid checkbox double-click toggle regression found during validation.
- [x] Refresh the DxUi Menu acrylic visual baseline so the golden image uses a deterministic owner-backed backdrop instead of environment-dependent desktop pixels.
- [x] Refresh DxUi Rendering visual baselines after the row/font metric updates, with core-control baselines forced to reduced motion so hover/focus animation progress cannot vary by suite order.
- [x] Run the full `DxUiTests` suite after focused metric/baseline fixes and archive perf JSONL evidence.
- [x] Refresh `RedSalamander/SelfTest/Commands/Commands.SelfTest*.cpp` coverage for navigation, dialogs, menus, preferences, plugins, and view commands that exercise migrated surfaces.
  - [x] Refresh the focused command families touched in the 2026-04-23 row/helper slice: Preferences plugins, plugin configuration, Connection Manager, Compare Directories options, NavigationView, and Shortcuts row-tooltip tracking.
  - [x] Fix and rerun the isolated file-system plugin configuration roundtrip dirty-state guard after the broad commands attempt exposed persisted local test configuration drift; archived green at `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_185757/`.
  - [x] Refresh the remaining migration-relevant 2026-04-23 command slices: credential prompt family green at `.../2026-04-23_193330/`, drive-menu shell stability at `.../2026-04-23_195938/`, pane status-bar sort at `.../2026-04-23_195940/`, change-case prompts at `.../2026-04-23_195954/`, and the Shortcuts shell/live-search row-tooltip/open-close contract at `.../2026-04-23_200627/` plus `.../2026-04-23_201922/`.
  - [x] Split the remaining broad `cmd_app_shortcuts_*` grouped-collapse / reorder / persisted-layout / search choreography failures into the ongoing DxUi roadmap after the focused rerun at `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_202549/` showed the red cluster is separate shortcuts-state work rather than a visible typography or GDI seam regression.
- [x] Archive perf and validation evidence under `Specs/TestRuns/` for each migrated cluster.
- [x] Update authoritative specs and move this plan to `Specs/Plans/Done/` only after the audits, tests, and perf evidence are complete.

## 0.1 Closeout Snapshot

As of 2026-04-23, the migration is complete for the shared typography phase, the visible font-unification work, and the full `ThemedControls.h/.cpp` retirement:

- `Common/DxUi/DxUi.Typography.h` is the shared Segoe UI Variable contract for both DirectWrite and transitional `HFONT` creation.
- DxUi text roles, FolderView, NavigationView rendering, File Operations popup, AlertOverlay, FunctionBar, FolderWindow status/dialog surfaces, Preferences, Connection Manager, Manage Plugins, Compare Directories, and the migrated viewer windows now route font selection through that helper.
- the hidden main-menu owner-draw path is retired, and `ViewerPETests` now assert that `ViewerWeb`, `ViewerText`, `ViewerImgRaw`, and `ViewerSpace` keep a plain hidden `HMENU` model with zero owner-draw items.
- `RedSalamander/ManagePluginsDialog.cpp` no longer keeps hidden legacy owner-draw command buttons or the hidden legacy owner-draw toggle path behind the visible DxUi shell, and the full `cmd_plugin_configuration_dialog_` family is archived green on 2026-04-22.
- `RedSalamander/Preferences.Dialog.cpp`, `RedSalamander/ManagePluginsDialog.cpp`, and `RedSalamander/ConnectionManagerDialog.cpp` no longer consume the visible `ThemedControls` owner-draw button/toggle, modern combo, list-view, or list-header APIs.
- `RedSalamander/ThemedControls.h/.cpp` are deleted. The residual generic HWND helpers first moved through the temporary `RedSalamander/Win32UiHelpers.h/.cpp` staging module; that module is now deleted, surviving pure helpers live in `RedSalamander/UiMetrics.h/.cpp`, and the focused source search has zero `ThemedControls` hits across `RedSalamander`, `Common`, `Plugins`, and `Tests`.
- Validation found and fixed a shared DxUi Grid pointer edge case: checkbox double-clicks now toggle like two checkbox clicks instead of being swallowed by row double-click activation.
- DxUi Grid and Tree now clamp interactive text rows to the shared 20 DIP Segoe UI Variable Body line-height minimum, with focused Grid/Tree tests covering row cadence and hit-test geometry. The tooltip layer also clamps wrapped tooltip outer width to the viewport, and the Shortcuts row-tooltip command selftest now uses deterministic layout plus DPI-scaled hover points.
- Editable ComboBox popup-hit selftests now derive click points from the control's debug popup item geometry, so TextField coverage tracks the shared popup row metrics instead of a hardcoded pre-retune coordinate.
- `Common/DxUi/DxUi.Accessibility.cpp` and `RedSalamander/ShortcutsWindow.cpp` now support explicit accessible names for unlabeled visible DxUi text inputs, so the Shortcuts search field keeps a non-empty UI Automation name/value contract without relying on placeholder text.
- The strict active-menu source search is now clean: no product/test hits remain for `NONCLIENTMETRICS`, `lfMenuFont`, `GetTextExtentPoint32W`, `GetTextMetricsW`, or `CreateMenuFontForDpi` under `RedSalamander`, `Plugins`, `Common`, and `Tests`.
- The DxUi Menu acrylic visual baseline no longer depends on live desktop pixels; the test injects a deterministic owner-backed backdrop after confirming the popup enabled app-rendered acrylic blur, and the focused Menu suite is archived green with perf JSONL.
- DxUi Rendering visual baselines are refreshed for the Segoe UI Variable metric output; core-control baselines now opt into reduced motion to make hover/focus visual states deterministic in both focused and full-suite runs.
- The full `DxUiTests` suite is archived green with perf JSONL after the Menu, TextField, and Rendering metric/baseline fixes.
- Fresh archived verification exists under:
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
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_184838/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_185143/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_191302/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201333/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201346/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201347/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201749/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201750/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_201751/`
  - `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_201900_functionbar_textpath/`
  - `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_204713_viewer_text_shell_contract/`
  - `Specs/TestRuns/4cb089111a23/ViewerPETests/2026-04-22_204720_viewer_text_diff_modes_green/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_203653/`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_204739/`
  - `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_204845_viewer_text_shell_directwrite/`
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
  - `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_154258_menu_hfont_measurement_closure/`
  - `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_162341_grid/`
  - `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_162341_tree/`
  - `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_162341_tooltip/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_162951/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_163057/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_163705/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_163728/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_163738/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_170933/`
  - `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_171000_segoe_variable_gdi_retirement_final/`
  - `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_173710_menu/`
  - `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_174000_textfield/`
  - `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_174500_rendering/`
  - `Specs/TestRuns/7d3a1247382a/DxUiTests/2026-04-23_174600_full/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_185757/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_193330/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_195938/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_195940/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_195954/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_200627/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_201922/`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_202549/`
  - `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_202647_segoe_variable_gdi_retirement_closeout/`

This plan is no longer blocked on visible Segoe UI Variable migration work. The former red broad `cmd_app_shortcuts_*` cases exercise grouped collapse, reorder, persisted-layout, and search/sort choreography in `ShortcutsWindow`; they were tracked as follow-on DxUi behavior work in `Specs/Plans/Done/DxUI_MigrationRoadmap.md` and do not reopen any visible typography, menu-HFONT, or `ThemedControls` seam.

### 0.2 Follow-on Boundary

Do not reopen this plan for the former broad ShortcutsWindow stateful-suite failures. The migration-specific closeout contract is:

1. visible typography / menu-HFONT seams are retired,
2. app-owned surfaces use the shared Segoe UI Variable contract or documented helper allowlist,
3. targeted migrated-surface command coverage is green and archived,
4. the former shortcuts red cluster is separate DxUi state-management debt and is closed in `Specs/Plans/Done/DxUI_MigrationRoadmap.md`.

Latest follow-on anchors:

- broad shortcuts family red archive: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_201922/`
- focused long-run scrolling follow-on rerun after the selftest baseline fix: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_202549/`
- authoritative follow-on tracker: `Specs/Plans/Done/DxUI_MigrationRoadmap.md`

### 0.3 Closeout Evidence

Use these anchors when auditing this completed plan later:

1. `.\build.ps1 -Configuration Debug -ProjectName RedSalamander`
2. `powershell -ExecutionPolicy Bypass -File .\Tools\Get-VisibleTypographyAudit.ps1`
3. the strict source search archived in `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_202647_segoe_variable_gdi_retirement_closeout/strict_source_search.txt`
4. the refreshed shared-helper allowlist output archived in `.../visible_typography_audit.txt`
5. the targeted green command archives listed above for plugin configuration, connection credentials, drive menus, status bar, change-case prompts, and Shortcuts shell/live-search coverage

Since the last plan refresh, `RedSalamander/CompareDirectoriesWindow.Options.cpp` has also dropped the hidden owner-draw footer-button and toggle fallback path behind the visible DxUi options surface. The targeted command selftests for that slice now have archived red/green evidence:

- red proof of the old hidden owner-draw footer path: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_153114/`
- green baseline after retirement: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_153404/`
- green access-key follow-through after explicit DxUi footer mnemonics: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_154109/`
- red family rerun that isolated the remaining scroll-stability seam: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_161055/`
- focused repro of that seam before the geometry-stabilizing selftest fix: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_161617/`
- green focused scroll-stability repro after pinning the window-specific snapshot path and forcing a deterministic smaller viewport: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_163431/`
- green full Compare Directories options family rerun: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_163502/`
- green focused hidden-bridge replacement rerun: `Specs/TestRuns/4cb089111a23/DxUi/2026-04-22_182650_textinputbridge_custom_service_green/`
- red proof that the Connection Manager hidden toggle fallback still needed explicit hidden-host state sync after owner-draw retirement: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_184838/`
- green full `cmd_connection_manager_window_` family rerun after retiring the hidden owner-draw command/toggle/form-action path and restoring hidden-host toggle sync: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_185143/`
- green full `cmd_plugin_configuration_dialog_` family rerun after retiring the hidden owner-draw command-button and toggle fallback behind the visible DxUi shell: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_191302/`
- green focused pane status-bar sort-popup regression after aligning the command selftest with the shared DxUi popup contract: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_195523/`
- green focused pane status-bar reruns after removing the remaining HFONT-backed sizing fallback in `RedSalamander/FolderWindow.StatusBar.cpp`: `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_195812/` and `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_195813/`
- refreshed visible-typography audit proving the status-bar file no longer appears in the visible seam inventory: `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_195820_statusbar_typography_audit/`

Focused NavigationView command selftests are green after the dead owner-draw popup cleanup, the fresh visible-typography audit no longer reports `Plugins/FileSystem/FileSystem.Menu.cpp`, `Plugins/FileSystemS3/FileSystemS3.Menu.cpp`, `RedSalamander/ConnectionManagerDialog.cpp`, `RedSalamander/ManagePluginsDialog.cpp`, `RedSalamander/FunctionBar.cpp`, `RedSalamander/FolderWindow.StatusBar.cpp`, or `Plugins/ViewerText/ViewerText.cpp`, focused `ViewerPETests` are green after the `ViewerWeb` / `ViewerText` dead menu-theme retirement plus the `ViewerImgRaw` / `ViewerSpace` owner-draw menu cleanup and now also assert that the active ViewerText shell exposes zero visible legacy GDI/HFONT text surfaces, the focused `DxUiTests --suite=TextInputBridge` run is green after the custom service-window bridge replacement, the full Compare Directories options family is green after the hidden owner-draw retirement plus the deterministic scroll-stability resize fix, the full focused `cmd_connection_manager_window_` family is green after retiring the hidden owner-draw command/toggle/form-action path behind the visible DxUi shell, the full `cmd_plugin_configuration_dialog_` family is green after retiring the hidden owner-draw command-button and toggle fallback behind the visible DxUi shell, the Function Bar slice is now green on focused command coverage with the visible text path moved onto a DirectWrite/D2D renderer, and the pane status-bar slice is green with the remaining HFONT-backed sizing fallback removed. Same-machine `render.function_bar.paint_us` dropped from `1953.2 us` average in `2026-04-22_201347` to `250.0 us` average in `2026-04-22_201751` for `cmd_app_toggleUiChrome_keeps_navigation_shell_stable`, while `render.status_bar.paint_us` is slightly lower in the focused sort-popup scenario (`3038.7 us` avg in `2026-04-22_195523` vs `2875.3 us` avg in `2026-04-22_195812`). The ViewerText root-shell cleanup adds `viewer.chrome.paint_us` to `viewer_text_diff_perf`; open-to-first-visible stays effectively flat or better across the archived baseline/candidate pair (`435051 -> 434129 us`, `171691 -> 161185 us`, `204673 -> 185379 us`), while the large side-by-side theme-switch and scroll repaint timings are higher in the same-machine candidate (`14708 -> 73924 us` and `5839 -> 8918 us`), so no perf win is claimed for that slice. The status-bar shell-stability scenario remains noisy across tiny three-sample reruns (`2958.0` to `4609.0 us` avg across `2026-04-22_195525`, `..._195813`, `..._195904`, `..._195905`), so that part stays directional rather than a claimed perf win. The plan is now closed; the remaining broad `cmd_app_shortcuts_*` grouped-collapse / reorder / persisted-layout / search-state failures are follow-on DxUi behavior work and do not reopen the visible typography / GDI retirement scope.

The NavigationView validation slice is now refreshed as well. `RedSalamander/NavigationViewInternal.h` no longer carries the dead shared-font helper seam, the fresh audit now reports only `Common/DxUi/DxUi.Typography.h` transitional helper hits, and the focused `cmd_pane_navigationView_path_doubleClick_enters_edit_mode`, `cmd_pane_navigationView_full_path_popup_edit_route`, `cmd_pane_navigationView_edit_suggest_keyboard_routing`, and full `cmd_pane_navigationView_` family reruns archived on 2026-04-22 are green after the selftests switched to `NavigationViewDebugSnapshot` instead of enumerating descendant native `Edit` windows.

Archived evidence for that refresh:

- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_212811/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_212817/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_212835/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_212845/`
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_212930_navigationviewinternal_contract/`

The first `ThemedControls.cpp` cleanup slice is now in as well. The dead `CenterEditTextVertically()` helper is removed, the inactive Compare Directories legacy edit branch no longer references it, the full `cmd_compare_directories_options_` family rerun archived on 2026-04-22 stays green, and a focused audit archive records `0` remaining references to that helper. This is a correctness-only cleanup on an inactive fallback path, so no new perf win is claimed for this slice.

Archived evidence for that refresh:

- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_213733/`
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_2138_themedcontrols_compare_options_helper_cleanup/`

The FolderWindow file-system cleanup slice is in too. `RedSalamander/FolderWindow.FileSystem.cpp` no longer carries the dead fenced change-case and selection-mask dialog implementations, their unused window-proc hook helpers, or the temporary include-only typography/input-frame support that existed only for those disabled paths. Focused verification stayed green: the archived `cmd_pane_changeCase_` family rerun is `4/4` passed, an exact `cmd_pane_selection_mask_dialogs` rerun was archived with exit code `0` plus copied `last_run` trace, and a matching source audit records `0` remaining references to the retired legacy dialog symbols. No perf claim is made for this slice because it removes dead disabled code rather than changing an active hot path.

Archived evidence for that refresh:

- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_215346/`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-22_215513_selection_mask_dead_dialog_retirement/`
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-22_215513_folderwindow_dead_dialog_source_cleanup/`

The active `ThemedControls.cpp` caller cleanup slice is now in. Preferences no longer routes visible footer/page buttons, switches, or combo theme/dropdown work through `ThemedControls`; Manage Plugins no longer uses the legacy toggle/combo seam; Connection Manager uses native combo/list theming and no custom list-header subclass painting. `ThemedControls.cpp` is reduced to non-visible compatibility helpers, and the retired visible API set has zero source references across `RedSalamander`, `Plugins`, `Common`, and `Tests`.

Focused verification stayed green after a shared DxUi Grid fix found during the Preferences plugin checkbox test. The red exact checkbox run showed a rapid second checkbox click being treated as a double-click no-op; `Common/DxUi/DxUi.Grid.cpp` now toggles checkbox cells on the double-click down as a second click, and the exact case plus focused Preferences plugin family are green. The full `cmd_plugin_configuration_dialog_` and `cmd_connection_manager_window_` families are green as well. The archived perf metrics are correctness/steady-state evidence for the dialog-shell cleanup; no performance win is claimed for this slice.

A broad `cmd_preferences_dialog_` run is also archived red at `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_143409/` with multiple state-order and settling failures outside the retired `ThemedControls` API seam. That is why the broader commands-suite refresh and final closeout checklist items remain unchecked.

Archived evidence for that refresh:

- initial exact shell guard rerun: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_142709/`
- broad red Preferences family rerun: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_143409/`
- red exact Preferences checkbox proof: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_143442/`
- green exact Preferences checkbox rerun: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_143955/`
- focused Preferences shell/page/viewers/general/panes/plugin reruns: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144026/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144030/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144034/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144037/`, `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144040/`, and `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144136/`
- focused plugin configuration and Connection Manager reruns: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144214/` and `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_144824/`
- refreshed audit/source-search proof: `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_144844_themedcontrols_visible_surface_cleanup/`

The final `ThemedControls.h/.cpp` file-retirement slice is now in. The files had no remaining reason to exist as "themed controls"; their only remaining content was generic HWND helper code. That code first moved through `Win32UiHelpers.h/.cpp`, then the later remaining-Win32-UI closeout retired that staging module as well. Surviving pure helpers now live in `UiMetrics.h/.cpp`, the remaining explicitly allowlisted HDC paint/interop seams are centralized through `D2DHdcPaint.h/.cpp`, and current source searches remain empty for `ThemedControls`, `Win32Ui::`, `using Win32Ui`, and `Win32UiHelpers` across product and test source.

Focused validation for that refresh:

- `.\build.ps1 -Configuration Debug -ProjectName RedSalamander` passed.
- `cmd_preferences_dialog_plugins_` passed `29/29`: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152530/`
- `cmd_plugin_configuration_dialog_` first exposed one order-sensitive access-key failure (`9/10`) at `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152611/`; the exact access-key case passed at `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152627/`, and the full family passed `10/10` on rerun at `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_152701/`.
- `cmd_connection_manager_window_` passed `13/13`: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_153313/`
- `cmd_compare_directories_options_` passed `9/9`: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-23_153329/`
- refreshed audit/source-search proof: `Specs/TestRuns/7d3a1247382a/Audit/2026-04-23_153351_themedcontrols_file_retirement/`

No performance win is claimed for this slice; the archived perf metrics are correctness and regression evidence for a module move with no intended runtime behavior change.

## Scope

This plan covers two coupled application-wide goals:

1. move all visible application typography to one shared contract centered on Segoe UI Variable,
2. remove the remaining visible GDI-rendered UI surfaces so typography, spacing, DPI behavior, and theming are all driven by the same DirectWrite/D2D-or-DxUi stack.

It also explicitly includes retirement of the hidden DxUi text-input bridge from the active product path. That bridge is in scope for closure, not a permanent exception.

This is intentionally broader than a simple `FontSpec` change in `DxUi.WindowHost.cpp`. The current typography drift comes from multiple parallel rendering systems, not from one wrong family name.

## Current Code Truth

As of 2026-04-23, the repo no longer has four equally active typography systems. It now has one shared Segoe UI Variable contract plus a shrinking set of legacy exceptions:

1. **Shared typography contract**
   - `Common/DxUi/DxUi.Typography.h` owns the Windows 11 family constants, DirectWrite format creation, and transitional `HFONT` / `LOGFONT` creation.
   - `Common/DxUi/DxUi.WindowHost.cpp` and the hidden text-input bridge both consume that helper.

2. **Visible app-owned DirectWrite and hybrid surfaces already on the shared contract**
   - `RedSalamander/FolderView.Rendering.cpp`
   - `RedSalamander/NavigationView.Rendering.cpp`
   - `RedSalamander/NavigationView.Edit.cpp`
   - `RedSalamander/NavigationView.FullPathPopup.cpp`
   - `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
   - `RedSalamander/Ui/AlertOverlay.h`
   - the migrated viewer surfaces in `Plugins/ViewerPE/`, `ViewerVLC/`, `ViewerWeb/`, `ViewerSpace/`, `ViewerImgRaw/`, and `ViewerText/`

3. **Visible HWND/HFONT surfaces already on the shared contract**
   - `RedSalamander/NavigationView.cpp`
   - `RedSalamander/FunctionBar.cpp`
   - `RedSalamander/FolderWindow.cpp`
   - `RedSalamander/FolderWindow.StatusBar.cpp`
   - `RedSalamander/Preferences.Dialog.cpp`
   - `RedSalamander/ConnectionManagerDialog.cpp`
   - `RedSalamander/ManagePluginsDialog.cpp`
   - `RedSalamander/CompareDirectoriesWindow.cpp`
   - `RedSalamander/CompareDirectoriesWindow.Options.cpp`
   - `RedSalamander/FolderWindow.FileSystem.cpp`

4. **Remaining visible drift buckets**
   - the fresh audit no longer reports active compare-directories or file-system menu-source drift; those surfaces now flow through the DxUi menu paths and focused command selftests,
   - `RedSalamander/FolderWindow.FileSystem.cpp` now carries only the live DxUi prompt routes for change-case and selection-mask commands; the dead fenced legacy dialog code behind those commands has been removed,
   - `RedSalamander/ThemedControls.h/.cpp` are deleted; the later remaining-Win32-UI closeout also deleted `RedSalamander/Win32UiHelpers.h/.cpp`, with surviving pure helpers in `RedSalamander/UiMetrics.h/.cpp`,
   - the former non-green broad shortcuts grouped-collapse / reorder / persisted-layout / search-state suite is tracked separately in `Specs/Plans/Done/DxUI_MigrationRoadmap.md`; it is no longer treated as visible typography or menu-HFONT drift.

The result is that font-family drift has been collapsed and closed out for this plan. Later work continues in the DxUi roadmap for broader ShortcutsWindow state choreography, not for visible typography/GDI retirement.

## Target End State

The target end state is:

- every visible application-owned text surface uses one shared typography contract,
- the primary UI family is Segoe UI Variable,
- icon glyphs use `Segoe Fluent Icons`,
- emoji uses `Segoe UI Emoji`,
- monospace uses the repo-approved monospace fallback chain,
- no visible application-owned surface uses `DEFAULT_GUI_FONT`,
- no visible application-owned surface derives text from `NONCLIENTMETRICS.lfMenuFont`,
- no visible application-owned surface hardcodes `CreateTextFormat(..., L"Segoe UI", ...)` or `CreateFontW(..., L"Segoe UI")` outside the shared compatibility helper,
- no visible application-owned surface relies on GDI text measurement or GDI owner-draw painting for routine UI.

The target is explicitly **not**:

- removing all Win32 HWND hosting from the process,
- removing hidden compatibility helpers that are not visible surfaces in the same milestone,
- removing shell-icon conversion paths that still legitimately require GDI/WIC bridging,
- or rewriting OS-native common dialogs that are not application-owned chrome.

## Keep List Until Separate Closure

The following are allowed to remain temporarily even after the visible-surface migration is complete, but must be documented as deliberate exceptions:

- shell icon extraction / premultiplication bridges used to convert shell imagery into D2D-ready bitmaps,
- non-app-owned system dialogs and shell UI.

These exceptions are not visible typography surfaces and should not block closure of the visible GDI-retirement milestone.

## Phase 0. Contract Lock and Inventory Refresh

Current risk:

- The spec says `Segoe UI Variable`, but the implementation currently uses `Segoe UI`.
- The shared Windows 11 Segoe UI Variable contract still needs to be encoded in one place instead of being split across local helpers and hardcoded strings.
- The remaining visible GDI surface inventory is spread across menus, bars, dialog chrome, and older hybrid windows.

Required work:

1. Lock the exact Windows 11 Segoe UI Variable runtime family contract used by the app:
   - define the shared family-name constants,
   - define any role-specific family usage if needed,
   - and remove downlevel fallback planning from this effort.
2. Update `Specs/UI/UI_DxUiWinUIDesign.md` so the typography section names:
   - the exact preferred family or families,
   - the Windows 11 baseline assumption for this effort,
   - permitted non-text exceptions such as icon, emoji, and monospace.
3. Create a repo audit helper for visible typography drift, covering:
   - `DEFAULT_GUI_FONT`
   - `lfMenuFont`
   - visible `CreateFontW/CreateFontIndirectW`
   - visible `CreateTextFormat(... "Segoe UI" ...)`
   - visible GDI text measurement and drawing entry points.
4. Refresh the inventory of remaining visible GDI-rendered surfaces and group them by migration bucket, not just by API hit count.

Engineering execution details:

1. Add a shared audit pass under `Tools/` that reports:
   - visible `DEFAULT_GUI_FONT` consumers,
   - visible `NONCLIENTMETRICS.lfMenuFont` consumers,
   - visible `CreateFontW` / `CreateFontIndirectW` call sites,
   - visible `CreateTextFormat(... "Segoe UI" ...)` call sites,
   - visible `GetTextExtentPoint32W` / `GetTextMetricsW` / owner-draw menu rendering paths.
2. Keep the audit scoped to:
   - `Common/DxUi/`
   - `RedSalamander/`
   - `Plugins/`
   - and exclude non-product tooling plus explicitly allowed shell-icon conversion helpers.
3. Check in the audit summary as a durable UI document, preferably a new `Specs/UI/UI_VisibleTypographyAudit.md`, instead of burying the inventory only in command output.
4. Refresh the initial inventory with the currently known hotspots:
   - `Common/DxUi/DxUi.WindowHost.cpp`
   - `Common/DxUi/DxUi.TextInput.cpp`
   - `RedSalamander/AppTheme.cpp`
   - `RedSalamander/FunctionBar.cpp`
   - `RedSalamander/FolderWindow.StatusBar.cpp`
   - `RedSalamander/NavigationView.cpp`
   - `RedSalamander/FolderView.Rendering.cpp`
   - `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
   - `RedSalamander/Ui/AlertOverlay.h`
   - `RedSalamander/UiMetrics.cpp` and `RedSalamander/D2DHdcPaint.cpp` for the residual helper paths that were split out after the temporary `Win32UiHelpers.cpp` staging module was retired
   - the viewer/plugin files listed in Phase 6.
5. Update `Specs/UI/UI_DxUiWinUIDesign.md`, `Specs/UI/UI_FolderView.md`, and `Specs/UI/UI_NavigationView.md` so the baseline typography contract and the current migration intent all say the same thing before code starts moving.
6. Record the expected verification entry points up front:
   - `Tests/DxUiTests`
   - `Tests/ViewerPETests`
   - `Tests/ViewerSqliteTests`
   - `RedSalamander/SelfTest/Commands/Commands.SelfTest*.cpp`
   - any plugin/viewer harnesses that must be added because no test coverage exists yet.

Exit criteria:

- The typography family contract is unambiguous in spec form.
- The repo has a repeatable audit for visible typography drift and visible GDI text/drawing surfaces.
- This plan carries an explicit inventory of the remaining buckets.

## Phase 1. Shared Typography Layer

Current problem:

- Typography policy is duplicated in `DxUi.WindowHost`, custom DWrite windows, and numerous GDI/HFONT call sites.
- A family-name flip inside `DxUi.WindowHost` alone would leave most visible surfaces inconsistent.

Required work:

1. Introduce one shared typography definition layer for app-owned UI text.
2. Move the `FontRole` to family/weight/size mapping into a shared helper that can be consumed by:
   - `DxUi::WindowHost`,
   - other DirectWrite-only windows,
   - and a temporary HFONT compatibility wrapper for legacy HWND surfaces.
3. Centralize the Windows 11 Segoe UI Variable family constants and add diagnostics if the expected family is unavailable; do not add a separate downlevel product contract in this plan.
4. Keep icons, emoji, and monospace explicit rather than inheriting the body-text family.
5. Add cache invalidation rules for:
   - DPI changes,
   - font-family availability changes,
   - theme or accessibility policy changes that affect typography.

Files likely touched:

- `Common/DxUi/DxUi.WindowHost.cpp`
- `Common/DxUi/DxUi.h`
- `Common/DxUi/DxUi.TextInput.cpp`
- new shared typography helper under `Common/DxUi/` or `Common/Ui/`

Engineering execution details:

1. Create a shared typography API with both DirectWrite and GDI/HFONT entry points. The helper should own:
   - family constants,
   - role-to-weight/size mapping,
   - emoji/icon/monospace exceptions,
   - DPI-aware size conversion,
   - text-format creation,
   - LOGFONT/HFONT creation for temporary HWND consumers.
2. Remove the local `FontSpec` source of truth from `DxUi.WindowHost.cpp`; `WindowHost::GetTextFormat(...)` and device-independent format initialization should call the shared helper instead.
3. Update `Common/DxUi/DxUi.TextInput.cpp` so the bridge no longer owns its own `Segoe UI` constant or private HFONT policy during the transition period.
4. Add debug logging or assertions that make an unexpected missing Windows 11 Segoe UI Variable family visible during development instead of silently falling back to an unrelated app font.
5. Keep the helper narrow enough that later bridge removal does not force a second typography rewrite; the hidden-bridge code should consume the same helper the visible surfaces use.
6. Add focused unit coverage for role mapping, DWrite text-format creation, and any new helper-level fallback/diagnostic behavior in `Tests/DxUiTests`.

Exit criteria:

- One shared typography helper exists for all app-owned visible UI text.
- `DxUi::WindowHost` no longer owns a private hardcoded `Segoe UI` policy.
- Legacy HWND-based surfaces have a temporary shared HFONT creation path instead of ad hoc font creation.

## Phase 2. Unify Existing DirectWrite Surfaces

Current problem:

- Several non-DxUi windows already use DirectWrite, but they bypass the shared typography contract and hardcode `Segoe UI`.

Known inventory:

- `RedSalamander/FolderView.Rendering.cpp`
- `RedSalamander/NavigationView.Rendering.cpp`
- `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- `RedSalamander/Ui/AlertOverlay.h`
- any other custom DWrite surfaces found by the Phase 0 audit

Required work:

1. Route all custom DirectWrite text creation through the shared typography layer.
2. Remove local hardcoded text family strings for visible text.
3. Normalize body/detail/header/overlay sizes and weights to the same role contract used by DxUi.
4. Refresh measurement-sensitive layouts where Segoe UI Variable metrics differ from the old Segoe UI assumptions.
5. Add focused rendering tests for affected surfaces where text metrics drive hit testing, truncation, or row heights.

Engineering execution details:

1. Start with the largest always-visible DWrite surfaces:
   - `RedSalamander/FolderView.Rendering.cpp`
   - `RedSalamander/NavigationView.Rendering.cpp`
   - `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
   - `RedSalamander/Ui/AlertOverlay.h`
2. Replace local hardcoded family strings with calls into the shared typography helper while preserving each surface's semantic role split:
   - body vs details vs metadata in FolderView,
   - breadcrumb/path vs separator/icon roles in NavigationView,
   - header/body/button/small/graph overlay/status icon roles in File Operations popup,
   - title/body/button/icon roles in AlertOverlay.
3. Recompute size-sensitive constants that were implicitly tuned to Segoe UI metrics:
   - row heights,
   - truncation widths,
   - watermark and empty-state scaling,
   - overlay text placement,
   - badge and focus-cue layout.
4. Refresh rendering and behavior coverage where metric drift can break behavior:
   - `Tests/DxUiTests/DxUiTests.Rendering.cpp`
   - `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`
   - file-operations and dialog selftests under `RedSalamander/SelfTest`.
5. Capture before/after visual baselines for FolderView rows, NavigationView, file-operations popup cards, and alert overlays.

Exit criteria:

- Zero remaining visible custom DirectWrite text paths hardcode `Segoe UI` for normal UI text.
- All visible DWrite surfaces use the same role-based typography contract.
- Layout and hit-test regressions caused by metric drift are covered by tests or baselines.

## Phase 3. Remove HFONT/System-Font Drift From Remaining HWND UI

Current problem:

- Many visible HWND-based surfaces still derive fonts from `DEFAULT_GUI_FONT` or `lfMenuFont`.
- Even before the full GDI-retirement phase, this keeps typography inconsistent.

Known inventory:

- `RedSalamander/FunctionBar.cpp`
- `RedSalamander/FolderWindow.StatusBar.cpp`
- `RedSalamander/NavigationView.cpp` path font
- `RedSalamander/Preferences.Dialog.cpp`
- `RedSalamander/ConnectionManagerDialog.cpp`
- `RedSalamander/ManagePluginsDialog.cpp`
- `RedSalamander/CompareDirectoriesWindow.cpp`
- additional `WM_SETFONT` consumers discovered by the Phase 0 audit

Required work:

1. Replace direct `DEFAULT_GUI_FONT` usage on visible app surfaces with the shared compatibility font helper.
2. Replace direct `CreateFontW(..., L"Segoe UI")` usage with the same helper.
3. Stop deriving app-owned visible UI from `NONCLIENTMETRICS.lfMenuFont` on non-menu surfaces immediately; retain the menu-specific path only until Phase 4 menu retirement removes it entirely.
4. Rework title/semibold/italic/secondary variants so they derive from the shared app typography contract rather than cloning whatever the OS default GUI font happens to be.
5. Document which HWND-based surfaces are still temporary and scheduled for full GDI retirement in later phases.

Engineering execution details:

1. Convert the non-menu HWND/HFONT surfaces in this order:
   - `RedSalamander/FunctionBar.cpp`
   - `RedSalamander/FolderWindow.StatusBar.cpp`
   - `RedSalamander/NavigationView.cpp`
   - `RedSalamander/CompareDirectoriesWindow.cpp`
   - `RedSalamander/Preferences.Dialog.cpp`
   - `RedSalamander/ConnectionManagerDialog.cpp`
   - `RedSalamander/ManagePluginsDialog.cpp`
   - `RedSalamander/FolderWindow.FileSystem.cpp`
2. Replace raw base-font cloning with explicit role-based requests from the shared helper so title, bold, italic, caption, and comment fonts all derive from app policy rather than ambient OS defaults.
3. Keep temporary HFONT usage behind one compatibility seam so later GDI retirement becomes “delete the seam” rather than “audit every dialog again”.
4. Use this phase to delete `RedSalamander/ThemedControls.h/.cpp` once all visible themed-control callers are gone; any residual helper must move to a neutral module rather than staying in a themed-control file. The later closeout retired the temporary `RedSalamander/Win32UiHelpers.h/.cpp` staging module in favor of `RedSalamander/UiMetrics.h/.cpp` and documented allowlisted HDC interop.
5. Refresh command/selftest coverage for dialogs and bars that currently assert visible DX or themed shell behavior:
   - `Commands.SelfTest.Dialogs.cpp`
   - `Commands.SelfTest.Connections.cpp`
   - `Commands.SelfTest.Preferences*.cpp`
   - `Commands.SelfTest.ViewCommands.cpp`
6. Capture targeted visual verification for FunctionBar, status bars, preferences/dialog shells, and compare-directories chrome after each bucket lands.

Exit criteria:

- No visible app-owned HWND surface uses raw `DEFAULT_GUI_FONT`.
- No visible app-owned HWND surface derives its body/title fonts from ad hoc `GetObjectW(baseFont, ...)` clones of the OS default font.
- Remaining menu-font usage, if any, is explicitly transitional and tracked.

## Phase 4. Retire Remaining Visible GDI Surface Families

Current problem:

- Typography unification will still look inconsistent until the remaining GDI-rendered visible surfaces are gone.
- GDI menus, bars, and owner-draw helpers are also the last major source of divergent measurement, rasterization, hover/press visuals, and theme behavior.

Workstreams:

### 4.1 Menus and Popup Surfaces

Targets:

- main app menu rendering in `RedSalamander/RedSalamander.cpp`
- compare-directories command/menu chrome
- remaining viewer/plugin owner-draw menus identified by the audit

Required work:

1. Finish routing these surfaces onto DxUi menu/popup infrastructure or an equivalent D2D/DWrite path.
2. Retire GDI text measurement and owner-draw painting for visible menu items.
3. Apply the shared Segoe UI Variable typography contract to menu text once the surface is no longer native-font driven.
4. Revalidate standard Windows menu-loop keyboard behavior while changing rendering ownership.

Engineering execution details:

1. Remove active dependence on:
   - `RedSalamander/AppTheme.cpp` `CreateMenuFontForDpi(...)`
   - owner-draw measurement in `RedSalamander/RedSalamander.cpp`
   - plugin/viewer menu-HFONT paths identified in Phase 6.
2. Route menus to shared DxUi popup/menu infrastructure or an equivalent D2D/DWrite pipeline with:
   - text measurement from DirectWrite,
   - icon measurement from Fluent/shell bitmap paths,
   - shared mnemonic / shortcut layout,
   - and shared theme/typography policy.
3. Keep Windows-standard menu-loop behavior while swapping rendering ownership:
   - keyboard root switching,
   - submenu close/open delays,
   - stationary-mouse behavior,
   - `Alt`, `F10`, `Tab`, and `Escape` dismissal.
4. Refresh the menu regression suite in `Tests/DxUiTests/DxUiTests.Menu.cpp` and the app command selftests that exercise the main menu and navigation popups.

### 4.2 Function Bar and Status Bars

Targets:

- `RedSalamander/FunctionBar.cpp`
- `RedSalamander/FolderWindow.StatusBar.cpp`
- any remaining monitor/app status-strip-like surfaces still on GDI paths

Required work:

1. Replace visible GDI drawing and text measurement with DxUi/D2D surfaces.
2. Retire `HFONT`-driven paint logic and `DEFAULT_GUI_FONT` fallbacks.
3. Reuse shared typography roles for function labels, modifiers, status text, and emphasis text.

Engineering execution details:

1. Replace `GetTextExtentPoint32W`, `GetTextMetricsW`, and HFONT-select painting in:
   - `RedSalamander/FunctionBar.cpp`
   - `RedSalamander/FolderWindow.StatusBar.cpp`
2. Prefer existing DxUi primitives where they match the product surface:
   - `StatusStrip`
   - shared text measurement/layout helpers
   - any already-landed toolbar/status-strip work from the WinUI-alignment rollout
3. Where a custom retained control is still needed, keep the text path DirectWrite-only and source all typography from the shared helper.

### 4.3 Dialog Shells and Retired ThemedControls Consumers

Targets:

- legacy dialog chrome previously painted or measured through `ThemedControls.cpp`
- dialog shells still relying on `WM_SETFONT` plus owner-draw helpers
- remaining hybrid dialog pages in Preferences and related windows

Required work:

1. Replace visible GDI owner-draw surfaces with DxUi or D2D/DWrite equivalents.
2. Stop using GDI for visible title/body/comment/toggle/button text on migrated shells.
3. Keep `RedSalamander/ThemedControls.h/.cpp` retired; generic HWND helper code must not return to a themed-control module, and the final remaining-Win32 closeout keeps pure helpers in `RedSalamander/UiMetrics.h/.cpp` instead of reviving `Win32UiHelpers`.

Engineering execution details:

1. Confirm every former visible `ThemedControls.cpp` caller remains migrated; classify any newly found visible themed-control helper as:
   - replaceable with shared typography without changing rendering ownership,
   - or requiring a move to DxUi / D2D retained rendering.
2. Remove text measurement helpers that still depend on:
   - `GetTextExtentPoint32W`
   - `GetTextMetricsW`
   - `DEFAULT_GUI_FONT`
3. Keep the shell transition incremental by moving one dialog family at a time, but do not leave any visible dialog text on ambient OS font defaults after the phase completes.

Exit criteria for Phase 4:

- No visible application-owned menu, function bar, status bar, or dialog shell remains GDI-rendered.
- Text on those surfaces comes from the shared typography layer.

## Phase 5. Navigation and Hybrid Window Closure

Current problem:

- `NavigationView` already renders most visible content with D2D/DWrite, but still mixes rendering systems:
  - explicit GDI path-font creation,
  - transient Win32 edit control usage in edit mode,
  - and older popup/menu compatibility layers.

Required work:

1. Remove the remaining visible font split between NavigationView’s DWrite path rendering and its GDI/HFONT-backed helpers.
2. Retire any remaining visible GDI fallback painting that survives pre-D2D startup or edit-mode transitions.
3. Replace the transient visible edit control path with a pure DxUi text-entry path or another non-visible helper path that does not reintroduce a visible Win32/GDI edit surface.
4. Keep the NavigationView spec aligned with the final rendering ownership and font contract.

Engineering execution details:

1. Treat `NavigationView` and the hidden DxUi bridge as one closure bucket:
   - visible breadcrumb/path/edit surfaces,
   - any visible edit popup or full-path popup text surfaces,
   - and the hidden `Common/DxUi/DxUi.TextInput.cpp` dependency.
2. Rework the input path so IME, clipboard, undo/redo, caret, and accessibility behavior are preserved without keeping a hidden product-time edit control.
3. Move bridge-specific regression coverage into direct host/control coverage where possible:
   - `Tests/DxUiTests/DxUiTests.TextField.cpp`
   - `Tests/DxUiTests/DxUiTests.MultilineText.cpp`
   - `Tests/DxUiTests/DxUiTests.TextInputBridge.cpp` while the bridge still exists,
   - then replace bridge-only assertions with end-state host/control assertions as the bridge is removed.
4. Refresh `Commands.SelfTest.Navigation.cpp` and any NavigationView-specific close/reopen, edit-mode, popup, and DPI tests after each migration slice.

Exit criteria:

- NavigationView presents one visible typography/rendering path.
- Any remaining Win32 helper is non-visible and explicitly documented as such.

## Phase 6. Viewer and Plugin Surface Cleanup

Current problem:

- The broader `Specs/Plans/Done/DxUI_MigrationRoadmap.md` historically tracked plugin/viewer windows and several custom UI surfaces with visible GDI dependencies; newer window-migration breadth was merged into `Specs/Plans/Done/UI_DxUiRemainingMigrationCloseoutPlan.md`.
- Typography will not be globally uniform until those surfaces are migrated too.

Required work:

1. Refresh the plugin/viewer inventory from `Specs/Plans/Done/DxUI_MigrationRoadmap.md` and the completed closeout record in `Specs/Plans/Done/UI_DxUiRemainingMigrationCloseoutPlan.md`.
2. Separate surfaces into:
   - already D2D/DWrite,
   - visible GDI surfaces that must migrate in this plan,
   - non-visible compatibility or hosting seams that can remain.
3. For each visible GDI viewer/plugin surface, define:
   - target rendering owner,
   - target typography roles,
   - verification owner and tests.

Engineering execution details:

1. Use the current inventory as the minimum in-scope file set:
   - `Plugins/ViewerText/ViewerText.cpp`
   - `Plugins/ViewerText/ViewerText.Text.cpp`
   - `Plugins/ViewerText/ViewerText.Hex.cpp`
   - `Plugins/ViewerText/ViewerText.MenuTheme.cpp`
   - `Plugins/ViewerWeb/ViewerWeb.cpp`
   - `Plugins/ViewerSpace/ViewerSpace.cpp`
   - `Plugins/ViewerImgRaw/ViewerImgRaw.cpp`
   - `Plugins/ViewerPE/ViewerPE.cpp`
   - `Plugins/ViewerVLC/ViewerVLC.cpp`
   - `Plugins/FileSystem/FileSystem.Menu.cpp`
   - `Plugins/FileSystemS3/FileSystemS3.Menu.cpp`
2. For each plugin window, explicitly separate:
   - visible UI text surfaces,
   - visible owner-draw menus,
   - hidden/non-visible hosting seams,
   - monospace/document-view text that legitimately keeps a non-UI font role.
3. Move plugin UI fonts onto the shared helper before deleting menu/GDI paths so typography and rendering retirement do not drift apart.
4. Where viewer content intentionally uses monospace or document fonts, keep that as an explicit role-based exception rather than an implicit hardcoded family string.
5. Add or extend verification for plugin windows through:
   - `Tests/ViewerPETests/ViewerPETests.cpp`
   - `Tests/ViewerSqliteTests/ViewerSqliteTests.cpp`
   - command selftests that open plugin windows or plugin-backed commands,
   - and any missing viewer harnesses that must be added for windows without current coverage.

Exit criteria:

- Every remaining visible GDI viewer/plugin surface in the repo is migrated in this plan.

## Validation Requirements

### Static Audits

The closeout audit must prove that visible surfaces no longer depend on:

- `GetStockObject(DEFAULT_GUI_FONT)`
- `NONCLIENTMETRICS.lfMenuFont`
- visible `CreateFontW(...)` / `CreateFontIndirectW(...)` family selection outside the shared compatibility helper
- visible `CreateTextFormat(..., L"Segoe UI", ...)` outside the shared typography helper or documented fallback path
- visible GDI text/drawing APIs on application-owned UI surfaces

Engineering execution details:

- The audit must scan both `RedSalamander/` and `Plugins/`.
- The audit output should group hits by migration bucket: shared helper, DWrite surface, HWND/HFONT surface, menu/popup surface, hidden/non-visible helper, and allowed exception.
- The audit should fail or at least highlight new hits introduced outside the allowlist once the first migration slices land.

The audit must also keep an explicit allowlist for:

- shell icon conversion helpers
- non-app-owned native dialogs

### Functional Validation

At minimum, rerun and refresh coverage for:

- menu keyboard loops and submenu behavior
- DPI changes on main-window chrome and popups
- FolderView and NavigationView truncation, hit testing, and hover behavior
- dialog focus/default/cancel flows after text-measurement changes
- theme changes and high-contrast behavior on migrated text surfaces

Engineering execution details:

- Refresh the focused `Tests/DxUiTests` suites:
  - `DxUiTests.Menu.cpp`
  - `DxUiTests.Rendering.cpp`
  - `DxUiTests.TextField.cpp`
  - `DxUiTests.MultilineText.cpp`
  - `DxUiTests.TextInputBridge.cpp` until the bridge is removed
  - `DxUiTests.WindowHost.cpp`
- Refresh or add command selftests under `RedSalamander/SelfTest/Commands/` for:
  - menus and view commands,
  - navigation edit/popup flows,
  - preferences/dialog shells,
  - plugin-backed dialogs or windows,
  - file-operations popup surfaces.
- Refresh viewer/plugin harnesses for windows that do not participate in the command selftest layer.

### Visual Validation

Refresh or add visual baselines for:

- menus
- FolderView rows and empty states
- NavigationView path bar and dropdowns
- function bar
- status bars
- migrated dialog shells
- plugin/viewer menu and chrome surfaces that currently use HFONT/GDI text paths

### Performance Validation

Because these changes affect hot UI paths, each migrated cluster must archive evidence under `Specs/TestRuns/` for:

- window open / first render,
- menu popup show latency,
- steady-state hover and repaint behavior,
- layout churn under DPI change,
- and any large-list or large-folder surfaces where text layout cost matters.

Engineering execution details:

- Archive at least one before/after run for each bucket:
  - app menus and navigation popups,
  - FolderView,
  - NavigationView,
  - file-operations popup,
  - FunctionBar/status bars,
  - dialog shells,
  - each migrated plugin/viewer family.
- Where metric drift changes cached-layout behavior, include explicit layout/create counts or equivalent instrumentation in the evidence.

## Completion Criteria

This plan is complete only when all of the following are true:

- the authoritative UI specs describe the final typography contract,
- visible application-owned text is unified on the shared Segoe UI Variable contract with explicit fallback rules,
- the hidden DxUi text-input bridge is removed from the active product path,
- no visible app-owned surface remains on raw GDI text rendering or system-font-driven drift paths,
- static audits have a documented allowlist and otherwise pass,
- focused automated tests and visual baselines pass for the migrated clusters,
- required perf evidence is archived,
- and this plan is moved to `Specs/Plans/Done/` with the learned durable rules merged into the authoritative specs rather than left only here.

## Resolved Early Decisions

These decisions are fixed for execution of this plan:

1. Windows 11 is the baseline for this effort, and Segoe UI Variable is the required UI family contract.
2. Closure includes all visible app-owned surfaces plus the hidden DxUi text-input bridge.
3. Menus switch to the shared app typography only when the GDI owner-draw menu path is fully retired.
4. All plugin/viewer windows are in scope for this plan.

No additional product-scope questions are currently blocking execution of the plan. The remaining choices are implementation details such as helper API shape, migration order inside each phase, and validation sequencing.
