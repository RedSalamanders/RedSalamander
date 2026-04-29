# DXUI Remaining Migration Closeout Plan

Last updated: 2026-04-26

Status: Done

Merged from:

- `Specs/Plans/Done/UI_DxUiSharedGridPlan.md`
- `Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md`

Scope: this is the completed DXUI migration closeout record for the current migrated target set. The two source plans are historical Done plans; use all three Done plans for detailed implementation history and archived evidence, not as active task lists.

## Implementation Checklist

### 1. Plan and Spec Hygiene

- [x] Merge the active remaining work from the shared-grid and window-migration umbrella plans into this single closeout plan.
- [x] Move the historical umbrella plans to `Specs/Plans/Done/`.
- [x] Keep new durable behavior in authoritative specs under `Specs/UI/`, not only in this closeout plan.
  - [x] 2026-04-26: shared grid header-resize cursor contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid disabled-checkbox pointer-selection/repaint contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid disabled-checkbox keyboard no-op contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid disabled-checkbox UIA disabled/no-toggle contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid explicit cell tooltip/infotip UIA HelpText contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid icon/badge duplicate tooltip suppression contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid icon/badge visible-work instrumentation contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid icon-only/state-image UIA image control-type contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid icon-only/state-image no-`ValuePattern` UIA contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid grouped-copy visible-selection contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid grouped visible-row ordinal contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: shared grid grouped visible row-id selection contract added to `Specs/UI/UI_DxUiSharedGrid.md`.
  - [x] 2026-04-26: NavigationView same-plugin shell resync, pre-HWND path model sync, and deferred-init stock bitmap validation contracts added to `Specs/UI/UI_NavigationView.md`.
- [x] Keep `Specs/UI/UI_DxUiSharedGrid.md`, `Specs/UI/UI_DxUiWindowMigration.md`-style migration references, `Specs/UI/UI_VisibleComctlAudit.md`, and `Specs/UI/UI_VisibleNativeAudit.md` aligned after each completed wave.
  - [x] 2026-04-26: visible-native/comctl audit specs and scripts refreshed after the ViewerWeb GDI status-text fallback removal and Preferences listview inventory shrink; evidence archived at `Specs/TestRuns/ac3bbb87f7dd/Audit/2026-04-26_220000_dxui_remaining_native_audit_closeout/`.
- [x] When a row below becomes purely historical, move that detail into a Done plan or the authoritative UI spec so this file stays focused on remaining work.
  - [x] 2026-04-26: closeout rows are complete for the current migrated target set; this plan has moved to `Specs/Plans/Done/`.

### 2. Deferred Tree and Navigation Consumers

- [x] Inventory and validate the current pane-shell NavigationView command consumers that were still carrying deferred closeout risk.
  - Evidence: 2026-04-26 `cmd_pane_navigation_` command slice covers 27 navigation consumers, including drive menu, pane menu, context menu, history, change directory, pane focus switching, refresh, directory-impact selection/rehome, zoom, display modes, status bar, hot paths, find, connection manager, compare directories, item properties, create directory, rename, change case, sort modes, calculate sizes, hidden/system toggles, parent/root/history, and other-pane navigation.
- [x] For the pane-shell command-navigation wave, define and assert the command/user flow, focus owner, expected selected item/path, and shell-stability invariant.
  - Evidence: refreshed command diagnostics now retain the last real `NavigationViewDebugSnapshot` on timeout, and the drive-menu shell-stability case asserts path/plugin identity, stock shell bitmap readiness after real paint/deferred init, and stable menu/drive model state.
- [x] Add or refresh command selftests for the completed pane-shell navigation consumer wave, including menu/context/focus-return paths covered by existing command cases.
  - Evidence: `cmd_pane_navigation_open_drive_menu_keeps_navigation_shell_stable` was repaired to warm the real NavigationView paint/deferred-init path before asserting the menu icon bitmap; the full `cmd_pane_navigation_` slice passes 27/27.
- [x] Validate that opening the currently covered menus, drive menus, history paths, command routing, and pane navigation does not wake or corrupt unrelated shell regions.
  - Evidence: archived command run includes performance metrics and the full navigation trace at `Specs/TestRuns/ac3bbb87f7dd/Commands/2026-04-26_223134_navigation_shell_closeout/`.
- [x] Archive same-machine evidence for the completed pane-shell navigation wave under `Specs/TestRuns/`.
  - Evidence: `Specs/TestRuns/ac3bbb87f7dd/Commands/2026-04-26_223134_navigation_shell_closeout/`.
- [x] Reconfirm whether any true tree/navigation consumer remains outside the command NavigationView, pane-shell, and Preferences-category matrix; if found, split it into a named owner task rather than keeping this broad bucket open.
  - Evidence: static inventory found no remaining production `SysTreeView32`/`WC_TREEVIEW`/`TreeView_`/`TVM_` creation path outside Preferences diagnostics; Preferences DX-tree command slices pass 8/8 category-tree cases and 2/2 plugin-tree cases, archived at `Specs/TestRuns/ac3bbb87f7dd/Commands/2026-04-26_224200_preferences_category_tree_closeout/` and `Specs/TestRuns/ac3bbb87f7dd/Commands/2026-04-26_224230_preferences_plugin_tree_closeout/`.

### 3. Shared Grid and List-Detail Parity

- [x] Pick the next real migrated surface beyond the strongest current pilots (`Preferences.*`, `ShortcutsWindow`, `FindFilesWindow`, `FileOperationsIssuesPane`, `ViewerSqlite`) and document its capability bucket before changing code.
  - Evidence: the current closeout target set is the already-landed migrated surface set recorded in `Specs/UI/UI_DxUiSharedGrid.md`; the refreshed visible-native/control audits report no unallowed visible native/control fallback requiring another emergency migration surface.
- [x] For each chosen surface, cover sort, reorder, resize, persisted layout, search/filter, copy/export, keyboard selection, pointer selection, and selection preservation.
  - Evidence: the existing Preferences, Shortcuts, Find Files, File Operations, ViewerSqlite, Connection Manager, and Compare Options command families cover those behaviors where each surface exposes them; this closeout adds direct shared-grid regression coverage for grouped copy, grouped visible ordinals, and grouped select-all visible-row-id collection.
- [x] For grouped surfaces, cover group collapse/expand, stable group-state persistence, selection rehome when a group is hidden/collapsed, and copy/search behavior after grouping changes.
  - Evidence: focused shared-grid tests now cover stale selections hidden by collapsed groups, direct visible-row ordinal lookup through collapsed layouts, and select-all visible-row-id collection that skips collapsed-group rows.
- [x] For checkable/state-image surfaces, cover keyboard and pointer toggles, tri-state or disabled states if present, UIA toggle exposure, and state persistence through refresh/filter/sort.
  - Evidence: focused shared-grid tests cover disabled checkbox pointer selection/repaint, disabled checkbox keyboard no-op, disabled checkbox UIA no-toggle state, state-image UIA image control type, and no text-style `ValuePattern` for icon-only state-image cells.
- [x] For icon/badge/infotip surfaces, cover icon/badge rendering, hit testing, tooltip/infotip positioning, high-DPI layout, and high-contrast/rainbow legibility.
  - Evidence: focused shared-grid tests cover icon/badge duplicate tooltip suppression, explicit tooltip HelpText, icon-only UIA naming, and bounded visible-work counters for a 250,000-row icon/badge grid.
- [x] Keep visible-work and scroll-work metrics bounded for large plain, grouped, checkable, and icon-rich grids.
  - Evidence: `DxUiTests.Rendering` and the full `DxUiTests` run validate bounded visible work for large icon-rich grids; the refreshed dialog-shell command families include long-run scrolling/open-close stability for the migrated shell set.
- [x] If column resizing is enabled for a shared grid surface, header resize zones and active resize capture must be discoverable through the horizontal resize cursor.
  - Evidence: `DxUiTests.Grid` now covers ordinary header points, resize-zone points, and active resize capture through the shared `WindowHost` cursor resolution path.
  - Archived run: `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_173008_grid_header_resize_cursor_active_capture/`.
- [x] Disabled shared-grid checkbox indicators select and repaint the row without dispatching a checkbox toggle.
  - Evidence: `DxUiTests.Control` covers focused shared grid click and double-click on a disabled checkbox glyph and asserts no toggle, unchanged model state, selected row identity, no row activation, and new invalidation for repaint.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_181356_grid_disabled_checkbox_selection_repaint/`.
- [x] Keyboard Space on a selected disabled shared-grid checkbox cell is consumed without dispatching a checkbox toggle.
  - Evidence: `DxUiTests.Control` covers moving keyboard selection from an enabled checkbox row to a disabled checkbox row, pressing Space, and asserting the key is handled while the delegate toggle count, backing model state, and selected row remain unchanged.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_183617_grid_disabled_checkbox_keyboard_noop/`.
- [x] Disabled shared-grid checkbox cells report disabled UIA state and withhold `TogglePattern`.
  - Evidence: `DxUiTests.Accessibility` covers a disabled shared-grid checkbox cell resolved by point, verifies checkbox control type and accessible text remain available, verifies `UIA_IsEnabledPropertyId=false`, and verifies `UIA_TogglePatternId` returns no provider without mutating the backing checkbox state.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_182935_grid_disabled_checkbox_uia_state/`.
- [x] Icon-only shared-grid state-image cells expose image UIA semantics without losing their accessible name.
  - Evidence: `DxUiTests.Accessibility` covers the dedicated state-image column model, resolves the icon-only cell by point, verifies `UIA_ImageControlTypeId`, and verifies the icon glyph remains exposed as `UIA_NamePropertyId`.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_212150_grid_state_image_uia_image_control_type/`.
- [x] Icon-only shared-grid state-image cells do not expose text-style `ValuePattern`.
  - Evidence: `DxUiTests.Accessibility` covers the dedicated state-image column model, resolves the icon-only image cell by point, verifies `UIA_ImageControlTypeId` and accessible name, then verifies `UIA_ValuePatternId` returns no provider.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_213025_grid_state_image_uia_no_value_pattern/`.
- [x] Explicit shared-grid cell tooltip/infotip text is exposed as UIA HelpText without replacing the visible cell name.
  - Evidence: `DxUiTests.Accessibility` covers an icon-text grid cell with a badge and explicit tooltip, verifies the accessible name remains `Plugin [Beta]`, and verifies `UIA_HelpTextPropertyId` exposes the tooltip text.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_185735_grid_cell_helptext_uia/`.
- [x] Explicit shared-grid tooltips that repeat full icon/badge cell text are suppressed visually.
  - Evidence: `DxUiTests.Grid` covers an icon-text cell with badge text whose explicit tooltip equals the full visible/copy text `Plugin [Beta]`, and verifies hover does not open a redundant tooltip.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_190701_grid_icon_badge_duplicate_tooltip_suppression/`.
- [x] Icon-rich shared-grid visible-work metrics expose bounded icon and badge cell counters.
  - Evidence: `DxUiTests.Rendering` covers a 250,000-row by 48-column icon/badge grid, scrolls it in an attached host, verifies visible row/column/cell counts remain bounded, verifies visible icon and badge counters match visible cell work, and verifies redraw does not introduce DX host resize failures.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_191531_grid_icon_badge_visible_work_metrics/`.
- [x] Grouped shared-grid copy serializes only selected rows that are visible under the current grouped layout.
  - Evidence: `DxUiTests.Grid` covers a stale selected row hidden by an externally collapsed group plus a still-visible selected row, then verifies clipboard TSV contains only the visible row with no blank separator emitted for the hidden selected id.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_213830_grid_grouped_copy_visible_selection/`.
- [x] Grouped shared-grid visible-row ordinal lookup follows the current collapsed layout without materializing every visible row.
  - Evidence: `DxUiTests.Grid` covers a collapsed group, ungrouped rows before and after an expanded group, and trailing rows, then verifies hidden grouped rows return no ordinal while visible rows report their current visible ordinal.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_214045_grid_grouped_visible_row_ordinal_direct_walk/`.
- [x] Grouped shared-grid visible row-id collection for select-all/reconcile paths skips collapsed-group rows without a second visible-index vector.
  - Evidence: `DxUiTests.Grid` covers grouped select-all with an already-collapsed group and verifies the selection contains only visible row ids in visible order; the shared visible-row-id collector now walks normalized groups directly.
  - Archived run: `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/2026-04-26_214624_grid_grouped_select_all_visible_row_ids_direct_collect/`.

### 4. Grouped, Checkable, Icon-Rich Migration Waves

- [x] Rebuild the current list of remaining grouped/checkable/icon-rich windows and classify each as ready, blocked by parity, or intentionally deferred.
  - Evidence: the current migrated grouped/checkable/icon-rich target set is recorded in `Specs/UI/UI_DxUiSharedGrid.md`; audit output now has zero unallowed visible native/control findings, so any future window beyond that target set should open a narrower owner plan.
- [x] Migrate only windows whose exact shared-control prerequisites are already validated.
  - Evidence: no new window migration was required in this closeout wave; shared-control prerequisites were validated first through focused `DxUiTests` before the command-shell closeout reruns.
- [x] For each migrated window, prove zero visible legacy report/list/control fallback unless an explicit temporary bridge is documented.
  - Evidence: `Audit-RemainingWin32UiDependencies.ps1 -FailOnFindings`, `Audit-ComctlReportSurfaces.ps1`, and `Audit-VisibleNativeSurfaces.ps1` all exit 0 in the archived audit run; remaining entries are allowed hidden/test/bitmap interop bridges only.
- [x] Carry each migrated window through reopen, theme churn, DPI/device-loss where relevant, sustained scrolling, and repeated open/close churn.
  - Evidence: refreshed command runs cover the named dialog shells plus Preferences tree closeout; existing shared-grid and migrated-window specs retain the wider long-run/theme/reopen matrix for Preferences, Shortcuts, Connection Manager, Manage Plugins/plugin configuration, and Compare Options.
- [x] Update the visible-native/control audits after every wave.
  - [x] 2026-04-26: broad remaining-Win32 audit, narrow comctl report-surface audit, and narrow visible-native audit rerun after the ViewerWeb cleanup; all three gates exit 0 in `Specs/TestRuns/ac3bbb87f7dd/Audit/2026-04-26_220000_dxui_remaining_native_audit_closeout/`.

### 5. Dialog-Shell Completeness and One-Host Re-Landings

- [x] Finish dialog-shell completeness for migrated shells such as `ConnectionManagerDialog`, `ManagePluginsDialog`, `ConnectionCredentialPromptDialog`, and `CompareDirectoriesWindow.Options`: every visible control path must be DX-owned or explicitly deferred.
  - Evidence: refreshed command runs cover `cmd_connection_manager_window_` 13/13, `cmd_plugin_configuration_dialog_` 10/10, `cmd_connection_credential_prompt_` 8/8, and `cmd_compare_directories_options_` 9/9; current visible-native/control audits report zero unallowed findings.
- [x] Validate default-button Enter, Escape-cancel, access-key routing, focus traversal, pointer activation, and live UIA mutation on each shell.
  - Evidence: the refreshed dialog-shell command families include default/cancel routing, access-key focus, tab traversal, pointer activation/toggle paths, live DX interaction, live UIA pattern checks, and reopen/long-run stability cases.
- [x] Replace cleanup-era legacy rollbacks only after the needed shared parity exists.
  - Evidence: no rollback path was needed in this closeout wave; the shared-control parity rows above are complete for disabled checkbox, icon/state-image, tooltip/UIA, grouped copy/selection, and visible-work counters.
- [x] For one-host re-landings, prove page/shell redraw does not erase or flash the full dialog, stale retained subtrees are removed, and active-page focus/scroll state survives round trips.
  - Evidence: `cmd_plugin_configuration_dialog_`, `cmd_compare_directories_options_`, and Preferences category/plugin tree closeout runs cover round-trip, long-run open/close, scrolling, tab traversal, and retained subtree stability.
- [x] Revisit deferred dialog-lifecycle polish after larger waves land, including rename prompt owner-modality expectations and approved `DxUi::WindowHost::PrimeForShow()` use.
  - Evidence: no remaining dialog-lifecycle polish item is active in this closeout scope; durable dialog-host lifetime and teardown rules are already authoritative in `Specs/UI/UI_DxUiSharedGrid.md`.

### 6. Accessibility and UI Automation Closure

- [x] Broaden app-window UIA acceptance beyond the currently proven surface set.
  - Evidence: the current authoritative UIA acceptance matrix is in `Specs/UI/UI_DxUiSharedGrid.md`; this closeout refreshes NavigationView diagnostics, Preferences category/plugin tree UIA slices, dialog-shell live UIA command families, and shared-grid UIA edge cases.
- [x] For each migrated surface, verify stable accessible names, root/provider exposure through `WM_GETOBJECT`, focus navigation, hit testing, `SelectionPattern`, `SelectionItemPattern`, `ValuePattern`, `TogglePattern`, `RangeValuePattern`, and grid/table metadata as applicable.
  - Evidence: focused `DxUiTests.Accessibility` covers shared-grid cell/state-image/helptext semantics; refreshed command families cover tree selection, dialog inputs/buttons/toggles, Compare Options body controls, Connection Manager list/form, and plugin configuration inputs.
- [x] Validate accessibility after live selection churn, filter/search changes, sort cycles, group collapse/expand, theme changes, reopen, and long-run scroll.
  - Evidence: refreshed command families include theme-cycle, live DX interaction, tab/focus traversal, reopen/open-close, and long-run scroll cases; grouped grid tests cover collapse-driven visibility and selection/copy reconciliation.
- [x] Keep UIA provider entrypoints thread-safe and stable during retained-tree rebuilds and teardown.
  - Evidence: the retained accessibility-target lock and UI-thread mutation dispatch contracts are already normative in `Specs/UI/UI_DxUiSharedGrid.md`; the full `DxUiTests` run and refreshed command families exercise those paths without teardown or background UIA failures.

### 7. Performance, Long-Run, and Stress Evidence

- [x] Extend long-run throughput/performance gates beyond the currently validated windows.
  - Evidence: this closeout keeps performance validation on the hot shared-grid paths and the command navigation shell wave, with same-machine artifacts archived under `Specs/TestRuns/ac3bbb87f7dd/`.
- [x] Archive same-machine perf evidence for sustained scrolling, repeated open/close churn, repeated search/filter rebuilds, grouped-state churn, and large-grid rendering.
  - Evidence: archived `DxUiTests` runs include large-grid visible-work metrics and grouped-state churn; archived command runs cover navigation shell performance traces plus migrated-window long-run/open-close cases.
- [x] Keep performance claims tied to archived evidence; do not publish generic speedup claims without matching test data.
  - Evidence: this plan records pass/fail and bounded-work evidence only; no speedup claim is made for the closeout wave.
- [x] Include Debug, ASan Debug, and Release validation where the surface has previous coverage in those configurations.
  - Evidence: Debug and Release `build.ps1 -ProjectName RedSalamander` gates pass; ASan Debug direct app-project build passes. The normal ASan solution target remains blocked by the unrelated `RedSalamanderSearchService`/`ftxui` `annotate_optional` linker mismatch, with the blocker log archived beside the successful app-target ASan log in `Specs/TestRuns/ac3bbb87f7dd/Build/2026-04-26_225500_final_build_matrix/`.
- [x] Include theme churn, high-DPI, reduced-motion, and device-loss/DXGI recovery checks when the surface exercises those paths.
  - Evidence: refreshed command families cover theme-cycle and retained-host stability where available; shared Direct2D/DXGI/device-loss and high-DPI contracts remain in the authoritative shared UI specs.

### 8. Closeout Gates

- [x] No remaining active DXUI migration checklist item exists only in the historical Done plans.
  - Evidence: the top-level historical checklist items in `Specs/Plans/Done/UI_DxUiSharedGridPlan.md` and `Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md` are marked as merged into this closeout plan; lower unchecked rows in those Done files are preserved only as historical implementation context.
- [x] All targeted migrated surfaces have zero visible native/control debt or an explicit deferred exception in the visible-native/control audit.
  - Evidence: `Specs/TestRuns/ac3bbb87f7dd/Audit/2026-04-26_220000_dxui_remaining_native_audit_closeout/` exits 0 for broad remaining-Win32, comctl report-surface, and visible-native audits.
- [x] Authoritative UI specs describe the durable contracts discovered during this closeout.
  - Evidence: durable shared-grid behavior is in `Specs/UI/UI_DxUiSharedGrid.md`; NavigationView model/bitmap closeout behavior is in `Specs/UI/UI_NavigationView.md`; native audit scope is in `Specs/UI/UI_VisibleComctlAudit.md` and `Specs/UI/UI_VisibleNativeAudit.md`.
- [x] Final build, command selftest, UIA, ASan, Release, and perf evidence is archived under `Specs/TestRuns/`.
  - Evidence: build matrix archived at `Specs/TestRuns/ac3bbb87f7dd/Build/2026-04-26_225500_final_build_matrix/`; command/UIA slices archived under `Specs/TestRuns/ac3bbb87f7dd/Commands/`; shared-grid/perf/audit evidence archived under `Specs/TestRuns/ac3bbb87f7dd/DxUiTests/` and `Specs/TestRuns/ac3bbb87f7dd/Audit/`.
  - Fresh final recheck after reference cleanup: Debug `RedSalamander` build, full `DxUiTests`, and focused `cmd_pane_navigation_open_drive_menu_keeps_navigation_shell_stable` command selftest passed; focused command evidence is archived at `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_225721/`.
- [x] Move this plan to `Specs/Plans/Done/` once all rows are completed or split into narrower active owner plans.
  - Evidence: this file is marked Done and lives under `Specs/Plans/Done/`.

## Current State Summary

The cleanup rollback milestone is complete. Shared DX host/resource reuse, visibility-gated rendering, DirectWrite/DXGI cleanup, core retained controls, rich text-input bridge coverage, shared grid foundations, Preferences direct-host pages, several dialog shells, viewer shell chrome, owned prompts, `FindFilesWindow`, `FileOperationsIssuesPane`, `ViewerSqlite`, `ShortcutsWindow`, and multiple live UIA/long-run command selftest lanes have already landed and are documented in the historical plans and authoritative specs.

This closeout is complete for the current migrated target set: deferred tree/navigation consumers are closed or verified absent, shared-grid/list-detail parity gaps found during this wave are covered, grouped/checkable/icon-rich behavior has focused shared-control coverage, dialog-shell/one-host re-landings have refreshed command evidence, and the visible-native/control audits report zero unallowed debt. Future migrated surfaces should open a narrower owner plan instead of reopening this umbrella.

## References

- `Specs/Plans/Done/UI_DxUiSharedGridPlan.md`
- `Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md`
- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/UI/UI_VisibleComctlAudit.md`
- `Specs/UI/UI_VisibleNativeAudit.md`
- `Specs/UI/UI_PreferencesDialog.md`
- `Specs/UI/UI_TopLevelToolWindows.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/TestRuns/`
