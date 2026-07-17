# Preferences Dialog Specification

Last updated: 2026-06-05

## Purpose

This document is the authoritative UX and integration contract for the live Preferences dialog.

Implementation history and phased migration notes live in `Specs/UI/UI_PreferencesDialog_MigrationHistory.md`.
Shared `DxUi` control behavior, accessibility, visible-native retirement, and migrated-window acceptance rules live in `Specs/UI/UI_DxUiSharedGrid.md`.

## Scope

This specification applies to:

- `RedSalamander/Preferences.cpp`
- `RedSalamander/Preferences.Dialog.cpp`
- `RedSalamander/Preferences.Internal.cpp`
- `RedSalamander/Preferences.General.cpp`
- `RedSalamander/Preferences.Panes.cpp`
- `RedSalamander/Preferences.FileActions.cpp`
- `RedSalamander/Preferences.Viewers.cpp`
- `RedSalamander/Preferences.Editors.cpp`
- `RedSalamander/Preferences.Keyboard.cpp`
- `RedSalamander/Preferences.Mouse.cpp`
- `RedSalamander/Preferences.Themes.cpp`
- `RedSalamander/Preferences.Plugins.cpp`
- `RedSalamander/Preferences.FileOperations.cpp`
- `RedSalamander/Preferences.CompareDirectories.cpp`
- `RedSalamander/Preferences.HotPaths.cpp`
- `RedSalamander/Preferences.Monitor.cpp`
- `RedSalamander/Preferences.Advanced.cpp`

Related specs:

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/UI/UI_TopLevelToolWindows.md`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/Core/Core_CompareDirectories.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/UI/UI_ManagePluginsDialog.md`

## Window Contract

- Preferences MUST open as a modeless, independent top-level tool window following `Specs/UI/UI_TopLevelToolWindows.md`.
- Preferences MUST apply the persisted `ui.windowBackdrop` setting through the shared window chrome/backdrop helper path with tool-window target semantics.
- Preferences MUST be single-instance within the app process; re-opening the command reuses and activates the existing window.
- The visible shell consists of:
  - a left navigation tree,
  - a right scrollable page host that owns the page title, page description, and page body,
  - `OK`, `Cancel`, and `Apply` command buttons.
- `Apply` MUST remain disabled when there are no pending changes.
- `OK` MUST validate, persist, and apply the current working settings, then close.
- `Cancel`, window close, and `Esc` MUST close directly when there are no pending changes. When pending changes exist, they MUST first prompt to save before closing: `Yes` saves and closes, `No` discards and closes, and `Cancel` keeps editing.
- `Apply` MUST persist and apply the current working settings while leaving the dialog open.

## Navigation Contract

Root item order is:

1. General
2. Panes
3. Viewers
4. Editors
5. User Menu
6. Keyboard
7. Mouse
8. Themes
9. Plugins
10. File Operations
11. Compare Directories
12. Hot Paths
13. Monitor
14. Advanced

Additional navigation rules:

- `Plugins` is the only root node with children.
- Selecting `Plugins` shows the plugin list page.
- Selecting a plugin child node shows the schema-driven subpage for that plugin.
- The visible root row order is the navigation contract. It MUST NOT be inferred from `PrefCategory` enum values because the enum order may differ from the displayed tree order.
- Category switches MUST atomically keep the tree selection, shell title, shell description, and visible page body synchronized.
- A category switch initiated while the native category tree owns keyboard focus MUST keep focus on that tree after the new page is installed. Deferred page-host focus restoration may repair focus that temporarily falls into the replaced page surface, but MUST NOT steal focus from the tree, footer commands, or another external control.
- Page-text refresh MUST invalidate the dialog shell and active page surfaces without forcing an unrelated category-tree repaint. Rapid category changes are accepted only when the existing strict render ceiling remains satisfied.
- Returning to a previously visited page MUST restore that page's retained UI state that is part of the user contract, including persisted search/filter and selection state where defined for the page.
- Page-host scroll position is retained per root/plugin page, not globally. Switching away from a scrolled page MUST NOT leak that page's scroll offset or scroll extent into the next page; the destination page enters at its own retained offset, clamped to its own measured content range, and shows a vertical scrollbar only when that destination page's current content overflows the page host.
- Routed page-host mouse-wheel scrolling is screen-hit-tested. Wheel input outside the dialog MUST be ignored; wheel input over the page host scrolls the page host when the hit target is the host or when the hit DxUi child does not handle the wheel itself. Nested scrollable controls that handle the wheel consume it before the page-host fallback runs.
- The page-host `DxUi::ScrollPanel` is the single interactive owner for right-pane page scroll. It MUST span the same top-to-bottom content band as the category tree, own the page title, description, and body, and keep `OK`, `Cancel`, and `Apply` fixed outside the scroll viewport. The page host MUST NOT show or track a competing native `WS_VSCROLL` thumb.
- Programmatic page-scroll commands such as retained-position restore, focus-into-view, and debug/test `WM_VSCROLL` reset MAY update the logical page offset, but they MUST immediately synchronize the visible `DxUi::ScrollPanel` offset so page state, thumb position, and painted content cannot diverge.
- On note-style pages with no page-local focusable controls, such as `Mouse`, baseline Tab traversal from the focused category tree MUST enter the shell commands in visible enabled order (`Reset All`, `OK`, `Cancel`) and then wrap native focus back to the category tree; reverse Tab traversal MUST mirror that order. When native focus wraps back to the category tree, the shell DxUi host retains the last logical shell command target according to the shared DxUi focus-retention contract.

## Current Live Page Contract

The live narrowed direct-host scope includes:

- `General`
- `Panes`
- `Viewers`
- `Editors`
- `UserMenu`
- `Keyboard`
- `Mouse`
- `Themes`
- `Plugins`
- `FileOperations`
- `CompareDirectories`
- `HotPaths`
- `Monitor`
- `Advanced`

Per-page rules:

- `Mouse` is a note-style page.
- `Viewers`, `Editors`, `Keyboard`, `Themes`, and `Plugins` are list/search/detail style pages and MUST preserve their page-local retained state across category round-trips.
- `Viewers` and `Editors` share the file-actions page implementation but remain separate visible categories because their command columns differ.
- Rapid category-switch validation MUST assert the active page's current contract: `Editors` is a file-actions page and legitimately exposes editable ValuePattern descendants, while `Mouse` is a note-style page and MUST NOT expose stale edit/combo/value/toggle descendants from previously active pages.
- `Plugins` root page MUST expose plugin enablement, custom-path management, and navigation into schema-driven child pages.
- The plugin child page MUST embed the schema-driven configuration editor and MAY still offer the dedicated advanced configuration dialog entry point.
- The File Operations page edits host-owned global defaults only: pre-calculation enable/workers, default copy/move speed limit, cross-file-system bridge buffer size, and the `Auto-dismiss Success` toggle.
- The default copy/move speed-limit edit uses the same binary-throughput unit, decimal, alias, rounding, saturation, and formatter round-trip contract as `Specs/UI/UI_FileOperationsPopup.md`. Preferences preserves its existing boundary policy of trimming every C0/control code unit through U+0020; this is intentionally broader than the popup's six-character ASCII whitespace policy.
- The File Operations page MUST expose an `Auto-dismiss Success` toggle that edits `workingSettings.fileOperations.autoDismissSuccess` (bool, default `false`). When enabled, completed successful or canceled file-operation tasks auto-dismiss from the File Operations progress popup instead of staying as result cards. Toggling it MUST mark Preferences dirty, and the dirty-close `No` path MUST discard the unapplied change.
- The File Operations page MUST NOT duplicate plugin-owned concurrency, recycle-bin batching, or search-walker controls; it instead shows a note that those settings live under `Preferences -> Plugins -> File System`.
- The Compare Directories page edits the same persisted defaults described in `Specs/Core/Core_CompareDirectories.md`.
- The Hot Paths page edits the persisted hot-path definitions and their menu-visibility flag.
- The Monitor page edits RedSalamanderMonitor display/filter defaults only. It MUST load from and save to settings app id `RedSalamanderMonitor`, and MUST NOT persist those values under the main `RedSalamander` settings file.
- The Monitor page MUST group the filter preset and the Text, Error, Warning, Info, Perf, and Debug message-type toggles into one card before the settings-file card. The card MUST NOT expose a numeric filter mask edit; custom masks are edited through the message-type toggles. The toggles remain enabled only when the preset is `Custom`.
- The Monitor page MUST include a final hyperlink-style command card that opens the current user's RedSalamanderMonitor settings JSON file, resolved as `Common::Settings::GetSettingsPath(L"RedSalamanderMonitor")`. If the file is missing, invoking the command MAY create it with the current monitor settings before opening it. Invoking it MUST NOT mark Preferences dirty.
- The Advanced page MUST include a bottom hyperlink-style command that opens the current user's main settings JSON file with the shell default editor, resolved as `Common::Settings::GetSettingsPath(appId)`. If the file is missing, invoking the command MAY create it with the current main settings before opening it. Invoking it MUST NOT mark Preferences dirty.

### Viewers And Editors Page Contract

The `Viewers` and `Editors` pages MUST expose the same mental model:

- **Actions** are named things RedSalamander can launch.
- **Associations** are rules that choose which action a command uses for a file type, pattern, default row, and optional computer override.

The `Viewers` page:

- MUST show `Actions` first, then `Associations`.
- The Associations table columns are `Match`, `Computer`, `F3 View`, `Alt+F3 Alternate View`, and `Status`.
- The Actions table columns are `Name`, `Type`, `Applies To`, `Computer`, and `Status`.
- Viewer-plugin actions MUST choose the stored `pluginId` through a non-editable combo that displays viewer plugin names; users MUST NOT type plugin IDs by hand. Persisted IDs that no longer match a discovered viewer plugin may be shown as a missing entry so the setting can still be inspected.
- `F3 View` and `Alt+F3 Alternate View` action selectors MUST list configured viewer actions plus `(none)` where clearing is valid.

The `Editors` page:

- MUST show `Actions` first, then `Associations`.
- The Associations table columns are `Match`, `Computer`, `F4 Edit`, `Ctrl+Shift+F4 Alternate Edit`, `Shift+F4 Edit New`, and `Status`.
- The Actions table columns are `Name`, `Type`, `Applies To`, `Computer`, and `Status`.
- Editor actions are external-program actions and MUST NOT expose a plugin-ID input.
- `F4 Edit`, `Ctrl+Shift+F4 Alternate Edit`, and `Shift+F4 Edit New` action selectors MUST list configured editor actions plus `(none)` where clearing is valid.

Both pages:

- MUST keep padded content margins inside the tab content area so grids and form controls are not flush to the tab edge.
- MUST give the Associations and Actions grids a stable minimum height, then grow those grids to consume spare page-host height before adding page-level overflow, so the fixed edit form remains visible while tall dialogs show more rows instead of leaving unused space below the tab surface.
- MUST render visible side and bottom borders around the active tab content area plus one tab-strip separator line that is interrupted under the active tab, so no line runs directly below the active tab rectangle.
- MUST edit `workingSettings.fileActions` only; they MUST NOT expose the removed root `viewers`/`editors` shape or `extensions.openWithViewerByExtension`.
- MUST use the shared file-action resolver for the visible preview row so the page explains the same priority the command layer uses: computer-specific extension/pattern, global extension/pattern, computer default, global default.
- MUST show the selected test file path, command, resolved action name, and reason in the preview.
- MUST mark Preferences dirty when an association or action changes, and `Apply` / `OK` MUST persist the changed `fileActions` graph without dropping unrelated settings.
- MUST expose the shared association form with `Match kind`, `Match value`, optional `Computer`, command-specific action selectors, and a `Save Association` command. The live UIA names for editing the match text and saving the association are the shared form labels, not older Viewers-only column/button captions.
- Association-table command columns MUST resolve action IDs through the page's configured actions: configured actions show their display names, and missing actions show the localized `Missing: {id}` text. Tests that validate configured action display names MUST seed both the action definitions and the association rows.
- Association-table header reorder/resize and `Ctrl+C` copy MUST follow the current visible column order. For Viewers, moving `F3 View` before `Match` makes copied row text begin with the `F3 View` cell; moving `Computer` before `Match` makes it begin with the computer override cell.
- Association-table header resize MUST be a real pointer resize: the target column visibly widens, the adjacent visible column shifts, and test diagnostics' grid resize-move counter advances for that drag. Search and sort round-trips MUST preserve the resized layout without reporting page-host resize failures.
- Header sort clicks on the Associations and Actions tables MUST cycle ascending, descending, and no-sort order for the clicked model column. Sorting compares the visible cell text case-insensitively, uses the `Match`/name cell as a tie-breaker for non-primary columns, and restores source order when the sort returns to none. Search/filter rebuilds MUST preserve the active sort spec until the user cycles it back to none.
- On the Associations tab, page-local Tab traversal MUST cover the shared form controls in visible order: `Search`, associations grid, `Match kind`, `Match value`, `Computer`, command-specific action selectors, `Test file`, `Save Association`, `Remove`, `Reset Defaults`, the tab header, then wrap to `Search`. For Viewers the command selectors are `F3 View` then `Alt+F3 Alternate View`; for Editors they are `F4 Edit`, `Ctrl+Shift+F4 Alternate Edit`, then `Shift+F4 Edit New`. Reverse Tab traversal MUST mirror the same order.
- MUST show every association row that participates in resolution, including the Default mapping row. Resetting to defaults restores the full default association set including that Default row.
- MUST replace the selected association row when `Save Association` edits that row to a non-duplicate key. If the edited key already exists elsewhere, saving MUST update by key rather than creating duplicate `(match, computer)` rows. With no selected row and no existing key, saving appends a new association.
- MUST keep the debug selected-extension/action state synchronized with the current grid selection. Fresh page creation may select a valid first row; destructive association mutations such as Remove and Reset Defaults MUST clear the association selection after rebinding so stale row details do not survive the mutation.
- MUST route test/debug association-list scrolling through the same DxUi grid wheel path used by the live surface. Negative wheel detents scroll down from the top, update the vertical scroll offset, and participate in the normal grid invalidation/render-count path used by sustained-scroll validation.
- MUST remain DxUi-owned, theme-aware, and accessible through the shared page-host and grid patterns.

### General Page Contract

The `General` page MUST expose three visible groups in this order:

1. `Display`
2. `DxUI`
3. `Startup`

The `Display` group currently contains:

- `Menu bar`
- `Function bar`
- `Language`

The `Language` control MUST expose `System Language`, embedded English, and any discovered satellite cultures. `System Language` persists `ui.language = "system"` and follows the Windows preferred UI language chain at runtime. Embedded English persists `ui.language = "en"` and uses the executable/plugin embedded resources. Satellite culture entries persist their BCP 47 culture tag, such as `fr-FR`.

The `DxUI` group currently contains:

- `Compact mode`
- `Animations`
- `Window backdrop`

Normative behavior:

- `Compact mode` is a live previewed app-wide density preference. Missing `ui.compactMode` settings default to enabled; an explicit `false` value opts the app back into standard density. Toggling it inside Preferences MUST immediately restyle the visible Preferences page host, then become the persisted value only on `Apply` or `OK`.
- `Animations` MUST expose `System`, `On`, and `Off`. `System` follows the OS preference, `On` forces full DxUI motion, and `Off` forces reduced motion while still editing `ui.reducedMotion`.
- `Window backdrop` MUST expose `Default`, `None`, `Mica`, `Mica Alt`, and `Acrylic`.
- Changing `Language` inside Preferences MUST update only `workingSettings` until `Apply` or `OK`; choosing `No` from the dirty-close prompt MUST discard the unapplied language edit.
- Changing `Window backdrop` inside Preferences MUST immediately preview the selected backdrop policy on the Preferences window itself and on DxUI popup/menu materials owned by the dialog.
- Choosing `No` from the dirty-close prompt MUST discard any unapplied `Compact mode`, `Animations`, or `Window backdrop` preview changes and restore the previously persisted runtime state.
- `Apply` / `OK` MUST persist the new `ui.*` settings, re-apply the current localization preference through the shared settings pipeline, refresh supported top-level windows and app-owned captioned utility/dialog windows through the shared settings/backdrop pipeline, and keep the General page controls synchronized with the committed state.
- Page-local Tab traversal MUST visit focusable General controls in visible order: `Menu bar`, `Function bar`, `Language`, `Compact mode`, `Animations`, `Window backdrop`, `Splash screen`, then wrap to `Menu bar`. Reverse Tab traversal MUST mirror that same order.
- Main folder window backdrop behavior is not part of the Preferences dialog acceptance contract; Preferences owns the setting UI and the supported tool/dialog refresh pipeline.

## Settings And Schema Contract

- All visible first-party settings surfaced in Preferences MUST have stable `title` and `description` metadata in `Specs/SettingsStore.schema.json`.
- The dialog edits `workingSettings` and only commits those values through `OK` or `Apply`.
- After persistence, the host MUST notify the running app so live settings-dependent surfaces update coherently.
- The `General -> Display` controls edit `ui.language` in `workingSettings`.
- The `General -> DxUI` controls edit `ui.compactMode`, `ui.reducedMotion`, and `ui.windowBackdrop` in `workingSettings`.
- The Monitor page controls edit `workingMonitorSettings` and persist through the `RedSalamanderMonitor` settings/schema path.
- Main and Monitor settings are separate conflict-aware documents, not one atomic transaction. `Apply` / `OK`
  saves the main document first. If that succeeds but the Monitor save fails, Preferences MUST apply and advance
  the baseline for the committed main settings, MUST leave the Monitor baseline unchanged so those edits remain
  dirty, and MUST show a localized message that identifies the partial success and failed Monitor path.
- A Monitor-only revision conflict MAY be rebased once by loading the newest Monitor document and replacing its
  entire Monitor section with `workingMonitorSettings`, because Preferences owns that whole section. A second
  conflict or any other Monitor save failure remains pending; it MUST NOT roll back or misreport the already
  committed main document.
- The Advanced and Monitor settings-file links are command affordances, not persisted settings, and MUST NOT require schema metadata.
- Plugin child pages use plugin-provided configuration schema and current configuration payload as the source of truth for rendered fields.
- Plugin fields marked `x-ui-hidden: true` MUST remain JSON-only advanced settings and MUST NOT be rendered in the embedded editor.

## DXUI And Accessibility Contract

- The Preferences visible shell and current live page set MUST use the shared `DxUi` path and obey `Specs/UI/UI_DxUiSharedGrid.md`.
- Shared `DxUi` rules such as zero accepted visible native fallback, page-host ownership, redraw batching, retained-state rules, `WM_GETOBJECT`, UI Automation exposure, and direct-host validation are normative through that shared spec and are not duplicated here.
- The active page surface, not the outer dialog HWND, is the page-local accessibility target for page-specific validation.
- Preferences debug snapshots expose native category-tree focus separately from retained DxUi host focus targets. Tests MUST NOT interpret a retained shell focus target as active native shell focus when `categoryTreeFocused` is true.
- General, Panes, Viewers, and Editors page layout MUST use the Preferences-owned typography context and DirectWrite measurement for visible toggle/combo/card/hint text, not pane-local `HFONT` or GDI text measurement. Tests MUST keep `generalUsesDxUiTypographyContext`, `generalUsesDxUiTypographyMetrics`, `panesUsesDxUiTypographyContext`, `panesUsesDxUiTypographyMetrics`, `viewersUsesDxUiTypographyContext`, and `viewersUsesDxUiTypographyMetrics` true for the matching page snapshots.

## Verification Requirements

Before changing Preferences behavior, the affected work MUST keep these contracts green:

- category navigation stays synchronized with page title, page description, and active page content,
- page scrolling works with mouse wheel, scrollbar thumb drag, and track clicks,
- the right page host begins at the category-tree top edge and scrolls title, description, and body together while keeping footer buttons fixed,
- `OK`, `Cancel`, and `Apply` keep their expected persistence semantics,
- the current live page set preserves page-local retained state expected by the product contract,
- the File Operations page keeps live UI Automation access to its visible combo/edit controls and the dirty-close `No` path discards unapplied File Operations page edits,
- the active page surface exposes the required UI Automation patterns for its visible controls,
- the live DX path does not regress to accepted visible native fallback.

## Non-Goals

- This document does not carry phased migration backlog, resume notes, or refactor TODOs.
- Those items belong in the WIP plans and migration history documents, not in the normative contract.
