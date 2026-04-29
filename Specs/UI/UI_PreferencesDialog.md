# Preferences Dialog Specification

Last updated: 2026-04-01

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
- `RedSalamander/Preferences.Viewers.cpp`
- `RedSalamander/Preferences.Editors.cpp`
- `RedSalamander/Preferences.Keyboard.cpp`
- `RedSalamander/Preferences.Mouse.cpp`
- `RedSalamander/Preferences.Themes.cpp`
- `RedSalamander/Preferences.Plugins.cpp`
- `RedSalamander/Preferences.FileOperations.cpp`
- `RedSalamander/Preferences.CompareDirectories.cpp`
- `RedSalamander/Preferences.HotPaths.cpp`
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
  - a right scrollable page host,
  - shell title plus description,
  - `OK`, `Cancel`, and `Apply` command buttons.
- `Apply` MUST remain disabled when there are no pending changes.
- `OK` MUST validate, persist, and apply the current working settings, then close.
- `Cancel` MUST discard unapplied edits and close.
- `Apply` MUST persist and apply the current working settings while leaving the dialog open.

## Navigation Contract

Root item order is:

1. General
2. Panes
3. Viewers
4. Editors
5. Keyboard
6. Mouse
7. Themes
8. Plugins
9. File Operations
10. Compare Directories
11. Hot Paths
12. Advanced

Additional navigation rules:

- `Plugins` is the only root node with children.
- Selecting `Plugins` shows the plugin list page.
- Selecting a plugin child node shows the schema-driven subpage for that plugin.
- Category switches MUST atomically keep the tree selection, shell title, shell description, and visible page body synchronized.
- Returning to a previously visited page MUST restore that page's retained UI state that is part of the user contract, including persisted search/filter and selection state where defined for the page.

## Current Live Page Contract

The live narrowed direct-host scope includes:

- `General`
- `Panes`
- `Viewers`
- `Editors`
- `Keyboard`
- `Mouse`
- `Themes`
- `Plugins`
- `FileOperations`
- `CompareDirectories`
- `HotPaths`
- `Advanced`

Per-page rules:

- `Editors` and `Mouse` are note-style pages.
- `Viewers`, `Keyboard`, `Themes`, and `Plugins` are list/search/detail style pages and MUST preserve their page-local retained state across category round-trips.
- `Plugins` root page MUST expose plugin enablement, custom-path management, and navigation into schema-driven child pages.
- The plugin child page MUST embed the schema-driven configuration editor and MAY still offer the dedicated advanced configuration dialog entry point.
- The File Operations page edits host-owned global defaults only: pre-calculation enable/workers, default copy/move speed limit, and cross-file-system bridge buffer size.
- The File Operations page MUST NOT duplicate plugin-owned concurrency, recycle-bin batching, or search-walker controls; it instead shows a note that those settings live under `Preferences -> Plugins -> File System`.
- The Compare Directories page edits the same persisted defaults described in `Specs/Core/Core_CompareDirectories.md`.
- The Hot Paths page edits the persisted hot-path definitions and their menu-visibility flag.

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

- `Compact mode` is a live previewed app-wide density preference. Toggling it inside Preferences MUST immediately restyle the visible Preferences page host, then become the persisted default only on `Apply` or `OK`.
- `Animations` MUST expose `System`, `On`, and `Off`. `System` follows the OS preference, `On` forces full DxUI motion, and `Off` forces reduced motion while still editing `ui.reducedMotion`.
- `Window backdrop` MUST expose `Default`, `None`, `Mica`, `Mica Alt`, and `Acrylic`.
- Changing `Language` inside Preferences MUST update only `workingSettings` until `Apply` or `OK`; `Cancel` MUST discard the unapplied language edit.
- Changing `Window backdrop` inside Preferences MUST immediately preview the selected backdrop policy on the Preferences window itself and on DxUI popup/menu materials owned by the dialog.
- `Cancel` MUST discard any unapplied `Compact mode`, `Animations`, or `Window backdrop` preview changes and restore the previously persisted runtime state.
- `Apply` / `OK` MUST persist the new `ui.*` settings, re-apply the current localization preference through the shared settings pipeline, refresh supported top-level windows and app-owned captioned utility/dialog windows through the shared settings/backdrop pipeline, and keep the General page controls synchronized with the committed state.
- Main folder window backdrop behavior is not part of the Preferences dialog acceptance contract; Preferences owns the setting UI and the supported tool/dialog refresh pipeline.

## Settings And Schema Contract

- All visible first-party settings surfaced in Preferences MUST have stable `title` and `description` metadata in `Specs/SettingsStore.schema.json`.
- The dialog edits `workingSettings` and only commits those values through `OK` or `Apply`.
- After persistence, the host MUST notify the running app so live settings-dependent surfaces update coherently.
- The `General -> Display` controls edit `ui.language` in `workingSettings`.
- The `General -> DxUI` controls edit `ui.compactMode`, `ui.reducedMotion`, and `ui.windowBackdrop` in `workingSettings`.
- Plugin child pages use plugin-provided configuration schema and current configuration payload as the source of truth for rendered fields.
- Plugin fields marked `x-ui-hidden: true` MUST remain JSON-only advanced settings and MUST NOT be rendered in the embedded editor.

## DXUI And Accessibility Contract

- The Preferences visible shell and current live page set MUST use the shared `DxUi` path and obey `Specs/UI/UI_DxUiSharedGrid.md`.
- Shared `DxUi` rules such as zero accepted visible native fallback, page-host ownership, redraw batching, retained-state rules, `WM_GETOBJECT`, UI Automation exposure, and direct-host validation are normative through that shared spec and are not duplicated here.
- The active page surface, not the outer dialog HWND, is the page-local accessibility target for page-specific validation.
- General, Panes, and Viewers page layout MUST use the Preferences-owned typography context and DirectWrite measurement for visible toggle/combo/card/hint text, not pane-local `HFONT` or GDI text measurement. Tests MUST keep `generalUsesDxUiTypographyContext`, `generalUsesDxUiTypographyMetrics`, `panesUsesDxUiTypographyContext`, `panesUsesDxUiTypographyMetrics`, `viewersUsesDxUiTypographyContext`, and `viewersUsesDxUiTypographyMetrics` true for the matching page snapshots.

## Verification Requirements

Before changing Preferences behavior, the affected work MUST keep these contracts green:

- category navigation stays synchronized with shell title, description, and active page content,
- page scrolling works with mouse wheel, scrollbar thumb drag, and track clicks,
- `OK`, `Cancel`, and `Apply` keep their expected persistence semantics,
- the current live page set preserves page-local retained state expected by the product contract,
- the File Operations page keeps live UI Automation access to its visible combo/edit controls and `Cancel` discards unapplied File Operations page edits,
- the active page surface exposes the required UI Automation patterns for its visible controls,
- the live DX path does not regress to accepted visible native fallback.

## Non-Goals

- This document does not carry phased migration backlog, resume notes, or refactor TODOs.
- Those items belong in the WIP plans and migration history documents, not in the normative contract.
