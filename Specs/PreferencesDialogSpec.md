# Preferences Dialog Specification (RedSalamander.exe)

## Overview

RedSalamander needs a modern, themed **Preferences** dialog that:
- Uses the same theme system and control rendering as the plugin manager dialogs.
- Provides a left navigation **tree** (categories + plugin subpages) and a right content pane (editors).
- Supports **OK / Cancel / Apply** semantics.
- Uses the Settings Store JSON Schema to drive display text (titles/descriptions) and enable future UI generation with minimal duplication.

This document is both a UX spec and an implementation plan (phased) to refactor the former shortcuts-only Preferences dialog (now removed) into a multi-page preferences experience.

## Goals

- Themed UI consistent with `ManagePluginsDialog` (buttons, toggles, edit boxes, focus rendering).
- Left navigation tree of categories with right-side content panel.
- Page-level validation and safe apply (best-effort; never corrupt settings).
- Minimal code duplication: share “themed controls” code between Preferences and plugin manager dialogs.
- Schema-driven display metadata: every setting shown in the UI has a `title` and `description` available from the aggregated schema written next to the settings file.
- Modeless, **owned** dialog: keep the main app usable while Preferences is open, and keep Preferences above the main app in Z-order (not a separate taskbar/Alt-Tab window; not globally topmost).
- Use `Apply` for iterative testing while keeping the dialog open.
- Rainbow-mode parity: selection highlight uses `RainbowMenuSelectionColor` where appropriate (navigation tree uses rainbow only for selection, not for banding).
- Control polish parity: combobox/edit visuals and list headers/row banding should match the “modern” look of the plugin manager and `ShortcutsWindow` (no classic `WS_EX_CLIENTEDGE` borders in themed mode).
- Disabled-state parity: framed edits/comboboxes use theme-consistent disabled fill + text (no “classic white” disabled backgrounds in dark/rainbow themes).

## Current Status (Implemented)

- Host: `IDD_PREFERENCES` is modeless, owned, resizable, and uses OK/Cancel/Apply semantics.
- Navigation: TreeView with Plugins child nodes; selecting a plugin opens its settings subpage.
- Scroll host: `RedSalamanderPrefsPageHost` provides vertical scrolling, wheel routing, and reduced resize artifacts.
  - The active pane window is resized to the full content height (>= view height) so child controls are never clipped while scrolling.
  - Scrolling is applied by shifting the pane window(s) (`PrefsPaneHost::ApplyScrollDelta`) and forcing an immediate redraw to avoid “blank until hover” and stale owner-draw artifacts.
- Plugins: list management + custom paths + per-plugin subpage (embedded configuration editor) + `Test` / `Test All`.
  - Plugins list: `Configure...` navigates to the selected plugin child node in the tree.
  - Plugin subpage: displays schema-driven controls for the plugin configuration (falls back to read-only JSON preview + error message if the schema is unavailable or has no `fields`).
  - `Configure...` (on the plugin subpage) opens the existing configuration editor dialog for advanced editing.
- Disabled options: when an option is disabled, its title/description text is dimmed (but readable) to match the disabled-state UX.
- Input polish:
  - Framed edits/comboboxes use sibling `STATIC` “input frames” (rounded fill + neutral border).
- Single-line edits are implemented as multiline edits and vertically centered via `ThemedControls::CenterEditTextVertically()` (shared with other dialogs).
  - Also used by Create Directory / rename edits (so centering behavior is consistent across the app).
  - Themed comboboxes use `ThemedControls::CreateModernComboBox()` (custom control) to eliminate native borders and ensure stable first paint; high-contrast falls back to the native ComboBox.
  - Focus cue: a thin accent underline is drawn when the input has focus (mouse or keyboard); focus rectangles are reserved for keyboard navigation on toggles.
- **Schema-Driven UI Generation** (NEW): Main settings now support automatic UI generation from `SettingsStore.schema.json`:
  - Settings annotated with `x-ui-pane`, `x-ui-control`, `x-ui-order`, and `x-ui-section` attributes automatically generate UI controls.
  - Hybrid approach: complex panes (Keyboard, Themes, Viewers, Panes) marked as `x-ui-control: "custom"` preserve handcrafted UX.
  - Simple panes (General, Advanced) use hybrid: existing handcrafted controls + auto-generated controls from schema.
  - Implementation: `SettingsSchemaParser.h/.cpp` parses schema recursively, `PrefsUi::CreateSchemaControl()` generates controls based on `controlType`.
  - Supported control types: `toggle` (boolean), `edit` (string), `number` (integer), `combo` (enum), `custom` (handcrafted).
  - Adding new simple settings to schema with x-ui-* attributes automatically adds them to the appropriate pane without C++ changes.

## Current Bugs / Follow-ups (WIP)

This section is intentionally kept up-to-date so the next prompt can pick up work without re-auditing the codebase.

### Input Controls (Edit / ComboBox)

User-reported issues:
- ComboBox field shows an unwanted inner/system border (“double border” inside the custom input frame).
- ComboBox text sometimes appears only after hover/resizing (first paint invalidation issues).
- ComboBox dropdown lacked hover (hot) tracking, items were too small, and the selected-row accent indicator was too thin.
- Clicking an item in the dropdown could fail to update what the field shows after close.
- Keyboard Up/Down could land on an “empty” selection when no current selection existed.
- When the combobox has focus, pressing “non-selection” keys (example: Ctrl) or pressing Up/Down at the ends could leave the field visually blank (only the frame accent bar remained) until hover/resizing.
- Edit text is not vertically centered in some rows.

Status:
- Addressed by the current implementation (`CreateModernComboBox()` + `CenterEditTextVertically()` + dropdown hot tracking/commit fixes), plus:
  - Frame paint safety: the sibling input frame windows are created with `WS_CLIPSIBLINGS` so their repaint cannot overwrite the input control’s pixels.
  - ModernCombo paint safety: selection syncing always invalidates the combobox window (even when the dropdown list is closed).

Implementation notes (current approach):
- Inputs are wrapped by a sibling `STATIC` “frame” window created by `PrefsInput::CreateFramedEditBox` / `PrefsInput::CreateFramedComboBox` in `RedSalamander/Preferences.Internal.cpp`.
- The frame paints themed fill/border (rounded) via `PrefsInputFrameSubclassProc` (framed edits also get an accent underline in the frame).
- Because the frame and the input are **siblings** (not parent/child), the frame must use `WS_CLIPSIBLINGS`; otherwise a frame repaint can draw over the input and make it appear “blank” until the input repaints for some other reason.
- Single-line edits are implemented as multiline edits and vertically centered via `ThemedControls::CenterEditTextVertically()` (which uses `EM_SETRECTNP` + font metrics). This helper is shared with other dialogs so all “single-line multiline edits” behave consistently.
- Combos in themed mode use the custom `RedSalamanderModernComboBox` control created by `ThemedControls::CreateModernComboBox()`:
  - Draws the selection field (no native border) and owns a themed popup list for the dropdown.
  - Implements the small CB_* message subset used by Preferences (`CB_RESETCONTENT`, `CB_ADDSTRING`, `CB_SETCURSEL`, `CB_GETCURSEL`, `CB_GETCOUNT`, `CB_SET/GETITEMDATA`, `CB_GETLBTEXT(LEN)`, `CB_SHOWDROPDOWN`, `CB_GETDROPPEDSTATE`).
  - Dropdown list UX: selection tracks hover (single current highlight) and the closed field previews the hovered row; Esc restores the opened selection; Enter/click accepts and closes; clicking outside closes and accepts the current selection.
  - Non-selectable rows: blank/whitespace-only items act as separators and are skipped by keyboard navigation; clicking them is ignored.
  - Keyboard support: Up/Down moves within bounds (non-circular), Alt+Down/F4 opens, Home/End/PageUp/PageDown jump, type-to-select, Esc cancels, Enter accepts.
  - Mouse support: button-down shows pressed feedback; dropdown toggles on button-up; hover updates the current selection; click accepts; wheel scroll routed to the popup list while open.
  - Popup positioning: the dropdown is positioned so the currently selected row is roughly aligned with the combo field, and then clamped to the monitor work area so it stays fully visible (no “off-screen” dropdown).
  - Popup sizing: the dropdown is slightly wider than the closed field (≈ 2–3 DIPs extra padding on each side) so it matches modern Windows spacing.
  - Dialog integration: when the dropdown is open, the control claims Enter/Esc via `WM_GETDLGCODE` so the dialog manager does not trigger the default/cancel buttons.
  - Keyboard UX: when the combobox has focus, Enter opens the dropdown (so it can be operated entirely from the keyboard without triggering the dialog default button).
  - Notifications: to avoid re-entrant UI rebuilds during dropdown interaction, `CBN_SELCHANGE` is suppressed while the popup is open and emitted once on close if the selection changed.
  - Reused by other dialogs that need framed combos (example: plugin configuration dialogs) so the visual behavior is consistent.
- Focus cues:
  - Framed edits show a thin accent underline when focused.
  - Modern comboboxes show a vertical accent bar on the left for keyboard focus (`FocusViaMouse` not set) and draw hover/pressed feedback on the dropdown button.
  - Toggle focus rectangles are shown only for keyboard navigation (mouse clicks set `FocusViaMouse` to suppress focus rectangles).

If combobox borders or first-paint glitches are still observed:
- Confirm the themed path is using `CreateModernComboBox()` (high contrast intentionally uses the native ComboBox).
- Verify popup close/routing in `RedSalamander/ThemedControls.cpp` (`ModernComboWndProc`, `ModernComboPopupWndProc`, `ModernComboListSubclassProc`).

### Plugins Configuration Subpages

User-reported issues:
- Per-plugin configuration UI sometimes does not render in the dedicated plugin pane (falls back to “No configuration…”).
- Plugins list `Configure...` should navigate to the plugin child node (showing the configuration page) rather than opening a separate dialog.
- Missing/duplicated navigation guidance text (e.g., “Select Plugins…” shown twice) and missing `Test` / `Test All` buttons in some layouts.

Code pointers:
- Plugins UI: `RedSalamander/Preferences.Plugins.cpp`.
- Schema-driven control creation helpers: `RedSalamander/Preferences.Internal.cpp` (`PrefsUi::CreateSchemaControl`, `CreateSchema*`).

### Scrolling / Redraw Artifacts

User-reported issues:
- Mouse wheel and scrollbar interactions can fail in Preferences (including thumb drag / track clicks).
- First paint and scrolling/resizing can leave artifacts (stale text/control visuals until hover).

Code pointers:
- Scroll host + wheel routing: `RedSalamander/Preferences.Dialog.cpp` (`PreferencesPageHostSubclassProc`, `FindWheelTargetFromPoint`, `HandlePageHostMouseWheel`).
- Scroll application: `RedSalamander/Preferences.Internal.cpp` (`PrefsPaneHost::ScrollTo`, `PrefsPaneHost::ApplyScrollDelta`).

## Known Limitations / Future Improvements

- Plugin schemas drive the embedded editor. If a plugin exposes `fields: []`, there is nothing Preferences can render; improving the plugin’s schema is the correct fix.
- The embedded plugin editor currently targets the common “flat object with scalar fields + options” schema used by built-in plugins; richer JSON editing still belongs in the dedicated editor dialog.

## Validation Checklist (next prompt)

Use this as the “what to verify” list after building/running on Windows:

- Scrolling works in every page: mouse wheel (including smooth wheels/touchpads), scroll bar track clicks, and thumb drag (SB_THUMBTRACK).
- No paint glitches: no “blank until hover” controls on first show; no stale artifacts after scrolling/resizing; no overlapped controls after resize.
- Inputs are polished: single-line edits are vertically centered; combobox selection text appears immediately (no hover required); combobox dropdown + main field are themed.
- ComboBox dropdown UX is correct: hover highlight works, click selects and updates the field after close, and keyboard Up/Down never produces a blank/invalid selection.
- Plugins UX is correct:
  - Plugins root shows the plugin list page; plugin child nodes show the plugin subpage.
  - Plugins list `Configure...` navigates to the selected plugin child node (does not open a dialog).
  - Plugin subpage `Configure...` opens the existing advanced configuration dialog.
  - `Test` / `Test All` buttons are present and reachable (including in small window sizes via scrolling).
  - Non-bug: plugins with `fields: []` (example: `builtin/file-system`) will not show an embedded editor UI.
- Disabled text is dim but readable in cards (both title + description).
- Windowing behavior: Preferences is modeless and owned by the main window (stays above its owner, not globally topmost, and does not show as a separate taskbar/Alt-Tab window).

Key code pointers when debugging any of the above:

- Scroll host + wheel routing: `RedSalamander/Preferences.Dialog.cpp` (`PreferencesPageHostSubclassProc`, `FindWheelTargetFromPoint`, `HandlePageHostMouseWheel`).
- Scroll application: `RedSalamander/Preferences.Internal.cpp` (`PrefsPaneHost::ScrollTo`, `PrefsPaneHost::ApplyScrollDelta`).
- Plugin config editor embedding: `RedSalamander/Preferences.Plugins.cpp` (`EnsurePluginsDetailsConfigEditor`, `LayoutPluginsDetailsConfigEditorPanel`, `CommitPluginsDetailsConfigEditor`).
- Combobox repaint helper: `RedSalamander/Preferences.Internal.cpp` (`PrefsUi::InvalidateComboBox`).

## Non-Goals (initial milestones)

- Fully schema-generated UI for all settings in the first iteration.
- Complete mouse/editor plugin systems if the underlying setting model is not yet defined.

## UX Layout

### Window

- Resizable vertically (and optionally horizontally).
- Modeless: does not block the main app while open.
- Owned by the main app window: clicking the main app must not allow it to cover Preferences (Preferences stays above its owner), while still allowing interacting with the main app behind it.
- Layout:
  - Left: navigation tree (`SysTreeView32`) for categories + plugin subpages.
  - Right: content panel (scrollable when needed).
  - Bottom-right: buttons `OK`, `Cancel`, `Apply`.
- Minimum window width must ensure bottom buttons never overlap (even when resized).
- `Apply` is disabled unless there are pending changes.

### Navigation Tree

Root item order:
1. General
2. Panes
3. Viewers
4. Editors
5. Keyboard
6. Mouse
7. Themes
8. Plugins
9. Advanced

Tree structure rules:
- All items above are **root** items.
- Only **Plugins** has children (one per discovered plugin).
  - Selecting **Plugins** (root) shows the Plugins list page.
  - Selecting a **plugin child** shows a plugin settings subpage for that plugin.

Implementation note:
- The control ID remains `IDC_PREFS_CATEGORY_LIST` for compatibility, but the dialog resource uses a `SysTreeView32` control.

Visual behavior:
- No alternating “band” background in the navigation tree (matches Windows Settings).
- In Rainbow mode, only the selected item uses rainbow tinting; unselected items keep the normal background.

### OK / Cancel / Apply behavior

- `OK`: validate → save settings → apply to running app → close dialog.
- `Cancel`: discard pending edits → close dialog (no save, no apply).
- `Apply`: validate → save settings → apply to running app → keep dialog open.

Additional modeless semantics:
- Transient settings (example: Themes “Apply Temporarily”) are previewed immediately without saving.
- The Preferences dialog must also update its own visuals when transient settings are applied (example: theme preview changes the dialog theme).
- If the user presses `Cancel` after applying transient settings, the app and plugins must re-receive baseline settings and revert all transient changes.

Dynamic apply requirement:
- `Apply` (and transient applies) must not only mutate settings on disk; they must cause the running application and loaded plugins to update behavior and visuals immediately (not only at startup).

## Shared Themed Controls (code reuse)

Create a shared, reusable set of themed Win32 controls/utilities used by:
- `ManagePluginsDialog`
- the new Preferences dialog

Target capabilities:
- Themed push buttons (modern look, focus with low-contrast color).
- Themed toggles:
  - Preferences pages use a Windows Settings-style switch: state text (e.g. `On`/`Off`) + switch track/knob on the right, aligned to the row’s right edge.
    - Track uses the theme accent/highlight color for the “first” option (e.g. `On`/`True`/first enum value).
    - Knob is rendered as a dark “dot” for strong contrast (with disabled-state dimming).
  - Plugin manager configuration can keep using a 2-option labeled toggle (e.g. `True [o  ] False`) when it improves clarity for non-boolean choices.
- Themed combobox (for enums / >2 options), consistent with edit boxes:
  - Use a custom rounded “input frame” (like plugin manager) instead of `WS_EX_CLIENTEDGE`.
  - Both the selection field (main part) and the drop-down list must use themed background/text colors, including disabled state.
  - Width: don’t stretch to full row if the option labels are short; compute a compact width from the widest option text and keep dropdown width large enough to show the full labels (`CB_SETDROPPEDWIDTH`).
- Themed edit boxes (rounded border; width varies by value type) using the same “input frame” approach (no classic border in themed mode).
- Scrollable panel helper (vertical scrolling, correct range/position, keyboard navigation support):
  - Mouse wheel works reliably with both “notched” wheels and high-resolution wheels/touchpads (accumulate partial deltas; no lost scroll).
  - Wheel is routed by pointer location inside the dialog (e.g., if the pointer is over a list, that list scrolls; otherwise the page host scrolls), independent of which control currently has focus.
  - Scrollbar thumb dragging and clicking the track work (SB_THUMBTRACK/SB_PAGE*).
- Focus rendering rules consistent across dialogs.
- ListView header theming helper (custom header paint to avoid “classic” unthemed headers in dark/rainbow themes).

### Settings row layout (Windows Settings style)

For row-based pages (General / Panes / Advanced):
- Each setting is rendered as a rounded “card” row:
  - Left: title + description (description can wrap).
  - Right: control (toggle/combobox/edit) right-aligned inside the card.
- Cards use a subtle contrasting surface fill and border (theme-aware) and avoid classic 3D borders.
- Cards must span the available page width and reflow correctly when the dialog is resized and when the page host scrollbar appears/disappears.
- Static text must paint with a background that matches its surface: card surface for text inside cards, window background elsewhere (avoid transparency artifacts).

## Schema Display Metadata

The aggregated schema file written next to settings (`<AppId>.settings.schema.json`) is the source of truth for:
- `title` (display name)
- `description` (help text shown under the control)

### Requirements

- All settings surfaced in Preferences must have `title` and (where meaningful) `description` in schema.
- For complex objects (e.g. `theme` / `shortcuts`), nested properties should also provide `title`/`description`.
- Main application settings that are not user-editable (e.g. `windows` placement) must still be fully specified with clear constraints and documentation.
- Avoid schema fields with unclear meaning; when an option set is known, prefer explicit `enum` + description over a loosely typed `string`/`integer`.

### UI Annotations for Schema-Driven UI Generation (IMPLEMENTED)

Schema-driven UI generation is now implemented using custom `x-ui-*` annotations. Settings annotated with these attributes automatically generate UI controls in the appropriate pane:

**Implemented Annotations**:
- `x-ui-pane`: Target pane name (`"General"`, `"Advanced"`, `"Keyboard"`, `"Themes"`, `"Viewers"`, `"Panes"`, `"Plugins"`)
- `x-ui-order`: Integer ordering within a pane (lower numbers appear first)
- `x-ui-control`: Control type - `"toggle"` (boolean), `"edit"` (string), `"number"` (integer), `"combo"` (enum), `"custom"` (handcrafted)
- `x-ui-section`: Optional section header for grouping related settings

**Example**:
```json
"menuBarVisible": {
  "type": "boolean",
  "default": true,
  "title": "Menu Bar",
  "description": "Show or hide the main menu bar.",
  "x-ui-pane": "General",
  "x-ui-order": 10,
  "x-ui-control": "toggle",
  "x-ui-section": "Display"
}
```

**Hybrid Approach**:
- Simple settings (toggles, text inputs, numbers) → Auto-generated from schema
- Complex UX (keyboard shortcuts with conflict detection, theme creation workflow, viewer mapping) → Mark as `x-ui-control: "custom"` to preserve handcrafted quality
- Transitioning panes → Existing handcrafted controls + new schema-driven controls appended

**Implementation Files**:
- Schema parser: `RedSalamander/SettingsSchemaParser.h/.cpp`
- UI helpers: `RedSalamander/Preferences.Internal.cpp` (`PrefsUi::CreateSchemaControl`, `CreateSchemaToggle`, `CreateSchemaEdit`, `CreateSchemaNumber`)
- Integration: `RedSalamander/Preferences.Dialog.cpp` (loads schema), `RedSalamander/Preferences.General.cpp` (demo)

These annotations are additive and do not break JSON Schema validation.

## Page Specs (right panel)

### General

General application preferences (initial scope):
- Menu visibility settings (menu bar / function bar).
- Cache limits (if appropriate for non-advanced).

### Panes

Two sections:
- Left pane
- Right pane

Editable fields (initial scope):
- Display mode
- Sort by / direction
- Status bar visibility
- History max

### Viewers

Edit extension → viewer plugin mapping:
- Backed by `extensions.openWithViewerByExtension`.
- UI supports add/remove/edit mappings and validates extension format.

### Editors

Edit extension → editor mapping:
- Requires defining the underlying settings model (TBD).
- Initial milestone can be a placeholder page until the editor system is implemented.

### Keyboard (redesign)

Replace the current “two lists + manual editor” with:

Layout proposal:
- Top: filter row
  - Search box: filter by command name/ID and by assigned shortcut text.
  - Scope filter: FunctionBar / FolderView / All
  - Pane filter: Left / Right / Both (when applicable)
- Main: list view grid
  - Columns: Command, Current Shortcut, Scope, Context
- Bottom: actions
  - `Assign…` (captures next key chord)
  - `Remove`
  - `Reset to Defaults`
  - `Import…` / `Export…`

Behavior:
- Selecting a command shows its current binding(s).
- Pressing `Assign…` enters capture mode:
  - Show a live “currently pressed” chord preview.
  - Show potential conflicts inline (conflicted command name).
  - Provide an explicit confirmation button to commit the assignment.
  - Provide a clear cancel path (Esc and/or a button).
- Conflicts are detected and resolved without relying on modal message boxes.
- The list view should match `ShortcutsWindow` visuals:
  - Banded rows.
  - Rainbow-mode selection tinting.
  - Low-contrast focus rendering that respects theme/high-contrast mode.

Data:
- Uses `settings.shortcuts` (and defaults when missing).
- Import/export uses JSON (stable format) stored by user-chosen path.

### Mouse

Requires underlying model (TBD). Initial milestone can be placeholder.

### Themes

Features:
- Theme selector: builtin + custom themes (loaded from themes folder) with display names.
- Builtin themes are read-only, but their effective values should still be displayed (so users can use them as a starting point).
- Color editor control for color keys:
  - Text box accepts `#RRGGBB` / `#AARRGGBB`
  - Optional color picker button
- The color editor shows a live swatch (rounded square + border) for the current color.
- Actions:
  - `Load From File…` (clone into editable working copy)
  - `Apply Temporarily` (preview without saving)
  - `Save Theme…` (writes `*.theme.json5` to themes folder)

New theme workflow:
- The theme selector includes a final entry: `<New Theme>`.
- When `<New Theme>` is selected:
  - Show a “New theme name” edit next to the selector (or in the same row).
  - Allow selecting a Base theme: `None` (treated as `builtin/system`), `System`, `Light`, `Dark`, `Rainbow`, `High Contrast`.
  - Show the effective color list (base values + overrides) and allow editing keys without requiring manual key typing.
  - `Apply Temporarily` previews the new theme even before it is saved to disk.
  - `Save Theme…` prompts for a destination file and writes the theme with `id = user/<slug>` generated from the entered name.

Color list visuals:
- The theme colors list shows a color swatch per row (either a dedicated column or embedded in the value column) with a visible border for rainbow mode.
- When the Base theme changes, all non-overridden (“untouched”) effective values in the list update immediately to match the new base.

### Plugins

Integrated plugin management:
- List installed plugins (type/origin + enable/disable)
- Inspect per-plugin settings via navigation tree child nodes
- Plugin subpage: embedded schema-driven configuration editor (when the plugin exposes a schema); falls back to read-only JSON preview + error message otherwise.
- `Configure...` behavior:
  - Plugins list page: selects the plugin child node (navigates).
  - Plugin subpage: opens the existing configuration editor dialog (reused from `ManagePluginsDialog`) for advanced editing.
- Diagnostics: `Test` (selected plugin) and `Test All`.

Entry point:
- The main menu item `Plugin Manager...` opens Preferences with **Plugins** selected.

### Advanced

Expert options:
- Monitor settings
- Low-level cache knobs
- Future debug/perf flags

## Implementation Plan (Phased)

### Implementation Status (Detailed)

- Shared themed Win32 drawing utilities are extracted into `RedSalamander/ThemedControls.h` and `RedSalamander/ThemedControls.cpp`.
- A new Preferences host dialog resource exists as `IDD_PREFERENCES` (navigation tree + page host + OK/Cancel/Apply).
- The initial Preferences host implementation lives in `RedSalamander/Preferences.Dialog.cpp` and supports Apply via `WndMsg::kSettingsApplied` (public API in `RedSalamander/Preferences.h` + `RedSalamander/Preferences.cpp`).
- Preferences is modeless (`CreateDialogParamW`) and the main message loop routes keyboard navigation via `IsDialogMessageW`.
- Preferences is single-instance: the window handle is owned via `wil::unique_hwnd` and the dialog state is owned via `std::unique_ptr`, avoiding manual cleanup while preventing duplicate Preferences windows.
- Before saving, Preferences requests a live settings snapshot via `WndMsg::kPreferencesRequestSettingsSnapshot` to avoid losing runtime state (window placement, folder history/current paths).
- Phase 1 currently implements the **General** page toggles for `mainMenu.menuBarVisible` and `mainMenu.functionBarVisible`.
- Phase 2 implements the **Keyboard** page with search + scope filtering, list view editing, capture-based assign, remove/reset, and shortcuts import/export.
- Keyboard capture hint reflows dynamically; `MeasureStaticTextHeight(...)` adds padding to avoid truncation (still re-check on DPI/theme changes).
- Phase 3 adds `title`/`description` metadata for the settings currently surfaced (menu visibility + shortcuts) in `Specs/SettingsStore.schema.json` and keeps `Common/Common/SettingsStore.cpp` in sync.
- Phase 4 implements the **Themes** page: theme selection (builtin/file/user), effective value display (base + overrides), editing `user/*` themes (name/base/colors), `<New Theme>` creation, hex color editor + color picker, load/save `*.theme.json5`, and a non-persistent “Apply Temporarily” preview that is reverted on Cancel.
- Themes page: the color editor shows a live swatch (rounded square + border) and the effective colors list shows a per-row swatch column.
- Phase 5 has started with a **Panes** page (display/sort/status bar/history max), a **Viewers** page (extension → viewer mapping), and a **Plugins** page with an embedded list (type/origin + enable/disable + Configure navigation + Test/Test All + custom plugin paths) plus per-plugin configuration subpages.
- Viewers page: includes a search/filter row above the mappings list.
- Themes page: includes a search/filter row above the colors list (filters by key/value).
- Plugins page: includes a search/filter row above the plugins list (filters by name/id/type).
- Pane modules now exist for all categories (some are still placeholders with note-only UI):
  - Implemented pages: `General`, `Panes`, `Keyboard`, `Viewers`, `Themes`, `Plugins`, `Advanced`
  - Placeholder pages: `Editors`, `Mouse`
  - Root pane windows are owned via `wil::unique_hwnd` and are hosted inside the page host.
  - Pane root windows paint the shared background/cards in `WM_PAINT` / `WM_PRINTCLIENT` (to satisfy standard controls that request parent background), and forward control messages back to the page host.
  - Pane root windows forward `WM_COMMAND`/`WM_NOTIFY`/`WM_DRAWITEM`/`WM_MEASUREITEM`/`WM_CTLCOLOR*` back to the page host so the existing host logic continues to handle owner-draw + scrolling + background decisions.
- Shared pane-host plumbing lives in `RedSalamander/Preferences.Internal.h` + `RedSalamander/Preferences.Internal.cpp` (`PrefsPaneHost::*`).
  - Page-host scrolling helpers (shared across panes): `PrefsPaneHost::ScrollTo(...)`, `PrefsPaneHost::EnsureControlVisible(...)`.
- Shared filter helpers for search rows live in `PrefsUi` (`TrimWhitespace`, `ContainsCaseInsensitive`) and are reused by Viewers/Themes/Plugins.
- Shared folder-settings helpers live in `PrefsFolders` (pane slot ids, history max, pane view prefs, and “ensure working” helpers) and are reused by the Panes page + save/dirty logic.
- The page host (`IDC_PREFS_PAGE_HOST`) enables `WS_CLIPCHILDREN` at runtime so it never paints over the active pane window (prevents “blank until hover” artifacts).
- Page-host paint: double-buffered background/cards rendering for any uncovered regions + full invalidation (including children) on scroll/resize/activation to avoid stale pixels and scroll artifacts (no “ghost” remnants after scrolling).
- Scrollbar sizing: scroll range measurement ignores the pane container windows themselves (measures visible child controls + cards) so the vertical scrollbar only appears when content exceeds the view.
- Scroll correctness: the active pane container is resized to the computed content height so scrolled-in controls are inside the pane’s client rect (avoids “blank cards” where only the background renders).
- The per-pane window classes are extracted first; control creation is progressively migrating into per-pane modules:
  - `GeneralPane::CreateControls(...)`, `PanesPane::CreateControls(...)`, `ViewersPane::CreateControls(...)`, `KeyboardPane::CreateControls(...)`, `ThemesPane::CreateControls(...)`, `PluginsPane::CreateControls(...)`, `AdvancedPane::CreateControls(...)` now exist.
  - `AdvancedPane::Refresh(...)` owns the Advanced page state sync (preset selection, mask enablement, toggles).
  - `AdvancedPane::HandleCommand(...)` owns Advanced page edits (preset, mask, toggles) and the page host routes those `WM_COMMAND` events into the pane.
  - `PanesPane::Refresh(...)` owns the Panes page state sync (display/sort/status bar/history) and is invoked when switching to the Panes category.
  - `PanesPane::HandleCommand(...)` owns Panes page edits (combos, toggles, history) and the page host routes those `WM_COMMAND` events into the pane.
  - `ViewersPane::HandleCommand(...)` owns Viewers page edits (search + add/update/remove/reset) and the page host routes those `WM_COMMAND` events into the pane.
  - `KeyboardPane::HandleCommand(...)` owns Keyboard page edits (search/scope + assign/remove/reset + import/export) and the page host routes those `WM_COMMAND` events into the pane.
  - `KeyboardPane::HandleNotify(...)`, `ViewersPane::HandleNotify(...)`, `ThemesPane::HandleNotify(...)` own list focus/selection/infotip behavior (and swallow any legacy `NM_CUSTOMDRAW` paths now that lists are owner-drawn).
  - Layout/command handling is now routed into the per-pane modules for all implemented pages; the host keeps shared page header/scrolling/cards rendering + message dispatch (Editors/Mouse remain placeholders with note-only UI).
- Panes page: 2-option enums (Display mode, Sort direction) are rendered as a 2-state themed toggle in themed mode (combobox fallback in high-contrast mode).
- General, Panes, and Advanced pages now use Windows Settings-style “cards” (rounded rows) with right-aligned controls (themed mode; standard layout in high-contrast).
- Framed inputs: Keyboard search, Viewers search/extension, Themes name/search/key/color, Plugins search, and the Advanced filter mask use `PrefsInput::CreateFramedEditBox(...)` in themed mode (`WS_EX_CLIENTEDGE` only in high-contrast).
- ListView theming is centralized in `ThemedControls::ApplyThemeToListView(...)` and applied consistently to the Keyboard/Viewers/Themes/Plugins lists.
- Preferences ListViews use `WS_EX_CLIENTEDGE` only in high-contrast mode (borderless in themed mode) to match the modern “Settings” look.
- Boolean toggles in Preferences use a Windows Settings-style switch (`ThemedControls::DrawThemedSwitchToggle(...)`): state text + switch, accent track for the “On” state, black knob, and low-contrast focus.
- Navigation list: no banded background (only rainbow selection tinting when enabled).
- Keyboard list: owner-drawn with banded rows + rainbow selection tinting (including inactive selection dimming and low-contrast focus).
- Visual parity: Viewers mapping list + Themes colors list share the same owner-drawn banded/rainbow list style as Keyboard.
- ListView headers are custom-themed via `ThemedControls::EnsureListViewHeaderThemed(...)` (Preferences lists + reuse in other dialogs).
- Shared UI helpers are being consolidated into `PrefsUi` (`RedSalamander/Preferences.Internal.h` + `RedSalamander/Preferences.Internal.cpp`) to shrink `Preferences.Dialog.cpp` and reduce duplication (combo helpers, toggle state, parsing, window text helpers).
- Resize polish: dialog layout uses measured button minimum widths + a non-overlap fallback + `SWP_NOCOPYBITS`; `WM_ERASEBKGND` + `WM_EXITSIZEMOVE` force repaint to prevent “ghost” overlap trails while resizing.

## Refactor Plan (Module Split + Renaming)

This is the next major step: turn the current “V2” implementation into the permanent Preferences code layout and remove legacy/temporary names.

### Goals

- Remove `PreferencesDialog.*` legacy code and resources once the new dialog fully replaces it.
- Remove all temporary naming (`V2`, `Preferences2`, `Prefs2`, `_PREFS2_`, etc.). (Done: resources now use `IDD_PREFERENCES` + `IDC_PREFS_*`.)
- Each Preferences pane becomes its own window + class, living in its own `.h/.cpp`, and reuses shared drawing/subclassing helpers (no duplicated UI plumbing).
- Consistent control visuals across panes:
  - Switch toggle for all boolean / 2-choice fields (On/Off, True/False, etc.).
  - Unified framed edit + framed combobox look.
  - Page title/description use a larger font than normal body text.
  - Any pane with a list includes a search/filter row at the top (like Keyboard).

### Target File Layout (end-state)

- `RedSalamander/Preferences.h` (public API)
- `RedSalamander/Preferences.cpp` (optional glue / helpers, if needed)
- `RedSalamander/Preferences.Dialog.h`
- `RedSalamander/Preferences.Dialog.cpp`
- `RedSalamander/Preferences.Internal.h` / `RedSalamander/Preferences.Internal.cpp` (shared helpers for panes)
- `RedSalamander/Preferences.General.h/.cpp`
- `RedSalamander/Preferences.Panes.h/.cpp`
- `RedSalamander/Preferences.Viewers.h/.cpp`
- `RedSalamander/Preferences.Editors.h/.cpp`
- `RedSalamander/Preferences.Keyboard.h/.cpp`
- `RedSalamander/Preferences.Mouse.h/.cpp`
- `RedSalamander/Preferences.Themes.h/.cpp`
- `RedSalamander/Preferences.Plugins.h/.cpp`
- `RedSalamander/Preferences.Advanced.h/.cpp`

### Incremental Migration Strategy (keep the build working)

1. **Public API + wiring**
   - Introduce `Preferences.h` and move the external API (`ShowPreferencesDialog*`, `GetPreferencesDialogHandle`) to it.
   - Update `RedSalamander.cpp` to include `Preferences.h`.
   - Keep the existing implementation file in place initially (to avoid a huge “all at once” change).

2. **Remove legacy shortcuts-only dialog**
   - Delete `PreferencesDialog.cpp/.h` once no longer referenced.
   - Remove the legacy `IDD_PREFERENCES` dialog resource and related strings/icons if unused.

3. **Dialog host extraction**
   - Move the host dialog proc/state + navigation tree + right host plumbing into `Preferences.Dialog.*`.
   - Move shared helpers (cards, framed input, list banding/header theming, scrolling host, focus rules) into `Preferences.Internal.*`.
   - Leave pane-specific logic in place temporarily but routed through the new modules.

4. **Pane window classes**
   - Introduce a small pane interface (create window, apply theme, load-from-settings, commit-to-settings, handle validation).
   - Migrate panes one-by-one into their own window classes:
     - Start with `General`, then `Keyboard`, then `Panes`, then `Themes`, then `Viewers`, then `Advanced`, then `Plugins`.
     - `Editors` / `Mouse` can remain placeholders until their data model is finalized.
   - Each migrated pane should own its root window via RAII (e.g., `wil::unique_hwnd`) and treat child `HWND` as non-owning if parent-owned.

5. **Search rows for list panes**
   - Add a consistent search/filter row to each pane that has a list control (Viewers, Themes colors list, Plugins list, etc.).
   - Reuse the Keyboard filtering model where possible.

6. **Rename pass (no temporary names)**
   - Done: renamed files and symbols to remove `V2`, `Preferences2`, and `_prefs2` naming.
   - Done: renamed resources: `IDD_PREFERENCES2` → `IDD_PREFERENCES`, `IDC_PREFS2_*` → `IDC_PREFS_*` (and updated all code references).
   - Remove any compatibility wrappers once all call sites are updated.

7. **Validation**
   - Ensure tab navigation works across panes and that `IsDialogMessageW` routing remains correct.
   - Verify Apply/Cancel semantics (including transient theme preview revert).
   - Verify resizing + scroll behavior (cards reflow correctly; no overlap; controls never disappear unexpectedly).

### Phase 0 — Preparation (safe refactor)

- Extract reusable “themed controls” and scroll panel helpers from `ManagePluginsDialog` into a shared module.
- Keep existing `IDD_PREFERENCES` functional during refactor by building a new dialog implementation side-by-side.

### Phase 1 — Preferences host dialog skeleton

- Add a new dialog resource (e.g. `IDD_PREFERENCES`) with:
  - Left navigation tree
  - Right placeholder panel
  - `OK/Cancel/Apply`
- Wire selection changes to swap right-side page windows.
- Add theming + focus rendering to match plugin dialogs.

### Phase 2 — Keyboard page migration

- Implement the redesigned keyboard UI.
- Reuse existing shortcut parsing/formatting and `ShortcutManager` conflict detection.
- Add import/export of shortcuts JSON.

### Phase 3 — Schema metadata enrichment

- Add `title`/`description` to `Specs/SettingsStore.schema.json` for all settings shown.
- Ensure the runtime base schema source stays in sync with `Specs/SettingsStore.schema.json` (see Phase 6).

### Phase 4 — Themes page + color control

- Implement color editor control + theme file operations.
- Support temporary apply and save to file.

### Phase 5 — Remaining pages

- General, Panes, Viewers mapping UI.
- Plugins embedded page (incrementally replacing sub-dialog).
- Advanced.

### Phase 6 — Modeless + schema distribution + refactor follow-ups

- Implement modeless Preferences (`CreateDialogParamW` + message-loop support via `IsDialogMessageW`).
- Ensure transient applies update the Preferences dialog itself and are fully reverted on `Cancel` (theme preview must not clobber unrelated live settings).
- Stop embedding the Settings Store schema as a giant C++ string; ship `Specs/SettingsStore.schema.json` with the app build output (near the exe) and load it at runtime when needed.
- Split `RedSalamander/Preferences.Dialog.cpp` into logical files (host + per-pane modules/classes).
- Implemented: Keyboard page split into `RedSalamander/Preferences.Keyboard.h` + `RedSalamander/Preferences.Keyboard.cpp` with shared state in `RedSalamander/Preferences.Internal.h` (`KeyboardPane`).
- Implemented: Viewers page split into `RedSalamander/Preferences.Viewers.h` + `RedSalamander/Preferences.Viewers.cpp` (`ViewersPane`).
- Prepare for localization by moving all user-visible UI strings to `.rc` STRINGTABLE (stable IDs).

## Compatibility / Migration Notes

- The settings file format remains compatible; UI changes must not require schema version bumps unless new persisted fields are introduced.
- If new settings fields are required (Editors/Mouse), update:
  - `Common::Settings::Settings` model
  - `Specs/SettingsStore.schema.json`
  - Load/Save logic in `Common.dll`

## Continuation Notes / Remaining Work

### Current implementation status (so far)

- Implemented categories: `General`, `Panes` (starter), `Viewers` (starter), `Keyboard`, `Themes`, `Plugins` (starter), `Advanced` (Monitor starter).
- Placeholder pages: `Editors`, `Mouse` (note-only UI; data model TBD).
- Refactor status: per-pane modules exist for all categories; most page-specific `WM_COMMAND`/`WM_NOTIFY` handling is routed into pane modules (General/Panes/Viewers/Keyboard/Themes/Plugins/Advanced). Layout is being migrated out of `Preferences.Dialog.cpp` pane-by-pane (`General`/`Panes`/`Plugins`/`Viewers`/`Themes`/`Keyboard`/`Advanced` moved so far).
- ListView notify handling is routed into pane modules (`*Pane::HandleNotify(...)`); the legacy list-specific `WM_NOTIFY` handling in `Preferences.Dialog.cpp` has been removed.
- Remaining UI polish: convert remaining pages to Windows Settings-style “cards” with right-aligned controls (notably Viewers/Themes list + filter layouts) and keep input frames visually consistent on card surfaces.

### Where the implementation lives

- New dialog implementation: `RedSalamander/Preferences.Dialog.cpp` (modeless host; `IDD_PREFERENCES`) with public API in `RedSalamander/Preferences.h` + `RedSalamander/Preferences.cpp`.
- Lifetime / ownership: single-instance via `g_preferencesDialog` (`wil::unique_hwnd`) + `g_preferencesState` (`std::unique_ptr<PreferencesDialogHost>`). The state is heap-owned because raw pointers are stored in dialog/subclass userdata.
- RAII policy: all owning Win32 resources use WIL RAII wrappers (`wil::unique_*`); HWNDs stored in `PreferencesDialogState` are non-owning unless explicitly wrapped (pane/config panels use `wil::unique_hwnd`).
- State model: `baselineSettings` (last applied) vs `workingSettings` (edited); `dirty` gates `Apply`.
- Apply flow (current): request live snapshot (`WndMsg::kPreferencesRequestSettingsSnapshot`) → merge edited sections → save settings + write aggregated schema → notify app via `WndMsg::kSettingsApplied`.
- Shared helpers: file I/O for shortcuts import/export and theme load/save is centralized in `RedSalamander/Preferences.Internal.h` + `RedSalamander/Preferences.Internal.cpp` (`PrefsFile::*`).
- Shared list helpers: list row height and two-column themed list drawing lives in `RedSalamander/Preferences.Internal.h` + `RedSalamander/Preferences.Internal.cpp` (`PrefsListView::*`).
- Panes page layout: implemented in `RedSalamander/Preferences.Panes.cpp` (`PanesPane::LayoutControls`); the host delegates layout into the pane.
- Keyboard list rendering: owner-draw measurement/painting is implemented in `RedSalamander/Preferences.Keyboard.cpp` (`KeyboardPane::OnMeasureList` / `KeyboardPane::OnDrawList`); the host only dispatches messages.
- Keyboard page layout: implemented in `RedSalamander/Preferences.Keyboard.cpp` (`KeyboardPane::LayoutControls`); the host delegates layout into the pane.
- Viewers list rendering: owner-draw measurement/painting is implemented in `RedSalamander/Preferences.Viewers.cpp` (`ViewersPane::OnMeasureList` / `ViewersPane::OnDrawList`).
- Viewers page layout: implemented in `RedSalamander/Preferences.Viewers.cpp` (`ViewersPane::LayoutControls`); the host delegates layout into the pane.
- Themes colors list + swatch rendering: owner-draw measurement/painting is implemented in `RedSalamander/Preferences.Themes.cpp` (`ThemesPane::OnMeasureColorsList` / `ThemesPane::OnDrawColorsList` / `ThemesPane::OnDrawColorSwatch`).
- Themes override editor sync: selection → key/value edit updates are implemented in `RedSalamander/Preferences.Themes.cpp` (`ThemesPane::UpdateEditorFromSelection`).
- Themes page: `ThemesPane::Refresh` and `ThemesPane::HandleCommand` now live in `RedSalamander/Preferences.Themes.cpp` (host delegates into the pane module).
- Themes page layout: implemented in `RedSalamander/Preferences.Themes.cpp` (`ThemesPane::LayoutControls`); the host delegates layout into the pane.
- Advanced page layout: implemented in `RedSalamander/Preferences.Advanced.cpp` (`AdvancedPane::LayoutControls`); the host delegates layout into the pane.

### Code Map / Resume Checklist (next prompt)

Use this checklist as a “where to start” guide if you resume work later:

1. **Host vs Pane responsibilities**
   - Host (modeless dialog + layout + apply/cancel/ok): `RedSalamander/Preferences.Dialog.cpp`.
   - Shared helpers/state: `RedSalamander/Preferences.Internal.h` + `RedSalamander/Preferences.Internal.cpp`.
   - Pane modules: `RedSalamander/Preferences.*.h/.cpp` (one per category).
   - If you find logic inside `Preferences.Dialog.cpp` that is clearly page-specific, move it into the corresponding pane module.

2. **WM_NOTIFY routing cleanup**
   - The Preferences host dialog handles navigation-tree notifications (`TVN_SELCHANGED*`, `NM_CUSTOMDRAW`) to switch pages and apply theme selection rendering.
   - Page controls inside the scroll host forward `WM_NOTIFY` through pane handlers (`KeyboardPane::HandleNotify`, `ViewersPane::HandleNotify`, `ThemesPane::HandleNotify`, `PluginsPane::HandleNotify`).
   - If a new page-owned list control is added, implement its notify handling inside the owning pane; the host should only dispatch.

3. **Pane layout migration**
   - Continue moving page-specific layout out of `Preferences.Dialog.cpp` into the owning pane module (`General`/`Panes`/`Plugins`/`Viewers`/`Themes`/`Keyboard`/`Advanced` moved; remaining panes still have host-managed layout).

4. **Naming / resource finalization**
   - Done: renamed resources and IDs to final names (`IDD_PREFERENCES`, `IDC_PREFS_*`).
   - Remove the old `PreferencesDialog.*` and its unused resources.

5. **Schema file distribution**
   - Replace any embedded schema string with runtime file loading of `Specs/SettingsStore.schema.json` copied next to the exe.
   - Ensure settings save also writes the aggregated schema (app + plugin schemas) next to the settings file for user editing.

6. **Plugins integration**
   - Implemented: embedded enable/disable list with search, Apply wiring, and `WndMsg::kPluginsChanged` notifications.
   - Implemented: embedded custom plugin paths list (add/remove).
   - Implemented: configure the selected plugin via the existing configuration dialog UI (stores into `plugins.configurationByPluginId` and applies on Apply).
   - Implemented: Preferences no longer uses the legacy `IDD_PLUGINS_MANAGER` dialog for plugin management.
   - Done: removed the legacy dialog/resources and entry point; only the reusable plugin configuration dialog UI remains active in `RedSalamander/ManagePluginsDialog.*`.

### Modeless conversion notes

- Implemented modeless behavior:
  - Hosted with `CreateDialogParamW` (not `DialogBoxParamW`).
  - Main message loop routes dialog messages using `IsDialogMessageW` for the active Preferences HWND.
  - Single-instance behavior (reuse/activate the existing window if already open).
- Apply/Cancel must handle transient theme previews:
  - `Apply Temporarily` updates app + Preferences visuals.
  - `Cancel` reverts app + Preferences visuals and re-applies baseline settings (including notifying plugins).

### Settings schema distribution (source of truth)

- Source of truth: `Specs/SettingsStore.schema.json` in the repo.
- Build output: copy `Specs/SettingsStore.schema.json` near the exe so runtime code can load it when generating schema exports.
- Runtime: replace the giant embedded `kSettingsStoreSchemaJsonUtf8` with file loading (cache in memory; fail gracefully if missing).

### Remaining pages (Phase 5+)

- **Panes** page:
  - Implemented UI (display/sort/status bar/history max) using themed toggles/combos/edits, rendered as Settings-style cards in themed mode (standard layout in high-contrast), and applied live via `WndMsg::kSettingsApplied`.
  - Schema metadata: `title` / `description` is now present for the `folders` settings used by the Panes page (display/sort/direction/status bar/history max). Keep it in sync as the UI expands.
  - Decide whether to surface layout/active pane in a later milestone.
- **Viewers** page:
  - Implemented basic extension → viewer mapping editor (search + list + add/update/remove + reset to defaults).
  - Schema metadata: `title` / `description` is now present for `extensions.openWithViewerByExtension`; clarify whether an “empty string disables viewing” behavior is supported or disallowed.
- **Editors** page:
  - Same mapping UI as Viewers, but for editor plugins (requires underlying model/schema if not present).
- **Mouse** page:
  - Placeholder until the shortcut/action model is finalized, then redesign similarly to Keyboard.
- **Plugins** page:
  - Implemented embedded plugin management page:
    - List installed plugins (type/origin + enable/disable + Configure... + Test/Test All).
    - Add/remove custom plugin DLL paths.
    - Navigation tree integration:
      - The Plugins root node shows the list page.
      - Each plugin appears as a child node; selecting it shows a plugin settings subpage (configuration preview + Configure...).
        - Preview displays the stored `plugins.configurationByPluginId[pluginId]` payload when present, otherwise it displays the plugin’s current/default configuration from the plugin manager.
      - Plugins list `Configure...` navigates by selecting the corresponding child node.
  - Done: deleted the legacy “Manage Plugins…” dialog; `RedSalamander/ManagePluginsDialog.*` remains as reusable plugin configuration UI.
- **Advanced** page:
  - Implemented: Monitor menu toggles + filter preset + filter mask + per-type filter toggles (custom editor).
  - Implemented: cache limits (Directory Info Cache max bytes/watchers/MRU); remaining: other advanced-only knobs.

### UX + theming polish

- Scrollable page host: implemented basic vertical scrolling (scrollbar + mouse wheel) and "scroll focused control into view".
  - Implemented: fixed header (title/description stays visible while page content scrolls).
  - Implemented: mouse wheel forwarding from framed inputs and toggles so the page scrolls even when a child control has focus.
  - Implemented: mouse wheel targets the hovered pane even if focus is still in the left navigation tree (prevents the tree from "stealing" wheel scroll intended for the page).
  - Implemented: page-host scrollbar is fully interactive (thumb drag, track clicks).
    - `IDC_PREFS_PAGE_HOST` uses a custom window class (`RedSalamanderPrefsPageHost`) registered with `DefWindowProcW` so standard scrollbar non-client behavior works reliably.
    - Preferences then subclasses this host to add custom client painting (cards/background) and custom scrolling behavior.
    - Implementation detail: scrollbar style toggles may trigger spurious `WM_SIZE`; the host suppresses re-entrant layout during internal `WS_VSCROLL` updates and reflows explicitly after the scroll range is finalized.
- Scrollbar rules:
  - Only vertical scrolling is supported for pages; never show a horizontal scrollbar.
  - Implemented: the vertical scrollbar does not reserve width when not needed (`WS_VSCROLL` is toggled on/off; content height measurement ignores the full-height pane container windows).
  - The scrollbar visuals should match the current theme (e.g. `DarkMode_Explorer` in dark mode), including when `WS_VSCROLL` is toggled at runtime (force `WM_THEMECHANGED` when the style changes).
- Implemented: bottom button bar (OK/Cancel/Apply) never overlaps; if the dialog becomes narrow, buttons shrink (down to a safe minimum) and gaps tighten to fit.
- Minimum window size:
  - Min width is computed from the measured button texts and a minimum `list + host` width.
  - Min height is computed from the measured button height + a small minimum `list + host` height so the dialog can be shrunk vertically (content scrolls).
  - Implemented: window chrome includes Minimize; Maximize is customized to expand vertically (work-area height) while preserving the current width.
- Resize redraw: on `WM_SIZE`, the dialog lays out; page host suppresses redraw during layout and then invalidates children to reduce resize artifacts. Also force a final repaint on `WM_EXITSIZEMOVE` and avoid `SWP` copy-bits to prevent “ghost” artifacts with owner-drawn controls.
- Card alignment: ensure all right-side controls (toggles/combos/edits) align to a common right edge within a page (notably Advanced), with a consistent gap from the left text block.
- Localization pass: implemented (Preferences host + pane UI strings moved to `.rc` STRINGTABLE, including file dialog filters and import/export error text).
- Keyboard page follow-ups:
  - Implemented: Conflict resolution UX (show which command currently owns the shortcut; offer clear/swap actions).
  - Improve filtering by pane/scope if needed (left/right/both).
- Themes page follow-ups:
  - Implemented: "Duplicate theme" flow (create `user/*` theme from builtin/file theme).
  - Implemented: Make "New theme name" a first-class workflow (dedicated edit next to selector; ensure `id = user/<slug>` stays in sync with the name when saving).
  - Implemented: Show which keys are explicitly overridden vs inherited (e.g., an "override" indicator column or style).
- Combobox polish:
  - Implemented: `ThemedControls::ApplyThemeToComboBox(...)` applies `DarkMode_Explorer` theming for both the field + dropdown in dark themes; input controls forward `WM_CTLCOLOR*` from combobox internals back to the host for consistent background/text.
  - Implemented: Preferences page host paints on `WM_PRINTCLIENT` without excluding child rects (fixes "white after hover"/blank parent background) and defers combo theming on dropdown open/close (avoids comctl32 re-entrancy crashes).
  - Implemented: dropdown-open theming avoids touching the combobox field (prevents the dropdown from immediately closing on open).
  - Implemented: pane windows handle `WM_PRINTCLIENT` + `WM_PAINT` so child controls using `DrawThemeParentBackgroundEx` receive the correct card/window background and hover artifacts are cleared.
  - Implemented: disabled framed inputs (combo/edit) use a disabled surface fill consistent with the input frame.
  - Implemented: modern comboboxes toggle the dropdown on mouse button-up (with button-down pressed feedback), show button hover feedback, and render a keyboard-focus left accent bar; the dropdown popup uses the same rounded-corner style.
  - Remaining: visually verify combobox hover states and input-frame consistency across all pages.
  - Implemented: cap combobox field widths for "card" rows (dropdown still expands to full label width via `CB_SETDROPPEDWIDTH`).
- ListView polish:
  - Implemented: Themes colors list and Viewers mapping list match Keyboard/ShortcutsWindow banded + rainbow selection visuals.
  - Implemented: Preferences list views are borderless in themed mode (no `WS_EX_CLIENTEDGE`).
- Rainbow-mode support:
  - Implemented: navigation tree uses rainbow for selection only; Keyboard list uses banding + rainbow selection tinting.

### Refactor / rename end-state

- Split `RedSalamander/Preferences.Dialog.cpp` into logical files (host + per-page code).
- Done: removed remaining temporary naming (`IDD_PREFERENCES2`, `IDC_PREFS2_*`, `Prefs2` identifiers) and renamed resources to their final names.
- Rename/refactor `RedSalamander/ManagePluginsDialog.*` into a Plugins page module aligned with the Preferences page architecture.
- RAII/handles:
  - Use WIL RAII wrappers for owned Win32 resources (HBRUSH/HPEN/HFONT/HICON/HDC/etc.).
  - Implemented: eliminate raw owned handles in Preferences state (use WIL wrappers like `wil::unique_any<HIMAGELIST,...>` directly).
  - Store child-control `HWND` values as non-owning references (parent destroys them); avoid manual `DestroyWindow` for dialog-owned children.
  - Prefer `wil::unique_hwnd` for top-level windows created outside dialog ownership.
  - Never patch vcpkg/WIL sources to “work around” compiler/IntelliSense issues; fix in project code (or project settings) instead.

### Future improvements
The Red Salamander preferences system is a sophisticated Windows dialog with 9 category panes, JSON-backed storage, and modern card-based UI. The architecture is modular but suffers from a monolithic dialog file (3,220 lines), repetitive control creation patterns. Key improvements focus on reducing code duplication and enhancing UX.

Plugins nav tree refresh — Rebuild the Plugins subtree when plugins are refreshed (e.g., after adding/removing custom plugin paths) so the navigation tree stays in sync without reopening Preferences.

Steps
Extract control factory utilities — Create a reusable SettingRowBuilder class in Preferences.Utilities.cpp to replace repetitive CreateWindowExW patterns for toggle/combo/edit rows (currently duplicated ~20× across panes).

Refactor PreferencesDialogState — Group the ~100 HWND members into per-pane sub-structs (e.g., GeneralPaneControls, ThemesPaneControls) in PreferencesInternal.h to improve organization and reduce struct size.

Split Preferences.Dialog.cpp — Extract navigation tree handling, scrollable host logic, and theming into separate files (Preferences.NavTree.cpp, Preferences.ScrollHost.cpp) to reduce the 3,220-line god file.

Control factory scope —  SettingRowBuilder support start with with toggle + combo (most common), and expand iteratively.

Unimplemented pane priorities — Which pane should be implemented first? Recommend: Editors pane first (more commonly requested than Mouse settings).

Settings search complexity — Full-text search requires indexing all setting labels/descriptions. Should this be static metadata or runtime-generated? Recommend: Static metadata array with localized strings for searchability.

### Schema follow-ups

- Every new setting surfaced in Preferences must add `title` / `description` in `Specs/SettingsStore.schema.json`.
- Keep the shipped schema file in sync with the repo spec (no duplicated schema sources).
- Optional future: adopt `x-ui` annotations (`category`, `order`, `control`, `advanced`) once pages become more schema-driven.
