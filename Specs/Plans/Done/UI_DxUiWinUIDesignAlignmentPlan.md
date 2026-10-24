# DxUI WinUI/Fluent Design Alignment Plan

**Author:** Ripley (Lead / Reviewer)
**Date:** 2026-04-04
**Closed:** 2026-04-13
**Status:** Complete. Core DxUI phases 0-4 plus the application-wide rollout are implemented and validated on `squad/dxui-winui-design`.
**Scope:** Historical execution checklist, rollout closeout, and validation evidence for the DxUI Fluent-alignment effort
**Authoritative Spec:** `Specs/UI/UI_DxUiWinUIDesign.md`

---

## 0. Checklist

Use this section as the short status view for the plan.

### Done in DxUI on `squad/dxui-winui-design`

- [x] Phase 0 foundation: token/palette/type-ramp updates, focus-ring update, button sizing refresh, and TextField accent focus border
- [x] Phase 1 menu system: `DxContextMenu`, cascading submenus, FolderView context-menu migration, FolderWindow `DxMenuBar`, keyboard/pointer root switching, and matching selftests/perf markers
- [x] Phase 2 primitives: `RadioButton` / `RadioButtons`, `ProgressBar`, shared scrollbar extraction in `DxUi.Scrollbar.cpp`, `Toolbar`, and multi-section `StatusStrip`
- [x] Phase 3 control upgrades: button variants, toggle refresh, checkbox indeterminate, TextField clear button alignment, ComboBox popup shadow, smoke overlay, and CardPanel formalization
- [x] Phase 4 foundations currently in tree: reusable `DrawDropShadow()` plus `WindowHost::SetSystemBackdrop(...)` for DWM Mica/Acrylic hints
- [x] Phase 4 motion work: connected animations and `PageHost` page transition animations
- [x] Acrylic policy decision: use a hybrid model. Keep DWM backdrops for whole windows, keep DxUI menus/flyouts solid or lightly translucent by default, and leave any faux-Acrylic popup treatment as an optional later refinement
- [x] Validation refresh: `DxUiTests.exe`, menu-bar command selftests, and archived perf evidence under `Specs/TestRuns/7d3a1247382a/Commands/2026-04-12_173916` through `2026-04-12_173923`
- [x] Automated visual baseline coverage now exists in `Tests/DxUiTests/Baselines/` and is validated by the focused rendering suite with archived evidence under `Specs/TestRuns/7d3a1247382a/DxUi/2026-04-12_184252`
- [x] Deferred framework controls are now implemented in-tree: `TabControl`, `Slider`, tracking tooltip hide-delay behavior, and inherited RTL / `FlowDirection` support with focused `NewControls`, `Tooltip`, and `Accessibility` coverage
- [x] Oversized menus/flyouts now scroll within the monitor work area, including keyboard visibility management, mouse-wheel scrolling, and archived validation/perf evidence under `Specs/TestRuns/7d3a1247382a/DxUi/2026-04-13_013704`
- [x] Per-monitor DPI handling is now unified for retained DxUI hosts and the remaining hybrid shell windows: `WindowHost` routes `WM_DPICHANGED` and `WM_DPICHANGED_AFTERPARENT` through a shared invalidation path, child/custom windows refresh on parent DPI changes, and `NavigationView` popups survive live DPI transitions; validated by `DxUiTests.exe --suite=WindowHost` plus archived command/selftest evidence under `Specs/TestRuns/7d3a1247382a/Commands/2026-04-13_233341` and `2026-04-13_233702`

### Remaining to implement in DxUI

None in the current DxUI framework checklist after the 2026-04-13 validation refresh.

### Remaining to update across the application

- [x] Replace remaining `ThemedControls` / owner-draw surfaces in `CompareDirectoriesWindow*`, including the Compare Directories menu bar and banner command row; refreshed command/selftest and perf evidence is archived under `Specs/TestRuns/7d3a1247382a/Commands/2026-04-13_065243` and `2026-04-13_065248`
- [x] Migrate `ConnectionManagerDialog`, `ManagePluginsDialog`, and `ConnectionCredentialPromptDialog` further toward DxUI-native controls where practical; refreshed command/selftest and perf evidence is archived under `Specs/TestRuns/7d3a1247382a/Commands/2026-04-13_065250`, `2026-04-13_065301`, `2026-04-13_065316`, `2026-04-13_065319`, `2026-04-13_065345`, `2026-04-13_065348`, `2026-04-13_065349`, `2026-04-13_065412`, and `2026-04-13_065415`
- [x] Migrate remaining `FolderWindow.FileSystem*` dialogs and popups that still use legacy/hybrid themed controls; refreshed command/selftest and perf evidence is archived under `Specs/TestRuns/7d3a1247382a/Commands/2026-04-13_192348` through `2026-04-13_192613`, plus navigation-shell stability guards under `2026-04-13_192048` through `2026-04-13_192132`
- [x] Continue rolling DxUI-native controls through `Preferences.*` and other hybrid surfaces such as `NavigationView`; refreshed command/selftest and perf evidence is archived under `Specs/TestRuns/7d3a1247382a/Commands/2026-04-13_193038` through `2026-04-13_193151`, `2026-04-13_191924`, and `2026-04-13_192746` through `2026-04-13_192832`
- [x] Audit app windows for actual use of the new DxUI `Toolbar`, `StatusStrip`, `ProgressBar`, `RadioButtons`, and `MenuBar` where this spec expects them; runtime validation refreshed under `Specs/TestRuns/7d3a1247382a/Commands/2026-04-13_204158`, `2026-04-13_204209`, and `2026-04-13_204218`, with plugin-window verification from `ViewerSqliteTests.exe TestViewerWindowUsesDxUiHostWithNoVisibleChildControls`
- [x] Migrate `RedSalamanderMonitor` chrome from legacy common-controls `TOOLBARCLASSNAMEW` / `STATUSCLASSNAMEW` to DxUI `Toolbar` / multi-section `StatusStrip`; focused monitor validation/perf evidence is archived under `Specs/TestRuns/7d3a1247382a/Monitor/2026-04-13_212547`
- [x] For each migrated surface, refresh selftests and archive perf evidence before calling the rollout complete; the monitor closeout run is archived under `Specs/TestRuns/7d3a1247382a/Monitor/2026-04-13_212547`, and the rollout-close build refresh is the successful full `.\build.ps1` run from 2026-04-13

None; the application rollout checklist is complete after the 2026-04-13 monitor validation refresh.

### Audit Findings (2026-04-13)

- [x] `MenuBar`: live DxUI menu bars are present in the main app window and Compare Directories, validated by `cmd_app_menuBar_persistent_surface_renders_nonempty_labels` and `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics`
- [x] `StatusStrip`: live DxUI status strips are present and visible in Find Files and ViewerSqlite; the audit refresh tightened the debug snapshots and validation so both surfaces now assert retained `StatusStrip` presence plus usable visible height
- [x] `Toolbar` / Monitor `StatusStrip`: RedSalamander Monitor now uses retained DxUI `Toolbar` and multi-section `StatusStrip` hosts, validated by `RedSalamanderMonitor.exe --chrome-selftest --wait-instance` and archived under `Specs/TestRuns/7d3a1247382a/Monitor/2026-04-13_212547`
- [x] `ProgressBar` and `RadioButtons`: framework implementation and focused `DxUiTests` coverage are complete, but this spec does not currently assign either control to a required product-window rollout target, so there is no additional application audit failure to close for those two controls yet

### Rollout-Close Discoveries (2026-04-13)

- [x] `RedSalamanderMonitor` chrome was the last remaining product-window surface still using legacy toolbar/status-bar common controls
- [x] The monitor needed its own focused `--chrome-selftest` archive path under `Specs/TestRuns/<MachineHash>/Monitor/` because it is a separate executable and does not participate in the main app selftest harness
- [x] Command-backed toolbar toggles need explicit retained-state synchronization after menu-command handling so the DxUI control state stays aligned with the menu checkmarks and underlying view state
- [x] Multi-section product-window status-strip validation benefits from section-level text introspection; `StatusStrip` now exposes section text so focused app-window validation can assert real content instead of just host existence
- [x] `ProgressBar` and `RadioButtons` remain framework-complete with focused `DxUiTests` coverage, but this specification still does not define an additional mandatory product-window rollout target for either control

---

---

## 1. Migration Priority & Phasing

**Priority vs. Phase:** Priority labels (P0, P1, P2) indicate **importance** — P0 controls are the most impactful to build. Phase numbers (0–4) indicate **execution order** — Phase 0 is foundation work that all later phases depend on. A P0 control (like Menu System) appears in Phase 1 because it depends on Phase 0 token updates. A P1 control (like RadioButton) appears in Phase 2 because it has no earlier dependencies.

### Phase 0: Design Token Update (Foundation)

Update `ThemePalette` and visual style resolvers to new values. No new controls.

1. Update `ControlCornerRadius` from 2 → 4 DIP across all existing controls
2. Update `OverlayCornerRadius` to 8 DIP for PopupLayer, ComboBox dropdown, TooltipLayer
3. Add new palette fields (`cardBackground`, `accentHover`, `accentPressed`, `borderDefault`, `borderStrong`, `focusStrokeOuter`, `focusStrokeInner`, `smokeOverlay`)
4. Implement double-stroke focus ring in `Control` base class
5. Update icon size from 15 → 16 DIP
6. Add new `FontRole` entries (`BodyStrong`, `BodyLarge`, `Subtitle`, `Display`)
7. Update `Button` measurements to match spec (32 DIP height, 12 DIP padding, 120 DIP min-width)
8. Implement accent bottom border on `TextField` focus

**Files:** `DxUi.h`, `DxUi.Theme.cpp`, `DxUi.Controls.cpp`, `DxUi.TextInput.cpp`, `DxUi.WindowHost.cpp`

### Phase 1: Menu System (P0 — Largest GDI Dependency)

1. Implement `DxContextMenu` with all menu item types
2. Implement cascading submenu support
3. Implement `DxMenuBar` for horizontal menu bar
4. Migrate FolderView context menu as pilot
5. Migrate FolderWindow main menu
6. Remove `ThemedControls` owner-draw menu code

**Files:** New `DxUi.Menu.cpp`, `DxUi.h` additions, `FolderView.cpp`, `FolderWindow.cpp`, `ThemedControls.cpp`

### Phase 2: Missing Primitives (P1)

1. **RadioButton** + **RadioButtons group** control
2. **ProgressBar** (determinate + indeterminate/marquee)
3. **Standalone Scrollbar** extraction
4. **Toolbar** for Monitor
5. **StatusBar** multi-section upgrade

**Files:** `DxUi.Controls.cpp` (RadioButton, ProgressBar), new `DxUi.Scrollbar.cpp`, new `DxUi.Toolbar.cpp`, `DxUi.h` additions, `RedSalamanderMonitor/StatusStrip.cpp`

### Phase 3: Enhanced Controls (P1-P2)

1. Button variants (DropDown, Split, Hyperlink, Icon-only, Repeat)
2. Toggle switch visual update (knob size animation, WinUI state colors)
3. Checkbox indeterminate state
4. TextField clear button behavior alignment
5. ComboBox popup corner radius + shadow
6. Smoke overlay for modal dialogs
7. Card pattern formalization in CardPanel

**Files:** `DxUi.Controls.cpp` (Button variants, Checkbox indeterminate), `DxUi.Toggle.cpp`, `DxUi.TextInput.cpp` (clear button), `DxUi.ComboBox.cpp` (popup update), `DxUi.WindowHost.cpp` (smoke overlay), `DxUi.CardPanel.cpp`

### Phase 4: Advanced Effects (P2 — Optional)

1. Shadow rendering (D2D effects or pre-rendered)
2. Optional faux-Acrylic popup style (do not require a full DxUI blur/material pipeline)
3. Connected animation support
4. Page transition animations

**Files:** `DxUi.WindowHost.cpp` (shadow rendering pipeline / DWM backdrop hookup), `DxUi.Controls.cpp` (connected animation), `DxUi.Theme.cpp` (optional faux-Acrylic visual token integration)

### 1.1 Branch Snapshot (2026-04-12)

The `squad/dxui-winui-design` branch currently carries:

1. Phase 0 token/theme/type-ramp updates.
2. Phase 1 `DxContextMenu` + cascading submenu support, FolderView context-menu migration, and FolderWindow `DxMenuBar` top-level popup switching.
3. Phase 2-4 control work, including `PageHost` connected animations / page transitions, automated visual baselines, and the earlier `DrawDropShadow()` / DWM backdrop work, with unit coverage and perf markers.

Latest validation refresh:

- Passing command-selftest archives: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-12_173916` through `2026-04-12_173923`.
- `cmd_app_menuBar_arrow_switches_top_level_popup` records `DxUI::PopupShow` with `detail:"root-switch-keyboard"` in `.../2026-04-12_173920`.
- `cmd_app_menuBar_hover_switches_top_level_popup` records `DxUI::HitTest` plus `DxUI::PopupShow` with `detail:"root-switch-pointer"` in `.../2026-04-12_173921`.
- Focused `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` and `--suite=Rendering` passes were archived in `Specs/TestRuns/7d3a1247382a/DxUi/2026-04-12_184252`, alongside `animation_perf.jsonl`, `rendering_perf.jsonl`, and the updated `Tests/DxUiTests/Baselines/` golden PNGs.
- `RedSalamander\RedSalamander.vcxproj` built successfully in `Debug|x64` on 2026-04-12 after the motion / baseline changes.
- Full `.\.build\x64\Debug\DxUiTests.exe` coverage is still blocked by the unrelated pre-existing `Grid` clipboard setup failure (`clipboard initialized before grid copy`); that caveat is documented in `Specs/TestRuns/7d3a1247382a/DxUi/2026-04-12_184252/notes.md`.
- Earlier exploratory command archives from `2026-04-12_172832` through `2026-04-12_172839` remain useful for before/after comparison and the initial assertion-fix investigation.

---

## 2. Verification

### Visual Verification

- Compare rendered controls side-by-side with WinUI 3 Gallery app screenshots
- Test all states: rest, hover, pressed, focused, disabled
- Verify light theme, dark theme, high-contrast (black and white), and rainbow mode
- Verify at 100%, 125%, 150%, 200% DPI scales
- Verify reduced-motion mode (all animations should be instant or ≤83ms)
- **Automated baseline comparison:** Implemented in `DxUiTests.Rendering.cpp` via `AttachedHostWindow` + `WindowHost::DebugCaptureBitmap(...)`, with golden PNGs stored in `Tests/DxUiTests/Baselines/`. Current acceptance threshold: ≤2% differing pixels with per-channel tolerance 8 to absorb subpixel variation.

### Performance Verification

- **Frame budget:** All control rendering must complete within **16ms** (60fps target). Measure via ETW TraceLogging events in `WindowHost::Paint()`.
- **Menu popup latency:** Context menu must appear within **100ms** of right-click. Measure from `WM_CONTEXTMENU` to first `Paint()` of the menu surface.
- **Animation smoothness:** No frame drops during standard animations (hover transitions, tree expand). Monitor via `AnimationDispatcher` tick timing — flag any tick > 20ms.
- **Text layout caching:** `IDWriteTextLayout` objects must be cached and reused. Re-creation only on text change, font change, or DPI change. Verify via a layout-creation counter in debug builds.
- **Add ETW TraceLogging events** for: `DxUI::Paint` (per-frame), `DxUI::Layout` (per-measure), `DxUI::HitTest` (per-pointer-move), `DxUI::PopupShow`, `DxUI::FocusChange`.

### Functional Verification

- Menu keyboard navigation: arrow keys, Enter, Escape, mnemonics, Home/End
- Radio button group: arrow keys select, Ctrl+arrow moves focus, Tab enters/exits group
- Focus ring visible on all interactive controls when using keyboard
- Tab navigation visits all focusable controls in visual order
- Modal dialog traps focus (Tab cycles within dialog only)
- Popup light dismiss: click outside closes, Escape closes topmost

### Accessibility Verification

UI Automation must report correct control types and patterns for every new control:

| Control | UIA ControlType | Required Patterns | Required Properties |
|---------|----------------|-------------------|---------------------|
| Button | `Button` | Invoke | Name, IsEnabled, IsKeyboardFocusable |
| Toggle | `Button` | Toggle | Name, ToggleState |
| Checkbox | `CheckBox` | Toggle | Name, ToggleState |
| RadioButton | `RadioButton` | SelectionItem | Name, IsSelected |
| TextField | `Edit` | Value | Name, Value, IsReadOnly |
| ComboBox | `ComboBox` | ExpandCollapse, Selection, Value | Name, ExpandCollapseState |
| Menu | `Menu` | — | Name |
| MenuItem | `MenuItem` | Invoke (or Toggle) | Name, IsEnabled, AccessKey |
| Tree | `Tree` | Selection | Name |
| TreeItem | `TreeItem` | ExpandCollapse, SelectionItem | Name, IsExpanded |
| Grid | `DataGrid` | Grid, Table, Selection | RowCount, ColumnCount |
| ProgressBar | `ProgressBar` | RangeValue | Name, Value, Minimum, Maximum |
| Slider | `Slider` | RangeValue | Name, Value, Minimum, Maximum |
| ToggleSwitch | `Button` | Toggle | Name, ToggleState |
| Dialog | `Window` | Window | Name, IsModal |
| Tooltip | `ToolTip` | — | Name |

### Theme Validation Checklist

Each control must be verified across all theme combinations:

| | Light | Dark | HC Black | HC White | Rainbow | Custom (×6) |
|---|---|---|---|---|---|---|
| Rest | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Hover | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Pressed | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Focused | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Disabled | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |

High contrast: all colors from system HC tokens, no hardcoded values. Focus indicators: always visible, ≥2px stroke width.

### Regression

- Run existing DxUI test suite (`Tests/DxUiTests/` — 14 test modules)
- Run self-test commands that exercise Preferences pages (already DxUI)
- Run RedSalamander self-tests (`Tools/Run-AllTests.ps1` — Commands/309, Compare/141, FileOps/68 cases)
- Manual smoke test of all migrated dialogs
- Visual baseline comparison (automated, see above)

---

## 3. Reference

- [WinUI Design Principles](https://learn.microsoft.com/en-us/windows/apps/design/design-principles)
- [WinUI Typography](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/typography)
- [WinUI Color](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/color)
- [WinUI Geometry (Rounded Corners)](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/geometry)
- [WinUI Motion](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/motion)
- [WinUI Materials](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/materials)
- [WinUI Layering & Elevation](https://learn.microsoft.com/en-us/windows/apps/design/signature-experiences/layering)
- [WinUI Buttons](https://learn.microsoft.com/en-us/windows/apps/design/controls/buttons)
- [WinUI Menus](https://learn.microsoft.com/en-us/windows/apps/design/controls/menus)
- [WinUI Radio Buttons](https://learn.microsoft.com/en-us/windows/apps/design/controls/radio-button)
- [WinUI Toggle Switch](https://learn.microsoft.com/en-us/windows/apps/design/controls/toggles)
- [WinUI Text Box](https://learn.microsoft.com/en-us/windows/apps/design/controls/text-box)
- [Fluent 2 Design System](https://fluent2.microsoft.design/)
- [Windows UI Kit for Figma](https://aka.ms/WinUI/3.0-figma-toolkit)
- Authoritative contract: `Specs/UI/UI_DxUiWinUIDesign.md`

