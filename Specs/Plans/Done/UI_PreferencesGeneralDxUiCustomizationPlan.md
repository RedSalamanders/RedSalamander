# Preferences General DxUI Customization Plan

**Author:** Codex
**Date:** 2026-04-15
**Status:** Complete (validated on build 232; ready to move to `Plans/Done`)
**Scope:** Add persisted user-facing DxUI customization controls to `Preferences -> General` for compact mode, reduced motion, and whole-window backdrop selection, then update the authoritative specs and settings schema to match.

---

## 0. Checklist

- [x] Confirm the persisted settings model for DxUI customization (`compactMode`, `reducedMotion`, `windowBackdrop`) and keep it backward-compatible with the current settings reader/writer.
- [x] Add the new settings metadata to `Specs/SettingsStore.schema.json` with stable `title`, `description`, `x-ui-pane`, `x-ui-order`, and section placement for `Preferences -> General`.
- [x] Extend `Common::Settings` load/save logic so the new settings round-trip, hot-reload safely, and default correctly when absent.
- [x] Extend `Preferences -> General` with a new DxUI customization section and wire the currently-implemented controls to `workingSettings`, `Apply`, `OK`, and `Cancel`.
- [x] Implement the first real DxUI compact-mode slice and expose a live `Compact mode` toggle in `Preferences -> General`.
- [x] Finish the remaining compact-density rollout promised by this plan across the broader framework surfaces (`Toolbar`, `StatusStrip`, and the remaining shared form-control metrics).
- [x] Finish the remaining compact-density rollout on shared data surfaces so compact mode visibly affects search results and other DxUI `Grid` / list-style views.
- [x] Fix compact-density regressions in shared form controls, especially the current `ComboBox` clipping / left-padding bug in compact mode.
- [x] Make the compact-mode `ComboBox` field and popup geometry visually correct: no clipped field, no clipped popup row, and restored left padding/gap before text.
- [x] Add a user override path for reduced motion instead of relying only on the current system-derived `ThemePalette::reducedMotion`.
- [x] Make `Window backdrop` visibly affect DxUI popup/menu materials and preview immediately in Preferences.
- [x] Finish the shared whole-window backdrop rollout across every supported top-level window and keep one authoritative policy/helper path.
- [x] Keep popup/menu materials visually aligned with the chosen backdrop policy so the setting is obvious in the live product, not only in top-level DWM state.
- [x] Finish the remaining menu/chrome rollout details required for this plan:
  - drive and disk dropdown menus in `NavigationView` must use DxUI instead of legacy themed `TrackPopupMenu`,
  - the top menu bar must honor right-justified items such as `Help`,
  - navigation/history popup menus must not dismiss immediately on the opening click/release path.
- [x] Revalidate the live navigation dropdown click/hover paths so drive/history/menu popups stay open on the opening click and do not dismiss until an intentional follow-up action.
- [x] Keep the top menu bar layout contract honest by validating that `Help` stays right-aligned whenever the menu definition marks it with `MFT_RIGHTJUSTIFY`.
- [x] Fix the `Compare Directories` DxUI options body scroll stability regression so scrolling to lower cards does not jump back to the top.
- [x] Keep the `Compare Directories` options surface scrollable all the way to the lower ignore-pattern cards without focus or layout updates snapping the viewport back to the top.
- [x] Update the authoritative specs after implementation: `Specs/UI/UI_DxUiWinUIDesign.md`, `Specs/UI/UI_PreferencesDialog.md`, `Specs/Core/Core_SettingsStore.md`, and `Specs/UI/UI_TopLevelToolWindows.md` if backdrop behavior becomes normative there.
- [x] Add focused deterministic selftests, settings round-trip coverage, and archived perf evidence for the current implementation slice.
- [x] Refresh the final whole-plan validation/archive pass before moving this plan out of `WIP`.

Closeout notes:

- The `Preferences -> General -> DxUI` section is now complete and honest in the live product:
  - `Compact mode` previews immediately and persists through `Apply` / `OK`,
  - `Reduced motion` exposes `System / On / Off`,
  - `Window backdrop` exposes `Default / None / Mica / Mica Alt / Acrylic`, previews immediately in Preferences, and applies to supported top-level windows through one shared fallback-aware helper path.
- Compact density is no longer limited to the first rollout slice. The completed scope now includes the Preferences General page, menu bar, popup menus / flyouts, drive/history navigation menus, combo boxes, Monitor toolbar/status strip metrics, Compare Directories option cards, and Find Files results/status surfaces.
- The regressions raised during rollout are closed as plan exit criteria, not left as polish:
  - drive and history menus now use the DxUI popup contract and do not dismiss on the opening click,
  - right-justified menu items such as `Help` stay anchored to the trailing edge,
  - compact-mode combo field/popup geometry keeps the left text inset and no longer clips,
  - Compare Directories options can scroll to the lower ignore-pattern cards without snapping back,
  - Find Files results visibly shrink in compact mode.
- Final current-tree validation/evidence for the completed plan is archived under:
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-16_082958` — `settings_ui_customization_roundtrip`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-16_082959` — `cmd_preferences_dialog_general_window_backdrop_apply_updates_supported_windows`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-16_083002` — `cmd_pane_find_dialog_compact_mode_shrinks_results_grid_metrics`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-16_083004` — `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-16_083006` — `cmd_pane_navigation_show_folders_history_keeps_navigation_shell_stable`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-16_083007` — `cmd_app_menuBar_right_justified_items_anchor_to_trailing_edge`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-16_083008` — `cmd_app_menuBar_hover_switches_top_level_popup`
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-16_083010` — `cmd_app_openDriveMenus`
  - `Specs/TestRuns/7d3a1247382a/Monitor/2026-04-16_083023` — `RedSalamanderMonitor.exe --chrome-selftest --wait-instance`
- Focused framework regression coverage also passed on the same closeout tree:
  - `DxUiTests.exe --suite=ComboBox`
- Key closeout perf markers:
  - `dxui.menu.popup.paint`: `560-710 us` in the current menu hover / drive-menu runs
  - `render.present_us`: `243-292 us` steady-state during the current menu popup runs
  - `monitor.ui.chrome_ready_us`: `14536 us`
  - `monitor.ui.toolbar_toggle_us`: `40280 us`
  - `monitor.ui.status_filter_sync_us`: `2211 us`
  - `monitor.ui.toolbar_present_failure_count`: `0`
  - `monitor.ui.status_present_failure_count`: `0`
- `Specs/SettingsStore.schema.json` already carries the final `ui.compactMode`, `ui.reducedMotion`, and `ui.windowBackdrop` schema contract from the implementation slice; no additional schema-shape change was needed during closeout.

---

## 1. Background

Current state in tree:

- `Preferences -> General` now exposes:
  - `mainMenu.menuBarVisible`
  - `mainMenu.functionBarVisible`
  - `ui.compactMode`
  - `ui.reducedMotion`
  - `ui.windowBackdrop`
  - `startup.showSplash`
- `DxUi` already has a real `ThemePalette::reducedMotion` flag, and this plan now adds a user override pipeline on top of the OS-derived default.
- `DxUi` already has `WindowHost::SetSystemBackdrop(...)`, and the app now drives it from persisted settings on several top-level windows, but the rollout is not yet complete across every supported window owner.
- Compact mode is specified in `Specs/UI/UI_DxUiWinUIDesign.md`, and this plan now implements the initial runtime `Density` contract in `Common/DxUi`; the remaining work is to finish the broader control-surface rollout promised below.

That means:

- `reduced motion` is now user-configurable,
- `window backdrop` is now both persisted and visibly previewable, but its full top-level rollout/spec closeout is still open,
- `compact mode` is now honest to expose, but the framework rollout remains incomplete beyond the first implemented surfaces.

---

## 2. Product Decision

### 2.1 Preferences Placement

These settings belong in `Preferences -> General`, not `Themes`:

- they are app-wide interaction and chrome preferences,
- they affect behavior as well as appearance,
- they should be available even when the user never edits a custom theme.

Recommended `General` page grouping after implementation:

1. `Display`
   - Menu Bar
   - Function Bar
2. `DxUI`
   - Compact mode
   - Reduced motion
   - Window backdrop
3. `Startup`
   - Splash screen

### 2.2 Persisted Settings Model

Recommended storage model: add a new top-level `ui` object in the settings store rather than overloading `theme`, `mainMenu`, or `startup`.

Recommended C++ model:

```cpp
namespace Common::Settings
{
    enum class ReducedMotionMode : uint8_t
    {
        System,
        On,
        Off,
    };

    enum class WindowBackdropMode : uint8_t
    {
        Default,
        None,
        Mica,
        MicaAlt,
        Acrylic,
    };

    struct UiSettings
    {
        bool compactMode = false;
        ReducedMotionMode reducedMotion = ReducedMotionMode::System;
        WindowBackdropMode windowBackdrop = WindowBackdropMode::Default;
    };
}
```

Recommended JSON shape:

```json
{
  "ui": {
    "compactMode": false,
    "reducedMotion": "system",
    "windowBackdrop": "default"
  }
}
```

Rationale:

- `compactMode` is naturally a boolean in the persisted product contract even if the framework uses an internal `Density` enum.
- `reducedMotion` needs a tri-state because `System / On / Off` is the correct user model.
- `windowBackdrop` needs an enum because it is a policy choice, not a boolean.

### 2.3 Settings Versioning Recommendation

Recommended approach: keep `schemaVersion` at `11` for this change.

Reason:

- the new `ui` object is additive and optional,
- the current reader already ignores unknown top-level keys,
- the current loader rejects unsupported schema versions and has no generic migration path,
- bumping to `12` would create avoidable compatibility churn for a backward-compatible settings extension.

If maintainers want strict version increments for all schema changes, that must be coupled with explicit `v11` and `v12` load support in `Common/Common/SettingsStore.cpp`. That is a separate compatibility decision and should not be hidden inside this feature.

---

## 3. User-Facing Contract

### 3.1 Compact Mode

General-pane control:

- `Compact mode`
- type: toggle
- default: `Off`

Behavior:

- Controls whether supported DxUI containers use `Density::Standard` or `Density::Compact`.
- Must preview inside the live Preferences dialog as soon as the working setting changes so the user can see the density difference before `Apply`.
- Must only become the persisted app-wide default on `Apply` or `OK`.

Initial required rollout surfaces:

- `Preferences.*` direct-host pages
- `Toolbar`
- `StatusStrip`
- `MenuBar`
- popup menus / flyouts
- basic form controls used on Preferences and schema-driven pages:
  - `Button`
  - `Toggle`
  - `Checkbox`
  - `RadioButtons`
  - `TextField`
  - `ComboBox`

Initial non-goals for the first compact-mode pass:

- `Grid` and `Tree` row-density redesign
- plugin-specific custom renderers that do not use DxUI control layout
- rewriting old GDI surfaces just to honor compact mode

### 3.2 Reduced Motion

General-pane control:

- `Reduced motion`
- type: combo or segmented choice
- values: `System`, `On`, `Off`
- default: `System`

Behavior:

- `System`: use current OS reduced-motion preference.
- `On`: force reduced motion on all DxUI surfaces even if the OS setting is off.
- `Off`: force normal motion even if the OS setting prefers reduced motion.

Preview behavior:

- Must update the live Preferences dialog immediately from `workingSettings`.
- Must update the rest of the app only after `Apply` / `OK` or after external settings reload.

### 3.3 Window Backdrop

General-pane control:

- `Window backdrop`
- type: combo
- values: `Default`, `None`, `Mica`, `Mica Alt`, `Acrylic`
- default: `Default`

Behavior:

- Applies only to supported top-level windows.
- Must gracefully fall back to `None` when the OS, window style, or accessibility mode does not support the requested backdrop.
- High contrast remains authoritative; backdrop preference must never break high-contrast readability.

Recommended `Default` policy:

- primary app windows: prefer `Mica` when supported,
- secondary/tool windows: prefer `Mica Alt` or `None` depending on visual fit,
- unsupported environments: `None`.

The plan does not require faux-Acrylic popup rendering. This remains outside scope and stays governed by `Specs/UI/UI_DxUiWinUIDesign.md`.

---

## 4. Implementation Workstreams

### 4.1 Settings Store And Schema

Files:

- `Common/SettingsStore.h`
- `Common/Common/SettingsStore.cpp`
- `Specs/SettingsStore.schema.json`
- `Specs/Core/Core_SettingsStore.md`

Work:

1. Add `UiSettings` and the new enums to `Common::Settings`.
2. Add `std::optional<UiSettings> ui;` to `Common::Settings::Settings`.
3. Update load/save logic for:
   - absent `ui` object -> defaults,
   - invalid enum strings -> load failure on strict path, recovery/defaults on tolerant path,
   - canonical write ordering with omitted-default behavior aligned to current settings-store rules.
4. Extend the base schema with a new `uiSettings` definition and top-level `ui` property.
5. Add stable schema metadata for the three fields:
   - `title`
   - `description`
   - `x-ui-pane: "General"`
   - `x-ui-order`
   - `x-ui-section`

Recommended schema metadata:

- `ui.compactMode`
  - section: `DxUI`
  - order: `40`
  - control: `toggle`
- `ui.reducedMotion`
  - section: `DxUI`
  - order: `50`
  - control: `custom`
- `ui.windowBackdrop`
  - section: `DxUI`
  - order: `60`
  - control: `custom`

### 4.2 Preferences General Pane

Files:

- `RedSalamander/Preferences.General.cpp`
- `RedSalamander/Preferences.h`
- `RedSalamander/Preferences.Dialog.cpp`
- `RedSalamander/resource.h`
- `RedSalamander/*.rc`

Work:

1. Add a new `DxUI` card section to the General page.
2. Add retained controls and debug-focus identifiers for:
   - compact mode
   - reduced motion
   - window backdrop
3. Keep the existing direct-host/card contract intact:
   - no visible native fallback,
   - card layout remains stable under DPI/theme changes,
   - `Apply` / `Cancel` semantics remain correct.
4. General-page preview rules:
   - compact mode relayouts the page immediately,
   - reduced motion affects the page host immediately,
   - backdrop previews on the Preferences top-level window immediately if possible,
   - `Cancel` restores the previous effective settings.

### 4.3 Compact Mode Framework Work

Files:

- `Common/DxUi/DxUi.h`
- `Common/DxUi/DxUi.cpp`
- `Common/DxUi/DxUi.Controls.cpp`
- `Common/DxUi/DxUi.Menu.cpp`
- `Common/DxUi/DxUi.ComboBox.cpp`
- `Common/DxUi/DxUi.TextInput.cpp`
- `Common/DxUi/DxUi.Scrollbar.cpp`

Work:

1. Add an inherited density concept to DxUI, analogous to `FlowDirection`.
2. Add a framework-facing API such as:

```cpp
enum class Density : uint8_t
{
    Standard,
    Compact,
};
```

3. Allow containers to set density and child controls to inherit it.
4. Update supported controls to honor compact metrics where the authoritative spec already defines them.
5. Keep minimum hit targets and accessibility behavior intact even when visuals shrink.
6. Add rendering and layout validation for both standard and compact density.

Important constraint:

- Do not expose the General-pane `Compact mode` setting until this work exists. A toggle that only changes a few hardcoded Preferences row heights would not satisfy the `DxUi` spec and would create a misleading product contract.

### 4.4 Reduced Motion Override Pipeline

Files:

- `RedSalamander/AppTheme.h`
- `RedSalamander/AppTheme.cpp`
- `RedSalamander/DxUiThemePalette.h`
- `Common/DxUi/DxUi.Theme.cpp`
- any viewer/plugin hosts that still build a `ThemePalette` directly

Work:

1. Introduce an app-level reduced-motion preference resolver:
   - system preference
   - user override
   - resolved effective bool
2. Feed the resolved value into every DxUI palette creation path, not just the main app shell.
3. Ensure plugin/viewer windows that derive `ThemePalette` from `ViewerTheme` also receive the override when the app settings apply to them.
4. Verify that reduced-motion override keeps current behavior correct for:
   - page transitions,
   - connected animations,
   - sort glyph animations,
   - expand/collapse animations,
   - toggle/radio animations,
   - scrollbar auto-hide visibility rules.

### 4.5 Whole-Window Backdrop Rollout

Files:

- `Common/DxUi/DxUi.h`
- `Common/DxUi/DxUi.WindowHost.cpp`
- top-level window owners such as:
  - `RedSalamander/RedSalamander.cpp`
  - `RedSalamander/Preferences.Dialog.cpp`
  - `RedSalamander/FindFilesWindow.cpp`
  - `RedSalamander/CompareDirectoriesWindow.cpp`
  - `RedSalamander/ConnectionManagerDialog.cpp`
  - `RedSalamander/ManagePluginsDialog.cpp`
  - `RedSalamanderMonitor/RedSalamanderMonitor.cpp`

Work:

1. Refactor the DWM backdrop application logic so it can be applied to real top-level HWNDs, not only to existing `WindowHost` owners.
2. Define one shared helper/policy path for mapping the persisted setting to the effective DWM backdrop.
3. Apply the policy consistently to supported top-level windows.
4. Ensure top-level windows re-apply the backdrop when:
   - theme changes,
   - settings apply,
   - DPI/theme/window recreation paths occur.
5. Fallback cleanly to `None` without noisy error logging when the OS does not support the requested backdrop.

Important constraint:

- A child DxUI host is not the same as a whole top-level window. The implementation must not pretend that setting a backdrop on an embedded child host solves the product requirement.

---

## 5. Authoritative Spec And Schema Updates Required At Closeout

These documents must be updated as part of the implementation, not left for later:

### 5.1 `Specs/UI/UI_DxUiWinUIDesign.md`

Update or add normative text for:

- density / compact mode as an implemented runtime contract rather than a spec-only section,
- reduced-motion override semantics (`System / On / Off`) and precedence over the OS default,
- window-backdrop policy, fallback rules, and supported surface types.

### 5.2 `Specs/UI/UI_PreferencesDialog.md`

Update:

- `General` page contract,
- settings-and-schema mapping,
- control list / behavior for the new DxUI customization section,
- `Apply` / `Cancel` preview semantics if local preview is added.

### 5.3 `Specs/Core/Core_SettingsStore.md`

Update:

- settings data model to document the new `ui` object,
- field names, defaults, allowed values, and persistence semantics,
- schema compatibility decision for this additive change.

### 5.4 `Specs/SettingsStore.schema.json`

Update:

- top-level `ui` object,
- `$defs` for `uiSettings`,
- stable `title` / `description` metadata,
- `x-ui-pane` / `x-ui-order` / `x-ui-section` annotations used by Preferences documentation and export tooling.

### 5.5 `Specs/UI/UI_TopLevelToolWindows.md`

Update only if backdrop behavior becomes part of the shared tool-window contract for:

- when backdrop may be used,
- when it must fall back,
- whether preview/apply behavior is required on Preferences and other long-lived tool windows.

### 5.6 Optional Follow-On Doc Touches

Only if implementation meaningfully changes these contracts:

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/TestRuns/README.md`

---

## 6. Validation And Perf Plan

### 6.1 Settings Round-Trip

Required:

- add unit/selftest coverage in `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` for:
  - missing defaults,
  - persisted round-trip of all three new settings,
  - invalid enum strings rejected on strict load path,
  - hot-reload preserving runtime-only state while applying the new persisted values.

### 6.2 Preferences General Page

Required:

- extend `General` page debug focus targets in `RedSalamander/Preferences.h`,
- add command/selftests for:
  - `General` page exposes the new DxUI controls,
  - `Cancel` discards unapplied changes,
  - `Apply` persists and updates effective runtime state,
  - `General -> other page -> General` round-trips keep retained state and do not recreate visible native fallback.

### 6.3 Compact Mode

Required:

- focused `DxUiTests` rendering/layout coverage for standard vs compact density,
- at least one General-page selftest proving the working-setting toggle relayouts the page and keeps controls accessible,
- archived visual/perf evidence for compact-mode rendering and layout.

### 6.4 Reduced Motion

Required:

- `DxUiTests` coverage proving `System / On / Off` maps to the effective `ThemePalette::reducedMotion` value,
- at least one animation/path selftest proving the override changes behavior immediately and deterministically,
- no caret/input regressions on text entry surfaces under reduced motion.

### 6.5 Window Backdrop

Required:

- focused app-window selftests proving the selected backdrop mode is applied to supported windows and falls back safely,
- debug signal or query path to verify the effective backdrop without relying only on screenshots,
- no `Present failed`, no resize-failure churn, and no activation/focus regressions during apply/revert.

### 6.6 Perf Evidence

Use existing instrumentation where it is already sufficient, and add focused markers only where the new work would otherwise be opaque.

Expected evidence:

- build + selftests green,
- archived runs under `Specs/TestRuns/<MachineHash>/Commands/...`,
- if new markers are needed, add scoped timings for:
  - applying General-page DxUI customization changes,
  - backdrop application/re-application,
  - compact-mode relayout if it becomes measurable work.

---

## 7. Recommended Implementation Order

1. Lock the persisted settings model and schema shape.
2. Implement reduced-motion override end-to-end.
3. Implement top-level backdrop policy and preview/apply behavior.
4. Implement true DxUI compact mode.
5. Extend `Preferences -> General` to expose all three settings together.
6. Refresh authoritative docs and archive final validation/perf evidence.

Reason:

- reduced motion and backdrop already have framework footholds,
- compact mode is the only item that still requires real framework work before the product setting is honest,
- landing the model first avoids churn in Preferences and the docs.

---

## 8. Closeout Criteria

This plan is complete only when all of the following are true:

- the settings persist and reload correctly,
- the General pane exposes the three settings on the shared DxUI path,
- compact mode is a real DxUI capability, not a General-page-only layout hack,
- reduced-motion override affects all supported DxUI surfaces consistently,
- backdrop selection applies to the agreed supported top-level windows with safe fallback,
- the authoritative specs and schema are updated in the same branch,
- selftests and archived perf evidence exist for the final tree.
