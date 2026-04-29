# Preferences Dialog Win32 Removal: Remaining Work Plan

Last updated: 2026-03-25

## Status

The implementation migration is finished architecturally, and the automated verification phase is now complete. The only remaining phase is the final manual pane-by-pane Preferences validation pass, which requires human-driven UI interaction.

What is now true in the code:

- `RedSalamander/Preferences.Internal.h` no longer includes `<commctrl.h>`, no longer carries pane-child-window helper seams, and now exposes the shared shell HWND state with clearer names such as `categoryTreeWindow` and `pageHostWindow`.
- Pane headers now use the simplified `InitializePage(...)`, `Refresh(...)`, `LayoutPage(...)`, and optional `OnVisibilityChanged(...)` contract instead of the old `EnsureCreated` / `Hwnd` / `CreateControls` / `LayoutControls` / `HandleCommand` / `HandleNotify` seam.
- `RedSalamander/Preferences.Dialog.cpp` no longer includes `<commctrl.h>`, no longer initializes common-control classes for Preferences, and now acts as the dialog shell and DxUi host coordinator.
- Pane-local and shared page-content layout now flows through DxUi surfaces rather than hidden Win32 child-control fallback branches.
- Static searches now reach zero remaining Preferences-layer hits for `<commctrl.h>`, `ListView_*`, `TreeView_*`, `NMHDR`, and `INITCOMMONCONTROLSEX`.
- The full automated Preferences self-test sweep passes in `Debug|x64`, including activation, retained-state, search, selection, scrolling, shell chrome, and long-run list validation cases.

This plan now tracks only the remaining human verification closure rather than implementation migration.

## Top Checklist

### Architecture

- [x] Lock the final boundary: keep only the top-level dialog HWND and DxUi host HWNDs; remove all remaining Preferences-specific Win32/common-control child UI and routing.
- [x] Replace the pane contract (`EnsureCreated`, `Hwnd`, `CreateControls`, `LayoutControls`, `HandleCommand`, `HandleNotify`) with a DxUi-first pane interface.
- [x] Remove `WM_COMMAND`, `WM_NOTIFY`, and `PostMessageW` callback bridges from pane interaction flows.
- [x] Remove remaining Preferences-specific common-control dependencies and `<commctrl.h>` includes.
- [x] Remove remaining hidden legacy child controls and helper HWND wrappers from the Preferences layer.
- [x] Collapse remaining Win32 layout code to DxUi layout only.
- [x] Simplify `PreferencesDialogState` so it only stores dialog state, settings state, and top-level DxUi host references.

### Pane Migration

- [x] General
- [x] Panes
- [x] Viewers
- [x] Editors
- [x] Keyboard
- [x] Mouse
- [x] Themes
- [x] Plugins
- [x] Advanced
- [x] Compare Directories
- [x] Hot Paths

### Verification

- [x] Add or update self-tests for pane activation, input, scrolling, resize, and Apply/OK flows.
- [ ] Run a full manual Preferences pass across every pane after the final cleanup. This item is human-only and cannot be closed by automated self-tests alone.
- [x] Verify `Debug|x64` and `Release|x64` builds.
- [x] Confirm no remaining Preferences-layer `WM_NOTIFY` / common-control regressions.

## Target End State

The target is not "remove every HWND from the process". The realistic end state is:

- Keep the top-level Preferences dialog HWND.
- Keep the minimum HWNDs required to host DxUi surfaces (`WindowHost` attachment points).
- Remove all pane-owned Win32 child controls used as visual controls, hidden mirrors, message routers, or layout placeholders.
- Remove Preferences-specific common-control usage from pane logic.
- Make pane interaction flow:

```text
DxUi callback -> update pane/dialog state -> call pane logic directly -> re-sync DxUi controls if needed
```

The target is explicitly not:

- rewriting the overall top-level dialog shell away from Win32,
- removing the need for `WindowHost` attachment HWNDs,
- or changing unrelated app windows outside the Preferences feature area.

## Completed Foundation

The following work is already done and should not be re-planned as remaining work:

- Most per-setting storage HWNDs were removed and replaced with state fields.
- Keyboard, Themes, Viewers, and Plugins already have real DxUi surfaces and retained hosts.
- Plugin schema-driven configuration fields already render through DxUi controls instead of dynamic Win32 edit/combo/toggle creation.
- Many visual-only statics, frames, and routing-only controls were already deleted.
- Several retained-host interaction and resize bugs were already fixed.

The remaining work is verification of the finished migration, not more architectural conversion.

## Remaining Workstreams

### 1. Replace the Pane Contract [completed]

Completed in this phase:

- All Preferences pane headers now expose the simplified DxUi-first contract: `InitializePage(...)`, `Refresh(...)`, `LayoutPage(...)`, and optional `OnVisibilityChanged(...)`.
- `Hwnd()`, `EnsureCreated(...)`, `ResizeToHostClient(...)`, `CreateControls(...)`, `LayoutControls(...)`, `HandleCommand(...)`, and `HandleNotify(...)` are removed from pane interfaces.
- `Preferences.Dialog.cpp` no longer queries pane-owned HWNDs, no longer resizes pane child windows, and no longer treats panes as child-HWND owners.
- `Preferences.Internal.*` no longer exposes the old pane-host helper trio (`EnsureCreated(...)`, `ResizeToHostClient(...)`, `Show(...)`).

Verification completed:

- Static search reaches zero remaining pane-interface hits for `EnsureCreated(`, `ResizeToHostClient(`, `Hwnd(`, `CreateControls(`, `LayoutControls(`, `HandleCommand(`, and `HandleNotify(` under `RedSalamander/Preferences*`.
- Focused self-tests passed for category switching, retained-page state, and the migrated DxUi search/update flows.

Residual note:

- This does not finish the overall Win32-removal plan. The remaining phases still need to remove common-control dependencies, hidden helper controls, and pane-local Win32 layout scaffolding.

### 2. Replace Fake Command Routing With Direct or Explicit Deferred Actions

Current problem:

- Several DxUi callbacks still post fake `WM_COMMAND` messages to the dialog or page host.
- That posted-command path is doing two different jobs today:
  - legacy command routing,
  - and deferral out of a live DxUi callback to avoid reentrancy problems.
- The pane logic still depends on command IDs and notify codes instead of explicit pane actions.

Evidence:

- `RedSalamander/Preferences.Keyboard.cpp`
- `RedSalamander/Preferences.Themes.cpp`
- `RedSalamander/Preferences.Plugins.cpp`
- `RedSalamander/Preferences.Viewers.cpp`
- `RedSalamander/Preferences.Dialog.cpp`

Required work:

- Split the current bridge usage into two categories:
  - direct pane method calls for safe synchronous work,
  - explicit deferred pane actions for work that must run after the active DxUi callback unwinds.
- Keep deferral for reentrancy-sensitive flows such as:
  - page refresh/rebuild,
  - selection/model changes that can invalidate the active callback path,
  - page navigation,
  - text-input bridge and retained-host rebinding,
  - nested UI launches or heavy theme/plugin actions.
- Replace fake `WM_COMMAND` routing with either:
  - direct method calls, or
  - a typed Preferences deferred-action message/payload path.
- Move any remaining command-ID dispatch bodies into named pane methods.
- Delete now-dead command-ID switch arms from `Preferences.Dialog.cpp`.
- Remove `WM_NOTIFY` forwarding once the last pane stops using it.

Exit criteria:

- No DxUi callback in Preferences posts `WM_COMMAND` or `WM_NOTIFY` just to invoke local pane logic.
- Reentrancy-sensitive flows still defer, but do so through an explicit typed pane-action mechanism rather than fake command IDs.
- `Preferences.Dialog.cpp` no longer has pane-specific `WM_NOTIFY` routing.

### 3. Remove Remaining Common-Control Dependencies

Current problem:

- Preferences still includes `<commctrl.h>` in multiple units.
- Some panes still keep `ListView_*` compatibility helpers even when the visible list is now DxUi.

Evidence:

- `RedSalamander/Preferences.Dialog.cpp`
- `RedSalamander/Preferences.Internal.cpp`
- `RedSalamander/Preferences.Internal.h`
- `RedSalamander/Preferences.Plugins.cpp`
- `RedSalamander/Preferences.Themes.cpp`
- `RedSalamander/Preferences.Plugin.Configuration.cpp`

Required work:

- Audit every remaining `<commctrl.h>` include under `RedSalamander/Preferences*`.
- Delete common-control code that is now only legacy scaffolding:
  - `ListView_InsertColumn`
  - `ListView_SetColumnWidth`
  - `ListView_GetHeader`
  - any no-longer-used `NMHDR` / notify shims
- Remove `InitCommonControlsEx(... ICC_LISTVIEW_CLASSES ...)` if nothing in Preferences requires it anymore.

Exit criteria:

- No Preferences source file includes `<commctrl.h>` unless it is still justified by an intentional keep.
- No pane depends on `ListView_*`, `TreeView_*`, or `NMHDR` for its active UI behavior.

### 4. Collapse Remaining Win32 Layout Paths

Current problem:

- The pane-local fallback layout branches are now gone, but the shared Preferences shell still mixes Win32 child positioning with DxUi page layout.
- The remaining cleanup is about deleting shared shell/layout scaffolding rather than finishing individual pane content surfaces.

Evidence:

- `RedSalamander/Preferences.Dialog.cpp`
- `RedSalamander/Preferences.Internal.cpp`

Required work:

- Move every remaining pane to a single `LayoutDxPage(...)` path.
- Delete legacy control sizing code once each pane is fully on DxUi.
- Stop hiding inactive legacy peers during layout; instead remove them entirely.
- Stop using child-HWND visibility as a state-management mechanism.

Exit criteria:

- No pane-specific `SetWindowPos` / `ShowWindow` / `EnableWindow` calls remain for active page content controls.
- Layout changes flow only through DxUi tree/layout state plus top-level host sizing.

### 5. Simplify Shared Preferences Infrastructure

Current problem:

- `PreferencesDialogState` still mixes real dialog state with old shell/HWND plumbing and compatibility helpers.
- `Preferences.Internal.*` still exposes helper functions aimed at legacy pane HWND ownership.

Evidence:

- `RedSalamander/Preferences.Internal.h`
- `RedSalamander/Preferences.Internal.cpp`

Required work:

- Separate dialog shell state from pane interaction state if that clarifies ownership.
- Remove helper APIs that only exist for pane child HWND lifecycle:
  - `EnsureCreated(...)`
  - `ResizeToHostClient(...)`
  - `Show(...)`
  - legacy visibility/count helpers tied to child windows
- Rename any remaining misleading fields such as `categoryTree` if the surface is actually a DxUi host/control and not a Win32 tree anymore.
- Reduce state to:
  - settings/baseline/working copies,
  - current category/selection/filter state,
  - top-level host references,
  - DxUi control references that are truly required across functions.

Exit criteria:

- `Preferences.Internal.h` no longer reads like a pane-HWND management layer.
- `PreferencesDialogState` has no pane-control HWND fields.

## Pane-by-Pane Remaining Work

### General

Phase 4 result:

- Uses the final pane contract.
- Uses only retained DxUi card/toggle layout.
- No pane-local `ShowWindow` / `SetWindowPos` / `EnableWindow` / `CreateWindowExW` fallback code remains.

### Panes

Phase 5 result:

- Uses the final pane contract.
- Uses only retained DxUi card/toggle/combo/edit layout.
- No pane-local `SetWindowPos` / `ShowWindow` / `EnableWindow` fallback branch remains.

### Viewers

Phase 5 result:

- Uses the final pane contract plus typed deferred pane actions where reentrancy still matters.
- No pane-local `HandleCommand(...)`, `HandleNotify(...)`, or fake `WM_COMMAND` callback bridge remains.
- Removed the ListView-era column helper path; the pane stays on the DxUi grid/editor surface only.

### Editors

Phase 4 result:

- Placeholder pane remains intentionally simple.
- Uses the final pane contract and only the shared DxUi note/empty-state surface.
- No pane-local child-HWND state or layout scaffolding remains.

### Keyboard

Phase 5 result:

- Uses the final pane contract plus typed deferred pane actions for search, scope, and button flows.
- No pane-local `HandleNotify(...)` or fake `WM_COMMAND` callback bridge remains.
- Removed the no-longer-used ListView helper code while preserving the DxUi-backed keyboard capture and grid flows.

### Mouse

Phase 4 result:

- Matches the Editors cleanup.
- Uses the final pane contract and only the shared DxUi note/empty-state surface.
- No pane-local child-HWND state or layout scaffolding remains.

### Themes

Phase 5 result:

- Uses the final pane contract plus typed deferred pane actions for theme/base/name/search flows.
- No pane-local `HandleNotify(...)` or fake `WM_COMMAND` callback bridge remains.
- Removed the last dead legacy visibility/layout helper from the pane; the active workflow is the DxUi grid, swatch, and editor surface only.

### Plugins

Phase 5 result:

- Uses the final pane contract plus typed deferred pane actions for search and plugin actions.
- No pane-local `HandleNotify(...)` or fake `WM_COMMAND` callback bridge remains.
- Removed the stale HWND button/layout branches; overview, feedback, and configuration remain on the shared DxUi surface.

### Advanced

Phase 6 result:

- Uses the final pane contract.
- Uses only the retained DxUi card/combo/edit/toggle layout path.
- Removed the dead hidden-control fallback branch and zero-binding HWND cleanup scaffolding.

### Compare Directories

Phase 6 result:

- Uses the final pane contract.
- Uses only the retained DxUi layout path for toggles, combo, and ignore-pattern editors.
- Removed the dead per-control child-host wrapper structs and the hidden Win32 fallback layout branch.

### Hot Paths

Phase 4 result:

- Uses the final pane contract.
- Uses only retained DxUi card/row layout for slot editing and the assign-preferences toggle.
- No pane-local `SetWindowPos` / `ShowWindow` / `EnableWindow` / `CreateWindowExW` fallback code remains.

## Recommended Execution Order

### Phase 1: Callback bridge cleanup [completed]

Completed in code on 2026-03-25:

- [x] Introduced a typed Preferences deferred-action path instead of fake `WM_COMMAND` routing for retained DxUi callbacks.
- [x] Migrated `Keyboard`, `Viewers`, `Themes`, and `Plugins` off pane-local `PostMessageW(... WM_COMMAND ...)` callback bridges.
- [x] Removed pane-specific `HandleNotify(...)` methods from the migrated DxUi-first panes.
- [x] Removed Preferences dialog `WM_NOTIFY` forwarding for pane-local logic.
- [x] Removed dialog `WM_COMMAND` forwarding for the migrated DxUi-first panes.
- [x] Added focused self-tests for deferred search actions on `Keyboard`, `Viewers`, `Themes`, and `Plugins`.

Notes:

- This phase intentionally did not replace the entire pane contract yet.
- Remaining `WM_COMMAND` handling is still present for panes that continue to depend on legacy Win32 command routing (`General`, `Panes`, `Advanced`, `Compare Directories`, `Hot Paths`).

### Phase 2: Pane command-routing cleanup [completed]

Completed in code on 2026-03-25:

- [x] Removed the last pane `HandleCommand(...)` compatibility methods from `General`, `Panes`, `Advanced`, `Compare Directories`, and `Hot Paths`.
- [x] Removed the final Preferences dialog pane-command fanout from `WM_COMMAND`.
- [x] Reached zero remaining `HandleCommand(...)` pane methods under `RedSalamander/Preferences*`.

Notes:

- The fake-command bridge seam is gone.
- The next phase is the pane-contract replacement and dialog-shell cleanup.

### Phase 3: Contract and dialog-shell cleanup [completed]

- [x] Replace the pane interface everywhere.
- [x] Stop treating pane code as child-HWND owners.
- [x] Remove pane `EnsureCreated` / `Hwnd` / `ResizeToHostClient` usage from `Preferences.Dialog.cpp`.
- [x] Remove pane-host helper remnants from `Preferences.Internal.*`.
- [x] Build `Debug|x64`.
- [x] Run focused Preferences self-tests for category switching, retained state, and migrated DxUi action flows.

Why third:

- The callback-bridge cleanup already removed the highest-risk reentrancy seam, so the remaining contract cleanup can proceed pane by pane.

Next active phase:

- Phase 4: Finish low-risk panes.

### Phase 4: Finish low-risk panes [completed]

- [x] Editors
- [x] Mouse
- [x] General
- [x] Hot Paths
- [x] Remove remaining pane-local `SetWindowPos` / `ShowWindow` / `EnableWindow` fallback code from `General` and `Hot Paths`.
- [x] Trim leftover placeholder pane state from `Editors` and `Mouse`.
- [x] Build `Debug|x64`.
- [x] Run focused self-tests for `General`, `Hot Paths`, `Editors`, and `Mouse`.

Why fourth:

- These are smaller and clarify the final pattern for the larger panes.

Next active phase:

- Phase 5: Finish medium-complexity panes.

### Phase 5: Finish medium-complexity panes [completed]

- [x] Panes
- [x] Viewers
- [x] Keyboard
- [x] Themes
- [x] Plugins
- [x] Remove pane-local ListView-era column helpers from `Keyboard` and `Viewers`.
- [x] Remove stale fallback `SetWindowPos` / `ShowWindow` layout branches from `Panes` and `Plugins`.
- [x] Remove the last dead legacy layout helper from `Themes`.
- [x] Build `Debug|x64`.
- [x] Run focused self-tests for `Panes`, `Viewers`, `Keyboard`, `Themes`, and `Plugins`.

Why fifth:

- These already have substantial DxUi surfaces; most remaining work is seam removal.

Verification completed:

- `.\build.ps1 -ProjectName RedSalamander` passed on 2026-03-25.
- Focused self-tests passed:
  - `cmd_preferences_dialog_panes_page_uses_dxui_statics_and_toggles`
  - `cmd_preferences_dialog_panes_roundtrip_restores_dxui_surface`
  - `cmd_preferences_dialog_viewers_search_action_updates_dxui_surface`
  - `cmd_preferences_dialog_keyboard_page_uses_dxui_shell_chrome`
  - `cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface`
  - `cmd_preferences_dialog_themes_search_action_updates_dxui_surface`
  - `cmd_preferences_dialog_plugins_search_action_updates_dxui_surface`
  - `cmd_preferences_dialog_plugins_search_roundtrip_preserves_retained_state`

Next active phase:

- Phase 6: Finish the hybrid-heavy panes.

### Phase 6: Finish the hybrid-heavy panes [completed]

- [x] Advanced
- [x] Compare Directories
- [x] Remove the dead hidden-control fallback layout path from `Preferences.Advanced.cpp`.
- [x] Remove the dead per-control wrapper structs and fallback Win32 layout path from `Preferences.CompareDirectories.cpp`.
- [x] Build `Debug|x64`.
- [x] Run focused self-tests for `Advanced` and `Compare Directories`.

Why sixth:

- They contained the largest remaining amount of hidden Win32 layout scaffolding among the pane implementations.

Verification completed:

- `.\build.ps1 -ProjectName RedSalamander` passed on 2026-03-25.
- Focused self-tests passed:
  - `cmd_preferences_dialog_advanced_page_uses_dxui_statics_and_toggles`
  - `cmd_preferences_dialog_advanced_roundtrip_restores_dxui_surface`
  - `cmd_preferences_dialog_compare_directories_page_uses_dxui_statics`
  - `cmd_preferences_dialog_compare_directories_roundtrip_restores_dxui_surface`

Next active phase:

- Phase 7: Final infrastructure cleanup.

### Phase 7: Final infrastructure cleanup [completed]

- [x] Remove dead helpers from `Preferences.Internal.*`.
- [x] Remove `<commctrl.h>` from Preferences sources.
- [x] Remove common-control initialization from the Preferences dialog.
- [x] Re-audit `PreferencesDialogState` and rename the remaining shared shell HWND fields to `categoryTreeWindow` / `pageHostWindow`.
- [x] Build `Debug|x64`.
- [x] Build `Release|x64`.
- [x] Run focused shared Preferences self-tests after the infrastructure cleanup.

Why seventh:

- This was the final shared-shell cleanup after all pane migrations were already complete.

Verification completed:

- Static search reaches zero remaining Preferences-layer hits for `<commctrl.h>`, `ListView_*`, `TreeView_*`, `NMHDR`, and `INITCOMMONCONTROLSEX`.
- `.\build.ps1 -ProjectName RedSalamander` passed on 2026-03-25.
- `.\build.ps1 -ProjectName RedSalamander -Configuration Release` passed on 2026-03-25.
- Focused self-tests passed:
  - `cmd_preferences_dialog_page_host_uses_dxui_surface`
  - `cmd_preferences_dialog_category_tree_keyboard_navigation_updates_category`
  - `cmd_preferences_dialog_scroll_host_preserves_retained_page_state`
  - `cmd_preferences_dialog_plugins_roundtrip_restores_dxui_surface`
  - `cmd_preferences_dialog_themes_roundtrip_restores_dxui_surface`

Next active work:

- Phase 8: Manual full-Preferences verification from the regression checklist below.
- This phase is human-only. It remains open until a person drives the UI across every pane and confirms behavior, layout, scrolling, resize, and Apply/OK flows.

## Regression and Verification Plan

### Static checks

- Search for `HandleNotify(` under `RedSalamander/Preferences*` and reach zero remaining pane usages.
- Search for `PostMessageW(... WM_COMMAND ...)` under `RedSalamander/Preferences*` and reach zero remaining pane-bridge usages.
- Search for `ListView_`, `TreeView_`, `NMHDR`, and `<commctrl.h>` under `RedSalamander/Preferences*` and justify or remove every hit.
- Search for pane-interface methods (`EnsureCreated`, `Hwnd`, `CreateControls`, `LayoutControls`) and remove them from all pane headers.

### Manual test matrix

- Open Preferences on every category.
- Switch categories repeatedly with mouse and keyboard.
- Resize the window while on every complex page.
- Verify scrolling, focus, hover, wheel, combo, toggle, text, grid, and button behavior on every pane.
- Verify OK, Cancel, Apply, Reset, import/export, browse, configure, and test flows.
- Verify per-plugin configuration pages, placeholder pages, and warning cards.

### Build matrix

- `.\build.ps1 -ProjectName RedSalamander`
- `.\build.ps1 -ProjectName RedSalamander -Configuration Release`

### Exit criteria

- Preferences no longer depends on Win32 common controls for active UI behavior.
- Preferences pane implementations are DxUi-first and do not require child-HWND lifecycle plumbing.
- `Preferences.Dialog.cpp` acts as a dialog shell and host coordinator, not as a pane command/notify broker.
- `Preferences.Internal.*` no longer contains pane-child-window management helpers.

## Explicit Non-Goals

- Replacing the top-level Preferences dialog HWND.
- Removing `WindowHost` attachment HWNDs needed by DxUi.
- Refactoring unrelated non-Preferences dialogs in the same pass.
- Folding unrelated viewer/window-lifetime fixes into this plan.
