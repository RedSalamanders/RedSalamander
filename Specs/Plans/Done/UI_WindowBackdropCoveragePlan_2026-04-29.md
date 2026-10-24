# Window Backdrop Coverage Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` to execute this plan.

**Goal:** Ensure the persisted `ui.windowBackdrop` setting applies Acrylic, Mica, and Mica Alt consistently to every in-scope supported captioned application window, with the main folder window explicitly out of scope.

**Architecture:** Keep one shared Win32/DWM backdrop policy path, use tool-window backdrop semantics for dialogs and utility windows, keep DxUi menu/context-menu material app-rendered, and close the work by updating the authoritative specs plus deterministic selftests.

**Tech Stack:** C++23, Win32, DWM `DWMWA_SYSTEMBACKDROP_TYPE`, WIL, Direct2D/DxUi theme palette, command selftests, existing `Common::WindowBackdrop` policy helpers.

## Opening Checklist

- [x] Confirm scope: main folder window is excluded from this plan and from new acceptance checks.
- [x] Resolve the File Operations question by treating each captioned File Operations HWND as an explicit target.
- [x] Preserve the persisted `ui.windowBackdrop` schema and current Acrylic/Mica/Mica Alt setting values.
- [x] Add or adjust deterministic selftests before changing production behavior.
- [x] Apply the shared backdrop policy to File Operations captioned windows that currently only apply title-bar theming.
- [x] Apply the shared backdrop policy to newly audited FolderView utility prompts and Item Properties.
- [x] Verify Preferences and Connection windows with regression coverage, even if production code is already correct.
- [x] Verify DxUi menus and context menus use app-rendered overlay material and do not request a DWM system backdrop.
- [x] Update authoritative specs so durable behavior is not left only in this WIP plan.
- [x] Run focused tests, build checks, and `git diff --check`.
- [x] Move this plan to `Specs/Plans/Done/` after implementation, verification, and spec closeout are complete.

## Scope And Non-Goals

In scope:

- Preferences Settings window.
- Connection Manager window and connection credential prompt windows that participate in app theme policy.
- File Operations progress popup window.
- File Operations issues pane window.
- File Operations speed-limit prompt window.
- FolderView utility prompt windows: pane filter, selection mask, rename, create directory, and change case.
- Item Properties window.
- Other supported captioned tool/dialog windows discovered by the title-bar/backdrop audit, when they are app-owned and can safely use the shared tool-window backdrop policy.
- DxUi menu and context-menu visual material, as an app-rendered surface driven by the same setting.

Out of scope:

- Main folder window behavior and tests. Existing behavior may remain untouched, but this plan must not add new main-window requirements.
- Native OS dialogs not owned/rendered by RedSalamander.
- Adding DWM system backdrop to DxUi popup/menu HWNDs. These surfaces should keep system backdrop disabled and use app-rendered Acrylic/Mica/Mica Alt material.
- Changing setting names, persisted values, migration behavior, or theme JSON schema.

## File Operations Question

The File Operations question is that "File Operations window" is not a single implementation target. The codebase has multiple captioned windows that users reasonably see as File Operations UI:

- `FileOperationsPopup` in `RedSalamander/FolderWindow.FileOperations.Popup.cpp`: the progress window.
- `FileOperationsIssuesPane` in `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp`: the issues/details window.
- `FileOperationsSpeedLimitPromptWindow` in `RedSalamander/FolderWindow.FileOperations.Popup.cpp`: the modal speed-limit prompt.
- The File Operations page inside Preferences is not a standalone File Operations window; it inherits the Preferences window backdrop contract.

The problem is that the standalone File Operations captioned windows currently apply title-bar theming but do not consistently apply the shared DWM backdrop helper. When `ui.windowBackdrop` is Acrylic, Mica, or Mica Alt, the title bar can look themed while the window background misses the requested system backdrop effect. This plan treats the three standalone File Operations HWNDs above as in scope and requires tests for each one.

## Existing Contracts To Preserve

- `Common::WindowBackdrop::Resolve(...)` remains the single policy source for Default/Acrylic/Mica/Mica Alt/None resolution.
- High contrast must resolve to no system backdrop.
- Tool/dialog windows must use tool-window target semantics.
- DxUi popups and menus must keep DWM system backdrop disabled and render their material through `DxUi::ThemePalette::overlayMaterial`.
- Existing supported windows that already call `ApplyWindowBackdropTheme(...)` should keep equivalent behavior.

## Implementation Notes

- Audit 2026-04-29: Preferences, Connection Manager, credential prompts, Find Files, Manage Plugins, and Shortcuts already called the shared backdrop helper on create/theme paths.
- Audit 2026-04-29: File Operations progress popup, issues pane, and speed-limit prompt applied title-bar theming without the shared backdrop helper.
- Audit 2026-04-29: FolderView pane filter, selection mask, rename, create-directory, change-case prompts, and Item Properties are app-owned captioned utility windows and should use tool-window backdrop semantics.
- Audit 2026-04-29: DxUi menus/context menus are app-rendered overlay surfaces and should keep DWM system backdrop disabled.
- Audit 2026-04-29: main folder window calls remain excluded from implementation changes and new acceptance checks.

## Implementation Tasks

### 1. Audit Current Window Coverage

- [x] Search for every `ApplyTitleBarTheme(...)` and `ApplyWindowBackdropTheme(...)` call site.
- [x] Classify each call site as:
  - [x] in-scope supported captioned tool/dialog window,
  - [x] main window, excluded,
  - [x] native/OS window, excluded,
  - [x] DxUi popup/menu, app-rendered material only.
- [x] Record the in-scope list in the implementation notes and mirror the durable result in `Specs/UI/UI_TopLevelToolWindows.md`.
- [x] Confirm Preferences and Connection windows already call `ApplyWindowBackdropTheme(...)` on create and theme changes.
- [x] Confirm File Operations progress popup, issues pane, and speed-limit prompt are missing shared backdrop application before changing them.

### 2. Add Failing Regression Coverage First

- [x] Add or adjust a command selftest that validates supported tool/dialog windows follow `ui.windowBackdrop=Acrylic`.
- [x] Ensure the test excludes the main folder window from assertions.
- [x] Add a File Operations selftest covering:
  - [x] progress popup window receives the expected tool backdrop kind,
  - [x] issues pane window receives the expected tool backdrop kind,
  - [x] speed-limit prompt window receives the expected tool backdrop kind.
- [x] Add or preserve coverage that Preferences applies the selected backdrop to its own HWND.
- [x] Add or preserve coverage that Connection Manager applies the selected backdrop to its own HWND.
- [x] Use `Common::WindowBackdrop::TryGetAppliedWindowBackdropKind(hwnd)` or the existing equivalent test hook rather than visual pixel assertions.
- [x] Keep DxUi menu/context-menu test coverage separate: assert app-rendered overlay material changes while system backdrop remains disabled.

### 3. Centralize The Chrome Application Path Where It Reduces Drift

- [x] Evaluate adding a small helper such as `ApplyWindowChromeTheme(HWND, const AppTheme&, WindowBackdropTarget, bool active) noexcept`.
- [x] If added, implement it in `RedSalamander/AppTheme.cpp` and declare it in `RedSalamander/AppTheme.h`.
- [x] The helper should call the existing shared backdrop helper and title-bar helper without changing policy:
  - [x] `ApplyWindowBackdropTheme(hwnd, theme, target)`,
  - [x] `ApplyTitleBarTheme(hwnd, theme, active)`.
- [x] Convert only in-scope tool/dialog windows where the helper removes missed-backdrop risk.
- [x] Do not convert main-window code as part of this plan.

### 4. Apply Backdrop To File Operations Windows

- [x] Update `FileOperationsPopup` creation/theme paths to apply tool-window backdrop policy.
- [x] Update `FileOperationsPopup` activation/theme-change handling so title-bar active state remains correct.
- [x] Update `FileOperationsIssuesPane` creation/theme paths to apply tool-window backdrop policy.
- [x] Update `FileOperationsIssuesPane` activation/theme-change handling so title-bar active state remains correct.
- [x] Update `FileOperationsSpeedLimitPromptWindow` creation/theme paths to apply tool-window backdrop policy.
- [x] Update `FileOperationsSpeedLimitPromptWindow` activation/theme-change handling so title-bar active state remains correct.
- [x] Keep fallback rendering usable when the resolved backdrop is None or when DWM backdrop APIs are unavailable.

### 5. Verify Preferences And Connection Windows

- [x] Keep or normalize existing Preferences calls to `ApplyWindowBackdropTheme(...)`.
- [x] Keep or normalize existing Connection Manager calls to `ApplyWindowBackdropTheme(...)`.
- [x] Confirm live setting changes refresh these windows when they are already open.
- [x] Confirm Apply/OK persistence behavior remains governed by existing settings contracts.
- [x] Confirm Cancel/discard behavior in Preferences remains unchanged.

### 6. Keep Menus And Context Menus App-Rendered

- [x] Verify `DxUi::ThemePalette::overlayMaterial` maps `WindowBackdropType::Acrylic` to Acrylic overlay material.
- [x] Verify the same path distinguishes Mica and Mica Alt.
- [x] Keep popup/menu HWND system backdrop disabled.
- [x] Extend tests only if current DxUi menu tests do not cover Acrylic/Mica/Mica Alt material selection.

### 7. Update Normative Specs

- [x] Update `Specs/UI/UI_TopLevelToolWindows.md` with the durable rule: supported app-owned captioned tool/dialog windows apply the persisted window backdrop setting, using tool-window target semantics unless the domain spec says otherwise.
- [x] In the same spec, explicitly state that the main folder window is outside this plan's acceptance scope.
- [x] Update `Specs/UI/UI_PreferencesDialog.md` if the Preferences window preview/apply contract needs clearer wording.
- [x] Update `Specs/Core/Core_ConnectionManager.md` or the authoritative Connection Manager spec with the Connection Manager backdrop contract.
- [x] Update the File Operations authoritative spec with the three standalone File Operations HWND contracts:
  - [x] progress popup,
  - [x] issues pane,
  - [x] speed-limit prompt.
- [x] Update `Specs/UI/UI_DxUiWinUIDesign.md` if needed to preserve the rule that menus/context menus use app-rendered material, not DWM system backdrop.
- [x] Update test coverage documentation if the repo has an index for these command selftests.

### 8. Verification

- [x] Build the RedSalamander project in Debug.
- [x] Run the focused Preferences backdrop selftest.
- [x] Run the focused Connection backdrop selftest.
- [x] Run the focused File Operations backdrop selftest.
- [x] Run DxUi menu/material tests if any menu/context-menu behavior or tests changed.
- [x] Run any existing window-backdrop/theme selftest suite.
- [x] Run `git diff --check`.
- [x] Inspect the diff to confirm no main-window production behavior or new main-window acceptance assertions were introduced.

## Acceptance Criteria

- [x] With `ui.windowBackdrop=Acrylic`, the Preferences Settings window applies the effective Acrylic tool-window backdrop when supported by the OS.
- [x] With `ui.windowBackdrop=Acrylic`, the Connection Manager window applies the effective Acrylic tool-window backdrop when supported by the OS.
- [x] With `ui.windowBackdrop=Acrylic`, each in-scope standalone File Operations captioned window applies the effective Acrylic tool-window backdrop when supported by the OS.
- [x] Equivalent behavior is covered for Mica and Mica Alt through either direct tests or shared policy tests.
- [x] High contrast still disables system backdrop.
- [x] DxUi menus and context menus visually follow the selected material while keeping DWM system backdrop disabled.
- [x] No new acceptance requirement is added for the main folder window.
- [x] Normative specs describe the durable behavior.

## Verification Evidence

- Debug build passed: `.build/logs/msbuild-20260429_195854_228.log`.
- Focused command selftests passed:
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_200019/` — `cmd_pane_fileops_issues_pane_uses_dxui_host_without_visible_child_controls`.
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_200025/` — `cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface`.
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_200030/` — `cmd_preferences_dialog_general_window_backdrop_apply_updates_supported_windows`.
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_200035/` — `cmd_connection_manager_window_applies_selected_tool_backdrop`.
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_200042/` — `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible`.
  - `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_200146/` — `settings_ui_customization_roundtrip`.
- Perf evidence archived with the speed-limit/progress-popup run:
  - `FileOps.InfoTask.EnsurePopupVisibleUs`: count 1, 61,164 us.
  - `FileOps.InfoTask.EnsurePopupVisible.CreateUs`: count 1, 18,687 us.
  - `FileOps.InfoTask.EnsurePopupVisible.RedrawWindowUs`: count 1, 35,388 us.
  - `FileOps.Popup.Render.TotalUs`: count 4, avg 6,705.5 us, max 8,823 us.
  - `FileOps.Popup.WmPaintUs`: count 4, avg 13,648.2 us, max 34,817 us.
- `git diff --check` passed with only line-ending normalization warnings.
- DxUi menu/context-menu behavior and tests were not changed; the existing `UI_DxUiWinUIDesign.md` contract and `DxUi.Menu` snapshot/test coverage already keep popup HWND system backdrop disabled while exercising distinct Mica/Mica Alt/Acrylic app-rendered materials.
- Diff inspection found no `RedSalamander/RedSalamander.cpp` changes and no added main-window backdrop acceptance assertions; removed assertions are only the old Preferences test checks against the main folder window.

## Risks And Mitigations

- Risk: DWM backdrop availability varies by Windows version.
  Mitigation: Tests should assert the stored attempted/applied policy through existing helper hooks and allow documented OS fallback semantics.

- Risk: Applying backdrop during every activation message creates unnecessary DWM calls.
  Mitigation: Apply backdrop on create/theme-change paths and title-bar active state on activation, unless the shared helper is confirmed idempotent and cheap.

- Risk: File Operations test setup can become flaky because operations are asynchronous.
  Mitigation: Reuse existing File Operations selftest harness helpers and wait predicates instead of adding sleeps.

- Risk: Menus/context menus are mistaken for DWM-backed windows.
  Mitigation: Keep an explicit spec and test distinction between top-level captioned windows and app-rendered DxUi popup surfaces.

## Self-Review Notes

- [x] The plan starts with a checklist.
- [x] The main window is explicitly out of scope.
- [x] The File Operations ambiguity is answered with concrete HWND targets.
- [x] Tests precede implementation.
- [x] Normative spec updates are required before closeout.
- [x] Verification includes focused tests and diff inspection.
