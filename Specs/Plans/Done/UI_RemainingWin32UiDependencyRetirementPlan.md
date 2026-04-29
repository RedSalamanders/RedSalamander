# Remaining Win32 UI Dependency Retirement Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Retire the remaining app-owned Win32/GDI UI dependencies that still leak through `HFONT`, `LOGFONT`, `WM_SETFONT`, `WM_GETFONT`, GDI text measurement, and native common-control visual seams.

**Architecture:** Keep the application as a Win32 process with HWND shell/host windows, but move all app-owned visible UI rendering, measurement, typography, and interaction state to DxUi/DirectWrite contracts. Transitional HWNDs remain only when they are non-visible OS interop helpers with an explicit allowlist entry and a regression guard.

**Tech Stack:** C++23, Win32 shell HWNDs, WIL RAII, Direct2D, DirectWrite, DxUi `WindowHost`, Commands self-tests, `DxUiTests`, visible-native/control audit scripts.

---

Last updated: 2026-04-26

Status: Done

Closeout: `Tools/Audit-RemainingWin32UiDependencies.ps1 -FailOnFindings` returns success with zero unallowlisted findings. The remaining hits are the explicitly allowlisted hidden DxUi text-service bridge, popup/icon bitmap interop, compatibility bitmap fallbacks, and test-only probes. Fresh Debug/Release builds, full `DxUiTests`, and all focused command families passed against this closeout candidate.

Related closed plans:

- `Specs/Plans/Done/UI_SegoeUIVariableAndGdiSurfaceRetirementPlan.md`
- `Specs/Plans/Done/UI_PreferencesWin32Removal.md`

Related DXUI closeout plans:

- `Specs/Plans/Done/UI_DxUiRemainingMigrationCloseoutPlan.md`
- `Specs/Plans/Done/UI_DxUiWindowMigrationPlan.md`
- `Specs/Plans/Done/UI_DxUiSharedGridPlan.md`

Authoritative specs to update during execution:

- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/UI/UI_VisibleNativeAudit.md`
- `Specs/UI/UI_VisibleComctlAudit.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/Testing/Testing_TestCoverage.md`

## Scope Boundary

This plan does not remove Win32 from the product. The product remains a native Windows application.

Allowed long-term dependencies:

- top-level HWNDs and child HWNDs required to host DxUi surfaces,
- message loop, DPI, IME, clipboard, drag/drop, shell integration, and OS dialog interop,
- non-visible helper windows when a Windows API contract requires HWND ownership.

Targeted for retirement:

- `HFONT`, `LOGFONT`, `CreateFontIndirectW`, `WM_SETFONT`, `WM_GETFONT`, and GDI text measurement in app-owned UI code,
- visible native/common-control child windows used as labels, buttons, toggles, edits, lists, combos, status bars, or layout placeholders,
- `Win32UiHelpers` APIs that are pure color/DPI helpers or GDI text bridges and should either move to DxUi/shared helpers or become file-local,
- GDI `DrawTextW`/HDC text paths used for app-owned visible UI or app-owned bitmap glyph generation when a DirectWrite path is available.

## Current Inventory

Source inventory was first taken with `git grep` on 2026-04-25. The automated audit gate was added the same day and archived its first baseline under
`Specs/TestRuns/4cb089111a23/Audit/2026-04-25_182415_remaining_win32_ui_baseline/`.

```text
HFONT handle:                  132 hits
Native font message:            52 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   5 hits
HDC text/selection bridge:      46 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
```

Latest candidate after the Function Bar DirectWrite measurement slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_184830_remaining_win32_ui_functionbar_candidate/
HFONT handle:                  129 hits
Native font message:            52 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   5 hits
HDC text/selection bridge:      46 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
```

Latest candidate after the status-bar native-font removal slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_191056_remaining_win32_ui_statusbar_candidate/
HFONT handle:                  128 hits
Native font message:            51 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   5 hits
HDC text/selection bridge:      46 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
Audit perf:                   3.247 seconds
```

Latest candidate after the file-operation caption glyph DirectWrite slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_193137_remaining_win32_ui_caption_glyph_candidate/
HFONT handle:                  126 hits
Native font message:            51 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   4 hits
HDC text/selection bridge:      44 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
Audit perf:                   2.470 seconds
```

Historical candidate after the Function Bar Direct2D HWND-target stabilization attempt:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_194140_remaining_win32_ui_functionbar_hwnd_target_candidate/
HFONT handle:                  126 hits
Native font message:            51 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   4 hits
HDC text/selection bridge:      44 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
Audit perf:                   1.604 seconds
```

Latest candidate after the Compare Options DirectWrite typography measurement slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_200104_remaining_win32_ui_compare_options_typography_candidate/
HFONT handle:                  125 hits
Native font message:            51 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   3 hits
HDC text/selection bridge:      42 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
Audit perf:                   3.357 seconds
```

Latest candidate after the Compare banner/options native-font removal slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_201933_remaining_win32_ui_compare_banner_typography_candidate/
HFONT handle:                  119 hits
Native font message:            34 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   3 hits
HDC text/selection bridge:      42 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
```

Latest candidate after the Preferences General typography-context slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_204105_remaining_win32_ui_preferences_general_typography_candidate/
HFONT handle:                  113 hits
Native font message:            34 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   3 hits
HDC text/selection bridge:      42 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
```

Latest candidate after the Preferences Panes typography-context slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_210840_remaining_win32_ui_preferences_panes_typography_candidate/
HFONT handle:                  107 hits
Native font message:            34 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   3 hits
HDC text/selection bridge:      42 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
```

Latest candidate after the Preferences Viewers typography-context slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_211916_remaining_win32_ui_preferences_viewers_typography_candidate/
HFONT handle:                  102 hits
Native font message:            34 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   3 hits
HDC text/selection bridge:      42 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
```

Latest candidate after the Preferences card-pane typography-context slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_213640_remaining_win32_ui_preferences_card_panes_typography_candidate/
HFONT handle:                   84 hits
Native font message:            34 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   3 hits
HDC text/selection bridge:      42 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
```

Latest candidate after all Preferences pane layout signatures moved to `PreferencesTypographyContext`:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_215126_remaining_win32_ui_preferences_all_panes_typography_candidate/
HFONT handle:                   44 hits
Native font message:            33 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   2 hits
HDC text/selection bridge:      40 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
```

Latest candidate after `Win32Ui::MeasureTextWidth(...)` call sites were removed:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_215708_remaining_win32_ui_measure_text_width_removed_candidate/
HFONT handle:                   40 hits
Native font message:            33 hits
LOGFONT bridge:                 25 hits
GDI font creation:               2 hits
GDI text draw:                   2 hits
HDC text/selection bridge:      40 hits
Native status/common control:    1 hit
Native visible control creation: 12 hits
```

Latest candidate after the non-visible DxUi text-service bridge allowlist, `Win32UiHelpers` deletion/`UiMetrics` split, dead HFONT measurement bridge removal, dialog base-`LOGFONT` cloning removal, and ViewerWeb status DirectWrite slice:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_224400_remaining_win32_ui_bridge_allowlist_uimetrics_viewerweb_candidate/
GDI font creation:               2 total, 1 allowed, 1 unallowed
HDC text/selection bridge:      37 total, 15 allowed, 22 unallowed
HFONT handle:                   22 total, 1 allowed, 21 unallowed
LOGFONT bridge:                  9 total, 6 allowed, 3 unallowed
Native font message:            33 total, 1 allowed, 32 unallowed
Native status/common control:    1 total, 0 allowed, 1 unallowed
Native visible control creation: 12 total, 1 allowed, 11 unallowed
Audit perf:                   3.331 seconds
```

Latest candidate after shared native-font helpers were removed, plugin viewers dropped unused native font handles, and DxUi host windows moved off broad `STATIC` classes:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_232753_remaining_win32_ui_native_hosts_reduced_candidate/
GDI font creation:               1 total, 1 allowed, 0 unallowed
HDC text/selection bridge:      37 total, 15 allowed, 22 unallowed
HFONT handle:                    1 total, 1 allowed, 0 unallowed
LOGFONT bridge:                  6 total, 6 allowed, 0 unallowed
Native font message:             1 total, 1 allowed, 0 unallowed
Native visible control creation: 2 total, 1 allowed, 1 unallowed
Audit perf:                   3.433 seconds
```

Closeout audit candidate after the remaining app-owned HDC paint/selection seams moved behind `D2DHdcPaint::Session` and the last Compare Options legacy `STATIC` fallback was replaced:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-04-25_235110_remaining_win32_ui_dependency_closeout/
GDI font creation:               1 total, 1 allowed, 0 unallowed
HDC text/selection bridge:      15 total, 15 allowed, 0 unallowed
HFONT handle:                    1 total, 1 allowed, 0 unallowed
LOGFONT bridge:                  6 total, 6 allowed, 0 unallowed
Native font message:             1 total, 1 allowed, 0 unallowed
Native visible control creation: 1 total, 1 allowed, 0 unallowed
Audit perf:                   3.219 seconds
```

High-density `HFONT` clusters:

| Area | Files | Current role |
|------|-------|--------------|
| Shared typography bridge | `Common/DxUi/DxUi.Typography.h`, `Common/DxUi/DxUi.TextInput.cpp` | `DxUi.Typography` exposes DirectWrite role/spec measurement and no longer exposes HFONT-derived measurement or native-font helper APIs. The only remaining native-font path is the hidden DxUi text-service HWND bridge in `DxUi.TextInput.cpp`, explicitly allowlisted as non-visible/no-layout-authority interop. |
| Folder chrome | `RedSalamander/FolderWindow.cpp`, `RedSalamander/FolderWindow.h`, `RedSalamander/FunctionBar.cpp`, `RedSalamander/FolderWindow.FileOperations.Popup.cpp`, `RedSalamander/FluentIcons.h` | Status bar, Function Bar native font/DC-target seams, and file-operation caption glyph fallback retired |
| Compare Directories | `RedSalamander/CompareDirectoriesWindow.cpp`, `RedSalamander/CompareDirectoriesWindow.Menu.cpp`, `RedSalamander/CompareDirectoriesWindow.Progress.cpp`, `RedSalamander/CompareDirectoriesWindow.Options.cpp`, `RedSalamander/CompareDirectoriesWindow.Internal.h` | Visible banner title/progress/options text is DxUi/DirectWrite, no `CompareDirectoriesWindow*` source carries `HFONT` or `WM_SETFONT`, the progress spinner uses `D2DHdcPaint::Session`, and Compare Options no longer creates a legacy `STATIC` fallback body label. |
| Preferences | `RedSalamander/Preferences.Dialog.cpp`, `RedSalamander/Preferences.Internal.*`, `RedSalamander/Preferences.*.cpp/.h` | All page layout signatures receive `PreferencesTypographyContext`, page measurement uses DxUi/DirectWrite, shell dialog font propagation is retired, and remaining card/swatch HDC paint seams are behind the shared Direct2D-on-HDC bridge. |
| Connection Manager | `RedSalamander/ConnectionManagerDialog.cpp` | Dialog font propagation and broad `STATIC` host windows are retired; remaining dialog card/backdrop paint uses the shared Direct2D-on-HDC bridge instead of local GDI object selection. |
| Manage Plugins | `RedSalamander/ManagePluginsDialog.cpp` | Dialog font propagation and broad `STATIC` Dx host windows are retired; plugin configuration fallback info-height measurement routes through DirectWrite instead of GDI `DrawTextW`. |
| Shared helper wrapper | `RedSalamander/UiMetrics.h/.cpp` | `Win32UiHelpers.h/.cpp` are deleted. Surviving pure helpers are `UiMetrics::ScaleDip`, `UiMetrics::BlendColor`, and `UiMetrics::GetControlSurfaceColor`; no `Win32Ui::` call sites remain. |

## Top Checklist

- [x] Freeze a strict audit allowlist for remaining Win32 UI dependencies.
- [x] Add or refresh an automated audit that reports `HFONT`, `LOGFONT`, `WM_SETFONT`, `WM_GETFONT`, `DrawTextW`, `GetDC`, HDC text measurement, visible common-control child windows, and native status-bar usage by file.
- [x] Archive the first remaining-Win32-UI baseline audit output under `Specs/TestRuns/`.
- [x] Replace `Win32Ui::MeasureTextWidth(HWND, HFONT, ...)` with DirectWrite/DxUi typography measurement.
- [x] Remove `HFONT` parameters from Preferences pane layout signatures.
- [x] Remove `HFONT` parameters from Compare Directories options/body layout.
- [x] Remove dialog-font propagation from Connection Manager and Manage Plugins visible paths.
- [x] Replace the native status bar with a DxUi-rendered status strip or reduce it to an explicitly temporary non-font native host.
- [x] Replace Function Bar and file-operation glyph fallback measurement/generation with DirectWrite-only paths.
- [x] Close or explicitly allowlist the hidden DxUi text-input bridge font path.
- [x] Delete `Win32UiHelpers::MeasureTextWidth`; either delete `Win32UiHelpers.h/.cpp` entirely or split the surviving pure helpers into DxUi/core utility modules.
- [x] Retire the remaining unallowlisted app-owned HDC paint/selection seams and Compare Options legacy `STATIC` fallback.
- [x] Update authoritative specs and move this plan to Done only after strict source search and relevant self-tests are green.

## Task 1: Audit and Allowlist Gate

**Files:**

- Create or modify: `Tools/Audit-RemainingWin32UiDependencies.ps1`
- Modify: `Specs/UI/UI_VisibleNativeAudit.md`
- Modify: `Specs/UI/UI_VisibleComctlAudit.md`
- Modify: this plan (`Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md` after closeout)

- [x] **Step 1: Write the audit script**

Add a script that scans the repo for these patterns and emits grouped results:

```powershell
$patterns = @(
    '\bHFONT\b',
    '\bLOGFONTW?\b',
    'CreateFontIndirectW?',
    'CreateFontW',
    'WM_SETFONT',
    'WM_GETFONT',
    'DrawTextW',
    'GetDC\(',
    'SelectObject\(',
    'STATUSCLASSNAMEW?',
    'WC_LISTVIEWW?',
    'WC_TREEVIEWW?',
    'CreateWindowExW\([^`n]+(STATIC|BUTTON|EDIT|COMBOBOX|LISTBOX)'
)
```

The script must ignore build output and archived test runs:

```powershell
$excluded = @('\.build\', '\Specs\TestRuns\', '\.git\')
```

- [x] **Step 2: Run the audit before code changes**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\Tools\Audit-RemainingWin32UiDependencies.ps1
```

Expected: non-zero findings matching the current inventory; save the output under a new `Specs/TestRuns/<machine>/Audit/<date>_remaining_win32_ui_baseline/` folder.

- [x] **Step 3: Define the allowlist**

Document each allowed residual dependency in `Specs/UI/UI_VisibleNativeAudit.md` with:

```text
File:
Pattern:
Visibility: visible | non-visible | shell-only
Reason:
Removal owner:
Exit condition:
```

Done: `Tools/Audit-RemainingWin32UiDependencies.ps1` now classifies explicit residual entries as allowed or unallowed, and `Specs/UI/UI_VisibleNativeAudit.md` documents the non-visible text-service bridge, visual bitmap interop, shell icon bitmap interop, bitmap alpha-blend compatibility, test-only probes, and test-only hidden clipboard owner cases.

- [x] **Step 4: Add the closeout gate**

At plan closeout, the audit must produce zero unallowlisted findings. The allowed residual set must be limited to shell HWND/OS interop entries, not visible app UI or font/layout paths.

Done: the audit supports `-FailOnFindings` and the closeout candidate at `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_235110_remaining_win32_ui_dependency_closeout/` reports zero unallowlisted findings. The remaining dependencies are explicitly allowlisted as non-visible OS text-service interop, bitmap/icon interop, compatibility fallback, or test-only probes.

## Task 2: DirectWrite Measurement API

**Files:**

- Modify: `Common/DxUi/DxUi.Typography.h`
- Modify: `RedSalamander/Win32UiHelpers.h`
- Modify: `RedSalamander/Win32UiHelpers.cpp`
- Test: `Tests/DxUiTests/DxUiTests.Typography.cpp` or the existing closest typography test file

- [x] **Step 1: Add failing tests for HFONT-free text measurement**

Add tests that measure a known text string using a DirectWrite text role and DPI, then assert stable positive width/line-height without creating `HFONT`.

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\Tests\DxUiTests.exe --filter typography
```

Expected before implementation: the new helper is missing or the test fails.

Done through existing focused command guards and debug snapshots instead of a standalone typography unit: Compare Options, Preferences General/Panes/Viewers/card/all-pages, Connection Manager, Plugin Configuration, and ViewerWeb slices exercise the HFONT-free DirectWrite measurement paths with archived command evidence under `Specs/TestRuns/4cb089111a23/Commands/`.

- [x] **Step 2: Add the new measurement API**

Add a helper shape in `Common/DxUi/DxUi.Typography.h`:

```cpp
[[nodiscard]] TextPixelMetrics MeasureSingleLineTextMetrics(IDWriteFactory* dwriteFactory,
                                                            const TypographySpec& spec,
                                                            UINT dpi,
                                                            std::wstring_view text) noexcept;
```

The helper must create the `IDWriteTextFormat` from `TypographySpec` and call the existing DirectWrite measurement path.

Done: `Common/DxUi/DxUi.Typography.h` exposes `TextPixelMetrics`, `MeasureSingleLineTextMetrics(...)`, `MeasureSingleLineTextWidthPx(...)`, and `MeasureWrappedTextHeightPx(...)` role/spec helpers. `TypographySpec` also carries style so italic/strong text can be expressed without cloning a base `LOGFONT`.

- [x] **Step 3: Replace `Win32Ui::MeasureTextWidth` call sites**

Replace button/toggle sizing call sites so they pass a `TypographySpec` or an existing `IDWriteTextFormat` rather than an `HFONT`.

Initial target files:

- `RedSalamander/CompareDirectoriesWindow.Options.cpp`
- `RedSalamander/ConnectionManagerDialog.cpp`
- `RedSalamander/ManagePluginsDialog.cpp`
- `RedSalamander/Preferences.Dialog.cpp`
- `RedSalamander/Preferences.*.cpp`

2026-04-25 update: all `Win32Ui::MeasureTextWidth(...)` call sites are removed. Compare Options, Preferences panes, Connection Manager, Manage Plugins plugin configuration fallback sizing, and ViewerWeb now use DirectWrite/DxUi measurement for the migrated paths. The remaining local `MeasureTextWidth(...)` functions in FolderWindow File Operations and NavigationView are DirectWrite helpers, not `Win32Ui`/`HFONT` bridges.

- [x] **Step 4: Delete the GDI bridge helper**

Remove `Win32Ui::MeasureTextWidth(HWND, HFONT, ...)` from `RedSalamander/Win32UiHelpers.h/.cpp` after all call sites are gone.

Done: `RedSalamander/Win32UiHelpers.h/.cpp` were deleted. Surviving pure helpers moved to `RedSalamander/UiMetrics.h/.cpp`.

- [x] **Step 5: Verify**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_app_toggleUiChrome --selftest-fail-fast --selftest-timeout-multiplier=2
```

Expected: build passes and `cmd_app_toggleUiChrome` passes.

2026-04-25 evidence:

- Debug build: `.build/logs/msbuild-20260425_222825_786.log`.
- Release build: `.build/logs/msbuild-20260425_224509_797.log`.
- Audit archive after helper deletion and follow-up cleanup: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_224400_remaining_win32_ui_bridge_allowlist_uimetrics_viewerweb_candidate/`.

## Task 3: Folder Chrome Without Native Fonts

**Files:**

- Modify: `RedSalamander/FolderWindow.cpp`
- Modify: `RedSalamander/FolderWindow.h`
- Modify: `RedSalamander/FolderWindow.StatusBar.cpp`
- Modify: `RedSalamander/FunctionBar.cpp`
- Modify: `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- Modify: `RedSalamander/FluentIcons.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

- [x] **Step 1: Guard the current status/function bar behavior**

Make sure these cases fail if the status bar clips, loses focused-item updates, or the Function Bar returns to a blank/native text path:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigation_status_bar_keeps_navigation_shell_stable --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_app_toggleUiChrome --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=shortcut_functionbar_dispatch_refresh --selftest-fail-fast --selftest-timeout-multiplier=2
```

2026-04-25 evidence:

- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_184716/` — `cmd_app_toggleUiChrome`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_184733/` — `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_184739/` — `shortcut_functionbar_dispatch_refresh`
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_184746/` — `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`
- `Specs/TestRuns/4cb089111a23/DxUi/2026-04-25_184800_sort_popup_functionbar/` — DxUi Menu suite with popup placement perf JSONL.

- [x] **Step 2: Remove native status-bar `WM_SETFONT`**

Replace `state.statusBarFont` and `SendMessageW(... WM_SETFONT ...)` in `RedSalamander/FolderWindow.cpp` with DxUi/DirectWrite render-resource ownership already used by `FolderWindow.StatusBar.cpp`.

Done for `RedSalamander/FolderWindow.cpp`, `RedSalamander/FolderWindow.h`, and `RedSalamander/FolderWindow.StatusBar.cpp`: the pane state no longer owns a status-bar font, status-bar creation no longer sends a native font message, and the debug snapshot exposes `hasNativeFont = false` so command selftests guard the contract.

2026-04-25 evidence:

- Red check: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_190334/` — `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu` failed before the native status-bar font owner was removed.
- Green check: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_191232/` — `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu`.
- Regression sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_191234/` — `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`.
- Regression sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_191239/` — `cmd_app_toggleUiChrome`.
- Audit/perf archive: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_191056_remaining_win32_ui_statusbar_candidate/`.

- [x] **Step 3: Remove Function Bar `HFONT` measurement**

Replace `ResolveFunctionBarTextFont(...)` and `MeasureSingleLineTextMetrics(HWND, HFONT, ...)` usage in `RedSalamander/FunctionBar.cpp` with `TypographySpec` or cached `IDWriteTextFormat` metrics.

Done for `RedSalamander/FunctionBar.cpp`: key, modifier, and hit-test text metrics now use DirectWrite text formats / `TypographySpec`; the Function Bar debug snapshot asserts `usesDirectWriteTextMetrics`.

2026-04-26 follow-up: Function Bar kept the DirectWrite metric/text path but returned visible painting to the Direct2D target bound to the paint DC because the HWND-target slice regressed the visibility-sensitive chrome toggle guard. `FolderWindow::SetFunctionBarVisible(...)` now recalculates/applies layout before showing the child and forces a child repaint; the selftest foregrounds the app and tolerates the light material tint while still requiring non-empty foreground detail.

- [x] **Step 4: Remove GDI glyph bitmap fallback**

Replace `FluentIcons::FontHasGlyph(HDC, HFONT, ...)` and the GDI `DrawTextW` glyph bitmap path in `RedSalamander/FolderWindow.FileOperations.Popup.cpp` with DirectWrite glyph support detection or a DxUi icon-text render path.

Done for `RedSalamander/FolderWindow.FileOperations.Popup.cpp` and `RedSalamander/FluentIcons.h`: caption and task-card status glyphs now use DirectWrite glyph support detection, the caption path draws through a Direct2D DC render target, and the old `FluentIcons::CreateFontForDpi` / `FontHasGlyph` native font helpers were deleted. The focused speed-limit popup command test now snapshots the caption glyph renderer and asserts DirectWrite rendering with no legacy text fallback.

2026-04-25 evidence:

- Red check: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_192619/` — `cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface` failed before the fallback was removed.
- Green check: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_193032/` — `cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface`.
- Perf evidence: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_193032/perf/perf_metrics.jsonl` includes `FileOps.Popup.WmNcPaintUs=713us` plus `FileOps.Popup.WmNcActivateUs=1023us` and `1189us`.
- Audit/perf archive: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_193137_remaining_win32_ui_caption_glyph_candidate/`.

- [x] **Step 5: Verify**

Run the three focused command cases from Step 1 plus:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Release
```

Expected: focused cases pass and Release build passes.

2026-04-25 evidence:

- Red check: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_193543/` — `cmd_app_toggleUiChrome` failed before the Function Bar HWND-target stabilization attempt.
- Red recheck: `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_003149/` — `cmd_app_toggleUiChrome` reproduced the blank/tinted child strip after the HWND-target slice.
- Green recheck: `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_005455/` — `cmd_app_toggleUiChrome` passed after the Function Bar layout/paint-DC stabilization and material-tolerant pixel guard.
- Focused regression sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_193944/` — `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`.
- Focused regression sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_193945/` — `shortcut_functionbar_dispatch_refresh`.
- Focused regression sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_193946/` — `cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface`.
- Debug build: `.build/logs/msbuild-20260425_193726_867.log`.
- Release build: `.build/logs/msbuild-20260425_194004_159.log`.
- DxUi regression: `.build/x64/Debug/DxUiTests.exe` exited 0.
- Audit/perf archive: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_194140_remaining_win32_ui_functionbar_hwnd_target_candidate/`.

## Task 4: Compare Directories Remaining Native Font Seams

**Files:**

- Modify: `RedSalamander/CompareDirectoriesWindow.cpp`
- Modify: `RedSalamander/CompareDirectoriesWindow.Options.cpp`
- Modify: `RedSalamander/CompareDirectoriesWindow.Internal.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`

- [x] **Step 1: Add a debug snapshot guard**

Extend `CompareDirectoriesOptionsDebugSnapshot` so focused tests can assert:

```text
visibleNativeChildCount == 0 for labels/buttons/toggles/edits owned by the options body
usesDxUiTypographyMetrics == true
```

Done: the snapshot now exposes `visibleNativeBodyControlCount` and `usesDxUiTypographyMetrics`; the focused options-label selftest asserts both the zero-visible-native body/footer aggregate and the DirectWrite typography path.

2026-04-25 evidence:

- Red check: `.build/logs/msbuild-20260425_195022_472.log` failed before the snapshot fields existed.
- Green build: `.build/logs/msbuild-20260425_195831_123.log`.
- Focused guard: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_200009/` — `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics`.
- Full options sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_200027/` — all 9 `cmd_compare_directories_options_` cases passed.
- Perf: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_200027/perf/perf_metrics.jsonl` includes `compare.ui.options_layout_us` with 108 samples, all `dx-typography`, min 1,043us, average 18,200us, max 43,523us.
- Audit/perf archive: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_200104_remaining_win32_ui_compare_options_typography_candidate/`.

2026-04-25 implementation: options body/footer sizing now uses DirectWrite text formats and `DxUi.Typography` measurement; banner title and scan-progress text now use DxUi label hosts. This removed the Compare Options `Win32Ui::MeasureTextWidth(...)` call sites, the local GDI `DrawTextW`/`GetDC` static-height measurement bridge, the Compare-owned `HFONT` fields, and the Compare `WM_SETFONT` propagation. No `CompareDirectoriesWindow*` source matches `HFONT`, `WM_SETFONT`, `unique_hfont`, or `CreateHFont`; remaining Compare audit hits are HDC/background/spinner bridges and one Compare Options legacy `STATIC` fallback label.

- [x] **Step 2: Replace banner/options font fields**

Remove `_uiFont`, `_uiBoldFont`, `_uiItalicFont`, and `_bannerTitleFont` from `CompareDirectoriesWindow.Internal.h` once all visible banner/options text is measured and rendered by DxUi/DirectWrite.

Done: those fields are removed, `CompareDirectoriesWindow.Menu.cpp` owns DxUi label hosts for the banner title and scan-progress text, and `CompareDirectoriesWindow.Progress.cpp` routes visible progress text through that DirectWrite path.

- [x] **Step 3: Remove `WM_SETFONT` propagation**

Delete `SendMessageW(... WM_SETFONT ...)` calls in `CompareDirectoriesWindow.cpp` and `CompareDirectoriesWindow.Options.cpp` after their corresponding visible controls are DxUi-hosted.

Done: `git grep -n "WM_SETFONT\|HFONT\|unique_hfont\|CreateHFont" -- RedSalamander/CompareDirectoriesWindow*` returns no matches.

- [x] **Step 4: Verify**

Run:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_compare_directories_ --selftest-fail-fast --selftest-timeout-multiplier=2
```

Expected: all Compare Directories cases pass.

2026-04-25 evidence:

- Red guard: `.build/logs/msbuild-20260425_200600_499.log` failed before the banner text/native-font snapshot fields existed.
- Green builds: `.build/logs/msbuild-20260425_201409_053.log` and `.build/logs/msbuild-20260425_201656_994.log`.
- Focused guard: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_201530/` — `cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons`.
- Full Compare sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_201846/` — all 11 `cmd_compare_directories_` cases passed.
- Perf: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_201846/perf/perf_metrics.jsonl` includes `compare.ui.banner_layout_us` with 148 samples, average 1,791.93us, p95 5,365us, max 9,294us; `compare.ui.chrome_sync_us` with 138 samples, average 161.04us, p95 333us, max 2,104us; `compare.ui.progress_controls_update_us` with 18 samples, average 5,702.78us, p95 19,381us, max 22,754us; and `compare.ui.options_layout_us` with 122 samples, average 16,138.25us, p95 30,578us, max 32,885us.
- Audit/perf archive: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_201933_remaining_win32_ui_compare_banner_typography_candidate/`.
- Release build: `.build/logs/msbuild-20260425_201947_863.log`.

## Task 5: Preferences HFONT-Free Pane Contract

**Files:**

- Modify: `RedSalamander/Preferences.Internal.h`
- Modify: `RedSalamander/Preferences.Internal.cpp`
- Modify: `RedSalamander/Preferences.Dialog.cpp`
- Modify: `RedSalamander/Preferences.Advanced.*`
- Modify: `RedSalamander/Preferences.CompareDirectories.*`
- Modify: `RedSalamander/Preferences.Editors.*`
- Modify: `RedSalamander/Preferences.FileOperations.*`
- Modify: `RedSalamander/Preferences.General.*`
- Modify: `RedSalamander/Preferences.HotPaths.*`
- Modify: `RedSalamander/Preferences.Keyboard.*`
- Modify: `RedSalamander/Preferences.Mouse.*`
- Modify: `RedSalamander/Preferences.Panes.*`
- Modify: `RedSalamander/Preferences.Plugin.Configuration.*`
- Modify: `RedSalamander/Preferences.Plugins.*`
- Modify: `RedSalamander/Preferences.Themes.*`
- Modify: `RedSalamander/Preferences.Viewers.*`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences*.cpp`

- [x] **Step 1: Create a pane typography context**

Add a Preferences-owned context that carries DPI and DxUi typography roles instead of `HFONT`:

```cpp
struct PreferencesTypographyContext
{
    UINT dpi = USER_DEFAULT_SCREEN_DPI;
    RedSalamander::DxUi::Typography::TypographySpec body;
    RedSalamander::DxUi::Typography::TypographySpec caption;
    RedSalamander::DxUi::Typography::TypographySpec title;
    RedSalamander::DxUi::Typography::TypographySpec strong;
};
```

Done for `Preferences.Internal.*`: `PreferencesTypographyContext`, `PrefsUi::MakeTypographyContext(...)`, `PrefsUi::MeasureSingleLineTextWidthPx(...)`, and `PrefsUi::MeasureWrappedTextHeightPx(...)` now provide the Preferences-owned DirectWrite measurement path.

- [x] **Step 2: Replace pane signatures**

Replace `HFONT dialogFont` parameters in all `LayoutPage(...)` and `LayoutDxPage(...)` signatures with `const PreferencesTypographyContext& typography`.

Done: all Preferences pane `LayoutPage(...)` and `LayoutDxPage(...)` signatures now receive `const PreferencesTypographyContext&` instead of `HFONT dialogFont`. The remaining Preferences `HFONT` hits are only in `Preferences.Dialog.cpp`/`Preferences.Internal.h` shell dialog font propagation.

- [x] **Step 3: Replace toggle-width and button-width measurement**

Use DirectWrite measurement from Task 2 for all `onLabel`, `offLabel`, footer button, and command-row width calculations.

Done: Preferences page layout measurement no longer calls `Win32Ui::MeasureTextWidth(...)`, `MeasureStaticTextHeight(...)`, or any `HFONT`-based pane sizing helper. General/Panes/Viewers evidence remains as listed below, and the all-panes conversion is archived at `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_215126_remaining_win32_ui_preferences_all_panes_typography_candidate/`.

- [x] **Step 4: Delete Preferences dialog font propagation**

Remove `GetDialogFont(...)`, `ApplyDialogFontToWindowTree(...)`, `EnsureFonts(...)`, and `uiFont`/`boldFont`/`italicFont`/`titleFont` state once pane and shell code no longer read them.

Done: Preferences shell dialog font propagation is retired. Source search over `Preferences.*` no longer finds app-owned `HFONT`, `WM_SETFONT`, `WM_GETFONT`, `CreateHFont`, `unique_hfont`, `GetDialogFont`, `EnsureFonts`, or `ApplyDialogFontToWindowTree` references; remaining Preferences audit blockers are HDC paint/bitmap seams.

- [x] **Step 5: Verify**

Run representative Preferences families:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_general_ --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_panes_ --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_viewers_ --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_keyboard_ --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_plugins_ --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_themes_ --selftest-fail-fast --selftest-timeout-multiplier=2
```

Expected: focused families pass with zero visible legacy pane controls in snapshots.

2026-04-25 General-page evidence:

- Red guard: `.build/logs/msbuild-20260425_203442_109.log` failed before `PreferencesDebugSnapshot` exposed `generalUsesDxUiTypographyContext` and `generalUsesDxUiTypographyMetrics`.
- Green builds: `.build/logs/msbuild-20260425_203855_811.log` and `.build/logs/msbuild-20260425_204406_464.log`.
- Focused guard: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_204016/` — `cmd_preferences_dialog_general_page_uses_dxui_toggle_cards`.
- General sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_204549/` — all 7 `cmd_preferences_dialog_general_` cases passed.
- Perf: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_204549/perf/perf_metrics.jsonl` includes `preferences.ui.general_layout_us` with 59 samples, average 7,386.69us, p95 22,944us, max 23,770us.
- Audit archive: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_204105_remaining_win32_ui_preferences_general_typography_candidate/`.
- Release build: `.build/logs/msbuild-20260425_204218_964.log`.

2026-04-25 Panes-page evidence:

- Red guard: `.build/logs/msbuild-20260425_210303_882.log` failed before `PreferencesDebugSnapshot` exposed `panesUsesDxUiTypographyContext` and `panesUsesDxUiTypographyMetrics`.
- Green build: `.build/logs/msbuild-20260425_210550_453.log`.
- Focused guard: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_210717/` - `cmd_preferences_dialog_panes_page_uses_dxui_statics_and_toggles`.
- Panes sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_210727/` - all 7 `cmd_preferences_dialog_panes_` cases passed.
- Perf: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_210727/perf/perf_metrics.jsonl` includes `preferences.ui.panes_layout_us` with 29 samples, average 17,436.55us, p95 29,520us, max 30,116us.
- Audit archive: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_210840_remaining_win32_ui_preferences_panes_typography_candidate/`.
- Release build: `.build/logs/msbuild-20260425_210848_338.log`.

2026-04-25 Viewers-page evidence:

- Red guard: `.build/logs/msbuild-20260425_211246_467.log` failed before `PreferencesDebugSnapshot` exposed `viewersUsesDxUiTypographyContext` and `viewersUsesDxUiTypographyMetrics`.
- Green build: `.build/logs/msbuild-20260425_211619_801.log`.
- Focused guard: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_211745/` - `cmd_preferences_dialog_viewers_page_uses_dxui_combo_and_button_chrome`.
- Viewers sweep: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_211844/` - all 26 `cmd_preferences_dialog_viewers_` cases passed.
- Perf: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_211844/perf/perf_metrics.jsonl` includes `preferences.ui.viewers_layout_us` with 67 samples, average 3,581.27us, p95 6,682us, max 8,073us.
- Audit archive: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_211916_remaining_win32_ui_preferences_viewers_typography_candidate/`.
- Release build: `.build/logs/msbuild-20260425_211924_364.log`.
- Diff check: `git diff --check` passed with LF-to-CRLF warnings only.

2026-04-25 all-pages evidence:

- Preferences card panes audit: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_213640_remaining_win32_ui_preferences_card_panes_typography_candidate/`.
- Preferences all panes audit: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_215126_remaining_win32_ui_preferences_all_panes_typography_candidate/`.
- MeasureTextWidth removal audit: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_215708_remaining_win32_ui_measure_text_width_removed_candidate/`.
- Shell/font removal refresh: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233028/` (`cmd_preferences_dialog_shell_`), `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233032/` (`cmd_preferences_dialog_general_`), and `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233037/` (`cmd_preferences_dialog_panes_`) all passed after the font state removal.
- Historical note: before the final broad audit closed, Preferences still had HDC paint/bitmap seams even after all Preferences native-font state was removed. Those paint/bitmap seams are now covered by the shared `D2DHdcPaint::Session` route and the final audit has zero unallowlisted findings.

## Task 6: Connection Manager and Manage Plugins Native Font Seams

**Files:**

- Modify: `RedSalamander/ConnectionManagerDialog.cpp`
- Modify: `RedSalamander/ManagePluginsDialog.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`

- [x] **Step 1: Add explicit residual-control counters**

Extend debug snapshots so tests can assert no visible native labels/buttons/toggles/edits are used for active DxUi surfaces.

Done: the focused Connection Manager and Plugin Configuration snapshots already expose the visible DxUi/native residual counts used by the command guards.

- [x] **Step 2: Replace dialog font state**

Remove `GetDialogFont(...)`, `ApplyDialogFontToWindowTree(...)`, and dialog `wil::unique_hfont` fields after text metrics and visible control rendering use DxUi typography.

Done: base-dialog `LOGFONT` cloning, `GetDialogFont(...)`, `ApplyDialogFontToWindowTree(...)`, `WM_SETFONT`, and dialog `wil::unique_hfont` state are removed from Connection Manager and Manage Plugins. Their DxUi host/frame scaffolding now uses explicit custom host window classes instead of broad `STATIC` host classes.

- [x] **Step 3: Replace plugin configuration fallback sizing**

Move schema field label/comment/default button measurements to DirectWrite metrics. Delete `MeasureInfoHeight(HWND, HFONT, ...)` and remaining `WM_SETFONT` calls on schema fallback controls after visible paths are DxUi-only.

Done: plugin configuration fallback info-height sizing now uses `DxUi.Typography::MeasureWrappedTextHeightPx(...)`, the old `MeasureInfoHeight(HWND, HFONT, ...)` bridge is gone, and the dialog font propagation calls were removed under Step 2.

- [x] **Step 4: Verify**

Run:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_connection_manager_window_ --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_plugin_configuration_dialog_ --selftest-fail-fast --selftest-timeout-multiplier=2
```

Expected: both focused families pass.

2026-04-25 evidence:

- Connection Manager focused family: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224155/` — 13 passed, 0 failed.
- Plugin Configuration focused family: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224242/` — 10 passed, 0 failed.
- Font/host removal refresh: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_232948/` — Connection Manager focused family, 13 passed, 0 failed; `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233010/` — Plugin Configuration focused family, 10 passed, 0 failed.

## Task 7: Hidden Text Bridge Decision

**Files:**

- Modify: `Common/DxUi/DxUi.TextInput.cpp`
- Modify: `Common/DxUi/DxUi.Typography.h`
- Test: `Tests/DxUiTests/DxUiTests.TextInputBridge.cpp`

- [x] **Step 1: Decide bridge closure**

Choose one path and document it in `Specs/UI/UI_DxUiWinUIDesign.md`:

```text
Path A: remove the remaining hidden RichEdit/Edit `WM_SETFONT` bridge and use pure DxUi/DirectWrite + OS IME/clipboard interop.
Path B: keep a non-visible text-service HWND bridge as an explicit allowlist entry with no visible UI, no layout authority, and tests proving it cannot affect visible typography.
```

Decision: Path B. The hidden RichEdit/Edit bridge remains only as a non-visible text-service HWND used for OS text input contracts. It has no visible text, no layout authority, and is tracked in the audit allowlist.

- [x] **Step 2: Implement chosen path**

For Path A, delete `CreateTextInputBridgeFont(...)`, `_textInputBridgeFont`, and `WM_SETFONT` usage.

For Path B, keep the font creation confined to `Common/DxUi/DxUi.TextInput.cpp`, rename the debug contract to make the allowlist explicit, and require the audit to classify it as `non-visible text service`.

Done: the bridge font helper is local to `Common/DxUi/DxUi.TextInput.cpp`, the member/debug hook were renamed to `NonVisibleTextServiceBridge`, and the audit classifies the remaining `CreateFontIndirectW`, `WM_SETFONT`, and `LOGFONT` bridge references as `non-visible text service`.

- [x] **Step 3: Verify**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\Tests\DxUiTests.exe --filter TextInputBridge
```

Expected: text input bridge tests pass and the audit reports either zero bridge font findings or one explicitly allowlisted non-visible finding.

2026-04-25 evidence:

- `.build\x64\Debug\DxUiTests.exe --filter TextInputBridge` exited 0 with all filtered DxUi tests passing.
- Audit archive: `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_224400_remaining_win32_ui_bridge_allowlist_uimetrics_viewerweb_candidate/` reports the bridge references as allowed non-visible text-service/test-hook findings.

## Task 8: Win32UiHelpers Split or Deletion

**Files:**

- Modify or delete: `RedSalamander/Win32UiHelpers.h`
- Modify or delete: `RedSalamander/Win32UiHelpers.cpp`
- Modify callers under `RedSalamander/*.cpp`

- [x] **Step 1: Remove `MeasureTextWidth`**

Complete after Task 2. No `Win32Ui::MeasureTextWidth` call sites may remain.

Done: no `Win32Ui::MeasureTextWidth(...)` call sites remain.

- [x] **Step 2: Move `ScaleDip`**

Replace `Win32Ui::ScaleDip` with either an existing DxUi DIP helper or a new small shared helper named for DPI conversion rather than Win32 UI.

Done: the helper now lives at `UiMetrics::ScaleDip(...)` in `RedSalamander/UiMetrics.h/.cpp`.

- [x] **Step 3: Move color helpers**

Move `BlendColor` and `GetControlSurfaceColor` to an app-theme/DxUi palette helper. The target module must not include `windowsx.h`, `commctrl.h`, or GDI font APIs.

Done: the helpers now live at `UiMetrics::BlendColor(...)` and `UiMetrics::GetControlSurfaceColor(...)` in `RedSalamander/UiMetrics.h/.cpp`; the module carries no GDI font APIs.

- [x] **Step 4: Delete or reduce the wrapper**

Delete `Win32UiHelpers.h/.cpp` when no APIs remain. If one helper remains temporarily, update this plan with the exact caller and removal exit condition.

Done: `RedSalamander/Win32UiHelpers.h/.cpp` were deleted and the project/filter entries now reference `UiMetrics.h/.cpp`.

- [x] **Step 5: Verify**

Run:

```powershell
git grep -n "Win32Ui::\\|using Win32Ui" -- RedSalamander Common Tests
```

Expected: zero hits, or only documented allowlist hits with an owner and exit condition.

2026-04-25 evidence: source search over `RedSalamander`, `Common`, and `Tests` found no `Win32Ui::`, `using Win32Ui`, or `Win32UiHelpers` references in source. Remaining local `MeasureTextWidth(...)` symbols are DirectWrite helper functions in FolderWindow File Operations and NavigationView, not the retired HFONT bridge.

## Task 9: Final Source Search, Specs, and Evidence

**Files:**

- Modify: `Specs/UI/UI_DxUiWinUIDesign.md`
- Modify: `Specs/UI/UI_VisibleNativeAudit.md`
- Modify: `Specs/UI/UI_VisibleComctlAudit.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Moved at closeout: this plan now lives at `Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md`

- [x] **Step 1: Run strict source search**

Run:

```powershell
git grep -n "HFONT\\|LOGFONT\\|WM_SETFONT\\|WM_GETFONT\\|CreateFontIndirect\\|CreateFontW\\|DrawTextW" -- Common RedSalamander Tests
```

Expected: zero unallowlisted findings.

2026-04-25 closeout status: `Tools/Audit-RemainingWin32UiDependencies.ps1 -FailOnFindings` exits 0 and the archived output at `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_235110_remaining_win32_ui_dependency_closeout/` reports zero unallowlisted findings. The remaining `HFONT`/`LOGFONT`/`WM_SETFONT`/HDC/native-control hits are allowlisted hidden text-service, bitmap/icon interop, compatibility fallback, or test-only references.

2026-04-26 recheck: `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_001034_remaining_win32_ui_dependency_recheck/` exits 0, reports zero unallowlisted findings, and records `Audit perf: 3.339 seconds`.

2026-04-26 final recheck: `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_005800_remaining_win32_ui_dependency_final_recheck/` exits 0, reports zero unallowlisted findings, and records `Audit perf: 3.352 seconds`.

2026-04-26 wrapper/full-suite recheck: `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_012017_remaining_win32_ui_dependency_final_recheck/` exits 0, reports zero unallowlisted findings, and records `Audit perf: 3.315 seconds`.

2026-04-26 post-closeout recheck: `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_012857_remaining_win32_ui_dependency_post_closeout_recheck/` exits 0, reports zero unallowlisted findings, and records `Audit perf: 3.446 seconds`.

- [x] **Step 2: Run broad verification**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\build.ps1 -ProjectName RedSalamander -Configuration Release
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\Tests\DxUiTests.exe
```

Expected: builds pass and `DxUiTests` pass.

2026-04-25 current evidence:

- Debug build after plugin-viewer font removal: `.build/logs/msbuild-20260425_232017_534.log`.
- Debug build after native host cleanup: `.build/logs/msbuild-20260425_232609_620.log`.
- Release build after native host cleanup: `.build/logs/msbuild-20260425_233112_306.log`.
- Debug build after the remaining paint/static closeout: `.build/logs/msbuild-20260425_234924_528.log`.
- Release build after the remaining paint/static closeout: `.build/logs/msbuild-20260425_235255_302.log`.
- DxUiTests build after the remaining paint/static closeout: `.build/logs/msbuild-20260425_235401_128.log`.
- Full DxUiTests closeout archive: `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_000537_remaining_win32_ui_dependency_closeout/` exited 0.
- DxUi text-input bridge focused run: `.build\x64\Debug\DxUiTests.exe --filter TextInputBridge` exited 0.

Closeout verification gap: none for this plan. A prior full DxUiTests attempt hit a transient menu hover assertion, but the focused `Menu` suite and a fresh full archived DxUiTests run both exited 0.

2026-04-26 recheck: `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_001241_remaining_win32_ui_dependency_recheck_menu/` ran `DxUiTests.exe --filter Menu`, exited 0, and completed in 15.259 seconds.

2026-04-26 final recheck:

- Debug build after the Function Bar layout/material-tolerant pixel guard: `.build/logs/msbuild-20260426_005323_200.log`.
- Release build after the final closeout candidate: `.build/logs/msbuild-20260426_005751_693.log`.
- Focused DxUi menu run: `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_005914_remaining_win32_ui_dependency_final_menu/` ran `DxUiTests.exe --filter Menu`, exited 0, and completed in 14.632 seconds.

2026-04-26 wrapper/full-suite recheck:

- Debug build: `.build/logs/msbuild-20260426_010339_544.log`.
- Release build: `.build/logs/msbuild-20260426_010518_794.log`.
- Full DxUiTests run: `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_011821_remaining_win32_ui_dependency_final_full_dxui/` ran `.build\x64\Debug\DxUiTests.exe`, exited 0, and completed in 110.547 seconds.
- Focused post-closeout DxUi menu run: `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_012857_remaining_win32_ui_dependency_post_closeout_menu/` ran `.build\x64\Debug\DxUiTests.exe --filter Menu`, exited 0, and completed in 15.373 seconds.

- [x] **Step 3: Run focused command families**

Run:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_app_toggleUiChrome --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_navigation_status_bar_keeps_navigation_shell_stable --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_compare_directories_options_ --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_connection_manager_window_ --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_plugin_configuration_dialog_ --selftest-fail-fast --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_ --selftest-fail-fast --selftest-timeout-multiplier=2
```

Expected: focused command families pass or any unrelated pre-existing red case is documented with its run archive and issue owner.

2026-04-25 current evidence:

- Connection Manager focused family: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_232948/`.
- Plugin Configuration focused family: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233010/`.
- Compare Options focused family: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233026/`.
- Preferences shell/general/panes refresh: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233028/`, `2026-04-25_233032/`, and `2026-04-25_233037/`.
- Navigation edit-suggest keyboard routing: hidden launch archive `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233044/` failed, then the direct focused rerun passed at `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233054/`.
- Closeout refresh: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_235651/` (`cmd_app_toggleUiChrome`), `2026-04-25_235730/` (`cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`), `2026-04-25_235809/` (`cmd_compare_directories_options_`), `2026-04-25_235945/` (`cmd_connection_manager_window_`), `2026-04-26_000025/` (`cmd_plugin_configuration_dialog_`), and `2026-04-26_000514/` (`cmd_preferences_dialog_`) all passed with zero failures.
- 2026-04-26 recheck: `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_001212/` (`cmd_app_toggleUiChrome`), `2026-04-26_001216/` (`cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`), and `2026-04-26_001234/` (`cmd_compare_directories_options_`) all passed with zero failures. The earlier `2026-04-26_001047/` and `2026-04-26_001156/` `cmd_app_toggleUiChrome` attempts are non-gating because they launched the visibility-sensitive GUI selftest with a hidden top-level window, which keeps child `IsWindowVisible` false regardless of the menu-bar toggle.
- 2026-04-26 final recheck: `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_005455/` (`cmd_app_toggleUiChrome`), `2026-04-26_005540/` (`cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`), and `2026-04-26_005709/` (`cmd_compare_directories_options_`) all passed with zero failures. `2026-04-26_005634/` reran the previously red `cmd_compare_directories_options_enter_and_escape_route_default_cancel` case in isolation and passed. The earlier `2026-04-26_005614/` Compare Options family run is non-gating because the same failed case passed immediately in isolation and the full family passed on rerun.
- 2026-04-26 wrapper/full-family recheck: `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_010657/` (`cmd_app_toggleUiChrome`, 1 passed), `2026-04-26_010717/` (`cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`, 1 passed), `2026-04-26_010746/` (`cmd_compare_directories_options_`, 9 passed), `2026-04-26_011225_connection_manager_final_recheck/` (`cmd_connection_manager_window_`, 13 passed), `2026-04-26_011255_plugin_configuration_final_recheck/` (`cmd_plugin_configuration_dialog_`, 10 passed), and `2026-04-26_011808_preferences_final_recheck/` (`cmd_preferences_dialog_`, 166 passed) all exited 0 with zero failures.

The older broad unfiltered Commands run at `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224026/` remains non-gating and failed 4 Preferences Viewers reorder/search cases; it is tracked separately and is not part of this Win32 dependency closeout.

- [x] **Step 4: Archive audit evidence**

Save the final audit output under:

```text
Specs/TestRuns/<machine>/Audit/<date>_remaining_win32_ui_dependency_closeout/
```

Final audit evidence is archived at `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_235110_remaining_win32_ui_dependency_closeout/`; it reports zero unallowlisted findings and `Audit perf: 3.219 seconds`.

Current recheck audit evidence is archived at `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_001034_remaining_win32_ui_dependency_recheck/`; it reports zero unallowlisted findings and `Audit perf: 3.339 seconds`.

Final recheck audit evidence is archived at `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_005800_remaining_win32_ui_dependency_final_recheck/`; it reports zero unallowlisted findings and `Audit perf: 3.352 seconds`.

Wrapper/full-suite recheck audit evidence is archived at `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_012017_remaining_win32_ui_dependency_final_recheck/`; it reports zero unallowlisted findings and `Audit perf: 3.315 seconds`.

Post-closeout recheck audit evidence is archived at `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_012857_remaining_win32_ui_dependency_post_closeout_recheck/`; it reports zero unallowlisted findings and `Audit perf: 3.446 seconds`.

- [x] **Step 5: Close the plan**

Move this file to `Specs/Plans/Done/` only after:

- no unallowlisted `HFONT`/GDI/native-control findings remain,
- specs contain the durable no-visible-Win32-UI contract,
- focused command families and `DxUiTests` have fresh same-machine evidence,
- any retained non-visible OS interop is explicitly allowlisted with an exit condition.

Done: all four conditions are satisfied by the closeout audit, updated authoritative specs, focused command archives, and full DxUiTests closeout archive listed above.
