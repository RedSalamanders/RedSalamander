# Visible Native Surface Audit

Last updated: 2026-05-17

## Scope

This document tracks the remaining non-report visible-native surface references that are still intentionally outside the current validated DX replacement set.

Related audits:

- `Specs/UI/UI_VisibleComctlAudit.md`
- `Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md`

Validation helper:

- `Tools/Audit-VisibleNativeSurfaces.ps1`
- `Tools/Audit-RemainingWin32UiDependencies.ps1` for broader HFONT/GDI/native-control cleanup inventory

## Current Inventory

The narrow visible-native surface audit remains empty for the original non-report scope, but the broader remaining-Win32-UI audit now owns the HFONT/GDI/native-control closeout gate.

Latest broad closeout candidate:

```text
Archive: Specs/TestRuns/4cb089111a23/Audit/2026-05-17_2005_dxui_native_textinput_bridge_removed_final_recheck/
HDC text/selection bridge:      16 total, 16 allowed, 0 unallowed
HFONT handle:                    1 total, 1 allowed, 0 unallowed
Native visible control creation: 1 total, 1 allowed, 0 unallowed
```

Closeout gate status: `Tools/Audit-RemainingWin32UiDependencies.ps1 -FailOnFindings` exits 0 with zero unallowlisted findings.

2026-05-17 native text-input update: the remaining production DxUi text-input bridge window, bridge window messages, hidden text-service `CreateFontIndirectW` / `WM_SETFONT` path, and `LOGFONTW` debug hook/test allowlist entries have been removed from source. The archive above is the refreshed broad-audit run without those DxUi text-service font rows or bridge-message allowlist dependencies.

2026-05-17 native text-input command-line update: the FolderWindow pane command-line input no longer creates visible native `STATIC` / `EDIT` controls, no longer subclasses an edit HWND, and no longer uses `DEFAULT_GUI_FONT` / `WM_SETFONT`. It is now a `FolderWindow` DxUi host with a native-backend retained `TextField`, guarded by `cmd_pane_command_line_insertion_and_execute`.

Allowed residuals in the broad audit:

| File | Pattern | Visibility | Reason | Removal owner | Exit condition |
|------|---------|------------|--------|---------------|----------------|
| `Common/DxUi/DxUi.ComboBox.cpp`, `Common/DxUi/DxUi.Menu.cpp`, `Tests/DxUiTests/DxUiTests.Menu.cpp` | `GetDC`, `SelectObject` | visual bitmap interop / test-only | Popup backdrop capture/test bitmap work, not visible text or native control layout. | DxUi popup/menu | Replace when popup backdrop snapshots no longer require GDI-compatible bitmap capture. |
| `RedSalamander/FolderView.Icons.cpp` | `GetDC` | test-only bitmap interop | `ENABLE_TESTS` synthetic thumbnail generation uses a screen DC only to create deterministic DIB thumbnails. | FolderView thumbnail self-tests | Replace if thumbnail self-tests gain a non-HDC bitmap factory helper. |
| `RedSalamander/FolderView.Rendering.cpp`, `RedSalamander/IconCache.cpp` | `GetDC`, `SelectObject` | shell icon bitmap interop | Shell icon and shortcut overlay conversion into D2D-compatible bitmaps. | FolderView/IconCache | Replace when shell icon extraction has a pure WIC/D2D path. |
| `RedSalamander/NavigationViewInternal.h` | `SelectObject` | bitmap alpha-blend compatibility | Compatibility DIB alpha-blend fallback, not text or control layout. | NavigationView | Replace when the fallback alpha blend path is retired. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`, `Tests/DxUiTests/DxUiTestHelpers.h`, `Tests/ViewerPETests/ViewerPETests.cpp` | HDC/HFONT mentions | test-only | Pixel probe, hidden clipboard owner, and assertion text. | Test owners | Remove when equivalent test helpers avoid native probes/text. |

No unallowlisted blockers remain in the broad audit. The formerly unallowlisted app-owned HDC paint/selection seams now route through `RedSalamander/D2DHdcPaint.*`, and the last Compare Options legacy `STATIC` fallback was replaced by the custom Dx host class.

## Current Conclusion

- The non-report visible-native inventory is now empty.
- The earlier `FolderWindow` pane status-bar native class has been retired in favor of an owned window class.
- The earlier `FolderWindow` pane command-line native `STATIC` / `EDIT` input has been retired in favor of a DxUi host with a native-backend retained `TextField`.
- The broader visible-native cleanup audit has zero unallowlisted findings; remaining hits are explicit hidden/test/interop allowlist entries.
- The broader verification work is closed by `Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md`.
- The first automated broad baseline is archived at `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_182415_remaining_win32_ui_baseline/`; the latest broad audit is archived at `Specs/TestRuns/4cb089111a23/Audit/2026-05-17_2005_dxui_native_textinput_bridge_removed_final_recheck/`.
- `Win32UiHelpers.h/.cpp` are retired; surviving pure helpers now live in `RedSalamander/UiMetrics.h/.cpp`.
- ViewerWeb status drawing now uses Direct2D/DirectWrite with no visible GDI `DrawTextW` fallback; if DirectWrite status rendering cannot initialize, the status text is skipped rather than reopening a visible native text path.
- Shared native-font helpers are removed from `Common/DxUi/DxUi.Typography.h`; the former hidden DxUi text-service bridge font and LOGFONT test-hook allowlist rows are removed.
- The broad audit gate is clean, and `UI_RemainingWin32UiDependencyRetirementPlan.md` has moved to Done with fresh focused verification.
- The narrow visible-native and comctl audit scripts now use native PowerShell scanning so they are not dependent on a local `rg.exe` install.
