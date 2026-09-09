# Visible Native Surface Contract

Last verified: 2026-08-04

## Scope

This document tracks the remaining non-report visible-native surface references that are still intentionally outside the current validated DX replacement set.

Related audits:

- `Specs/UI/UI_VisibleComctlAudit.md`
- `Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md`

Validation helper:

- `Tools/Audit-VisibleNativeSurfaces.ps1`
- `Tools/Audit-RemainingWin32UiDependencies.ps1` for broader HFONT/GDI/native-control cleanup inventory

## Current Inventory

The narrow visible-native surface audit remains empty for the original non-report scope. The broader remaining-Win32-UI audit owns the HFONT/GDI/native-control gate and must exit successfully with zero unallowlisted findings:

```powershell
pwsh -File Tools/Audit-RemainingWin32UiDependencies.ps1 -FailOnFindings
```

The FolderWindow pane command-line input is absent; `cmd/pane/bringCurrentDirToCommandLine` inserts the current directory through the opposite-pane Terminal. AlertOverlay footer buttons and close affordances use shared DxUi chrome. Compare Directories and Manage Plugins may retain hidden compatibility controls, but focused selftests require their active surfaces to use DxUi and expose zero visible legacy controls.

Allowed residuals in the broad audit:

| File | Pattern | Visibility | Reason | Removal owner | Exit condition |
|------|---------|------------|--------|---------------|----------------|
| `Common/DxUi/DxUi.cpp`, `Common/DxUi/DxUi.ComboBox.cpp`, `Common/DxUi/DxUi.Menu.cpp` | `GetDC`, `SelectObject` | visual bitmap interop | Shared/popup backdrop capture, not visible text or native control layout. | DxUi backdrop capture | Replace when popup backdrop snapshots no longer require GDI-compatible bitmap capture. |
| `RedSalamander/Ui/AlertOverlayWindow.cpp` | `GetDC`, `SelectObject` | visual bitmap interop | Modal AlertOverlay backdrop capture before Direct2D scrim composition, not visible text or native control layout. | AlertOverlayWindow | Replace when alert-overlay backdrop snapshots no longer require GDI-compatible bitmap capture. |
| `RedSalamander/FolderView.Icons.cpp` | `GetDC` | test-only bitmap interop | `ENABLE_TESTS` synthetic thumbnail generation uses a screen DC only to create deterministic DIB thumbnails. | FolderView thumbnail self-tests | Replace if thumbnail self-tests gain a non-HDC bitmap factory helper. |
| `RedSalamander/FolderView.Rendering.cpp`, `RedSalamander/IconCache.cpp` | `GetDC`, `SelectObject` | shell icon bitmap interop | Shell icon and shortcut overlay conversion into D2D-compatible bitmaps. | FolderView/IconCache | Replace when shell icon extraction has a pure WIC/D2D path. |
| `RedSalamander/NavigationViewInternal.h` | `SelectObject` | bitmap alpha-blend compatibility | Compatibility DIB alpha-blend fallback, not text or control layout. | NavigationView | Replace when the fallback alpha blend path is retired. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp`, `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`, `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`, `Plugins/ViewerVLC/ViewerVLC.cpp`, `Tests/DxUiTests/DxUiTestHelpers.h`, `Tests/DxUiTests/DxUiTests.Menu.cpp`, `Tests/DxUiTests/DxUiTests.WindowHost.cpp`, `Tests/ViewerPETests/ViewerPETests.cpp` | HDC/HFONT/native probe mentions | test-only | Directed-input warning, pixel/backdrop probes, hidden message-queue/clipboard windows, synthetic VLC child, and assertion text. | Test owners | Remove when equivalent test helpers avoid native probes/text. |

No unallowlisted blockers may remain in the broad audit. Generated/dependency trees such as `vcpkg_installed` are outside the product-source boundary and must not create false findings.

## Current Conclusion

- The non-report visible-native inventory is now empty.
- FolderWindow status-bar and command-line product surfaces use owned/DxUi/Terminal paths rather than native common controls.
- The broader visible-native cleanup audit has zero unallowlisted findings; remaining hits are explicit hidden/test/interop allowlist entries.
- `Win32UiHelpers.h/.cpp` are retired; surviving pure helpers now live in `RedSalamander/UiMetrics.h/.cpp`.
- ViewerWeb status drawing now uses Direct2D/DirectWrite with no visible GDI `DrawTextW` fallback; if DirectWrite status rendering cannot initialize, the status text is skipped rather than reopening a visible native text path.
- Shared native-font helpers are absent from `Common/DxUi/DxUi.Typography.h`; DxUi text services do not own a native font bridge.
- The broad audit gate must remain clean; an unavailable historical archive is not part of this contract.
- The narrow visible-native and comctl audit scripts now use native PowerShell scanning so they are not dependent on a local `rg.exe` install.
