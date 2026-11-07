# Visible Native Surface Contract

Last reviewed: 2026-10-31

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
| `RedSalamander/Ui/AlertOverlayWindow.cpp` | `GetDC`, `SelectObject` | visual bitmap interop | Modal AlertOverlay backdrop capture before Direct2D scrim composition, not visible text or native control layout. | AlertOverlayWindow | Replace when alert-overlay backdrop snapshots no longer require GDI-compatible bitmap capture. |
| `RedSalamander/FolderView.Icons.cpp` | `GetDC` | test-only bitmap interop | `ENABLE_TESTS` synthetic thumbnail generation uses a screen DC only to create deterministic DIB thumbnails. | FolderView thumbnail self-tests | Replace if thumbnail self-tests gain a non-HDC bitmap factory helper. |
| `RedSalamander/FolderView.Rendering.cpp`, `RedSalamander/IconCache.cpp` | `GetDC`, `SelectObject` | shell icon bitmap interop | Shell icon and shortcut overlay conversion into D2D-compatible bitmaps. | FolderView/IconCache | Replace when shell icon extraction has a pure WIC/D2D path. |
| `RedSalamander/NavigationViewInternal.h` | `SelectObject` | bitmap alpha-blend compatibility | Compatibility DIB alpha-blend fallback, not text or control layout. | NavigationView | Replace when the fallback alpha blend path is retired. |
| `RedSalamander/SelfTest/Commands/Commands.Tests.Runner.cpp`, `RedSalamander/SelfTest/Commands/Commands.Tests.ViewCommands.cpp`, `RedSalamander/SelfTest/Commands/Commands.Tests.Search.cpp`, `Plugins/ViewerVLC/ViewerVLC.cpp`, `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp` | HDC/HFONT/native probe mentions | test-only | Directed-input warning, pixel/backdrop probes, hidden message-queue/clipboard windows, synthetic VLC child, and assertion text. | Test owners | Remove when equivalent test helpers avoid native probes/text. |
| `RedSalamander/SelfTest/Commands/Commands.Tests.Search.cpp` | exact hidden `STATIC` parent with `HWND_MESSAGE` and zero style | test-only message-only parent | The private-thread Find queue-order probe owns a parent and child HWND without desktop activation; it does not create a visible control. | Find dialog command self-tests | Replace when a non-HWND message-queue probe provides equivalent ordering coverage. |
| `RedSalamander/FolderWindow.FileOperations.Popup.cpp` | `CreateFontIndirectW`, `HFONT` system-font fallback, and `WM_SETFONT` | production emergency failure surface | Only `EnterFailureSurface` uses native fonts after the DxUi host is absent or repeated frame-preparation failures cross the declared threshold; the emergency error text and Cancel All/Close recovery actions remain readable. | File Operations popup failure surface | Replace when the failure surface uses a supported DxUi/DirectWrite diagnostic and recovery surface without native font fallback. |
| `RedSalamander/SelfTest/Commands/Commands.Tests.FileOps.cpp` | `GetDC(nullptr)` at clipboard bitmap serialization | test-only clipboard bitmap serialization | `CF_BITMAP` pixels are copied with `GetDIBits` into the bounded owned byte snapshot; the DC does not draw text or controls. | Commands File Operations clipboard self-tests | Replace when a bitmap-only snapshot helper preserves the byte-level clipboard contract. |
| `RedSalamander/SelfTest/Commands/Commands.Tests.Dialogs.cpp` | `drainWindow` `HWND_MESSAGE` | test-only message-only payload receiver | A real payload registry window proves Change Case failed-post cleanup and payload draining. | Commands Change Case payload self-tests | Replace when the payload registry has an injectable non-HWND dispatch test surface. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.PathDetailFocus.cpp` | hidden `ownershipChild` | test-only hidden focus-ownership child | A real child HWND supplies ancestry to the focus policy matrix without focus or foreground writes. | File Operations path-detail focus self-tests | Replace when ownership can be exercised with a non-HWND ancestry probe. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Phases10_13.cpp` | `c1DeferredCloseTarget` `HWND_MESSAGE` | test-only message-only close receiver | The message-only target receives and destroys on `WM_CLOSE` to prove C1 deferred application-close ordering. | File Operations C1 deferred-close self-test | Replace when the close-order contract has a non-HWND message-target seam. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.QueuedAdmission.cpp` | `tokenWindow` `HWND_MESSAGE` | test-only message-only payload registry window | A real registry token is retired and replayed as stale during queued admission. | File Operations queued-admission self-tests | Replace when the stale-token contract runs through an injectable payload registry without an HWND. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Runner.cpp` | `Owned completion teardown fixture` `WS_POPUP` | test-only hidden FolderWindow owner | The hidden popup is the real parent for the secondary FolderWindow teardown regression; it is created without `WS_VISIBLE`. | `RunFileOpsCompletionOwnerDestroyChecks` self-test | Replace when the ownership contract can create the FolderWindow against a non-window host fixture. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Runner.cpp` | `target` deferred-close `HWND_MESSAGE` | test-only message-only deferred-close receiver | The target receives and destroys on `WM_CLOSE` to witness deferred-exit refusal ordering without visible product UI. | `RunFileOpsDeferredExitRefusesNewWorkChecks` self-test | Replace when deferred exit-close ordering has a non-HWND close-target seam. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Runner.cpp` | `closeTarget` `HWND_MESSAGE` | test-only message-only compare-close receiver | The target receives `WM_CLOSE` to verify Compare-window close-consent ordering without visible product UI. | `RunFileOpsCompareWindowCloseConsentChecks` self-test | Replace when Compare-window close-consent ordering has a non-HWND close-target seam. |
| `RedSalamander/Ui/AlertOverlayWindow.cpp` | `foreign HWND reuse witness` | test-only hidden UIA identity probe | A real foreign HWND proves a stale AlertOverlay provider rejects a reused handle. | AlertOverlay UI Automation identity self-tests | Replace when provider identity reuse has a deterministic non-HWND handle-identity seam. |
| `RedSalamander/Ui/AlertOverlayWindow.cpp` | visible `createAnchor` `WS_POPUP` | test-only visible tracking anchor | The tracking matrix exercises real owner geometry and visible-window transitions. | AlertOverlay tracking self-tests | Replace when an owned application test window provides equivalent geometry and lifecycle coverage. |
| `RedSalamander/Ui/AlertOverlayWindow.cpp` | hidden `hiddenTarget` `WS_CHILD` | test-only hidden focus-return target | A real HWND verifies focus-return fallback and ownership predicates without making the target visible. | AlertOverlay modeless focus self-tests | Replace when target liveness and ownership use a non-HWND identity seam. |
| `RedSalamander/Ui/AlertOverlayWindow.cpp` | visible focus-return `target` `WS_CHILD \| WS_VISIBLE` | test-only visible focus-return target | The blocking-alert test captures a real child HWND and verifies restoration after dismissal. | AlertOverlay modeless focus self-tests | Replace when an owned application child supports equivalent visibility and focus observations. |
| `Tests/PluginContractTests/PluginContract.Tests.Runner.cpp` | hidden `Terminal` parent `WS_OVERLAPPED` | test-only hidden Terminal parent window | A valid parent is passed to `Terminal.Open` while resource-owner activation and unsupported-location behavior are checked. | Terminal plugin contract tests | Remove when this diagnostic context may be parentless. |
| `Tests/ProductUiTests/ProductUi.Tests.Runner.cpp` | `messageWindow` `HWND_MESSAGE` | test-only message-only pump window | A real thread queue and HWND message validate `PumpMessagesUntil` dispatch without visible UI. | ProductUi message-pump contract tests | Replace when an injected queue has equivalent dispatch behavior. |

No unallowlisted blockers may remain in the broad audit. Generated/dependency trees such as `vcpkg_installed` are outside the product-source boundary and must not create false findings.

Default broad-audit success output contains only structured findings; its human-readable category
summary uses the information stream. `-AsMarkdown` emits Markdown text. The narrow audits support
empty inventories under strict mode. Current tests also require classified nonempty findings and
an isolated unclassified-font negative control that retains the real failure diagnostic.
Residual line patterns match the shared scanner's trimmed `Line` field; an anchored pattern must
not require source indentation. The scanner's whitespace control and actual audit execution cover
that producer/consumer boundary.

## Current Conclusion

- The non-report visible-native inventory is now empty.
- FolderWindow status-bar and command-line product surfaces use owned/DxUi/Terminal paths rather than native common controls.
- The broader cleanup audit classifies hidden/test/interop dependencies and the explicit File Operations emergency recovery font surface. That production exception applies only after the hosted surface fails; it does not authorize native controls for ordinary popup content.
- `Win32UiHelpers.h/.cpp` are retired; surviving pure helpers now live in `RedSalamander/UiMetrics.h/.cpp`.
- ViewerWeb status drawing now uses Direct2D/DirectWrite with no visible GDI `DrawTextW` fallback; if DirectWrite status rendering cannot initialize, the status text is skipped rather than reopening a visible native text path.
- Shared native-font helpers are absent from the public `<DxUi/Typography.h>`; DxUi text services do not own a native font bridge.
- The broad audit gate must remain clean; an unavailable historical archive is not part of this contract.
- The narrow visible-native and comctl audit scripts now use native PowerShell scanning so they are not dependent on a local `rg.exe` install.
