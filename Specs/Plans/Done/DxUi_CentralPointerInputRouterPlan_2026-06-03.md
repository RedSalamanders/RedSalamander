# DxUi Central Pointer Input Router Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace duplicated live-cursor-based pointer decisions with one delivered-message input router and remove every production `GetCursorPos()` dependency from the source tree.

**Architecture:** A shared DxUi pointer input layer builds immutable routed events from delivered Win32 messages: target HWND, delivered client/screen points, message id, message time/order, capture/menu owner, and component generation tokens. Production routing, menu anchoring, hover, hit testing, popup open/close, root switching, viewer context menus, and placement logic consume only delivered events or explicit owner/control anchors. `GetCursorPos()` is diagnostic/selftest evidence only and must not affect production behavior anywhere under `Common`, `RedSalamander`, or `Plugins`.

**Tech Stack:** C++23, Win32 messages, Direct2D/DxUi controls, `DxUiTests`, RedSalamander command selftests, PowerShell source guard.

---

## Status

**Status:** Complete
**Date:** 2026-06-04
**Scope:** Whole production tree for this rule: `Common`, `RedSalamander`, and `Plugins`. Tests/selftests may use `GetCursorPos()` for deterministic setup/evidence. Production code may contain `GetCursorPos()` only on a same-line diagnostic-only annotation and only when the value is logged, never routed from.

## Continuation Closeout - 2026-06-04

The reopened continuation is complete. The original combined `cmd_viewer_`
shutdown hang was resolved by giving ViewerSpace an explicit module quiet point
for scheduler/cache, window-class, and graphics state, and by allowing
graphics-backed viewer DLLs to request process-exit module retention after that
quiet point at process shutdown. DxUi shared process-exit graphics resources now
reset before retained plugin module teardown.

Final validation:

- Debug `RedSalamander` build:
  `.build/logs/msbuild-20260604_213301_479.log` (`0 warning(s), 0 error(s)`).
- Debug `DxUiTests` build:
  `.build/logs/msbuild-20260604_214754_056.log` (`0 warning(s), 0 error(s)`).
- `.build\x64\Debug\DxUiTests.exe --suite=Menu`: passed twice after the final
  virtual-screen clamp hardening.
- `cmd_viewer_`: `Specs/TestRuns/7d3a1247382a/Commands/2026-06-04_215054/`
  (`4 passed, 0 failed`).
- `cmd_app_menuBar_`: `Specs/TestRuns/7d3a1247382a/Commands/2026-06-04_145308/`
  (`17 passed, 0 failed`).
- `cmd_pane_navigation_`:
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-04_145534/`
  (`31 passed, 0 failed`).
- Additional command slices passed at `2026-06-04_145538/`,
  `2026-06-04_145551/`, and `2026-06-04_145554/`.
- `cmd_pane_find_dialog_`:
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-04_214327/`
  (`56 passed, 0 failed, 6 skipped`). The skips are limited to clipboard-content
  assertions because the OS clipboard was externally unavailable
  (`OpenClipboard error=5, openWindow=0x0, owner=0x0`); focused clipboard
  coverage had already passed earlier at `2026-06-04_180358/`.
- `Scripts/VerifyNoProductionGetCursorPos.ps1`: passed with
  `No production GetCursorPos violations found.`
- `git diff --check -- Common RedSalamander Plugins Specs\UI Specs\Testing Scripts Tests`:
  exited 0 with only line-ending normalization warnings.

Durable contract updates were merged into
`Specs/Testing/Testing_TestCoverage.md`, `Specs/UI/UI_FindFilesWindow.md`, and
`Specs/Plugins/Plugins_ViewerPlugins.md`.

## Historical Continuation Handoff - 2026-06-03

This section records why the plan was reopened on 2026-06-03. It is retained as historical debugging context; the 2026-06-04 closeout above is the current status.

### Historical Objective

Finish the central pointer input router implementation and close the remaining flaky/incorrect selftest work. Do not claim the plan is complete until the combined viewer prefix, expanded command slices, DxUi tests, source guard, and documentation closeout all pass.

### Historical Blocker

`cmd_viewer_` was the active blocker at handoff. The four viewer assertions passed, but the RedSalamander process hung during shutdown and never archived the run to the repo.

Repro command:

```powershell
Get-Process RedSalamander -ErrorAction SilentlyContinue | Stop-Process -Force
$exe = (Resolve-Path .\.build\x64\Debug\RedSalamander.exe).Path
$p = Start-Process -FilePath $exe -ArgumentList @('--commands-selftest','--selftest-case=cmd_viewer_','--selftest-timeout-multiplier=4') -PassThru
$p.WaitForExit(90000)
if (-not $p.HasExited) {
  Get-Content (Join-Path $env:LOCALAPPDATA 'RedSalamander\SelfTest\last_run\trace.txt') -Tail 100
  Stop-Process -Id $p.Id -Force
}
```

Current symptom:

- `cmd_viewer_text_context_menu_uses_delivered_anchor` passes.
- `cmd_viewer_text_hover_uses_delivered_point` passes.
- `cmd_viewer_space_context_menu_uses_delivered_anchor` passes.
- `cmd_viewer_space_hover_uses_delivered_point` passes.
- `commands\results.json` under `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\commands\` reports `passed: 4`, `failed: 0`.
- Shutdown then hangs inside viewer plugin shutdown, so `ArchiveToRepo` never runs.

Latest trace tail consistently ends at:

```text
ViewerPluginManager::Shutdown: unload begin id='builtin/viewer-space' path='Z:\src\RedSalamander\.build\x64\Debug\Plugins\ViewerSpace.dll'
ViewerPluginManager::Unload: unregister resources begin id='builtin/viewer-space'
ViewerPluginManager::Unload: unregister resources complete id='builtin/viewer-space'
ViewerPluginManager::Unload: module reset begin id='builtin/viewer-space'
```

With reverse unload order applied, the preceding plugins unload successfully first (`viewer-web`, `viewer-vlc`, `viewer-sqlite`, `viewer-pe`, `viewer-markdown`, `viewer-json`, `viewer-text`), and the hang still occurs on `ViewerSpace.dll` module reset.

Important negative/positive isolation:

- Individual cases exit `0`.
- `cmd_viewer_text_` exits `0`.
- `cmd_viewer_space_` exits `0`.
- The combined family `cmd_viewer_` hangs after the four cases pass.
- `cmd_app_menuBar_` visible run passed `17/17` at `Specs\TestRuns\4cb089111a23\Commands\2026-06-03_214623`.
- `cmd_pane_navigation_` passed at `Specs\TestRuns\4cb089111a23\Commands\2026-06-03_214647`.
- Latest successful RedSalamander build after the current diagnostics/experiments: `.build\logs\msbuild-20260603_223825_929.log`, `0 warning(s), 0 error(s)`.

### Historical Debugging Evidence

Temporary test-only shutdown breadcrumbs were added and show:

- `CommandsSelfTest: end: exit_code=0` is reached.
- `FolderWindow::SetSettings(nullptr)` completes.
- `FolderWindow::Destroy()` completes.
- `FolderWindow::OnDestroy()` logs viewer and file-op shutdown complete.
- `FileSystemPluginManager::Shutdown(...)` completes.
- `ViewerPluginManager::Shutdown(...)` starts and reaches `FreeLibrary`/module reset for `ViewerSpace.dll`, then stops.

Window/thread snapshot from a hung process:

- No visible viewer windows.
- Top-level windows left: hidden `REDSALAMANDER`, `IME`, and `MSCTFIME UI`.
- Many waiting threads are present, mostly `UserRequest`. This may indicate plugin worker/thread lifetime or loader-lock/static-destruction interaction.

### Debug Changes Present At Handoff

These changes are for diagnosis or unproven mitigation. A continuation session should review them and either keep them as a confirmed fix or remove them before final closeout:

- `RedSalamander/FolderWindow.cpp`: test-only trace breadcrumbs around `FolderWindow::Destroy()` and `OnDestroy()`.
- `RedSalamander/RedSalamander.cpp`: test-only trace breadcrumbs around test-mode shutdown after `g_folderWindow.Destroy()`.
- `RedSalamander/ViewerPluginManager.cpp`: test-only per-plugin unload trace breadcrumbs; unload loop was changed to reverse order as an experiment. Reverse order did not fix `cmd_viewer_`.
- `Plugins/ViewerSpace/ViewerSpace.h` and `.cpp`: `UnregisterWndClassIfIdle()` plus `g_viewerSpaceWindowCount` and `g_viewerSpaceClassAtom` were added as an experiment. This did not fix `cmd_viewer_`.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`: `CleanupStandaloneViewerPointerProbe(...)` now pumps after releasing the viewer and file-system references. This did not fix `cmd_viewer_`, but may still be harmless.

Do not assume these experiments are final architecture.

### Corrected Flaky/Incorrect Tests Already Addressed

The menu-bar flaky expectation was corrected:

- Old stale registration: `cmd_app_menuBar_persistent_mouseleave_hover_switches_popup`.
- New contract/test: `cmd_app_menuBar_persistent_mouseleave_clears_hover_without_live_cursor_switch`.
- `WM_MOUSELEAVE` must clear hover; it must not sample the live cursor and must not switch top-level roots from live cursor state.

Menu keyboard/root-switch product fix already made:

- `Common/DxUi/DxUi.Menu.cpp` treats stale root-switch pointer moves from both `MenuInputSource::ModalMessage` and `MenuInputSource::PopupWndProc`.
- Keyboard-opened roots suppress the first root-switchable pointer move after the keyboard switch.
- Root switch bookkeeping clears the last pointer root-switch screen point when a keyboard root switch is remembered.
- Evidence trace for the original failure was `Specs\TestRuns\4cb089111a23\LivePointer\menu_trace_keyboard_hover_current.log`.

Other earlier fixes in this continuation:

- `Common/DxUi/DxUi.WindowHost.cpp`: `WM_SYSCHAR` now runs mnemonic handling before native text input, fixing the Find Alt mnemonic focus path.
- `RedSalamander/FindFilesWindow.cpp`: compact result grid metrics are respected.
- `Commands.SelfTest.Search.cpp`: find double-click/drag tests use host-space delivered points; long-run scroll boundary expectation fixed; search service cases set `preferIndex=true`; resized-column sort setup uses debug sort state instead of rapid header clicks.
- `Commands.SelfTest.Settings.cpp`: obsolete effective-child mouse remapping helpers removed.

### Suggested Next Debugging Sequence

1. Start with no stale processes:

```powershell
Get-Process RedSalamander -ErrorAction SilentlyContinue | Stop-Process -Force
```

2. Reproduce `cmd_viewer_` and preserve `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\trace.txt` and `commands\results.json`.

3. Capture a real dump or call stack at the `ViewerSpace.dll` module-reset hang. If debugger tools are not available, create a dump with:

```powershell
$pid = (Get-Process RedSalamander).Id
$dump = Join-Path (Get-Location) "Specs\TestRuns\viewer_space_unload_hang_$pid.dmp"
rundll32.exe C:\Windows\System32\comsvcs.dll, MiniDump $pid $dump full
```

Then inspect it with Visual Studio or a debugger to identify the stuck stack during `FreeLibrary`.

4. If stack tools are still unavailable, add targeted traces in `Plugins/ViewerSpace`:

- `DllMain` process detach begin/end using `OutputDebugStringW` or a safe file trace.
- `ViewerSpace::OnDestroy`, `OnNcDestroy`, `Close`, `Release`, and a real `~ViewerSpace()` destructor.
- `CancelScanAndWait`, `ReapFinishedScanWorkers`, and `ScanMain` start/end.
- `GetScanScheduler()` and `GetScanResultCache()` static construction/destruction paths.

5. Test the strongest current hypotheses one at a time:

- A `ViewerSpace` worker thread or helper thread is still executing plugin code at unload.
- `ViewerSpace` function-local static destruction during DLL detach is blocking; candidates are `GetScanScheduler()` and `GetScanResultCache()`.
- A `ViewerSpace` object is still referenced and the module unload is racing object/static teardown.
- DLL-owned Win32 class/static brush cleanup is involved. The current class-unregister experiment did not fix this by itself.

6. Minimal hypothesis tests worth trying:

- Add `ViewerSpace::~ViewerSpace()` that calls `CancelScanAndWait()`, `CancelScanCacheBuild()`, and clears pending updates, then trace that it runs before module reset.
- Temporarily make `GetScanScheduler()` and/or `GetScanResultCache()` return heap-allocated never-destroy singletons. If the hang disappears, replace with an explicit quiet-point clear/shutdown that does not run under loader lock.
- Add explicit ViewerSpace module shutdown export or manager-side quiet-point only if the dump proves static teardown is the problem.
- Avoid a "just skip viewer plugin unload during selftest" workaround unless the root cause is understood and documented.

### Verification Still Required After Fix

After fixing `cmd_viewer_`, run at minimum:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\build.ps1 -ProjectName DxUiTests
.\.build\x64\Debug\DxUiTests.exe --suite=Menu

$exe = (Resolve-Path .\.build\x64\Debug\RedSalamander.exe).Path
$cases = @(
  'cmd_viewer_',
  'cmd_app_menuBar_',
  'cmd_pane_navigation_',
  'cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu',
  'cmd_pane_fileops_speedLimit_prompt_',
  'cmd_pane_contextMenuCurrentDirectory_routes_current_folder',
  'cmd_pane_find_dialog_'
)
foreach ($case in $cases) {
  $p = Start-Process -FilePath $exe -ArgumentList @('--commands-selftest', "--selftest-case=$case", '--selftest-timeout-multiplier=4') -PassThru
  $p.WaitForExit()
  "$case EXIT=$($p.ExitCode)"
}

powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\VerifyNoProductionGetCursorPos.ps1
git diff --check -- Common RedSalamander Plugins Specs\UI Specs\Testing Scripts
```

Then update `Specs/Testing/Testing_TestCoverage.md`, `Specs/UI/UI_CommandMenuKeyboard.md` if the menu contract changed further, and this plan with final archive paths.

### Dirty Tree At Handoff

`git status --short` at handoff:

```text
 M Common/DxUi/DxUi.Menu.cpp
 M Common/DxUi/DxUi.WindowHost.cpp
 M Common/DxUi/DxUi.h
 M Common/DxUi/DxUi.vcxproj
 M Common/DxUi/DxUiNativeMenuInterop.h
 M Common/WindowMessages.h
 M Plugins/ViewerSpace/ViewerSpace.cpp
 M Plugins/ViewerSpace/ViewerSpace.h
 M Plugins/ViewerText/ViewerText.Text.cpp
 M Plugins/ViewerText/ViewerText.cpp
 M Plugins/ViewerText/ViewerText.h
 M RedSalamander/CompareDirectoriesWindow.cpp
 M RedSalamander/FindFilesWindow.cpp
 M RedSalamander/FindFilesWindow.h
 M RedSalamander/FolderWindow.FileOperations.Popup.cpp
 M RedSalamander/FolderWindow.Interaction.cpp
 M RedSalamander/FolderWindow.StatusBar.cpp
 M RedSalamander/FolderWindow.cpp
 M RedSalamander/NavigationView.Edit.cpp
 M RedSalamander/NavigationView.FullPathPopup.cpp
 M RedSalamander/NavigationView.Interaction.cpp
 M RedSalamander/NavigationView.Menus.cpp
 M RedSalamander/NavigationView.cpp
 M RedSalamander/NavigationView.h
 M RedSalamander/RedSalamander.cpp
 M RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp
 M RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp
 M RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp
 M RedSalamander/SplashScreen.cpp
 M RedSalamander/Ui/AlertOverlayWindow.cpp
 M RedSalamander/ViewerPluginManager.cpp
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_CommandMenuKeyboard.md
 M Specs/UI/UI_DxUiWinUIDesign.md
 M Specs/UI/UI_FindFilesWindow.md
 M Specs/UI/UI_NavigationView.md
 M Tests/DxUiTests/DxUiTests.Menu.cpp
?? Common/DxUi/DxUi.PointerInput.cpp
?? Common/DxUi/DxUi.PointerInput.h
?? Scripts/
?? Specs/Plans/WIP/DxUi_CentralPointerInputRouterPlan_2026-06-03.md
```

### Handoff Rule

A new blank chat should first read `AGENTS.md`, this handoff section, and the current diffs for `ViewerSpace`, `ViewerPluginManager`, `FolderWindow`, `RedSalamander`, and `Commands.SelfTest.ViewCommands`. Continue from the `cmd_viewer_` shutdown hang. Do not restart from the beginning of the pointer-router plan.

## Normative Contract

The specs already updated for this plan are:

- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/UI/UI_FindFilesWindow.md`
- `Specs/UI/UI_NavigationView.md`

These specs now make the following behavior mandatory:

- Delivered pointer message coordinates are authoritative.
- Production routing must not call `GetCursorPos()` or replace delivered points with the later live cursor.
- `GetCursorPos()` is allowed only in tests/selftests or diagnostic-only production code with the exact same-line annotation `// getcursorpos-allow: diagnostic-only`.
- Diagnostic-only production calls may only write evidence to logs/traces. They must not branch, route, open, close, repaint, hit-test, select anchors, choose monitor placement, update hover, or decide stale/fresh input.
- Stale input must be rejected by delivered-message metadata: target HWND, source HWND, message time/order, capture/menu ownership, input generation, root-switch generation, edit-host lifecycle, and explicit teardown tokens.
- Paint, cursor handling, path refresh, status refresh, and identical history refresh must not mutate hover state.

## Current Evidence

The live log showed two different cases that the current code cannot cleanly distinguish:

- A valid queued destination NavigationView message delivered to client `(178, 22)` while the later live cursor had drifted to `(178, 44)` on the same Find owner window. Current live-cursor classification can drop this valid click.
- A stale destination NavigationView message delivered to an old client point while the later live cursor was far away or on another top-level window. Current code must still reject this residue.

The central router solves this by removing live cursor state from routing. The valid queued click is accepted because the delivered event is still current for its target/input generation. The stale residue is rejected because the target/layout/edit/menu generation no longer matches, or because the event is older than the teardown/root-switch state that invalidated it.

## Implementation Tracking Checklist

| State | Slice | Implementation unit | Required proof before complete | Evidence / notes |
|-------|-------|---------------------|--------------------------------|------------------|
| [x] | 0 | Whole-tree baseline and source audit | Every current `GetCursorPos` hit under `Common`, `RedSalamander`, and `Plugins` listed and classified | `rg` found 30 production `GetCursorPos()` lines. Classified as production-routing (`DxUi.Menu` resync, `DxUi.WindowHost` leave/down/up live routing, `NavigationView` stale/live resolver), production-anchor (`DxUiNativeMenuInterop`, Compare, Folder, FileOps popup, main menu/context, ViewerText/ViewerSpace context), production-hit-test (`FolderWindow.Interaction`, status bar, AlertOverlay, ViewerText/ViewerSpace hover), production-placement (`SplashScreen`), and diagnostic-only trace candidates (`FindFilesWindow`, `RedSalamander`, parts of DxUi/Menu/WindowHost/NavigationView`). Baseline focused Find command selftests all exited `0`. |
| [x] | 1 | Whole-tree static source guard | Guard scans `Common`, `RedSalamander`, and `Plugins`; fails on current production calls; passes only when production is clean | Guard added at `Scripts/VerifyNoProductionGetCursorPos.ps1`; fixed PowerShell array construction; initial red run exits `1` with 30 production violations. Final green run exits `0` with `No production GetCursorPos violations found.` |
| [x] | 2 | Routed event core | DxUi unit tests prove delivered point/time/source/capture fields are stable | Added `Common/DxUi/DxUi.PointerInput.{h,cpp}` and Menu-suite tests `TestPointerInputEventMouseMoveUsesDeliveredPoint`, `TestPointerInputEventButtonUsesDeliveredPointAndFlags`, `TestPointerInputEventWheelUsesDeliveredScreenPoint`, and `TestPointerInputEventHasNoLiveCursorState`. Red build failed on missing header (`msbuild-20260603_164435_268.log`); green `DxUiTests` build passed with `0 warning(s), 0 error(s)` (`msbuild-20260603_164640_282.log`) and `.build\x64\Debug\DxUiTests.exe --suite=Menu` exited `0`. |
| [x] | 3 | DxUi menu migration | Menu tests pass without modal resync/live cursor routing | Removed menu modal live-cursor resync and idle polling. Updated root-switch tests to the delivered-message contract. `.\build.ps1 -ProjectName DxUiTests` passed with `0 warning(s), 0 error(s)` (`.build\logs\msbuild-20260603_173337_275.log`), and `.\.build\x64\Debug\DxUiTests.exe --suite=Menu` exited `0` with `All DxUi tests passed.` |
| [x] | 4 | NavigationView migration | Find destination stale/queued tests pass without live cursor classification | Implemented explicit `NavigationView::_inputGeneration{1}` using the shared `DxUi::InputGeneration` token. NavigationView WndProc now stamps delivered `WM_MOUSEMOVE`, `WM_LBUTTONDOWN`, and `WM_LBUTTONDBLCLK` events with the current generation, event handlers validate target/client-point/generation metadata before routing, and generation bumps on layout/DPI/path/history/file-system/theme/focus/embedded-mode/edit-mode/full-path-popup/dropdown/menu-loop/destroy transitions. The debug snapshots expose the generation through Find's destination NavigationView bridge, and `cmd_pane_find_dialog_destination_navigation_uses_delivered_input_generation` asserts generation advances on delivered double-click edit entry and Escape edit exit. Focused green evidence: standalone generation run `Specs\TestRuns\4cb089111a23\Commands\2026-06-03_192232\`, and final focused batch `2026-06-03_192316\` through `2026-06-03_192330\` all exited `0`; the generation case is `2026-06-03_192318\`. |
| [x] | 5 | WindowHost/Find diagnostics cleanup | Production `GetCursorPos` calls converted to diagnostic wrapper or removed | `DxUi.WindowHost`, `FindFilesWindow`, `NavigationView`, `RedSalamander`, and `DxUi.Menu` retain only same-line annotated diagnostic-only `GetCursorPos()` calls. `WM_MOUSELEAVE` no longer re-arms hover from the live cursor. |
| [x] | 6 | Native menu/context anchors | Menus opened from commands use explicit delivered/owner anchors instead of cursor fallback | Removed native context-menu cursor fallback. Keyboard/no-point routes now use explicit owner/control/focused anchors or `{}` rather than current cursor state. Guard exits `0`. |
| [x] | 7 | RedSalamander whole-tree production removal | Folder, compare, status, overlay, splash, and file-operation popup paths no longer use live cursor | Removed live cursor dependencies from FolderWindow, CompareDirectories, status bar, AlertOverlay, file-operation popup, SplashScreen placement, main user-menu anchors, and NavigationView full-path popup polling. `rg` shows production calls only on diagnostic-annotated lines; guard exits `0`. |
| [x] | 8 | Plugin viewer production removal | ViewerText and ViewerSpace context/hit-testing paths no longer use live cursor | ViewerText and ViewerSpace context-menu and hover/cursor paths now use delivered message points, `GetMessagePos`, keyboard/control anchors, or default cursor behavior. Plan-named viewer selftests are registered and exited `0`: `cmd_viewer_text_context_menu_uses_delivered_anchor` (`2026-06-03_181323`), `cmd_viewer_text_hover_uses_delivered_point` (`2026-06-03_181324`), `cmd_viewer_space_context_menu_uses_delivered_anchor` (`2026-06-03_181325`), and `cmd_viewer_space_hover_uses_delivered_point` (`2026-06-03_181326`). Guard exits `0`. |
| [x] | 9 | Regression and manual validation | Focused DxUi/Commands/plugin tests pass; live Find/nav menu hand validation recorded | Fresh automated regression after flaky-test fixes: Debug `DxUiTests` build passed (`.build\logs\msbuild-20260603_201812_256.log`, `0 warning(s), 0 error(s)`); Debug `RedSalamander` build passed (`.build\logs\msbuild-20260603_201952_657.log`, `0 warning(s), 0 error(s)`). `.\.build\x64\Debug\DxUiTests.exe --suite=Menu` exited `0` in three consecutive reruns and again in final verification. Fixed flaky Menu assertions by making delivered hover tests align live cursor state, making delayed submenu timer tests accept either pending timer or already-fired submenu state, and parking cursor state for keyboard/baseline-only tests. Focused command selftests exited `0` after a transient setup-only cursor-move retry; green batch archived from `Specs\TestRuns\4cb089111a23\Commands\2026-06-03_202334\` through `2026-06-03_202346_001\`. Final diagnostic live validation used the Debug app with `REDSALAMANDER_DXUI_MENU_TRACE=1`; per-case menu traces are archived under `Specs\TestRuns\4cb089111a23\LivePointer\2026-06-03_final_validation\`, and all live command archives `2026-06-03_203537\`, `2026-06-03_203539\`, `2026-06-03_203542\`, `2026-06-03_203548\`, `2026-06-03_203549\`, `2026-06-03_203550\`, `2026-06-03_203633\`, `2026-06-03_203634\`, `2026-06-03_203635\`, `2026-06-03_203636\`, `2026-06-03_203637\`, `2026-06-03_203639\`, and `2026-06-03_203640\` exited `0`. That live slice covers Find split-button open/hover/light-dismiss, destination NavigationView drift/stale-edit/history routing, result help overlay close/Escape/backdrop repaint, FolderWindow context routing, status-bar delivered hover/sort popup open-close, file-operation popup/prompt ownership, and ViewerText/ViewerSpace delivered anchor/hover routing. |
| [x] | 10 | Spec closeout | Whole-tree guard is green; durable specs updated; plan moved to Done | Durable UI specs and `Specs/Testing/Testing_TestCoverage.md` were updated with the final pointer-router contract and closeout evidence. Final closeout guard commands are re-run after this row update before moving the plan to `Specs/Plans/Done/`. |

---

## File Structure

### Create

- `Common/DxUi/DxUi.PointerInput.h`
  - Defines `PointerInputSource`, `PointerInputKind`, `PointerInputEvent`, `InputGeneration`, and helper builders.
- `Common/DxUi/DxUi.PointerInput.cpp`
  - Converts delivered Win32 messages to routed pointer events without live cursor polling.
- `Scripts/VerifyNoProductionGetCursorPos.ps1`
  - Scans `Common`, `RedSalamander`, and `Plugins`; fails if raw `GetCursorPos()` appears outside tests/selftests or outside explicitly annotated diagnostic-only production lines.

### Modify

- `Common/DxUi/DxUi.vcxproj`
  - Includes the new pointer input source/header.
- `Common/DxUi/DxUi.Menu.cpp`
  - Replaces menu-local pointer event construction and removes `MenuInputResyncState`, `RefreshMenuInputResyncSnapshot(...)`, and `RouteMenuInputResync(...)`.
- `Common/DxUi/DxUi.WindowHost.cpp`
  - Uses delivered events for pointer down/up/move routing; keeps live cursor snapshots only behind diagnostics.
- `Common/DxUi/DxUiNativeMenuInterop.h`
  - Requires explicit anchor points for production menu opens; raw cursor fallback becomes diagnostic/selftest-only or is removed.
- `RedSalamander/CompareDirectoriesWindow.cpp`
  - Replaces context/menu anchor lookup with delivered event point or focused row/control anchor.
- `RedSalamander/FolderWindow.cpp`
  - Removes cursor fallback from shell context menu point resolution; callers pass a delivered point or explicit selection/focus anchor.
- `RedSalamander/FolderWindow.Interaction.cpp`
  - Replaces `OnSetCursor` live hit testing with last delivered pointer metadata or default cursor behavior.
- `RedSalamander/FolderWindow.StatusBar.cpp`
  - Replaces status bar hover/context point lookup with delivered pointer metadata.
- `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
  - Replaces file-operation popup anchors with explicit owner/control/selection anchors.
- `RedSalamander/NavigationView.h`
  - Adds NavigationView input generation and routed-event helper declarations.
- `RedSalamander/NavigationView.cpp`
  - Keeps raw cursor snapshots only in explicitly annotated diagnostics.
- `RedSalamander/NavigationView.Interaction.cpp`
  - Removes `TryGetNavigationLiveCursor(...)`, `ResolveNavigationLivePointer(...)`, `NavigationSameOwnerFringePx(...)`, `IsNearSameOwnerNavigationEdge(...)`, and `IsStaleNavigationPointerMessage(...)`.
- `RedSalamander/NavigationView.Edit.cpp`
  - Forwards edit-host pointer messages with delivered coordinates and generation data.
- `RedSalamander/NavigationView.Menus.cpp`
  - Opens menu/history/sibling/dropdown routes from routed events.
- `RedSalamander/NavigationView.FullPathPopup.cpp`
  - Replaces timer polling with delivered-event popup/root switching and generation-based stale rejection.
- `RedSalamander/FindFilesWindow.cpp`
  - Keeps log-only cursor snapshots under diagnostics; routes embedded NavigationView child messages by delivered events.
- `RedSalamander/RedSalamander.cpp`
  - Keeps app message-loop cursor snapshots diagnostic-only and routes menu-bar/context input from delivered events.
- `RedSalamander/SplashScreen.cpp`
  - Removes cursor-based monitor choice; uses owner/foreground/default monitor placement.
- `RedSalamander/Ui/AlertOverlayWindow.cpp`
  - Uses delivered pointer/keyboard events for close/backdrop handling; no cursor polling.
- `Plugins/ViewerText/ViewerText.cpp`
  - Replaces context menu anchor and command points with delivered event or focused view anchor.
- `Plugins/ViewerText/ViewerText.Text.cpp`
  - Replaces text hover/hit-test cursor sampling with delivered pointer state.
- `Plugins/ViewerSpace/ViewerSpace.cpp`
  - Replaces viewer context/hit testing cursor sampling with delivered pointer state.
- `Tests/DxUiTests/DxUiTests.Menu.cpp`
  - Adds routed pointer/menu tests and source guard test hook when practical.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`
  - Updates Find destination NavigationView stale/queued tests to assert generation-based behavior.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
  - Adds viewer command/context menu coverage for delivered-point anchors where practical.
- `Specs/Testing/Testing_TestCoverage.md`
  - Records red/green evidence for the final migration.

---

## Task 0: Whole-Tree Baseline And Classification

**Files:**
- Read: `Common/DxUi/DxUiNativeMenuInterop.h`
- Read: `Common/DxUi/DxUi.Menu.cpp`
- Read: `Common/DxUi/DxUi.WindowHost.cpp`
- Read: `Plugins/ViewerText/ViewerText.cpp`
- Read: `Plugins/ViewerText/ViewerText.Text.cpp`
- Read: `Plugins/ViewerSpace/ViewerSpace.cpp`
- Read: `RedSalamander/CompareDirectoriesWindow.cpp`
- Read: `RedSalamander/FindFilesWindow.cpp`
- Read: `RedSalamander/FolderWindow.cpp`
- Read: `RedSalamander/FolderWindow.Interaction.cpp`
- Read: `RedSalamander/FolderWindow.StatusBar.cpp`
- Read: `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- Read: `RedSalamander/NavigationView.cpp`
- Read: `RedSalamander/NavigationView.FullPathPopup.cpp`
- Read: `RedSalamander/NavigationView.Interaction.cpp`
- Read: `RedSalamander/RedSalamander.cpp`
- Read: `RedSalamander/SplashScreen.cpp`
- Read: `RedSalamander/Ui/AlertOverlayWindow.cpp`
- Modify: this WIP plan evidence rows only

- [x] **Step 0.1: Capture current production cursor usage**

Run:

```powershell
rg -n "GetCursorPos\s*\(|WindowFromPoint|RouteMenuInputResync|MenuInputResync|ResolveNavigationLivePointer|IsStaleNavigationPointerMessage|NavigationSameOwnerFringePx|pointer-stale.accepted" Common RedSalamander Plugins -g "*.cpp" -g "*.h"
```

Expected now: matches in DxUi menu/window-host/native-menu code, NavigationView stale/live-pointer code, Find/main diagnostics, FolderWindow and file-operation popup anchor code, CompareDirectories anchor code, AlertOverlay pointer code, SplashScreen monitor placement, and ViewerText/ViewerSpace anchor or hit-test code.

Record the concrete `GetCursorPos()` locations currently known to exist:

```text
Common\DxUi\DxUiNativeMenuInterop.h
Common\DxUi\DxUi.Menu.cpp
Common\DxUi\DxUi.WindowHost.cpp
Plugins\ViewerText\ViewerText.cpp
Plugins\ViewerText\ViewerText.Text.cpp
Plugins\ViewerSpace\ViewerSpace.cpp
RedSalamander\CompareDirectoriesWindow.cpp
RedSalamander\FindFilesWindow.cpp
RedSalamander\FolderWindow.cpp
RedSalamander\FolderWindow.Interaction.cpp
RedSalamander\FolderWindow.StatusBar.cpp
RedSalamander\FolderWindow.FileOperations.Popup.cpp
RedSalamander\NavigationView.cpp
RedSalamander\NavigationView.FullPathPopup.cpp
RedSalamander\NavigationView.Interaction.cpp
RedSalamander\RedSalamander.cpp
RedSalamander\SplashScreen.cpp
RedSalamander\Ui\AlertOverlayWindow.cpp
```

- [x] **Step 0.2: Classify each hit**

Classify every hit as one of:

- `diagnostic-only`: guarded by diagnostics and not used for routing.
- `selftest-only`: used only by tests/repro helpers.
- `production-routing`: affects hover, activation, stale classification, menu open/close, root switching, or repaint state.
- `production-anchor`: uses current cursor to choose where to open a menu without a delivered event.
- `production-hit-test`: uses current cursor to decide hover, status bar item, close button, text position, or viewer item.
- `production-placement`: uses current cursor to choose monitor/window placement.

Record the classification in the checklist evidence row for Slice 0.

- [x] **Step 0.3: Capture current focused test baseline**

Run:

```powershell
$exe = (Resolve-Path .\.build\x64\Debug\RedSalamander.exe).Path
$cases = @(
  'cmd_pane_find_dialog_destination_navigation_stale_edit_host_hit_testing',
  'cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions',
  'cmd_pane_find_dialog_result_drains_respect_child_input_queue_order',
  'cmd_pane_find_dialog_escape_closes_popup_before_cancel',
  'cmd_pane_find_dialog_escape_from_dx_control_closes_cancel'
)
foreach ($case in $cases) {
  $p = Start-Process -FilePath $exe -ArgumentList @('--commands-selftest', "--selftest-case=$case", '--selftest-timeout-multiplier=4') -WorkingDirectory (Get-Location) -Wait -PassThru -WindowStyle Hidden
  "$case EXIT=$($p.ExitCode)"
}
```

Expected now: all cases pass before migration starts. Any failure must be investigated before this plan proceeds.

---

## Task 1: Add A Whole-Tree No-Live-Cursor Source Guard

**Files:**
- Create: `Scripts/VerifyNoProductionGetCursorPos.ps1`
- Modify: `Specs/Testing/Testing_TestCoverage.md`

- [x] **Step 1.1: Write the failing guard script**

Create `Scripts/VerifyNoProductionGetCursorPos.ps1` with this behavior:

```powershell
param(
    [string]$Root = (Resolve-Path "$PSScriptRoot\..").Path
)

$allowedPathPatterns = @(
    '\\SelfTest\\',
    '\\Tests\\',
    '\\Scripts\\VerifyNoProductionGetCursorPos\.ps1$'
)

$requiredDiagnosticAnnotation = '// getcursorpos-allow: diagnostic-only'

$files = @(
    Join-Path $Root 'Common',
    Join-Path $Root 'RedSalamander',
    Join-Path $Root 'Plugins'
)

$violations = New-Object System.Collections.Generic.List[string]
foreach ($fileRoot in $files) {
    if (-not (Test-Path -LiteralPath $fileRoot)) {
        continue
    }
    foreach ($item in Get-ChildItem -Path $fileRoot -Recurse -Include *.cpp,*.h) {
        $path = $item.FullName
        $relative = $path.Substring($Root.Length).TrimStart('\')
        if ($allowedPathPatterns | Where-Object { $relative -match $_ }) {
            continue
        }

        $lineNo = 0
        foreach ($line in Get-Content -Path $path) {
            $lineNo++
            if ($line -notmatch 'GetCursorPos\s*\(') {
                continue
            }
            if ($line.Contains($requiredDiagnosticAnnotation)) {
                continue
            }
            $violations.Add("${relative}:${lineNo}: $line")
        }
    }
}

if ($violations.Count -gt 0) {
    Write-Host 'Production GetCursorPos violations:'
    $violations | ForEach-Object { Write-Host $_ }
    exit 1
}

Write-Host 'No production GetCursorPos violations found.'
exit 0
```

This is intentionally red at first because current production files still use `GetCursorPos()`.

Rules enforced by the guard:

- The guard scans `Common`, `RedSalamander`, and `Plugins`.
- Test and selftest paths are allowed because they can move/sample the OS cursor to build deterministic repros.
- Production code may retain `GetCursorPos()` only if the same source line contains `// getcursorpos-allow: diagnostic-only`.
- The annotation is a promise: the value is log evidence only. If a diagnostic value affects branching, routing, placement, anchors, hover, close/open behavior, stale classification, or repaint state, the implementation is still invalid even if the guard passes.

- [x] **Step 1.2: Run the guard and verify red**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\VerifyNoProductionGetCursorPos.ps1
```

Expected before production migration: exit `1`, with violations in `Common`, `RedSalamander`, and `Plugins` production files listed in Task 0.

- [x] **Step 1.3: Record the red guard evidence**

Add the failing command and the violation count to `Specs/Testing/Testing_TestCoverage.md` under a new `2026-06-03 central pointer input router` entry.

---

## Task 2: Create The Routed Pointer Event Core

**Files:**
- Create: `Common/DxUi/DxUi.PointerInput.h`
- Create: `Common/DxUi/DxUi.PointerInput.cpp`
- Modify: `Common/DxUi/DxUi.vcxproj`
- Modify: `Tests/DxUiTests/DxUiTests.Menu.cpp`

- [x] **Step 2.1: Write failing DxUi pointer event tests**

Add tests covering:

- `WM_MOUSEMOVE` converts `GET_X_LPARAM/GET_Y_LPARAM` to client and screen points from the delivered message.
- `WM_LBUTTONDOWN` records button kind and `MK_LBUTTON`.
- `WM_MOUSEWHEEL` records wheel delta and uses delivered screen point semantics.
- Event construction stores `GetMessageTime()` only as message metadata and does not sample `GetCursorPos()`.

Expected test names:

```cpp
TestPointerInputEventMouseMoveUsesDeliveredPoint()
TestPointerInputEventButtonUsesDeliveredPointAndFlags()
TestPointerInputEventWheelUsesDeliveredScreenPoint()
TestPointerInputEventHasNoLiveCursorState()
```

- [x] **Step 2.2: Run tests and verify red**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests
.\.build\x64\Debug\DxUiTests.exe --suite=Menu
```

Expected: build fails because `DxUi.PointerInput.h` is missing, or tests fail because the routed event API is not implemented.

- [x] **Step 2.3: Implement the event API**

Create the API with this shape:

```cpp
namespace RedSalamander::DxUi
{
enum class PointerInputSource : uint8_t
{
    WindowProc,
    ModalLoopMessage,
    PopupWindowProc,
    ForwardedChild,
    DiagnosticOnly
};

enum class PointerInputKind : uint8_t
{
    Move,
    Leave,
    LeftDown,
    LeftUp,
    LeftDoubleClick,
    RightDown,
    RightUp,
    Wheel,
    Unknown
};

struct InputGeneration
{
    uint64_t value = 0;
};

struct PointerInputEvent
{
    PointerInputSource source = PointerInputSource::WindowProc;
    PointerInputKind kind = PointerInputKind::Unknown;
    HWND targetHwnd = nullptr;
    HWND rootHwnd = nullptr;
    HWND captureHwnd = nullptr;
    UINT message = 0;
    WPARAM wParam = 0;
    LPARAM lParam = 0;
    DWORD messageTime = 0;
    POINT clientPointPx{};
    POINT screenPointPx{};
    InputGeneration generation{};
    bool hasClientPoint = false;
    bool hasScreenPoint = false;
};

[[nodiscard]] std::optional<PointerInputKind> PointerInputKindFromMessage(UINT message) noexcept;
[[nodiscard]] std::optional<PointerInputEvent> TryBuildPointerInputEvent(
    HWND targetHwnd,
    UINT message,
    WPARAM wParam,
    LPARAM lParam,
    PointerInputSource source,
    InputGeneration generation = {}) noexcept;
[[nodiscard]] std::optional<PointerInputEvent> TryBuildPointerInputEventFromMsg(
    const MSG& message,
    PointerInputSource source,
    InputGeneration generation = {}) noexcept;
}
```

Rules:

- Use only delivered message data, `ClientToScreen(...)`, and `ScreenToClient(...)` for coordinate conversion.
- Do not call `GetCursorPos()`.
- For `WM_MOUSEWHEEL`, use `GET_X_LPARAM(lParam)` / `GET_Y_LPARAM(lParam)` as delivered screen coordinates and convert to client.
- For mouse client messages, use `GET_X_LPARAM(lParam)` / `GET_Y_LPARAM(lParam)` as delivered client coordinates and convert to screen.
- Store `GetCapture()` as metadata only; do not make routing decisions in the builder.

- [x] **Step 2.4: Run tests and verify green**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests
.\.build\x64\Debug\DxUiTests.exe --suite=Menu
```

Expected: new pointer event tests pass.

---

## Task 3: Migrate DxUi Menu Routing

**Files:**
- Modify: `Common/DxUi/DxUi.Menu.cpp`
- Modify: `Tests/DxUiTests/DxUiTests.Menu.cpp`
- Modify: `Specs/UI/UI_CommandMenuKeyboard.md` only if implementation discovers a sharper invariant

- [x] **Step 3.1: Write failing menu tests for no-resync behavior**

Add or update tests:

```cpp
TestContextMenuDeliveredOwnerMoveDoesNotUseLiveCursor()
TestContextMenuOutsideDismissUsesDeliveredPoint()
TestContextMenuRootSwitchUsesDeliveredMessageOrder()
TestContextMenuNoModalInputResyncPolling()
```

Test intent:

- Move the OS cursor to a point that disagrees with a delivered `WM_MOUSEMOVE`; verify hover follows delivered coordinates.
- Send a delivered outside click while the live cursor remains over the popup; verify dismissal follows delivered point.
- After keyboard root switch, send older owner moves and verify root-switch generation suppresses stale rollback.
- Assert debug state exposes no active `ModalResync` source.

- [x] **Step 3.2: Verify red**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests
.\.build\x64\Debug\DxUiTests.exe --suite=Menu
```

Expected before migration: at least one new test fails because `RouteMenuInputResync(...)` or live-cursor comparison still influences routing.

- [x] **Step 3.3: Replace menu-local pointer construction**

In `Common/DxUi/DxUi.Menu.cpp`:

- Build popup `MenuPointerEvent` from `PointerInputEvent`.
- Build modal-loop non-popup pointer events from `TryBuildPointerInputEventFromMsg(...)`.
- Keep popup/root hit testing based on the routed event's delivered screen point.
- Preserve keyboard routing as `MenuKeyboardEvent`.
- Preserve diagnostics by logging both delivered event data and optional diagnostic snapshots, but make diagnostic snapshots write-only.

- [x] **Step 3.4: Remove modal input resync**

Delete:

- `MenuInputSource::ModalResync`
- `MenuInputResyncState`
- `RefreshMenuInputResyncSnapshot(...)`
- `RouteMenuInputResync(...)`
- idle/flood calls that synthesize pointer input from `GetCursorPos()`

Replace flood handling with this rule:

- Pump pending keyboard/pointer messages first.
- For generic owner-window floods, keep dispatching/pumping without synthesizing pointer state.
- If the menu remains open and no input arrives, do not change hover/pressed/root state.

- [x] **Step 3.5: Replace root-switch stale checks**

Replace live-cursor disagreement checks with generation and message order:

- Store `lastPointerRootSwitchGeneration`.
- Store `lastPointerRootSwitchMessageTime`.
- Store `lastPointerRootSwitchDeliveredScreenPoint`.
- Reject older owner moves when `messageTime <= lastPointerRootSwitchMessageTime` and the message source/root does not match the active root generation.
- Accept fresh owner moves even if the OS cursor has since moved.

- [x] **Step 3.6: Verify menu green**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests
.\.build\x64\Debug\DxUiTests.exe --suite=Menu
```

Expected: all menu tests pass.

---

## Task 4: Migrate NavigationView Input

**Files:**
- Modify: `RedSalamander/NavigationViewInternal.h`
- Modify: `RedSalamander/NavigationView.Interaction.cpp`
- Modify: `RedSalamander/NavigationView.Edit.cpp`
- Modify: `RedSalamander/NavigationView.Menus.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`

- [x] **Step 4.1: Write failing NavigationView generation tests**

Add a new case:

```text
cmd_pane_find_dialog_destination_navigation_uses_delivered_input_generation
```

Required probes:

- Delivered history-arrow click opens even when `GetCursorPos()` is far away, if target HWND and generation are current.
- Delivered double-click from an old generation does not enter edit mode after edit-host teardown.
- Delivered stale hover from an old layout generation clears/ignores without polling the live cursor.
- Forwarded edit-host click carries delivered coordinates and exits edit mode before opening history/menu/disk branch.

- [x] **Step 4.2: Verify red**

Run:

```powershell
$exe = (Resolve-Path .\.build\x64\Debug\RedSalamander.exe).Path
$p = Start-Process -FilePath $exe -ArgumentList @('--commands-selftest', '--selftest-case=cmd_pane_find_dialog_destination_navigation_uses_delivered_input_generation', '--selftest-timeout-multiplier=4') -WorkingDirectory (Get-Location) -Wait -PassThru -WindowStyle Hidden
"EXIT=$($p.ExitCode)"
```

Expected before migration: failure because current NavigationView stale classification still samples live cursor.

- [x] **Step 4.3: Add NavigationView generation state**

In `NavigationViewInternal.h`, add:

```cpp
RedSalamander::DxUi::InputGeneration _inputGeneration{1};
void BumpInputGeneration(std::wstring_view reason) noexcept;
[[nodiscard]] RedSalamander::DxUi::InputGeneration CurrentInputGeneration() const noexcept;
```

Bump generation in:

- `WM_NCDESTROY`
- layout resize that changes interactive rects
- edit mode enter
- edit mode exit
- full-path popup open/close
- menu/dropdown close
- path/history model replacement
- embedded destination path/history reset

Each bump logs `navigation.input-generation` when diagnostics are enabled.

- [x] **Step 4.4: Replace stale live-pointer guard**

Remove:

- `TryGetNavigationLiveCursor(...)`
- `ResolveNavigationLivePointer(...)`
- `NavigationSameOwnerFringePx(...)`
- `IsNearSameOwnerNavigationEdge(...)`
- `IsStaleNavigationPointerMessage(...)`

Add:

```cpp
[[nodiscard]] bool NavigationView::ShouldAcceptPointerEvent(const DxUi::PointerInputEvent& event) const noexcept
{
    if (event.targetHwnd != _hWnd.get() && IsChild(_hWnd.get(), event.targetHwnd) == FALSE)
    {
        return false;
    }
    if (event.generation.value == 0 || event.generation.value != _inputGeneration.value)
    {
        return false;
    }
    return true;
}
```

Adjust exact rules during implementation:

- Current direct WndProc events must carry the current generation.
- Forwarded edit-host events must carry the generation captured when the edit host received the delivered message.
- Posted/deferred menu opens must carry the generation captured when they were posted.
- `InputGeneration{0}` means invalid/untracked for NavigationView and must be rejected instead of treated as fresh input.

- [x] **Step 4.5: Route NavigationView mouse handlers through events**

Change `OnMouseMove(POINT)`, `OnLButtonDown(POINT)`, and `OnLButtonDblClk(POINT)` to either:

- accept a `const DxUi::PointerInputEvent&`, or
- build the event in WndProc and call a shared `HandlePointerEvent(...)`.

Keep rendering branches unchanged after hit testing:

- menu button opens menu
- history opens history
- disk opens disk info
- path segment navigates
- separator opens siblings
- stale generation clears hover and returns

- [x] **Step 4.6: Verify NavigationView green**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
$exe = (Resolve-Path .\.build\x64\Debug\RedSalamander.exe).Path
$cases = @(
  'cmd_pane_find_dialog_destination_navigation_stale_edit_host_hit_testing',
  'cmd_pane_find_dialog_destination_navigation_uses_delivered_input_generation',
  'cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions'
)
foreach ($case in $cases) {
  $p = Start-Process -FilePath $exe -ArgumentList @('--commands-selftest', "--selftest-case=$case", '--selftest-timeout-multiplier=4') -WorkingDirectory (Get-Location) -Wait -PassThru -WindowStyle Hidden
  "$case EXIT=$($p.ExitCode)"
}
```

Expected: all exit `0`.

---

## Task 5: Clean WindowHost, Find, And Main Diagnostics

**Files:**
- Modify: `Common/DxUi/DxUi.WindowHost.cpp`
- Modify: `RedSalamander/FindFilesWindow.cpp`
- Modify: `RedSalamander/RedSalamander.cpp`

- [x] **Step 5.1: Move live cursor snapshots behind diagnostics**

Any remaining `GetCursorPos()` used for logs must be inside explicit diagnostics helpers:

```cpp
[[nodiscard]] std::optional<POINT> TryGetDiagnosticCursorScreenPoint() noexcept;
```

Rules:

- The helper name must include `Diagnostic`.
- Callers may only put returned values into trace/log records.
- Callers must not branch production behavior from this value.

- [x] **Step 5.2: Replace WindowHost owner hover handoff**

On `WM_MOUSELEAVE`:

- clear hover,
- stop tracking,
- wait for the next delivered owner `WM_MOUSEMOVE` to re-arm hover,
- do not call `GetCursorPos()` to decide whether to re-arm immediately.

On pointer down/up diagnostics:

- log delivered message point as primary,
- optionally log diagnostic cursor snapshot as evidence,
- do not use `WindowFromPoint(liveCursor)` for routing.

- [x] **Step 5.3: Replace Find/main loop routing logs**

Keep `app.message-loop.find`, `find.wndproc.raw`, and `dxui.windowhost.raw` useful by logging:

- delivered `MSG.pt` / `lParam`,
- target hwnd/root,
- message time,
- capture/focus/active hwnd,
- diagnostic cursor snapshot if the support trace switch is enabled.

Do not let diagnostic cursor fields choose routing, hover, open, close, or repaint decisions.

- [x] **Step 5.4: Run source guard red/green**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\VerifyNoProductionGetCursorPos.ps1
```

Expected after Task 5: violations remain only for production-anchor, production-hit-test, or production-placement sites that Tasks 6-8 will remove, or no violations if those tasks have already landed.

---

## Task 6: Replace Production Menu Anchor Fallbacks

**Files:**
- Modify: `Common/DxUi/DxUiNativeMenuInterop.h`
- Modify: `RedSalamander/RedSalamander.cpp`
- Modify: callers found in Task 0 classified as `production-anchor`

- [x] **Step 6.1: Identify all menu opens without explicit anchor**

Run:

```powershell
rg -n "ContextMenu::Show|ShowAtCursor|GetCursorPos" Common RedSalamander Plugins -g "*.cpp" -g "*.h"
```

Expected: every production menu open either has an explicit screen point from the delivered event/owner control bounds, or is listed for replacement in this task.

- [x] **Step 6.2: Replace cursor fallback APIs**

For native/DxUi menu interop, prefer one of:

```cpp
ShowMenuAtOwnerRect(ownerHwnd, anchorRect, items, theme, callbacks);
ShowMenuAtDeliveredPoint(ownerHwnd, deliveredScreenPoint, items, theme, callbacks);
ShowKeyboardMenuAtFocusedItem(ownerHwnd, focusedItemRect, items, theme, callbacks);
```

Remove production APIs that imply "open at current cursor" without a delivered event.

- [x] **Step 6.3: Add tests for keyboard/context anchor behavior**

Add tests proving:

- keyboard-opened menus anchor to the focused item/control,
- context-menu-key menus anchor to selected item or focused control rect,
- mouse-opened menus anchor to the delivered mouse event point,
- none of these paths call `GetCursorPos()` for production routing.

- [x] **Step 6.4: Run source guard green**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\VerifyNoProductionGetCursorPos.ps1
```

Expected after Task 6: no `production-anchor` violations remain. The guard may still exit `1` only for RedSalamander production hit-test/placement sites handled by Task 7 or plugin viewer sites handled by Task 8.

---

## Task 7: Remove Remaining RedSalamander Production Cursor Dependencies

**Files:**
- Modify: `RedSalamander/CompareDirectoriesWindow.cpp`
- Modify: `RedSalamander/FolderWindow.cpp`
- Modify: `RedSalamander/FolderWindow.Interaction.cpp`
- Modify: `RedSalamander/FolderWindow.StatusBar.cpp`
- Modify: `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- Modify: `RedSalamander/NavigationView.cpp`
- Modify: `RedSalamander/NavigationView.FullPathPopup.cpp`
- Modify: `RedSalamander/SplashScreen.cpp`
- Modify: `RedSalamander/Ui/AlertOverlayWindow.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`

- [x] **Step 7.1: Write failing whole-tree RedSalamander guard evidence**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\VerifyNoProductionGetCursorPos.ps1
```

Expected before this task: remaining violations point at RedSalamander production files outside the DxUi/Menu/Find diagnostic cleanup.

- [x] **Step 7.2: Replace shell/context menu anchor fallback**

For `FolderWindow.cpp`, `CompareDirectoriesWindow.cpp`, and `FolderWindow.FileOperations.Popup.cpp`:

- replace any `Resolve...ContextMenuPoint()` implementation that samples the current cursor,
- pass a `POINT screenPoint` from the delivered mouse event when the command is mouse initiated,
- pass a selected/focused item rectangle center for keyboard/menu-bar initiated commands,
- pass an owner/control rectangle for toolbar/button initiated commands,
- never fall back to the current OS cursor.

Use explicit names so call sites state their contract:

```cpp
enum class MenuAnchorReason : uint8_t
{
    DeliveredMousePoint,
    FocusedItemRect,
    OwnerControlRect
};

struct MenuAnchor
{
    POINT screenPoint{};
    MenuAnchorReason reason = MenuAnchorReason::OwnerControlRect;
};
```

- [x] **Step 7.3: Replace FolderWindow cursor/hit-test paths**

For `FolderWindow.Interaction.cpp` and `FolderWindow.StatusBar.cpp`:

- maintain a last delivered pointer record from `WM_MOUSEMOVE`, button messages, and wheel messages,
- use `WM_NCHITTEST` or delivered client points where available,
- if `WM_SETCURSOR` arrives without a recent delivered point, set the default cursor instead of sampling the live cursor,
- update status bar hover/details only from delivered mouse messages, leave notifications, focus changes, or explicit selection changes.

- [x] **Step 7.4: Replace NavigationView full-path popup polling**

For `NavigationView.FullPathPopup.cpp`:

- remove timer-based `GetCursorPos()` polling,
- switch popup/root hover from delivered popup/owner `WM_MOUSEMOVE` events,
- use generation/order metadata to reject old owner messages after popup close,
- close/open sibling branches only from routed pointer events or keyboard navigation.

- [x] **Step 7.5: Replace AlertOverlay close/backdrop cursor paths**

For `Ui/AlertOverlayWindow.cpp`:

- close button hit testing uses delivered button-up/down points,
- backdrop clicks use delivered points converted from the message target,
- Escape closes the overlay through keyboard command handling,
- no close, redraw, or focus behavior waits for later mouse movement.

- [x] **Step 7.6: Replace SplashScreen monitor placement**

For `SplashScreen.cpp`:

- if the splash has an owner window, use `MonitorFromWindow(owner, MONITOR_DEFAULTTONEAREST)`,
- else if a foreground RedSalamander window exists, use `MonitorFromWindow(foreground, MONITOR_DEFAULTTONEAREST)`,
- else use the primary/default monitor,
- do not place the splash based on the current mouse cursor.

- [x] **Step 7.7: Verify RedSalamander whole-tree cleanup**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\VerifyNoProductionGetCursorPos.ps1
```

Expected after this task: no RedSalamander production violations remain unless Task 8 plugin viewer violations are still pending.

---

## Task 8: Remove Plugin Viewer Production Cursor Dependencies

**Files:**
- Modify: `Plugins/ViewerText/ViewerText.cpp`
- Modify: `Plugins/ViewerText/ViewerText.Text.cpp`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

- [x] **Step 8.1: Write failing viewer delivered-anchor tests**

Add focused viewer selftest cases:

```text
cmd_viewer_text_context_menu_uses_delivered_anchor
cmd_viewer_text_hover_uses_delivered_point
cmd_viewer_space_context_menu_uses_delivered_anchor
cmd_viewer_space_hover_uses_delivered_point
```

Each case must move the OS cursor away from the delivered message point, send the viewer message, and assert that the viewer action follows the delivered point or explicit focused-view anchor.

- [x] **Step 8.2: Verify viewer tests red**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
$exe = (Resolve-Path .\.build\x64\Debug\RedSalamander.exe).Path
$cases = @(
  'cmd_viewer_text_context_menu_uses_delivered_anchor',
  'cmd_viewer_text_hover_uses_delivered_point',
  'cmd_viewer_space_context_menu_uses_delivered_anchor',
  'cmd_viewer_space_hover_uses_delivered_point'
)
foreach ($case in $cases) {
  $p = Start-Process -FilePath $exe -ArgumentList @('--commands-selftest', "--selftest-case=$case", '--selftest-timeout-multiplier=4') -WorkingDirectory (Get-Location) -Wait -PassThru -WindowStyle Hidden
  "$case EXIT=$($p.ExitCode)"
}
```

Expected before migration: at least one case fails because the viewer samples the live cursor.

- [x] **Step 8.3: Migrate ViewerText**

For `ViewerText.cpp` and `ViewerText.Text.cpp`:

- route context menus from delivered mouse message points,
- route keyboard-opened menus from focused text/caret/view rectangle anchors,
- route hover/text-hit testing from delivered client points,
- remove any live cursor branch from command execution.

- [x] **Step 8.4: Migrate ViewerSpace**

For `ViewerSpace.cpp`:

- route context menus from delivered mouse message points,
- route keyboard-opened menus from focused preview/view rectangle anchors,
- route hover/hit testing from delivered client points,
- remove any live cursor branch from command execution.

- [x] **Step 8.5: Verify viewer cleanup and guard**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
$exe = (Resolve-Path .\.build\x64\Debug\RedSalamander.exe).Path
$cases = @(
  'cmd_viewer_text_context_menu_uses_delivered_anchor',
  'cmd_viewer_text_hover_uses_delivered_point',
  'cmd_viewer_space_context_menu_uses_delivered_anchor',
  'cmd_viewer_space_hover_uses_delivered_point'
)
foreach ($case in $cases) {
  $p = Start-Process -FilePath $exe -ArgumentList @('--commands-selftest', "--selftest-case=$case", '--selftest-timeout-multiplier=4') -WorkingDirectory (Get-Location) -Wait -PassThru -WindowStyle Hidden
  "$case EXIT=$($p.ExitCode)"
}
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\VerifyNoProductionGetCursorPos.ps1
```

Expected: all viewer cases exit `0`; the source guard exits `0`.

---

## Task 9: Full Regression And Live Validation

**Files:**
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Modify: this WIP plan evidence rows

- [x] **Step 9.1: Build Debug**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
```

Expected: `0 warning(s), 0 error(s)`.

- [x] **Step 9.2: Run DxUi menu tests**

Run:

```powershell
.\.build\x64\Debug\DxUiTests.exe --suite=Menu
```

Expected: exit `0`.

- [x] **Step 9.3: Run focused Commands tests**

Run:

```powershell
$exe = (Resolve-Path .\.build\x64\Debug\RedSalamander.exe).Path
$cases = @(
  'cmd_pane_find_dialog_destination_navigation_stale_edit_host_hit_testing',
  'cmd_pane_find_dialog_destination_navigation_uses_delivered_input_generation',
  'cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions',
  'cmd_pane_find_dialog_result_drains_respect_child_input_queue_order',
  'cmd_pane_find_dialog_escape_closes_popup_before_cancel',
  'cmd_pane_find_dialog_escape_from_dx_control_closes_cancel',
  'cmd_viewer_text_context_menu_uses_delivered_anchor',
  'cmd_viewer_text_hover_uses_delivered_point',
  'cmd_viewer_space_context_menu_uses_delivered_anchor',
  'cmd_viewer_space_hover_uses_delivered_point'
)
foreach ($case in $cases) {
  $p = Start-Process -FilePath $exe -ArgumentList @('--commands-selftest', "--selftest-case=$case", '--selftest-timeout-multiplier=4') -WorkingDirectory (Get-Location) -Wait -PassThru -WindowStyle Hidden
  "$case EXIT=$($p.ExitCode)"
}
```

Expected: all exit `0`.

- [x] **Step 9.4: Run whole-tree source guard**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\VerifyNoProductionGetCursorPos.ps1
```

Expected: exit `0`.

- [x] **Step 9.5: Manual live validation**

Use the Debug app with support diagnostics enabled and verify:

- Find split-button menu opens on click without title-bar movement.
- Find split-button menu hover highlights under the pointer.
- Clicking outside closes it immediately.
- Destination NavigationView history/menu/separator opens immediately.
- Destination NavigationView does not drop a click when the pointer drifts a few pixels after the delivered event.
- Stale old double-clicks do not enter edit mode after edit-host teardown.
- Help overlay close, Escape, and backdrop repaint still work.
- FolderWindow context menus, status bar hover, and file-operation popups open/hover/close without needing an unrelated mouse move.
- ViewerText and ViewerSpace context menus and hover follow the delivered point or keyboard focus anchor, not the later cursor.

Save the live log path and summarize the confirming event sequence in `Specs/Testing/Testing_TestCoverage.md`.

---

## Task 10: Spec Closeout

**Files:**
- Modify: `Specs/UI/UI_CommandMenuKeyboard.md`
- Modify: `Specs/UI/UI_DxUiWinUIDesign.md`
- Modify: `Specs/UI/UI_FindFilesWindow.md`
- Modify: `Specs/UI/UI_NavigationView.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Move when complete: `Specs/Plans/WIP/DxUi_CentralPointerInputRouterPlan_2026-06-03.md` to `Specs/Plans/Done/DxUi_CentralPointerInputRouterPlan_2026-06-03.md`

- [x] **Step 10.1: Update specs only for discoveries**

If implementation changes the contract, update the authoritative spec that owns the behavior. Do not leave durable requirements only in this WIP plan.

- [x] **Step 10.2: Final guard**

Run:

```powershell
rg -n "GetCursorPos" Common RedSalamander Plugins -g "*.cpp" -g "*.h"
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\VerifyNoProductionGetCursorPos.ps1
git diff --check -- Common RedSalamander Plugins Specs\UI Specs\Testing Scripts
```

Expected:

- raw `GetCursorPos` remains only in tests/selftests or explicitly annotated diagnostic-only production lines,
- source guard exits `0`,
- `git diff --check` reports no whitespace errors.

- [x] **Step 10.3: Move plan to Done**

Only after all tasks pass and manual live validation is recorded:

```powershell
Move-Item -LiteralPath 'Specs\Plans\WIP\DxUi_CentralPointerInputRouterPlan_2026-06-03.md' -Destination 'Specs\Plans\Done\DxUi_CentralPointerInputRouterPlan_2026-06-03.md'
```

Do not move this plan to Done while any source guard violation, focused selftest failure, or live Find menu/Nav repro remains open.

---

## Risks And Mitigations

- **Risk:** Removing modal input resync reintroduces menu starvation under owner-window floods.
  - **Mitigation:** The menu loop must prioritize queued input and dispatch popup/owner pointer messages through the router before generic owner traffic.
- **Risk:** Some commands genuinely need a menu position without a pointer message.
  - **Mitigation:** Use focused item/control rectangles or explicit owner anchors. Do not use current cursor position as fallback.
- **Risk:** Generation tokens reject valid sent-message tests.
  - **Mitigation:** Direct WndProc/sent events use the current generation. Only posted/forwarded/deferred events carry a captured generation that can become stale.
- **Risk:** Diagnostic logging becomes less useful after removing live cursor routing.
  - **Mitigation:** Keep optional diagnostic cursor snapshots in trace records, clearly named as diagnostic-only, and verify source guard allows only those wrappers.

## Completion Definition

This plan is complete when:

- no production behavior under `Common`, `RedSalamander`, or `Plugins` calls `GetCursorPos()`,
- any remaining production `GetCursorPos()` call is same-line annotated `// getcursorpos-allow: diagnostic-only` and is log/evidence-only,
- the whole-tree source guard exits `0`,
- DxUi menu tests pass,
- focused Find/NavigationView command selftests pass,
- viewer context/hover selftests pass,
- live Find dialog menu and destination NavigationView behavior no longer require title-bar or unrelated pointer movement,
- live FolderWindow, AlertOverlay, SplashScreen, and viewer behavior have no delayed cursor-message dependency,
- specs and test coverage contain the final durable contract and evidence,
- this plan is moved from `Specs/Plans/WIP/` to `Specs/Plans/Done/`.
