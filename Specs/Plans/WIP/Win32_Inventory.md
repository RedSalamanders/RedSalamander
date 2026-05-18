# Win32 Inventory

Last updated: 2026-05-18

Status: WIP - inventory captured; follow-up audits remain open.

## Scope

This inventory covers Win32-facing code in the RedSalamander solution: production app code, shared `Common` code, plugins, RedConfigure, RedLauncher, RedSalamanderMonitor, RedSalamanderSearchService, resource scripts, and in-tree selftests. It excludes generated/build folders, `.git`, `.vs`, and PoC-only projects unless noted.

The scan is regex-assisted, not a semantic compiler analysis. Counts are useful for locating hot spots, but each follow-up item still needs code review before changing behavior.

## Static Signal Summary

Primary app/library scan excluding `Tests/`, `PoC/`, and `RedSalamander/SelfTest/`: 449 C++/header/resource files.

| Signal | Files | Matches | Meaning |
| --- | ---: | ---: | --- |
| Handles | 230 | 3880 | `HWND`, `HDC`, `HINSTANCE`, `HMENU`, `HICON`, `HANDLE`, `HKEY`, `HBITMAP`, `HRGN`, etc. |
| Messages | 77 | 2371 | `WM_*`, `WndMsg::`, `PostMessage*`, `SendMessage*`, payload helpers. |
| WndProc/hooks | 70 | 1731 | WndProc/DialogProc, `GWLP_USERDATA`, `GWLP_WNDPROC`, `DefWindowProc`, subclass hooks. |
| RAII | 176 | 2035 | `wil::unique_*`, `wil::com_ptr`, `wil::scope_exit`, WIL COM/event helpers. |
| Window classes | 39 | 200 | `RegisterClass*`, `CreateWindow*`, resource dialogs. |
| DPI/theme | 49 | 322 | `WM_DPICHANGED`, `WM_DPICHANGED_AFTERPARENT`, DWM/theme APIs. |
| Manual cleanup | 45 | 112 | `DestroyWindow`, `DeleteObject`, `DestroyMenu`, `Release()`, etc. Needs semantic review. |
| Shell/COM UI | 47 | 198 | shell APIs, file dialogs, `ShellExecuteW`, `MessageBoxW`, `TrackPopupMenu`, etc. |

Including in-tree `RedSalamander/SelfTest/` raises the production-tree scan to 483 files and 17K+ Win32-related matches, mostly because command selftests drive UI by posting/sending Win32 messages.

## Highest-Density Files

| File | Why it matters |
| --- | --- |
| `RedSalamander/RedSalamander.cpp` | Main window registration, message loop, menu dispatch, app-level custom messages, full screen, shell launches, startup/shutdown. |
| `RedSalamander/Preferences.Dialog.cpp` | Resource dialog plus DxUi hosts, wheel routing, deferred actions, page host payload drain, DPI/theme routing. |
| `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp` | Multiple custom prompt/dialog windows for file-system commands and background task completion. |
| `RedSalamander/ManagePluginsDialog.cpp` | Plugin config dialog, dynamic native child controls, Dx host hooks, dialog proc. |
| `RedSalamander/FolderWindow.FileSystem.cpp` | Prompt windows and navigation/file-system dialogs with repeated WndProc patterns. |
| `RedSalamander/FolderWindow.FileOperations.Popup.cpp` | Popup window, D2D/DWrite render target, DPI, timers/deferred speed-limit prompt. |
| `RedSalamander/CompareDirectoriesWindow.Options.cpp` | Resource dialog plus DxUi host windows, WndProc hooks, wheel routing, DPI/theme. |
| `RedSalamander/FolderView.cpp` | Custom pane view WndProc, posted enumeration/icon/thumbnail payloads, `WM_DPICHANGED_AFTERPARENT`. |
| `RedSalamander/NavigationView.cpp` | Custom view plus Dx host WndProc, posted suggest/path payloads, DPI-after-parent handling. |
| `Common/DxUi/DxUi.WindowHost.cpp` | Shared Direct3D/DXGI/DirectWrite host, swap chains, `WM_DPICHANGED` and `WM_DPICHANGED_AFTERPARENT`. |
| `Common/DxUi/DxUi.Menu.cpp` | Native-style menu flyout window, nested message loop, popup WndProc, D2D/DWrite painting. |
| `Plugins/ViewerText/ViewerText.cpp` | Viewer top-level/embedded windows, text and hex child windows, file combo host hook, async payloads. |
| `Plugins/ViewerVLC/ViewerVLC.cpp` | Multiple registered window classes for main/video/HUD/overlay/seek preview and message forwarding. |
| `Plugins/ViewerWeb/ViewerWeb.cpp` | WebView host window, file combo hook, async load payloads, file dialogs, shell launch. |
| `RedSalamanderMonitor/ColorTextView.cpp` | Custom text view, worker payload messages, subclassed find panel/edit controls, Direct3D/DXGI. |

## Central Win32 Infrastructure

- `Common/WindowMessages.h` is the custom message registry. Keep all `WM_APP`/`WM_USER` IDs there unless a component-local constant is intentionally private.
- `Common/Helpers.h` owns `PostMessagePayload(...)`, `TakeMessagePayload<T>(...)`, `InitPostedPayloadWindow(...)`, and `DrainPostedPayloadsForWindow(...)`.
- `Common/Win32CallbackHelpers.h` centralizes no-throw wrappers for `SetWindowLongPtrW`, `CallWindowProcW`, dialog creation, and WndProc hook install/restore.
- `Common/ViewerFileComboHost.h` standardizes viewer file-combo host subclassing.
- `Common/WindowSizing.h` and `Common/WindowBackdropPolicy.h` carry reusable DPI/min-track/backdrop helpers.
- `Common/DxUi/*` is the shared retained DirectX UI layer sitting on top of HWNDs, message routing, native menu interop, accessibility, native text input, and D2D/DWrite rendering.

## Custom Message Map

Current message ranges in `Common/WindowMessages.h`:

| Range | Owner |
| --- | --- |
| `WM_APP + 0x300..0x309` | Folder view sync, enumeration, icon, thumbnail, directory cache/update payloads. |
| `WM_APP + 0x350` and `0x380..0x388` | Edit suggest results and navigation menu/view actions. |
| `WM_APP + 0x400..0x460` | Pane focus/selection/file operation/function bar messages. |
| `WM_APP + 0x500..0x543` | Host alerts/prompts/settings/connections, compare directories, find, preferences, item properties, file-system command completions. |
| `WM_APP + 0x600..0x622` | Viewer async/debug messages plus ColorTextView async layout/width/ETW messages. |
| `WM_APP + 0x6F0..0x6F1` | Splash screen text/recenter. |
| `WM_USER + 0x500..0x504` | Modern combo private control messages. |

The main payload-queue policy is already visible in the codebase: receivers that can lose queued heap payloads generally call `InitPostedPayloadWindow(...)` at create time and `DrainPostedPayloadsForWindow(...)` at `WM_NCDESTROY`.

## Posted Payload Inventory

Helper-based cross-thread payload posting is used by:

- `RedSalamander/RedSalamander.cpp` for main-window settings/name payloads and compare command forwarding.
- `RedSalamander/FolderWindow.cpp`, `FolderView.cpp`, `NavigationView.cpp`, `Preferences.Dialog.cpp`, `FindFilesWindow.cpp`, `CompareDirectoriesWindow.cpp`, `FolderWindow.FileSystem*.cpp`, `FolderWindow.SelectionSize.cpp`, `FolderWindow.ItemProperties.cpp`, and `HostServices.cpp`.
- Viewer plugins: `ViewerText`, `ViewerPE`, `ViewerWeb`, `ViewerImgRaw`, `ViewerSqlite`, and `ViewerVLC`.
- `RedSalamanderMonitor/ColorTextView.cpp` for layout and width workers.
- `DirectoryInfoCache.cpp`, `SettingsHotReload.cpp`, and file-operation queue completion.

Scan note: no production match was found for raw `PostMessageW(...release())` or raw `PostMessageW(...new ...)` heap-payload posts. Raw `PostMessageW` remains in use for ordinary `WM_CLOSE`, `WM_COMMAND`, key/mouse simulation in selftests, and one file-operations popup deferred task-id message.

## UI And Window Surfaces By Area

### RedSalamander Shell

- Main app window: `RedSalamander/RedSalamander.cpp`.
- Pane composition: `FolderWindow.cpp`, `FolderView.cpp`, `NavigationView.cpp`, `FunctionBar.cpp`, `FolderWindow.StatusBar.cpp`.
- Shell and icon integration: `FolderWindow.cpp`, `FolderView.*`, `IconCache.cpp`, `FolderWatcher.cpp`, `AppDataPaths.cpp`, `AppTheme.cpp`.
- Host callbacks and marshaling: `HostServices.cpp`, `FileSystemPluginManager.cpp`, `ViewerPluginManager.cpp`.

### Dialogs, Popups, And Panels

- Preferences: `Preferences.Dialog.cpp`, `Preferences.Internal.cpp`, and `Preferences.*.cpp`.
- Compare Directories: `CompareDirectoriesWindow.cpp`, `.Options.cpp`, `.Menu.cpp`, `.Progress.cpp`.
- Find Files: `FindFilesWindow.cpp`.
- Plugin management/configuration: `ManagePluginsDialog.cpp`.
- Connections and credentials: `ConnectionManagerWindow.cpp`, `ConnectionCredentialPromptDialog.cpp`, `ConnectionSecrets.cpp`.
- Shortcuts and command surfaces: `ShortcutsWindow.cpp`, `CommandRegistry.cpp`, `CommandDispatch.Debug.h`.
- File operations: `FolderWindow.FileOperations*.cpp`, `FolderWindow.FileOperations.Popup.*`, `FolderWindow.FileOperations.IssuesPane.*`.
- Item properties and file-system command prompts: `FolderWindow.ItemProperties.cpp`, `FolderWindow.FileSystem*.cpp`, `FolderWindow.FileSystem.Commands.Part.cpp`.
- Overlays and visual helpers: `Ui/AlertOverlayWindow.cpp`, `Ui/AnimationDispatcher.h`, `SplashScreen.cpp`, `ThemedInputFrames.cpp`.

### Common/DxUi

- `Common/DxUi/DxUi.WindowHost.cpp` is the central HWND-backed DirectX host.
- `Common/DxUi/DxUi.Menu.cpp` and `DxUiNativeMenuInterop.h` bridge retained menu models and native menu semantics.
- `Common/DxUi/DxUi.NativeTextInput.cpp`, `DxUi.TextInput.cpp`, and `DxUi.TextStoreACP.cpp` own TSF/native text integration.
- `Common/DxUi/DxUi.Accessibility.cpp` owns UI Automation provider plumbing and has many manual `Release()` calls that should be treated as COM-boundary code, not assumed safe or unsafe without local review.

### Viewers

- `Plugins/ViewerText`: main viewer plus text/hex child windows, file combo host, D2D/DWrite, async open payloads.
- `Plugins/ViewerPE`: viewer window, file combo host, export file dialog, async parse payloads.
- `Plugins/ViewerSqlite`: viewer and grid-backed DxUi surface, async open/query payloads, debug messages.
- `Plugins/ViewerWeb`: WebView2 host, file combo host, save dialog, shell-open external links, async load payloads.
- `Plugins/ViewerImgRaw`: image viewer, export dialog, async decode/export payloads, D2D/DWrite overlays.
- `Plugins/ViewerVLC`: main/video/HUD/overlay/seek-preview child windows, wheel forwarding, async open payloads.
- `Plugins/ViewerSpace`: custom viewer window, D2D/DWrite, shell open parent path, host menu interop.

### File-System Plugins And Services

- `Plugins/FileSystem` is the deepest Win32 file/shell integration: `CreateFileW`, `FindFirstFileExW`, reparse handles, shell links, shell file operation COM, recycle bin integration, directory watcher.
- `Plugins/FileSystemCurl` and `Plugins/FileSystemMicrosoftDrive` use temp-file handles for streaming bodies.
- `Plugins/FileSystem7z` wraps archive DLL loading and Win32 file handles.
- `Plugins/FileSystemGoogleDrive`, `FileSystemS3`, and dummy plugins are lighter on direct Win32 but still expose COM/plugin boundaries.
- `RedSalamanderSearchService/Main.cpp` and `Common/SearchServiceBroker.cpp` use console handles, events, named pipes, and process/pipe I/O.
- `RedLauncher/Main.cpp` uses file handles, `CreateProcessW`, wait handles, and error message boxes.

### Monitor And Configure Tools

- `RedSalamanderMonitor/RedSalamanderMonitor.cpp` owns a separate main Win32 app, toolbar/status bar, native dialogs, and Dx host.
- `RedSalamanderMonitor/ColorTextView.cpp` is a heavy custom HWND with Direct3D/DXGI/DirectWrite and background layout workers.
- `RedConfigure/Main.cpp` owns its top-level Win32 window and menu; most UI is routed through RedConfigure app/root helpers.

## Resource Inventory

Production resource scan: 53 `.rc` files.

| Resource type | Count |
| --- | ---: |
| Menus | 26 |
| Dialogs | 21 |
| String tables | 71 |
| Accelerators | 2 |
| Icons | 14 |

Primary resource owners:

- `RedSalamander/RedSalamander.rc` and localized variants: main menu, dialogs, accelerators, strings, icons.
- `RedSalamander/PluginManagerResources.rc`: plugin configuration dialog resources.
- `RedSalamanderMonitor/RedSalamanderMonitor.rc` and localized variants: monitor dialogs/menu/accelerator/icon resources.
- `RedConfigure/RedConfigure.rc` and localized variants: configuration menu/icon/string resources.
- Viewer and file-system plugin resource files: plugin-local menus and string tables.

## Existing Guardrails

- WIL RAII is widely used for handles and COM (`wil::unique_*`, `wil::com_ptr`, `wil::scope_exit`).
- Custom message IDs are centralized in `Common/WindowMessages.h`.
- Posted heap payloads mostly use `PostMessagePayload(...)` and `TakeMessagePayload<T>(...)`.
- Several child/custom surfaces already handle `WM_DPICHANGED_AFTERPARENT`: `FolderView`, `NavigationView`, `FunctionBar`, `FolderWindow.StatusBar`, `Ui/AlertOverlayWindow`, and `FolderWindow.ItemProperties`.
- Production scan did not find a clear `std::thread(...).detach()` usage; `detach()` hits in production were COM pointer ownership transfers or out-params, while raw detached threads appeared only in selftest code.

## Follow-Up Audit Checklist

- [ ] Replace owned-window manual destroy calls with `_hWnd.reset()` or explicit ownership transfer:
  - `Plugins/ViewerText/ViewerText.cpp:621`
  - `Plugins/ViewerText/ViewerText.cpp:631`
  - `RedSalamander/RedSalamander.cpp:648`
  - `RedSalamander/RedSalamander.cpp:1040`
- [ ] Review long WndProc/DialogProc switches against `.github/skills/win32-wndproc/SKILL.md`: keep switch cases small, cast `WPARAM`/`LPARAM` at the boundary, and route behavior into `On*` handlers when touching these files.
- [ ] Audit manual `Release()` sites in factories, plugin async callbacks, and `Common/DxUi/DxUi.Accessibility.cpp`. Some are likely COM-boundary return semantics, but owning raw COM members or avoidable manual ref-counting should move to `wil::com_ptr`.
- [ ] Normalize raw `CoInitializeEx`/`CoUninitialize` pairs to WIL RAII where practical, especially worker and dialog helper paths. Keep explicit comments where manual COM lifetime is required.
- [ ] For every touched `PostMessagePayload(...)` sender/receiver, verify the full teardown contract: producer cancellation, UI-thread boundary, `InitPostedPayloadWindow`, `TakeMessagePayload<T>`, and `DrainPostedPayloadsForWindow`.
- [ ] Audit child/custom windows with DPI-sensitive chrome that handle only `WM_DPICHANGED`; add `WM_DPICHANGED_AFTERPARENT` where the window is child-hosted and needs parent DPI transitions.
- [ ] Keep resource edits aligned with localization rules: source `.rc` strings use positional `{0}` placeholders, and satellite translations preserve source placeholder tokens.
- [ ] Any refactor touching startup, folder rendering, viewer loading, find/search, compare, file operations, plugin I/O, or retained UI must include focused selftests and archived perf evidence under `Specs/TestRuns/` when responsiveness can change.

## Verification Notes

No build was run for this inventory because the change is documentation only. Minimum closeout verification for future code changes from this inventory should include:

- `.\build.ps1 -ProjectName RedSalamander`
- Targeted viewer/plugin/test project builds for the touched area.
- Relevant command selftest case(s) from `RedSalamander/SelfTest/Commands`.
- `git diff --check`

## Closeout Criteria

This WIP can move to `Specs/Plans/Done/` after the follow-up checklist is either completed or split into narrower WIP plans, and any durable Win32 contract changes are merged into `AGENTS.md`, `.github/skills/win32-wndproc/SKILL.md`, `.github/skills/wil-raii/SKILL.md`, or the relevant `Specs/<Domain>/` document.
