# Developer Guide

This guide is the developer entry point for the public documentation set. It now
includes **subsystem deep dives** that explain the big functional parts of the
codebase: what each does, the real files and types involved, the control flow,
the threading rules, and the gotchas worth knowing before you touch them. The
authoritative engineering rules still live in [../AGENTS.md](../AGENTS.md) and the
durable subsystem contracts live under `Specs/`; the deep dives complement those
contracts, they do not replace them.

## Repository layout

| Path | Purpose |
| --- | --- |
| `RedSalamander/` | Main dual-pane shell, dialogs, settings, command routing, and app theming. |
| `Common/` | Shared libraries, plugin interfaces, settings helpers, Win32 helpers, and `Common/DxUi`. |
| `Plugins/` | Built-in file-system and viewer plugins. |
| `Tests/` | Deterministic selftests and component test executables, including `Tests/DxUiTests`. |
| `Specs/` | Authoritative product, UI, file-system, testing, and implementation-plan specs. |
| `docs/` | User and developer documentation that can be published with GitHub Pages. |

## Build and test

Use the repository build wrapper from the repo root:

```powershell
.\build.ps1
.\build.ps1 -Configuration Release
.\build.ps1 -ProjectName RedSalamander
.\build.ps1 -ProjectName DxUi
.\build.ps1 -ProjectName DxUiTests
```

The canonical local verification command is the full test runner (it builds, then
runs the full suite); pass `-SkipBuild` to reuse an existing Debug build:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full
.\Tools\Run-AllTests.ps1 -SkipBuild
```

Useful outputs:

- `.build\x64\Debug\RedSalamander.exe`
- `.build\x64\Debug\RedSalamanderMonitor.exe`
- `.build\x64\Debug\DxUiTests.exe`
- `.build\x64\Release\DxUiTests.exe`

Run focused DxUi suites while working on the shared UI layer:

```powershell
.\build\x64\Debug\DxUiTests.exe --suite=WindowHost
.\build\x64\Debug\DxUiTests.exe --suite=Control
.\build\x64\Debug\DxUiTests.exe --suite=Grid
.\build\x64\Debug\DxUiTests.exe --suite=NativeTextInput
.\build\x64\Debug\DxUiTests.exe --suite=Accessibility
```

Generate the public theme/control gallery screenshots:

```powershell
.\build\x64\Debug\DxUiTests.exe --suite=Gallery --gallery-output-directory=docs\res
.\build\x64\Debug\DxUiTests.exe --suite=ButtonContrast --button-audit-output=docs\res\theme-button-states-after-fix.png
```

## Development rules to remember

- Use WIL RAII wrappers for Windows resources and COM pointers.
- Keep shared UI work on the UI thread unless a spec explicitly says otherwise.
- Treat perf validation as part of the feature, not a later cleanup.
- Localized user-facing strings belong in resources with positional
  `std::format` placeholders.
- Update durable specs and the relevant `docs/` page before closing a user-facing
  change.
- Do not leave completed behavior only in `Specs/Plans/WIP/` or
  `Specs/Plans/Done/`.

## Subsystem deep dives

These sections map the major functional areas of the codebase for new
contributors. Each names the concrete files, types, message IDs, and threading
rules involved, and links to the normative `Specs/` contract for that area. They
are intentionally implementation-oriented: for product behavior see the matching
user page in [the docs index](README.md), and for normative requirements see
`Specs/`. (Section names mirror the source layout; line numbers are approximate
and drift as the code changes.)

- [Application Shell, Startup & Lifecycle](#application-shell-startup--lifecycle)
- [Command Registry, Dispatch, and the FolderWindow Host](#command-registry-dispatch-and-the-folderwindow-host)
- [FolderView: File-List Model, Rendering & Interaction](#folderview-file-list-model-rendering--interaction)
- [NavigationView: Address Bar, Paths & History](#navigationview-address-bar-paths--history)
- [File Operations Engine: Queue, Scheduler & Data Safety](#file-operations-engine-queue-scheduler--data-safety)
- [Plugin Host Model & the Cross-File-System Bridge](#plugin-host-model--the-cross-file-system-bridge)
- [File-System Plugins (Local, Archive, Remote, Cloud, Object Store)](#file-system-plugins-local-archive-remote-cloud-object-store)
- [Viewer Plugin Host & Built-in Viewers](#viewer-plugin-host--built-in-viewers)
- [DxUi: The Shared DirectX UI Layer](#dxui-the-shared-directx-ui-layer)
- [Settings & SettingsStore: schema, hot reload, migration](#settings--settingsstore-schema-hot-reload-migration)
- [Find Files, Search Backends & the Search Service](#find-files-search-backends--the-search-service)
- [Compare Directories Engine & Window](#compare-directories-engine--window)
- [Batch Rename, Change Case & Rename Batching](#batch-rename-change-case--rename-batching)
- [Connection Manager, Secrets & Windows Hello](#connection-manager-secrets--windows-hello)
- [Theming (AppTheme) & the Preferences Dialog](#theming-apptheme--the-preferences-dialog)
- [Diagnostics: ETW/TraceLogging, Debug Logging & Monitor](#diagnostics-etwtracelogging-debug-logging--monitor)
- [Localization, Resources, Build & Test Infrastructure](#localization-resources-build--test-infrastructure)

### Application Shell, Startup & Lifecycle

_How RedSalamander.exe boots, instruments startup, shows a delayed splash, handles crashes, persists window placement, and shuts down cleanly._

RedSalamander.exe has no `main.cpp`; the entry point is `wWinMain` in `RedSalamander/RedSalamander.cpp` (line ~7779). It first calls `Common::MinimumOsVersion::EnsureCurrentWindowsVersionSupported`, installs the crash front door via `CrashHandler::Install()`, then runs the body inside a `__try/__except` whose filter is `CrashHandler::WriteDumpForException`. All real work happens in `RunApplication`, so any SEH escape lands in the dump path and a localized fatal-error dialog (`IDS_FATAL_ERROR_CAPTION`).

#### Boot sequence

`RunApplication` parses argv (`--help`, `--crash-test`, `--etw`, `--perf[=PATH]`, and `ENABLE_TESTS`-only `--selftest*`), opens a `Debug::Perf::Scope` named `App.Startup.UntilMessageLoop`, and calls `StartupMetrics::Initialize()` to stamp `t0`. It then, in order: `CoInitializeEx(COINIT_APARTMENTTHREADED)`; loads themes; loads settings via `Common::Settings::LoadSettingsWithRecoveryInfo`; runs `CrashQuarantine::OfferPluginDisableIfPreviousCrashDetected`; arms the delayed splash; warms `NavigationView` device resources and the icon cache on the thread pool; initializes shortcuts, file-system and viewer plugin managers; and finally `InitInstance` creates the main window. After `LoadAccelerators`, `startupPerf` is reset and the `GetMessage` loop runs. Two `wil::scope_exit` guards (`comCleanup`, `shutdownProcessSingletons`) tear down COM and process singletons (closing the splash, sweeping DxUi window hosts and the current-thread native text-input TSF manager) on every exit path.

#### Startup metrics & lifecycle markers

`StartupMetrics` (`StartupMetrics.h/.cpp`) emits four once-only ETW milestones through `Debug::Perf::Emit`, guarded by `std::atomic_flag`: `MarkFirstWindowCreated` (after `CreateWindowW` in `InitInstance`, ~7919), `MarkFirstPaint` (`OnMainWindowPaint`), `MarkInputReady`, and `MarkFirstPanePopulated` (called from `FolderView.Enumeration.cpp:1780`). `InitInstance` restores placement, sets the splash owner, calls `ShowWindow`/`UpdateWindow`, then `PostMessageW(hWnd, WndMsg::kAppStartupInputReady, 0, 0)`. `OnMainWindowStartupInputReady` marks input-ready, calls `SplashScreen::RequestCloseIfExist()`, and `CrashHandler::ShowPreviousCrashUiIfPresent`.

#### Splash, crash, session state

`SplashScreen` runs on its own `std::jthread`/`WindowHost` UI thread (`SplashScreen.cpp`); `BeginDelayedOpen(300ms, ...)` only materializes a window if startup is slow, gated by `startup.showSplash`. `CrashHandler` writes a minidump plus a marker under `%LOCALAPPDATA%\RedSalamander\Crashes\last_crash.txt`. `SessionState` writes a small marker of active file-system plugin ids per operation (Browse/Copy/Compare); on next launch `CrashQuarantine` cross-references the crash marker with `SessionState::TryRead()` to offer disabling the suspect plugin before plugins initialize.

#### Window placement & shutdown

Per-window bounds/DPI/maximized state live in `settings.windows[windowId]`. `WindowPlacementPersistence::Save/Restore` (header-only) handle auxiliary windows; the main window is saved inline in `CaptureRuntimeSettings` and restored in `InitInstance`. `WindowMaximizeBehavior::ApplyVerticalMaximize` (header-only) gives auxiliary windows a vertical-only maximize via `WM_GETMINMAXINFO`. On `WM_CLOSE`, `OnMainWindowClose` confirms file-op cancellation, then `CloseUnownedTopLevelRedSalamanderWindowsForShutdown` enumerates same-process unowned `RedSalamander.*` windows and sends `WM_CLOSE` with `SendMessageTimeoutW`, then posts `kFinalizeMainWindowCloseMessage` (`WM_APP+0x47`) which calls `DestroyWindow`. `OnMainWindowDestroy` calls `SaveAppSettings`, `SessionState::Clear()`, and `PostQuitMessage`.

#### Key files/types

| File | Role |
| --- | --- |
| `RedSalamander/RedSalamander.cpp` | `wWinMain`, `RunApplication`, `InitInstance`, WndProc, shutdown |
| `StartupMetrics.h/.cpp` | Once-only ETW startup milestones |
| `SplashScreen.h/.cpp` | Delayed splash on worker UI thread |
| `CrashHandler.h/.cpp` | Front door, minidump, next-launch UX |
| `CrashQuarantine.*` / `SessionState.*` | Plugin quarantine after a crash |
| `WindowPlacementPersistence.h` / `WindowMaximizeBehavior.h` | Window geometry helpers |

#### Threading, invariants, gotchas

The message loop, `InitInstance`, and all `OnMainWindow*` handlers are UI-thread only; warmups use `TrySubmitThreadpoolCallback`. The splash owns a separate thread, so it must be joined/closed before the DxUi host sweep (the `shutdownProcessSingletons` order is load-bearing). Closing auxiliary windows before teardown is mandatory: the Direct2D debug layer breaks if D2D objects survive to process exit. Marker files are removed before prompting to avoid repeat prompts. Spec: `Specs/Core/Core_StartupBootstrap.md`. See also [MainWindow.md](MainWindow.md), [Monitor.md](Monitor.md), and [DeveloperGuide.md](DeveloperGuide.md).

### Command Registry, Dispatch, and the FolderWindow Host

_How canonical cmd/* IDs are catalogued, bound to keyboard chords, and dispatched through the main window into FolderWindow and FolderView._

RedSalamander routes every menu click, function-key press, and configurable shortcut through a single command vocabulary: canonical string IDs of the form `cmd/area/action` (e.g. `cmd/pane/refresh`, `cmd/app/swapPanes`). This section explains how those IDs are catalogued, bound to keys, and turned into behavior.

#### Architecture

The command catalog is a `constexpr std::array<CommandInfo, 126> kCommands` in `RedSalamander/CommandRegistry.cpp`, sorted by `id` (enforced by a `static_assert`). Each `CommandInfo` maps a `wstring_view id` to a display-name string ID, a description string ID, and an optional `wmCommandId` (an `IDM_*` resource constant). A second table, `kParameterizedCommandPrefixes`, lists prefixes (e.g. `cmd/pane/hotPath/`, `cmd/app/theme/select/`) so that an ID carrying a trailing argument canonicalizes back to its base command. Lookups (`FindCommandInfo`, `TryGetWmCommandId`, `FindCommandInfoByWmCommandId`) first run `CanonicalizeCommandId` then a `std::lower_bound`.

Keyboard bindings live in `ShortcutManager` (`ShortcutManager.cpp`). `Load()` reads `Common::Settings::ShortcutsSettings` (two scopes: `functionBar` and `folderView`) and builds chord→command and reverse command→chord maps. A chord is packed by `MakeChordKey(vk, modifiers)` into `vk | (mods<<8)`, where modifiers are `kModCtrl/kModAlt/kModShift`. `ShortcutDefaults.cpp` seeds defaults via `EnsureShortcutsInitialized` and migrates deprecated IDs.

#### Control flow

The actual dispatch is **not** in a `CommandDispatch.cpp` (only `CommandDispatch.Debug.h` exists). It lives in `RedSalamander.cpp`. The message pump in `RunApplication` pre-translates keystrokes before `TranslateAccelerator`: `TryHandleShortcutKeyDown` (F1-F12 / function-bar scope) and `TryHandleFolderViewShortcutKeyDown` (folder-view scope) call `g_shortcutManager.FindFunctionBarCommand`/`FindFolderViewCommand`, then `DispatchShortcutCommand` → `ExecuteCommandById`. `ExecuteCommandById` peels off parameterized prefixes (theme, viewWith, editWith, userMenu, hotPath, goDriveRoot, ...), handles those inline, and for everything else calls `TryGetWmCommandId` and `SendMessageW(ownerWindow, WM_COMMAND, ...)`. The window proc's `OnMainWindowCommand` switch then either runs app-level handlers or calls into the global `g_folderWindow` (e.g. `IDM_VIEW_FUNCTIONBAR`, `IDM_VIEW_SWITCH_PANE_FOCUS`). Unmapped IDs reach `ShowCommandNotImplementedMessage`. The `FunctionBar` widget posts `WndMsg::kFunctionBarInvoke` to the owner; `OnFunctionBarInvoke` mirrors the keydown path. `CompareDirectoriesWindow` has parallel handlers (`DispatchShortcutCommandToCompareWindow`) that post `kCompareDirectoriesExecuteCommand` for non-WM commands.

| File / type | Role |
|---|---|
| `CommandRegistry.cpp` (`kCommands`, `CommandInfo`) | Canonical command catalog + lookups |
| `ShortcutManager.cpp` (`ShortcutManager`) | Chord↔command maps per scope |
| `ShortcutDefaults.cpp` | Default/migrated bindings |
| `RedSalamander.cpp` (`ExecuteCommandById`, `DispatchShortcutCommand`, `OnMainWindowCommand`) | Dispatch core + WM_COMMAND host |
| `FunctionBar.cpp` (`FunctionBar`) | Function-key bar UI; posts `kFunctionBarInvoke` |

#### Threading and invariants

All dispatch and `g_folderWindow` access is UI-thread only; the FolderWindow host is accessed through the global `g_folderWindow` guarded by `g_hFolderWindow` (an atomic checked with `memory_order_acquire`). Key invariants: `kCommands` must stay sorted by `id`; every parameterized prefix's `canonicalId` must exist in `kCommands` (both `static_assert`-enforced). Edit-control focus (`IsEditControlFocused`) bypasses shortcut and accelerator handling for text-edit safety, per `Specs/UI/UI_CommandMenuKeyboard.md`. Unbound F1-F12 are deliberately consumed so reconfigured shortcuts never fall through to stale accelerators.

#### Extension points and gotchas

To add a command: append a sorted `CommandInfo` row, add `IDS_CMD_*`/`IDS_CMD_DESC_*` resources (short-name strings use the `IDS_CMD_SHORT_BASE` offset), and either set a `wmCommandId` with an `OnMainWindowCommand` case or handle it inline in `ExecuteCommandById`. Gotchas: the canonical spec is `Specs/UI/UI_CommandMenuKeyboard.md` (`UI_KeyboardManagement.md` is a redirect stub); a `wmCommandId` of `0` means "handled inline, no menu/accelerator"; and self-tests reach commands via `DebugDispatchShortcutCommand` in `CommandDispatch.Debug.h`. See [MainWindow.md](MainWindow.md) for the user-facing surface and `Specs/UI/UI_FolderWindow.md` for pane semantics.

### FolderView: File-List Model, Rendering & Interaction

_The per-pane Direct2D file browser: async plugin enumeration into a zero-copy item model, virtualized column-grid rendering, and all mouse/keyboard/drag interaction._

`FolderView` is the per-pane file browser. One instance lives inside each `FolderWindow` pane below its `NavigationView`. It is a custom Win32 child window (class `RedSalamanderFolderView`) that renders with Direct3D 11 + Direct2D on a DXGI flip swap chain. The class is large and split across ~12 translation units; the public surface is `RedSalamander/FolderView.h` and shared internals live in `RedSalamander/FolderViewInternal.h`. The authoritative behavior spec is `Specs/UI/UI_FolderView.md` (grid math, theming, drag-drop, thumbnails, empty/error states); shared-grid invariants are in `Specs/UI/UI_DxUiSharedGrid.md`.

#### Key files & types

| File / type | Responsibility |
|---|---|
| `FolderView.cpp` | Window lifecycle + `WndProc` message dispatch |
| `FolderView::FolderItem` (in `.h`) | One row; `displayName` is a zero-copy `wstring_view` into the arena |
| `FolderView.Enumeration.cpp` | Background `EnumerationWorker` / `ExecuteEnumeration`, sort, cache refresh |
| `FolderView.Rendering.cpp` | D3D/D2D/DWrite resources, `Render`, `DrawItem` |
| `FolderView.Layout.cpp` + `FolderViewColumnLayout.h` | Grid layout, `HitTest`, `GetVisibleItemRange`, `EnsureVisible`, scroll metrics |
| `FolderView.Interaction.cpp` | Mouse/keyboard/scroll, incremental search |
| `FolderView.Icons.cpp` | Async icon + thumbnail loading, UI-thread bitmap creation |

#### Data & control flow

`SetFolderPath` -> `EnumerateFolder` bumps `_enumerationGeneration` and posts the path to the `_enumerationThread` (a `std::jthread`). `ExecuteEnumeration` calls `IFileSystem::ReadDirectoryInfo` (via `DirectoryInfoCache`), walks the `FileInfo` buffer by `NextEntryOffset`, builds `FolderItem`s pointing into the COM `IFilesInformation` arena (kept alive in `EnumerationPayload::arenaBuffer`), applies hidden/system/name-filter/hidden-names filtering, sorts directories-first, and resolves icon indices in parallel via the Windows thread pool and `IconCache`. The payload is delivered to the UI thread by `PostMessagePayload(... kFolderViewEnumerateComplete)`; `ProcessEnumerationResult` drops stale generations, then swaps `_items`/`_itemsArenaBuffer`, preserves selection/rendering state for surviving names on refresh, restores remembered focus, and triggers `LayoutItems`. `LayoutItems` (using `FolderViewColumnLayout::Resolve`) assigns items to top-to-bottom columns with per-column variable width; `Render` clips to the dirty rect, draws only `GetVisibleItemRange`, and creates DWrite layouts lazily. Icons/thumbnails are queued for visible items first (`QueueIconLoading`/`QueueThumbnailLoading`), extracted off-thread, and converted to `ID2D1Bitmap1` back on the UI thread (`kFolderViewCreateIconBitmap`/`kFolderViewCreateThumbnailBitmap`).

#### Threading rules

The window, all `_items` access, D2D rendering, and bitmap creation are UI-thread only. The single `_enumerationThread` initializes COM as MTA and performs enumeration, icon extraction, and shell/WIC thumbnail decode. Cross-thread results always arrive as posted messages; `_enumerationGeneration` (atomic) is the staleness fence checked everywhere. `_enumerationMutex`/`_enumerationCv` gate worker wakeups; `_errorOverlayMutex` guards overlay state read on both threads; `_nameFilter`/`_hiddenNames` are atomic `shared_ptr`s so the worker reads a stable snapshot.

#### Invariants & gotchas

- DirectX init is deferred until after first paint via `kFolderViewDeferredInit`; before then the first `WM_PAINT` does a GDI fill and `QueueIconLoading` no-ops until `_d2dContext` exists, re-running from `OnDeferredInit`.
- `FolderItem::displayName` is only valid while `_itemsArenaBuffer` lives; never outlive the arena with a view.
- Per-column widths are variable: rendering, hit testing, `EnsureVisible`, and horizontal scroll must use `_columnLayout`, never a global stride (see `Specs/UI/UI_DxUiSharedGrid.md`). Horizontal scroll snaps to column stops on release.
- Focus and selection are independent; refresh transfers selection by name (with rename-chain hints).

#### Extension points

Host wiring is via callbacks on `FolderView` (`SetFileOperationRequestCallback`, `SetDetailsTextProvider`, `SetViewFileRequestCallback`, etc.); file operations MUST delegate through `FileOperationRequestCallback` rather than calling plugins directly. See [FileOperations.md](FileOperations.md) and [MainWindow.md](MainWindow.md) for the user-facing pane behavior.

### NavigationView: Address Bar, Paths & History

_The breadcrumb/address bar at the top of each pane: a Direct2D-rendered Win32 child that parses, displays, edits, and navigates locations and feeds the shared folder history._

`NavigationView` is the navigation bar drawn above each FolderView pane. It is a single Win32 child window class (`RedSalamander.NavigationView`) that renders four sections with Direct2D/DirectWrite to a per-window DXGI swap chain: Drive/Menu button, Path breadcrumb, History button, and Disk Info. It owns address-bar edit mode, the full-path popup, plugin-sourced dropdowns, and the path-string model that bridges the UI and the file-system plugins.

#### Key files and types

| File | Responsibility |
|------|----------------|
| `RedSalamander/NavigationView.cpp` | Window lifecycle, `WndProc` dispatch, `SetPath`/`SetHistory`, model state |
| `RedSalamander/NavigationViewInternal.h` | Private helpers: path sniffing, `TryParseEditSuggestQuery`, layout/measure utilities, `NavigationDxTextHost` |
| `RedSalamander/NavigationLocation.h` | Pure header: `Location`, `TryParseLocation`, `FormatHistoryPath`, `FormatEditPath`, `NormalizePluginPath*` |
| `RedSalamander/NavigationView.Edit.cpp` | `EnterEditMode`/`ExitEditMode`, validation popup, edit-suggest worker thread |
| `RedSalamander/NavigationView.Menus.cpp` | `ShowMenuDropdown`/`ShowHistoryDropdown`/`ShowSiblingsDropdown` (DxUi popups) |
| `RedSalamander/NavigationView.FullPathPopup.cpp` | Overflow ellipsis popup window |

#### The path model (the central contract)

`NavigationLocation` defines two strings per location. **pluginPath** is what plugins (`IFileSystem`, `IDriveInfo`) receive: a Windows absolute path for `file`, or a `/`-rooted path otherwise — never a `<shortId>:` prefix or a `|` mount delimiter. **canonical/history path** adds those back: `<shortId>:<pluginPath>` or `<shortId>:<instanceContext>|<pluginPath>`. `SetPath` runs `TryParseLocation`, then computes three stored optionals — `_currentPluginPath` (breadcrumb display), `_currentEditPath` (`FormatEditPath`, what edit mode shows), `_currentPath` (`FormatHistoryPath`, the canonical identity). This split is the most important thing to understand; `Specs/UI/UI_NavigationView.md` "Path Model" is authoritative.

#### Control flow

The host wires `SetPathChangedCallback` and `SetRequestFolderViewFocusCallback`. Display updates flow **in** via `SetPath`/`SetHistory`; navigation requests flow **out** via `RequestPathChange`, which calls `_pathChangedCallback` (FolderWindow re-navigates and calls `SetPath` back). Accepting an edited path: `ExitEditMode(true,...)` normalizes the typed text (`NormalizeUserTypedLocationText`, env-var expansion, quote stripping), runs `ValidatePath`, and on success calls `RequestPathChange`; on failure it keeps edit mode, shows a themed DirectWrite warning popup, and preserves the typo. Plugins can request navigation through `NavigationMenuRequestNavigate`, which posts `WndMsg::kNavigationMenuRequestPath`.

#### Threading and UI-thread rules

Everything visible is UI-thread only. Two background `std::jthread`s exist: the **edit-suggest worker** (enumerates directory children for autocomplete, posts results back via `WndMsg::kEditSuggestResults` with `request_stop`/CV cancellation) and the **sibling-prefetch worker**. `Destroy`/`OnDestroy` join both and clear pending queries. Dropdown opening is deferred through posted messages (`kNavigationViewShowHistoryDropdown`, `kNavigationViewShowMenuDropdown`, `kNavigationMenuShowSiblingsDropdown`) so the popup opens from the resolved input branch.

#### Invariants and gotchas

- The child WndProc must return `HTCLIENT`/`MA_ACTIVATE` so breadcrumb clicks reach it directly; `NavigationDxTextHost` edit children must self-retire (`HTTRANSPARENT`) when edit mode is inactive.
- Input routing must use delivered message `lParam`, never `GetCursorPos()` (diagnostics only). The hover timer runs only inside menu loops.
- `SetPath`/`SetHistory` short-circuit identical input but must still rebuild on case-only changes; D2D init is deferred until first paint (`kNavigationViewDeferredInit`).
- Edit text flows only through the retained DxUi `TextField`; do not introduce a native `EDIT` or subclass the bridge.

See `Specs/UI/UI_NavigationView.md` and `Specs/UI/UI_FolderWindow.md`; user docs are [NavigationAndPaths.md](NavigationAndPaths.md).

### File Operations Engine: Queue, Scheduler & Data Safety

_Background copy/move/delete execution: per-task worker threads, a Wait/Parallel start-gating queue, a shared per-item scheduler, host-driven pre-calc and cross-FS bridging, and inline conflict/data-safety arbitration._

This subsystem runs every long file operation (Copy `F5`, Move `F6`, Delete `Del`/`Shift+Del`, clipboard paste, pane drag-drop, pack/unpack, Find Files commands) off the UI thread and surfaces it in the File Operations popup. The entry point is `FolderWindow::FileOperationState::StartOperation(...)`, which validates preconditions (provider capabilities via `CanSameFileSystemOperation`, same-folder/overlap rejection, destructive confirmation prompts), snapshots host-owned settings (pre-calc, cross-FS bridge buffer, default bandwidth) into a `Task`, then spawns one `std::jthread` per task running `Task::ThreadMain`.

#### Key files and types

| File | Role |
|------|------|
| `FolderWindow.FileOperationsInternal.h` | `FileOperationState`, `Task` (impl `IFileSystemCallback` + `IFileSystemDirectorySizeCallback`), `ConflictArbiter`, `CompletedTaskSummary`, `TaskDiagnosticEntry` |
| `FolderWindow.FileOperations.State.Runtime.Part.cpp` | `StartOperation`, `ApplyQueueMode`, `CancelAll`, `NotifyQueueChanged`, `Shutdown` |
| `FolderWindow.FileOperations.State.Queue.Part.cpp` | `EnterOperation`, `LeaveOperation`, `UpdateQueuePausedTasks`, `PostCompleted` |
| `FolderWindow.FileOperations.State.cpp` | `ThreadMain`, `ExecuteOperation`, `RunPreCalculation`, callbacks, `PerItemTaskScheduler`, conflict arbiter |
| `FolderWindow.FileOperations.State.Diagnostics.Part.cpp` | `RecordCompletedTask`, diagnostics log + Issues capture |
| `FolderWindow.FileOperations.Popup.cpp` / `.IssuesPane.cpp` | D2D/DWrite UI; reads snapshots, posts user actions |

#### Control flow

`ThreadMain` enters the queue first via `EnterOperation`: if `_waitForOthers`, the task FIFO-blocks on `_queueCv` until `_activeOperations == 0` and it is the queue head (Wait mode); otherwise it bumps `_activeOperations` and proceeds (Parallel). Pre-calc runs while holding the slot. For Copy with pre-calc enabled, "5F early admission" runs `RunPreCalculation` on a side `std::jthread` (via `TryStartPreCalculationThread`) concurrently with the transfer so bytes move before the recursive scan finishes; Move/Delete keep serial pre-calc-then-execute because they mutate the source. `ExecuteOperation` either calls the same-context `IFileSystem` bulk/per-item APIs or, when `_destinationFileSystem` is set, drives the host cross-filesystem bridge. Per-item work for `maxConcurrency > 1` is dispatched through the process-wide `PerItemTaskScheduler` (`GetPerItemTaskScheduler()`), a shared bounded worker pool issuing short `processIndex(i)` work items so one large operation never pins every worker. Completion calls `PostCompleted`, which sets `_taskFinished`, builds a `CompletedTaskSummary` (`RecordCompletedTask`), and posts `WndMsg::kFileOperationCompleted` to the FolderWindow (`OnFileOperationCompleted`).

#### Threading rules

`IFileSystem::*` and all callbacks run on worker threads; the popup updates on a ~100ms timer reading published atomics/snapshots. Per the "Popup Windowing And Locking Contract", never call `ShowWindow`/`SetWindowPos` while holding `_mutex`. Callbacks for one logical operation are serialized; parallel streams are identified by `(cookie, progressStreamId)`.

#### Data safety and conflicts

`FileSystemIssue` classifies failures into `ConflictBucket`s and routes them through the `ConflictArbiter`: `BeginConflictPrompt` enforces at most one active prompt per task (`prompt.active`), `WaitForConflictDecision` blocks the worker on `decisionEvent`, and the popup posts back via `Task::SubmitConflictDecision`. Invariants: no implicit overwrite; Overwrite/ReplaceReadOnly grants are one-shot unless "Apply to all"; Retry is capped to one per `(item, bucket)` and never cached; directory-vs-directory is a merge, not a conflict; cross-volume Move partial failure surfaces `ERROR_PARTIAL_COPY` with "source preserved" diagnostics. Cancel sets `_cancelled` and wakes every wait primitive; pause blocks inside callbacks via `WaitWhilePaused`. The Riptide/Floodgate plans (`Specs/Plans/WIP/Operation_Riptide_*.md`, `Operation_Floodgate_*.md`) track remediation of silent data-loss and conflict-routing parity.

#### Extension points and gotchas

New providers must implement `IFileSystem::GetCapabilities()` (fail-closed) and emit serialized callbacks with stable stream IDs. `--fileops-selftest` cases are mandatory for any change touching liveness/data safety (see `Specs/FileSystem/FileSystem_FileOperations.md`). Gotcha: `_perf` `FileOps.*` metrics must be maintained; `ENABLE_TESTS` hooks (bridge fault injection, post-finished pause) gate deterministic tests. See [FileOperations.md](FileOperations.md) and `Specs/Plugins/Plugins_VirtualFileSystem.md`.

### Plugin Host Model & the Cross-File-System Bridge

_How RedSalamander discovers, loads, and drives COM-style file-system plugins, exposes host services back to them, and copies/moves files between two different providers via a capability-gated streaming bridge._

RedSalamander's file managers never touch Win32 file APIs directly. Every pane talks to an `IFileSystem` instance produced by a plugin DLL. This section covers the three pillars that make that work: plugin discovery/loading, the host-services object plugins call back into, and the bridge that moves bytes between two unrelated providers.

#### Discovery and loading

`FileSystemPluginManager` (`RedSalamander/FileSystemPluginManager.cpp`, a process singleton via `GetInstance()`) owns the plugin set. `Discover()` builds a candidate list from three origins (`PluginOrigin::Embedded` for `Plugins\FileSystem.dll`/`ViewerText.dll` next to the exe, `Optional` for every DLL under `Plugins\`, and `Custom` from `settings.plugins.customPluginPaths`). For each DLL it calls the exported `RedSalamanderEnumeratePlugins` (see `Common/PlugInterfaces/Factory.h`) to read `PluginMetaData` rows — a single DLL may advertise several logical plugins. `EnsureLoaded()` then `LoadLibraryExW`s the module, resolves `RedSalamanderCreate`, and constructs the instance, passing `GetHostServices()` as the `IHost*`. It `QueryInterface`s for `IInformations`, validates that the reported `id`/`shortId` match what enumeration advertised (mismatch fails the entry), and applies persisted JSON config. Duplicate `id` or `shortId` values are rejected. One plugin is "active" per `_activePluginId`; `SetActivePlugin()` flips it and mirrors it into `settings.plugins.currentFileSystemPluginId` and `SessionState`.

Unload is centralized in `Unload()`: release COM instances first, then call the optional `RedSalamanderPluginShutdown` quiet-point export, unregister the localization resource owner, and — only on `ModuleUnloadMode::ProcessShutdown` — honor `RedSalamanderPluginRetainModuleUntilProcessExit` by leaking the `HMODULE` so process teardown does not race driver cleanup.

#### Host services

`GetHostServices()` (`RedSalamander/HostServices.cpp`) returns a never-deleted singleton implementing `IHost` plus `IHostAlerts`, `IHostPrompts`, `IHostConnections`, `IHostPaneExecute`, and `IHostViewers` (all in `Common/PlugInterfaces/Host.h`), discoverable by `QueryInterface`. The hard rule: these methods can be called on any plugin worker thread, but the UI work must run on the FolderWindow thread. Each method checks `IsCurrentThreadWindowThread(g_hFolderWindow)`; if off-thread it marshals a heap payload to the window. Fire-and-forget calls (`ShowAlert`, `ClearAlert`, `ExecuteInActivePane`) use `PostMessagePayload`; result-returning calls (`ShowPrompt`, connection-secret and viewer calls) use blocking `SendMessageW`. `TryHandleHostServicesWindowMessage` (called from the window proc) unpacks the payload via the `WndMsg::kHost*` IDs and runs the real `*OnUiThread` handler. All request structs carry `version`/`sizeBytes` and are rejected with `E_INVALIDARG` on mismatch — the ABI evolution contract in `Specs/Plugins/Plugins_PluginAPI.md`.

#### The cross-file-system bridge

When a copy/move crosses providers (source and destination panes resolve different `IFileSystem` instances), the host cannot delegate to one plugin. `FolderWindow.FileOperations.cpp` first gates the operation through `CanCrossFileSystemCopyMove()`, which parses each provider's mandatory `GetCapabilities` JSON (`crossFileSystem.export`/`import` plugin-id allow-lists, plus `read`/`write`/`delete` flags) — a missing or unparseable document fails closed via `ReportCapabilitiesContractViolationOnce`. The actual transfer is the `CrossFileSystemBridge` class in `FolderWindow.FileOperations.State.cpp`: it `QueryInterface`s `IFileSystemIO` on both providers, opens an `IFileReader` on the source and an `IFileWriter` on a CSPRNG-named temp path on the destination, pumps a buffer (sized by `ResolveAdaptiveCrossFsBridgeBufferBytes` from `GetTransferHints`/`GetStorageCharacteristics` plus `crossFsBridgeBufferSizeKB`), `Commit()`s, then `PromoteTempToFinalPath` renames into place for atomic-commit semantics. It supports serial and producer/consumer pipeline modes, carries one-shot overwrite grants, and reports progress through `BridgeCallback` with per-worker `progressStreamId`s.

#### Key files/types

| File / type | Role |
|---|---|
| `FileSystemPluginManager.cpp` | Discover/load/activate/unload plugins |
| `HostServices.cpp` (`HostServices`) | `IHost` services, thread-marshaled UI |
| `FolderWindow.FileOperations.cpp` | Capability parsing + same/cross-FS gating |
| `CrossFileSystemBridge` (State.cpp) | Streaming reader→temp→promote transfer |
| `FileActionResolver` / `FileActionLauncher` | Resolve and launch external view/edit actions |

#### Gotchas

Capabilities are fail-closed: a plugin that returns `E_NOTIMPL` from `GetCapabilities` gets all gated operations disabled. Never call host-services UI methods expecting them to run synchronously off-thread — prompts block via `SendMessage`, alerts don't. See `Specs/Plugins/Plugins_VirtualFileSystem.md` for the `pathIdentity` and capability JSON contracts and [Plugins.md](Plugins.md) for the user-facing plugin list.

### File-System Plugins (Local, Archive, Remote, Cloud, Object Store)

_COM-style IFileSystem plugin DLLs that back every address-bar prefix (file:, 7z:, ftp/sftp/scp/imap:, s3/s3table:, gdrive:, onedrive/sharepoint:, fk:), sharing one buffer-based enumeration ABI and a JSON capability contract._

Every storage backend reachable in the address bar is a plugin DLL under `Plugins/`. They all implement the COM-style `IFileSystem` (and optional sibling interfaces) defined in `Common/PlugInterfaces/FileSystem.h`, so the host (FolderView/FolderWindow) treats local disks, archives, remote servers, cloud drives, and object stores uniformly. The normative ABI is `Specs/Plugins/Plugins_VirtualFileSystem.md`; per-provider behavior lives in `Specs/FileSystem/FileSystem_*.md`.

#### Common skeleton
Each DLL exports three C entry points (see any `Factory.cpp`): `RedSalamanderEnumeratePlugins` (lists `PluginMetaData` for the DLL — one DLL can expose several plugins, e.g. `FileSystemCurl` returns FTP/SFTP/SCP/IMAP and `FileSystemS3` returns `s3`/`s3table`), `RedSalamanderCreate` (instantiates the right class by `pluginId`), and `RedSalamanderGetConfigurationSchema`. The provider class implements `IFileSystem` plus optional `IFileSystemIO`, `IFileSystemDirectoryOperations`, `IFileSystemDirectoryWatch`, `IFileSystemSearch`, `IInformations`, `INavigationMenu`, and `IDriveInfo`; the host discovers what each instance supports via `QueryInterface`. `IFileSystemInitialize` is used by `FileSystem7z` to bind a mounted archive path (e.g. `7z:C:\x.zip|/`).

#### Enumeration ABI
`ReadDirectoryInfo` returns an `IFilesInformation` holding one contiguous buffer of `FileInfo` records linked by `NextEntryOffset` (NOT NUL-terminated names — use `FileNameSize`). The built-in `FileSystem` streams Win32 `FindFirstFile`/`GetFileInformationByHandleEx` directly into the buffer; remote/cloud/archive providers build a `std::vector<Entry>` then call a `BuildFromEntries` packer (e.g. `FilesInformationCurl`, `FilesInformationS3`). Once returned the buffer is immutable.

#### Capabilities and the host bridge
`GetCapabilities` MUST return a UTF-8 JSON doc with `version`, `operations`, `concurrency`, `crossFileSystem`, and `pathIdentity`.

A failed/missing/unparseable capability doc (`version`, `operations`, `concurrency`, `crossFileSystem`, `pathIdentity`) disables capability-gated ops. Missing/invalid/unstable `pathIdentity` is fail-closed for identity-sensitive planners such as Batch Rename and same-context Copy/Move/Delete. Read-only providers advertise `copy/move/delete=false` (7z, S3 Table, Google Drive). Cross-filesystem copy/move is performed by the *host* bridge using `IFileSystemIO::CreateFileReader`/`CreateFileWriter` (read→write), so even `s3:`→`file:` works when both sides allow it; `GetTransferHints`/`GetStorageCharacteristics` feed buffer sizing and auto-concurrency.

#### Threading and UI rules
Plugin calls run on host worker threads, never the UI thread. Per-call `IFileSystemCallback` (progress/cancel/`FileSystemIssue` conflict prompts) may block and may run on background threads; plugins MUST NOT invoke it after the op returns, MUST NOT call it concurrently for one op, and parallel workers MUST use distinct `progressStreamId`. Watch callbacks must use PostMessage (never SendMessage) and `UnwatchDirectory` MUST synchronously drain in-flight callbacks without holding the delivery lock — remote/cloud providers synthesize watch events via `SyntheticWatchRegistration`.

#### Key files / types
| File / type | Role |
|---|---|
| `Common/PlugInterfaces/FileSystem.h` | `IFileSystem`, `IFilesInformation`, `FileInfo`, callbacks, `FileSystemArena` |
| `Plugins/FileSystem/FileSystem*.cpp` | local Win32 provider (`builtin/file-system`), search, watch, recycle bin |
| `Plugins/FileSystem7z/FileSystem7z.cpp` | archive mount via `IFileSystemInitialize`, async index build |
| `Plugins/FileSystemCurl/` | FTP/SFTP/SCP (`.CopyMove`,`.DirectoryOps`) + IMAP (`.Imap`) |
| `Plugins/FileSystemS3/` | S3 + S3 Table via AWS S3Crt client |
| `Plugins/FileSystemGoogleDrive`, `FileSystemMicrosoftDrive` | Graph/Drive REST over curl + token cache |
| `Plugins/FileSystemDummy/FileSystemDummy.cpp` | deterministic in-memory FS for tests |

#### Extension points and gotchas
To add a provider, copy the `Factory.cpp` skeleton, implement `IFileSystem` + the optional interfaces you support, return correct capability JSON, and pack entries via a `FilesInformation*` helper. Secrets resolve through `IHostConnections` (`GetConnectionSecret`/`PromptForConnectionSecret`) for `/@conn:<name>/...` paths; plugin-settings passwords are plain text. Directory-vs-directory existence is a *merge*, not an `ERROR_ALREADY_EXISTS` conflict; overwrite grants are one-shot per child. Cloud providers (Microsoft/Google) double-buffer their configuration JSON so previously returned pointers stay valid. See [Plugins.md](Plugins.md), [RemoteFileSystems.md](RemoteFileSystems.md), [CloudDrives.md](CloudDrives.md), [S3AndS3Table.md](S3AndS3Table.md), and `Specs/FileSystem/FileSystem_FileOperations.md`.

### Viewer Plugin Host & Built-in Viewers

_How RedSalamander discovers, loads, instantiates, and hosts COM-style viewer plugins (standalone F3 windows and embedded preview), and how the built-in viewers implement the IViewer contract._

RedSalamander opens files in viewers through a COM-style plugin contract (`IViewer` + `IInformations`), with both a discovery/lifetime manager (`ViewerPluginManager`) and per-window hosting logic in `FolderWindow`. The seven built-in viewer DLLs all implement the same factory and interface contracts.

#### Architecture & key types

| File / type | Role |
| --- | --- |
| `Common/PlugInterfaces/Viewer.h` (`IViewer`, `IViewerCallback`, `ViewerOpenContext`, `ViewerTheme`) | The ABI. `IViewer` UUID `d1da10b7-...`; `ViewerTheme.version == 4`. |
| `RedSalamander/ViewerPluginManager.{h,cpp}` (singleton) | Discovery, `LoadLibraryExW` load, factory invocation, enable/disable, JSON config, unload. |
| `RedSalamander/FolderWindow.Viewers.cpp` (`ViewerInstance`, `ViewerCallbackState`) | Per-window viewer lifecycle: open/reopen, theme push, callback wiring, preview pane hosting. |
| `Common/EmbeddedViewerBase.h` (`EmbeddedViewerBase<Derived>`) | Shared base for embedded-capable viewers: callback drain, theme storage, HWND-reuse check. |
| `Plugins/Viewer*/Factory.cpp` | `RedSalamanderCreate` / `RedSalamanderEnumeratePlugins` exports per DLL. |

#### Discovery and loading

`ViewerPluginManager::Discover` (called from `Initialize`/`Refresh`) builds a candidate list: embedded DLLs hard-coded under `<exeDir>\Plugins` (`ViewerText.dll`, `ViewerSpace.dll`, `ViewerImgRaw.dll`), every DLL in that `Plugins` dir (`PluginOrigin::Optional`), then `settings.plugins.customPluginPaths`. For each it probes `RedSalamanderEnumeratePlugins(IID_IViewer, ...)`: present and non-empty means a multi-plugin DLL (e.g. `ViewerWeb` exposing `builtin/viewer-web`, `-json`, `-markdown`), and one `PluginEntry` per `PluginMetaData` is created with `factoryPluginId` set. `EnsureLoaded` does the real load, calls `RegisterResourceOwner` for localization, reads metadata (via enumerate or by creating an instance and `QueryInterface(IInformations)->GetMetaData`), and validates id/shortId uniqueness. Discovery is a health gate: zero loadable plugins returns `ERROR_NOT_FOUND`.

#### Open / control flow

`FolderWindow::TryViewFileWithViewer` (F3/`IDM_PANE_VIEW`) resolves a plugin id via `FileActionResolver::ResolveViewerAction` against `settings.fileActions.viewers`, defaulting to `builtin/viewer-text`. It builds a `ViewerOpenContext` (focused path, selection, `otherFiles` peers, owner HWND) and calls `OpenViewerWithPluginInternal`, which: calls `ViewerPluginManager::CreateViewerInstance` (factory + `ApplyConfigurationFromSettings` from `plugins.configurationByPluginId`), wraps it in a `ViewerInstance` stored in `_viewerInstances`, then calls `SetTheme(BuildViewerTheme())`, `SetCallback(&_viewerCallback, cookie)`, and `Open(&openContext)`. On open failure the callback is cleared and `Close()` called. The preview pane path (`OpenPreviewFocusedPathWithViewer`) sets `VIEWER_OPEN_FLAG_EMBEDDED`, consults `DefaultViewerFileActionsSettings` when only the text default matched, and reuses the live instance via `ReopenViewerInstance` (calling `Open()` again, never `Close()`) when the same embedded-capable plugin (`SupportsEmbeddedPreviewViewer`) resolves.

#### Threading / UI-thread rules

All host viewer calls run on the UI thread. The fragile contract is the callback drain: `IViewer::SetCallback(nullptr, nullptr)` synchronously drains in-flight `ViewerClosed` deliveries, so `OnViewerClosed` deliberately does **not** clear its own callback (that would deadlock). `ShutdownViewers` and `ClosePreviewViewer` clear callbacks before `Close()` from the host side. Plugins must copy all `ViewerOpenContext`/`ViewerTheme` pointers (host-owned, ephemeral past `Open()`/`SetTheme()`) and `AddRef` `fileSystem` if retained.

#### Invariants, extension points, gotchas

Each `IViewer` must also expose `IInformations`; ids must be unique and `shortId` alphanumeric. Theme is push-only — plugins must never read global app state; `ApplyViewerTheme` re-pushes to all live instances on theme change. Config round-trips via `IInformations::SomethingToSave`/`GetConfiguration`; `PersistViewerConfiguration` only writes when JSON differs from the initial snapshot. Optional `RedSalamanderPluginShutdown` / `RedSalamanderPluginRetainModuleUntilProcessExit` exports handle driver-teardown hazards (e.g. graphics-backed viewers leaving the module pinned at process exit; see `ViewerPluginManager::Unload`). Gotcha: multi-plugin DLLs must back `PluginMetaData` strings with DLL-lifetime storage, not temporaries.

See `Specs/Plugins/Plugins_ViewerPlugins.md` and per-viewer specs (`Specs/Plugins/Plugins_ViewerText.md`, etc.). User-facing reference: [Viewers.md](Viewers.md); related host UI in `Specs/UI/UI_FolderWindow.md`.

### DxUi: The Shared DirectX UI Layer

_Retained, Direct2D/DirectWrite UI toolkit in Common/DxUi: a WindowHost binds an HWND to a shared D3D11/D2D device and drives a tree of Control objects with built-in focus, theming, text input, accessibility, and a frame-timed animation dispatcher._

DxUi is RedSalamander's shared retained UI toolkit. It moves interactive chrome off native Win32/comctl controls onto one Direct3D 11 + Direct2D + DirectWrite path: a `WindowHost` attaches to a caller-owned `HWND`, owns the DirectX device/swap-chain, and drives a retained tree of `Control` objects with built-in focus, theming, text input, accessibility, tooltips, pointer capture, high-DPI, and animation. The public surface lives in `Common/DxUi/DxUi.h`; private helpers in `Common/DxUi/DxUi.Internal.h`. The authoritative behavior contracts are `Specs/UI/UI_DxUiSharedGrid.md`, `Specs/UI/UI_DxUiWinUIDesign.md`, and `Specs/UI/UI_VisualStyle.md`. A substantial usage guide already exists at [DxUi.md](DxUi.md); this section adds the architectural model behind it.

#### Key files and types

| File | Responsibility |
| --- | --- |
| `DxUi.h` | Public API: `Control`, `Panel`, `WindowHost`, `ThemePalette`, all controls, grid/tree interfaces. |
| `DxUi.WindowHost.cpp` | HWND attach/detach, shared device lifecycle, `HandleMessage` routing, render, focus, DPI, animation subscription. |
| `DxUi.Controls.cpp` / `DxUi.ComboBox.cpp` | Panels, labels, buttons, toggles, tabs, status strips; combo boxes. |
| `DxUi.Grid.cpp` / `DxUi.Tree.cpp` | Virtualized data surfaces over `IDxGridModel`/`IDxTreeModel`. |
| `DxUi.TextInput.cpp` / `DxUi.NativeTextInput.cpp` / `DxUi.TextStoreACP.cpp` | `TextField` editing, IME/TSF, caret windows. |
| `DxUi.Theme.cpp` | `MakeDefaultThemePalette`, `MakeThemePaletteFromViewerTheme`, resolved visual styles. |
| `DxUi.Accessibility.cpp` | `AccessibilityProvider` UIA fragment root answering `WM_GETOBJECT`. |
| `DxUi.FrameRuntime.h/.cpp` | `FrameClock`, `FrameStage`, `MotionPolicy`, hitch-clamp smoothing. |
| `RedSalamander/Ui/AnimationDispatcher.h` | Shared 8 ms `WM_TIMER` dispatcher (note: lives in the app, not `Common/DxUi`). |

#### Control tree and rendering

`Control` is the abstract base (`Paint`, `Tick`, `OnMouse*`/`OnKey*` virtuals, bounds in DIPs). `Panel::AddChild<T>(...)` owns children via `unique_ptr`; `WindowHost::SetRoot(...)` takes the root. Each control holds a `_lifetimeToken` (`shared_ptr<int>`) so async callbacks can hold a `weak_ptr` and detect destruction. `WindowHost` keeps only non-owning observers (`_focusedControl`, `_hoveredControl`, `_capturedControl`, `_defaultButton`, `_cancelButton`) — these are pruned via `PruneStaleInteractionState`/`ResetRootInteractionState` before the tree changes. Painting goes through `WindowHost::Render`, which uses cached `IDWriteTextFormat` per `FontRole` and an `ID2D1SolidColorBrush` cache keyed by packed color (`GetSolidBrush`).

#### Shared device model

All hosts share one D3D11/D2D/DXGI/DirectWrite stack via `GetSharedWindowHostGraphicsResources()` keyed by a `generation`. `EnsureDeviceResources` rebinds when the generation changes (device loss); `Attach`/`Detach` ref-count the shared bucket (`Register/ReleaseSharedWindowHostAttachment`) so the last detach tears it down — important for plugin smoke tests and clean process exit. Per the spec, this device singleton is single-UI-thread-owned and must fail fast on cross-thread access.

#### Threading and frame rules

Everything — render, hit-test, focus, capture, animation — is UI-thread-only. Worker threads must post payloads to the UI owner, then mutate models and call `NotifyDataChanged`. Animation is cooperative: a control calls `WindowHost::RequestAnimation()`, which subscribes to `Ui::AnimationDispatcher` (a hidden `HWND_MESSAGE` window pumping `WM_TIMER` at 8 ms). Each tick calls `OnAnimationTick`, which calls `_root->Tick(...)`; returning `false` (no work left) auto-unsubscribes. `FrameClock::SmoothDeltaUs` clamps hitches to 50 ms and targets 120 Hz. During `FrameStage::Render`, layout mutation is blocked (`EmitDxUiRenderMutationBlockedForDebug`).

#### Invariants, extension points, gotchas

- Store `WindowHost` in stable storage; route the owner WndProc through `HandleMessage` and always `Detach()` on `WM_NCDESTROY` before destroying owner state.
- Grid/tree models are non-owning and must outlive the control; never read/mutate them off-thread.
- Extend by subclassing `Control`/`Panel` or implementing `IDxGridModel`/`IDxGridDelegate`; add a `ThemePalette` token in `DxUi.Theme.cpp` rather than hardcoding colors.
- Gotchas: no Win32 background brush (DxUi paints its own surface); hidden hosts must not present (`IsHostWindowEffectivelyVisible`); `TextInputBackend` is `Native` only — do not reintroduce hidden edit/RichEdit backers; keep TSF `ITfThreadMgr` active at UI-thread scope and shut it down through `ShutdownNativeTextInputForCurrentThread()`, not per modal prompt or focus cycle.

### Settings & SettingsStore: schema, hot reload, migration

_Versioned JSON5 settings persistence with atomic writes, file-watcher hot reload, an auto-generated JSON Schema (base + plugin configs), and replace-on-mismatch "migration"._

RedSalamander persists all per-user state as a single JSON5 file under `%LocalAppData%\RedSalamander\Settings\`. The on-disk shape is the strongly-typed `Common::Settings::Settings` struct (`Common/SettingsStore.h`); the store engine, an auto-generated JSON Schema, and a live file watcher keep the running app, the on-disk file, and the Preferences UI in sync.

#### Architecture and key files

| File / type | Role |
|---|---|
| `Common/SettingsStore.h` (`Common::Settings::Settings`) | Canonical typed model; `schemaVersion = 16` |
| `Common/Common/SettingsStore.cpp` | Load/save engine, path resolution, atomic write, recovery/backup |
| `RedSalamander/AppDataPaths.cpp` | Resolves `%LOCALAPPDATA%` (mirrors the store's own resolver) |
| `RedSalamander/SettingsHotReload.cpp` | Directory watcher thread, save+schema wrapper, stamp dedupe, conflict prompts, theme apply |
| `RedSalamander/SettingsSchemaExport.cpp` | Builds aggregated schema = base + per-plugin config `$defs` |
| `RedSalamander/SettingsSchemaParser.cpp` | Parses `x-ui-*` schema annotations to drive Preferences panes |
| `RedSalamander/SettingsSave.h` (`SettingsSave::PrepareForSave`) | Drops default-valued optional sections before writing |

#### Data and control flow

Load: `LoadSettings`/`LoadSettingsWithRecoveryInfo` call `ResolveSettingsLoadPath` (debug build prefers `<AppId>-debug.settings.json`, then versioned `<AppId>-<Major>.<Minor>.settings.json`, then legacy `<AppId>.settings.json`), then `LoadSettingsFromResolvedPath` parses JSON5 via yyjson and dispatches per-section `Parse*` helpers. Save goes through `SaveSettings` → `WriteFileBytesAtomic`, which writes `<file>.tmp`, `FlushFileBuffers`, then `MoveFileExW(MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)` — never a partial in-place write.

Hot reload: `SettingsHotReload::Start(hwnd, appId)` spawns a `std::jthread` running `WatchSettingsDirectoryThread`, which uses `FindFirstChangeNotificationW` on the settings directory (filters: file name, last write, size). On a change it posts `WndMsg::kSettingsFileChanged` to the main window. The UI-thread handler `OnMainWindowSettingsFileChanged` (in `RedSalamander.cpp`) calls `TryLoadChangedSettings`, which compares a `SettingsFileStamp` (volume serial + file index + last-write + size) against `lastAppliedStamp`/`lastRejectedStamp` to ignore self-writes, retries up to 6× on sharing/lock errors via `TryLoadSettingsNoRecovery`, then `MergeDiskSettingsWithRuntimeSession` overlays live window placement and pane folders before `ApplyCurrentSettingsToRunningApp` and `NotifyParticipants` (posts `WndMsg::kSettingsReloadedFromDisk` to registered dialogs).

#### Schema generation

`SaveSettingsAndSchema` writes both the settings file and `<AppId>.settings.schema.json`. `SaveAggregatedSettingsSchema` deep-copies the shipped base schema (`GetSettingsStoreSchemaJsonUtf8`, cached from `SettingsStore.schema.json` next to the exe), strips app-irrelevant sections (e.g. `monitor` for RedSalamander), and injects each plugin's `GetConfigurationSchema()` payload as a `$defs/pluginConfig_*` entry referenced under `pluginsSettings.configurationByPluginId`. The parser reads `x-ui-pane`, `x-ui-control`, `x-ui-section`, `x-ui-order` to build Preferences panes.

#### Migration

There is no field upgrader: `schemaVersion` must equal `16` exactly. Any other value yields `SettingsLoadRecoveryReason::UnsupportedSchemaVersion`; on the cold-start recovery path the bad file is renamed `*.bad.<UTC>` and defaults are restored (see `Specs/Core/Core_SettingsStore.md`). A v16 file still carrying legacy `viewers`/`editors` keys is rejected as `LegacyShape`. Legacy plugin IDs are normalized at parse time.

#### Threading, invariants, gotchas

The watcher runs off-thread but only ever **posts** messages; all mutation of `g_settings` happens on the UI thread. `g_state` is guarded by one `std::mutex`. Live reload of an invalid file is non-destructive (keeps current settings, shows a modeless alert via `ShowInvalidReloadAlert`); only cold start backs up. Always save through `SettingsHotReload::SaveSettingsAndSchema` so the applied stamp updates and the watcher does not echo your own write. See [SettingsFile.md](SettingsFile.md) and `Specs/SettingsStore.schema.json`.

### Find Files, Search Backends & the Search Service

_The modeless Find window, its worker-threaded session controller, the layered local search backends (service → local-index → scan), and the out-of-process SQLite-backed search service._

RedSalamander search is a host-owned feature: one Find experience that works for every `IFileSystem` plugin, prefers a fast indexed/service backend for local roots, and always has a correctness baseline (live scan). `Specs/Core/Core_Search.md` is the authoritative spec; the user-facing page is [FindFiles.md](FindFiles.md).

#### Architecture and main types

The `cmd/pane/find` command (defaults `Alt+F7` / `Ctrl+F`) calls `ShowFindFilesWindow(...)` (`RedSalamander/FindFilesWindow.h`) with a `FindFilesPaneContext` snapshotting the target pane's `IFileSystem`, `pluginId`, `instanceContext`, and root path. Each invocation creates an independent modeless `FindFilesWindow` (`RedSalamander/FindFilesWindow.cpp`), a DxUI-hosted window implementing `IDxGridDelegate`.

| File / type | Role |
|---|---|
| `RedSalamander/FindFilesWindow.cpp` — `FindFilesWindow` | UI, result grid, set algebra, status lines |
| `FindFilesWindow.cpp` — `SearchSessionController` | Owns the `std::jthread` worker, cancellation, completion |
| `FindFilesWindow.cpp` — `SearchCallbacks` | `IFileSystemSearchCallback`; batches results, host extensions |
| `RedSalamander/SearchFallbackEngine.cpp` — `Execute` | Generic scan baseline |
| `Plugins/FileSystem/FileSystem.Search.cpp` — `SelectSearchBackend` | Local backend routing |
| `Common/SearchServiceBroker.{h,cpp}` | Named-pipe client + `RunServer` |
| `Common/LocalSearchIndexCore.{h,cpp}` — `Repository` | Index core / query planning |
| `Common/SqliteIndexStore.{h,cpp}` | `index-v2.sqlite3` schema v2 |
| `RedSalamanderSearchService/Main.cpp` | Service host + FTXUI dashboard |

#### Data and control flow

`BeginSearch` builds a `SearchRequest` from the combos/checkboxes and calls `SearchSessionController::Start`, which spawns the worker (`Run`). `Run` fills a `FileSystemSearchQuery` and `QueryInterface`s the plugin for `IFileSystemSearch`. If present it calls `Search(...)`; if that returns an unsupported HRESULT, or if the plugin has no search interface, it falls back to `SearchFallbackEngine::Execute`, which traverses via `IFileSystem::ReadDirectoryInfo` (a stack-based `SearchDirectoryTree`) and scans content via `IFileSystemIO::CreateFileReader`. For the built-in local `file` plugin the host sets `query.reserved = FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1` and passes `FileSystemSearchHostExtensions` as the cookie, after seeding `SearchServiceBroker::GetStatus`. The plugin's `SelectSearchBackend` then routes by `searchBackendPreference` (`auto`/`service`/`local-index`/`scan`); `FILESYSTEM_SEARCH_FORCE_SCAN` (set when "Prefer indexed backend" is unchecked) always wins. The service path goes through the broker pipe (`\\.\pipe\RedSalamander.SearchService.v3`, Debug suffix `.Debug.v3`, protocol v3) to `RunServer`, which queries a `LocalSearchIndexCore::Repository` backed by SQLite or live scan.

#### Threading / UI-thread rules

`IFileSystemSearch::Search(...)` is synchronous, so it MUST run on the worker. Results never touch the UI thread directly: `SearchCallbacks` batch into `FindResultRecord`s and `PostMessagePayload` them as `WndMsg::kFindSearchResults`, `kFindSearchProgress`, and `kFindSearchComplete`, consumed in `WindowProc` by `OnSearchResults/Progress/Complete`. Flushing is by capacity (`kBatchSize`) or age so the first batch never waits on a low-priority timer. Cancellation is an `std::atomic<bool>` checked at traversal boundaries (`CheckSearchCancelled`).

#### Invariants and extension points

Result identity is `pluginId + normalized instanceContext + normalized fullPath` — used for de-dup and the Find/Append/Intersect/Subtract set operations. Regex is validated before traversal; invalid/unsafe patterns return `E_INVALIDARG` with one `REGEX_REJECTED` completion. The extension seam is `IFileSystemSearch` (`Specs/Plugins/Plugins_VirtualFileSystem.md`); without it, scan fallback still provides name search. The service is independent: `--register`, `--run-foreground` (FTXUI dashboard), `--compact`, `--request-compact`.

#### Gotchas

Debug and Release services use separate ProgramData roots and pipes so they never share SQLite state. SQLite bootstrap failures are fail-open — the service still answers via live scan and reports `DEGRADED_NO_INDEX` (distinct from `SERVICE_UNAVAILABLE`). `DirectoryInfoCache` (`Specs/Core/Core_DirectoryInfoCache.md`) caches enumerations for panes but is not the Find traversal path; Find scans through `ReadDirectoryInfo` directly.

### Compare Directories Engine & Window

_Dual-pane directory-comparison feature: a background diff engine with worker pools plus a FolderWindow-hosted window that drives compare mode, options, progress, and sync._

The Compare Directories feature opens a dedicated dual-pane window that compares two folder trees (Left vs Right), matches entries by name under the same relative path, and surfaces only differing items by default. It splits cleanly into a UI-agnostic engine (`CompareDirectoriesEngine.*`) and a window/host layer (`CompareDirectoriesWindow.*`).

#### Architecture and key files

| File / Type | Role |
|---|---|
| `CompareDirectoriesSession` (`CompareDirectoriesEngine.h/.cpp`) | Owns roots, settings, decision cache, scan + content-compare worker pools, versioning |
| `CompareDirectoriesFileSystem` (engine .cpp) | Per-pane `IFileSystem`/`IInformations` wrapper; cache-only enumeration |
| `CompareDirectoriesWindow` (`CompareDirectoriesWindow.Internal.h` + `.cpp`) | Top-level window, banner/menu, splitter, embeds a `FolderWindow` |
| `.Options.cpp` / `.Progress.cpp` / `.Menu.cpp` | Options panel (DxUi cards), progress UI + File Operations task card, themed menu bar |
| `CompareDirectoriesFolderDecision` / `...ItemDecision` | Per-folder/per-item diff results keyed by `WStringViewNoCaseLess` ordered map |

Entry point: `ShowCompareDirectoriesWindow(owner, settings, theme, shortcuts, left, right)`. Difference state is a bitmask of `CompareDirectoriesDiffBit` (`OnlyInLeft`, `Size`, `Content`, `ContentPending`, `SubdirPending`, ...).

#### Data and control flow

Enumeration is driven through the wrapper, not synchronously. `CompareDirectoriesFileSystem::ReadDirectoryInfo` delegates to the base filesystem when compare is disabled or the path is outside roots; otherwise it calls `session->TryGetCachedDecision(rel)` (cache-only, never I/O). On a miss it calls `RequestScanForFolder(rel)` and returns an empty list. Background `ScanWorker` threads run `ComputeDecisionForFolder` (enumerate both sides via `IFileSystem::ReadDirectoryInfo`, deliberately bypassing `DirectoryInfoCache`), match names ordinally, compare metadata, and enqueue `ContentCompareJob`s when `compareContent` is on. `ContentCompareWorker` threads open files via `IFileSystemIO::CreateFileReader` and stream-compare (256 KB buffer), then post results.

#### Threading and UI-thread rules

The session is fully thread-safe under one `_mutex` plus atomics; workers are `std::jthread` pools (content workers run `BELOW_NORMAL`, `CoInitializeEx` MTA). Progress reaches the UI only via posted payload messages in `Common/WindowMessages.h`: `kCompareDirectoriesScanProgress` (WM_APP+0x521), `kCompareDirectoriesContentProgress` (+0x524), `kCompareDirectoriesDecisionUpdated` (+0x523), `kCompareDirectoriesDeferredStart` (+0x520), `kCompareDirectoriesExecuteCommand` (+0x522). Receivers use `TakeMessagePayload<T>(lParam)`. Decision updates are coalesced: `kCompareDirectoriesDecisionUpdated` calls `ScheduleDecisionRefresh()`, and a `kCompareDecisionRefreshTimerId` tick (200 ms) drains updates in bounded batches via `FlushPendingContentCompareUpdatesBudgeted` / `FlushPendingSubdirUpdatesBudgeted`. The UI thread must never trigger subtree traversal or synchronous compare I/O.

#### Key invariants and contracts

An atomic `uint64_t _version` is the coherence mechanism: it increments on `SetRoots`, `SetSettings`, `Invalidate`. Cached decisions are tagged with the version and treated stale on mismatch; in-flight jobs carry the version and bail when it changes. A separate `_uiVersion` suppresses redundant refreshes. Progress messages are tagged with a run id (`_compareRunId` / version at scan start) and stale updates are ignored. Start is two-phase (`PrepareCompareRun` then `ExecutePreparedCompareRun` via the deferred message) so a Rescan replaces, not accumulates, run state. Cancel calls `SetBackgroundWorkEnabled(false)` for responsiveness.

#### Extension points and gotchas

New diff criteria go through `CompareDirectoriesDiffBit` + `ComputeDecisionForFolder` + `BuildDetailsTextForCompareItem`. Content compare requires `IFileSystemIO` on both sides (`IsContentCompareSupported`); otherwise `compareContent` must be forced off. Gotchas: name matching ignores trailing spaces/dots and uses an ordered `std::map` (never `unordered_map`) to avoid hash/ordinal-equality mismatches; enumeration failures are marked different and NOT cached (retried); when `keepIdenticalItems` is off, per-file `ContentPending` placeholders may be elided but folder-level `anyPending` must stay accurate. After file operations the session calls `InvalidateForAbsolutePath(path, includeSubtree=true)`. See `Specs/Core/Core_CompareDirectories.md` and user docs [CompareDirectories.md](CompareDirectories.md); related sync behavior in [FileOperations.md](FileOperations.md).

### Batch Rename, Change Case & Rename Batching

_Preview-first bulk rename engine plus the Change Case command and the shared arena-backed batch-rename marshaller that drives every multi-item leaf rename._

This subsystem turns a set of files/folders into a validated rename plan, previews it, and applies it through the plugin file-system layer. It is split into a pure engine, a modeless UI window, helper-menu logic, a standalone Change Case command, and the low-level batch marshaller they all share.

#### Key files and types

| File | Role |
|------|------|
| `RedSalamander/BatchRenameEngine.{h,cpp}` | Pure `BatchRename::BuildPlan(targets, rules)` -> `Plan` (macros, search/replace, case, validation). No I/O, `noexcept`. |
| `RedSalamander/BatchRenameWindow.{h,cpp}` | Modeless `BatchRenameWindow` (DxUi): scope/collection, preview grid, threading envelope, settings, undo, debug API. |
| `RedSalamander/BatchRenameExecutionEngine.{h,cpp}` | Window-free Batch Rename provider-dispatch engine: dependency layers, cycle temp hops, per-item outcome bookkeeping, undo entries, directory-move path rewriting, and execution metrics. |
| `RedSalamander/BatchRenameMenus.{h,cpp}` | Template/regex/replacement helper flyouts and caret-aware insertion. |
| `RedSalamander/MaskSyntax.{h,cpp}` | `*.*`-style include/exclude wildcard masks + MRU history. |
| `RedSalamander/ChangeCase.{h,cpp}` | `cmd/pane/changeCase` casing logic + `ApplyToPaths` recursive driver. |
| `RedSalamander/FileSystemRenameBatch.{h,cpp}` | Arena-backed `Execute()` that calls `IFileSystem::RenameItems`, falling back to per-item `RenameItem`. |

#### Control and data flow

Launch goes through `FolderWindow::CommandBatchRename` (`RedSalamander/FolderWindow.cpp`), which fills a `BatchRenamePaneContext` (file system, plugin id/instance, root, initial paths, `onSuccessfulRename`, `onRevealPath`) and calls `ShowBatchRenameWindow`. The window collects targets (folder-scope collection runs on a background `std::jthread`; explicit selection may stay on the UI thread), then on every edit calls `RequestPreviewRebuild` -> debounced `OnPreviewRebuildTimer` (~150 ms) -> `RebuildPreview` -> `BuildPlan` + `ApplyContextualPreviewValidation`.

`BuildPlan` applies transforms in a fixed order: macro expansion (`ExpandTemplate`/`ResolveMacro`, canonical `{macro}` plus `$(Macro)` aliases via `NormalizeAliasTokens`), then search/replace (`ApplyReplacement`: literal `ReplaceLiteral` or one compiled `std::wregex`), then `ApplyCaseTransforms` (which delegates to `ChangeCase::TransformLeafName`), then `ValidateLeafName`/`AddWarningIssues`/`MarkDuplicateTargets`. Issues carry stable IDs (`name_empty`, `name_duplicate`, `name_reserved_device`, `macro_unknown`, `macro_invalid_format`, `regex_invalid`, `regex_match_failed`, ...) localized by `LocalizeBatchRenameIssueId`. The hot path uses allocation-light leaf splitting for `{stem}`/`{ext}` and helper-owned `FileSystemPathIdentity` keys for duplicate detection when the profile can safely provide them; every key hit is still verified with identity equality.

Execution: `BatchRenameWindow::ExecuteRename` builds non-no-op `BatchRenameExecutionOp`s, sorts deepest-first, and hands them to a worker wrapper (MTA COM). The wrapper keeps the provider owner alive until before `CoUninitialize`, calls `RunBatchRenameExecutionEngine`, and posts `WndMsg::kBatchRenameCompleted` back to the UI thread. `BatchRenameExecutionEngine` processes depth groups, builds dependency layers, breaks pure cycles with a `.rsren-<guid>` temp hop, and dispatches each layer through `FileSystemRenameBatch::Execute`. Per-item outcomes come from the engine callback's `FileSystemItemCompleted`; `OnExecutionCompleted` refreshes targets, fires `DirectoryInfoCache::NotifyPathMoved`, stores the report/undo plan, and invokes `onSuccessfulRename` -> `FolderWindow::RefreshPanesAfterBatchRename`.

#### Threading and UI-thread rules

`BuildPlan`, validation, local revalidation, all grid/model updates, and settings persistence are UI-thread only. Only target collection and the provider-dispatch phase run on the single `_taskWorker` jthread. Progress arrives via `WndMsg::kBatchRenameTaskUpdate` (throttled ~100 ms). Reserved message IDs live in `Common/WindowMessages.h`. Re-entrancy guard: `_executing`/`_collecting` block `Rename`; the footer `Cancel` sets `_cancelRequested`. Close path calls `CancelAndJoinBackgroundTask` then drains posted payloads in `WM_NCDESTROY`. Stale results are discarded via `_taskGeneration`.

#### Invariants, extension points, gotchas

- `RenameItems`/`RenameItem` take leaf names only; macro/manual output with path separators is a blocking error.
- New macros: extend `ResolveMacro` and the helper specs in `BatchRenameMenus.cpp` together. New issue IDs need a `LocalizeBatchRenameIssueId` mapping + RC string.
- Change Case is invoked two ways: standalone `FolderWindow::CommandChangeCase` -> `ChangeCase::ApplyToPaths` (recursive, deepest-first, 64-item batches, `std::stop_token`), and inline in the engine via `TransformLeafName`.
- **Gotcha:** `Specs/UI/UI_BatchRenameWindow.md` mandates a single resolved `FileSystemPathIdentity` profile at every identity-sensitive "same path?" site. Batch Rename now routes duplicate/case-only validation, dependency layering, provider callback pairing fallback, destination probing, cache notification, target refresh, provider-selection matching, and undo path finalization through that profile. Same-context Copy/Move/Delete also rejects missing, malformed, unsupported, or unstable `pathIdentity` before task creation.
- **Gotcha:** folder-scope collection parses the mask once per collection run, and provider directory walks poll cancellation only after validating each entry's buffer bounds. Do not move cancellation checks ahead of malformed-buffer validation.
- **Perf contract:** keep `batchrename.preview.build_plan_us`, `batchrename.validation.us`, `batchrename.collect.us`, and `batchrename.execute.target_refresh_match.us` healthy; the 10,000-row preview and target-refresh selftests are the canaries for accidental O(n^2) work.
- **Gotcha:** undo entries record only the NET original->final transition; temp hops are hidden, and an orphaned temp is restored best-effort.

Cross-links: `Specs/UI/UI_BatchRenameWindow.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, and user-facing [UserGuide.md](UserGuide.md).

### Connection Manager, Secrets & Windows Hello

_Host-managed connection profiles with non-secret data in Settings, secrets in Windows Credential Manager, and Windows Hello gating before secrets are released to plugins._

This subsystem lets all virtual-file-system plugins (FTP/SFTP/SCP/IMAP, S3/S3 Table, Google Drive, OneDrive/SharePoint) share one connection list. Non-secret profile fields live in the Settings Store (`connections` section), secrets live in Windows Credential Manager (WinCred), and Windows Hello can gate secret release. Plugins never see secrets in URIs; they resolve a profile via the host-reserved `/@conn:<name>/...` path prefix and pull secrets through `IHostConnections`. Authoritative behavior is in `Specs/Core/Core_ConnectionManager.md`.

#### Key files & types

| File | Role |
|------|------|
| `RedSalamander/ConnectionManagerWindow.{h,cpp}` | Single-canvas DxUi window (`WindowImpl`); `ShowConnectionManagerWindow` (modeless) + `ShowConnectionManagerDialog` (synchronous facade) |
| `RedSalamander/ConnectionSecrets.{h,cpp}` | WinCred generic-credential I/O, Quick Connect in-memory store, secret-access authorization cache |
| `RedSalamander/WindowsHello.{h,cpp}` | `RedSalamander::Security::VerifyWindowsHelloForWindow` (WinRT `UserConsentVerifier`) |
| `RedSalamander/ConnectionCredentialPromptDialog.{h,cpp}` | Themed secret / user+password prompts |
| `RedSalamander/ConnectionProfileUtils.{h,cpp}` | `/@conn:` parsing, name lookup, display-URL build, `extra` JSON getters |
| `RedSalamander/HostServices.cpp` | `IHostConnections` implementation marshaled onto the UI thread |

#### Data & control flow

The WinCred target name is built by `BuildCredentialTargetName(connectionId, SecretKind)` as `RedSalamander/Connections/<id>/<password|sshKeyPassphrase|refreshToken>`. Secrets are UTF-16 NUL-terminated blobs written `CRED_PERSIST_LOCAL_MACHINE` via `SaveGenericCredential`/`LoadGenericCredential`/`DeleteGenericCredential`.

Editing flow: while the user types a secret, `WindowImpl::StageSecretForProfile` holds it in `_stagedPasswordById` / `_stagedPassphraseById` (never written until Connect/Close). On save, `CommitSecretsForProfile` writes or deletes WinCred entries based on `savePassword` and `authMode` (OAuth2 PKCE never stores a password; `requireWindowsHello` is committed too). Quick Connect (`@quick`, id `...0001`) is intercepted by `CommitQuickConnectSecretsAndProfile` and kept purely in memory. Connect posts `WndMsg::kConnectionManagerConnect` (defined in `Common/WindowMessages.h`) with the copied connection name to the owner pane.

Plugin retrieval: `IHostConnections::GetConnectionSecret` (and `SetConnectionSecret`, `PromptForConnectionSecret`, `UpgradeFtpAnonymousToPassword`) may be called from any thread; `HostServices` marshals to the UI thread via `SendMessageW(WndMsg::kHostGetConnectionSecret, ...)` etc. `GetConnectionSecretOnUiThread` checks the in-memory session cache first, then (for persisted profiles) gates on Hello, then loads from WinCred and returns a `CoTaskMemAlloc` buffer.

#### Windows Hello gating

Both the reveal path (`VerifySecretRevealForProfile` in the window) and the plugin path (`GetConnectionSecretOnUiThread`) apply the same policy: skip if `!requireWindowsHello` or `bypassWindowsHello`; otherwise reuse a recent authorization. `IsSecretAccessAuthorized(id, reauthTimeoutMs)` checks `windowsHelloReauthTimeoutMinute` (default 10; `0` = always prompt). Crucially, `HasSecretAccessAuthorization(id)` is also accepted so a once-authorized connection is not re-prompted during long background copy/compare. `VerifyWindowsHelloForWindow` runs the WinRT async op under a message-pumping wait (`WaitForOperationWithMessagePump`) and maps results to `S_OK` / `ERROR_CANCELLED` / `ERROR_NOT_SUPPORTED`.

#### Threading, invariants, gotchas

- Window code is UI-thread only; `IHostConnections` enforces this by marshaling (returns `ERROR_INVALID_THREAD_ID` if called on the wrong thread directly).
- The window class must register `CS_DBLCLKS` for DxUi text fields.
- Authorization ticks use `GetTickCount64`; `ENABLE_TESTS` exposes `SetSecretAccessAuthorizationTickForTesting` and `SetWindowsHelloTestVerifier` to avoid real prompts/sleeps.
- Removing a profile from the list deletes its three WinCred entries at save time (baseline-id diff). Never log secrets.

See [Connections.md](Connections.md) and `Specs/UI/UI_TopLevelToolWindows.md` for the window contract.

### Theming (version 2 authored themes, AppTheme, and Preferences)

_How the shared version 2 theme model resolves palettes, references, static/event/paint functions, and Rainbow inheritance into the `AppTheme` consumed by GDI/D2D/DWM surfaces._

Theme input has one shared representation and parser. `Common::Settings::ThemeDefinition` requires `formatVersion == 2` and stores authored `ThemeColorSource` values in both `palette` and semantic `colors` maps. `ThemeDefinitionIo` reads/writes that representation losslessly; there is no version-1 reader or flattened compatibility writer. `ThemeExpression` parses, validates, formats, resolves dependencies, reports cycles/missing references/limits, and compiles allowlisted paint-time sources. Runtime adapters then apply `ResolvedThemeColors` to value-typed `AppTheme` and Monitor theme structures.

#### What this part does

`AppTheme` (`RedSalamander/AppTheme.h`) aggregates sub-themes for each surface: `FolderViewTheme`, `NavigationViewTheme`, `MenuTheme`, `TitleBarTheme`, `FileOperationsTheme`, `ViewerDiffTheme`, plus scalar fields (`dark`, `highContrast`, `accent`, `compactMode`, `reducedMotionOverride`, backdrop types, and `windowBackground`). `ResolveAppTheme(requestedMode, rainbowSeed, accentOverride)` constructs the built-in base. `MakeAppThemeResolutionContext` exposes inherited semantic values and system colors to `ResolveThemeDefinition`; the resolved static map and immutable compiled programs are then applied atomically. System High Contrast wins before authored or Rainbow colors.

Five built-in `ThemeMode` values exist: `System`, `Light`, `Dark`, `Rainbow`, and `HighContrast`. Rainbow follows the effective Windows light/dark base, derives repeatable colors from stable 32-bit hashes, and retains application-wide plugin/viewer Rainbow flags only when the base is `builtin/rainbow`. A static authored value suppresses inherited Rainbow for that one token. Other bases may use an allowlisted `seededRainbow`/`seededChoice` token without enabling unrelated Rainbow surfaces.

#### Architecture and key files

| File | Role |
|------|------|
| `Common/ThemeExpression.h`, `Common/Common/ThemeExpression.cpp` | Authored source model, parser/formatter, graph resolver, dependency maps, system/perceptual functions, compiled dynamic evaluator |
| `Common/ThemeDefinitionIo.h`, `Common/Common/ThemeDefinitionIo.cpp` | Strict standalone and lenient inline version 2 JSON5 I/O |
| `Common/SettingsStore.h`, `Common/Common/SettingsStore.cpp` | Inline user-theme persistence and opaque recovery for unusable entries |
| `AppTheme.h/.cpp` | Built-in base palette, shared resolution context, resolved override application, dynamic programs, DWM helpers |
| `DxUiThemePalette.h` | `MakeAppThemeDxPalette` — adapts `AppTheme` to `DxUi::ThemePalette` (see [DxUi.md](DxUi.md)) |
| `ThemedInputFrames.h/.cpp` | Subclasses native edit/combo controls + their frame to paint themed borders/backgrounds via D2D-on-HDC |
| `Preferences.h/.cpp` | Public `ShowPreferencesDialog*` entry points + `UpdatePreferencesWindowsTheme` |
| `Preferences.Dialog.cpp` | Modeless shell: category tree, scrollable page host, OK/Cancel/Apply, per-page dispatch |
| `Preferences.Internal.h` | `PreferencesDialogState`, `PrefCategory`, layout constants, per-namespace settings accessors |
| `Preferences.Themes.cpp` | Lossless theme list/duplicate/reset/import/export and temporary application |
| `RedConfigure/Themes/ThemePreviewModel.*` | Advanced palette/source editing, dependency inspection, fixed-seed preview resolution |

`ColorFromCOLORREF`/`ColorToCOLORREF` bridge GDI `COLORREF` and D2D `D2D1::ColorF`. `MakeAppThemeDxPalette` mixes accent into hovered/pressed states and selects the overlay material (Mica/MicaAlt/Acrylic/Solid) from the backdrop type.

#### Preferences dialog flow

`ShowPreferencesDialog{,Plugins,HotPaths,UserMenu}` call `PreferencesDialog::Show(...)` with an initial `PrefCategory`. The dialog is **modeless and single-instance** (`GetHandle`). `PreferencesDialogHost` extends `PreferencesDialogState` and owns one pane object per category (`_generalPane`, `_themesPane`, …) plus three `DxUi::WindowHost`s (`_categoryTreeHost`, `_shellHost`, `_pageHostHost`). Pages are created lazily: `EnsurePreferencesPageInitialized` switches on `currentCategory`, calls `pane.InitializePage(...)` once (guarded by `paneFirstCreateDone[]`), and toggles `paneWrapperPanels[]` visibility. `SelectCategory` (line ~4132) and `RefreshPreferencesPage` (the `case` dispatch around line 3847) route to each pane's `Refresh`/`LayoutPage`.

Settings use a two-copy model in `PreferencesDialogState`: `baselineSettings` vs `workingSettings` (monitor settings tracked separately). `SetDirty` enables Apply; external on-disk changes set `staleFromExternalReload` and trigger `ReloadPreferencesDialogFromDisk`.

#### Theme preview and commit

The Themes page edits `workingSettings.theme`, preserving `palette` and `colors` source objects through duplicate, reset, import, and export. `ApplyThemeTemporarily` copies the working selection into live settings, resolves the complete graph before replacement, repaints, and posts `WndMsg::kSettingsApplied` to the owner. `RestorePreviewAppliedPreferencesOnCancel` restores the baseline selection. Advanced graph authoring remains RedConfigure-owned so Preferences does not maintain a second expression designer.

Static and event-time sources resolve on theme application or the relevant system color/theme notification. Paint-time `seededRainbow` and `seededChoice` programs are parsed and compiled during resolution; `EvaluateDynamicThemeColor` receives only the stable runtime hash and performs no parsing, allocation, locking, system query, callback, or I/O. Theme/settings hot reload builds a complete replacement first and retains the last valid live theme on failure.

#### Threading, invariants, gotchas

Theme application and event-time system-color resolution are **UI-thread only**. Theme resources held in state (`backgroundBrush`, `cardBrush`, `inputBrush`…) are `wil::unique_hbrush` (RAII). Never retain a stale `AppTheme&` or a pointer into a prior `ResolvedThemeColors` across selection/hot reload. Paint consumers may read only the immutable compiled program owned by the atomically installed theme. Semantic keys must be real runtime keys recognized by all intended adapters; preview-only aliases are forbidden. See `Specs/Core/Core_SettingsStore.md`, `Specs/Core/Core_RedConfigure.md`, `Specs/UI/UI_VisualStyle.md`, and `Specs/UI/UI_PreferencesDialog.md`.

### Diagnostics: ETW/TraceLogging, Debug Logging & Monitor

_How RedSalamander emits structured diagnostics over ETW/TraceLogging via Debug:: helpers, and how RedSalamanderMonitor consumes and renders them in real time._

RedSalamander has no log files: every diagnostic is a structured ETW (Event Tracing for Windows) event emitted through a single TraceLogging provider, and `RedSalamanderMonitor.exe` is the real-time consumer. This decouples producers from consumers - apps emit regardless of whether a monitor is running.

#### Architecture: emit -> transport -> consume

All of the producer side lives in `Common/Helpers.h` inside `namespace Debug`. The provider is declared once with `TRACELOGGING_DECLARE_PROVIDER(g_RedSalamanderProvider)` and defined in exactly one `.cpp` per module via `#define REDSAL_DEFINE_TRACE_PROVIDER`. Critically, the GUID `{440c70f6-6c6b-4ff7-9a3f-0b7db411b31a}` is shared by every module but each EXE/DLL owns its own provider storage (TraceLogging cannot share handles across DLL boundaries). The same GUID is hard-coded in `EtwListener::kProviderGuid` so the consumer subscribes to all producers at once.

| File / Type | Role |
|---|---|
| `Common/Helpers.h` `Debug::Error/Warning/Info/ErrorWithLastError` | Public logging API; all funnel into `Debug::Out` |
| `Common/Helpers.h` `Debug::Perf::Scope`/`Emit*` | Perf events (`PerfScope`) + JSONL sink |
| `Common/Helpers.h` `EmitEtwEvent`, `InfoParam` | Writes `DebugMessage` event; metadata struct |
| `Common/Helpers.h` `CallTracer` / `TRACER*` macros | Per-thread indentation + timing |
| `RedSalamanderMonitor/EtwListener.{h,cpp}` | Real-time ETW session + TDH decode |
| `RedSalamanderMonitor/ColorTextView.cpp` | Cross-thread queue + batch intake + render |
| `RedSalamanderMonitor/MonitorDiagnostics.h` | Self-event filtering policy |

#### Data and control flow

`Debug::Error(...)` and friends build an `InfoParam` (FILETIME, processID, threadID, `Type`) via `BuildInfoParam`, format the message, and call `EmitEtwEvent`, which calls `TraceLoggingWrite` with event name `"DebugMessage"` under `kDebugKeyword` (0x1). `Debug::Perf` emits `"PerfScope"` events under `kPerfKeyword` (0x2). Write/fail counters (`g_etwWritten`/`g_etwFailed`) are exposed via `GetTransportStats()`.

On the consumer side `EtwListener::Start` stops any stale `RedSalamanderMonitor_ETW_Session`, calls `StartTrace` (256 KB buffers, 8-128 buffers), `EnableTraceEx2` on the provider, `OpenTrace`, then runs `ProcessTrace` on a `std::jthread`. `EventRecordCallback` -> `HandleEvent` -> `ExtractEventData` decodes properties with the TDH API (slow but off the UI thread) and synthesizes `[perf] ...` text for Perf events. The user callback (wired in `OnCreate` near line 3687) applies `ShouldAcceptEtwEventForDisplay`, then calls `ColorTextView::QueueEtwEvent`.

#### Threading and UI-thread rules

`QueueEtwEvent` runs on the ETW worker thread: it pushes onto `_etwEventQueue` (a `std::deque` under `wil::critical_section _etwQueueCS`) and posts `WndMsg::kColorTextViewEtwBatch` only when the queue was empty (coalescing). `OnAppEtwBatch` runs on the UI thread, drains up to `kMaxBatchSize`=200 entries, reposts if overflow remains, and calls `Document::AppendInfoLines` under a single write lock. Append/layout/paint must stay on the UI thread; everything else (the trace session) stays on the worker. See `Specs/Core/Core_RedSalamanderMonitor.md` for the two-mode (AUTO_SCROLL/SCROLL_BACK) rendering contract.

#### Invariants and gotchas

- Self-event suppression: by default Monitor filters out events whose `processID == GetCurrentProcessId()` and shows no startup text - only `--etw` (which sets `SetRuntimeMonitorDiagnosticsEnabled`) opts into self-diagnostics.
- Build gating: Debug/ASan Debug builds emit Info/Perf/Debug by default; Release emits only Error/Warning unless `--etw`. See `ShouldEmitMonitorDiagnosticMessageType`.
- ETW session start can fail with `ERROR_ACCESS_DENIED` - run `init-etw-trace.ps1` (Performance Log Users) or elevate; the UI offers a relaunch prompt.
- `TRACER` logs only on exit; use `TRACER_INOUT` for entry+exit. Indentation is thread-local.
- Perf has a second sink: JSONL via `--perf[=PATH]` (`WritePerfJsonl`), used by self-test perf gates. See [Monitor.md](Monitor.md) and [Troubleshooting.md](Troubleshooting.md).

### Localization, Resources, Build & Test Infrastructure

_How RedSalamander stores localizable UI in .rc resources with satellite DLLs, builds via build.ps1/MSBuild, and validates via in-product self-tests and the Run-AllTests harness._

This subsystem covers three intertwined concerns a new contributor touches constantly: where user-facing text lives, how the solution builds, and how changes are validated.

#### Localization and resources

Every user-facing string and every static menu/dialog/accelerator lives in `.rc` resources, never hardcoded in C++ (`Specs/Core/Core_Localization.md`). English resources are embedded in each owner module (`RedSalamander/RedSalamander.rc`, `RedSalamanderMonitor/RedSalamanderMonitor.rc`, per-plugin `.rc`). The main `.rc` defines `LANGUAGE LANG_ENGLISH, SUBLANG_ENGLISH_US`, menus like `IDR_FOLDERVIEW_CONTEXT` and `IDR_COMPARE_DIRECTORIES_MENU`, dialogs (`IDD_PREFERENCES`, `IDD_COMPARE_DIRECTORIES_OPTIONS`), an `IDC_REDSALAMANDER ACCELERATORS` table, and many `STRINGTABLE` blocks.

Translations are resource-only satellite DLLs under `RedSalamander/Lang/<culture>/`, named `<OwnerStem>-<culture>.dll` (e.g. `RedSalamander-fr-FR.dll`). Each culture folder holds a generated `.rc` plus a `.vcxproj` that outputs to `.build\<Platform>\<Configuration>\Lang\`. Runtime lookup goes through the `Localization` namespace in `Common/LocalizationManager.h`: owners call `RegisterResourceOwner(ownerName, hInstance)` at startup (plugins pass `g_hInstance`, unregister on unload via `UnregisterResourceOwner`); `LoadString`, `LoadMenuResource`, `LoadAcceleratorsResource`, and `FindLocalizedResourceHandle` try the selected culture chain in satellites first, then fall back to embedded English. App code uses the `Common/Helpers.h` wrappers `LoadStringResource`, `FormatStringResource`, `MessageBoxResource`. Stable non-translatable tokens use `LoadEmbeddedStringResource`/`FormatEmbeddedStringResource` (owner-scoped, satellite-bypassing on purpose).

**Contracts/invariants:** formatted strings use positional `std::format` placeholders (`{0}`, `{1:08X}`) — bare `{}`, `{:08X}`, and printf-style `%s`/`%d` are forbidden so translators can reorder safely. Command labels need both full (`IDS_CMD_*`) and short (`IDS_CMD_SHORT_BASE + IDS_CMD_*`, ids 20000-21999) forms. `FindLocalizedResourceHandle` results are transient — copy immediately, never cache across a language change. `Tools/Tests/ResourceLocalizationContracts.Tests.ps1` enforces placeholder parity and the language-neutral inventory; run it after editing braced strings. See `.github/skills/localization/SKILL.md`.

#### Build

`build.ps1` (repo root) is the wrapper over MSBuild on `RedSalamander.sln`. It locates MSBuild via `vswhere`/path search (prefers VS 2026 / toolset v145), stamps versions through `Tools\Versioning.ps1`, and supports `-Configuration {Debug|Release|ASan Debug}`, `-Platform {x64|ARM64}`, `-ProjectName`, `-Clean`/`-Rebuild`, and packaging (`-Msix`, `-Msi`, `-Zip`, `-GenerateWingetManifest`). Helper modules under `Tools/` (`BuildProjectSelection.ps1`, `MSBuildInvocation.ps1`, `ProcessStreaming.ps1`) drive project selection and streamed logging to `.build\logs\`. Language satellite projects are validated to land in the `Lang\` output folder.

#### Tests

In-product self-tests are debug-only (`ENABLE_TESTS`) and split by suite under `RedSalamander/SelfTest/` (`Commands/`, `CompareDirectories/`, `FileOperations/`, shared `Common/SelfTestCommon.h`). They run via CLI flags parsed in `RedSalamander.cpp`: `--commands-selftest`, `--compare-selftest`, `--fileops-selftest`, with `--selftest-case=`, `--selftest-fail-fast`, `--selftest-list-cases`, `--selftest-timeout-multiplier=`. The shared `SelfTest::RunCase` template records each declared case as `passed`/`failed`/`skipped` with a reason into `results.json` (artifacts under `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\`, archived to `Specs/TestRuns/`). `Tools/Run-AllTests.ps1 -Suite Full` builds Debug, runs the suites plus standalone native/CppUnitTest/Pester tests, cross-checks each suite against `--selftest-list-cases` for coverage drift, and writes `run-all-tests-results.json`. See `Specs/Testing/Testing_SelfTests.md`, `Specs/Testing/Testing_TestCoverage.md`, `Specs/Testing/Testing_PerformanceValidation.md`.

#### Threading / UI-thread rules

GUI self-tests must run foreground (focus/pointer routing); launch the GUI-subsystem exe with `Start-Process -Wait -PassThru` to read its exit code. Runtime language changes must run on the UI thread: reload menu handles from `LoadMenuResource`, rebuild dynamic theme/plugin sections, and resync the DxUi menu model.

#### Key files/types

| File / symbol | Role |
| --- | --- |
| `RedSalamander/RedSalamander.rc` | Embedded English strings, menus, dialogs, accelerators |
| `RedSalamander/Lang/<culture>/*.rc + *.vcxproj` | Satellite translation DLLs |
| `Common/LocalizationManager.h` (`Localization::`) | Owner registration + localized resource lookup |
| `Common/Helpers.h` | `LoadStringResource`/`FormatStringResource`/`MessageBoxResource` |
| `build.ps1` + `Tools/*.ps1` | MSBuild wrapper, versioning, packaging |
| `RedSalamander/SelfTest/Common/SelfTestCommon.h` | `RunCase`, result/artifact contract |
| `Tools/Run-AllTests.ps1` | Unified runner + coverage cross-check |

#### Gotchas

Satellite resources must never duplicate language-neutral IDs (it breaks the parity test). Self-tests need a Debug build; a case must stay declared and `skipped` (with reason) when a precondition is missing rather than vanish. `--selftest-timeout-multiplier` is clamped to `[0.1, 100.0]`. Cross-link: [Themes.md](Themes.md), [Monitor.md](Monitor.md), and [dev/Localization.md](dev/Localization.md) for the full localization reference.

## Technical guides

- [DxUi Technical Guide](DxUi.md) explains the shared DirectX UI layer, host
  lifecycle, theme/background model, controls, examples, and tests.
- [Winget Integration](WingetIntegration.md) explains package manifest
  generation and validation.
- [Documentation Coverage Map](DocumentationMap.md) tracks what the public docs
  cover and which areas are still thin.
- [Monitor.md](Monitor.md) introduces RedSalamanderMonitor for ETW/log
  inspection.

## Specs worth reading first

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/UI/UI_VisualStyle.md`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/Testing/Testing_TestCoverage.md`

Use specs for normative behavior and `docs/` for onboarding, workflows, and
public explanations.


