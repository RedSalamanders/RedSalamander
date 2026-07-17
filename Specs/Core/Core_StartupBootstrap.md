# Startup & Bootstrap Specification (RedSalamander.exe)

## Goals

- Improve startup visibility with **ETW-backed metrics** (no-op when ETW is not active).
- Keep **RedSalamanderMonitor** as an **ETW-only** viewer (no shared-memory crash plumbing).
- Provide a **delayed splash screen** for slow startups that does not add startup tax for fast runs.
- Provide an **in-process crash front door** with **minidumps** and a **crash-on-next-launch** UX.

## Process and DLL Search Hardening

Both `RedLauncher.exe` and `RedSalamander.exe` establish
`LOAD_LIBRARY_SEARCH_SYSTEM32 | LOAD_LIBRARY_SEARCH_APPLICATION_DIR` through
`SetDefaultDllDirectories` before optional or runtime DLL loading. Startup fails and records an error if the
safe search policy cannot be installed; the processes do not continue with the inherited ambient DLL search
order and do not use process-global `SetDllDirectoryW`.

`RedLauncher.exe` resolves `RedSalamander.exe` relative to the launcher's final package path and passes that
trusted package directory as the child's current directory. User-supplied or inherited working directories
therefore cannot affect the initial child working directory or bare-name DLL discovery.

## ETW Startup Metrics

RedSalamander records one-time startup milestone events via `Debug::Perf::Emit()`. Debug and ASan Debug builds emit Info/Perf/Debug ETW diagnostics and write JSONL perf capture to the default path by default. Release builds keep ETW traffic quiet unless launched with `--etw`; Release JSONL perf capture is explicitly enabled with `--perf` for the default path or `--perf=PATH` for a custom path (or by selftest/perf harness configuration). `RedSalamanderMonitor` self diagnostics use the same runtime `--etw`, `--perf`, and `--perf=PATH` switches.

- `App.Startup.Metric.TimeToFirstWindow`  
  First successful main window creation (`CreateWindowW` returns an `HWND`).
- `App.Startup.Metric.TimeToFirstPaint`  
  First `WM_PAINT` observed on the main window.
- `App.Startup.Metric.TimeToInputReady`  
  Posted after `ShowWindow` + `UpdateWindow` complete and the window is ready to accept user input.
- `App.Startup.Metric.TimeToFirstPanePopulated`  
  First time a pane applies an enumeration result (items are available).

Implementation:

- Code: `RedSalamander/StartupMetrics.h`, `RedSalamander/StartupMetrics.cpp`
- Trigger points:
  - Created: `StartupMetrics::MarkFirstWindowCreated()`
  - Painted: `StartupMetrics::MarkFirstPaint()`
  - Input-ready: `WndMsg::kAppStartupInputReady` → `StartupMetrics::MarkInputReady()`
  - Pane populated: `FolderView` enumeration apply path → `StartupMetrics::MarkFirstPanePopulated()`

## Splash Screen (Delayed)

### Behavior

- Splash is shown **only if startup lasts longer than 300ms**.
- Splash is **opt-in via settings**:
  - `startup.showSplash` (default: `true`)
- The splash closes automatically when the main window is marked input-ready.

### UX / Rendering

- Borderless, topmost window, centered over:
  - the main window once available, otherwise
  - the active monitor work area.
- Uses embedded `res/logo.png` (resource `IDR_SPLASH_LOGO_PNG`) and paints a branded background.
- Optional status text can be updated from the app while it is visible.

Implementation:

- Code: `RedSalamander/SplashScreen.h`, `RedSalamander/SplashScreen.cpp`
- Resource: `RedSalamander/RedSalamander.rc` (`IDD_SPLASH`, `IDR_SPLASH_LOGO_PNG`)
- Preferences: `Preferences > General > Splash screen`

## Crash Handling (In-Process)

### Unified “Front Door”

RedSalamander installs a best-effort crash handler that:

- Writes a minidump using `MiniDumpWriteDump`.
- Writes a marker file so the next launch can present a crash UX.

Installation sources:

- `SetUnhandledExceptionFilter` (SEH)
- `std::set_terminate`
- CRT purecall + invalid parameter handlers

### Dump Location

- Folder: `%LOCALAPPDATA%\\RedSalamander\\Crashes`
- Marker: `%LOCALAPPDATA%\\RedSalamander\\Crashes\\last_crash.txt` (UTF-16, contains dump path)

### Crash-on-Next-Launch UX

On a subsequent run, when the main window is ready, RedSalamander:

- Detects the marker file.
- Prompts the user and offers to open the crash folder in Explorer.
- Removes the marker before prompting to avoid repeated prompts.

### Deliberate Crash Test

- Command line flag: `--crash-test`
- Forces a non-continuable exception to validate dump + marker + next-launch UI.

Implementation:

- Code: `RedSalamander/CrashHandler.h`, `RedSalamander/CrashHandler.cpp`
- Strings: `IDS_CRASH_DETECTED_*` in `RedSalamander/RedSalamander.rc`

## Debug-only: Auto-launch RedSalamanderMonitor

To capture startup ETW events during development, RedSalamander debug builds:

- Attempt to launch `RedSalamanderMonitor.exe` asynchronously as early as possible.
- Avoid launching if the monitor instance mutex exists:
  - `Local\\RedSalamanderMonitor_Instance`
- Release startup ETW capture is opt-in: launch `RedSalamander.exe --etw`. Debug and ASan Debug enable it by default.
- Release startup JSONL perf capture is opt-in: launch `RedSalamander.exe --perf` or `RedSalamander.exe --perf=PATH`. Debug and ASan Debug write the default JSONL path by default.

Implementation:

- Code: `QueueRedSalamanderMonitorLaunch()` in `RedSalamander/RedSalamander.cpp`

## Shutdown (Debug-layer safe)

When the user closes the main window (`WM_CLOSE`), RedSalamander MUST make a best-effort pass to shut down auxiliary top-level windows before the process exits.

Behavior:

- Confirm cancellation of file operations (do not exit while file operations are still running unless the user explicitly chooses to cancel).
- Close any other **unowned top-level** RedSalamander windows in the same process (Compare Directories, viewers, File Operations popup, item properties, etc.) before destroying the main window.
- If any such window refuses to close or does not respond within the timeout, log a warning and continue shutdown. Do not trap the user behind an unresponsive auxiliary window.
- Post a deferred final-close message after the best-effort pass so auxiliary windows that accepted `WM_CLOSE` can process posted teardown before the main window posts `WM_QUIT`.

Implementation notes:

- Window discovery uses `EnumWindows` filtered by:
  - current process id
  - `GetParent(hwnd) == nullptr` and `GetWindow(hwnd, GW_OWNER) == nullptr`
  - `GetClassNameW(hwnd)` starts with `RedSalamander.`
- Close windows by sending `WM_CLOSE` via `SendMessageTimeoutW` (use a finite timeout; do not hang shutdown on a stuck window).
- If the window is still alive after `WM_CLOSE`, treat the auxiliary close as incomplete but continue process shutdown after logging.

Rationale:

- With the Direct2D debug layer enabled (`d2d1debug3.dll`), outstanding D2D objects can trigger a debug break during process teardown. Closing windows first ensures they run their normal destruction paths and release D2D/D3D resources before exit.

Code:

- `CloseUnownedTopLevelRedSalamanderWindowsForShutdown()` and `kFinalizeMainWindowCloseMessage` handling in `RedSalamander/RedSalamander.cpp`

## Windows session-end durability

Windows logoff, restart, and shutdown notifications use a settings-only path that is deliberately separate
from normal interactive `WM_CLOSE` teardown.

- `WM_QUERYENDSESSION` returns `TRUE` promptly. It does not prompt, cancel File Operations, close windows,
  unload plugins, or write settings.
- A confirmed `WM_ENDSESSION` (`wParam != FALSE`) captures current main-window placement, pane paths and
  view/history state, active pane, and menu/function-bar visibility, then submits the prepared
  RedSalamander settings document as a bounded final request to the serialized save coordinator.
- The final request fences later submissions and is ordered after any already-active write, so an older
  debounced snapshot cannot replace the session-end snapshot.
- A canceled `WM_ENDSESSION` (`wParam == FALSE`) is a no-op.
- Session-end persistence is idempotent. One atomic save owner arbitrates the confirmed-session path and the
  later normal `WM_DESTROY` path, so repeated notifications and normal destruction cannot write the runtime
  snapshot twice.
- The session-end handler does not run normal close prompts, close viewers, destroy the folder window, shut
  down either plugin manager, or update the plugin-configuration schema sidecar. Normal process destruction
  still performs its required teardown after a session-end save, but skips the duplicate settings/schema save.
- Save failure is best-effort: log one error with the settings path and HRESULT, without showing UI or
  preventing Windows from ending the session.

Performance and regression contract:

- Emit `App.Shutdown.SessionEndSettingsSave` with the complete capture/write duration and result HRESULT.
- Commands case `cmd_app_session_end_persists_runtime_state_without_teardown` covers prompt-free query,
  canceled and repeated delivery, changed pane/menu capture, one writer call, live-window preservation, and
  zero normal-teardown entries. It writes `session_end_settings_metrics.json` as deterministic perf evidence.

## Settings

### Schema

- Settings schema version: **16** (the splash setting was introduced in v8)
- New section:
  - `startup.showSplash` (bool, default `true`)

Files:

- Schema source: `Specs/SettingsStore.schema.json`
- Parser/writer: `Common/Common/SettingsStore.cpp`, `Common/SettingsStore.h`

