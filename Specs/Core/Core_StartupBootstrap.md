# Startup & Bootstrap Specification (RedSalamander.exe)

## Goals

- Improve startup visibility with **ETW-backed metrics** (no-op when ETW is not active).
- Keep **RedSalamanderMonitor** as an **ETW-only** viewer (no shared-memory crash plumbing).
- Provide a **delayed splash screen** for slow startups that does not add startup tax for fast runs.
- Provide an **in-process crash front door** with **minidumps** and a **crash-on-next-launch** UX.

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

When the user closes the main window (`WM_CLOSE`), RedSalamander MUST ensure any auxiliary top-level windows are shut down before the process exits.

Behavior:

- Confirm cancellation of file operations (do not exit while file operations are still running unless the user explicitly chooses to cancel).
- Close any other **unowned top-level** RedSalamander windows in the same process (Compare Directories, viewers, File Operations popup, item properties, etc.) before destroying the main window.
- If any such window refuses to close or does not respond within the timeout, abort the main window close (do not proceed to process teardown while any window is still alive).

Implementation notes:

- Window discovery uses `EnumWindows` filtered by:
  - current process id
  - `GetParent(hwnd) == nullptr` and `GetWindow(hwnd, GW_OWNER) == nullptr`
  - `GetClassNameW(hwnd)` starts with `RedSalamander.`
- Close windows by sending `WM_CLOSE` via `SendMessageTimeoutW` (use a finite timeout; do not hang shutdown on a stuck window).
- If the window is still alive after `WM_CLOSE`, treat shutdown as canceled and keep the app running.

Rationale:

- With the Direct2D debug layer enabled (`d2d1debug3.dll`), outstanding D2D objects can trigger a debug break during process teardown. Closing windows first ensures they run their normal destruction paths and release D2D/D3D resources before exit.

Code:

- `CloseUnownedTopLevelRedSalamanderWindowsForShutdown()` in `RedSalamander/RedSalamander.cpp`

## Settings

### Schema

- Settings schema version: **10** (the splash setting was introduced in v8)
- New section:
  - `startup.showSplash` (bool, default `true`)

Files:

- Schema source: `Specs/SettingsStore.schema.json`
- Parser/writer: `Common/Common/SettingsStore.cpp`, `Common/SettingsStore.h`

