# Startup & Bootstrap Specification (RedSalamander.exe)

## Goals

- Improve startup visibility with **ETW-backed metrics** (no-op when ETW is not active).
- Keep **RedSalamanderMonitor** as an **ETW-only** viewer (no shared-memory crash plumbing).
- Provide a **delayed splash screen** for slow startups that does not add startup tax for fast runs.
- Provide an **in-process crash front door** with **minidumps** and a **crash-on-next-launch** UX.

## ETW Startup Metrics

RedSalamander emits one-time startup milestone events via `Debug::Perf::Emit()`:

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

Implementation:

- Code: `QueueRedSalamanderMonitorLaunch()` in `RedSalamander/RedSalamander.cpp`

## Settings

### Schema

- Settings schema version: **8**
- New section:
  - `startup.showSplash` (bool, default `true`)

Files:

- Schema source: `Specs/SettingsStore.schema.json`
- Parser/writer: `Common/Common/SettingsStore.cpp`, `Common/SettingsStore.h`

