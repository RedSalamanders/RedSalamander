# Diagnostics: ETW, Debug Logging & Perf

RedSalamander has no log files: every diagnostic is a structured ETW (Event Tracing for Windows) event emitted through a single TraceLogging provider, and `RedSalamanderMonitor.exe` is the real-time consumer. This page is the developer reference for how to emit diagnostics, the on-the-wire event schema, the consumer pipeline, the build/runtime gating that decides what is emitted, and the separate `--perf` JSONL sink. For the wider architecture see [../DeveloperGuide.md](../DeveloperGuide.md) ("Diagnostics" deep dive); for the consumer UI see [../Monitor.md](../Monitor.md); for access/permission issues see [../Troubleshooting.md](../Troubleshooting.md).

## How to emit diagnostics

All of the producer side lives in `Common/Helpers.h` inside `namespace Debug`. Include the header and call the helper that matches the severity; never write your own `OutputDebugString`/log file. Messages use positional `std::format` placeholders (`{0}`, `{1:08X}`), not printf-style.

| Helper | Type | Use when |
| --- | --- | --- |
| `Debug::Error(fmt, ...)` | `Error` | An unexpected failure happened. In Debug builds the formatted text is also mirrored to `OutputDebugStringW`. |
| `Debug::ErrorWithLastError(fmt, ...)` | `Error` | A Win32 call failed; appends ` --> (code) message` from `GetLastError()` via `FormatMessage`. Returns the captured `DWORD` error. |
| `Debug::Warning(fmt, ...)` | `Warning` | A recoverable problem worth surfacing. |
| `Debug::Info(fmt, ...)` | `Info` | Normal but useful flow detail (gated off in Release by default). |
| `Debug::Out(L"...")` / `Debug::Out(type, fmt, ...)` | any | Low-level sink the helpers above funnel into. |

```cpp
#include "Helpers.h"

if (!::SetThreadDpiAwarenessContext(ctx))
{
    Debug::ErrorWithLastError(L"SetThreadDpiAwarenessContext failed for window {0}", windowId);
}
Debug::Info(L"Enumerated {0} items in {1}", itemCount, path);
```

Do not log normal control flow at `Error`/`Warning`. All helpers are `noexcept`: a formatting error falls back to a placeholder string, and `std::bad_alloc` is treated as fatal (`std::terminate`) so the crash pipeline can capture a dump.

### Perf scopes and counters

Timing and counter data is emitted through `namespace Debug::Perf`. A `Perf::Scope` measures a block with RAII and writes one event on destruction:

```cpp
{
    Debug::Perf::Scope scope(L"FolderView.Enumerate");
    scope.SetDetail(path);
    scope.SetValue0(itemCount);
    // ... work ...
} // duration emitted here
```

For point values use `Debug::Perf::Emit(name, detail, durationUs, value0, value1, hr)` or the convenience wrappers `EmitCounter`, `EmitValue`, and `EmitDurationUs`. Perf events are only assembled when a sink is active (`Perf::IsCaptureEnabled()`), so the `Scope` constructor captures a start time only when ETW perf or the JSONL sink is on; an idle scope is nearly free.

### Call tracing (TRACER macros)

The `TRACER*` macros (also in `Common/Helpers.h`) add per-thread call indentation shared with `Debug::Info`/`Warning`/`Error`, plus an elapsed-time print on exit. Indentation is thread-local, so nested calls on the same thread visually nest. They are active only when debug ETW output is enabled (`IsDebugEtwEnabled()`).

| Macro | Logs |
| --- | --- |
| `TRACER` | Exit only (with elapsed ms). Indentation still applies to nested logs. |
| `TRACER_CTX(ctx)` | Exit only, with a context string. |
| `TRACER_INOUT` | Entering and Exiting. |
| `TRACER_INOUT_CTX(ctx)` | Entering and Exiting, with a context string. |

```cpp
void ExpensiveStep()
{
    TRACER; // logs "ExpensiveStep Exiting (12.345ms)" on return
    // ...
}
```

### One provider per module rule

The TraceLogging provider is declared once via `TRACELOGGING_DECLARE_PROVIDER(g_RedSalamanderProvider)` in `Common/Helpers.h`. TraceLogging cannot share a provider handle across DLL boundaries, so every EXE/DLL must own its own storage. Define it in **exactly one** `.cpp` file per module by setting the macro before the include:

```cpp
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"   // TRACELOGGING_DEFINE_PROVIDER lives here
```

All other translation units in the same module just `#include "Helpers.h"` (declaration only). Every module uses the **same** GUID `{440c70f6-6c6b-4ff7-9a3f-0b7db411b31a}`, so a single consumer session subscribes to all producers at once. Registration is lazy and idempotent (`EnsureTraceLoggingRegistered`, guarded by `std::call_once`).

## ETW event schema

Two event names are written to the provider, distinguished by keyword.

| Constant | Value | Events |
| --- | --- | --- |
| `kDebugKeyword` | `0x1` | `DebugMessage` (Error/Warning/Info/Debug/Text) |
| `kPerfKeyword` | `0x2` | `PerfScope` (perf scopes, counters, values) |

Provider name: `"RedSalamanderMonitor"`. Provider GUID: `{440c70f6-6c6b-4ff7-9a3f-0b7db411b31a}` (shared by every producer module and hard-coded in the consumer as `EtwListener::kProviderGuid`). All events are written at level `TRACE_LEVEL_INFORMATION`.

### DebugMessage (`kDebugKeyword = 0x1`)

Built by `EmitEtwEvent` from an `InfoParam` (`BuildInfoParam` stamps the time, PID, TID, and type).

| Field | TraceLogging type | Meaning |
| --- | --- | --- |
| `Type` | `UInt32` | `InfoParam::Type` value (see below). |
| `ProcessId` | `UInt32` | Emitting process id. |
| `ThreadId` | `UInt32` | Emitting thread id. |
| `FileTime` | `UInt64` | `FILETIME` of emit (UTC, 100ns ticks). |
| `Message` | `CountedWideString` | The formatted payload text. |

`InfoParam::Type` is a bit value (`Text=0x0`, `Error=0x1`, `Warning=0x2`, `Info=0x4`, `Perf=0x8`, `Debug=0x10`, `All=0x3F`). The Monitor maps these to its 6-bit display filter mask via `FilterBitForType` (Text `0x01`, Error `0x02`, Warning `0x04`, Info `0x08`, Perf `0x10`, Debug `0x20`).

### PerfScope (`kPerfKeyword = 0x2`)

Written by `Perf::Scope`'s destructor and by `Perf::Emit`.

| Field | TraceLogging type | Meaning |
| --- | --- | --- |
| `Name` | `CountedWideString` | Metric name (e.g. `App.Startup.UntilMessageLoop`). |
| `Detail` | `CountedWideString` | Free-form detail (`"counter"`, `"value"`, `"duration"`, or caller text). |
| `DurationUs` | `UInt64` | Elapsed microseconds (0 for counters/values). |
| `Value0` | `UInt64` | Caller value 0 (e.g. count). |
| `Value1` | `UInt64` | Caller value 1. |
| `Hr` | `UInt32` | `HRESULT` outcome. |

`EmitEtwEvent` increments `g_etwWritten`/`g_etwFailed`; the counters are exposed through `Debug::GetTransportStats()` (a `TransportStats { etwWritten, etwFailed }`).

## Consumer pipeline

`RedSalamanderMonitor.exe` consumes the same provider in real time. The chain is **emit -> ETW session -> TDH decode -> queue -> batch -> render**:

1. **Session start** — `EtwListener::Start` stops any stale `RedSalamanderMonitor_ETW_Session`, calls `StartTrace` (real-time mode, 256 KB buffers, 8-128 buffers), then `EnableTraceEx2` on the provider GUID at `TRACE_LEVEL_VERBOSE` matching any keyword, and `OpenTrace`.
2. **Worker thread** — `ProcessTrace` runs on a `std::jthread` (`ProcessTraceThread`); `EventRecordCallback` -> `HandleEvent` -> `ExtractEventData` decodes each event's properties with the TDH API (slow, but off the UI thread). When a record has no `Message` but has a `Name`, it is treated as a `PerfScope` and a `[perf] ...` line is synthesized (with a duration warning/error marker over 500ms / 1s).
3. **Display filter** — the wired callback applies `ShouldAcceptEtwEventForDisplay` (see self-event suppression below), then calls `ColorTextView::QueueEtwEvent`.
4. **Cross-thread queue** — `QueueEtwEvent` (still on the worker) pushes onto `_etwEventQueue` (a `std::deque` under `wil::critical_section _etwQueueCS`) and posts `WndMsg::kColorTextViewEtwBatch` **only when the queue was empty** (coalescing to one pending message).
5. **Batch intake** — `OnAppEtwBatch` runs on the UI thread, drains up to `kMaxBatchSize` = 200 entries under the lock, reposts the batch message if overflow remains, and appends the chunk via `Document::AppendInfoLines` under a single document write lock. Append/layout/paint stay on the UI thread.

The two-mode (AUTO_SCROLL / SCROLL_BACK) rendering model and its perf metrics are documented in `Specs/Core/Core_RedSalamanderMonitor.md`.

## Build & runtime gating

What actually reaches ETW depends on the build flavor and command-line/environment opt-ins. The decision for `DebugMessage` types is `ShouldEmitMonitorDiagnosticMessageType`; the perf path additionally checks `Perf::IsEnabled()`.

| Producer build | Error / Warning | Info / Perf / Debug | Default JSONL perf |
| --- | --- | --- | --- |
| Debug, ASan Debug | Emitted | Emitted | On |
| Release | Emitted | Off unless opted in | Off unless opted in |

Opt-ins (both `RedSalamander.exe` and `RedSalamanderMonitor.exe` parse the same flags):

| Switch / variable | Effect |
| --- | --- |
| `--etw` | Calls `SetRuntimeMonitorDiagnosticsEnabled(true)` and sets `REDSALAMANDER_DIAGNOSTICS_ETW=1`, enabling Info/Perf/Debug emission and, for the Monitor, its own self-diagnostics and startup status text. |
| `REDSALAMANDER_DIAGNOSTICS_ETW` | Environment opt-in checked by `IsRuntimeMonitorDiagnosticsEnabled` (any value other than empty or `0` enables it). |
| `--perf` | Enables the JSONL perf sink at the default path (see below). |
| `--perf=PATH` | Enables the JSONL perf sink at a custom path. |

Build-flavor detection is `IsMonitorDiagnosticsBuild()` — true for `_DEBUG` / `RS_ASAN_DEBUG_BUILD`, unless `RS_DIAGNOSTICS_RUNTIME_OPT_IN` is defined. Note `--etw` controls *emission* in the producer and *self-event display* in the consumer; it is distinct from the ETW session privilege needed to *start listening*.

### Self-event suppression (Monitor)

By default the Monitor must stay quiet about itself. `ShouldAcceptEtwEventForDisplay` (in `RedSalamanderMonitor/MonitorDiagnostics.h`) drops events whose `processID == GetCurrentProcessId()` and suppresses startup status text. Launching the Monitor with `--etw` (which sets the runtime flag) opts into showing its own self-originated events and startup status.

## The `--perf` JSONL sink

Perf events have a second, independent sink: newline-delimited JSON (JSONL), implemented in `Common/Common/PerfJsonl.cpp` and exported from `Common.dll` so every module shares one set of state. It is wired through `Debug::Perf::ConfigureJsonlOutput(...)` and used by self-test perf gates. Crucially it is independent of ETW enablement: `Perf::Scope` and `Perf::Emit` always write to the JSONL sink when `HasPerfJsonlOutput()` is true, even if ETW perf is off.

- **Default path**: `%LocalAppData%\RedSalamander\Perf\<AppName>_<timestamp>.jsonl` (`GetDefaultPerfJsonlPath`). Selected when `--perf` is passed with no value, or by default in Debug/ASan Debug producer builds of `RedSalamander.exe`.
- **Custom path**: `--perf=PATH`.
- **Environment**: the path and run metadata can also come from `REDSALAMANDER_PERF_JSONL_PATH` (plus `_SCENARIO`, `_BUILD`, `_BRANCH`, `_COMMIT`, `_MACHINE_HASH`, `_RUN_ID`); `ConfigureJsonlOutput` also exports these so child processes inherit the sink.

Each line is one JSON object appended with `FILE_APPEND_DATA` (concurrent-writer safe):

| Field | Meaning |
| --- | --- |
| `timestamp` | UTC ISO-8601 with milliseconds. |
| `process`, `thread` | Emitting PID / TID. |
| `metric`, `detail` | Perf name and detail. |
| `durationUs`, `value0`, `value1`, `hr` | Same payload as the ETW `PerfScope` event. |
| `value`, `unit` | Convenience: `durationUs`/`us` when timed, else `value0`/`count`. |
| `scenario`, `build`, `branch`, `commit`, `machineHash`, `runId` | Run metadata from `ConfigureJsonlOutput`/env. |

## Getting ETW access

Starting a real-time ETW session can fail with `ERROR_ACCESS_DENIED` on machines where the user is not in the **Performance Log Users** group. Run the repo-root helper once, then sign out/in (or reboot):

```powershell
.\init-etw-trace.ps1
# undo with: .\init-etw-trace.ps1 -Remove
```

The Monitor also surfaces the failing error code (`EtwListener::GetLastError`/`GetLastErrorCode`) and offers a relaunch prompt. See [../Monitor.md](../Monitor.md) for the full access walkthrough and [../Troubleshooting.md](../Troubleshooting.md) for general recovery steps.

## Key files

| File / symbol | Role |
| --- | --- |
| `Common/Helpers.h` (`Debug::Error/Warning/Info/ErrorWithLastError`) | Public logging API; all funnel into `Debug::Out`. |
| `Common/Helpers.h` (`Debug::Perf::Scope`, `Emit*`) | Perf events (`PerfScope`) + JSONL feed. |
| `Common/Helpers.h` (`EmitEtwEvent`, `InfoParam`, `kDebugKeyword`, `kPerfKeyword`) | `DebugMessage` writer and metadata struct. |
| `Common/Helpers.h` (`CallTracer`, `TRACER*`) | Per-thread indentation + timing. |
| `Common/Common/PerfJsonl.cpp` | JSONL perf sink (exported from `Common.dll`). |
| `RedSalamanderMonitor/EtwListener.{h,cpp}` | Real-time ETW session + TDH decode. |
| `RedSalamanderMonitor/ColorTextView.cpp` | Cross-thread queue + batch intake + render. |
| `RedSalamanderMonitor/MonitorDiagnostics.h` | Self-event filtering policy. |
