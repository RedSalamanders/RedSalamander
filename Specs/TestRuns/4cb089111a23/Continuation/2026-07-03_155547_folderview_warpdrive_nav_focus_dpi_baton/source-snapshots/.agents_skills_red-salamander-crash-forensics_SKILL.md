---
name: red-salamander-crash-forensics
description: Debug RedSalamander crashes, Fatal Error dialogs, WER/minidump artifacts, sidecar callstack .txt files, selftest hangs/timeouts, access violations, PerfJsonl failures, DxUi prompt teardown crashes, and regression runs that stop without a clean test result. Use when triaging RedSalamander crash dumps, user-pasted call stacks, .build/codex-runs logs, SelfTest/last_run traces, or Specs/TestRuns continuation evidence.
---

# RedSalamander Crash Forensics

## Core Rule

Preserve evidence before rerunning. A new selftest can overwrite `RedSalamander/SelfTest/last_run`, and a new crash can bury the useful dump in a larger pile.

## First Pass

1. Identify whether this is a crash, hang, timeout, fatal dialog, or bad assertion.
2. Locate the newest crash artifacts before opening code.
3. Search for sidecar `.txt` callstack files near every `.dmp`; they are often faster and clearer than loading the dump.
4. Copy the relevant wrapper logs, traces, dump pointers, sidecar stacks, and repro command into a continuation archive under `Specs/TestRuns/<machine>/Continuation/`.
5. Only then rerun a focused repro.

## Artifact Search

Check these locations first. RedSalamander's own crash handler writes paired `.dmp` and `.txt` callstack files under `C:\Users\<username>\AppData\Local\RedSalamander\Crashes` (`$env:LOCALAPPDATA\RedSalamander\Crashes`); check that before generic WER paths.

```powershell
Get-ChildItem "$env:LOCALAPPDATA\RedSalamander\Crashes" -File | Sort-Object LastWriteTime -Descending | Select-Object -First 30
Get-ChildItem "$env:LOCALAPPDATA\RedSalamander\Crashes" -File -Filter "*.txt" | Sort-Object LastWriteTime -Descending | Select-Object -First 10
Get-ChildItem "$env:LOCALAPPDATA\CrashDumps" -File | Sort-Object LastWriteTime -Descending | Select-Object -First 20
Get-ChildItem "$env:LOCALAPPDATA\CrashDumps" -File -Filter "*RedSalamander*.txt" | Sort-Object LastWriteTime -Descending
Get-ChildItem "C:\ProgramData\Microsoft\Windows\WER\ReportArchive","C:\ProgramData\Microsoft\Windows\WER\ReportQueue" -ErrorAction SilentlyContinue
Get-ChildItem ".build\codex-runs" -File | Sort-Object LastWriteTime -Descending | Select-Object -First 30
Get-ChildItem "RedSalamander\SelfTest\last_run" -Recurse -File -ErrorAction SilentlyContinue
```

Use targeted text search when the user provides a process id, exception code, function, or source file:

```powershell
rg -n "ProcessId=73832|ExceptionCode=0xC0000005|PerfJsonl.cpp|FolderView::ProcessThumbnailLoadQueue" "$env:LOCALAPPDATA\RedSalamander\Crashes" "$env:LOCALAPPDATA\CrashDumps" ".build\codex-runs" "Specs\TestRuns"
```

If a broad AppData search is necessary, constrain it by filename and recency. Avoid blind full-recursive searches through all of AppData unless the focused paths fail.

When preserving PowerShell snapshots, pipe formatted output through `Out-String` before `Set-Content`; `Format-List | Set-Content` writes formatter object names instead of readable process details.

## Reading The Stack

Start from the deepest RedSalamander frame that still owns the object lifetime or UI action. NTDLL, CRT, and Win32 top frames often show only where corruption was observed, not where it was created.

Always verify source line numbers against the current worktree. A stack from a previous build can point near, not exactly at, the current line.

For worker-thread stacks, identify what UI or shared state the worker touched and whether the failing data outlives the perf/logging scope, posted message, COM object, or window.

## Known Signatures

### PerfJsonl Access Violation

Signature:

```text
WideCharToMultiByte
Debug::detail::ConvertUtf8
Debug::detail::WritePerfJsonl
Debug::Perf::Scope::~Scope
```

Primary suspicion: a `Debug::Perf::Scope` stored a `std::wstring_view` to a temporary detail/name string, then wrote JSONL after the temporary died. Watch for calls like:

```cpp
perf.SetDetail(path.parent_path().native());
perf.SetDetail(std::format(...));
```

Preferred fix: make the perf scope own the metric/detail strings when ETW or JSONL sinks are enabled, or otherwise extend the producer string lifetime through the scope destructor.

### DxUi Prompt Teardown Crash

Signature: a selftest worker confirms or cancels a DxUi modal prompt, then crashes in TSF/native text input or window teardown.

Primary suspicion: cross-thread synchronous `SendMessageW` destroyed a modal DxUi HWND inside the worker call stack.

Preferred fix: post confirm/cancel close commands for modal prompt teardown. Keep synchronous sends only for non-destructive snapshot/set commands that do not destroy the HWND.

### Fatal Error Dialog Without A Fresh Dump

The app can catch an unhandled exception and display `RedSalamander - Fatal Error` without producing a normal WER dump. Preserve:

- the visible dialog state,
- the process id,
- `.build/codex-runs` wrapper log,
- `RedSalamander/SelfTest/last_run/trace.txt`,
- any sidecar stack text,
- the exact command and filter.

After collecting evidence, kill the hung process before starting another selftest.

### Misleading Selftest Filters

Do not build filters from command labels or free text in trace output. Use exact selftest names from source or a reviewed list. Bad filters can silently run extra commands and create false failures.

### Shared Selftest State

Do not run multiple `Tools/Run-AllTests.ps1` invocations in parallel. They share volatile `RedSalamander/SelfTest/last_run` state and can overwrite each other's evidence.

## Regression Loop

1. Reproduce or collect a narrow artifact.
2. Patch the root cause, not just the crashing call site, when the stack reveals a reusable contract violation.
3. Build with `.\build.ps1 -ProjectName RedSalamander -Configuration Debug`.
4. Run the narrow selftest first.
5. Run the broader affected suite after the narrow gate passes.
6. Archive all commands, logs, stacks, and remaining risks in `Specs/TestRuns/<machine>/Continuation/`.
7. Update the WIP plan with verified status, exact logs, and unresolved blockers.
