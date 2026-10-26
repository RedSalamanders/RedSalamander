# File-operation exit access violation — preserved evidence

Captured from the user report on 2026-08-27 at 13:33:51 +02:00 before running a new reproduction.

## Reported scenario

- Exit RedSalamander while a file operation is in flight.
- Result: access violation.

## Crash signature

```text
ExceptionCode=0xC0000005
ExceptionAddress=0x7ff70d70d8f7
ProcessId=41792
ThreadId=10744
Callstack:
00 0x00007FF70D70D8F7 FolderWindow::FileOperationState::StartOperation+0x1E47 D:\RedSalamander\RedSalamander\FolderWindow\.FileOperations.State.Runtime.cpp:931(+236)
01 0x00007FF70D41765A FolderWindow::FileOperationState::AdmitOperation+0x35EA D:\RedSalamander\RedSalamander\FolderWindow\.FileOperations.cpp:3249(+41)
02 0x00007FF70D420EDA FolderWindow::CommandCopyToOtherPane+0x52A D:\RedSalamander\RedSalamander\FolderWindow\.FileOperations.cpp:4732(+635)
03 0x00007FF70D69B283 `anonymous namespace'::OnMainWindowCommand+0x2403 D:\RedSalamander\RedSalamander\RedSalamander.cpp:11573(+0)
04 0x00007FF70D69D40A WndProc+0x72A D:\RedSalamander\RedSalamander\RedSalamander.cpp:12067(+15)
05 0x00007FFF7D7EC366 CallWindowProcW+0x6A6
06 0x00007FFF7D7EB673 SendMessageW+0x553
07 0x00007FFF7D7EB424 SendMessageW+0x304
08 0x00007FFF7D7EB20F SendMessageW+0xEF
09 0x00007FF70D68CDF6 `anonymous namespace'::ExecuteCommandById+0x27F6 D:\RedSalamander\RedSalamander\RedSalamander.cpp:6821(+9)
10 0x00007FF70D68D129 `anonymous namespace'::TryHandleShortcutKeyDown+0x1A9 D:\RedSalamander\RedSalamander\RedSalamander.cpp:6876(+30)
11 0x00007FF70D691D94 RunApplication+0x3544 D:\RedSalamander\RedSalamander\RedSalamander.cpp:8448(+12)
12 0x00007FF70D692657 wWinMain+0x97 D:\RedSalamander\RedSalamander\RedSalamander.cpp:8541(+10)
13 0x00007FF70D7C1409 __scrt_common_main_seh+0x109 D:\a\_work\1\s\src\vctools\crt\vcstartup\src\startup\exe_common.inl:288(+33)
14 0x00007FFF7D5DCCB7 BaseThreadInitThunk+0x17
15 0x00007FFF7DDEAD6C RtlUserThreadStart+0x2C
```

## Evidence lookup

- No matching fresh `.dmp` or `.txt` sidecar was present under `%LOCALAPPDATA%\RedSalamander\Crashes` or `%LOCALAPPDATA%\CrashDumps`.
- Both `D:\RedSalamander` and this worktree were at commit `c5e960bfea65` when inspected.
- The reported line 931 is the `return S_FALSE` immediately after `ConfirmNonRevertableFileOperation(...)`; the instruction offset must be resolved against the owning function and lifecycle rather than treated as an exact expression-level location.

## Root cause

`HostServices::ShowPromptOnUiThread` runs a nested `PeekMessageW` loop while the File Operations confirmation overlay is active. If that loop dispatched `WM_CLOSE`, `OnMainWindowClose` accepted exit and posted the final main-window destroy. The prompt loop then continued draining the queue and could dispatch that destroy before returning to `StartOperation`. `FolderWindow::OnDestroy` consequently shut down and reset `FileOperationState` underneath the still-active admission stack, producing the reported use-after-free on return from the confirmation boundary.

The shortcut call stack also exposes a second admission window: a queued command could reach `OnMainWindowCommand` after exit was accepted but before the posted final destroy.

## Implemented lifetime fix

- Main-window exit now commits an atomic command fence before auxiliary teardown and before posting final destruction.
- Host prompt shutdown is committed at the same boundary. An active prompt stops its nested queue drain immediately after the close message that committed shutdown, returns the escape result, and unwinds its caller before the queued owner destroy can run.
- New prompts fail with `ERROR_SHUTDOWN_IN_PROGRESS` once that fence is active.
- File Operations admission checks its own shutdown state initially and after each synchronous confirmation boundary.
- The authoritative File Operations lifetime contract and both runtime/source-contract regressions were updated.

## Validation on 2026-08-27

### Compile and link

```text
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -MaxCpuCount 4
MSBuild log: .build/logs/msbuild-20260827_142158_393-pid42140-47b70b89.log
Compiler/linker diagnostics: 0 warnings, 0 errors
Produced: .build/x64/Debug/RedSalamander.exe
```

Visual Studio Professional 2026 updated from 18.6 to 18.10 during validation, temporarily removing and then reinstalling the C++ toolset, Windows 11 SDK, and vcpkg integration. After installation completed, compilation and linking were clean. The repository wrapper subsequently rejected the new vcpkg tracking-log shape because `Common.write.1u.tlog` now lists source dependency DLLs under `.build/vcpkg_installed/...` as outputs outside `.build/x64/Debug`; therefore the wrapper exited nonzero after successful linking and did not publish a build receipt. No artifact-contamination marker was created.

### Exact selftests against the freshly linked executable

Each process completed with exit code 0 and advanced its suite `results.json`:

```text
--commands-selftest --selftest-case=file_operations_phase0_typed_contract_source_guard
  passed=1 failed=0 skipped=0

--commands-selftest --selftest-case=cmd_app_prompt_uses_alert_overlay_window
  passed=1 failed=0 skipped=0
  Includes live shutdown-unwind validation: the real DxUi prompt returned Cancel,
  removed its overlay, and observed the main owner window still alive at return.

--fileops-selftest --selftest-case=Phase10_PermanentDelete
  passed=3 failed=0 skipped=0 (Setup + exact case + Cleanup)

--fileops-selftest --selftest-case=Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker
  passed=3 failed=0 skipped=0 (Setup + exact case + Cleanup)
```

`git diff --check` passed; Git emitted only the repository's existing LF-to-CRLF conversion warnings.
