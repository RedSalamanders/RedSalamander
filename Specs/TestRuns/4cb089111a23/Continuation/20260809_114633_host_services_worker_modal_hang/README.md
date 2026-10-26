# Host-services worker modal hang continuation

## Reproduction

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter host_services_worker_marshaling_and_connection_manager_abi
```

- Debug x64 executable: `.build/x64/Debug/RedSalamander.exe`
- TestSandbox run: `20260809T094154Z-88728-062ae7c4f9a5465ca5f5f0a875b55329`
- Process: PID 31088, started 2026-08-09 11:41:55, visible `RedSalamander.ConnectionManagerWindow` titled `Connections`
- At capture time the process reported `Responding=False`; `SendMessageTimeout(WM_NULL, SMTO_ABORTIFHUNG, 1000ms)` timed out and externally posted `IDCANCEL` commands were not dispatched.
- No fresh crash dump or sidecar existed. ProcDump/CDB/WinDbg were not available on `PATH` or under `Tools`; the live selftest traces are preserved beside this note.

## Triage

The prompt accept, cancel, and destroyed-target phases had completed because the only visible test surface was the first modal Connection Manager. Its automation thread called Connection Manager debug snapshot/click helpers directly, crossing the UI-thread ownership boundary. The modal UI thread then stopped dispatching even external `IDCANCEL`, consistent with cross-thread DxUi state access deadlocking message handling. The correction is to drive the modal terminal action exclusively through posted window messages and keep all debug/DxUi access on the UI thread.

This archive is failure evidence only and is not a passing package gate.

## Resolution

- Removed all worker-thread Connection Manager debug-model access; automation now waits on native
  HWND state and posts `IDOK` / `IDCANCEL` for the UI thread to dispatch.
- Full Debug x64 contamination-clearing rebuild:
  `.build/logs/msbuild-20260809_114844_709.log` (0 warnings, 0 errors).
- Final targeted Debug x64 build:
  `.build/logs/msbuild-20260809_115343_331.log` (0 warnings, 0 errors).
- Exact retry passed 1 / 0 in 2.3 seconds with zero disk-audit issues:
  TestSandbox run `20260809T095605Z-39640-5414d40b6d4a4d149d987cc55af9ac78`.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` passed 154 / 154 after the correction.
