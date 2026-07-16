# Fix S3 shutdown crash: AWS background thread after DLL unload

Purpose: prevent the shutdown access violation where AWS SDK background threads (e.g. `AwsHostResolver`) execute inside unloaded CRT DLLs (`aws-c-io.dll` etc.) after `FileSystemS3.dll` is unmapped.

Status (2026-07-02 folder review): WIP — mitigations are landed and verified in code, but the validation evidence is not archived and the runtime-refresh unload path remains unproven (see REMAINING EXPOSURE).

## Landed mitigations (verified 2026-07-02)

- `AwsSdkLifetime` reference counting around SDK init/shutdown — `Plugins/FileSystemS3/FileSystemS3.Shared.cpp:7-58`, used by `FileSystemS3.Core.cpp:55/92` and `FileSystemS3.IO.cpp:299/501`.
- Finite 30s default `requestTimeoutMs` wired into S3 client configuration — `FileSystemS3.h:240`, `FileSystemS3.Shared.cpp:705`; clamped in `FileSystemS3.Configuration.cpp:93-98`.
- Cached S3 clients cleared before releasing the SDK lifetime reference — `FileSystemS3.Core.cpp:85-93`.
- Process-shutdown module retention: plugin exports `RedSalamanderPluginRetainModuleUntilProcessExit` (`FileSystemS3.Factory.cpp:142-145`), honored by the host in `RedSalamander/FileSystemPluginManager.cpp:1290-1314` (selftest-guarded), so AWS CRT DLLs stay mapped until OS teardown.
- The required shutdown ordering contract (stop producers → stop UI postings → release clients → `Aws::ShutdownAPI` → unload) is durably documented in `Specs/Plugins/Plugins_VirtualFileSystem.md` (via FSDeepAudit TrackI).

Symptom, root cause, code touchpoints, and the required-shutdown-ordering prose formerly in this plan are now redundant with the landed code and `Specs/Plugins/Plugins_VirtualFileSystem.md`, which owns that content.

## REMAINING EXPOSURE (this is why the plan stays open)

Module retention applies ONLY in ProcessShutdown mode. Runtime plugin refresh/rediscovery paths (`RedSalamander/FileSystemPluginManager.cpp:524, 809-810, 942, 959, 1084`) unload in FreeLibrary mode and unconditionally unmap `FileSystemS3.dll` and its AWS CRT dependency chain. `FileSystemS3` exports no `RedSalamanderPluginCanUnloadNow`, and its `RedSalamanderPluginShutdown` (`FileSystemS3.Factory.cpp:137-140`) is a no-op. On that path the only barrier against the original AV (`AwsHostResolver` running in unmapped `aws-c-io.dll`) is the UNVERIFIED assumption that `Aws::ShutdownAPI` (`FileSystemS3.Shared.cpp:45`) joins all CRT threads synchronously before returning.

## Open items

- [ ] CL-1: Run and archive the ASan heavy-S3 shutdown validation using the verification recipe below.
- [ ] CL-2: Prove AWS CRT-thread quiescence on the runtime-refresh unload path (FreeLibrary mode), or close the gap in code — e.g. export a `RedSalamanderPluginCanUnloadNow` that answers busy until the SDK has quiesced.
- [ ] CL-3: Record the evidence here, then move the plan to Done.

## Verification recipe

1. Enable ASan builds and reproduce: start a heavy S3 operation (compare content or enumeration), then close the app mid-run.
2. Confirm the log ordering:
   - All compare sessions stop background work.
   - All `FileSystemS3` instances are released.
   - `S3: Shutting down AWS SDK` (and completion) occurs **before** plugin unload.
3. Confirm no AWS CRT threads (e.g., `AwsHostResolver`) are executing after AWS DLLs unload.
   - If needed, use WinDbg to list threads/stacks during teardown and confirm quiescence.

## Next action

CL-1/CL-2: run the ASan recipe with a runtime plugin refresh under heavy S3 load included in the scenario.
