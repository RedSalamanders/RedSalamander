# Fix S3 shutdown crash: AWS background thread after DLL unload

Purpose: prevent the shutdown access violation where AWS SDK background threads (for example `AwsHostResolver`) execute inside unloaded AWS CRT DLLs after `FileSystemS3.dll` is unmapped.

Status: **DONE 2026-07-17** through Operation Observatory Track 11.

## Closed contract

- AWS initialization is a synchronized, fallible state transition. A runtime reference is published only after `Aws::InitAPI` succeeds; a named exception is factory-visible and leaves the module non-unloadable instead of exposing an invalid SDK.
- Live `FileSystemS3` and ranged-reader objects own AWS runtime references. Cached clients are cleared before the final reference is released.
- `RedSalamanderPluginShutdown()` closes new acquisition and schedules pending multipart-abort reconciliation.
- `RedSalamanderPluginCanUnloadNow()` remains false while cleanup is pending/active, a runtime owner exists, or AWS initialization/shutdown is incomplete. It becomes true only after `Aws::ShutdownAPI` returns.
- Multipart cleanup callbacks carry a module pin through the Windows threadpool callback-return boundary. A queued cleanup item owns its filesystem instance, which in turn keeps the AWS runtime alive.
- Process shutdown continues to honor `RedSalamanderPluginRetainModuleUntilProcessExit()` as an additional defense: normal quiet-point shutdown runs, while the plugin and AWS dependency chain remain mapped for OS teardown.

The durable requirements live in `Specs/FileSystem/FileSystem_S3.md` and `Specs/Plugins/Plugins_VirtualFileSystem.md`.

## Closeout evidence

- [x] CL-1: focused ASan lifecycle validation archived. The original credential-dependent “heavy live S3 operation” recipe was replaced with a stronger deterministic local gate for this defect: eight real `LoadLibrary` → `Aws::InitAPI` → owner-held shutdown → final release/`Aws::ShutdownAPI` → `FreeLibrary` cycles. Every cycle verifies the busy/safe unload states and confirms the module is actually unmapped; the run completed without an AddressSanitizer report.
- [x] CL-2: runtime-refresh unload gap closed in code with explicit shutdown and `CanUnloadNow` exports, synchronized lifetime state, owner gating, and module-pinned cleanup.
- [x] CL-3: code, authoritative specs, deterministic selftests, source contracts, and focused evidence are recorded.

Evidence archive: `Specs/TestRuns/f4e0c8c3bed8/FileSystemS3/2026-07-17_114500_observatory_track11_s3_contracts/`.

Key results:

- Debug and ASan S3/plugin-contract builds: zero warnings and zero errors.
- S3 debug selftests: 161 passed, 0 failed.
- Plugin contracts: passed, including eight runtime-refresh unload cycles.
- Source contracts: 144 passed, 0 failed.
- ASan plugin contracts: passed in 1.928 seconds with no sanitizer report.

No live AWS credentials or external S3 service were used. That is an intentional boundary: the defect was DLL/runtime lifetime, and the deterministic test exercises the actual AWS SDK initialization/shutdown and physical module unload without making backend availability part of the safety proof.
