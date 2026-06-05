# Fix S3 shutdown crash: AWS background thread after DLL unload

Status: WIP - mitigation code is present, but the shutdown-crash validation evidence is not archived.

## Closeout audit (`2026-04-25`)

The code now has `AwsSdkLifetime` reference counting around S3 instances and IO objects, a finite default `requestTimeoutMs` of 30000, cached S3 client cleanup before SDK release, request-timeout wiring into client config, and process-shutdown module retention for `FileSystemS3.dll` so its imported AWS CRT DLLs remain mapped until OS teardown. This plan stays in WIP because the required ASan/heavy-S3 shutdown reproduction and proof of AWS CRT thread quiescence are not present in the plan or archived test runs.

Remaining closeout checklist:

- [x] Reference-count AWS SDK initialization/shutdown with `AwsSdkLifetime`.
- [x] Use a finite default request timeout and wire it into S3 client configuration.
- [x] Clear cached S3 clients before releasing the AWS SDK lifetime reference.
- [x] Request process-shutdown module retention so AWS CRT dependencies are not explicitly unloaded during process teardown.
- [ ] Run and archive the ASan heavy-S3 shutdown validation from the recipe.
- [ ] Confirm no `AwsHostResolver` or other AWS CRT thread executes after AWS DLL unload.
- [ ] Update the shutdown ordering evidence here, then move the plan to Done.

## Symptom

Crash during shutdown (often under ASan) with an access violation executing code in an **unloaded AWS DLL**, e.g.:

- A background thread such as `AwsHostResolver` continues running.
- The process unloads `aws-c-io.dll` (dependency of the S3 plugin).
- The background thread later executes inside the unmapped module → AV.

## Root cause

The S3 plugin and/or host unload ordering allows AWS SDK background threads to outlive the last module reference to AWS CRT DLLs.

Even if `Aws::ShutdownAPI()` is called, a shutdown is only safe if **all AWS client objects and internal threadpool resources are released first**, and the SDK has actually quiesced before the last `FreeLibrary` on AWS DLLs occurs.

## Code touchpoints

- S3 SDK lifetime:
  - `Plugins/FileSystemS3/FileSystemS3.Shared.cpp` (`AwsSdkLifetime::AddRef/Release`, `Aws::InitAPI`, `Aws::ShutdownAPI`)
  - `Plugins/FileSystemS3/FileSystemS3.Core.cpp` (`FileSystemS3` ctor/dtor calls `AwsSdkLifetime`)
- Host/plugin module unload:
  - `RedSalamander/FileSystemPluginManager.cpp` (`Shutdown`, `Unload` → `entry.module.reset()`)
  - `RedSalamander/RedSalamander.cpp` (`SaveAppSettings` calls plugin manager shutdown)
- Compare window safety:
  - `RedSalamander/CompareDirectoriesWindow.cpp` holds `wil::unique_hmodule` for per-pane plugin instances until background cleanup runs.

## Required shutdown ordering (quiet point)

To prevent use-after-free across module unload:

1. **Stop producers**: cancel compare workers / stop creating new AWS requests.
2. **Stop UI postings**: stop callbacks posting into windows that may be closing.
3. **Release all AWS clients/streams/readers**: ensure no `S3CrtClient` instances (and no response streams) remain alive.
4. **Shutdown the AWS SDK**: call `Aws::ShutdownAPI()` only after step (3) is true.
5. **Only then unload modules**: allow `wil::unique_hmodule` / plugin manager to drop the last reference to `FileSystemS3.dll` and its dependency chain (`aws-c-io.dll`, etc.).

Notes:

- Infinite request timeouts (`requestTimeoutMs = 0`) make step (1) unreliable because requests can hang forever. Prefer a finite timeout (30s) so cancellation/exit is bounded.
- Do not “fix” this by forcing UI-thread joins; shutdown must stay responsive.

## Verification recipe

1. Enable ASan builds and reproduce: start a heavy S3 operation (compare content or enumeration), then close the app mid-run.
2. Confirm the log ordering:
   - All compare sessions stop background work.
   - All `FileSystemS3` instances are released.
   - `S3: Shutting down AWS SDK` (and completion) occurs **before** plugin unload.
3. Confirm no AWS CRT threads (e.g., `AwsHostResolver`) are executing after AWS DLLs unload.
   - If needed, use WinDbg to list threads/stacks during teardown and confirm quiescence.

