# Operation FS Deep Audit Track I - Threading, Secrets, and Plugin Quiet Points - 2026-06-26

## Status

- State: Done.
- Closed: 2026-06-27.
- Priority: P2 lifecycle and credential correctness.
- Scope: Host services secret state, provider/session shared state, plugin refresh/reload safety, and unload quiet points.

## Problem

Host and provider mutable state paths need explicit lock, persistence, and lifetime guarantees so failures do not become wrong-provider operations, credential-state corruption, or callback-after-unload hazards.

## Targets

- `RedSalamander/HostServices.cpp`
- `RedSalamander/FolderWindow.FileSystem.cpp`
- Provider connection/session code identified during drift checks
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `AGENTS.md` only if a repo-wide rule is missing

## Tasks

1. Done - Inventoried shared mutable state: connection JSON, session secret cache, persistent connection data, provider refresh/reload boundaries, callback cookies, and plugin unload sequencing.
2. Done - Verified the touched connection-secret paths stay serialized on the host UI thread and do not run UI-thread bodies on workers when the host window is unavailable.
3. Done - `SetConnectionSecret(persist=TRUE)` now persists or deletes durable credentials before mutating the live session cache or secret authorization state. Persistence failure leaves the old live value intact.
4. Done - Host connection/secret APIs return `HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE)` when the host UI window/owner state is unavailable instead of falling through to UI-thread bodies on the caller thread.
5. Done - Provider reload/refresh behavior was reviewed against existing neutralization contracts; no code change was needed beyond documenting the durable unload quiet point.
6. Done - Plugin unload quiet point is documented in `Specs/Plugins/Plugins_VirtualFileSystem.md`: stop host producers, clear callbacks, cancel/drain provider work, release COM instances, run optional shutdown, unregister resources, then release or retain the module according to process-shutdown state.
7. Done - Added source-contract coverage for host secret lifecycle ordering and preserved existing focused Compare selftests for Windows Hello cache and OAuth refresh token storage.
8. Done - Updated `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Core/Core_ConnectionManager.md`, and `Common/PlugInterfaces/Host.h`.

## Findings

- Real: public host connection/secret APIs could fall through to `*OnUiThread` bodies when `GetInitializedHostWindow()` failed, allowing host-owned state access from a worker path. Fixed by failing with `ERROR_INVALID_WINDOW_HANDLE`.
- Real: persisted secret updates mutated `_sessionSecretByConnectionId` before WinCred persistence. Fixed so durable success happens before live session exposure for persistent writes.
- Real: persisted secret deletion could clear session state before durable deletion succeeded. Fixed so delete failure leaves the session cache unchanged.
- Not found: a concrete callback-after-unload defect in the reviewed provider reload paths. The lasting value here is the documented quiet-point contract and regression guard, not a speculative code rewrite.
- 2026-06-28 revalidation: Compare/FileOps selftests can run before the normal application host window exists. Secret/profile tests now skip only the exact `HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE)` host-connection UI precondition; all other HRESULTs remain failures unless another explicit precondition applies. The durable selftest rule was added to `Specs/Testing/Testing_SelfTests.md`.

## Validation

```powershell
Invoke-Pester -Path Tools\Tests\TestHarnessSourceContracts.Tests.ps1
.\build.ps1 -ProjectName RedSalamander
.\.build\x64\Debug\PluginContractTests.exe
.\.build\x64\Debug\RedSalamander.exe --compare-selftest --selftest-case=windows_hello_cache,oauth_refresh_token_storage
git diff --check
```

Results:

- Pester source-contract suite: PASS, 31 passed.
- RedSalamander Debug build: PASS, `.build\logs\msbuild-20260627_120957_184.log`, 0 warnings / 0 errors.
- Plugin contract tests: PASS; FileSystem 35/0, MicrosoftDrive 55/0, S3 115/0, Curl 68/0 debug selftests.
- Focused Compare selftests: PASS for `windows_hello_cache` and `oauth_refresh_token_storage`.
- `git diff --check`: PASS.
- 2026-06-28 focused host-secret revalidation: `windows_hello_cache`, `oauth_refresh_token_storage`, Google Drive client/refresh-token cases, and OneDrive client-id case skipped only for `ERROR_INVALID_WINDOW_HANDLE`; `content_pending_elided` passed.
- 2026-06-28 full CompareDirectories revalidation: `162 passed / 0 failed / 29 skipped`. Evidence: `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-28_143051/compare_results.json`.
