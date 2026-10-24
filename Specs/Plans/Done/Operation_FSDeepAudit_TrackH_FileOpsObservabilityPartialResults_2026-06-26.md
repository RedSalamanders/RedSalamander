# Operation FS Deep Audit Track H - FileOps Observability and Partial Results - 2026-06-26

## Status

- State: Done - moved to `Specs/Plans/Done/` after implementation and focused validation.
- Priority: P2 observability; promote if a sub-item hides destructive failure.
- Scope: Directory size, watch overflow/rearm, diagnostics writes, issue deduplication, temp cleanup, and progress underestimation.

## Problem

Secondary paths can still hide partial work or collapse distinct failures, making incomplete operations harder for users and tests to detect.

## Targets

- `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`
- `Plugins/FileSystem/FileSystem.Watch.cpp`
- `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp`
- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp`
- `Common/MonitorIssueStore*`
- Relevant specs

## Tasks

1. [x] `GetDirectorySize` now preserves best-effort totals and sets `result.status = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)` for denied/vanished non-root descendants. Unexpected later traversal errors can still override the partial status as fatal.
2. [x] Local directory watch now treats zero-byte successful completions as overflow/resync-required and emits overflow when re-arming `ReadDirectoryChangesW` fails.
3. [x] File-operation issues report export now checks `WriteFile` byte counts for BOM, header, and issue rows; short writes fail export instead of silently producing truncated diagnostics.
4. [x] Issues pane de-duplication now includes task, operation, status/severity/category/message, source/destination path, concurrency mode, and source/destination storage class.
5. [x] `.rs_tmp_*` orphan policy is defined: failed/cancelled bridge transfers best-effort delete and log cleanup failures; crash leftovers are safe/user-visible cleanup items, hidden from search/index results, and no blind startup sweep is allowed without conservative age/handle checks.
6. [x] Progress underestimation is made explicit in specs through `preCalcSkipped`, best-effort totals, completed file/folder counts, and `precalc.result`; partial/unavailable pre-calc totals are advisory.
7. [x] Deterministic guards added where practical: Pester source-contract guards for DirectoryOps partials, Watch overflow/rearm, diagnostic short writes, Issues pane dedupe context, and `.rs_tmp_` indexing; compare/search low-hardening guard now requires `.rs_tmp_` filtering.
8. [x] Authoritative specs updated:
   - `Specs/FileSystem/FileSystem_FileOperations.md`
   - `Specs/Plugins/Plugins_VirtualFileSystem.md`
   - `Specs/Core/Core_Search.md`

## Implementation Notes

- `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`
  - Added partial-status marking for access-denied/path-vanished child traversal failures.
  - Root open failures remain hard failures.
  - Fatal traversal errors can override an earlier partial marker.
- `Plugins/FileSystem/FileSystem.Watch.cpp`
  - Zero-byte successful completions enqueue overflow.
  - Failed immediate re-arm logs and enqueues overflow so consumers resync.
- `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp`
  - `ExportTaskIssuesReport(...)` treats short writes as export failure.
- `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp`
  - Issue stable/dedupe key includes operation and storage/concurrency context.
- `Common/LocalSearchIndexCore.cpp`
  - `IsRedSalamanderStagedTempName(...)` now also filters bridge `.rs_tmp_*` crash leftovers.

## Validation

```powershell
Invoke-Pester -Path Tools\Tests\TestHarnessSourceContracts.Tests.ps1
.\build.ps1 -ProjectName RedSalamander
.\.build\x64\Debug\PluginContractTests.exe
.\.build\x64\Debug\RedSalamander.exe --compare-selftest --selftest-case=search_low_hardening_smoke
.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase7_WatcherChurn
git diff --check
```

Results:

- `Invoke-Pester -Path Tools\Tests\TestHarnessSourceContracts.Tests.ps1`: passed 30, failed 0.
- `.\build.ps1 -ProjectName RedSalamander`: passed, 0 warnings / 0 errors, log `.build\logs\msbuild-20260627_120058_336.log`.
- `.\.build\x64\Debug\PluginContractTests.exe`: passed; FileSystem debug selftests passed 35/0, S3 115/0, Curl 68/0, MicrosoftDrive 55/0.
- `.\.build\x64\Debug\RedSalamander.exe --compare-selftest --selftest-case=search_low_hardening_smoke`: passed (exit 0).
- `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase7_WatcherChurn`: passed (exit 0).
- `git diff --check`: passed.
