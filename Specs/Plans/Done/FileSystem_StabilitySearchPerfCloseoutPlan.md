# FileSystem Stability, Search, and Perf Closeout Plan

**Last updated:** 2026-04-26

**Status:** Done - moved after focused code, selftest, spec, and archived validation closeout.

**Scope:** Own FileSystem validation work that is broader than the completed recursive Copy/Move scheduler fix. This plan intentionally excludes the already-closed recursive copy correctness matrix from `Code_DxUiAndFileSystemImprovementPlan.md`.

## Current Checklist

### 1. Recursive Copy/Move Perf Follow-Ups

- [x] Add slower-destination or slower-device evidence if a real tuning decision depends on it.
  - Closed with deterministic virtual slower-provider evidence from `--fileops-selftest --selftest-case=Phase11_BridgePipelineDummyToDummyPerf`.
  - No real-device tuning decision remains in this closeout; future real-device tuning requires a new scoped plan and same-machine archived evidence.
- [x] Add queue-backpressure saturation stress only if future telemetry shows the bounded recursive queue to be tuning-sensitive.
  - Closed with existing bounded scheduler coverage from `--fileops-selftest --selftest-case=Phase7_SharedPerItemScheduler` and throughput/concurrency coverage from `Phase7_CopyMoveConcurrency16Perf`.
  - No current telemetry shows a remaining bounded-queue tuning decision.
- [x] Keep user-facing performance claims tied to archived same-machine evidence; do not publish generic speedup claims without matching `Specs/TestRuns/` data.
  - Durable wording lives in `Specs/FileSystem/FileSystem_FileOperations.md` and `Specs/Testing/Testing_PerformanceValidation.md`.

### 2. Local FileSystem Stability

- [x] Verify module lifetime pinning or equivalent quiet-point guarantees for current local FileSystem worker callback paths.
  - Evidence reviewed in `Plugins/FileSystem/FileSystem.Search.cpp`, `Plugins/FileSystem/FileSystem.Watch.cpp`, and `Plugins/FileSystem/FileSystem.FileOps.cpp`.
- [x] Verify producers stop before callback release and module unload for current search, file-op, and watch paths.
  - Durable quiet-point requirements were merged into `Specs/FileSystem/FileSystem_FileOperations.md`.
- [x] Refresh teardown stress that destroys or unloads the plugin while operations are queued, running, and cancelling.
  - Evidence: `--fileops-selftest --selftest-case=Phase14_PopupHostLifetimeGuard`, `--fileops-selftest --selftest-case=Phase7_WatcherChurn`, and `--commands-selftest --selftest-case=filesystem_local_watch_unwatch_drains_inflight_callback`.
- [x] Validate folder watching under rename/delete/recreate churn:
  - [x] watched-directory callbacks do not race stale state into the UI,
  - [x] callback payload ownership and teardown match the repository payload rules.
  - Evidence: `--fileops-selftest --selftest-case=Phase7_WatcherChurn`, archived below.

### 3. Regex/Search Validation

- [x] Add pathological-pattern search selftests for cancellation and timeout behavior.
  - [x] Added coverage that nested unbounded repetition is rejected before traversal and reports the same single completed progress payload as syntax-invalid regex.
  - [x] Added parallel scan cancellation/fan-in coverage for heavy ready-match batches via `search_local_plugin_parallel_cancel_fanin`.
- [x] Confirm invalid regex errors are reported once and do not leave worker state dirty.
  - Code fix: `Plugins/FileSystem/FileSystem.Search.cpp` now maps `std::regex_error` during name/content regex compilation through the same rejection-reporting path as safety validation failures.
  - Regression: `search_local_plugin_invalid_regex_reports_single_completion`.
- [x] Review current parallel search behavior:
  - [x] worker-count selection for built-in local recursive name-only scan (`searchMaxDirectoryWalkers`, clamped `1..8`; `1` keeps traversal serial),
  - [x] queue fairness between deep and broad directory trees,
  - [x] cancellation latency during heavy enumeration,
  - [x] progress/reporting fan-in when many search workers finish at once.
  - Code fix: `Plugins/FileSystem/FileSystem.Search.cpp` now polls `FileSystemSearchShouldCancel` while draining completed parallel scan result batches and exits terminal cancellation after workers reach the quiet point even when completed result batches remain queued.
  - Regression: `search_local_plugin_parallel_cancel_fanin`.
- [x] Replace any stale projected search speedup notes with measured same-machine evidence.
  - No remaining projected search speedup notes were found in the touched search specs/plans; durable claims remain bound to archived evidence by `Specs/Testing/Testing_PerformanceValidation.md`.

### 4. Closeout Gates

- [x] Update authoritative FileSystem/search specs for any durable behavior discovered here.
  - Updated `Specs/Core/Core_Search.md` and `Specs/FileSystem/FileSystem_FileOperations.md`.
- [x] Archive focused selftest/stress runs under `Specs/TestRuns/`.
  - `Specs/TestRuns/27b63c26e1eb/SearchPerfCloseout/2026-04-26_144913_regex_and_watcher/`
  - `Specs/TestRuns/27b63c26e1eb/SearchPerfCloseout/2026-04-26_153000_final_closeout/`
- [x] Move this plan to `Specs/Plans/Done/` when all rows are complete or split into narrower active owner plans.
  - All rows are complete; this plan is ready for `Specs/Plans/Done/`.

## References

- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Plugins/FileSystem/FileSystem.FileOps.cpp`
- `Plugins/FileSystem/FileSystem.Search.cpp`
- `Plugins/FileSystem/FileSystem.Watch.cpp`
- `Specs/TestRuns/`
