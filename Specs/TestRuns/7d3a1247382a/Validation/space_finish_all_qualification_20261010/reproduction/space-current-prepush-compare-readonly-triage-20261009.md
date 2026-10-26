# Read-only triage: Debug PrePush Compare Directories failure (2026-10-09)

## Run and immutable result

Candidate checkout: `C:\Users\PVB\.codex\worktrees\viewer-space-explorer-analyst\RedSalamander`, HEAD `993c3ff14ce60eb7331a0f36bd54475e227859e9`. Run: `20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f`. Its preserved Compare artifacts are in `D:\RedSalamander.Perf\runs\20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f\artifacts\archive\7d3a1247382a\CompareDirectories\2026-10-09_225204`; the run entry record is under `D:\RedSalamander.Perf\evidence\runs\20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f\entries\selftest.compare-directories`.

`compare_results.json` records `duration_ms: 154011`, `passed: 244`, `failed: 1`, `skipped: 27`. The sole failed case is `local_search_callback_contract` (10 ms), with reason: `Search abort-match case expected ERROR_CANCELLED. hr=0x00000000`. `compare_trace.txt` records the same case as failed and the suite as `CompareSelfTest: FAIL`. No rerun was performed. The outer run is reported by its owner as now running `selftest.commands`; this note does not inspect or alter that run's process state.

## Exact test contract

`RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.Tests.Cases.SearchAndIndex.cpp:32-85` prepares one `match.txt` file containing `contract`, then issues a recursive, include-files, force-scan query with wildcard `*.txt` and content search disabled. It exercises success, progress failure, cancel-callback failure, cancel, abort-progress, and finally abort-match in that order. Only the final assertion failed: `Search` should return `HRESULT_FROM_WIN32(ERROR_CANCELLED)` when the match callback aborts (`:81-85`).

`CompareDirectoriesEngine.Tests.Runner.cpp:3759-3795` shows `RecordingSearchCallback::FileSystemSearchMatch` returns `E_ABORT` immediately for `Mode::AbortMatch`. The failed result shows `Search` returned `S_OK` (`0x00000000`) rather than the expected cancellation HRESULT (`0x800704C7`). The result does not separately record callback-call counts or which subcall emitted each metric.

## Recorded metrics

`perf/perf_metrics.jsonl` contains two complete `FileSystem.Search` scans for the `search_callback_contract` root, both forced scan on the same one-file case tree:

- 20:50:29.077Z: `FileSystem.Search` 2,241 µs, `value0=1` (matched entries), `value1=2` (one scanned file plus one directory), `hr=0`; parallel walk 964 µs, one directory, max active workers 1.
- 20:50:29.083Z: `FileSystem.Search` 2,117 µs, `value0=0`, `value1=2`, `hr=0`; parallel walk 854 µs, one directory, max active workers 1.

These are the successful-match call and a later one-file scan consistent with abort-before-match-count-increment. Metrics do not identify a callback invocation directly, nor capture the worker-completion state at the point of cancellation. They support the subcase localization but do not alone prove the branch that lost the HRESULT.

## Source-level candidate, not confirmed runtime causality

The query satisfies `ShouldUseParallelDirectoryWalk` (`Plugins/FileSystem/FileSystem.Search.cpp:399-402`) because it is recursive and content search is disabled; forced scan selects the scan backend. `Search` therefore calls `SearchDirectoryTreeParallelNameOnly` for this directory query (`:2871-2873`). The coordinator drains worker matches and invokes `EmitSearchMatch` (`:955-958`); `EmitSearchMatch` maps callback `E_ABORT` to `ERROR_CANCELLED` (`:1443-1448`). On a callback failure, `RequestParallelDirectoryWalkStop` stores the failed HRESULT in `state.terminalHr` and sets `cancelFlag` (`:487-496`).

There is a concrete ordering path worth checking in that coordinator: `finalHr` starts as `S_OK` (`:893`). The top-of-loop terminal-error branch copies `terminalHr` into `finalHr` when workers are done (`:933-939`). But when the callback fails during result processing, `RequestParallelDirectoryWalkStop` updates the shared terminal HRESULT after the loop already took its local `terminalHr`/`workersDone` snapshot (`:905-929`, `:955-959`). The bottom `if (workersDone) break` (`:993-996`) can then leave the loop without refreshing `finalHr`. If the last worker had already completed at that snapshot, the scan can return the initialized `S_OK` despite the callback failure. The one-directory/one-active-worker metrics are compatible with that ordering, but they do not prove that `workersDone` was true in the failing execution.

This source-level path is a focused candidate for the observed `S_OK`; it is not a proven reconstruction of the exact run. Other runtime paths remain unexcluded by the preserved trace because it lacks per-subcall callback and coordinator-state instrumentation.

## Scope and follow-up

`git diff --name-only HEAD -- Plugins/FileSystem RedSalamander/SelfTest/CompareDirectories` returned no paths. Comparing the same scoped paths against parent `d713d469213659f20d47f842d6980c1c85b24609` also returned no paths. The candidate worktree's tracked changes are in ViewerSpace-related files/specs, `Common/WindowMessages.h`, and a Commands test; the FileSystem search implementation and Compare Directories test files are unchanged in that diff. This scope check does not establish the runtime cause.

Named follow-up: **FileSystem parallel name-only abort-match HRESULT propagation**. In a separately authorized focused validation after the canonical run completes, cover the existing forced-scan one-file abort-match scenario and capture whether callback E_ABORT reaches `EmitSearchMatch`, the worker-completion snapshot, shared `terminalHr`, and returned `Search` HRESULT. If changing code is later authorized, keep cancellation normalization and unrelated tests/coverage intact; validate both early coordinator completion and worker-still-active orderings. Do not alter test expectations to accept S_OK.

No source, tracked test, documentation, index, build, test evidence, or process state was changed. This ignored proposal is the only output.
