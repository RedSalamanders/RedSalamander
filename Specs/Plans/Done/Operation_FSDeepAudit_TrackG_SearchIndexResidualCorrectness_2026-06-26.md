# Operation FS Deep Audit Track G - Search and Index Residual Correctness - 2026-06-26

## Status

- State: Done 2026-06-27. Ready to move to `Specs/Plans/Done/`.
- Priority: P2 correctness; promote if a sub-item proves data loss/security impact.
- Scope: Search snapshot atomicity, root casing, hardlink/path modeling, conservative prefilters, buffered failure semantics, and SQLite robustness.

## Problem

Malformed local materialization and SQLite maintenance overlap were fixed earlier, but residual index/search correctness items remain and need explicit behavior plus tests.

## Targets

- `Common/LocalSearchIndexCore.cpp`
- `Common/SqliteIndexStore.cpp`
- `Common/SearchServiceBroker.cpp`
- `Specs/Core/Core_Search.md`
- Compare selftests

## Tasks

1. [x] Search snapshots now write to a sibling temp, validate short writes, flush the temp handle, and replace the final snapshot with `MoveFileExW(..., MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)` only after success.
2. [x] SQLite volume lookup now tries exact `root_path` first, then invariant folded root comparison; stored root casing remains the persistent/display text and differently-cased callers do not create duplicate volume rows.
3. [x] Hardlink/multiple-path alias behavior is explicit: the index models one canonical path per file ID. Seed, traversal hydration, and hard-link journal replay set `QueryStats::hardlinkAliasCoverageIncomplete` when they observe distinct paths for one file ID, and the broker preserves that flag.
4. [x] Literal, extension, and prefix prefilters were already post-filtered by the normal name matcher; existing prefilter tests and source guards keep them false-positive-only.
5. [x] Buffered service overflow/failure semantics were already deterministic from Track D/Floodgate: no-callback queries enforce count/storage ceilings, clear partial buffered candidates, and return `ERROR_BUFFER_OVERFLOW`.
6. [x] SQLite `VACUUM` maintenance now runs `PRAGMA quick_check` before manual compaction and automatic auto-vacuum-mode rewrite. Background checkpoint-only maintenance can defer `SQLITE_BUSY`/`SQLITE_LOCKED` with `S_FALSE`; manual/post-VACUUM checkpoint failure remains hard.
7. [x] Added/updated tests and guards:
   - `sqlite_index_store_root_lookup_case_insensitive`
   - `search_low_hardening_smoke` source guards for atomic snapshot replacement, SQLite `quick_check`, busy checkpoint deferral, and hardlink alias stats round-trip.
   - Existing source/SQLite prefilter and buffered overflow tests remain wired.
8. [x] Updated `Specs/Core/Core_Search.md` with the durable contracts.

## Implementation Notes

- `Common/LocalSearchIndexCore.cpp` now uses atomic flushed sibling-temp snapshot saves and reports non-exhaustive hardlink alias coverage.
- `Common/SqliteIndexStore.cpp` now performs folded root lookup fallback and quick-check-gates full `VACUUM`.
- `Common/SearchServiceBroker.cpp` now packs/unpacks the hardlink alias coverage stats flag.
- `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp` adds direct SQLite root-casing coverage and extends the low-hardening source guard.

## Validation

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild
git diff --check
```

2026-06-27 results:
- `.\build.ps1 -ProjectName RedSalamander` passed with 0 warnings / 0 errors (`.build\logs\msbuild-20260627_114901_080.log`).
- Focused Compare self-tests passed:
  - `search_low_hardening_smoke`
  - `sqlite_index_store_root_lookup_case_insensitive`
  - `sqlite_index_store_automatic*`
