# SQLite Viewer Plugin Plan

Last updated: 2026-03-13

Status: Done

## Progress

- [x] Phase 0. Scope, architecture, and repo-fit review
- [x] Phase 1. Dependency and solution wiring
- [x] Phase 2. Read-only SQLite engine
- [x] Phase 3. Viewer window, async execution, and paging UX
- [x] Phase 4. Host integration, defaults, and docs
- [x] Phase 5. Automated tests, build verification, and cleanup

## Goal

Add a new optional viewer plugin, `ViewerSqlite.dll`, that:

- opens SQLite databases in read-only mode
- works as an optional plugin loaded from `Plugins\`
- handles large tables without loading entire result sets into memory
- allows user-entered read-only SQL queries
- stays within RedSalamander project constraints: WIL RAII, C++23, Unicode UTF-16, no owning raw COM pointers, no `catch (...)`, bounded async work, and safe cross-thread payload posting

## Recommended Shape

Build this as:

- `Plugins/ViewerSqlite/`
  - plugin COM shell (`IViewer`, `IInformations`)
  - Win32 viewer window with common controls
  - async work dispatch using `TrySubmitThreadpoolCallback`
- `Plugins/ViewerSqlite/ViewerSqlite.Engine.*`
  - read-only SQLite source/session helpers
  - schema enumeration
  - paged table reads
  - read-only query validation and execution
- `Tests/ViewerSqliteTests/`
  - console test executable for deterministic engine tests

## Key Design Decisions

### 1. Keep the data view bounded

Do not build a v1 custom Direct2D data grid.

Use:

- standard Win32 controls for file selection, table selection, query entry, paging, and commands
- a report-style list view for row display
- paged reads (`LIMIT` / `OFFSET`) for table preview

Reason:

- this is the lowest-risk way to support very large tables
- memory use stays bounded by page size instead of table size
- it fits the existing Windows-native plugin style

### 2. Read-only means enforced, not advisory

Custom SQL must be:

- a single SQLite statement
- prepared successfully
- accepted only when `sqlite3_stmt_readonly(...)` returns true

Rejected examples:

- `INSERT`
- `UPDATE`
- `DELETE`
- `ATTACH`
- multiple statements in one entry

### 3. Handle non-local files by staging a local snapshot

SQLite works best on local files, especially for metadata queries and large scans.

Recommended open path:

- local builtin filesystem + absolute Win32 path: open directly with SQLite in read-only mode
- other filesystems / archives / remote sources: copy to a temp local snapshot through `IFileReader`, then open the snapshot read-only

This keeps the viewer compatible with the plugin filesystem model without forcing SQLite onto virtual paths.

### 4. Large-table UX

The v1 large-table path should be:

- auto-enumerate tables and views from `sqlite_schema`
- default preview query:
  - `SELECT * FROM "<table>" LIMIT <pageSize> OFFSET <pageOffset>`
- next/previous page controls
- status text showing current page and whether more rows exist

Custom SQL is supported, but table preview is the primary large-data experience.

## Planned Phases

### Phase 0. Scope, architecture, and repo-fit review

- [x] Review viewer plugin loading and optional plugin shape
- [x] Review async/threading constraints and posted-payload rules
- [x] Choose temp-snapshot strategy for non-local databases
- [x] Choose paged list-view UX over a custom grid for v1

### Phase 1. Dependency and solution wiring

- [x] Add `sqlite3` to `vcpkg.json`
- [x] Add `ViewerSqlite` plugin project
- [x] Add `ViewerSqliteTests` console test project
- [x] Add both projects to `RedSalamander.sln`
- [x] Add a new custom viewer async message id in `Common/WindowMessages.h`

### Phase 2. Read-only SQLite engine

- [x] Add UTF-8 / UTF-16 conversion helpers local to the plugin engine
- [x] Add SQLite RAII wrappers for database and statement handles
- [x] Add direct-open detection for `builtin/file-system`
- [x] Add temp snapshot copy for non-local files via `IFileReader`
- [x] Add schema enumeration from `sqlite_schema`
- [x] Add identifier quoting for table preview SQL
- [x] Add paged row loading for table preview
- [x] Add read-only validation for custom SQL
- [x] Add bounded row materialization and cell formatting

### Phase 3. Viewer window, async execution, and paging UX

- [x] Create the viewer window class and lifetime management
- [x] Initialize posted-payload tracking on create and drain on destroy
- [x] Add file combo, table combo, query edit, run/reset buttons, paging buttons, result list, and status line
- [x] Queue background open/query work with `TrySubmitThreadpoolCallback`
- [x] Hold module lifetime with `AcquireModuleReferenceFromAddress(...)`
- [x] Ignore stale async completions by request id
- [x] Implement `Open`, `Close`, `SetTheme`, and `SetCallback`

### Phase 4. Host integration, defaults, and docs

- [x] Add default extension mappings for SQLite file extensions
- [x] Add resources, localized strings, and plugin metadata
- [x] Update viewer docs to mention the SQLite viewer

### Phase 5. Automated tests, build verification, and cleanup

- [x] Add engine tests for table enumeration
- [x] Add engine tests for paged reads on a large table
- [x] Add engine tests for read-only query acceptance
- [x] Add engine tests for write-query rejection
- [x] Build the plugin and test project
- [x] Run the test executable and confirm pass/fail output

## Closeout Notes

- Work landed under `Plugins/ViewerSqlite/` with deterministic coverage under `Tests/ViewerSqliteTests/`.
- Archived validation evidence includes `Specs/TestRuns/4cb089111a23/ViewerSqliteTests/2026-04-22_131600_typography_phase2b/stdout.txt`, covering engine enumeration, paged reads, sorted reads, read-only/write-query behavior, DxUI host/UIA exposure, bounded scrolling, open-close churn, paging/sort flows, selection, tab traversal, and theme-cycle legibility.
- Additional focused bounded-scroll evidence is archived under `Specs/TestRuns/4cb089111a23/ViewerSqliteTests/2026-04-08_095009_000/`.
- No implementation work remains for this plan; follow-up DxUI grid hardening lives in the broader shared-grid and UI modernization plans.

## Verification Targets

The feature is not complete until all of the following work:

- opening a local `.sqlite` or `.db` file launches `ViewerSqlite.dll`
- the viewer lists tables/views without blocking the UI thread
- selecting a table shows only one page of rows at a time
- navigating pages does not grow memory with total table size
- entering a read-only `SELECT` query returns rows
- entering a write query is rejected before execution
- closing the window during background work does not leak payloads or unload the DLL early
- automated tests pass
