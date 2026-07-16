# ViewerSqlite Specification

Last updated: 2026-07-11

## Purpose

This document is the authoritative behavior contract for the built-in SQLite viewer plugin.

The dependency floor is SQLite 3.53.2 or newer. Do not lower the `vcpkg.json` minimum or ship a resolved build below that floor; SQLite parser/security fixes are part of the untrusted-database viewing contract.

It extends the general viewer contract from `Specs/Plugins/Plugins_ViewerPlugins.md` and the shared `DxUi` migration rules from `Specs/UI/UI_DxUiSharedGrid.md`.

## Scope

This specification applies to:

- `Plugins/ViewerSqlite/ViewerSqlite.cpp`
- `Plugins/ViewerSqlite/ViewerSqlite.h`
- `Plugins/ViewerSqlite/ViewerSqlite.Engine.*`

## Identity

- Plugin id: `builtin/viewer-sqlite`
- Title: `SQLite Viewer`

## Window Contract

- `ViewerSqlite` MUST open as a top-level viewer window following the shared viewer-window contract in `Specs/Plugins/Plugins_ViewerPlugins.md`.
- The visible shell MUST be fully DX-hosted.
- The window exposes:
  - an `Other Files` filename combo box that is enabled when more than one peer file is available,
  - a table selector combo box,
  - previous/next page buttons for table preview paging,
  - a query text field,
  - a `Run Query` command,
  - a `Table Preview` command,
  - a results grid,
  - a status strip.
- Embedded preview mode MUST preserve source-pane focus and must not show standalone title/menu/header chrome beyond the embedded DX surface.

## Keyboard Contract

- After a database page or query result is loaded, keyboard focus MUST land on the results grid as the main viewer surface.
- `Tab` advances from the results grid through the file combo, reload button, table combo, previous/next page buttons, query field, `Run Query`, `Table Preview`, and back to the results grid. `Shift+Tab` walks the same controls in reverse.
- While focus is inside a combo box or query field, that control owns its normal arrow/editing/Enter behavior. After file selection, table preview, paging, or query completion, focus SHOULD return to the results grid when rows are available.
- `Esc` from focused controls returns focus to the results grid; `Esc` from the focused results grid posts `WM_CLOSE` and closes the idle viewer.

## Data And Query Contract

- Opening the viewer loads the focused SQLite file or selected peer file from `otherFiles`.
- Every open is isolated from later source mutations:
  - a built-in local path is copied through SQLite's backup API while a source read transaction is held, so committed WAL content at open is included consistently;
  - the backup is written to a private snapshot whose lifetime handle denies `FILE_SHARE_DELETE`; statements and the cached connection close before that handle is reset and the file is deterministically deleted. Closed exact-prefix artifacts left by process termination are scavenged from the resolved Windows temp directory before the next snapshot is created, while live cross-process snapshots cannot be removed;
  - a virtual filesystem is copied through `IFileReader` into the same deterministic-close plus next-start stale-scavenging lifetime, with an 8 GiB byte cap, bounded chunks, cancellation checks, advertised-size-before/after validation, and exact byte-count validation;
  - a virtual main database whose SQLite header advertises WAL mode is rejected because the filesystem interface cannot atomically snapshot the `-wal`/`-shm` sidecars. Even in rollback mode, size checks cannot detect an adversarial same-size in-place rewrite, so virtual providers are expected to present a stable file for the duration of the copy.
- One read-only `SQLITE_OPEN_NOMUTEX | SQLITE_OPEN_PRIVATECACHE` connection is opened on the private snapshot and reused for the source lifetime. Every operation holds the source connection mutex, so that NOMUTEX handle is never entered concurrently.
- The table selector MUST drive read-only table preview mode.
- Table preview retains the public `LIMIT/OFFSET` page API. Page size is capped at 1000 rows; SQLite VM steps, elapsed time, materialized cell count, and materialized UTF-16 characters are bounded and cancellable by request generation. Deep-offset seek/keyset pagination is a separate optimization and is not claimed by this contract.
- The frozen snapshot makes sequential offset pages stable against concurrent source writes. Sorting applies only to table preview results, and a requested sort ordinal is validated against prepared table metadata before it is interpolated as a numeric `ORDER BY` ordinal.
- Custom SQL execution is read-only and bounded by the configured query row cap plus the same VM/time/result budgets.
- Table names retain an exact bounded internal identifier for quoting and selection, while table/header display strings remove control and bidirectional-override characters and are capped before application-owned materialization. Cell display text is capped near 4K UTF-16 code units. Raw SQLite diagnostic text is sanitized before it reaches status UI.
- The status strip MUST reflect the current state:
  - ready,
  - loading/open failure,
  - no tables available,
  - table preview row range,
  - custom query result count,
  - truncated query result set when the row cap is hit.

## Localization Contract

- User-visible chrome, status text, bounded-work/cancellation messages, snapshot failures, query failures, value labels, and diagnostic context MUST be loaded from `ViewerSqliteResources.rc` through `LoadStringResource(...)` or `FormatStringResource(...)`; engine helpers resolve the owning `ViewerSqlite` module resource instance instead of embedding English fallback text in the engine.
- Raw SQLite diagnostic detail may be retained only after the existing control/bidirectional sanitization and length cap. It is composed with a resource-backed localized context string; raw provider text must not replace the localized operation description.
- Every supported `ViewerSqlite` satellite MUST preserve the embedded source token set except for explicitly documented language-neutral ids. Formatted translations MUST preserve the exact indexed placeholder tokens and format specifications from the source string; translators may reorder positional tokens but may not add, drop, duplicate, renumber, or change them.
- `Tools/Tests/ResourceLocalizationContracts.Tests.ps1` is the authoritative token/satellite guard. It must continue to cover positional placeholder safety, documented embedded-only ids, and satellite id completeness for `Plugins/ViewerSqlite`. Compiled localization coverage must continue to prove that a registered resource owner resolves its satellite until the final matching unregister.

## Configuration Contract

The plugin configuration schema currently exposes:

- `pageSize`
  - maximum rows loaded for table preview pages,
  - valid range `1..1000`
- `queryRowCap`
  - maximum rows materialized for custom read-only queries,
  - valid range `1..100000`
- `directOpenLocalFiles`
  - legacy compatibility key; both values now preserve the private-snapshot contract,
  - built-in local files always use the consistent SQLite backup path and are never exposed as a changing live/WAL database,
  - the current schema labels this as the optimized local snapshot preference; it MUST NOT restore direct live-file viewing

## DXUI, Testing, And Performance Contract

- `ViewerSqlite` is part of the shared full-DX pilot set and MUST remain on the shared `DxUi` host plus `DxUi::Grid` path.
- The real window MUST answer `WM_GETOBJECT` and expose the visible DX subtree expected by the current shared accessibility contract.
- The viewer MUST keep repeated open/close churn stable with zero accepted visible child fallback.
- Results-grid scrolling MUST stay bounded-work and remain suitable for the shared DXUI sustained-scroll validation matrix.
- Deterministic engine coverage MUST prove one cached connection, serialized concurrent callers, exact sequential offset pages, invalid sort-ordinal rejection, generation cancellation, VM-budget interruption, WAL-inclusive local snapshot isolation, local/virtual byte ceilings, virtual-WAL refusal, deterministic close deletion, closed exact-prefix stale-artifact scavenging without deleting a live snapshot, bounded control/oversized table/header/cell rendering, and resource-backed status/diagnostic composition.
- Performance evidence uses `viewer.sqlite.snapshot_us` and `viewer.sqlite.page_us`; result records must identify local-backup versus virtual-copy acceptance/rejection and page offset/row counts.

## Related Specs

- `Specs/Plugins/Plugins_ViewerPlugins.md`
- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/Testing/Testing_SelfTests.md`
