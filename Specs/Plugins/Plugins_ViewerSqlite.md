# ViewerSqlite Specification

Last updated: 2026-05-15

## Purpose

This document is the authoritative behavior contract for the built-in SQLite viewer plugin.

It extends the general viewer contract from `Specs/Plugins/Plugins_ViewerPlugins.md` and the shared `DxUi` migration rules from `Specs/UI/UI_DxUiSharedGrid.md`.

## Scope

This specification applies to:

- `Plugins/ViewerSqlite/ViewerSqlite.cpp`
- `Plugins/ViewerSqlite/ViewerSqlite.h`
- `Plugins/ViewerSqlite/ViewerSqliteEngine.*`

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
- `Esc` closes the viewer from the results grid or any focused DX control.

## Data And Query Contract

- Opening the viewer loads the focused SQLite file or selected peer file from `otherFiles`.
- The table selector MUST drive read-only table preview mode.
- Table preview paging MUST be bounded by the configured page size.
- Custom SQL execution is read-only and MUST be bounded by the configured query row cap.
- Sorting applies to table preview results where the current page model supports sortable columns.
- The status strip MUST reflect the current state:
  - ready,
  - loading/open failure,
  - no tables available,
  - table preview row range,
  - custom query result count,
  - truncated query result set when the row cap is hit.

## Configuration Contract

The plugin configuration schema currently exposes:

- `pageSize`
  - maximum rows loaded for table preview pages,
  - valid range `1..1000`
- `queryRowCap`
  - maximum rows materialized for custom read-only queries,
  - valid range `1..100000`
- `directOpenLocalFiles`
  - when true, local SQLite files may be opened directly instead of being copied to a temporary snapshot first

## DXUI, Testing, And Performance Contract

- `ViewerSqlite` is part of the shared full-DX pilot set and MUST remain on the shared `DxUi` host plus `DxUi::Grid` path.
- The real window MUST answer `WM_GETOBJECT` and expose the visible DX subtree expected by the current shared accessibility contract.
- The viewer MUST keep repeated open/close churn stable with zero accepted visible child fallback.
- Results-grid scrolling MUST stay bounded-work and remain suitable for the shared DXUI sustained-scroll validation matrix.

## Related Specs

- `Specs/Plugins/Plugins_ViewerPlugins.md`
- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/Testing/Testing_SelfTests.md`
