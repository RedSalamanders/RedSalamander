# ViewerSpace Plugin Specification

## Overview

**ViewerSpace** (`builtin/viewer-space`) is a built-in **viewer plugin** that visualizes **disk usage** for a folder as a **Direct2D/DirectWrite treemap**.

Goals:
- compute folder and file sizes asynchronously (non-blocking UI)
- render progressively as results arrive (real-time expansion)
- allow interactive drill-down (into subfolders) and drill-up (back to parent)
- respect the host theme (colors, dark/high-contrast, rainbow, DPI)

## Invocation (Host Integration)

ViewerSpace is invoked from a **FolderView pane**:
- **Shortcut**: `Shift+F3`
- **FolderView context menu**: `View Space`

Target path selection:
1. If exactly **one directory** is selected in the active pane, open ViewerSpace for that directory.
2. Otherwise open ViewerSpace for the pane’s **current folder**.

The host opens the plugin directly by id (not via file-extension association).

## ViewerOpenContext Contract

The host passes:
- `ViewerOpenContext.fileSystem` (active filesystem instance)
- `ViewerOpenContext.fileSystemName` (localized filesystem name for display)
- `ViewerOpenContext.focusedPath` (filesystem-internal UTF-16 folder path)

`selectionPaths` and `otherFiles` are unused by ViewerSpace and SHOULD be null/empty.

Path semantics:
- `focusedPath` is treated as an **opaque filesystem-internal path string** (it may not be a valid Win32 path).
- ViewerSpace MUST pass the same path string (and child paths derived from it) back to `IFileSystem` APIs; it MUST NOT enumerate with Win32 file APIs.

## Configuration

ViewerSpace exposes a JSON configuration schema (`GetConfigurationSchema`) and accepts configuration via `SetConfiguration`.

Supported keys (defaults):
- `topFilesPerDirectory` (96): compatibility floor for largest-file retention per directory. The scanner may retain more file candidates when pixel/detail budgets allow it; remaining files are grouped into exact “Other” totals (`0` groups all files only when the effective adaptive floor is also disabled by implementation policy). The current adaptive budget is derived from the visible treemap area, then capped at the 20k file-candidate ceiling.
- `scanThreads` (1 in persisted config): number of background threads used to scan sibling subtrees in parallel within a single ViewerSpace scan. For the Win32 filesystem, the runtime default is at least `min(hardware_concurrency, 8)`; non-Win32 providers stay at `1` unless explicitly configured.
- `maxConcurrentScansPerVolume` (1): throttles how many ViewerSpace instances scan the same volume at once.
- `cacheEnabled` (`true`): enables the in-memory scan cache.
- `cacheTtlSeconds` (60): how long cached scans remain reusable.
- `cacheMaxEntries` (1): maximum cached roots kept in memory.

## UI / UX

### Layout
- **Header**: current path (left) + scan status (right) + live progress, with **Up** (left) and **Cancel** (right, while scanning).
  - Path display uses a middle ellipsis when needed, preserving the last segment where possible.
- **Treemap**: squarified treemap representing the current node’s children.

Auto-expansion:
- For **large-enough directory rectangles**, ViewerSpace renders the **next level inside the rectangle** (recursively), reserving a label header for the folder name and using the remaining area for children.
- Auto-expansion is driven by **relative area** (fraction of the visible treemap) plus pixel/detail guards: it keeps expanding dominant rectangles while scanning at roughly **8%** of the view, then spends the settled detail budget on still-large folders down to roughly **3.5%** of the view, with hard depth and draw-item caps.

### Progressive Rendering
While scanning:
- treemap begins rendering immediately
- rectangles expand/settle as sizes are discovered
- an indeterminate progress indicator runs in the header (theme accent color; rainbow-aware)
- the header also shows:
  - line 1: path (left) and `Scanning…` / `Queued…` (right)
  - line 2: `N items (X folders, Y files)` (left) and `Processing: <folderName>` (right, no full path)
  - line 3: human-readable total size and raw bytes
- All numeric values (counts, percentages, bytes) follow Windows regional settings (digit grouping + decimal separator); the formatting locale is cached and invalidated on `WM_SETTINGCHANGE`.
- directories that are not yet complete show a small “incomplete” spinner
- large-enough incomplete tiles show a diagonal watermark:
  - while a scan is active: `In Progress`
  - if a scan ends before completion: `Scan Incomplete`
  - completed tiles show no watermark
- For very large incomplete tiles, the scan state is also shown as a second line under the folder name in the tile header (and the diagonal watermark is suppressed to avoid clutter).
- tiles in `Queued` / `NotStarted` / `Scanning` state are dimmed; tiles in `Scanning` state also show centered spinner overlays (current view root children; up to `scanThreads` spinners, one per concurrently scanned subtree); if no eligible tile is visible/large enough, show a single fallback spinner centered in the treemap area (rainbow-aware, slower animation)
- When a scan completes (state transitions to `Done`), show a brief “Scan completed” toast/overlay in the treemap area.

### Interaction
- Mouse: click a directory rectangle to drill down; click header Up to drill up.
- Right-click a treemap tile to open a context menu:
  - `Focus in pane`: navigate the active FolderView pane to the tile’s parent folder and focus the file/folder (when resolvable).
  - `Zoom in` / `Zoom out`
  - Standard FolderView file/folder commands (Open/Open With/Delete/Move/Rename/Copy/Paste/Properties), excluding `View Space` and debug-only items.
  - The popup surface MUST use the shared DxUI context-menu renderer so the menu matches the viewer theme/backdrop contract instead of falling back to a native `TrackPopupMenu` surface.
- Up stays available during scanning; if the current scan root has no parent in the current model, Up restarts a scan at the parent path. Up is disabled when the scan root is already at a volume/share root.
- While scanning, click header **Cancel** or press `Esc` from the treemap surface to stop the scan without closing the viewer.
- Double-clicking an aggregated “Other” bucket triggers a focused action to explore that part (typically drilling into the parent directory or rescanning it with a higher `topFilesPerDirectory`).
- Hover: show an in-canvas Direct2D/DirectWrite tooltip overlay with name/path, size, and share for the tile under the cursor; scan state is shown only while in-progress (not shown for `Done`); tooltip tracks the cursor within the tile (no “stuck” tooltip position), wraps at about 420 DIPs, uses viewer theme/high-contrast colors, and must not create a native `TOOLTIPS_CLASS` window or send `TTM_*` messages.
- Keyboard:
  - `Backspace` / `Alt+Up`: drill up
  - `F5`: refresh/rescan current node (forces a full rescan; bypasses cache)
  - `Esc`: cancel scan if scanning; otherwise close viewer
- The window menu bar still sources commands from `IDR_VIEWERSPACE_MENU`, but the visible top chrome is rendered through the shared `RedSalamander.DxNativeMenuBar` host instead of a native owner-drawn menu bar. The live window menu handle is detached after attach, and `Alt`, `F10`, and menu mnemonics continue to route through the DxUi menu bar.

### Labels
- Item name + compact size label (when there is enough space).
- Text is ellipsized via DirectWrite trimming (avoids mid-glyph clipping without per-tile layouts).
- Aggregated “Other” buckets show item counts (and folder/file breakdown) plus size.
- Folder tiles reserve a header strip for the name (when there is enough height); file tiles use a small “dog-ear” corner fold that reveals the parent tile color behind it, and a matching cut-corner outline to reinforce the fold (stronger file/folder distinction).
- When a tile is large enough to show a line of text, ViewerSpace keeps at least one readable name line visible (prefer showing the beginning of the name).

## Rendering Requirements

- Direct2D 1.1 + DirectWrite
- ViewerSpace renders through WIL-owned D3D11/DXGI/Direct2D device-context resources (`ID2D1DeviceContext` + flip-model swap chain target bitmap). Ordinary `WM_SIZE` must resize the swap chain and rebuild only target-size-dependent resources; it must not recreate the D3D device or size-independent brushes/DirectWrite text formats. Handle `D2DERR_RECREATE_TARGET`, `DXGI_ERROR_DEVICE_REMOVED`, and `DXGI_ERROR_DEVICE_RESET` by discarding device-dependent resources, invalidating the static cache, and scheduling a follow-up paint that recreates the renderer.
- Settled treemap frames may be cached in a device-dependent bitmap. Hover/tooltip-only paints replay the static bitmap and draw dynamic overlays without rerecording the static cache. Static cache invalidation is required on model/layout changes, view-root changes, resize, DPI/theme changes, and device loss.
- Tile rendering uses precomputed LOD tiers. Tiny tiles are fill-only, small tiles avoid text, medium+ tiles may draw names, large/hero tiles may draw richer file/folder details. Render metrics must report tile, text, visible, and culled counts.
- Layout rebuilds compute stable item identities, target rectangles, LOD tiers, and effective tile areas before paint. The static treemap paint path may replay a settled bitmap; when it must record, it should avoid resolving item metadata repeatedly inside hot loops and must keep text/dog-ear/watermark work behind LOD thresholds.
- Use `ViewerTheme.backgroundArgb`, `textArgb`, and `accentArgb` for styling; in high-contrast, prioritize strong borders/readability.
- First-show background MUST match the current theme (avoid a white/high-contrast flash before the first Direct2D paint); implement via a themed `WNDCLASSEXW::hbrBackground` + `WM_ERASEBKGND` behavior consistent with `Specs/Plugins/Plugins_ViewerPlugins.md`.

## Performance Instrumentation

ViewerSpace emits the `viewer.space.*` metric family through `Debug::Perf` so behavior can be baselined and compared during renderer/layout/scanner optimization. The current contract covers:

- `viewer.space.open_to_first_paint_us`
- `viewer.space.scan.total_us`, `.enumerate_us`, `.active_workers`, `.frontier_depth`, `.frontier_depth_peak`, `.idle_worker_waits`, `.files`, `.folders`, `.bytes`
- `viewer.space.queue.posted_count`, `.coalesced_count`, `.pending_bytes`, `.drain_us`, `.lock_hold_us`
- `viewer.space.model.directory_count`, `.file_candidate_count`, `.synthetic_count`, `.cache_snapshot_bytes`, `.cache_skipped_large`
- `viewer.space.layout.rebuild_us`, `.draw_items`, `.visible_tiles`, `.culled_tiles`, `.candidate_cache_hits`, `.candidate_cache_misses`
- `viewer.space.render.paint_us`, `.static_cache_hit`, `.static_cache_miss`, `.static_cache_record_us`, `.static_cache_replay_us`, `.tile_draw_count`, `.text_draw_count`
- `viewer.space.hit_test_us`, `.hit_grid.cells`, `.hit_grid.max_candidates`

Debug builds with `ENABLE_TESTS` expose `WndMsg::kViewerSpaceDebugGetPerfSnapshot` for deterministic command selftests. The snapshot reports renderer mode, scan state, retained directory/file/synthetic counts, root/file-candidate/synthetic byte totals, the effective adaptive file-candidate budget used for the live scan and scan-cache key, pending queue count/bytes, draw/visible/cull counts, last tile/text draw counts, renderer create/resize/brush/text-format counters, swap-chain pixel dimensions, last renderer failure stage/HRESULT, layout generation, static render-cache counters/bytes, scan-cache snapshot bytes, layout candidate-cache counters, hit-grid cell/candidate counts, recent paint/layout/drain/hit-test timings, and a working-set sample. Retained synthetic counts/bytes are scan-model "Other" buckets only; transient layout-only "Other" buckets are excluded from those model totals. `WndMsg::kViewerSpaceDebugCompareHitTesting` samples visible tile centers and compares spatial-grid results with the reverse-linear fallback. `WndMsg::kViewerSpaceDebugForceRendererFault` is a test-only fault-injection hook for forcing the next render through the device-loss recovery path.

## Threading and Cancellation

- Folder scanning runs on background threads with cooperative cancellation (`std::stop_token`); `scanThreads` controls the in-process parallelism for a single ViewerSpace scan.
- Concurrency model: workers pull directories from a bounded shared frontier. Each enumerated directory appends discovered child directories back to the frontier and completion state propagates child byte totals to parents so final totals remain exact.
- UI updates are queued and periodically drained to avoid message storms. Repeated size/progress records are coalesced before drain when a newer record supersedes an older pending one. The UI thread swaps a bounded batch out of the producer queue under one short lock, processes outside the lock, and applies measured backpressure from pending bytes rather than raw queue count. Future queue representation changes must preserve the same generation checks, cancellation behavior, and `viewer.space.queue.*` metrics.
- Update draining and layout rebuild run on the viewer timer with a small time budget to keep `WM_PAINT` mostly render-only (responsive move/resize while scanning).
- When multiple ViewerSpace windows scan the same volume on the Win32 filesystem (`shortId == "file"`), scans are throttled via a per-volume concurrency limit (`maxConcurrentScansPerVolume`) regardless of `scanThreads`. For non-Win32 filesystems, throttling is per-filesystem-instance to avoid unrelated viewers blocking each other.
- A short-lived in-memory scan cache may be used to reuse recent results for the same root (configurable). To avoid collisions across mounts, ViewerSpace currently only uses the cache for the Win32 filesystem (`shortId == "file"`). Large snapshots are skipped when the estimated snapshot exceeds the scan-cache byte cap and must emit `viewer.space.model.cache_skipped_large`.
- Scans are canceled when:
  - viewer closes
  - user refreshes
  - a new scan starts
  - user presses **Cancel** during scanning
  - user presses `Esc` while scanning

## Traversal Rules

- Uses `IFileSystem::ReadDirectoryInfo()` to enumerate entries and compute sizes (works for any active filesystem plugin, not only Win32 paths).
- Skips reparse points to avoid cycles.
- Access-denied paths are treated as zero-sized and do not abort the scan.
- Child path construction is string-based: `childPath = parentPath + separator + entryName` (no `std::filesystem` normalization). Separator selection prefers the style already present in the root path; otherwise it defaults to `\` for Win32 paths and `/` for non-Win32 paths.
- To control memory usage, ViewerSpace stores directories as real nodes and stores retained files in a compact file-record arena with packed item ids. `topFilesPerDirectory` acts as a compatibility floor while the adaptive candidate policy derives a larger budget from the current treemap pixel area, up to the 20k retained-file ceiling per directory. Files beyond the retained candidate budget are grouped into an exact “Other” bucket so totals remain correct.
- Treemap item ids distinguish real directory nodes, compact file records, and synthetic “Other” buckets. Tooltip, path building, context menu dispatch, drill-down, and hit testing MUST resolve through that item identity instead of assuming every tile is a `Node`.
- Layout-only “Other” buckets use deterministic synthetic ids derived from the parent. Double-clicking a layout “Other” bucket increases that parent’s layout detail budget and, if needed, navigates to the parent so the user can explore the dense group. Scan-retained “Other” buckets preserve exact folded file totals and remain eligible for focused rescan/higher-detail behavior when the retained candidate floor policy allows it.

## Layout and Hit Testing

- Layout item admission is pixel-driven with a high measured safety ceiling (currently 50k draw items; 20k while actively scanning). The legacy 600/2600 caps are no longer the detail ceiling.
- Candidate selection may use a per-directory layout candidate cache keyed by child count, total bytes, scan state, and max item budget. Cache hits must reconstruct any synthetic “Other” bucket for the current layout pass and must be invalidated on model, view-root, size, DPI, theme, or relevant configuration changes.
- Layout scratch vectors are reused by depth through `LayoutWorkspace`; rebuilds keep stable last-rect slots keyed by directory id, file-record id, or deterministic synthetic id so animations do not require a per-rebuild previous-rect map.
- Dense layouts build a spatial hit grid over the treemap. Hover/click hit testing checks the relevant cell first and keeps a reverse-linear fallback for small or ungridded layouts. Debug tests compare sampled grid results with the linear fallback.
