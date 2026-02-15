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
- `topFilesPerDirectory` (96): number of largest files tracked per directory; remaining files are grouped into “Other” (`0` groups all files).
- `scanThreads` (1): number of background threads used to scan sibling subtrees in parallel within a single ViewerSpace scan.
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
- Auto-expansion is driven by **relative area** (fraction of the visible treemap): it keeps expanding “dominant” rectangles until they fall below a target threshold (roughly **10%** while scanning/idle, with a hard depth cap).

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
- Up stays available during scanning; if the current scan root has no parent in the current model, Up restarts a scan at the parent path. Up is disabled when the scan root is already at a volume/share root.
- While scanning, click header **Cancel** (or press `Esc`) to stop the scan.
- Double-clicking an aggregated “Other” bucket triggers a focused action to explore that part (typically drilling into the parent directory or rescanning it with a higher `topFilesPerDirectory`).
- Hover: show a tooltip with name/path, size, and share for the tile under the cursor; scan state is shown only while in-progress (not shown for `Done`); tooltip tracks the cursor within the tile (no “stuck” tooltip position).
- Keyboard:
  - `Backspace` / `Alt+Up`: drill up
  - `F5`: refresh/rescan current node (forces a full rescan; bypasses cache)
  - `Esc`: cancel scan if scanning; otherwise close viewer
- The window menu bar (`IDR_VIEWERSPACE_MENU`) is owner-drawn and themed like host menus (dark/high-contrast aware; rainbow selection when `rainbowMode` is enabled).

### Labels
- Item name + compact size label (when there is enough space).
- Text is ellipsized via DirectWrite trimming (avoids mid-glyph clipping without per-tile layouts).
- Aggregated “Other” buckets show item counts (and folder/file breakdown) plus size.
- Folder tiles reserve a header strip for the name (when there is enough height); file tiles use a small “dog-ear” corner fold that reveals the parent tile color behind it, and a matching cut-corner outline to reinforce the fold (stronger file/folder distinction).
- When a tile is large enough to show a line of text, ViewerSpace keeps at least one readable name line visible (prefer showing the beginning of the name).

## Rendering Requirements

- Direct2D 1.1 + DirectWrite
- Hwnd render target; handle `D2DERR_RECREATE_TARGET`
- Use `ViewerTheme.backgroundArgb`, `textArgb`, and `accentArgb` for styling; in high-contrast, prioritize strong borders/readability.
- First-show background MUST match the current theme (avoid a white/high-contrast flash before the first Direct2D paint); implement via a themed `WNDCLASSEXW::hbrBackground` + `WM_ERASEBKGND` behavior consistent with `Specs/PluginsViewer.md`.

## Threading and Cancellation

- Folder scanning runs on background threads with cooperative cancellation (`std::stop_token`); `scanThreads` controls the in-process parallelism for a single ViewerSpace scan.
- Concurrency model: enumerate the scan root, then scan its immediate child directories in parallel (each worker scans a subtree depth-first).
- UI updates are batched (queue + periodic drain) to avoid message storms.
- Update draining and layout rebuild run on the viewer timer with a small time budget to keep `WM_PAINT` mostly render-only (responsive move/resize while scanning).
- When multiple ViewerSpace windows scan the same volume on the Win32 filesystem (`shortId == "file"`), scans are throttled via a per-volume concurrency limit (`maxConcurrentScansPerVolume`) regardless of `scanThreads`. For non-Win32 filesystems, throttling is per-filesystem-instance to avoid unrelated viewers blocking each other.
- A short-lived in-memory scan cache may be used to reuse recent results for the same root (configurable). To avoid collisions across mounts, ViewerSpace currently only uses the cache for the Win32 filesystem (`shortId == "file"`).
- Scans are canceled when:
  - viewer closes
  - user refreshes
  - a new scan starts
  - user presses **Cancel** / `Esc` during scanning

## Traversal Rules

- Uses `IFileSystem::ReadDirectoryInfo()` to enumerate entries and compute sizes (works for any active filesystem plugin, not only Win32 paths).
- Skips reparse points to avoid cycles.
- Access-denied paths are treated as zero-sized and do not abort the scan.
- Child path construction is string-based: `childPath = parentPath + separator + entryName` (no `std::filesystem` normalization). Separator selection prefers the style already present in the root path; otherwise it defaults to `\` for Win32 paths and `/` for non-Win32 paths.
- To control memory usage, ViewerSpace stores all directories but only keeps the top-K files per directory; the remainder is grouped into an “Other” bucket.
