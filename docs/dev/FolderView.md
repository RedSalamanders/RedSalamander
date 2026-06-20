# FolderView Architecture (Developer)

`FolderView` is the per-pane Direct2D file list: it enumerates a folder on a background MTA thread into a zero-copy item model, renders a virtualized column grid on the UI thread, and loads icons and thumbnails asynchronously. This page documents the internal contracts that are easy to break: the enumeration handoff and its staleness fence, arena lifetime, deferred DirectX init with GDI fallback, the column-layout / scroll-stop rules, and the icon-vs-thumbnail pipeline. It goes deeper than the FolderView section of [../DeveloperGuide.md](../DeveloperGuide.md); see [../MainWindow.md](../MainWindow.md) for the user-facing pane behavior and `Specs/UI/UI_FolderView.md` for the normative spec.

## Source layout

The class is split across translation units that share one private header, `RedSalamander/FolderViewInternal.h`. The members described here live in `RedSalamander/FolderView.h`.

| File | Responsibility |
|---|---|
| `FolderView.cpp` | Window lifecycle, `WndProc` message dispatch, `OnPaint`, `OnDeferredInit` |
| `FolderView.Enumeration.cpp` | `EnumerationWorker` / `ExecuteEnumeration`, sort, cache refresh, `ProcessEnumerationResult` |
| `FolderView.Rendering.cpp` | D3D/D2D/DWrite resources, `Render`, `DrawItem` |
| `FolderView.Layout.cpp` | `LayoutItems`, `GetVisibleItemRange`, `HitTest`, `EnsureVisible`, scroll metrics |
| `FolderView.Interaction.cpp` | Mouse/keyboard/scroll, horizontal scroll-stop resolution, incremental search |
| `FolderView.Icons.cpp` | Async icon + thumbnail loading and UI-thread bitmap creation |
| `FolderViewColumnLayout.h` | Pure column-layout solver + scroll-stop math (header-only) |

## Enumeration: MTA jthread, generation fence, message handoff

### One worker thread, two job types

`FolderView` owns a single `std::jthread _enumerationThread` started lazily by `EnsureEnumerationThread()` and torn down by `StopEnumerationThread()`. The worker loop `EnumerationWorker(std::stop_token)` initializes COM as **MTA** (`CoInitializeEx(nullptr, COINIT_MULTITHREADED)`); this is required because icon extraction (`IconCache::ExtractSystemIcon` via `IImageList::GetIcon`) and WIC thumbnail decoding both call COM off the UI thread. The UI thread itself uses STA (`COINIT_APARTMENTTHREADED` in `OnCreate`).

The same worker services three kinds of work, gated by a condition variable `_enumerationCv` predicate over: a pending folder path, `_iconLoadingActive`, and `_thumbnailLoadingActive`. So one thread does folder enumeration, icon extraction, and thumbnail extraction in turn.

### The staleness fence: `_enumerationGeneration`

Every navigation or refresh bumps a monotonic counter:

```cpp
std::atomic<uint64_t> _enumerationGeneration{0};
```

`EnumerateFolder()` (navigation) and `RequestRefreshFromCache()` (cache-driven refresh) both do `_enumerationGeneration.fetch_add(1, ...) + 1`, stash the new generation alongside the path under `_enumerationMutex` (`_pendingEnumerationPath` / `_pendingEnumerationGeneration`), and notify the CV. `CancelPendingEnumeration()` also bumps the counter (without queueing new work) so any in-flight payload becomes stale.

`ExecuteEnumeration(folder, generation, stopToken)` re-checks `_enumerationGeneration.load() != generation` at every expensive boundary — the entry-buffer walk, the sort/merge, and each parallel icon-index query phase — and returns `nullptr` the instant the generation moves. This is the key data-safety property: a navigation that supersedes an in-flight enumeration causes the worker to abandon it rather than post a stale item list.

### Handoff via `kFolderViewEnumerateComplete`

When the worker finishes and the generation still matches, it hands the result to the UI thread by posting a window message with an owned payload:

```cpp
PostMessagePayload(_hWnd.get(), WndMsg::kFolderViewEnumerateComplete, 0, std::move(payload));
```

The `WndProc` case takes the payload back with `TakeMessagePayload<EnumerationPayload>(lParam)` and calls `ProcessEnumerationResult`. That handler re-checks the fence a **third** time on the UI thread: if `payload->generation != _enumerationGeneration.load()` the payload is dropped. This double/triple check matters because messages can sit in the queue across a navigation. `DrainPendingEnumerationPayloadMessages()` exists to flush any queued `kFolderViewEnumerateComplete` payloads (e.g. before teardown) without applying them.

| Step | Thread | Fence check |
|---|---|---|
| `EnumerateFolder` / `RequestRefreshFromCache` bump generation, queue path | UI | writes new generation |
| `ExecuteEnumeration` builds items, queries icon indices | worker (MTA) | re-checks at each loop / phase; returns `nullptr` if stale |
| Worker posts `kFolderViewEnumerateComplete` | worker | only if `generation == _enumerationGeneration` |
| `ProcessEnumerationResult` applies payload | UI | drops payload if `generation` mismatched |

`ScheduleBusyOverlay(generation, folder)` is armed when work is queued and a >300 ms busy overlay (with a Cancel action) is shown if the enumeration is slow; `CancelBusyOverlay(generation)` clears it when the matching payload arrives.

## Zero-copy `FolderItem` and arena lifetime

`ExecuteEnumeration` calls `IFileSystem` through `DirectoryInfoCache::BorrowDirectoryInfo(...)`, which yields an `IFilesInformation` whose internal arena owns a packed `FileInfo` buffer. Each `FolderItem` stores its name as a `std::wstring_view` pointing **into that arena** — no per-item string copy:

```cpp
struct FolderItem
{
    std::wstring_view displayName; // View into FileInfo::FileName in the arena
    uint16_t extensionOffset = 0;  // Offset of '.' within displayName
    uint32_t stableHash32    = 0;
    // ... attributes, bounds, icon/thumbnail bitmaps, text layouts ...
};
```

`GetExtension()` / `GetNameWithoutExtension()` are computed by slicing `displayName` at `extensionOffset` — still zero-copy. The buffer is traversed via `FileInfo::NextEntryOffset` with explicit bounds validation (`next < base || next + sizeof(FileInfo) > end` → `ERROR_INVALID_DATA`).

### Lifetime rule

The arena must outlive every `string_view` into it. The payload carries a COM reference that keeps it alive:

```cpp
struct EnumerationPayload
{
    uint64_t generation = 0;
    HRESULT status      = S_OK;
    std::vector<FolderItem> items;
    wil::com_ptr<IFilesInformation> arenaBuffer; // keeps the arena alive
    std::filesystem::path folder;                // recompute full paths on demand
};
```

`ProcessEnumerationResult` moves both the items and the arena into the live members in lockstep:

```cpp
_items            = std::move(payload->items);
_itemsArenaBuffer = std::move(payload->arenaBuffer); // keep arena alive for the views
_itemsFolder      = std::move(payload->folder);
```

**Never let a `FolderItem::displayName` (or any view derived from it) outlive `_itemsArenaBuffer`.** Full paths are reconstructed on demand from `_itemsFolder` + `displayName` (`GetItemFullPath`), so items intentionally do not store paths. The stable per-item hash is seeded from the folder path so rainbow rendering stays stable without keeping a path string per row.

Sorting in the worker uses `std::execution::par` only past a threshold (1000 items); directories sort before files. On the UI thread, `ApplyCurrentSort` re-sorts in place using the user's sort key/direction and preserves selection and focus **by display name** (refresh transfers selection across rename-chain hints in `_pendingRefreshSelectionRenames`). Refresh of the same folder also transfers cached layouts, details/metadata text, and the icon bitmap when the icon index is unchanged, so unchanged rows do not re-measure or re-decode.

## Deferred DirectX init, GDI fallback, and icon re-queue

To keep `WM_CREATE` fast, FolderView does **not** create its D3D device, swap chain, or D2D target during creation. `OnCreate` only initializes COM/OLE and registers the drop target. DirectX comes up after the first paint.

### First paint falls back to GDI

`OnPaint` checks whether the D2D pipeline is ready (`_d2dContext && (_swapChain || _swapChainLegacy) && _d2dTarget`). If not, it fills the client with a GDI brush (the menu background brush, or a solid brush from the theme background color) and posts `kFolderViewDeferredInit` exactly once, guarded by `_deferredInitPosted`:

```cpp
FillRect(paint_dc.get(), &rcPaint, fillBrush);
if (! _deferredInitPosted && _hWnd && _clientSize.cx > 0 && _clientSize.cy > 0)
{
    _deferredInitPosted = PostMessageW(_hWnd.get(), WndMsg::kFolderViewDeferredInit, 0, 0) != 0;
}
```

### `OnDeferredInit` brings up the pipeline

`OnDeferredInit` runs `EnsureDeviceIndependentResources()`, `EnsureDeviceResources()`, and `EnsureSwapChain()`, then clears `_deferredInitPosted`. It computes a diagnostic `missingMask` (bit `0x01` zero size, `0x02` no D2D context, `0x04` no swap chain, `0x08` no D2D target, `0x10` resize pending) before and after, emitted as perf values. If the pipeline still is not ready (`missingAfter != 0`, common at 0×0 size or during an active resize) it returns without invalidating, avoiding a repaint loop; the next real paint re-posts.

### Icon work must not block on swap-chain resize

A subtle but important rule (see `FolderViewInternal` comments and the spec): icon extraction/conversion needs only the **D2D device context**, not a completed swap-chain resize. `OnDeferredInit` therefore initializes `IconCache` and re-runs `QueueIconLoading()` (and `QueueThumbnailLoading()` when in thumbnail mode) as soon as `_d2dContext` exists, *before* the `missingAfter` early-return:

```cpp
if (_d2dContext)
{
    IconCache::GetInstance().Initialize(_d2dContext.get(), _dpi);
    QueueIconLoading();
    if (_thumbnailsVisible) { QueueThumbnailLoading(); }
}
if (missingAfter != 0) { return; }
```

This re-queue is necessary because enumeration can complete before DirectX is ready: `QueueIconLoading()` no-ops when `_d2dContext` is null. Without the re-queue from `OnDeferredInit`, fast startup would leave enumerated rows stuck on placeholder icons. (Pending swap-chain resizes are instead handled at the top of `OnPaint` via `TryResizeSwapChain` + `EnsureSwapChain`, so a resize never gates icon conversion.)

## Column layout and the scroll-stop contract

FolderView lays items out in **vertical columns, scrolled horizontally only** (the vertical scrollbar is always hidden). The pure solver lives in `FolderViewColumnLayout.h` and is driven by `LayoutItems` in `FolderView.Layout.cpp`.

### Variable-width columns

`FolderViewColumnLayout::Resolve` first computes `rowsPerColumn` from the client height and row stride, then walks items top-to-bottom assigning `rowsPerColumn` items per column. Each column's `widthDip` is sized to the **widest text in that column only** (label width, plus details/metadata widths in the richer display modes), clamped to `[minColumnWidth, clientWidth]`. Columns are then placed left to right with `kColumnSpacingDip` (18 DIP) gutters; the first column starts at `leftDip = columnSpacing`, so there is a leading gutter before column 0.

Because widths are per-column, all position-dependent paths **must read `_columnLayout[c].leftDip` / `.widthDip`** and must not assume a single global column stride. The data structures built by `LayoutItems`:

| Member | Purpose |
|---|---|
| `_columnLayout` | `vector<Column>` — variable `leftDip`/`widthDip` per column |
| `_columnCounts` | item count per column |
| `_columnPrefixSums` | prefix sums (`[c]` = items before column c, plus a sentinel) for O(1) column/row ↔ index conversion in keyboard nav |
| `_contentWidth` | `max(layout content width, max item right + gutter, client width)` |

`GetVisibleItemRange()`, `HitTest()`, `EnsureVisible()`, rendering, and horizontal scrolling all consume `_columnLayout` directly. `EnsureVisible` scrolls so the focused item's column edge (`column.leftDip`) is in view, clamped to `[0, _contentWidth - viewWidth]`.

### Scroll stops

Horizontal offset `0` is the canonical first-column stop and preserves the leading gutter. Line/page/wheel scrolling snaps to a column edge via three header functions, all of which iterate `columns[1..]` (column 0's stop is implicitly `0`):

| Function | Used by |
|---|---|
| `ResolveNextScrollStop(offset, max, columns)` | `SB_LINERIGHT`, page-right, wheel-right |
| `ResolvePreviousScrollStop(offset, max, columns)` | `SB_LINELEFT`, page-left, wheel-left |
| `ResolveNearestScrollStop(target, max, columns)` | `SB_THUMBPOSITION` (thumb release) |

Contract consequences encoded by this math:

- The **first right line-scroll from `0`** moves to the second column's `leftDip`, not back to the first column's gutter; the matching first left line-scroll from the second column returns to `0` in one step (a `0.5` DIP tolerance absorbs float noise).
- Thumb tracking may follow raw pixel offsets while dragging, but **release snaps** to the nearest valid stop (`ResolveNearestScrollStop`).
- Hit testing treats the leading gutter before column 0 as part of column 0 (`ResolveHitColumnIndex` clamps column 0's left to `min(0, leftDip)`) so the first visible strip stays clickable; gaps between later columns are dead hit-test space.

## Async icon vs thumbnail pipeline

Both pipelines are producer/consumer: the UI thread builds a bounded request queue, the MTA worker extracts off-thread, and a posted message creates the actual Direct2D bitmap back on the UI thread. Both carry `enumerationGeneration` so stale work is dropped, and both create D2D bitmaps only on the UI thread.

### Icons — grouped by system icon index

Icon indices are resolved during enumeration (extension batch query + per-file lookup for `.exe`/`.dll`/`.ico`/`.lnk`/`.url` and special folders). `QueueIconLoading()`:

1. No-ops if `_d2dContext` is null (re-queued later from `OnDeferredInit`).
2. Groups items by `iconIndex` so each icon is converted **once** and applied to all matching rows.
3. Immediately stamps any bitmap already cached for the current D2D device (no background work).
4. Orders **visible** groups first (sorted by first visible item), offscreen groups after, into `_iconLoadQueue`, then sets `_iconLoadingActive` and notifies the CV.

`ProcessIconLoadQueue()` (worker) pops groups, calls `IconCache::ExtractSystemIcon` (with a bounded retry on failure — visible groups re-queue to the front), and posts `kFolderViewCreateIconBitmap` with an owned `IconBitmapRequest`. `OnCreateIconBitmap` (UI thread) drops the request if `iconLoadBatchId` or `enumerationGeneration` is stale, converts the `HICON` via `ConvertIconToBitmapOnUIThread`, and applies the bitmap to every still-matching item (`item.iconIndex == request.iconIndex && ! item.icon`). `BoostIconLoadingForVisibleRange()` re-prioritizes the queue toward the current viewport on scroll. The worker reads the live `_d2dDevice` under `_d2dDeviceMutex` to test the per-device cache.

### Thumbnails — bounded to visible work, exclusive display mode

Thumbnails is an exclusive per-pane display mode; normal icons remain the fallback. `QueueThumbnailLoading()` queues **only visible items** (capped at `kMaxThumbnailQueueItems = 256`), evicts offscreen thumbnails when the cache exceeds `kMaxThumbnailCacheBytes` (64 MiB) — never evicting currently visible thumbnails — and bumps a `batchId` so a new pass invalidates the old one. `ProcessThumbnailLoadQueue()` (worker):

- Tries the shell thumbnail first via `IShellItemImageFactory::GetImage`.
- On shell failure for a likely-image extension (`HasLikelyWicImageExtension`), decodes a bounded first frame through **WIC**, scaling with `IWICBitmapScaler` and converting to `GUID_WICPixelFormat32bppPBGRA`. The WIC factory is created **once per worker** via `EnsureThumbnailWicFactory` and reused (the worker is MTA, as the factory requires).
- Drops a request whose `enumerationGeneration` or `thumbnailLoadBatchId` no longer matches (`staleDrops`), retries once on hard extraction failure, and otherwise posts `kFolderViewCreateThumbnailBitmap`.

`OnCreateThumbnailBitmap` (UI thread) creates the D2D bitmap from the shell `HBITMAP` or the WIC pixel buffer. Thumbnails preserve source aspect ratio inside the slot (letterbox/pillarbox); when neither shell nor WIC yields a thumbnail, the normal icon is drawn in the larger slot and `thumbnailFallbackResolved`/`thumbnailFallbackTargetPx` record the fallback for the current target size. `CancelThumbnailLoading()` bumps the batch id and clears the queue; navigation, refresh, sort, resize, horizontal scroll, ensure-visible, thumbnail-size changes, and leaving thumbnail mode all cancel or re-queue so stale payloads are never applied and newly visible columns do not stay on icon fallback.

### Message summary

| Message | Posted by | Handled by | Payload |
|---|---|---|---|
| `kFolderViewEnumerateComplete` | enumeration worker | `ProcessEnumerationResult` | `EnumerationPayload` (items + arena) |
| `kFolderViewDeferredInit` | `OnPaint` (GDI fallback) | `OnDeferredInit` | none |
| `kFolderViewCreateIconBitmap` | `ProcessIconLoadQueue` | `OnCreateIconBitmap` | `IconBitmapRequest` (HICON + item indices) |
| `kFolderViewCreateThumbnailBitmap` | `ProcessThumbnailLoadQueue` | `OnCreateThumbnailBitmap` | `ThumbnailBitmapRequest` (HBITMAP or BGRA pixels) |

## Invariants checklist

- Every cross-thread result carries `enumerationGeneration`; check it on both ends before applying.
- `FolderItem::displayName` is valid only while `_itemsArenaBuffer` lives — move the items and the arena together, never separately.
- The enumeration worker must be COM-MTA; the UI thread is COM-STA.
- D2D bitmaps (icons, thumbnails) are created only on the UI thread, from posted payloads.
- `QueueIconLoading` no-ops without `_d2dContext`; `OnDeferredInit` re-runs it once DirectX is up, and icon conversion must not wait on swap-chain resize.
- Position math (render, hit test, `EnsureVisible`, horizontal scroll) reads `_columnLayout`; never assume a global column stride. Horizontal scroll snaps to a column stop on release.
- Focus and selection are independent; refresh preserves both by display name (with rename-chain hints).
