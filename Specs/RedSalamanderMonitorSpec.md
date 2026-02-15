# RedSalamanderMonitor Overview and Improvement Plan

## Architecture: Append-Only Log Viewer

**RedSalamanderMonitor is a HIGH-THROUGHPUT, APPEND-ONLY log viewer optimized for real-time log streaming.**

### Core Constraints
- **Lines are NEVER modified** - only appended at the end
- **High throughput**: 10,000+ logs/second sustained
- **Auto-scroll (tail mode) is the primary use case** - viewing latest logs in real-time
- **Scroll-back is secondary** - reviewing historical logs when paused
- **Read-only display** - no editing operations

## Localization and Menu Resources

- The main menu and all static UI strings are defined in `.rc` resources (`RedSalamanderMonitor/RedSalamanderMonitor.rc`) for localization.
- Runtime code may only populate truly dynamic menu content (e.g., themes discovered from `Themes\\*.theme.json5` and custom themes from settings); see `Specs/LocalizationSpec.md`.

### Two-Mode Rendering Architecture

**Mode 1: AUTO-SCROLL (Hot Path - 99% of usage)**
- Optimized for blazing-fast append and display of new lines
- Maintains a **dynamic tail layout** of the last N lines where N = `max(100, visibleLines + 50)`
  - Adapts to viewport size and DPI
  - Ensures tail window always covers visible area plus margin
  - Prevents blank areas on large/high-DPI displays
- Direct rendering without complex virtualization or caching
- Synchronous layout updates for immediate visibility
- Target: <0.5ms append latency, zero dropped frames at 10k logs/sec

**Mode 2: SCROLL-BACK (Cold Path - occasional use)**
- Full virtualization with slice-based rendering when user scrolls up
- Slice size: 256 lines per block (`kSliceBlockLines`)
- Complex coverage checks and fallback layouts for historical review
- Offscreen bitmap caching for smooth scrolling (with 16384px D3D11 texture limit)
- Fallback to direct rendering when slice exceeds texture dimension limits
- Acceptable latency for non-real-time viewing

### Mode Transitions and Auto-Scroll Behavior

**Rendering Mode Selection**:
- Mode is determined by `ColorTextView::_renderMode` (enum `RenderMode`)
- AUTO-SCROLL mode when `_renderMode == RenderMode::AUTO_SCROLL`
- SCROLL-BACK mode when `_renderMode == RenderMode::SCROLL_BACK`

**Auto-Scroll State Management** (Properly Encapsulated):

ColorTextView owns and manages auto-scroll state through:
- `SetAutoScroll(bool)`: Switch render mode (AUTO_SCROLL vs SCROLL_BACK)
- `GetAutoScroll()`: Query whether AUTO_SCROLL mode is active
- Internal helper: `ShouldUseAutoScrollMode()` is a thin wrapper over `_renderMode == RenderMode::AUTO_SCROLL`

When **Auto-Scroll ON** (`_renderMode == RenderMode::AUTO_SCROLL`):
- New lines arrive → automatically scroll to show them (stay at bottom)
- User scrolls UP (wheel/scrollbar) → switches to **SCROLL-BACK** mode (`SwitchToScrollBackMode()`)
- User jumps to END (End key/SB_BOTTOM) → switches/keeps **AUTO-SCROLL** mode (`SwitchToAutoScrollMode()`)

When **Auto-Scroll OFF** (`_renderMode == RenderMode::SCROLL_BACK`):
- New lines arrive → **don't change display** (stay at current position)
- User scrolls anywhere → nothing changes (already off)
- User jumps to END → switches to **AUTO-SCROLL** mode (`SwitchToAutoScrollMode()`)

**Menu Synchronization**:
- ColorTextView manages its own auto-scroll state (no global variables)
- Main window calls `GetAutoScroll()` in `UpdateStatusBar()` and syncs menu checkmark
- Main window toggles via `SetAutoScroll()` when user clicks menu
- Clean API separation: ColorTextView = reusable component, main window = UI consumer

**User Actions**:
- Manual menu toggle: Main window calls `SetAutoScroll(!GetAutoScroll())` and updates menu
- Scroll up: ColorTextView switches to **SCROLL-BACK** mode internally (wheel/scrollbar handlers call `SwitchToScrollBackMode()` when scrolling up)
- Jump to end: ColorTextView switches to **AUTO-SCROLL** mode internally (End / SB_BOTTOM paths call `SwitchToAutoScrollMode()`)
- Main window queries state via `GetAutoScroll()` for menu synchronization

### Display Row Mapping System - VisibleLine Architecture

**Critical for multi-line content rendering with filtering support:**

The system uses a **VisibleLine Index Pattern** instead of sentinel values for efficient filtered line access:

1. **`newlineCount` field**: Tracks embedded newlines per line (0 for single-line content)
2. **VisibleLine struct**: Lightweight 12-byte mapping structure
   ```cpp
   struct VisibleLine {
       size_t sourceIndex;      // Index into source lines vector
       UINT32 displayRowStart;  // First display row for this visible line
   };
   ```
3. **`visibleLines` vector**: Computed view maintained alongside source `lines` vector
   - Contains only visible (non-filtered) lines
   - Sorted by sourceIndex (always ascending)
   - Updated incrementally during append operations
   - Rebuilt fully on filter mask changes
4. **Mapping functions**:
   - `displayRowForVisible(visibleIndex)` → Returns display row for visible line by visible index
   - `displayRowForSource(sourceIndex)` → Returns display row for source line by searching visibleLines
   - `visibleIndexFromDisplayRow(displayRow)` → Binary search to find visible index from display row

**Y Position Calculations:**
All rendering Y positions MUST use `displayRowForVisible()` or `displayRowForSource()` instead of logical line numbers:
```cpp
// CORRECT (using visible index):
const UINT32 displayRow = _document.displayRowForVisible(visibleIdx);
const float yBase = displayRow * lineHeight;

// CORRECT (using source index):
const UINT32 displayRow = _document.displayRowForSource(sourceIdx);
const float yBase = displayRow * lineHeight;

// WRONG (causes blank lines bug):
const float yBase = logicalLine * lineHeight;  // Ignores multi-line content AND filtering!
```

This applies to:
- Text layout rendering (`DrawScene`)
- Slice bitmap positioning (`_sliceBitmapYBase`)
- Tail layout rendering (AUTO_SCROLL mode)
- Fallback layout rendering
- Line number rendering (`DrawLineNumbers`)

**Filtering Integration:**
When filtering is active, the VisibleLine architecture provides:
- **Efficient access**: O(1) access to visible lines by visible index
- **Source mapping**: O(log n) binary search to map display rows to source indices
- **Incremental updates**: Single visible line added during append (O(1))
- **Full rebuild**: O(n) rebuild when filter mask changes
- **No sentinels**: Clean separation between visible and source line spaces

## Current behavior
- Windows monitor that receives log lines over **ETW (Event Tracing for Windows)** (see `Common/Helpers.h`): structured `Debug::InfoParam` records (time, process id, thread id, type Text/Error/Warning/Info/Debug).
- **ETW message intake**:
  - **`WM_APP_ETW_BATCH`**: ETW events queued and processed in batches from the EtwListener worker thread
  - Batch processing defers invalidation by calling `AppendInfoLine(..., deferInvalidation=true)` for each entry, then performing a single mode-specific update at the end of the batch.
  - Note: `BeginBatchAppend()`/`EndBatchAppend()` exists but the current `WM_APP_ETW_BATCH` path does not call them directly.
- Single main window with menu and toolbar: New/Open/Save As, Copy, toggle toolbar, toggle line numbers, show/hide IDs, auto-scroll, always-on-top; debug builds can start a random message generator.
- Display surface is `ColorTextView` rendered with Direct2D/DirectWrite on a D3D11/DXGI swap chain; per-monitor DPI aware; supports line numbers, colored metadata prefixes, keyword colorization for Error/Warning/Debug, and optional auto-scroll.
- Input pipeline normalizes CR/LF, appends text to the document, caches prefixes per line, and updates gutter width; selection and clipboard copy are supported (Ctrl+C, Ctrl+A); find bar via Ctrl+F with F3 navigation; mouse wheel scroll with Shift for horizontal.
- File open/save uses BOM detection (UTF-8/UTF-16LE), registry-backed config for email, and basic about dialog; device-loss and DPI-change handling recreate swap chain/targets.

## Existing performance/architecture traits
- **Two-mode rendering system**: AUTO-SCROLL mode uses dynamic tail layout (viewport-sized, min 100 lines) with direct rendering for <0.5ms append latency; SCROLL-BACK mode uses full virtualization with slice-based rendering (`kSliceBlockLines=256`) for historical review.
- **Display row mapping**: All Y position calculations use display-row offsets (`Document::DisplayRowForVisible()` / `Document::DisplayRowForSource()`) to correctly handle multi-line content with embedded newlines.
- **Batched message intake**: ETW events are processed in batches via `WM_APP_ETW_BATCH` to avoid per-message overhead at high throughput.
- **D3D11 texture limits**: Validates slice bitmap dimensions against 16384px limit, falls back to direct rendering when exceeded.
- Layout and width measurements in SCROLL-BACK mode run on a thread pool with slice prefetch; layouts are cached to skip reflow when the same slice is requested.
- AUTO-SCROLL mode uses synchronous layout updates and direct-to-backbuffer rendering; SCROLL-BACK mode uses offscreen slice bitmap when possible (within texture limits), otherwise direct rendering.
- Partial present is used for scroll deltas to minimize redraw.
- Line metadata (time/pid/tid/type) and brushes are cached; display-row offsets are precomputed for quick gutter/hit-testing.
- Mode transitions are explicit: menu toggle calls `SetAutoScroll()`, wheel/scrollbar scrolling up switches to SCROLL-BACK, and End/SB_BOTTOM switches to AUTO-SCROLL.
- **Mode detection timing (current implementation)**: `WM_APP_ETW_BATCH` appends the whole batch first, then applies AUTO_SCROLL vs SCROLL_BACK invalidation/layout behavior once per batch.

## Known Issues and Fixes Applied

### Blank Lines Bug (FIXED)
**Root Cause**: Y positioning used logical line numbers instead of display row offsets, causing misalignment when lines contained embedded newlines.

**Symptoms**: 
- Text rendering at wrong Y coordinates after ~1000 lines
- Line numbers display but text content missing (blank white lines)
- Issue appeared at document boundaries where multi-line content accumulated

**Fix Applied**:
- All Y position calculations now use `displayRowForVisible()` or `displayRowForSource()`
- Updated: text layout rendering, slice bitmap positioning, tail layout, fallback layout
- Display row offset maintained incrementally during append

**Validation**: Tested with 5000+ messages including multi-line content, scrolling throughout document

### Filter Display Synchronization Bug (FIXED)
**Root Cause**: When filter mask changed, `visibleLines` vector was not rebuilt, causing stale computed view with incorrect display row mappings.

**Symptoms**:
- Line numbers not synchronized with displayed content after filter change
- Empty lines appearing when toggling filters
- Lines disappearing during scrolling with active filters
- Incorrect Y positioning of text after filter changes
- Crash when accessing empty visibleLines vector

**Fix Applied** (VisibleLine Architecture Migration):

1. **VisibleLine Infrastructure** (Phase 1):
   - Introduced `VisibleLine` struct with sourceIndex and displayRowStart fields
   - Implemented `rebuildVisibleLines()` for full O(n) reconstruction
   - Added core mapping methods: `displayRowForVisible()`, `displayRowForSource()`, `visibleIndexFromDisplayRow()`
   - Migrated from sentinel-based array to clean index pattern

2. **Complete API Migration** (Phases 2-3):
   - Replaced all obsolete method calls throughout codebase
   - `getLineAll()` → `getSourceLine()` (15 locations)
   - `lineFromDisplayRowAll()` → `visibleIndexFromDisplayRow()` (7 locations)
   - `displayRowForLine()` → `displayRowForSource()` (9 locations)
   - Removed obsolete methods and sentinel-based code

3. **Incremental visibleLines Maintenance** (Phase 3f):
   - `appendInfoLine()`: Incremental update when single line added (O(1))
   - `appendText()`: Full rebuild after bulk append
   - `setText()`: Clear and rebuild when replacing all text
   - `clear()`: Clear visibleLines when document cleared
   - `setFilterMask()`: Full rebuild via `rebuildVisibleLines()`

4. **Deadlock Fixes** (Phase 3e):
   - Split `isLineVisible()` into public (with lock) and `isLineVisibleUnsafe()` (no lock)
   - Use unsafe version within `appendInfoLine()` to avoid lock acquisition while holding unique_lock
   - Pattern: Methods ending in `Unsafe()` assume caller holds lock

5. **Bounds Checking** (Phase 3d):
   - Added empty visibleLines checks at all access points
   - Prevents crash when visibleLines not yet populated

**Synchronization Flow**:
```text
User toggles filter → SetFilterMask(newMask)
  ├─> Document::setFilterMask(newMask)
  │    ├─> _filterMask = newMask
  │    ├─> rebuildVisibleLines() [O(n) full rebuild]
  │    │    ├─> visibleLines.clear()
  │    │    ├─> Iterate all source lines
  │    │    └─> For visible: visibleLines.push_back({srcIdx, displayRow})
  │    └─> markAllDirtyUnsafe()
  ├─> _contentHeight = totalDisplayRows() * lineHeight
  ├─> ClampScroll() [prevent out-of-bounds]
  ├─> UpdateScrollBars() [update ranges]
  └─> Full invalidation + redraw
```

**Validation Requirements**:
- Toggle filters repeatedly while scrolled at various positions
- Verify no empty lines appear
- Verify line numbers stay synchronized with visible content
- Verify smooth scrolling without content disappearing
- Verify scrollbar ranges update immediately
- Test rapid filter changes (stress test)
- Test all filter combinations (single type, multiple types, all types)

## Public API Reference

### ColorTextView Core Methods

**Document Management:**
```cpp
void SetText(const std::wstring& text);           // Replace all content
void AppendText(const std::wstring& text);        // Append plain text
void AppendInfoLine(const std::wstring& text,     // Append structured log entry
                    const Debug::InfoParam& info);
void Clear();                                      // Clear all content
```

**Filtering:**
```cpp
void SetFilterMask(uint32_t mask);                // Set type filter (5-bit mask)
size_t GetVisibleLineCount() const;               // Count of visible lines
size_t GetTotalLineCount() const;                 // Count of all lines (filtered + unfiltered)
```

**Auto-Scroll:**
```cpp
void SetAutoScroll(bool enable);                  // Enable/disable auto-scroll
bool GetAutoScroll() const;                       // Query auto-scroll state
```

**Display Options:**
```cpp
void EnableLineNumbers(bool enable);              // Toggle line number gutter
void EnableShowIds(bool enable);                  // Toggle PID/TID display
void SetTheme(const Theme& theme);                // Set color theme
```

**Theme Requirements:**
- Must support **Light**, **Dark**, **Rainbow**, and **System High Contrast** themes.
- Theme changes must not clear or reflow content unnecessarily; apply via brush updates and invalidation.
- High Contrast mode must prefer Windows system colors and maximize readability.

**File Operations:**
```cpp
bool LoadFromFile(const std::wstring& path);      // Load text file (BOM detection)
bool SaveToFile(const std::wstring& path);        // Save visible content
```

### Document Class Methods

**Line Access:**
```cpp
const Line& GetVisibleLine(size_t visibleIndex);  // Access by visible index (0-based)
const Line& GetSourceLine(size_t sourceIndex);    // Access by source index (all lines)
size_t VisibleLineCount() const;                  // Count of visible lines
size_t TotalLineCount() const;                    // Count of all lines
```

**Display Row Mapping:**
```cpp
UINT32 DisplayRowForVisible(size_t visIdx);       // Map visible index → display row
UINT32 DisplayRowForSource(size_t srcIdx);        // Map source index → display row
size_t VisibleIndexFromDisplayRow(UINT32 row);    // Map display row → visible index
UINT32 TotalDisplayRows() const;                  // Total display rows (accounts for multi-line)
```

**Filtering:**
```cpp
void SetFilterMask(uint32_t mask);                // Update filter mask
bool IsLineVisible(size_t sourceIndex) const;     // Check if source line is visible
```

## Performance Metrics

### Measured Performance (Development Build)

**Append Latency:**
- AUTO-SCROLL mode: <0.5ms per line (synchronous layout + render)
- Batch append (100 lines): ~15ms total (~0.15ms per line)

**Throughput:**
- Sustained rate: 10,000 logs/second without frame drops
- Peak burst: 50,000 logs/second (short duration)
- ETW batch processing: 100-500 messages per batch

**Memory Usage:**
- Base overhead: ~5MB (ColorTextView + resources)
- Per-line overhead: ~150 bytes (Line struct + cached strings)
- 100K lines: ~20MB total
- Icon cache: ~2MB (cached D2D bitmaps for prefixes)

**Rendering Performance:**
- AUTO-SCROLL: 60fps sustained during continuous append
- SCROLL-BACK: 60fps smooth scrolling with slice caching
- Fallback rendering: 30-45fps (acceptable for cold path)

**Layout Cache Efficiency:**
- Slice cache hit rate: >95% during scroll-back navigation
- Layout creation: ~5ms per 256-line slice (amortized via caching)
- Dirty region optimization: 70% reduction in redraw area

### Performance Targets vs. Actuals

| Metric | Target | Actual | Status |
|--------|--------|--------|--------|
| Append latency (AUTO-SCROLL) | <0.5ms | <0.5ms | ✅ Met |
| Throughput (sustained) | 10K logs/sec | 10K+ logs/sec | ✅ Met |
| Frame rate (AUTO-SCROLL) | 60fps | 60fps | ✅ Met |
| Frame rate (SCROLL-BACK) | 60fps | 60fps | ✅ Met |
| Memory (100K lines) | <50MB | ~20MB | ✅ Better |
| Slice cache hit rate | >90% | >95% | ✅ Better |

### Optimization Opportunities

**Future Improvements:**
1. **Adaptive tail layout size**: Dynamically adjust based on append rate
2. **GPU-accelerated text rendering**: Direct2D1.2 color glyph rendering
3. **Compressed line storage**: Reduce memory for inactive slices
4. **Incremental layout**: Update only changed regions instead of full slice
5. **Background layout creation**: Prefetch slices on worker thread

## Decisions and improvement plan
### Selection and copy
- Support copy modes: payload-only, payload+metadata prefix, CSV, and JSON; keep Ctrl+C for current selection and add menu/toolbar split-button for mode choice.
- Quick actions: copy current line, copy visible region, and copy filtered result set; expose hotkeys and context menu entries.
- Allow pause-on-select to freeze intake/auto-scroll while selecting; resume restores live tail.
- Stretch goals: block (column) selection for timestamps/PIDs and multi-range copy (spec stays optional but keep state machine ready).
- Selection highlight uses small rounded corners to match the rest of the UI (see `Specs/VisualStyleSpec.md`).

### Filtering and navigation

**Filtering System (IMPLEMENTED)**:

RedSalamanderMonitor implements a high-performance type-based filtering system using bit masks and a **VisibleLine Index Pattern** for efficient filtered line access:

**Type-Based Filtering**:
- 5-bit filter mask controls visibility by `InfoParam::Type`:
  - Bit 0: Text messages (0x01)
  - Bit 1: Error messages (0x02) 
  - Bit 2: Warning messages (0x04)
  - Bit 3: Info messages (0x08)
  - Bit 4: Debug messages (0x10)
- Default mask: `0x1F` (all types visible)
- Example masks:
  - Errors only: `0x02`
  - Errors + Warnings: `0x06`
  - Errors + Warnings + Info: `0x0E`
  - All types: `0x1F`

**Architecture - VisibleLine Index Pattern**:
- **No sentinel values** - clean separation between visible and source line spaces
- `visibleLines` vector contains only visible lines (no filtered entries)
- Each VisibleLine maps sourceIndex → displayRowStart
- Display rows are densely packed (0, 1, 2...) for visible lines only
- Source line numbers always shown in gutter (sourceIndex + 1)
- Efficient O(1) visible access, O(log n) source lookup

**Key Data Structures**:
```cpp
// Document class (ColorTextView::Document):
struct VisibleLine {
    size_t sourceIndex;      // Index into source lines vector
    UINT32 displayRowStart;  // First display row for this visible line
};

uint32_t _filterMask;                // 5-bit visibility mask
std::vector<Line> _lines;            // Source lines (all lines, unfiltered)
std::vector<VisibleLine> _visibleLines; // Computed view (visible lines only)
```

**Filter Operations**:
```cpp
// ColorTextView API:
void SetFilterMask(uint32_t mask);     // Change filter, triggers full display sync
size_t GetVisibleLineCount() const;    // Returns count of visible lines
size_t GetTotalLineCount() const;      // Returns total lines (unfiltered)

// Document API:
void setFilterMask(uint32_t mask);                   // Sets mask, invalidates caches
bool isLineVisible(size_t allLineIndex) const;       // Check visibility (with lock)
bool isLineVisibleUnsafe(size_t allLineIndex) const; // Check visibility (no lock)
const Line& getLine(size_t visibleIndex) const;      // Access by visible index
const Line& getLineAll(size_t allLineIndex) const;   // Access by ALL index
```

**Display Synchronization Flow**:
When filter mask changes via `SetFilterMask()`:
1. `Document::setFilterMask(mask)` updates mask
2. Invalidates display offsets (`_displayOffsetsValid = false`)
3. Invalidates visible index (`_visibleIndexValid = false`)
4. Marks all caches dirty (`CacheInvalidationReason::FilterChanged`)
5. Recalculates `_contentHeight` from `totalDisplayRows()`
6. Clamps scroll position to new valid range (`ClampScroll()`)
7. Updates scrollbar ranges (`UpdateScrollBars()`)
8. Invalidates all layouts (`_sliceBitmap`, `_tailLayoutValid`, `_fallbackValid`)
9. Forces full redraw (`RequestFullRedraw()`)

**VisibleLine Rebuild** (in `rebuildVisibleLines()`):
```cpp
// Rebuild visibleLines from scratch based on current filter mask:
visibleLines.clear();
UINT32 displayRow = 0;

for (size_t srcIdx = 0; srcIdx < lines.size(); ++srcIdx) {
    if (isLineVisibleUnsafe(srcIdx)) {
        visibleLines.push_back({srcIdx, displayRow});
        displayRow += lines[srcIdx].newlineCount + 1u; // Multi-line support
    }
}
// Result: visibleLines contains only visible entries with display row offsets
```

**Binary Search for Display Row** (in `visibleIndexFromDisplayRow()`):
```cpp
// Binary search on visibleLines vector:
auto it = std::lower_bound(
    visibleLines.begin(), visibleLines.end(),
    displayRow,
    [](const VisibleLine& vl, UINT32 row) { return vl.displayRowStart < row; }
);

if (it != visibleLines.begin()) --it; // Find line containing displayRow
return static_cast<size_t>(it - visibleLines.begin());
```

**Y Position Calculation** (in `displayRowForSource()`):
```cpp
// Binary search visibleLines to find source line:
auto it = std::lower_bound(
    visibleLines.begin(), visibleLines.end(),
    sourceIndex,
    [](const VisibleLine& vl, size_t src) { return vl.sourceIndex < src; }
);

if (it != visibleLines.end() && it->sourceIndex == sourceIndex)
    return it->displayRowStart; // Found exact match

// Filtered line - return display row of next visible line (or total rows if none)
return (it != visibleLines.end()) ? it->displayRowStart : totalDisplayRows();
```

**Text Layout Building with Filtering**:
When filtering is active, layout workers iterate visible lines directly:
```cpp
// In StartLayoutWorker(), RebuildTailLayout(), CreateFallbackLayoutIfNeeded():
const size_t visCount = _document.visibleLineCount();
for (size_t visIdx = firstVisIdx; visIdx <= lastVisIdx && visIdx < visCount; ++visIdx) {
    const auto& line = _document.getVisibleLine(visIdx);
    text += _document.buildDisplayText(line);
    if (visIdx < lastVisIdx) text += L'\n';
}
// No need to check isLineVisible() - visibleLines already filtered
```

**Line Number Rendering**:
```cpp
// In DrawLineNumbers():
const auto [startVisIdx, endVisIdx] = GetVisibleLineRange(); // Returns visible indices
for (size_t visIdx = startVisIdx; visIdx <= endVisIdx; ++visIdx) {
    const auto& vl = _document.visibleLines[visIdx];
    const UINT32 displayRow = vl.displayRowStart;
    const float yBase = displayRow * lineHeight;
    const UINT32 lineNumber = static_cast<UINT32>(vl.sourceIndex + 1); // Show source line number
    // Render lineNumber at yBase...
}
// Directly iterate visible lines - no filtering checks needed
```

**Performance Characteristics**:
- **O(1) visibility check**: Bit mask test on metadata type (isLineVisibleUnsafe)
- **O(1) visible line access**: Direct visibleLines[visIdx] access
- **O(1) display row by visible index**: visibleLines[visIdx].displayRowStart
- **O(log n) display row by source index**: Binary search visibleLines for sourceIndex
- **O(log n) display row to visible index**: Binary search visibleLines by displayRowStart
- **O(n) filter change**: Full rebuild of visibleLines vector
- **O(1) incremental append**: Single visibleLines.push_back() if line visible

**Invariants Maintained**:
1. **Character positions always reference source lines** (unfiltered document)
2. **Display rows are dense** (0, 1, 2... with no gaps for visible lines only)
3. **Line numbers show source positions** (sourceIndex + 1, not visible index)
4. **visibleLines sorted by sourceIndex** (ascending order, always consistent)
5. **visibleLines.size() matches visible line count** (no filtered entries)
6. **Content height reflects visible lines only** (totalDisplayRows * lineHeight)
7. **Scroll position clamped to valid range** (prevents scrolling beyond visible content)

**Content Filters (NOT YET IMPLEMENTED)**:
- Content filters: include/exclude tokens and regex with case toggle; default to highlight-in-place (do not drop) so context remains; optional "hide non-matching" view.
- Metadata filters: PID, TID, and time-window based on prefix timestamp; allow saving named presets per producer.
- Navigation: next/previous match works even while paused; show filter badges and counts; keep pause/resume intake from UI.

### Performance and throughput
- **AUTO-SCROLL mode optimizations**: 
  - Dynamic tail layout (viewport-sized, min 100 lines)
  - Synchronous updates with batched ETW intake
  - Direct rendering with display row mapping
  - No virtualization overhead 
  - → <0.5ms append latency achieved
- **SCROLL-BACK mode optimizations**: 
  - Full slice virtualization (256-line blocks)
  - Async workers with layout caching
  - Bitmap caching (when within 16384px texture limit)
  - Display row offset mapping for correct positioning
  - → Acceptable latency for historical review
- Replace content scanning in `AddLine` with severity-driven coloring from metadata; avoid per-append keyword rescans.
- Batch intake: queue ETW events (via `WM_APP_ETW_BATCH`), coalesce append/layout/width work per frame, and clamp history with a bounded ring (configurable cap, optional disk spill).
- Instrument append→layout→paint durations with metrics: target zero blank lines (ACHIEVED), zero dropped frames at 10k logs/sec in AUTO-SCROLL mode.

### Communication (ETW/TraceLogging Only)
- Uses TraceLogging provider to emit structured events for every debug message (type, pid, tid, filetime, payload).
- Debug-build call tracing and indentation: `TRACER`/`TRACER_CTX` adjust per-thread indentation and (by default) only log the Exiting message; use `TRACER_INOUT`/`TRACER_INOUT_CTX` to log Entering+Exiting 
- ETW-only architecture with counters tracking writes and failures.
- ETW events batched via `WM_APP_ETW_BATCH` message from the EtwListener worker thread for optimal UI performance.
- No window discovery dependency - applications emit ETW events regardless of consumer presence.
- Monitor surfaces ETW statistics in UI/status bar, logging write failures when they occur.

### UX and robustness
- Cache toolbar PNG per-DPI bucket; show status indicators for transport mode, active filters, pause state, and drop counts.
- Clipboard failures report to status bar/toast; large-copy operations warn on truncation.
- Stress the system regularly with scripted scenarios (below) and capture metrics for regressions.
- Handle multi-line log messages correctly with display row offset mapping system.

### Stress tests (automation outline)
- High-rate append: 200k messages over 30s (mixed severities, mixed lengths, including multi-line content), verify drop counts, layout latency, and scroll smoothness.
- Multi-line content: Generate logs with embedded newlines, verify correct Y positioning throughout document (no blank lines).
- Large viewport testing: Test on high-DPI displays with large windows to verify dynamic tail window sizing.
- **Filter stress testing**:
  - **Rapid filter toggling**: Toggle filter buttons 100 times while streaming 10k messages/sec, verify no empty lines or sync issues
  - **Filter with scroll position**: Apply filters at various scroll positions (top, middle, bottom), verify scroll position clamps correctly
  - **Heavy filtering**: Filter to show only 1% of messages (e.g., errors only), verify dense display with no gaps
  - **Filter during batch append**: Toggle filters while processing large ETW batches, verify content height syncs
  - **All filter combinations**: Test all 32 possible 5-bit filter masks (0x00 through 0x1F), verify correct visibility
  - **Line number synchronization**: Verify line numbers match actual document positions throughout filtering changes
  - **Selection persistence**: Maintain selection across filter changes (if selection includes filtered lines)
- Filter-under-load: enable regex include/exclude and severity filters while appending 50k lines; ensure navigation and copy respect filters.
- Selection/copy under churn: pause, select large ranges, copy in all formats; resume and ensure caret/selection stability.
- Device/DPI resilience: change DPI mid-scroll and simulate device-loss during present; verify recovery without crashes.
- History cap: run until ring buffer eviction occurs; validate drop indicators and retained-length invariants.
- Texture limit testing: Generate enough content to exceed 16384px slice dimensions, verify fallback to direct rendering.

## Implementation Details

### Constants
```cpp
static constexpr size_t kTailLines = 100;           // Minimum tail window size
static constexpr size_t kSliceBlockLines = 256;     // Slice size for virtualization
static constexpr UINT32 kMaxD3D11TextureDimension = 16384;  // D3D11 texture limit
```

### Key Functions

**ColorTextView API**:
- `SetAutoScroll(bool)`: Set auto-scroll preference (triggers mode transition)
- `GetAutoScroll()`: Query current auto-scroll state
- `ShouldUseAutoScrollMode()`: Returns `_renderMode == RenderMode::AUTO_SCROLL` (internal helper)
- `RebuildTailLayout()`: Builds dynamic-sized tail layout for AUTO_SCROLL mode
- `displayRowForVisible(visIdx)`: Maps visible index to display row (O(1) access)
- `displayRowForSource(srcIdx)`: Maps source index to display row (O(log n) binary search)
- `visibleIndexFromDisplayRow(row)`: Maps display row to visible index (O(log n) binary search)
- `BeginBatchAppend()`/`EndBatchAppend()`: Optional batch API for callers; current ETW batch path uses `AppendInfoLine(..., deferInvalidation=true)` and a single update at end of batch.
- `GetTotalLineCount()`: Efficient O(1) line count (lines.size())
- `GetVisibleLineCount()`: Efficient O(1) visible line count (visibleLines.size())

**Filtering API (ColorTextView)**:
- `SetFilterMask(uint32_t mask)`: Change filter mask, triggers full display synchronization
- `GetVisibleLineCount() const`: Returns count of visible (not filtered) lines
- `GetTotalLineCount() const`: Returns total line count (all lines, filtered or not)

**Document Filtering API**:
- `setFilterMask(uint32_t mask)`: Sets filter mask, invalidates caches, rebuilds visibleLines
- `getFilterMask() const`: Returns current filter mask
- `isLineVisible(size_t sourceIndex) const`: Check if line is visible (with lock)
- `isLineVisibleUnsafe(size_t sourceIndex) const`: Check visibility (no lock, for internal use)
- `rebuildVisibleLines()`: Full O(n) rebuild of visibleLines from current filter
- `visibleLineCount() const`: Returns count of visible lines (O(1) - visibleLines.size())
- `totalLineCount() const`: Returns total line count (O(1) - lines.size())
- `getVisibleLine(size_t visibleIndex) const`: Access line by visible index
- `getSourceLine(size_t sourceIndex) const`: Access line by source index
- `displayRowForVisible(size_t visibleIndex) const`: Get display row by visible index (O(1))
- `displayRowForSource(size_t sourceIndex) const`: Get display row by source index (O(log n))
- `visibleIndexFromDisplayRow(UINT32 displayRow) const`: Binary search display row → visible index
- `totalDisplayRows() const`: Returns total display rows for visible lines only

**Main Window**:
- `UpdateStatusBar()`: Polls `GetAutoScroll()` state and syncs menu checkmark

### Message Handlers
- `WM_APP_ETW_BATCH`: ETW event batch processing from EtwListener worker thread
- `WM_PAINT`: Two-mode rendering based on current state
- `WM_VSCROLL`/`WM_MOUSEWHEEL`: Mode transitions on scroll

