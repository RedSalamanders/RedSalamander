# RedSalamanderMonitor Contract

## Shared UI dependency

Monitor builds against the exact external `DxUi.lib` pin through the product MSBuild
imports. `MakeMonitorDxPalette` remains the application adapter for resolved Monitor
colors and density. Toolbar/status hosts and the neutral frame clock use public DxUi
headers; no shared implementation sources are compiled in this application. Existing
Monitor append, scrollback, filter, focus and resource selftests qualify the consumer.
Library diagnostics follow the standalone contract, with no legacy logging requirement.
The generated module sidecar is part of the build receipt.

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
- Runtime code may only populate truly dynamic menu content (e.g., themes discovered from `Themes\\*.theme.json5` and custom themes from settings); see `Specs/Core/Core_Localization.md`.
- The top-level order is File, Edit, View, Options, Help. View owns Toolbar,
  Line Numbers, and Theme. Options is ordered Auto Scroll, Show Process/Thread
  IDs, Always on Top, separator, Active Filter.
- Active Filter uses the six message-type labels Error, Warning, Information,
  Performance, Debug, and Trace, followed by the four presets Errors only,
  Errors and warnings, Errors/performance/debug, and All message types. The
  existing stable `IDM_FILTER_*` IDs and bit meanings remain unchanged; the
  `IDM_FILTER_TEXT` presentation label is Trace.
- Every actionable static sibling has one access key unique within its
  separator-delimited group in each supported satellite.

### Two-Mode Rendering Architecture

**Mode 1: AUTO-SCROLL (Hot Path - 99% of usage)**
- Optimized for blazing-fast append and display of new lines
- Maintains a **dynamic tail layout** of the last N lines where N = `max(100, visibleLines + 50)`
  - Adapts to viewport size and DPI
  - Ensures tail window always covers visible area plus margin
  - Prevents blank areas on large/high-DPI displays
- Direct rendering without complex virtualization or caching
- Synchronous layout updates for immediate visibility
- Current latency and frame budgets are the instrumented gates in the performance contract below; unarchived point estimates are not acceptance evidence

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
  - The cross-thread ETW queue is a bounded `std::deque` drained in chunks of at most 200 on the UI thread. When
    the configured producer cap is reached, the oldest queued event is dropped so current diagnostics continue;
    remaining work stays ordered and one follow-up message is posted.
  - Batch processing moves the drained chunk into `Document::AppendInfoLines(...)`, taking the document write lock once and performing a single mode-specific update at the end of the batch.
- Single main window with menu and toolbar: New/Open/Save As, Copy, toggle toolbar, toggle line numbers, show/hide IDs, auto-scroll, always-on-top; debug builds can start a random message generator.
- Display surface is `ColorTextView` rendered with Direct2D/DirectWrite on a D3D11/DXGI swap chain; per-monitor DPI aware; supports line numbers, colored metadata prefixes, keyword colorization for Error/Warning/Debug, and optional auto-scroll.
- Input pipeline normalizes CR/LF, appends text to the document, caches prefixes per line, and updates gutter width; selection and clipboard copy are supported (Ctrl+C, Ctrl+A); find bar via Ctrl+F with F3 navigation; mouse wheel scroll with Shift for horizontal.
- File open is cancellable worker I/O with strict streaming UTF-8/UTF-16LE validation, encoded/decoded-byte and
  line budgets, and move-publication of one decoded immutable record snapshot. Save As captures one exact start
  snapshot by sharing immutable record blocks, then converts/writes strict transactional UTF-8+BOM output on an owned worker. Registry-backed email configuration
  and the about dialog remain, and device-loss/DPI-change handling recreates swap-chain targets.

### Bounded pipeline and lifetime contract

- `monitor.retention` owns four validated settings. Defaults are `maxQueuedEvents=4096`,
  `maxRetainedLines=100000`, `maxRetainedTextBytes=67108864`, and `maxSearchMatches=100000`. Load clamps them
  respectively to 200–1,000,000 events, 1,000–5,000,000 lines, 1 MiB–4 GiB, and 1,000–1,000,000 matches.
- Queue overflow and retained-history eviction are oldest-first. The status strip includes the cumulative dropped
  count. `monitor.etw.queue_high_water_mark`, `monitor.etw.dropped_count`,
  `monitor.document.retained_text_bytes`, and `monitor.document.retained_lines` make the bound observable.
- `Document` owns lines in a `std::deque`; each logical line owns one reader-independent immutable
  `MonitorTextBlock` from `MonitorTextSnapshot.h`. Replacing or extending text publishes a new block, while ordinary
  ETW append moves a newly built block into the deque. The document tracks UTF-16 payload bytes and applies the line
  and byte ceilings in one eviction operation. Eviction drops only document ownership; an in-flight Save As snapshot
  keeps its shared blocks valid. Eviction must shift/clamp selection, caret, scrolling, display rows, width/layout generations,
  visible-line mappings, and search offsets together.
- Search indexes only appended ranges while query/case/filter policy is stable, tracks an explicit line/offset scan
  frontier, shifts matches and the frontier after oldest eviction, and stores no more than `maxSearchMatches`. When
  capacity reopens, it resumes at the oldest retained unscanned position before scanning the new tail. Query, case,
  filter, Show IDs, or full-text replacement may rebuild. Matches remain ordered for F3 navigation.
  `monitor.search.match_update_us` and `monitor.search.match_rebuild_us` distinguish the paths.
- `Document::GetVisibleLine`, `GetSourceLine`, `VisibleLines`, `GetDisplayText*`, and display-text batches return
  locked snapshots by value. No public API may return a line/container/string reference whose internal lock has
  already been released.
- ETW callback identity is per session through `EVENT_TRACE_LOGFILE::Context` / `EVENT_RECORD::UserContext`; a
  process-global listener pointer is forbidden. The worker receives its trace handle by value, passes only a local
  copy by address to `ProcessTrace`, and owns shared callback state. Stop disables callback acceptance, closes the
  consumer handle, waits at most five seconds, and may detach only the shared lifetime-owned consumer state after
  that real bounded wait. The MonitorTest shutdown drill covers both prompt join and bounded fallback.
- SCROLL_BACK layout and width workers capture the target HWND, `IDWriteFactory`, and `IDWriteTextFormat` COM
  references by value. They must not retain `ColorTextView*` or read UI-mutated COM/window members. Posted results
  use the registered payload helper and are drained at `WM_NCDESTROY`.
- Unicode copy uses `Common::Clipboard::TrySetUnicodeText`. The helper builds a `wil::unique_hglobal` payload before
  opening or emptying the clipboard, keeps WIL ownership through every failure, and calls `release()` only after
  successful `SetClipboardData`. Do not reintroduce per-Win32-call forwarding wrappers solely for fault injection.
- File open keeps the picker on the UI thread, then uses one owned `std::jthread` to read 64 KiB chunks. It rejects
  the encoded-byte budget from `GetFileSizeEx` before allocation, incrementally budgets retained UTF-16 bytes and
  lines before growth, carries partial UTF-8 scalars, UTF-16 code units, and surrogate pairs across chunks, and polls
  cancellation during both I/O and CPU decode. It accepts UTF-8 (optional BOM) and UTF-16LE BOM strictly and rejects
  UTF-16BE, overlong/truncated UTF-8, odd UTF-16, and unpaired surrogates. The worker publishes one
  `MonitorTextSnapshot`; `Document::SetTextSnapshot` moves those immutable block handles into the document without
  rebuilding a second whole-text payload. Only the current generation may publish on the UI thread.
- Reader record finalization is delimiter-aware. Each LF publishes exactly one record, including an empty record
  represented by a consecutive or trailing LF. EOF publishes only a nonempty unterminated record and never invents
  a record after a terminal LF. Empty and BOM-only input therefore contain zero records. CR is normalized away.
  Export writes one UTF-8 BOM followed by every record and exactly one LF terminator per record, so
  read/export/read is byte- and record-idempotent for canonical output and an exact line budget does not pay for a
  synthetic terminal record.
- File-open cancellation calls `request_stop` plus `CancelSynchronousIo`; close waits at most two seconds and may
  detach only the worker-owned path/limits/decoder payload after that wait. Progress and completion use the registered
  posted-payload helper, and `WM_NCDESTROY` drains/closes the registry so a late detached completion cannot target a
  reused HWND.
- Show IDs changes the display coordinate space. It therefore resets selection and caret to zero and rebuilds search
  matches/frontier; preserving stale display offsets across the toggle is forbidden.
- Open and Save worker admission is transactional. Their Monitor-local `noexcept` start helpers catch only
  `std::system_error` from thread construction, log once, map it to an HRESULT, roll back active/shared state, post no
  completion, and allow a later retry. `std::bad_alloc` remains fatal. The command path shows the one localized
  background-file-operation launch failure and leaves the live document and destination unchanged.
- Save As captures the complete document under its shared lock by copying immutable ownership handles, releases the
  lock before I/O, and starts an owned cancellable worker. All record blocks are immutable immediately, so the
  maximum mutable active tail and the text bytes copied by capture are both exactly zero. Capture allocates only the
  descriptor deque (currently 16 bytes per logical record). The captured snapshot is the exact export boundary:
  later append, retention, clear, or replacement is excluded and cannot invalidate it. Progress
  and completion use registered posted payloads. Close requests stop, cancels synchronous I/O, waits at most two
  seconds, and may detach only the worker-owned snapshot/state after the payload registry makes late posts harmless.
- Clear and full text replacement invalidate pending layout/width generations. Clear also resets caret, selection,
  mouse-selection, search, pending-scroll, and layout state so stale interaction state cannot survive an empty view.

## Existing performance/architecture traits
- **Two-mode rendering system**: AUTO-SCROLL mode uses dynamic tail layout (viewport-sized, min 100 lines) with direct rendering for <0.5ms append latency; SCROLL-BACK mode uses full virtualization with slice-based rendering (`kSliceBlockLines=256`) for historical review.
- **Display row mapping**: All Y position calculations use display-row offsets (`Document::DisplayRowForVisible()` / `Document::DisplayRowForSource()`) to correctly handle multi-line content with embedded newlines.
- **Batched message intake**: ETW events are processed in bounded batches via `WM_APP_ETW_BATCH` to avoid per-message overhead at high throughput while preventing a single mega-burst from monopolizing the UI thread.
- **D3D11 texture limits**: Validates slice bitmap dimensions against 16384px limit, falls back to direct rendering when exceeded.
- Layout and width measurements in SCROLL-BACK mode run on a thread pool with immutable HWND/COM snapshots and
  slice prefetch; layouts are cached to skip reflow when the same slice is requested.
- AUTO-SCROLL mode uses synchronous layout updates and direct-to-backbuffer rendering; SCROLL-BACK mode uses offscreen slice bitmap when possible (within texture limits), otherwise direct rendering.
- Partial present is used for scroll deltas to minimize redraw.
- Line metadata (time/pid/tid/type) and brushes are cached; display-row offsets are precomputed for quick gutter/hit-testing.
- Mode transitions are explicit: menu toggle calls `SetAutoScroll()`, wheel/scrollbar scrolling up switches to SCROLL-BACK, and End/SB_BOTTOM switches to AUTO-SCROLL.
- **Mode detection timing (current implementation)**: `WM_APP_ETW_BATCH` appends the whole batch first, then applies AUTO_SCROLL vs SCROLL_BACK invalidation/layout behavior once per batch.
- **ColorTextView frame metrics**: monitor paint paths emit per-frame aggregate rows `monitor.frame.total_us`, `monitor.frame.present_us`, `monitor.frame.mode`, and `monitor.frame.tail_layout_us`. AUTO_SCROLL append visibility is measured with `monitor.frame.append_to_visible_us`. `monitor.frame.scrollback_slice_us` remains optional and is only expected for real SCROLL_BACK scenarios.
- **Monitor scrollback drill**: `RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-scrollback-selftest` is the focused SCROLL_BACK frame gate. It appends enough ETW lines for historical navigation, toggles auto-scroll through `IDM_OPTION_AUTO_SCROLL`/`WM_COMMAND`, drives ColorTextView scroll input, forces a visible render, restores auto-scroll through the same command path, and adds `monitorScrollbackSelfTest.metricPresence` plus p50/p95/p99/max summaries for `monitor.frame.scrollback_slice_us`, `monitor.frame.mode`, `monitor.frame.total_us`, and `monitor.frame.present_us`.
- **ETW batch and scheduling metrics**: `monitor.etw.batch_drain_us` records UI-thread batch-drain cost,
  `monitor.etw.queue_depth` records queue depth at drain start, `monitor.etw.queue_high_water_mark` records the peak,
  `monitor.etw.dropped_count` records queue plus retained-history drops, and `monitor.etw.batch_repost_count` records
  whether overflow required a follow-up batch message. Paint scheduling is intentionally a no-op unless there is
  real pending work.
- **Monitor ETW latency drill**: `RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260` is the focused append-to-visible perf gate. It preserves the default chrome selftest path and adds `monitorEtwBurstLatency.metricPresence` plus p50/p95/p99/max `metricSummary` rows to `results.json`.
- **Monitor scheduling gate**: frame scheduling must preserve append order, avoid self-posting when no paint/append work is pending, and keep AUTO_SCROLL vs SCROLL_BACK behavior explicit. Any scheduling optimization requires same-machine `append_to_visible`, `batch_drain`, and frame metric evidence before being called an improvement. If the latency drill keeps `monitor.etw.batch_drain_us` p95 at or below `8,333us`, p99 at or below `16,667us`, and `monitor.frame.append_to_visible_us` p95 at or below `50,000us`, scheduling changes are a measured no-op and should not be promoted.
- **Retained-state I/O metrics**: `monitor.file_open.total_us` records worker read/decode time and carries encoded
  bytes plus decoded line count; `monitor.file_open.peak_retained_text_bytes` records peak decoded payload;
  `monitor.file_open.publish_us` records UI move-publication. `monitor.file_open.cancel_us` and
  `monitor.file_open.close_us` cover cancellation request and bounded shutdown. Save separates
  `monitor.file_save.snapshot_us`, `snapshot_lock_us`, `ui_return_us`, `total_us`, `throughput_bytes_per_sec`,
  `cancel_us`, and `close_us`. Capture also reports `retained_text_bytes`, `retained_line_count`,
  `shared_block_count`, `shared_block_bytes`, `copied_text_bytes`, `active_tail_copied_bytes`, and
  `peak_additional_snapshot_bytes`.
- **Immutable-snapshot Release gate**: the test-enabled Chrome drill runs 512, 4,096, and 16,384 retained records
  five times each and compares the legacy full-string copy with immutable capture. Candidate
  `copied_text_bytes` and `active_tail_copied_bytes` MUST remain zero and additional snapshot bytes MUST remain at
  or below `record_count * sizeof(MonitorTextBlock)`. On machine `4cb089111a23`, the 16,384-record/8,060,928-byte
  p95 snapshot-lock regression threshold is 1,500 us; the accepted 2026-08-17 result was 1,008 us versus 3,395 us
  for the faithful legacy deep-copy comparator. Evidence is
  `Specs/TestRuns/4cb089111a23/Monitor/2026-08-17_144304/`.

Historical blank-line and filter-synchronization fixes are represented by the
current display-row/filter invariants and regression tests. The closed bug diary
is available from Git history and is not part of the current contract.

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
void SetFilterMask(uint32_t mask);                // Set type filter (6-bit mask)
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
MonitorFileReadResult ReadMonitorTextFile(path, stopToken, limits, progress);
void ColorTextView::SetTextSnapshot(MonitorTextSnapshot&& snapshot);
MonitorTextSnapshot ColorTextView::CaptureTextSnapshot() const;
MonitorFileExportResult WriteMonitorTextSnapshot(path, snapshot, stopToken, progress);
```

- Monitor text export MUST encode every line with strict UTF-16-to-UTF-8 conversion and fail the entire save if
  malformed input is encountered. It must not silently replace invalid code units.
- Save writes MUST use `Common::Files::LocalFileTransaction`: write a UTF-8 BOM and the complete document to a
  unique sibling, verify/flush it, then atomically replace the requested local target. Conversion, write, flush,
  verification, or promotion failure MUST preserve the previous target and clean up the temporary sibling.
- The UI captures exactly one start snapshot and releases the document lock before worker conversion/I/O. A save
  already in progress is cancelled instead of silently starting a second exporter. Later ingestion is not included
  in the captured snapshot.
- Canonical export is a UTF-8 BOM plus exactly one LF-terminated encoding of each logical record. Empty/BOM-only
  input has zero records; terminal LF never synthesizes another empty record at EOF; consecutive LFs remain real
  blank records.
- `MonitorTest --document-model-selftest` is the focused regression guard for malformed-input preservation,
  successful UTF-8+BOM output, the empty/unterminated/terminated/CRLF/blank/trailing-blank/Unicode/budget matrix,
  exact start-snapshot behavior across append/retention/clear/replacement, cancellation/fault target preservation,
  zero-copy practical-large capture, and abandoned-temporary cleanup.

### Document Class Methods

**Line Access:**
```cpp
Line GetVisibleLine(size_t visibleIndex);         // Locked snapshot by visible index (0-based)
Line GetSourceLine(size_t sourceIndex);           // Locked snapshot by source index (all lines)
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

## Performance evidence contract

The authoritative current metrics, deterministic latency and scrollback drills,
and promotion thresholds are listed under Existing performance/architecture
traits above. Static development-build estimates and historical target/actual
tables are not acceptance evidence. Any new performance claim follows
`Specs/Testing/Testing_PerformanceValidation.md` and cites an archived
`Specs/TestRuns/` comparison.

## Filtering, selection, and export contract

- The type filter is a six-bit mask: Text `0x01`, Error `0x02`, Warning
  `0x04`, Info `0x08`, Perf `0x10`, and Debug `0x20`; `0x3F` shows all
  types. Filter changes rebuild the visible-line mapping, clamp scroll state,
  invalidate affected layouts, and keep source line numbers in the gutter.
- Selection remains in source-document coordinates. Copying while a type filter
  is active includes only visible selected text and uses the shared Unicode
  clipboard publication helper.
- File export writes the complete captured start snapshot through the strict asynchronous transactional contract in
  Public API Reference; it is not a clipboard-format selector. Ingestion after capture remains in the live document
  and is intentionally absent from that export.
- Optional clipboard formats, richer content/metadata filters, and alternative
  storage/rendering optimizations are not current behavior. Decisions are owned
  by
  `Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md`.
- Bounded ETW intake and retained line/text/search ceilings are current. Disk
  spill is not part of the contract.

## Communication (ETW/TraceLogging only)
- Uses TraceLogging provider to emit structured events for every debug message (type, pid, tid, filetime, payload).
- `RedSalamander.exe` Debug and ASan Debug builds emit Info/Perf/Debug ETW diagnostics and write JSONL perf capture to the default path by default; Release builds require `--etw` for ETW and use `--perf` for the default JSONL path or `--perf=PATH` for a custom path.
- Normal Debug and Release builds of `RedSalamanderMonitor` do not emit their own Info/Perf/debug-style ETW diagnostics unless launched with `--etw`. Error and warning diagnostics remain eligible for ETW visibility.
- Normal Debug and Release builds of `RedSalamanderMonitor` must keep its own display quiet: no startup/sample status text, and ETW events whose process id matches the monitor process are filtered out. Launching with `--etw` may display the monitor's own startup status and self-originated ETW events. Explicit JSONL perf capture uses `--perf` for the default path or `--perf=PATH` for a custom path.
- Debug-build call tracing and indentation: `TRACER`/`TRACER_CTX` adjust per-thread indentation and (by default) only log the Exiting message; use `TRACER_INOUT`/`TRACER_INOUT_CTX` to log Entering+Exiting 
- ETW-only architecture with counters tracking writes and failures.
- ETW events batched via `WM_APP_ETW_BATCH` message from the EtwListener worker thread for optimal UI performance.
- No window discovery dependency - applications emit ETW events regardless of consumer presence.
- Monitor surfaces ETW statistics in UI/status bar, logging write failures when they occur.
- Default Debug / ASan Debug builds of `RedSalamanderMonitor` MUST NOT enable `ENABLE_TESTS` automatically. Monitor-specific test hooks are opt-in only, otherwise the monitor can surface its own test/perf ETW chatter and pollute the default live display.
- Monitor chrome selftest coverage requires a test-enabled `RedSalamanderMonitor` build and must verify that `monitor.frame.total_us`, `monitor.frame.present_us`, `monitor.frame.append_to_visible_us`, `monitor.frame.tail_layout_us`, `monitor.frame.mode`, and `monitor.etw.batch_drain_us` are present. Its retained-state I/O drill must also verify every open/save metric listed under Existing performance/architecture traits, run the three-size/five-repeat immutable-snapshot matrix, compare exact canonical output bytes, and inject both open and Save thread-start failures through the production admission helpers. Failure must leave no active flag, shared worker state, joinable thread, posted completion, or destination mutation, and the immediate retry must start successfully. The drill creates and removes its own deterministic streaming-open/export fixtures and fails if cleanup leaves any file behind. The opt-in ETW latency burst mode must additionally verify `monitor.etw.selftest_burst_drain_us`, `monitor.etw.queue_depth`, and `monitor.etw.batch_repost_count`, and must include quantile summaries for append-to-visible, batch-drain, total-frame, and present timings. The opt-in scrollback mode must additionally verify `monitor.frame.scrollback_slice_us` without making that cold-path metric mandatory for the default chrome selftest. Test-enabled Debug and Release builds are produced with `RSBuildEnableTests=true`; Release `ClCompile` definitions must include `$(RSBuildTestDefinitions)` so the opt-in defines `ENABLE_TESTS` without enabling Monitor self diagnostics by default. Helper functions used only by these opt-in drills must also stay behind `ENABLE_TESTS` so normal builds do not carry unreferenced selftest code or warning noise.
- `MonitorTest.exe --document-model-selftest` is the focused non-UI Track 6 guard. It covers retained line/byte
  eviction, decoded expansion limits, split UTF-8/UTF-16 boundaries, malformed/truncated input, mid-decode
  cancellation, move-publication, exact start-snapshot export, cancellation/fault target preservation, Unicode
  clipboard storage construction, and ETW consumer shutdown through both the normal join path and a deterministic
  short-timeout lifetime-owned fallback.
- Invalid/dirty rectangle visualization MUST be opt-in only. Default Debug builds must not paint invalid rectangles in color; rebuilding with `RS_MONITOR_SHOW_INVALID_RECTS` enables that visual diagnostic. All brushes, palette state, and paint overlays for this diagnostic must stay behind the `RS_MONITOR_INVALID_RECT_VISUALIZATION_ENABLED` contract.

## Verification ownership

- `MonitorTest.exe --document-model-selftest` owns document batching/filtering,
  retention eviction, exact canonical and cancellable file reads, immutable-snapshot lifetime/scaling, transactional export,
  Unicode clipboard construction, and ETW shutdown lifetime.
- `RedSalamanderMonitor.exe --chrome-selftest` owns live window chrome, filter
  status, bounded queue/search-frontier/Show IDs behavior, transactional worker admission, retained-state I/O metric presence, payload-safe bounded
  open close, fixture cleanup, queue drain, and required frame metrics.
- `--monitor-etw-burst-mode=latency` and `--monitor-scrollback-selftest` own the
  focused hot-tail and cold-scrollback metric gates.
- DPI/device-loss, high-rate mixed multi-line input, filter changes, selection,
  and texture-limit fallback remain correctness requirements. New performance
  claims require the scenario and archived evidence contract above; a historical
  unchecked stress outline is not proof.
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
- `BeginBatchAppend()`/`EndBatchAppend()`: Optional legacy batch API for callers; current ETW batch path uses `Document::AppendInfoLines(...)` and a single update at end of batch.
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
- `appendInfoLines(std::vector<InfoLineInput>)`: Append structured monitor log entries under a single write lock
- `displayRowForVisible(size_t visibleIndex) const`: Get display row by visible index (O(1))
- `displayRowForSource(size_t sourceIndex) const`: Get display row by source index (O(log n))
- `visibleIndexFromDisplayRow(UINT32 displayRow) const`: Binary search display row → visible index
- `sourceLineForDisplayRow(UINT32 displayRow) const`: Binary search display row → source line without exposing the visible index vector
- `closestVisibleSourceLine(size_t sourceIndex) const`: Use the rebuilt visible index to preserve a filter anchor by exact/next visible source line, or the last visible line when the anchor is beyond the filtered tail
- `totalDisplayRows() const`: Returns total display rows for visible lines only

**Main Window**:
- `UpdateStatusBar()`: Polls `GetAutoScroll()` state and syncs menu checkmark

### Message Handlers
- `WM_APP_ETW_BATCH`: ETW event batch processing from EtwListener worker thread
- `WM_PAINT`: Two-mode rendering based on current state
- `WM_VSCROLL`/`WM_MOUSEWHEEL`: Mode transitions on scroll
