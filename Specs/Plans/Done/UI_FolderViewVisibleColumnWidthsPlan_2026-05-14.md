# FolderView Visible Column Widths Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox syntax for tracking.

**Goal:** Make FolderView's column-based pane layout size each vertical column from the items assigned to that column after the current pane height is known, instead of sizing every column from the widest item in the entire folder.

**Architecture:** First add the test-only audit hooks, crazy-folder scenario, failing behavior tests, and baseline measurement run against the current layout. Then add a small pure column-layout helper that receives per-item text-width estimates and current pane metrics, returns row count, per-column x/width/start/count, and content width. FolderView keeps its existing text measurement and rendering responsibilities, but stores per-column layout metrics and routes rendering, hit testing, horizontal scrolling, visible-range lookup, text layout widths, `EnsureVisible`, and scroll metrics through those metrics.

**Tech Stack:** C++23, Win32, DirectWrite text metrics, Direct2D FolderView rendering, WIL, RedSalamander command selftests, CppUnitTest, Debug perf JSONL.

---

## Progress Checklist

- [x] Start implementation on branch `codex/folderview-visible-column-widths`.
- [x] Add this live progress checklist at the beginning of the WIP plan.
- [x] Build current HEAD before product layout changes.
- [x] Add current-layout audit/debug snapshots without changing FolderView behavior.
- [x] Add `folderView_column_widths_audit` and generate the crazy-folder scenario matrix.
- [x] Run and archive the before-code audit/perf baseline.
- [x] Add desired-behavior command selftest and verify it fails on current layout.
- [x] Add pure column-layout CppUnitTests and verify the red phase.
- [x] Implement per-visible-column layout helper and FolderView integration.
- [x] Update horizontal scrolling, visible range, rendering, hit testing, and `EnsureVisible`.
- [x] Run candidate audit, desired-behavior selftest, helper tests, and large-folder perf case.
- [x] Compare before/after perf and width artifacts on the same machine.
- [x] Add testing-spec note that adversarial filename fixtures must budget for the full selftest temp path before claiming near-max coverage.
- [x] Update authoritative specs and test docs for any durable behavior discovered during coding.
- [x] Move this plan to `Specs/Plans/Done/` after code, tests, perf, and specs are complete.

## Current Evidence

- `Specs/UI/UI_FolderView.md:129` documents global per-mode column width calculation using the maximum text width across all items.
- `RedSalamander/FolderView.Layout.cpp:122` through `RedSalamander/FolderView.Layout.cpp:246` measures `maxLabelWidth`, `maxDetailsWidth`, and `maxMetadataWidth` across every item in `_items`.
- `RedSalamander/FolderView.Layout.cpp:253` through `RedSalamander/FolderView.Layout.cpp:268` assigns one global `_tileWidthDip`.
- `RedSalamander/FolderView.Layout.cpp:287` through `RedSalamander/FolderView.Layout.cpp:340` lays out every column using one `columnStride`.
- `RedSalamander/FolderView.Layout.cpp:385` passes one `labelWidth` to text-layout creation for all items.
- `RedSalamander/FolderView.Layout.cpp:583`, `RedSalamander/FolderView.Layout.cpp:1063`, and `RedSalamander/FolderView.Layout.cpp:1140` compute visible range, hit testing, and horizontal scrolling from the same global stride.
- `RedSalamander/FolderView.Rendering.cpp:1213` and `RedSalamander/FolderView.Rendering.cpp:2050` render with global `_tileWidthDip`.
- Current-layout audit baseline: `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_123634`, artifact `perf/folderView_column_widths_audit_metrics.json`.
- Desired-behavior red test: `folderView_visible_column_widths` failed on current layout in `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_124117` with equal column widths `c1=474.0`, `c2=474.0`.
- Helper implementation check: `.\build.ps1 -ProjectName PerformanceTests2` succeeded with 0 warnings / 0 errors, and `vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll /Tests:FolderViewColumnLayoutTests` passed 2 tests.
- Product implementation check: `.\build.ps1 -ProjectName RedSalamander` succeeded with 0 warnings / 0 errors after FolderView integration.
- Desired-behavior candidate test: `folderView_visible_column_widths` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_130036`.
- Candidate audit/perf: `folderView_column_widths_audit` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_130147`, artifact `perf/folderView_column_widths_audit_metrics.json`, with zero integrity failures.
- Large-folder candidate perf: `folderView_perf_large_folder_baseline` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_130206`.
- Before/after comparison on the same machine: audit case duration `67496 ms -> 59353 ms` (`-12.1%`), `render.frame_us` P95 `55465 us -> 53205 us` (`-4.1%`), `render.begin_to_enddraw_us` P95 `48058 us -> 46029 us` (`-4.2%`), and `render.draw_item_us` P95 `164 us -> 112 us` (`-31.7%`).

## Acceptance Contract

- Column width is computed after the pane height determines `rowsPerColumn`.
- Each column width is based only on the items assigned to that column.
- A long filename in column 2 must not widen columns 1, 3, or 4.
- The behavior applies to Brief, Detailed, ExtraDetailed, and Thumbnails display modes.
- Detailed mode considers the display-name line and details line for each item in that column.
- ExtraDetailed and Thumbnails modes consider display-name, details, and metadata lines for each item in that column.
- Widths retain existing minimums, icon/text padding, and client-width clamp behavior so a single column still fits the pane.
- Horizontal scrolling, visible-range calculation, drawing, selection bounds, `EnsureVisible`, and hit testing all use variable per-column left/right bounds.
- Hit testing inside inter-column spacing returns no item.
- `LayoutItems()` must not create DirectWrite layouts per visible cell as part of the new width calculation; it should keep measurement bounded to existing text metric work and cache invalidation.
- A deterministic command selftest proves the full FolderView path uses visible-column widths.
- A deterministic audit selftest measures current before-state display and horizontal scrolling before any product layout code is changed.
- A pure CppUnitTest proves the column-width math without a UI host.
- Perf evidence is archived under `Specs/TestRuns/` before closeout and includes layout/render metrics for the large-folder scenario.
- Final implementation is not accepted if it is slower or less stable than the current layout on the same machine without an explicit user decision. The candidate must be on par with baseline for render/layout/scroll metrics and clearly better for wasted horizontal width in ragged-column scenarios.

## Files To Create

- `RedSalamander/FolderViewColumnLayout.h`
- `Tests/PerformanceTests2/FolderViewColumnLayoutTests.cpp`

## Files To Modify

- `Tests/PerformanceTests2/PerformanceTests2.vcxproj`
- `Tests/PerformanceTests2/PerformanceTests2.vcxproj.filters`
- `RedSalamander/FolderViewInternal.h`
- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.Layout.cpp`
- `RedSalamander/FolderView.Interaction.cpp`
- `RedSalamander/FolderView.Rendering.cpp`
- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `Specs/UI/UI_FolderView.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Tests/README.md` if the new command selftest should be listed with named examples

## Mandatory Pre-Code Gate

No product layout change may start until this gate is complete:

- [x] Build current HEAD with `.\build.ps1 -ProjectName RedSalamander`.
- [x] Add only test/debug capture support needed to observe the current layout. This may include `#ifdef ENABLE_TESTS` snapshots and command selftests, but it must not change `LayoutItems()` behavior, scroll behavior, render behavior, or hit testing.
- [x] Add `folderView_column_widths_audit`, a command selftest that passes on the current layout and writes a JSON artifact describing display and scroll behavior for every scenario in the matrix below.
- [x] Run `folderView_column_widths_audit` on the unmodified product layout and archive the run under the current machine's `Specs/TestRuns` Commands area.
- [x] Add `folderView_visible_column_widths`, the desired-behavior command selftest, and verify it fails on current layout because all columns still share the global max width.
- [x] Add the pure CppUnitTest coverage before any implementation code. It may fail to compile or fail assertions until the helper exists; that is the expected red phase.
- [x] Record the exact baseline run directory and baseline artifact name in the implementation notes before changing the layout algorithm.

## Crazy Folder Scenario Matrix

The audit selftest must create all of these scenarios inside the command selftest temp root and drive them through the real left FolderView pane. Every scenario must run in Brief, Detailed, ExtraDetailed, and Thumbnails when the display mode is supported by the command test harness.

| Scenario | Folder Shape | Display Assertions | Scroll Assertions | Measurement Purpose |
|---|---|---|---|---|
| `column_poison_second_column` | Exactly `rowsPerColumn * 5` files sorted into five column blocks. Only block 2 has very long names. | Baseline records same global width for all columns. Candidate must widen only column 2. | Scroll left, one page right, right edge, one page left, and home. Visible item range must stay ordered. | Proves the user-reported bug and quantifies wasted width. |
| `column_poison_every_other_column` | Six column blocks alternating tiny names and long names. | Candidate widths must alternate narrow/wide without overlap. | Repeated `SB_LINERIGHT` and `SB_LINELEFT` must visit adjacent columns without skipping. | Catches stride assumptions left in scroll code. |
| `unicode_width_mess` | Names include CJK full-width characters, combining marks, right-to-left script, spaces, brackets, many dots, uppercase/lowercase extension variants, and emoji-equivalent surrogate pairs when NTFS accepts the name. | Item rectangles must not overlap and text width must remain finite. | End scroll and focus-last must keep the final item visible. | Protects DirectWrite measurement and UTF-16 handling. |
| `near_max_filename_lengths` | Files with names close to the Windows single-component limit while keeping total path length valid under the test temp root. | Long names clamp to pane width for their own column only. | Horizontal max must be finite and scrollbars must not produce negative offsets. | Catches overflow and clamp mistakes. |
| `tiny_and_empty` | Empty folder, one item, `rowsPerColumn - 1` items, and exactly `rowsPerColumn` items. | No phantom columns, no horizontal scroll for single-column content, empty-folder pseudo item unchanged. | `SB_RIGHT`, mouse horizontal wheel, Home, and End leave offsets clamped. | Protects small-count behavior. |
| `many_columns_large_folder` | At least 2400 files plus 48 directories across mixed extensions and extensionless names. Long names appear in only every tenth column. | Visible range remains bounded to visible columns plus existing preload buffer. | Focus first, middle, last, then first again; each focused item must become visible. | Protects responsiveness and icon/render churn. |
| `resize_changes_rows_per_column` | Start with `rowsPerColumn * 6` files, scroll to the middle, then resize the main window height enough to change `rowsPerColumn`. | Column recomputation must not leave stale item bounds. | Horizontal offset must clamp to valid content after resize and focused item must remain visible. | Catches stale column-cache and scroll-clamp bugs. |
| `details_metadata_poison` | Files with size/date/type detail strings and metadata-like display lines where only one column has long secondary text. | Detailed mode uses details only for that column; ExtraDetailed and Thumbnails include metadata only for that column. | Page-right/page-left must respect per-column widths. | Protects non-Brief modes. |

## Before/After Measurement Contract

`folderView_column_widths_audit` must write `folderView_column_widths_audit_metrics.json` through `SelfTest::GetPerfArtifactPath(...)`. The JSON must be an array containing one record per scenario and display mode; each record uses this shape:

```json
{
  "case": "folderView_column_widths_audit",
  "scenario": "column_poison_second_column",
  "displayMode": "Brief",
  "itemCount": 120,
  "clientWidthPx": 1280,
  "clientHeightPx": 720,
  "rowsPerColumn": 24,
  "columnCount": 5,
  "contentWidthDip": 1600.0,
  "maxHorizontalOffsetDip": 320.0,
  "integrity": {
    "overlapCount": 0,
    "outOfOrderColumnCount": 0,
    "negativeWidthCount": 0,
    "visibleRangeMismatchCount": 0,
    "hitTestSpacingFalsePositiveCount": 0
  },
  "scrollSamples": [
    {
      "name": "left",
      "horizontalOffsetDip": 0.0,
      "firstVisibleIndex": 0,
      "lastVisibleIndex": 72,
      "visibleColumnCount": 3
    }
  ],
  "columns": [
    {
      "index": 0,
      "startIndex": 0,
      "itemCount": 24,
      "leftDip": 0.0,
      "widthDip": 320.0,
      "rightDip": 320.0,
      "widestDisplayName": "a_000.txt"
    }
  ]
}
```

The baseline and candidate runs must be compared on the same machine with these commands run immediately after each selftest run:

```powershell
$commandRuns = Get-ChildItem Specs\TestRuns -Directory |
    ForEach-Object { Get-ChildItem (Join-Path $_.FullName 'Commands') -Directory -ErrorAction SilentlyContinue } |
    Sort-Object LastWriteTime
$latestCommandsRun = $commandRuns[-1].FullName
.\Tools\Show-PerfRuns.ps1 -Run $latestCommandsRun -Scenario folderView_column_widths_audit
```

After the candidate run:

```powershell
$commandRuns = Get-ChildItem Specs\TestRuns -Directory |
    ForEach-Object { Get-ChildItem (Join-Path $_.FullName 'Commands') -Directory -ErrorAction SilentlyContinue } |
    Sort-Object LastWriteTime
$baselineCommandsRun = $commandRuns[-2].FullName
$candidateCommandsRun = $commandRuns[-1].FullName
.\Tools\CompareTestRuns.ps1 $baselineCommandsRun $candidateCommandsRun -Suite Commands
.\Tools\Show-PerfRuns.ps1 -CompareRun $baselineCommandsRun,$candidateCommandsRun -Scenario folderView_column_widths_audit
.\Tools\Show-PerfRuns.ps1 -CompareRun $baselineCommandsRun,$candidateCommandsRun -Scenario folderView_perf_large_folder_baseline
```

Candidate acceptance thresholds:

- `integrity.overlapCount`, `integrity.outOfOrderColumnCount`, `integrity.negativeWidthCount`, `integrity.visibleRangeMismatchCount`, and `integrity.hitTestSpacingFalsePositiveCount` must be `0` for every candidate scenario.
- `column_poison_second_column` candidate `contentWidthDip` must be at least `25%` lower than baseline, unless baseline content already fits within one viewport; in that case candidate must fit too.
- `column_poison_every_other_column` candidate narrow-column widths must be at least `20%` lower than their adjacent poisoned columns.
- `details_metadata_poison` must show the long details or metadata line widening only the column that contains that item.
- `many_columns_large_folder` candidate `render.frame_us`, `render.begin_to_enddraw_us`, `render.present_us`, and `render.layout_items_us` p95 must be no more than `5%` worse than baseline and must not add more than `1.0 ms` absolute p95 latency for any metric.
- `many_columns_large_folder` candidate render call count, warm-render call count, icon queue count, and icon batch update count must not exceed baseline by more than `5%`.
- If a metric regresses beyond these thresholds, the implementation must be revised before closeout. If the regression is intentional, the user must approve that specific tradeoff with the baseline and candidate numbers visible.

## Implementation Steps

- [x] Add current-layout audit snapshot support before changing layout behavior.
  - Add `FolderView::DebugColumnLayoutEntry`, `FolderView::DebugVisibleItemEntry`, and `FolderView::DebugColumnLayoutSnapshot` under `#ifdef ENABLE_TESTS` in `RedSalamander/FolderView.h`.
  - Include for each column: `startIndex`, `itemCount`, `leftDip`, `widthDip`, and `rightDip`.
  - Include for each visible item: `index`, `displayName`, `bounds`, `columnIndex`, and `rowIndex`.
  - Include scroll state: `horizontalOffsetDip`, `maxHorizontalOffsetDip`, `contentWidthDip`, `clientWidthDip`, `clientHeightDip`, `firstVisibleIndex`, and `lastVisibleIndex`.
  - Implement `DebugGetColumnLayoutSnapshot()` against the current global `_tileWidthDip` layout first. It must describe current behavior without changing it.
  - Add `FolderWindow::DebugGetPaneColumnLayoutSnapshot(Pane pane, FolderView::DebugColumnLayoutSnapshot& out)` in `RedSalamander/FolderWindow.h`.
  - Implement the wrapper in `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`.

- [x] Add and run the current-layout audit selftest before product layout coding.
  - In `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`, add `TestFolderViewColumnWidthsAudit`.
  - Register the case as `folderView_column_widths_audit`.
  - Generate every folder in the Crazy Folder Scenario Matrix inside `SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands)`.
  - Drive each scenario through the real left pane by calling `g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, scenarioRoot)`.
  - For each display mode, set the view mode through existing pane view commands or direct test helpers already used by pane view-options tests.
  - Wait for enumeration, item count, and at least one render warmup.
  - Capture snapshots at scroll samples `left`, `line-right`, `page-right`, `right`, `page-left`, and `home`.
  - Send real scroll input through `WM_HSCROLL` with `SB_LINERIGHT`, `SB_PAGERIGHT`, `SB_RIGHT`, `SB_PAGELEFT`, and `SB_LEFT`; do not mutate `_horizontalOffset` directly.
  - For focus scrolling, focus first, middle, last, then first item using debug focus helpers and capture the post-`EnsureVisible` snapshot.
  - Compute overlap, ordering, negative-width, visible-range, and hit-test-spacing integrity counters from the snapshots.
  - Write `folderView_column_widths_audit_metrics.json` using `SelfTest::GetPerfArtifactPath(...)`.
  - Build current HEAD:
    ```powershell
    .\build.ps1 -ProjectName RedSalamander
    ```
  - Run current HEAD before any layout behavior changes:
    ```powershell
    .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_column_widths_audit
    ```
  - Confirm the run archives under the current machine's `Specs\TestRuns` Commands area and the artifact contains all eight scenario names.

- [x] Add desired-behavior tests and verify they fail before layout implementation.
  - Add the full-path command selftest `TestFolderViewVisibleColumnWidths` in `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`.
  - Register the case as `folderView_visible_column_widths`.
  - Reuse the `column_poison_second_column`, `column_poison_every_other_column`, `details_metadata_poison`, and `resize_changes_rows_per_column` scenario builders from the audit test.
  - Assert the desired variable-column behavior described in the Acceptance Contract.
  - Run current HEAD:
    ```powershell
    .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_visible_column_widths
    ```
  - Expected before implementation: FAIL with a message showing that at least one unpoisoned column shares the poisoned column width.

- [x] Add failing pure layout tests before layout implementation.
  - Create `Tests/PerformanceTests2/FolderViewColumnLayoutTests.cpp`.
  - Add it to `Tests/PerformanceTests2/PerformanceTests2.vcxproj` and `.filters`.
  - Add `VisibleColumnWidths_DifferentColumnsDoNotShareGlobalMax`.
    - Use 12 synthetic items, 3 rows per column, Brief-equivalent metrics.
    - Put the longest label only in the second column's item group.
    - Assert column 2 is wider than columns 1, 3, and 4.
    - Assert column 1 width is derived from its own max item plus padding, not the global max.
    - Assert each `leftDip` accumulates from the previous column's actual width plus spacing.
  - Add `VisibleColumnWidths_DetailedAndMetadataLinesStayColumnLocal`.
    - Use 8 synthetic items, 2 rows per column.
    - Enable details and metadata width participation.
    - Put the widest details line in column 3 and the widest metadata line in column 4.
    - Assert only those columns widen for those lines.
  - Run `.\build.ps1 -ProjectName PerformanceTests2` and confirm these tests fail to compile until the helper exists.

- [x] Create the pure column-layout helper.
  - Add `RedSalamander/FolderViewColumnLayout.h`.
  - Keep the helper header-only and free of HWND, Direct2D, DirectWrite, and FolderView object state so it is cheap to test.
  - Define:
    ```cpp
    namespace FolderViewColumnLayout
    {
        struct ItemTextMetrics
        {
            float labelWidthDip = 0.0f;
            float detailsWidthDip = 0.0f;
            float metadataWidthDip = 0.0f;
        };

        struct Input
        {
            float clientWidthDip = 0.0f;
            float clientHeightDip = 0.0f;
            float tileHeightDip = 0.0f;
            float rowSpacingDip = 0.0f;
            float iconSizeDip = 0.0f;
            float iconTextGapDip = 0.0f;
            float horizontalPaddingDip = 0.0f;
            float columnSpacingDip = 0.0f;
            float textWidthSafetyDip = 0.0f;
            bool includeDetailsLine = false;
            bool includeMetadataLine = false;
            std::span<const ItemTextMetrics> items;
        };

        struct Column
        {
            size_t startIndex = 0;
            size_t itemCount = 0;
            float leftDip = 0.0f;
            float widthDip = 0.0f;

            [[nodiscard]] float RightDip() const noexcept;
        };

        struct Result
        {
            int rowsPerColumn = 1;
            std::vector<Column> columns;
            float contentWidthDip = 0.0f;
            float maxColumnWidthDip = 0.0f;
        };

        [[nodiscard]] Result Resolve(const Input& input);
    }
    ```
  - In `Resolve`, compute `rowsPerColumn` from `clientHeightDip`, `tileHeightDip`, and `rowSpacingDip` using the same vertical packing rules FolderView currently uses.
  - For each sequential column group, compute text width from only `[startIndex, startIndex + itemCount)`.
  - Compute item width as `iconSizeDip + iconTextGapDip + maxTextWidth + horizontalPaddingDip + textWidthSafetyDip`.
  - Clamp each column width to `clientWidthDip` when `clientWidthDip > 0.0f`.
  - Clamp each column width to the existing minimum item width implied by icon, gap, padding, and safety.
  - Compute each column `leftDip` from the previous column's `RightDip()` plus `columnSpacingDip`.
  - Compute `contentWidthDip` from the last column right edge, clamped to at least `clientWidthDip`.
  - Re-run `.\build.ps1 -ProjectName PerformanceTests2`.
  - Run `vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll /Tests:FolderViewColumnLayoutTests`.

- [x] Replace FolderView's global column-width state with per-column metrics.
  - Include the helper from `RedSalamander/FolderView.h` or `RedSalamander/FolderViewInternal.h`.
  - In `RedSalamander/FolderView.h`, replace layout dependence on `_columnCounts` and `_columnPrefixSums` with:
    ```cpp
    std::vector<FolderViewColumnLayout::Column> _columnLayout;
    float _maxColumnWidthDip = 0.0f;
    ```
  - Keep `_columns`, `_rowsPerColumn`, `_contentWidth`, `_contentHeight`, `_horizontalOffset`, and `_tileWidthDip` for existing contracts; assign `_tileWidthDip` to the helper result's `maxColumnWidthDip` as a compatibility value until all remaining reads are removed or intentionally kept.
  - Add private helpers:
    ```cpp
    [[nodiscard]] float GetItemLabelWidthDip(const FolderItem& item) const noexcept;
    [[nodiscard]] std::pair<size_t, size_t> GetVisibleColumnRange(float leftDip, float rightDip) const noexcept;
    [[nodiscard]] const FolderViewColumnLayout::Column* TryGetColumnLayout(int column) const noexcept;
    ```
  - In `FolderView::LayoutItems()`, build a `std::vector<FolderViewColumnLayout::ItemTextMetrics>` from the same display-name, details, and metadata strings currently used for `maxLabelWidth`, `maxDetailsWidth`, and `maxMetadataWidth`.
  - Call `FolderViewColumnLayout::Resolve(...)` after item metrics, tile height, client size, spacing, and display mode are known.
  - Assign `_rowsPerColumn`, `_columns`, `_columnLayout`, `_maxColumnWidthDip`, `_tileWidthDip`, and `_contentWidth` from the helper result.
  - Lay out item bounds by iterating `_columnLayout`; each item receives `leftDip` from its column and `rightDip = leftDip + column.widthDip`.
  - Preserve existing empty-folder behavior, compact row spacing behavior, DPI-scaled constants, and scroll clamping.

- [x] Route text layout creation through item-local label width.
  - Change `UpdateItemTextLayouts(float labelWidth)` to either `UpdateItemTextLayouts()` or a name that reflects per-item widths.
  - Inside the loop, compute width with `GetItemLabelWidthDip(item)` from the item's current bounds, icon width, text gap, and padding.
  - Keep text layout invalidation tied to the same mode/theme/font/width changes as today.
  - Update all callers in `FolderView.Layout.cpp`.
  - In `FolderView.Rendering.cpp`, update `DrawItem()` so `EnsureItemTextLayout(item, labelWidth)` receives the item's own label width.

- [x] Route visible range, rendering, hit testing, and scrolling through variable columns.
  - In `GetVisibleItemRange()`, replace stride division with `GetVisibleColumnRange(viewLeft, viewRight)`.
  - Return the first item of the first visible column and the end item of the last visible column.
  - In the render loop in `FolderView.Rendering.cpp`, iterate visible columns from `_columnLayout` and then items in each column's `[startIndex, startIndex + itemCount)`.
  - In `HitTest()`, find the column whose `leftDip <= x < RightDip()` after horizontal-scroll adjustment.
  - Return no item when `x` is in column spacing or beyond the final column.
  - Compute row index from the column-local y coordinate and existing `tileHeight + rowSpacing` rule.
  - In `FolderView.Interaction.cpp`, update `OnHScrollMessage()` so line/page/thumb scrolling uses `_columnLayout` and real column boundaries instead of `_tileWidthDip + kColumnSpacingDip`.
  - In `EnsureVisible()`, use `item.bounds.left` and `item.bounds.right` instead of deriving x from `index / rowsPerColumn`.
  - Audit keyboard navigation and selection rectangle callers for assumptions that `column * columnStride` is valid; switch any such use to `_columnLayout[column].leftDip`.

- [x] Add perf instrumentation and archive evidence.
  - In `FolderView::LayoutItems()`, add `Debug::Perf::Scope layoutPerf(L"render.layout_items_us")`.
  - Set `layoutPerf.SetValue0(static_cast<int64_t>(_items.size()))`.
  - Set `layoutPerf.SetValue1(static_cast<int64_t>(_columnLayout.size()))` after resolving layout.
  - After implementation, run the candidate audit, desired behavior, and large-folder perf cases:
    ```powershell
    .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_column_widths_audit
    .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_visible_column_widths
    .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline
    ```
  - Compare baseline and candidate artifacts with existing perf tools:
    ```powershell
    $commandRuns = Get-ChildItem Specs\TestRuns -Directory |
        ForEach-Object { Get-ChildItem (Join-Path $_.FullName 'Commands') -Directory -ErrorAction SilentlyContinue } |
        Sort-Object LastWriteTime
    $baselineCommandsRun = $commandRuns[-2].FullName
    $candidateCommandsRun = $commandRuns[-1].FullName
    .\Tools\CompareTestRuns.ps1 $baselineCommandsRun $candidateCommandsRun -Suite Commands
    .\Tools\Show-PerfRuns.ps1 -CompareRun $baselineCommandsRun,$candidateCommandsRun -Scenario folderView_column_widths_audit
    .\Tools\Show-PerfRuns.ps1 -CompareRun $baselineCommandsRun,$candidateCommandsRun -Scenario folderView_perf_large_folder_baseline
    ```
  - Record in the implementation closeout whether `render.layout_items_us`, `render.frame_us`, `render.begin_to_enddraw_us`, and `render.present_us` stayed within the accepted range or improved.
  - Record the before and after `contentWidthDip` values for `column_poison_second_column`, `column_poison_every_other_column`, and `details_metadata_poison`.

- [x] Update durable docs and close the plan.
  - Update `Specs/UI/UI_FolderView.md` so the column calculation section states that FolderView first determines rows per column from the visible pane height, then computes each column width from only the items assigned to that column.
  - Mention Brief, Detailed, ExtraDetailed, and Thumbnails line participation in the spec.
  - Update `Specs/Testing/Testing_TestCoverage.md` with the new command selftest and pure layout test.
  - Update `Tests/README.md` if the named command selftest list is maintained there.
  - Move this file to `Specs/Plans/Done/UI_FolderViewVisibleColumnWidthsPlan_2026-05-14.md` only after implementation, tests, perf archives, and spec updates are complete.

## Verification Commands

```powershell
.\build.ps1 -ProjectName RedSalamander
.\build.ps1 -ProjectName PerformanceTests2
vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll /Tests:FolderViewColumnLayoutTests
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_column_widths_audit
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_visible_column_widths
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline
$commandRuns = Get-ChildItem Specs\TestRuns -Directory |
    ForEach-Object { Get-ChildItem (Join-Path $_.FullName 'Commands') -Directory -ErrorAction SilentlyContinue } |
    Sort-Object LastWriteTime
$baselineCommandsRun = $commandRuns[-2].FullName
$candidateCommandsRun = $commandRuns[-1].FullName
.\Tools\CompareTestRuns.ps1 $baselineCommandsRun $candidateCommandsRun -Suite Commands
.\Tools\Show-PerfRuns.ps1 -CompareRun $baselineCommandsRun,$candidateCommandsRun -Scenario folderView_column_widths_audit
.\Tools\Show-PerfRuns.ps1 -CompareRun $baselineCommandsRun,$candidateCommandsRun -Scenario folderView_perf_large_folder_baseline
```

## Rollback Plan

- The helper is additive and can remain tested even if integration is reverted.
- To revert product behavior, restore `FolderView::LayoutItems()`, `UpdateItemTextLayouts`, `GetVisibleItemRange`, `HitTest`, `EnsureVisible`, and render-loop callers to the previous global `_tileWidthDip` and `columnStride` path.
- Leave the spec update out of the closeout if the behavior is reverted before shipping.
