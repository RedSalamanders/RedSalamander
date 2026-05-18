# FolderView Performance Measurement Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the stale generic FolderView performance review with a measurement-first plan that proves which FolderView, IconCache, DirectoryInfoCache, and rendering paths are actually hot before any optimization lands.

**Architecture:** Establish same-machine baselines from existing FolderView selftests, add focused deterministic scenarios only where current coverage is too broad, and gate every optimization on archived `Specs/TestRuns/` evidence. The plan treats the earlier review findings as hypotheses, not facts.

**Tech Stack:** C++23, Win32, Direct2D/DirectWrite, WIL, `Debug::Perf`, `--commands-selftest`, `Specs/TestRuns/`.

---

## Checklist

- [x] Capture current same-machine baselines for existing FolderView performance cases before any production behavior change.
- [x] Record which earlier review findings are stale or already covered.
- [x] Add missing deterministic scenarios for cold open/icon churn, sort toggling, scroll/render invalidation, DirectoryInfoCache notification churn, and IconCache contention only where existing cases do not isolate the behavior.
- [x] Add instrumentation only when existing metric families cannot attribute the scenario.
- [x] Archive every baseline and candidate run under `Specs/TestRuns/`.
- [x] Decide optimizations from measured bottlenecks, not from the original checklist.
- [x] Update authoritative specs and move this plan to `Specs/Plans/Done/` when complete.
- [x] Follow-up: land a measured sort-path optimization with same-machine before/after evidence.
- [x] Follow-up: rerun sort correctness coverage after changing the sort algorithm.
- [x] Follow-up: land a measured inactive quick-search render guard with red/green same-machine evidence.
- [x] Follow-up: rerun active quick-search correctness and full FolderView perf aggregate after the render guard.
- [x] Follow-up: land a measured IconCache association-lock instrumentation optimization after repeated slow wait rows appeared in the final aggregate.
- [x] Follow-up: rerun the focused IconCache contention case and full FolderView perf aggregate after the IconCache instrumentation optimization.
- [x] Follow-up: validate the optimized candidate in a Release build with selftest hooks enabled and archive the focused plus aggregate runs.

## Ground Rules

- Do not implement any optimization until the scenario has a baseline archive and a metric showing the suspected cost.
- A finding is not actionable if it is already implemented, intentionally a sink parameter, outside the measured hot path, or below run-to-run noise.
- Prefer existing metrics:
  - `FolderView.ExecuteEnumeration.BuildItems`
  - `FolderView.ExecuteEnumeration.SortMerge`
  - `FolderView.ExecuteEnumeration.IconIndex.*`
  - `FolderView.ApplyCurrentSort`
  - `FolderView.IconLoading.ProcessQueue`
  - `FolderView.IconLoading.BatchUpdate`
  - `FolderView.IconLoading.BitmapConversion`
  - `FolderView.ThumbnailLoading.ProcessQueue`
  - `render.frame_us`
  - `render.layout_items_us`
  - `icons.*`
  - `thumbnails.*`
  - `iconcache.lock_wait_slow_us`
  - `iconcache.lock_hold_slow_us`
- Add new metrics only when the existing metrics cannot answer the decision.
- Use at least three same-machine samples for noisy UI scenarios before claiming a regression or win.
- Treat p95/max outliers as directional unless the artifact also records enough context to explain the work count.

## Rejected Stale Or Unproven Findings

The following items from the external review must not be implemented directly:

- `DirectoryInfoCache` per-character `towlower()` hash: current code already uses an ASCII fast path plus `LCMapStringEx` in `MakeCaseInsensitivePathKey(...)`.
- FolderView `_wcsicmp` sort replacement: current FolderView sorting uses `OrdinalString::Compare`, backed by `CompareStringOrdinal`.
- ColorTextView layout-cache LRU: current `ColorTextView` already promotes cache hits and evicts the oldest slice.
- Blanket `std::format` conversion for performance: this is a style/readability decision, not a proven speedup.
- FolderView theme brush batching: theme changes are not a repeated frame path unless a baseline proves otherwise.
- Pass-by-value sink constructors/factories such as `FolderWatcher(...)` and FileSystem7z reader/open callback constructors: these intentionally move into owned state unless a call-site metric proves a real copy hotspot.

## Measurement Scenarios

| Scenario | Existing Coverage | Measurement Decision |
|----------|-------------------|----------------------|
| Large folder cold open + icon churn | `folderView_perf_large_folder_baseline` | Re-run first. Add a narrower case only if artifact cannot separate enumeration, sort, icon index, icon queue, and first visible render cost. |
| Sort toggle stress | Partially covered by `FolderView.ApplyCurrentSort` metrics | Add focused case if existing baseline does not repeatedly toggle Name/Extension/Time/Size over a large adversarial set. |
| Scroll/render invalidation storm | Thumbnail scroll stress covers thumbnail mode; normal mode less isolated | Add focused normal-mode scroll/render case with frame/layout metrics and visible-work counts. |
| DirectoryInfoCache watcher notification churn | FileOps phases cover correctness; not clearly a FolderView perf case | Add scenario only if FolderView panes can receive rapid create/delete/move notifications and produce repeated refresh cost. |
| IconCache lock contention | Existing slow lock metrics exist | Reuse `iconcache.lock_*` metrics; add a deterministic contention driver only if normal baselines show slow lock rows. |
| Error/operation overlay invalidation | Current review calls it critical without evidence | Keep as low-priority smell; measure only if scroll/render or FileOps feedback artifacts show overlay churn. |
| Thumbnail mode regressions | Strong existing thumbnail cases | Do not duplicate; reuse `folderView_thumbnail_*` only as guardrail. |

## Task 1: Baseline Existing FolderView Cases

**Files:**
- Read: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Read: `Specs/Testing/Testing_TestCoverage.md`
- Update later with evidence: `Specs/Plans/WIP/FolderView_PerformanceMeasurementPlan_2026-05-15.md`

- [x] **Step 1: Build the current tree without production perf changes**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
```

Expected: build exits `0`. If the current unrelated `RedLauncher` solution issue blocks a full graph build, record the blocker and use the project-targeted `RedSalamander` build log.

- [x] **Step 2: List existing FolderView selftest cases**

Run:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-list-cases --selftest-case=folderView_
```

Expected: list includes `folderView_perf_large_folder_baseline`, `folderView_column_widths_audit`, and `folderView_thumbnail_*`.

- [x] **Step 3: Run large-folder baseline three times**

Run three samples:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4
```

Expected: each run passes and prints `ArchiveToRepo: Specs\TestRuns\<MachineHash>\Commands\<RunId>`.

- [x] **Step 4: Run existing layout and thumbnail guardrails**

Run:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_column_widths_audit --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_scroll_stress --selftest-timeout-multiplier=4
```

Expected: both pass and archive `perf/perf_metrics.jsonl` plus their JSON artifacts.

- [x] **Step 5: Record baseline evidence in this plan**

Add a `Baseline Evidence` section with:

- build log path,
- each `Specs/TestRuns/.../Commands/...` archive path,
- pass/fail counts,
- metric keys present in each `perf/perf_metrics.jsonl`,
- whether `iconcache.lock_wait_slow_us` or `iconcache.lock_hold_slow_us` appeared.

## Task 2: Add Metric Coverage Inventory

**Files:**
- Modify: `Specs/Plans/WIP/FolderView_PerformanceMeasurementPlan_2026-05-15.md`

- [x] **Step 1: Extract metric keys from each baseline**

Run for each archive:

```powershell
$path = 'Specs\TestRuns\<MachineHash>\Commands\<RunId>\perf\perf_metrics.jsonl'
Get-Content $path | ConvertFrom-Json | Group-Object metric | Sort-Object Name | Select-Object Name,Count
```

Expected: the inventory shows which scenario emits each FolderView/IconCache/render metric.

- [x] **Step 2: Mark each scenario as covered or needing a focused case**

Update the scenario table using these statuses:

- `covered`: existing metrics isolate the decision.
- `needs-focused-case`: existing run is too broad or lacks repeated stimulus.
- `needs-instrumentation`: no metric can answer the decision.
- `blocked`: the scenario cannot be made deterministic; state the exact blocker.

## Task 3: Focused Cold Open And Icon Churn Case

**Files:**
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify if inventory requires a new metric: `RedSalamander/FolderView.Enumeration.cpp`
- Modify if inventory requires a new metric: `RedSalamander/FolderView.Icons.cpp`
- Modify: `Specs/Testing/Testing_TestCoverage.md`

- [x] **Step 1: Decide whether the case is needed**

Use Task 2. Add this case only if `folderView_perf_large_folder_baseline` cannot distinguish enumeration build, sort merge, icon-index preparation, icon queue, and first visible render.

Decision: not needed. Task 2 shows the existing three-sample baseline emits separate enumeration, sort, icon-index, icon queue/batch, and render metric families for this gate.

- [x] **Step 2: Add selftest case `folderView_perf_cold_open_icon_churn`**

Skipped by the Step 1 gate. The existing large-folder baseline already separates the required stage metrics, so no new cold-open/icon-churn case was added.

Build a deterministic temp folder with:

- at least 2,000 files,
- repeated common extensions to exercise association cache hits,
- at least 100 unique extensions to exercise extension icon lookup,
- mixed directories and files,
- several long names and non-ASCII names to preserve correctness coverage.

The case must open the folder in a real pane, wait for enumeration and visible render settle, and write `folderView_perf_cold_open_icon_churn_metrics.json` with item counts, extension counts, visible rows, and extracted metric summary.

- [x] **Step 3: Run the new case before optimization**

Skipped by the Step 1 gate because the case was not needed.

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_cold_open_icon_churn --selftest-timeout-multiplier=4
```

Expected: pass, archive path printed, JSON artifact present.

## Task 4: Focused Sort Toggle Stress Case

**Files:**
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify if inventory requires a new metric: `RedSalamander/FolderView.Enumeration.cpp`
- Modify: `Specs/Testing/Testing_TestCoverage.md`

- [x] **Step 1: Add selftest case `folderView_perf_sort_toggle_stress`**

Use a deterministic folder with at least 5,000 entries and adversarial names:

- same prefix with numeric suffixes,
- mixed case variants,
- many repeated extensions,
- long extension-less names,
- non-ASCII names that must keep ordinal comparison semantics.

The case must repeatedly toggle sort by Name, Extension, Time, Size, and None, then restore the original view.

- [x] **Step 2: Capture sorting metrics**

Reuse `FolderView.ApplyCurrentSort` and `FolderView.ExecuteEnumeration.SortMerge`. If the artifact cannot attribute each sort request, add a scoped metric with detail text such as `name`, `extension`, `time`, `size`, or `none`.

- [x] **Step 3: Run the baseline**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_sort_toggle_stress --selftest-timeout-multiplier=4
```

Expected: pass, archive path printed, artifact records per-sort durations and final focused item correctness.

## Task 5: Normal-Mode Scroll And Render Invalidation Stress Case

**Files:**
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify if inventory requires a new metric: `RedSalamander/FolderView.Interaction.cpp`
- Modify if inventory requires a new metric: `RedSalamander/FolderView.Rendering.cpp`
- Modify if inventory requires a new metric: `RedSalamander/FolderView.Layout.cpp`
- Modify: `Specs/Testing/Testing_TestCoverage.md`

- [x] **Step 1: Add selftest case `folderView_perf_scroll_render_stress`**

Use a large normal-mode folder in Brief, Detailed, and Extra Detailed modes. Drive horizontal and vertical scrolling through real input/message paths, not by mutating offsets directly.

- [x] **Step 2: Reuse or add metrics**

Reuse:

- `render.frame_us`
- `render.layout_items_us`
- `render.draw_item_us`

Add metrics only if needed:

- `folder.scroll_input_to_paint_us`
- `folder.scroll_visible_item_count`
- `folder.scroll_frame_count`
- `folder.scroll_full_invalidate_count`

- [x] **Step 3: Run baseline**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
```

Expected: pass, archive path printed, artifact records per-mode frame count, max frame time, and visible item count.

## Task 6: DirectoryInfoCache Notification Churn Scenario

**Files:**
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify if inventory requires a new metric: `RedSalamander/DirectoryInfoCache.cpp`
- Modify: `Specs/Testing/Testing_TestCoverage.md`

- [x] **Step 1: Decide whether this belongs in Commands or FileOperations**

Use `--commands-selftest` if the scenario is pane-visible refresh responsiveness. Use `--fileops-selftest` if it only exercises cache correctness under file-operation mutations.

Decision: Commands. The new case opens a real FolderView pane, mutates the visible local folder externally, waits for pane refresh settle, and verifies final visible item count plus focus stability.

- [x] **Step 2: Add selftest case `folderView_perf_directory_change_storm` if pane-visible behavior is measurable**

Create a local folder, open it in FolderView, then apply a deterministic burst:

- create 200 files,
- rename 100 files,
- delete 100 files,
- create and delete 20 directories.

Wait for FolderView to settle and verify final visible item count and focus stability.

- [x] **Step 3: Reuse or add metrics**

Reuse existing enumeration metrics. Add cache metrics only if needed:

- `directorycache.notify_us`
- `directorycache.mark_dirty_us`
- `directorycache.post_refresh_count`
- `directorycache.subtree_dirty_count`

- [x] **Step 4: Run baseline**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_directory_change_storm --selftest-timeout-multiplier=4
```

Expected: pass, archive path printed, final folder state correct, no unbounded refresh loop.

## Task 7: IconCache Lock Contention Decision

**Files:**
- Modify only if baseline shows slow lock rows: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify only if baseline shows slow lock rows without attribution: `RedSalamander/IconCache.cpp`
- Modify: `Specs/Testing/Testing_TestCoverage.md`

- [x] **Step 1: Inspect existing baselines for slow lock rows**

Run:

```powershell
Get-Content Specs\TestRuns\<MachineHash>\Commands\<RunId>\perf\perf_metrics.jsonl |
    ConvertFrom-Json |
    Where-Object { $_.metric -in @('iconcache.lock_wait_slow_us', 'iconcache.lock_hold_slow_us') } |
    Select-Object metric,detail,durationUs,value0,value1
```

Expected: if no rows appear in three baseline samples, do not add an IconCache contention optimization.

- [x] **Step 2: Add focused contention driver only when rows appear**

If slow lock rows appear, add `folderView_perf_iconcache_contention` to open multiple panes or repeated icon-heavy folders and record lock wait/hold rows by detail.

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_iconcache_contention --selftest-timeout-multiplier=4
```

Expected: pass and archive. Candidate optimization work starts only if the baseline has repeated slow lock rows in the same detail stage.

## Task 8: Candidate Optimization Gate

**Files:**
- Modify only after measured evidence points to a bottleneck:
  - `RedSalamander/FolderView.Rendering.cpp`
  - `RedSalamander/FolderView.Layout.cpp`
  - `RedSalamander/FolderView.Enumeration.cpp`
  - `RedSalamander/FolderView.Icons.cpp`
  - `RedSalamander/DirectoryInfoCache.cpp`
  - `RedSalamander/IconCache.cpp`
  - `RedSalamander/FolderView.ErrorOverlay.cpp`

- [x] **Step 1: Create a measured bottleneck table**

For each candidate, record:

- baseline archive,
- metric key,
- sample count,
- median, p95, and max if available,
- candidate file/function,
- correctness risk,
- expected measurable direction.

- [x] **Step 2: Reject unproven optimizations**

Reject changes that only improve code aesthetics or theoretical cost. Examples: blanket `std::format`, sink-parameter rewrites, theme brush batching, and duplicate string-view conversions.

- [x] **Step 3: Implement one optimization per metric**

For each chosen optimization:

- write or extend the deterministic selftest first,
- run it as baseline and confirm the metric captures the current cost,
- implement the smallest production change,
- rerun the same test,
- compare same-machine baseline and candidate archives,
- keep the change only if correctness still passes and the metric improves or the trade-off is explicitly accepted.

Measured bottleneck table:

| Candidate | Baseline archive | Metric key(s) | Sample count / summary | Decision |
|-----------|------------------|---------------|------------------------|----------|
| Sort-path optimization | `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191333`, `2026-05-15_192148`, `2026-05-15_193019` | `folder.sort_toggle_us`, `FolderView.ApplyCurrentSort`, `FolderView.ExecuteEnumeration.SortMerge` | Baseline: 3 runs / 30 toggles; `folder.sort_toggle_us` median 278434 us, avg 290109 us, p95 526892 us; `FolderView.ApplyCurrentSort` median 132061 us, avg 139284 us, p95 268214 us | Implemented: keep medium folders on sequential `std::sort`, reserve parallel sort for 20000+ items. Final same-machine runs `2026-05-15_193548`, `2026-05-15_193637`, `2026-05-15_193730`: sort-toggle median 250308 us (-10.1%), avg 259959 us (-10.4%), p95 359338 us (-31.8%); apply-sort median 37563 us (-71.6%), avg 36294 us (-73.9%), p95 62527 us (-76.7%). |
| Inactive quick-search render guard | Pre-guard optimized-sort archives `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_193548`, `2026-05-15_193637`, `2026-05-15_193730` | `render.incremental_search_effect_updates`, `render.frame_us`, `render.begin_to_enddraw_us`, `render.draw_item_us` | Baseline: 3 runs / 126 rendered frames; inactive incremental-search effect updates 6327 over 6327 drawn items; `render.frame_us` median 44600 us, avg 45558 us; `render.draw_item_us` avg 132 us | Implemented: skip incremental-search drawing-effect clearing and match scans unless quick search is active with a non-empty query. Candidate same-machine runs `2026-05-15_200406`, `2026-05-15_200523`, `2026-05-15_200551`: inactive updates 0 (-100%); `render.frame_us` median 34736 us (-22.1%), avg 40205 us (-11.7%); `render.begin_to_enddraw_us` median 29255 us (-23.5%), avg 33053 us (-14.2%); `render.draw_item_us` avg 85 us (-35.6%). Full `folder.sort_toggle_us` was noisy as a set (+7.6% median, +2.1% avg), so this is accepted as a deterministic render-frame improvement, not an end-to-end sort-toggle win. |
| Normal-mode scroll/render optimization | `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191356` | `folder.scroll_input_to_paint_us`, `folder.scroll_frame_count`, `folder.scroll_visible_item_count`, `render.*` | 24 scroll steps; artifact median 78068 us, p95 129075 us, max 181503 us | No production change in this plan. The run provides focused attribution for future measured deltas. |
| DirectoryInfoCache notification optimization | `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191413` | `folder.directory_change_storm_mutation_us`, `folder.directory_change_storm_settle_us`, `directorycache.post_refresh_count` | mutation 441169 us, settle 111201 us, final 101 items, focus stable, no unbounded loop | Reject optimization for now; no repeated cache bottleneck was isolated beyond the new baseline. |
| IconCache contention optimization | `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191805` plus initial large-folder samples | `iconcache.lock_wait_slow_us`, `iconcache.lock_hold_slow_us`, `folder.iconcache_contention_cycle_us` | contention driver: 3 cycles, 11 slow hold rows, 0 slow wait rows; large-folder samples: 0 slow lock rows | Initial decision rejected production changes because evidence showed hold-only rows, not measured wait contention. Reopened later when the final aggregate produced repeated association slow-wait rows; see the next row. |
| IconCache association-lock instrumentation optimization | Final pre-IconCache-change aggregate `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_201709`; focused pre-change runs `2026-05-15_202546`, `2026-05-15_202747`, `2026-05-15_202948` | `iconcache.lock_wait_slow_us`, `folder.iconcache_contention_cycle_us` | Final aggregate showed repeated slow wait rows: 350 `iconcache.lock_wait_slow_us` rows totaling 16583.9 ms, dominated by `association_lookup` and `association_store`. Focused pre-change set was noisy but reproduced a first-run storm: 288 association wait rows totaling 18104.1 ms. | Implemented: emit association-cache lock wait/hold diagnostics after the association mutex is released. Final aggregate `2026-05-15_203429`: wait rows 350 -> 9 (-97.4%), wait sum 16583.9 ms -> 4.2 ms (-99.97%), IconCache contention case duration 5398 ms -> 3592 ms (-33.5%), and contention-cycle sum 3142.3 ms -> 1840.6 ms (-41.4%). |
| External-review stale items | baseline/inventory sections above | n/a | n/a | Reject blanket `std::format`, sink-parameter rewrites, theme brush batching, `_wcsicmp` sort replacement, ColorTextView LRU, and DirectoryInfoCache hash rewrite as stale or unproven. |

The production optimizations kept by this plan are the measured FolderView sort-path change, the inactive quick-search render guard, and the IconCache association-lock instrumentation change. Other candidates remain rejected or deferred until their focused metrics show a same-machine improvement target.

## Task 9: Closeout

**Files:**
- Modify: `Specs/UI/UI_FolderView.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Move: `Specs/Plans/WIP/FolderView_PerformanceMeasurementPlan_2026-05-15.md` to `Specs/Plans/Done/FolderView_PerformanceMeasurementPlan_2026-05-15.md`

- [x] **Step 1: Update authoritative FolderView performance contract**

Add any newly adopted scenarios and metric families to `Specs/UI/UI_FolderView.md`. Do not move speculative or rejected findings into the authoritative spec.

- [x] **Step 2: Update test inventory**

Add any new selftest cases and artifacts to `Specs/Testing/Testing_TestCoverage.md`.

- [x] **Step 3: Record final evidence**

Add closeout notes to this plan with:

- baseline archive paths,
- candidate archive paths,
- same-machine status,
- changed metrics,
- rejected optimizations and reasons,
- remaining blocked items with exact blocker.

- [x] **Step 4: Move the plan**

Move this file to `Specs/Plans/Done/` after the evidence and spec updates are complete.

## Baseline Evidence

Captured on same machine hash `4cb089111a23` before any production behavior change.

Build:

- `.build/logs/msbuild-20260515_185854_263.log` — `.\build.ps1 -ProjectName RedSalamander -Configuration Debug`, exit `0`, 0 warnings, 0 errors.

Case inventory:

- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-list-cases --selftest-case=folderView_` listed 16 Commands cases, including `folderView_perf_large_folder_baseline`, `folderView_column_widths_audit`, and the `folderView_thumbnail_*` family.

Archived runs:

| Run | Case | Result | Artifact | Relevant metric keys present | Slow IconCache lock rows |
|-----|------|--------|----------|------------------------------|--------------------------|
| `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_190139` | `folderView_perf_large_folder_baseline` | 1 passed / 0 failed / 0 skipped, 10754 ms | `perf/folderView_perf_large_folder_baseline_metrics.json` | `DirectoryInfoCache.ReadDirectoryInfo`; `FileSystem.DirectoryOps.Enumerate`; `FileSystem.DirectoryOps.TrimBuffer`; `folder.selftest.render_warmup_us`; `FolderView.ApplyCurrentSort`; `FolderView.ExecuteEnumeration.BuildItems`; `FolderView.ExecuteEnumeration.IconIndex.Prepare`; `FolderView.ExecuteEnumeration.IconIndex.QueryExtensions`; `FolderView.ExecuteEnumeration.SortMerge`; `FolderView.IconLoading.BatchUpdate`; `FolderView.IconLoading.BitmapConversion`; `FolderView.IconLoading.ProcessQueue`; `IconCache.WarmCommonExtensions`; `iconcache.convert_create_d2d_bitmap_us`; `iconcache.convert_draw_icon_us`; `iconcache.shgetfileinfo_us`; `icons.*`; `render.*` | none |
| `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_190204` | `folderView_perf_large_folder_baseline` | 1 passed / 0 failed / 0 skipped, 10887 ms | `perf/folderView_perf_large_folder_baseline_metrics.json` | same as `2026-05-15_190139` | none |
| `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_190215` | `folderView_perf_large_folder_baseline` | 1 passed / 0 failed / 0 skipped, 10594 ms | `perf/folderView_perf_large_folder_baseline_metrics.json` | same as `2026-05-15_190139` | none |
| `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_190351` | `folderView_column_widths_audit` | 1 passed / 0 failed / 0 skipped, 87789 ms | `perf/folderView_column_widths_audit_metrics.json` | large-folder families plus `FolderView.ThumbnailLoading.ProcessQueue`; `iconcache.get_bitmap_hit`; `iconcache.get_bitmap_miss`; `iconcache.miss_convert_us`; `iconcache.miss_extract_us`; `iconcache.miss_store_us`; `render.folderViewColumnAudit.us`; `thumbnails.*` | 51 `iconcache.lock_hold_slow_us`; 0 `iconcache.lock_wait_slow_us` |
| `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_190406` | `folderView_thumbnail_scroll_stress` | 1 passed / 0 failed / 0 skipped, 1830 ms | `perf/folderView_thumbnail_scroll_stress_metrics.json` | large-folder families plus `FolderView.ThumbnailLoading.ProcessQueue`; `thumbnails.*` | none |
| `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191333` | `folderView_perf_sort_toggle_stress` | 1 passed / 0 failed / 0 skipped, 12802 ms | `perf/folderView_perf_sort_toggle_stress_metrics.json` | `folder.sort_toggle_us`; `FolderView.ApplyCurrentSort`; `FolderView.ExecuteEnumeration.SortMerge`; `render.*`; `icons.*` | 4 `iconcache.lock_hold_slow_us`; 0 `iconcache.lock_wait_slow_us` |
| `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191356` | `folderView_perf_scroll_render_stress` | 1 passed / 0 failed / 0 skipped, 5833 ms | `perf/folderView_perf_scroll_render_stress_metrics.json` | `folder.scroll_input_to_paint_us`; `folder.scroll_frame_count`; `folder.scroll_visible_item_count`; `render.*`; `icons.*` | 6 `iconcache.lock_hold_slow_us`; 0 `iconcache.lock_wait_slow_us` |
| `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191413` | `folderView_perf_directory_change_storm` | 1 passed / 0 failed / 0 skipped, 718 ms | `perf/folderView_perf_directory_change_storm_metrics.json` | `folder.directory_change_storm_mutation_us`; `folder.directory_change_storm_settle_us`; `directorycache.post_refresh_count`; `FolderView.ExecuteEnumeration.*` | none |
| `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191805` | `folderView_perf_iconcache_contention` | 1 passed / 0 failed / 0 skipped, 4021 ms | `perf/folderView_perf_iconcache_contention_metrics.json` | `folder.iconcache_contention_cycle_us`; `iconcache.lock_*`; `FolderView.IconLoading.*`; `icons.*`; `render.*` | 11 `iconcache.lock_hold_slow_us`; 0 `iconcache.lock_wait_slow_us` |

Custom artifact notes:

- Large-folder baseline artifact records `itemCount=241`, `bitmapIconCount=241`, `enumerationCount=3`, `warmRenderInvocations=5`, `renderCalls=20`, `queueIconLoadingCalls=12`, `processIconQueueCalls=0`, and `batchIconUpdateCalls=5`.
- Thumbnail scroll stress artifact records `itemCount=640`, `thumbnailScrollStress.elapsedUs=722602`, `thumbnails.queued=30`, `thumbnails.completed=30`, `thumbnails.fallback=30`, `thumbnails.pending=0`, and `thumbnails.staleDrops=28`.
- Sort-toggle artifact records `itemCount=5000`, 10 sort operations, median `272724 us`, p95/max `391983 us`.
- Scroll/render artifact records `itemCount=1600`, 24 scroll steps, median `78068 us`, p95 `129075 us`, max `181503 us`.
- Directory-change artifact records `finalItemCount=101`, `mutationUs=441169`, `settleUs=111201`, `refreshDelta=0`, `enumerationCount=2`, and stable focus on `focus_anchor.txt`.
- IconCache contention artifact records 3 cycles across two 480-item panes with 160 unique extensions per pane; the run produced hold-only slow lock diagnostics and no wait rows.

## Metric Coverage Inventory

| Scenario | Status | Evidence / next decision |
|----------|--------|--------------------------|
| Large folder cold open + icon churn | `covered` | Three same-machine `folderView_perf_large_folder_baseline` samples emit separate enumeration, sort, icon-index, icon queue/batch, and render metric families. The artifact is not a 2,000-item churn fixture, but Task 3's gate is stage attribution, and the existing artifact can answer that decision. |
| Sort toggle stress | `covered` | Added `folderView_perf_sort_toggle_stress`, which repeatedly toggles Name/Extension/Time/Size/None over a 5,000-entry adversarial folder and emits `folder.sort_toggle_us` alongside existing sort metrics. |
| Scroll/render invalidation storm | `covered` | Added `folderView_perf_scroll_render_stress`, which drives normal-mode horizontal and vertical scroll messages across Brief, Detailed, and Extra Detailed modes and emits `folder.scroll_*` plus `render.*` metrics. |
| DirectoryInfoCache watcher notification churn | `covered` | Added `folderView_perf_directory_change_storm`, which opens a real pane, applies deterministic create/rename/delete/directory churn, and verifies final count/focus stability with `folder.directory_change_storm_*` metrics. |
| IconCache lock contention | `covered` | Added `folderView_perf_iconcache_contention` because slow hold rows appeared outside the three large-folder samples. The first focused driver produced 11 slow hold rows and 0 slow wait rows, so production changes were initially rejected; later final aggregate evidence produced repeated association wait rows and justified the association-lock instrumentation optimization. |
| Error/operation overlay invalidation | `blocked` | No current baseline showed overlay churn. This remains blocked until a scroll/render or FileOps-feedback artifact identifies a deterministic overlay invalidation stimulus. |
| Thumbnail mode regressions | `covered` | `folderView_thumbnail_scroll_stress` and the existing `folderView_thumbnail_*` family are the guardrail; no duplicate scenario is needed for this plan. |

## Closeout Notes

Authoritative specs updated:

- `Specs/UI/UI_FolderView.md` now lists the adopted FolderView perf scenarios and metric families.
- `Specs/Testing/Testing_TestCoverage.md` now lists `folderView_perf_sort_toggle_stress`, `folderView_perf_scroll_render_stress`, `folderView_perf_directory_change_storm`, and `folderView_perf_iconcache_contention`; runner-native Commands count is `632`, static Commands fallback count is `602`.

Final evidence:

- Baseline build before production behavior changes: `.build/logs/msbuild-20260515_185854_263.log`.
- Build after focused selftest additions: `.build/logs/msbuild-20260515_191119_392.log`.
- Build after IconCache contention driver addition: `.build/logs/msbuild-20260515_191618_618.log`.
- Build after sort-path optimization: `.build/logs/msbuild-20260515_193350_404.log`.
- Build after inactive quick-search render guard: `.build/logs/msbuild-20260515_200200_634.log`.
- Final build after quick-search stale-effect exit cleanup: `.build/logs/msbuild-20260515_201427_944.log`.
- Build after IconCache association-lock instrumentation optimization: `.build/logs/msbuild-20260515_203118_903.log`.
- Standard Release build before selftest validation: `.build/logs/msbuild-20260515_223252_366.log`, 0 warnings, 0 errors; Release selftests are compiled out by default.
- Release-with-selftest-hooks rebuild for validation: `.build/logs/msbuild-20260515_223506_104.log`, 1 warning (`C4883` in selftest-only `FolderWindow.FileOperations.SelfTest.cpp`), 0 errors.
- Standard Release rebuild after validation: `.build/logs/msbuild-20260515_223929_547.log`, 0 warnings, 0 errors; this restores the normal non-selftest Release binary in `.build/x64/Release/`.
- Existing baseline/guardrail archives: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_190139`, `2026-05-15_190204`, `2026-05-15_190215`, `2026-05-15_190351`, `2026-05-15_190406`.
- New baseline archives: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191333`, `2026-05-15_191356`, `2026-05-15_191413`, `2026-05-15_191805`.
- Measurement-coverage aggregate `folderView_perf_` archive before the production sort optimization: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_192148` with 5 passed / 0 failed / 0 skipped.
- Sort optimization baseline archives: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_191333`, `2026-05-15_192148`, `2026-05-15_193019`.
- Sort optimization candidate archives: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_193548`, `2026-05-15_193637`, `2026-05-15_193730`.
- Sort correctness guard archive: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_193810` with `cmd_pane_navigation_sort_modes_keep_navigation_shell_stable` passed.
- Optimized full `folderView_perf_` aggregate archive: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_194003` with 5 passed / 0 failed / 0 skipped; its sort case reported `folder.sort_toggle_us` median 265460 us and `FolderView.ApplyCurrentSort` median 34976 us.
- Inactive quick-search render guard RED archive: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_200144`; `folderView_perf_sort_toggle_stress` failed because normal sort rendering performed 171 inactive incremental-search drawing-effect updates on the first toggle.
- Inactive quick-search render guard candidate archives: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_200406`, `2026-05-15_200523`, `2026-05-15_200551`.
- Final active quick-search correctness guard archive: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_201626` with `cmd_pane_quickSearch_integrated_navigation` passed.
- Final full `folderView_perf_` aggregate archive after the render guard and stale-effect cleanup: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_201709` with 5 passed / 0 failed / 0 skipped; its sort case recorded `incrementalSearchEffectUpdates=0` for every toggle.
- IconCache association-lock focused baseline archives: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_202546`, `2026-05-15_202747`, `2026-05-15_202948`; all passed `folderView_perf_iconcache_contention`.
- IconCache association-lock focused candidate archives: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_203307`, `2026-05-15_203312`, `2026-05-15_203318`; all passed `folderView_perf_iconcache_contention`.
- Final full `folderView_perf_` aggregate archive after the IconCache instrumentation optimization: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_203429` with 5 passed / 0 failed / 0 skipped.
- Release-with-selftest-hooks focused validation archives: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_223732` (`folderView_perf_sort_toggle_stress` passed), `2026-05-15_223737` (`folderView_perf_iconcache_contention` passed), and `2026-05-15_223745` (`folderView_perf_scroll_render_stress` passed).
- Release-with-selftest-hooks aggregate validation archive: `Specs/TestRuns/4cb089111a23/Commands/2026-05-15_223838` with 5 passed / 0 failed / 0 skipped.
- Same-machine status: all listed archives are on machine hash `4cb089111a23`.
- Candidate change: `FolderView::ApplyCurrentSort` now uses `std::sort` with the existing total comparator and raises the parallel-sort cutoff from 1000 to 20000 items.
- Candidate change: `FolderView::DrawItem` now skips incremental-search drawing-effect clearing and match scans unless quick search is active with a non-empty query.
- Candidate change: `FolderView::ExitIncrementalSearch` clears existing DirectWrite layout effects once on quick-search exit so the render guard cannot leave stale highlighted text behind.
- Candidate change: `QueryAssociationIconIndex` now records association-cache lock wait/hold durations while holding the mutex but emits the slow-lock perf rows only after releasing it.
- Candidate result: across 3 same-machine baseline runs and 3 final sort-path runs, `FolderView.ApplyCurrentSort` median improved from 132061 us to 37563 us (-71.6%) and average improved from 139284 us to 36294 us (-73.9%). The full forced `folder.sort_toggle_us` median improved from 278434 us to 250308 us (-10.1%) and average from 290109 us to 259959 us (-10.4%).
- Candidate result: across 3 same-machine pre-guard and 3 post-guard runs, inactive incremental-search effect updates improved from 6327 to 0 (-100%), `render.frame_us` median improved from 44600 us to 34736 us (-22.1%), and `render.draw_item_us` average improved from 132 us to 85 us (-35.6%). The full forced `folder.sort_toggle_us` set was noisy (+7.6% median), so the accepted win is the deterministic render-frame reduction.
- Candidate result: comparing final full aggregates `2026-05-15_201709` -> `2026-05-15_203429`, IconCache association lock wait diagnostics dropped from 350 rows / 16583.9 ms to 9 rows / 4.2 ms, `folderView_perf_iconcache_contention` duration dropped from 5398 ms to 3592 ms (-33.5%), and its three-cycle metric sum dropped from 3142.3 ms to 1840.6 ms (-41.4%).
- Release validation result: aggregate `2026-05-15_223838` confirmed `FolderView.ApplyCurrentSort` is no longer a Release hotspot (35 rows, median 369 us, avg 753 us), inactive quick-search effect updates stayed at 0 across 197 aggregate render frames, and no `iconcache.lock_wait_slow_us` rows appeared in the Release aggregate. Release aggregate case durations were: large-folder baseline 9889 ms, sort-toggle stress 10847 ms, scroll/render stress 6426 ms, directory-change storm 642 ms, and IconCache contention 3762 ms.

Changed metrics / new selftest-local metric rows:

- `folder.sort_toggle_us`
- `render.incremental_search_effect_updates`
- `iconcache.lock_wait_slow_us`
- `iconcache.lock_hold_slow_us`
- `folder.scroll_input_to_paint_us`
- `folder.scroll_frame_count`
- `folder.scroll_visible_item_count`
- `folder.directory_change_storm_mutation_us`
- `folder.directory_change_storm_settle_us`
- `folder.iconcache_contention_cycle_us`
- `directorycache.post_refresh_count`

Rejected or deferred optimizations:

- No additional render, DirectoryInfoCache, or overlay optimization landed because no candidate met the same-machine before/after evidence gate.
- No deeper IconCache cache-design change landed; only the measured association-lock instrumentation critical-section fix is accepted.
- External-review findings listed in the rejected section remain stale or unproven and must not be implemented directly.

Remaining blocked item:

- Error/operation overlay invalidation remains blocked until a deterministic scroll/render or FileOps-feedback artifact shows overlay churn.

## Self-Review

- Spec coverage: The plan covers the original review's FolderView, DirectoryInfoCache, IconCache, rendering, string-parameter, sorting, and invalidation claims as measurement hypotheses.
- Unfinished-marker scan: No unresolved implementation markers remain.
- Scope check: The plan is focused on FolderView-adjacent performance and explicitly excludes unrelated style cleanups.
- Evidence gate: Every optimization path requires a baseline archive before production code changes.
