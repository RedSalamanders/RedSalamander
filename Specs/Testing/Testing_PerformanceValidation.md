# Performance Validation Specification

## Overview

Performance work in RedSalamander is a **normative engineering requirement**, not an optional cleanup step.

Any new feature, behavior change, or optimization that can affect:

- UI responsiveness,
- startup time,
- folder enumeration,
- rendering,
- search,
- Compare Directories,
- File Operations,
- plugin I/O,
- memory retention,
- background queueing,

MUST define its validation path up front and MUST ship with measurable coverage from the beginning.

Related documents:

- `Specs/Testing/Testing_SelfTests.md`
- `Specs/TestRuns/README.md`
- `Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md`

## Required Development Contract

### 1. Every perf-sensitive change MUST name its scenario

Before implementation is considered complete, the owning change MUST identify:

- the user-visible scenario being protected,
- the primary metric or metric family,
- the authoritative selftest or deterministic repro path,
- the expected direction of change.

Examples:

- “Find Files progress storm while results are visible”
- “Compare Directories content-progress queue churn”
- “Cold folder open with repeated icon indices”
- “FileOps pre-calc cancel latency”
- “Parsed ViewerText diff open, side-by-side scroll repaint, unchanged-text expansion, range-bounded referenced-file hydration, viewport rehydration, and unresolved placeholder fallback”

### 2. Every perf-sensitive change MUST have measurable evidence

A change MUST satisfy one of these:

- extend an existing metric family, or
- add new instrumentation, or
- explicitly justify why existing metrics already cover the scenario.

For optimizations, “seems faster” is not sufficient. The change MUST be supported by archived measurements or by a documented blocked reason.

### 3. New features MUST integrate tests and perf measurement from the start

When a new feature introduces a new hot path, queue, async pipeline, render surface, large-list path, or repeated callback flow:

- a deterministic selftest or deterministic repro harness MUST be added with the feature,
- at least one performance metric relevant to that path MUST be emitted with the feature,
- the expected archive path under `Specs/TestRuns/` MUST be part of the validation plan for the feature.

This requirement applies even when the first landing only establishes a baseline.

### 4. Optimizations MUST preserve correctness coverage

An optimization is incomplete if it improves a metric but lacks behavioral protection.

Any performance change MUST keep or add:

- correctness selftests,
- perf instrumentation,
- archived run evidence when the scenario is runnable on the current machine.

### 5. Claims of improvement MUST be grounded

A claimed performance improvement MUST state:

- the baseline run,
- the candidate run,
- whether the comparison is same-machine and same-suite,
- the metrics that improved,
- any material caveats.

If same-machine evidence is unavailable, the claim MUST be labeled directional rather than definitive.

## Perf Gate Template

Every perf-sensitive PR MUST include this gate in the PR description, review notes, or linked closeout spec before merge:

- Scenario: the exact user-visible path being protected.
- Subsystem and change type: feature, optimization, stabilization, regression-fix, or instrumentation-only.
- User-visible risk protected: the latency, throughput, responsiveness, queueing, memory, or correctness/perf interaction being guarded.
- Metric keys: the emitted metric family or the existing metrics that cover the path, including units and sample grain.
- Instrumentation: existing instrumentation reused or new instrumentation added.
- Deterministic validation: the selftest, focused harness, or deterministic repro that exercises the path, including exact command, case filter, timeout multiplier, perf budget path when applicable, and required environment variables.
- Build flavor: Debug diagnostic, test-enabled Release, or another explicitly justified flavor. Final throughput, latency, percentile, or budget claims require Release evidence unless the owning spec allows otherwise.
- Environment/root matrix: machine hash, CPU/load notes when material, `REDSALAMANDER_TEST_ROOT` override if any, filesystem capability reason, and whether the run is final evidence or diagnostic only. Path-sensitive local evidence that cannot use the workspace default should use `REDSALAMANDER_TEST_ROOT=C:\RSPerf`; `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox` and `C:\RST` must not be used for final path-sensitive evidence.
- Archived evidence: the `Specs/TestRuns/<MachineHash>/<Area>/<RunId>/` folder containing `results.json`, `trace.txt`, `run-all-tests-results.json` when the runner created one, and `perf_metrics.jsonl` when metrics are emitted.
- Analyzer and sample quality: the `Tools/Show-PerfRuns.ps1` command or owning analyzer, `-FailOnQuality` status for percentile claims, and explicit sample count sufficiency.
- Before/after delta: baseline run, candidate run, same-machine status, same-suite status, changed metric values, and caveats.
- Authoritative spec updates: the owning durable spec or guidance files that were updated, or the exact blocker if the update is deferred.

If any field is blocked, the PR MUST state the blocker explicitly and identify the follow-up owner/task. A perf-sensitive PR is not complete with only manual timing notes or unarchived terminal output.

## Instrumentation Rules

### Metric design

Metrics SHOULD be:

- scenario-specific,
- deterministic enough for repeated selftest use,
- attributable to one stage of work,
- named consistently with the owning subsystem.

Examples:

- `find.ui.*`
- `compare.ui.*`
- `FileOps.*`
- `render.*`
- `icons.*`
- `viewer.diff.*`
Current ViewerText diff baselines include `viewer.diff.open_to_first_visible_us`, `viewer.diff.visible_rows`, `viewer.diff.semantic_row_paint_us`, `viewer.diff.visible_styled_rows`, `viewer.diff.visible_context_rows`, `viewer.diff.visible_banner_rows`, `viewer.diff.visible_split_rows`, `viewer.diff.theme_switch_repaint_us`, `viewer.diff.scroll_repaint_us`, `viewer.diff.hunk_jump_to_visible_us`, `viewer.diff.expand_context_us`, `viewer.diff.viewport_rehydrate_us`, `viewer.diff.viewport_backtrack_us`, `viewer.diff.deferred_rows`, `viewer.diff.referenced_bytes_read`, `viewer.diff.viewport_referenced_bytes_read`, `viewer.diff.viewport_referenced_bytes_delta`, `viewer.diff.viewport_backtrack_referenced_bytes_delta`, `viewer.diff.placeholder_rows`, `viewer.diff.placeholder_bands`, and `viewer.chrome.paint_us`.
The `viewer_text_diff_perf` artifact also records semantic-row paint timing, context-row visibility, pane-local side-by-side layout activation, visible split-row counts, pane column widths, rainbow-mode theme-switch repaint timing, root-shell chrome repaint timing, hunk-jump latency, and expand-time, post-viewport, and post-backtrack referenced-byte counts so clickable hidden-banner reveal, calmer base-background side-by-side structure, combo-sync behavior on scroll repaint, root-shell typography changes, horizontal-scroll presentation switches, bounded referenced-file growth, and cache reuse can be reviewed alongside the metric stream.
Preserve-viewport diff rehydrate/backtrack rebuilds should reuse the existing section-combo model and avoid full combo repopulation or header relayout unless the visible section set actually changes.

### Preferred measurements

When applicable, instrument:

- queue depth or queue drain behavior,
- message coalescing and skipped work,
- input-to-visible or progress-to-visible latency,
- rebuild/repaint counts,
- bytes processed or rewritten,
- lock hold time,
- pre-calc / sort / refresh stage timings,
- bounded visible-work counters for large lists or grids,
- DirectWrite/glyph text-layout creation counts and timing (the `dwrite.text_layout.*` family in FolderView; see `Specs/UI/UI_FolderView.md`). 2026-06-19 same-machine evidence found FolderView text-layout creation non-material (about 1.2–1.4% of `render.layout_items_us`), so this family is reusable measurement infrastructure and not, by itself, a justification to add a text-layout cache without fresh evidence.
- Per-phase decomposition of a hot pass into named sub-stage timings (e.g. the `folder.layout.*_us` family decomposes `render.layout_items_us` in FolderView; see `Specs/UI/UI_FolderView.md`). 2026-06-19 this decomposition exposed that an apparent FolderView layout "bottleneck" (~82.6% in `UpdateItemTextLayouts`) was ~96% a *measurement artifact* — the JSONL sink opening/closing the file per metric row — not real work; the sink was fixed to keep the handle open (`Common/Common/PerfJsonl.cpp`). A worked example of why you decompose before optimizing: the dominant cost was outside the suspected mechanism.

### Anti-patterns

Avoid relying only on:

- ad hoc logging with no archived metric,
- one-off manual stopwatch measurements,
- broad “total duration” metrics when the decision requires stage attribution.
- repeated hot-path rows that only report pointer coordinates and sub-millisecond callback duration, such as per-hit-test/per-pointer-move traces, when the user-visible scenario needs input-to-visible latency, queue counts, scroll-apply cost, or paint cost instead.
- high-frequency per-item or per-creation emits that distort the very pass that hosts them. The JSONL sink keeps its file handle open (fixed 2026-06-19, `Common/Common/PerfJsonl.cpp`, so a row is a single `WriteFile` rather than an open/close), but each row still has a cost and grows the file, so coalesce hot-path telemetry into per-pass/per-frame aggregates (cf. the FolderView `dwrite.text_layout.frame_create_*` pattern).
- file-backed selftest trace writes inside the measured render, draw-item, icon-apply, or present hot path. Use `Debug::Perf` counters/scopes for measured paths, and keep `SelfTest::AppendSelfTestTrace(...)` outside the timed gesture or behind a diagnostic mode that is disabled for perf-budget runs. A 2026-06-29 FolderView closeout pass found that trace writes around `Present1` and per-batch icon apply distorted the huge quick-search and relayout matrices even though the product work was below budget.
- self-test-local perf sinks configured through `Debug::Perf::ConfigureJsonlOutput(...)` must immediately update the process-local sink cache. A previous "no sink configured" fast path must not suppress metrics after a later self-test configures `perf/perf_metrics.jsonl`; `ClearJsonlOutput()` must likewise clear the cache so following tests do not append to a stale sink.

## Test and Archive Requirements

### Selftests

Perf-sensitive features SHOULD prefer:

- `--commands-selftest` for window/UI behavior,
- `--compare-selftest` for Compare/search engine behavior,
- `--fileops-selftest` for File Operations.

If none of the existing suites can exercise the scenario, a new deterministic selftest case or equivalent harness MUST be added.

### Archival

Meaningful validation runs MUST be archived under `Specs/TestRuns/` per `Specs/TestRuns/README.md`.

If the repo archive path is unavailable, the blocked reason MUST be documented explicitly.

### Baselines

Once a scenario becomes important for ongoing work, a same-machine archived baseline SHOULD be recorded and reused for future comparisons.

### Evidence quality gates

Frame-time and latency claims MUST report the actual selftest build flavor carried
by the JSONL `build` field. Release runs are not acceptable evidence when the
artifact is mislabeled as Debug or when the build flavor is unknown.

For percentile-based decisions, `Tools/Show-PerfRuns.ps1` is the required
first-pass analyzer. Use PowerShell 7+ for this script. Any p95 claim SHOULD have
at least 200 samples for the requested metric unless an authoritative per-metric
budget declares a smaller `minimumSamples`; p99 SHOULD have at least 1000
samples. Automation that gates a p95 claim MUST pass `-FailOnQuality`, which
exits non-zero when a quality-gated requested metric has fewer than its p95
sample minimum. An explicitly supplied `-MinimumSamplesForP95` overrides budget
minimums; otherwise `-BudgetPath` selects the budget document used for those
minimums. If that selected budget file exists but cannot be parsed, the analyzer
MUST warn, latch quality failure, and exit non-zero when `-FailOnQuality` is set.
`-FolderViewPreset` uses the p95 rows in
`Specs/Testing/FolderViewPerfBudgets.json5` as its quality-gated metric set; the
remaining preset rows are still displayed but are informational for
`-FailOnQuality` unless requested explicitly with `-Metric`. In `-CompareRun`
mode, `-FailOnQuality` gates the candidate run only; a sparse baseline remains
visible in the comparison table but does not fail the command. Use
`-FolderViewPreset` for the core FolderView frame, scale/cold/slow, scroll,
icon, thumbnail, and refresh metric family when reviewing FolderView runs.
FolderView scroll/dirty-region reviews MUST check that the preset summarizes
`folder.scroll.product_paint_render_count`,
`folder.scroll.product_paint_full_client_count`, and
`folder.scroll.product_paint_dirty_rect_area_px` in addition to
`folder.scroll_input_to_paint_us` and the `folder.frame.*` rows. No-op scroll
requests that do not change the viewport must report zero product paint frames;
viewport-changing scrolls are expected to remain full-client paints unless a
separately gated DXGI scroll-rect implementation exists.
FolderView icon-pipeline reviews MUST check that the preset summarizes
`FolderView.ExecuteEnumeration.IconIndex.QueryExtensions`,
`FolderView.ExecuteEnumeration.IconIndex.QueryPerFileIcons`,
`FolderView.ExecuteEnumeration.IconIndex.BuildPerFilePaths`,
`FolderView.IconLoading.ProcessQueue`, `FolderView.IconLoading.BatchUpdate`,
`FolderView.IconLoading.BitmapConversion`, `icons.queue_wait_to_dequeue_us`,
`icons.extract_us`, `icons.batch_update_scan_us`, `iconcache.shgetfileinfo_us`,
`iconcache.lock_wait_slow_us`, `iconcache.lock_hold_slow_us`, and
`icons.recall_avoided_count`.

FolderView same-folder refresh reviews MUST check the aggregate `folder.refresh.*`
family. `preserve_count`, `rebuild_count`, `selection_preserve_count`,
`rename_transfer_count`, `debounce_delay_ms`, and `enumeration_count` are
one-row-per-refresh counters; `request_to_paint_us` is a request-to-present
latency row emitted only after the refresh result has reached the UI thread and
the pane has presented. Refresh latency MUST use the dedicated FolderView refresh
pending slot so input-to-paint and navigation-to-paint rows cannot overwrite or
misattribute it.

FolderView frame-producing perf cases that make frame p95 claims MUST archive a
`metricQuality` object with `folderFrameTotal.count`,
`folderFramePresent.count`, `samplesEnoughForP95`, `samplesEnoughForP99`, and
`buildConfiguration`. `samplesEnoughForP95` is a hard case gate for
`folder.frame.total_us` and `folder.frame.present_us` when the case claims p95
evidence, and strict budget runs still enforce the budget file's
`minimumSamples`. `folderView_perf_overlay_invalidation_stress` records
`samplesEnoughForP95` as advisory because animation cadence and host load can
reduce frame samples without proving an overlay correctness failure; cite p95 for
that case only when the analyzer quality gate passes. `samplesEnoughForP99` is
advisory until a case deliberately captures at least 1000 frame samples. Other
event-scoped rows exposed by the FolderView preset, such as scroll input latency
or one-shot icon/thumbnail counters, may report `P95Quality=fail`; do not cite a
p95 for those rows unless that exact metric also has enough samples.

Representative FolderView perf coverage MUST include:

- `folderView_perf_scroll_render_stress`: 1,600-item normal-mode scroll/render
  coverage across Brief, Detailed, and Extra Detailed modes. It must produce at
  least 200 `folder.frame.total_us` and `folder.frame.present_us` rows for p95
  claims, emit scroll input latency and product-paint delta metrics, assert
  repeated no-op boundary scrolls do not repaint, and archive whether each
  scroll step changed the viewport.
- `folderView_perf_huge_folder_scale`: synthetic dummy-provider scale coverage at
  10,000 items for routine runs and 50,000 items only when
  `REDSALAMANDER_FOLDERVIEW_HUGE_PERF=1` is set. Artifacts record item count,
  extension count, enumeration time, first visible paint, sort toggle,
  quick-search keystroke-to-paint, select-all scroll brush guard, and working-set
  / private-byte samples including bytes per item.
- `folderView_perf_cold_first_visit`: local first-visit coverage with unique paths
  and extensions after clearing the application `IconCache`. The artifact records
  first enumeration, first paint, icon-index lookup metric count, icon bitmap
  queue count, icon-settle time, and a deterministic forced first thumbnail
  fallback. The artifact MUST state that the OS shell cache is not controlled by
  the test.
- `folderView_perf_slow_virtual_provider`: deterministic slow-provider coverage
  for dummy-provider enumeration latency, icon extraction latency, live-path
  failed icon lookup negative-cache behavior (`icons.repeated_failed_lookup_count`
  must stay bounded), provider-allowed thumbnail lookup latency, and the paste
  shortcut latency coverage supplied by the ShellCommands paste-shortcut cases.
- `folderView_perf_icon_pipeline_cold_slow`: focused icon-pipeline coverage for
  per-file icon types (`.exe`, `.dll`, `.ico`, `.lnk`, `.url`, `.cpl`, `.scr`,
  `.msc`, `.ocx`), one delayed HICON extraction, analyzer-visible queue/extract/
  convert/apply metrics, and synthetic offline/recall placeholder avoidance.
  The artifact must prove at least one visible bitmap icon resolves before the
  delayed extraction finishes, `IconPathLiveLookup` is not consumed for the
  offline/recall dummy fixture, `icons.recall_avoided_count` emits for those
  items, and the thumbnail pass avoids provider-allowed shell I/O.
- `folderView_perf_relayout_churn_while_scrolled`: 10,000-item scrolled-pane
  relayout coverage that alternates DPI, size, light/dark/high-contrast theme
  triggers, emits `folder.relayout_to_paint_us`, archives repaint-burst sample
  quality and `fontRelayoutCovered=false`, and asserts focus/scroll survival.

Every representative FolderView scale/cold/slow/relayout artifact MUST include an
`environmentMatrix` object with build flavor, active DPI, display refresh rate,
display scale percent, local-console/RDP status, WARP availability, whether WARP
was actually executed, adapter name, driver-version availability, high-DPI run
status, and notes for matrix dimensions that require separate hardware or
process-level runs. FolderView WARP coverage is opt-in for selftest evidence:
set `REDSALAMANDER_FOLDERVIEW_FORCE_WARP=1` before launching the Commands
selftest so FolderView creates its D3D device with `D3D_DRIVER_TYPE_WARP`; the
archive must then report `environmentMatrix.warpRunExecuted=true`.

Release perf evidence that depends on selftest cases MUST use a test-enabled
Release build, because normal Release binaries may omit selftest entrypoints:

```powershell
try {
    $env:RSBuildEnableTests='true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Release
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
```

### Automated FolderView budget gates

The authoritative FolderView perf budget file is
`Specs/Testing/FolderViewPerfBudgets.json5`. It is intentionally machine-keyed
through a top-level `machines[]` array: each entry has a `machineHash` and its own
`budgets[]` list. Hard thresholds apply only when the current selftest
`machineHash` matches one of those machine entries, and each hard threshold MUST
cite its source archive, measured value, maximum, statistic, build flavor, and
`minimumSamples`.

Unknown machines are never silent. The native harness emits a visible warning with
a scaffold entry shape to fill, e.g. `{ "machineHash": "<current>", "budgets": [] }`.
The default behavior keeps non-applicable budgets as warnings so ad-hoc Debug runs
can proceed; strict runs add `--selftest-require-perf-budgets` (or
`Run-AllTests.ps1 -RequirePerfBudgets`) to fail when the budget path is missing, no
current-machine entry exists, or no hard entry matches the current build flavor.

Run the focused strict gate through the normal runner with:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -CaseFilter folderView_perf_scroll_render_stress -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5 -RequirePerfBudgets -TimeoutMultiplier 8
```

For an already-built test-enabled Release binary, the native equivalent is:

```powershell
.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-perf-budget=Specs\Testing\FolderViewPerfBudgets.json5 --selftest-require-perf-budgets --selftest-timeout-multiplier=8
```

For grouped validation, pass the comma-separated budgeted case list. The current
budgeted set is `folderView_perf_scroll_render_stress`,
`folderView_perf_overlay_invalidation_stress`,
`folderView_perf_huge_folder_scale`, `folderView_perf_slow_virtual_provider`,
`folderView_perf_relayout_churn_while_scrolled`, and
`folderView_thumbnail_cached_only_no_close_stall`.

Expected-failure smoke checks SHOULD use a local scratch budget with an
impossible maximum and MUST NOT commit that scratch file.

Selftest-only latency injection points MUST be implemented through
`RedSalamander/SelfTest/Common/SelfTestLatencyHooks.h/.cpp` and compiled as
no-ops outside `ENABLE_TESTS`.
These hooks exist to make slow shell/provider/file-system behavior deterministic;
they must not become production delays or runtime configuration.

## Spec Ownership Requirement

The owning normative spec for a subsystem MUST describe:

- the important user-visible performance scenarios,
- the expected test entrypoints,
- the authoritative metric families when they exist.

Performance requirements MUST NOT live only in a WIP plan when the subsystem has already adopted them as standard practice.

When a WIP plan is finished:

- the plan MUST move to `Specs/Plans/Done/`,
- any durable performance contract, verification entrypoint, or workflow requirement discovered in that work MUST be merged into the authoritative subsystem spec or repo-level guidance,
- the finished plan remains historical evidence, not the authoritative contract.

## Completion Criteria

A feature or optimization is performance-complete when:

1. the scenario is named,
2. the metrics exist or were shown to already exist,
3. deterministic coverage exists,
4. archived evidence exists or the block is documented,
5. the owning authoritative spec notes the lasting result when the change establishes or updates a baseline, and any finished WIP plan has been moved to `Specs/Plans/Done/`.
