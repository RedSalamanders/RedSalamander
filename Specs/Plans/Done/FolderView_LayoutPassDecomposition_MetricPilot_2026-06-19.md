# FolderView Layout-Pass Decomposition Measurement Pilot

> **SUPERSEDED / CORRECTED (2026-06-19).** The "82.6% in `UpdateItemTextLayouts`"
> headline below was later found to be **~96% a measurement artifact** — the perf
> JSONL sink opened and closed the file on every metric row, and the per-creation
> text-layout emits drove ~3,748 such writes per run inside that function. After the
> sink was fixed to keep the append handle open (`Common/Common/PerfJsonl.cpp`), the
> layout pass is ~15ms p95 (was ~250ms) — there is no real layout bottleneck. See
> `Specs/Plans/Done/FolderView_UpdateItemTextLayouts_Optimization_2026-06-19.md`. The
> phase-decomposition *metrics* added here remain valid and are now accurate; the
> conclusions drawn from the inflated numbers do not.

**Status:** **CLOSED (Done) 2026-06-19** (results corrected — see banner above).
Decomposition landed and measured; the apparent dominance of `UpdateItemTextLayouts`
was a perf-sink artifact, fixed in the follow-up. The phase metrics
(`folder.layout.*_us`) are retained and now accurate.
**Created:** 2026-06-19.
**Parent backlog:** `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`.
**Predecessor:** `Specs/Plans/Done/FolderView_TextLayout_MetricPilot_2026-06-19.md`
(closed as measured no-op; it surfaced this lead).
**Scope:** Decompose the dominant `render.layout_items_us` cost (p95 ≈ 200ms on
machine 7d3a1247382a) into named sub-phases so a future patch can target the phase
that actually owns the time. **No optimization in this patch — instrumentation only.**

## Results (measured 2026-06-19, machine 7d3a1247382a, Debug x64)

Build clean (0/0). Scenario `folderView_perf_scroll_render_stress`, five clean
same-machine runs (archived RunIds under `Specs/TestRuns/7d3a1247382a/Commands/`).

**Phase breakdown (share of total `render.layout_items_us` across runs):**

| Phase | Share | Worst-call p95 |
| --- | --- | --- |
| `update_text_layouts_us` | **82.6%** | ~252ms |
| `estimate_metrics_us` | 9.1% | ~48ms |
| `column_resolve_us` | 0.9% | ~1.9ms |
| `bounds_us` | 0.8% | ~1.7ms |
| `setup_us` | 0.7% | ~0.4ms |

The original hypothesis (estimate-metrics pass) was wrong: `UpdateItemTextLayouts`
dominates.

**Drill-down — it is per-item-bound, not range-bound.**
`folder.layout.update_visible_item_count` shows the pass iterates only **26–49
items** per call (p50 26, p95 49) out of 1,600, yet costs ~250–290ms — i.e.
**~5.75ms per item** worst case. So the cost is the per-item body on the visible
window, not the size of the range.

**Cross-reference with the predecessor pilot.** DirectWrite layout *creation* was
1.3% of layout time, so the ~81% inside `UpdateItemTextLayouts` is **not** the
`CreateTextLayout` calls. The strong hypothesis: the layout *finalization* that
sits outside the creation timer — `ConfigureLabelLayout` (sets trimming/ellipsis)
and `GetMetrics` (forces DirectWrite to compute line-breaking/glyph runs) — runs
per visible item every invalidating pass, because the estimate pass `reset()`s all
layouts so they are recreated and re-finalized. That is the next plan's target.

**Debug caveat.** These are Debug-build numbers; DirectWrite is far slower in Debug,
so the absolute ~250ms is inflated. The **relative** breakdown (82.6%,
per-item-bound) is the durable finding; a Release-with-tests baseline should
confirm absolute sizing before a fix is scoped.

**Verdict.** Dominant target identified: `UpdateItemTextLayouts` per-item layout
finalization. This is a measurement pilot — it closes green by producing reusable
instrumentation and a precise next target, not by optimizing. Instrumentation is
retained (gated on `Debug::Perf::IsCaptureEnabled()`, zero production cost;
`update_visible_item_count` doubles as an ongoing visible-work guard).

**Correctness gate: passed.** All six FolderView perf selftests still exit 0.

**Follow-up plan (not started):** sub-decompose the `UpdateItemTextLayouts` per-item
body — time `ConfigureLabelLayout`, `GetMetrics`, and `SetMaxWidth/SetMaxHeight`
separately — then test the fix direction (avoid the reset-then-recreate-and-
re-finalize churn; reuse finalized layouts across passes when item identity and
constraints are unchanged).

## Why this pilot

The text-layout pilot proved DirectWrite layout creation is ~1.3% of the layout
pass. The layout pass itself (`render.layout_items_us`, a single `FolderView::
LayoutItems()` call) has a p95 of ~200ms versus `folder.frame.total_us` p95 ~50ms,
so it is the real FolderView bottleneck. We do not yet know *which part* of
`LayoutItems()` owns the time. This pilot measures that before any optimization.

## Correcting the original hypothesis

The lead was framed as "sort vs icon hydration vs predictive-buffer pre-pass vs
visible-range compute." Source inspection of `FolderView::LayoutItems()`
([FolderView.Layout.cpp:62](../../../RedSalamander/FolderView.Layout.cpp)) shows:

- **Sort is not in `LayoutItems`** — it lives in `FolderView::ApplyCurrentSort` /
  `ExecuteEnumeration.SortMerge` (already has its own metrics).
- **Icon hydration is not in `LayoutItems`** — it lives in FolderView.Icons.cpp /
  the icon-loading queue.
- **Visible-range compute** (`GetVisibleItemRange`) is a cheap call inside
  `UpdateItemTextLayouts`, not a standalone phase.

What `LayoutItems()` actually does, in order, is the real decomposition target:

| Phase | Lines | What it does | Metric key |
| --- | --- | --- | --- |
| Setup | 67–70 | `EnsureDeviceIndependentResources` + `UpdateEstimatedMetrics` | `folder.layout.setup_us` |
| Metrics estimation | 134–251 | **O(N) over all items**, gated by `_itemMetricsCached`: builds details/metadata text via `_detailsTextProvider`/`_metadataTextProvider`/`BuildDetailsText`, estimates widths | `folder.layout.estimate_metrics_us` |
| Column resolve | 253–312 | tile-height + per-item width metrics + `FolderViewColumnLayout::Resolve` packing | `folder.layout.column_resolve_us` |
| Bounds assignment | 314–358 | column counts, prefix sums, per-item `bounds` rects | `folder.layout.bounds_us` |
| Update text layouts | 361 | `UpdateItemTextLayouts` (visible + predictive buffer) | `folder.layout.update_text_layouts_us` |

The **metrics-estimation pass** is the prime suspect: it is the only O(N)-over-the-
*full*-item-list block that does string building and provider callbacks, and it runs
whenever `_itemMetricsCached` is invalidated (mode change, item-set change). A
counter `folder.layout.metrics_estimate_pass_count` records how often the full pass
runs so cost can be correlated with cache misses.

## Instrumentation design

A gated `markLayoutPhase(name)` lambda local to `LayoutItems()` emits the elapsed
time since the previous mark as `folder.layout.<phase>_us`. Gating is on
`Debug::Perf::IsCaptureEnabled()` (same gate the `Debug::Perf::Scope` RAII uses), so
normal builds pay nothing. The five phase timings nest inside the existing
`render.layout_items_us` Scope and should sum to roughly its value, which is a
self-consistency check. Manual timestamps (not RAII `Scope`) are used because the
phases produce values consumed later in the function, so block-scoping a `Scope`
would break variable lifetimes.

New metric keys (timing + one counter; no caches, no behavior change):

- `folder.layout.setup_us`
- `folder.layout.estimate_metrics_us`
- `folder.layout.column_resolve_us`
- `folder.layout.bounds_us`
- `folder.layout.update_text_layouts_us`
- `folder.layout.metrics_estimate_pass_count` (counter)

## Selftest wiring

Add the five `folder.layout.*_us` keys to the `frameMetricPresence` vector in
`TestFolderViewPerfScrollRenderStress`
([Commands.SelfTest.ViewCommands.cpp](../../../RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp)).
The 1,600-item scroll/mode-change scenario drives `LayoutItems` on the non-empty
path, so all five fire.

## Build / run / archive

Same methodology as the predecessor pilot:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
# 5+ clean same-machine runs; Start-Process -Wait (GUI process), kill any prior
# instance between runs, read each archived RunId (not the pooled last_run).
```

Output auto-archives under `Specs/TestRuns/<MachineHash>/Commands/<RunId>/`.
GUI selftests require a foreground desktop session.

## Decision rule

1. Establish the per-phase p95 across ≥5 same-machine runs, plus the
   `render.layout_items_us` p95 noise floor.
2. The phase whose p95 is the largest fraction of `render.layout_items_us` p95 is
   the optimization target. If `estimate_metrics_us` dominates and correlates with
   `metrics_estimate_pass_count`, the fix is about cache invalidation / deferring
   provider work, not raw loop speed.
3. Open a **separate** optimization plan for the winning phase, with its own
   before/after gates. This patch ships only the decomposition.

## Correctness gates

- All six FolderView perf selftests still pass (pure instrumentation; the
  `markLayoutPhase` lambda has no behavioral effect).
- No new cross-thread access (LayoutItems is UI-thread).

## Spec closeout target

On a durable finding, record the phase metric family + the dominant-phase result in
`Specs/UI/UI_FolderView.md` and `Specs/Testing/Testing_PerformanceValidation.md` /
`Testing_TestCoverage.md`, then move this plan to `Specs/Plans/Done/`.
