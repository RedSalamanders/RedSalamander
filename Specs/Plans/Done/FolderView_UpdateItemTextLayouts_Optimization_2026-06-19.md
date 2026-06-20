# FolderView UpdateItemTextLayouts Optimization Plan

**Status:** **CLOSED (Done) 2026-06-19.** Phase 0 measurement revealed the dominant
"cost" was a **measurement artifact** — the perf JSONL sink opened and closed the
file on every metric row. **Fixing the sink** (keep the append handle open) collapsed
`render.layout_items_us` p95 from ~250ms to ~15ms and `folder.frame.total_us` p95
from ~50ms to ~16ms, with no behavior change and all selftests still passing. The
FolderView layout optimization itself is **closed as a measured no-op** — there was
no real layout bottleneck. The genuine, valuable fix was the perf sink
(`Common/Common/PerfJsonl.cpp`).
**Created:** 2026-06-19.
**Parent backlog:** `Specs/Plans/WIP/DxUi_FolderView_Monitor_FuturePerformanceIdeas_2026-05-20.md`.
**Predecessors (both closed, both measurement pilots):**
- `Specs/Plans/Done/FolderView_TextLayout_MetricPilot_2026-06-19.md` — DirectWrite
  text-layout *creation* is ~1.3% of layout time (cache = measured no-op).
- `Specs/Plans/Done/FolderView_LayoutPassDecomposition_MetricPilot_2026-06-19.md` —
  `UpdateItemTextLayouts` owns **~82.6%** of `render.layout_items_us`, per-item-bound
  on the ~26–49-item visible window.

**Scope:** Reduce the per-item cost of `FolderView::UpdateItemTextLayouts`
([FolderView.Layout.cpp:426](../../../RedSalamander/FolderView.Layout.cpp)). Unlike the
two predecessors this plan may ship a **behavior change**, so it carries full
correctness gates. It still starts with measurement.

---

## Problem statement (from measured evidence)

`UpdateItemTextLayouts` dominates the FolderView layout pass, and the cost is
per-item work on the small visible+predictive-buffer window, **not** range size and
**not** `CreateTextLayout` (1.3%). The leading hypothesis is per-item layout
**finalization** that sits outside the creation timer:

- The estimate pass in `LayoutItems` (`if (! _itemMetricsCached)`,
  FolderView.Layout.cpp ~151–268) calls `item.labelLayout.reset()` /
  `detailsLayout.reset()` / `metadataLayout.reset()` for **all** items when the cache
  is invalidated (mode change, item-set change).
- `UpdateItemTextLayouts` then **recreates** the layouts for the visible window and
  calls `ConfigureLabelLayout` (sets trimming/ellipsis sign) and `GetMetrics` (forces
  DirectWrite to compute line-breaking and glyph runs) per item — the expensive part.
- Net: a reset-then-recreate-and-re-finalize churn each invalidating pass.

> **Debug caveat:** the predecessor evidence is Debug (DirectWrite is far slower in
> Debug). **Phase 0 below MUST re-baseline in a Release-with-tests build** before any
> fix is sized — the fix must beat the Release noise floor, not the Debug one.

---

## Results (measured 2026-06-19) — the premise was a measurement artifact

Phase 0 sub-decomposed the per-item body of `UpdateItemTextLayouts`:

| Sub-op | Share of `UpdateItemTextLayouts` |
| --- | --- |
| `GetMetrics` | 2.2% |
| `SetMaxWidth/SetMaxHeight` | 0% |
| `CreateTextLayout` | 1.4% |
| "other" | **96.4%** |

The 96.4% "other" was the perf **JSONL sink** (`Common/Common/PerfJsonl.cpp`):
`WritePerfJsonl` did `create_directories` + a fresh `CreateFileW` (open) + write +
`CloseHandle` **per metric row** (~50–150µs each). The text-layout instrumentation
emits 2 rows per layout creation × ~1,874 creations = ~3,748 open/write/close cycles
per run, all inside `UpdateItemTextLayouts`. That syscall I/O — not layout work — was
the "82.6% / ~200ms". In production the sink is unconfigured and `WritePerfJsonl`
early-returns, so the cost never existed for users.

**Fix:** cache the append handle and reopen only on path change. `FILE_APPEND_DATA`
keeps appends correct across writes; in-process readers (the selftest presence scan)
still see writes via the OS file cache, so no flush is needed. Same-machine
before/after (Debug):

| p95 | Before | After | Δ |
| --- | --- | --- | --- |
| `render.layout_items_us` | ~250ms | 15.5ms | −94% |
| `update_text_layouts_us` | ~250ms | 9ms | −96% |
| `folder.frame.total_us` | ~50ms | 16ms | −67% |

All six FolderView perf selftests pass with `allPresent: true`, confirming
read-visibility is intact. The fix makes **every** perf measurement project-wide
accurate (the prior numbers were inflated by sink I/O) and selftests faster.

**Verdict.** The FolderView layout optimization is a **measured no-op** — there is no
real layout bottleneck; the dominant cost was outside the candidate's mechanism (the
governance's explicit reject criterion). The real, durable fix is the perf sink. The
diagnostic Phase 0 item-level timers were removed; the only shipped change is
`Common/Common/PerfJsonl.cpp`.

**Correction to prior docs.** The 2026-06-19 layout-decomposition finding
("`UpdateItemTextLayouts` owns ~82.6% of layout time") is now understood as a sink
artifact and is annotated as such in `UI_FolderView.md`,
`Testing_PerformanceValidation.md`, and the parent backlog.

---

The original measurement plan below is retained as historical context; it was
superseded by the artifact finding above.

## Phase 0 — Release baseline + sub-decomposition (measurement only)

1. Build Release-with-tests and capture a same-machine baseline of
   `render.layout_items_us`, `folder.layout.update_text_layouts_us`, and
   `folder.layout.update_visible_item_count` over ≥5 runs (Release commands are in
   `Specs/Testing/Testing_TestCoverage.md`). Record the Release noise floor.
2. Add sub-phase instrumentation **inside** the per-item loop of
   `UpdateItemTextLayouts`, separating:
   - `folder.layout.item.configure_us` — `ConfigureLabelLayout`,
   - `folder.layout.item.metrics_us` — `GetMetrics`,
   - `folder.layout.item.setmax_us` — `SetMaxWidth` / `SetMaxHeight` on existing layouts,
   - `folder.layout.item.provider_us` — `_detailsTextProvider` / `_metadataTextProvider`
     / `BuildDetailsText` (should be ~0 in steady state; confirm).
   Accumulate per-call and emit once per `UpdateItemTextLayouts` to avoid per-item emit
   overhead; gate on `Debug::Perf::IsCaptureEnabled()`.
3. Confirm which sub-op dominates. Expected: `metrics_us` (+ `configure_us`). If the
   data contradicts the hypothesis, re-scope the fix before writing any.

---

## Phase 1 — Fix (only the mechanism the data justifies)

Candidate fixes, in rough order of preference (pick per Phase 0 evidence):

1. **Stop the reset-then-recreate churn.** In the estimate pass, do not `reset()`
   layouts whose item identity and layout constraints (text, width, height, mode, DPI,
   font) are unchanged — keep the already-finalized layout instead of forcing a
   recreate + re-`GetMetrics`. This is the highest-leverage fix if `metrics_us`
   dominates.
2. **Defer / avoid `GetMetrics`** when the metric values are already known and the
   constraints did not change (the metrics are cached on the item;
   `item.labelMetrics` etc.).
3. **Avoid redundant `SetMaxWidth`/`SetMaxHeight`** when the constraint is unchanged
   from the previous pass.

Reuse keys (item identity + everything that changes finalized geometry): text/version,
width, height, mode, DPI, font generation, trimming/ellipsis policy, reading direction.
These mirror the cache-key fields already enumerated in parent Candidate 2.

---

## Required metrics

- Phase 0 sub-phase keys above.
- Reuse on the fix: `folder.layout.item.finalize_reused_count` /
  `folder.layout.item.finalize_recreated_count` so the churn reduction is provable.
- Existing: `folder.layout.update_text_layouts_us`, `render.layout_items_us`,
  `folder.layout.update_visible_item_count`, `dwrite.text_layout.create_count`.

## Acceptance gates

- Release `folder.layout.update_text_layouts_us` p95 improves by **≥ 20%** beyond the
  Release run-to-run noise floor (per the parent "Variance before verdicts" rule), and
  `render.layout_items_us` p95 improves materially.
- p99 does not regress; `update_visible_item_count` unchanged (no coverage loss).
- **Correctness (hard gates):** label/details/metadata text, ellipsis/trimming, column
  widths, selection/focus/anchor, hover, quick-search highlight, DPI/theme/font/locale
  re-finalization, and accessibility text all unchanged. A finalized layout MUST be
  re-finalized when any reuse key changes. Add a debug bypass to compare reuse-on vs
  reuse-off output.
- All six `folderView_perf_*` selftests pass.

## No-op criteria

- If Phase 0 shows no single sub-op dominates, or the dominant cost is unavoidable
  DirectWrite work on genuinely-changed items, close as measured no-op (keep the
  sub-phase instrumentation).

## Spec closeout target

On a durable win, record the behavior rule (when finalized layouts are reused vs
re-finalized) in `Specs/UI/UI_FolderView.md` and the metric/gate in
`Specs/Testing/Testing_PerformanceValidation.md` / `Testing_TestCoverage.md`, then move
this plan to `Specs/Plans/Done/`.

## Open questions

- Does Release shrink the per-item cost enough that this is no longer worth fixing?
  (Phase 0 answers this — do not assume the Debug 250ms is real.)
- Does layout reuse interact with the existing lazy `EnsureItemTextLayout` /
  `ProcessIdleLayoutBatch` paths? All three must agree on the reuse keys.
