# FolderView DirectWrite Text-Layout Measurement Pilot

**Status:** **CLOSED (Done) 2026-06-19.** Metric-only patch landed and measured;
the DirectWrite text-layout cache is closed as a **measured no-op**. Lasting
contract merged into `Specs/UI/UI_FolderView.md`,
`Specs/Testing/Testing_PerformanceValidation.md`, and
`Specs/Testing/Testing_TestCoverage.md`. Instrumentation retained. Single-machine
evidence accepted at closeout; cross-machine re-confirmation and the
`render.layout_items_us` follow-up are optional future work.
**Created:** 2026-06-19.
**Parent backlog:** `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
(Candidate 2, and the "Proposed First Implementation Plan").
**Scope:** Instrument per-item `IDWriteTextLayout` creation in FolderView so we can
decide — with archived same-machine evidence and a variance band — whether a text
layout cache is justified. **No cache and no dirty-region change in this patch.**

This plan satisfies the parent backlog's "Authority And Use" template: it names the
protected scenario, the regression being measured, the metric keys, the deterministic
selftest, the baseline archive, the correctness gates, and the spec that receives any
lasting contract.

---

## Results (measured 2026-06-19, machine 7d3a1247382a)

Build: Debug x64, clean (0 warnings, 0 errors). Scenario:
`folderView_perf_scroll_render_stress` (1,600 items). Six clean same-machine runs,
each its own archived RunId under `Specs/TestRuns/7d3a1247382a/Commands/`.

**Noise floor (validates the variance gate).** `folder.frame.total_us` p95 across
the six runs: min 44,824us, max 60,925us, mean 51,134us — run-to-run spread
**31.5%**, CoV **9.7%**. A single before/after pair is therefore not evidence; a
candidate must beat ~10% just to clear run-to-run noise.

**Materiality verdict — text-layout creation is NOT material.**

- `sum(dwrite.text_layout.create_us) / sum(render.layout_items_us)` = **1.18–1.44%**
  across all six runs.
- Only **5 of ~108 render frames** create any layout in the render path
  (`dwrite.text_layout.frame_create_*`), every run — creation contributes
  essentially nothing to per-frame render cost.
- Creation count is deterministic at **1,874** per run; per-creation cost is real
  (2–142us, long tail), so the metric is genuinely measuring DirectWrite work.

A cache would reclaim at best ~1.3% of layout-pass time — far below the 15%
acceptance gate and inside the 9.7% noise floor. **Decision: close the DirectWrite
text-layout cache (parent Candidate 2) as a measured no-op.** The instrumentation
is the deliverable and is retained: bounded, gated on
`Debug::Perf::IsCaptureEnabled()` (zero production cost), and the exact seam a
future cache would use if evidence ever changes.

**New lead surfaced by the data.** `render.layout_items_us` p95 is ~200ms while
`folder.frame.total_us` p95 is ~50ms, and text-layout creation is only ~1.3% of
that layout time. The dominant layout-pass cost is therefore something else
(sort, icon hydration, the predictive-buffer pre-pass, or visible-range
computation). That is the next thing to instrument — a separate candidate, not
this one.

**Correctness gate: passed.** All six FolderView perf selftests still pass
(`scroll_render_stress`, `large_folder_baseline`, `sort_toggle_stress`,
`overlay_invalidation_stress`, `directory_change_storm`, `iconcache_contention` —
all exit 0). The helper is a transparent wrapper around the same `CreateTextLayout`
call.

**Method note.** Reliable evidence requires `Start-Process -Wait` (the app is a
GUI process; piping/`*>` redirection does not block on it) and killing any prior
instance between runs (single-instance forwarding otherwise pools runs into one
`last_run` file and fakes zero variance). The archived per-RunId folders are the
authoritative per-run captures, not `last_run`.

**Closeout (2026-06-19):** single-machine no-op accepted; the lasting contract was
merged into the authoritative specs above and this plan moved to
`Specs/Plans/Done/`. Optional future work: re-confirm the no-op on a second machine
profile, and open a separate plan to instrument and attack the dominant
`render.layout_items_us` cost (not text-layout creation).

---

## Protected scenario

`folderView_perf_scroll_render_stress` (1,600 items, real horizontal/vertical scroll
across Brief, Detailed, and Extra Detailed modes). This scenario drives all three
item-layout creation paths and is the variance vehicle. Secondary signal:
`folderView_perf_large_folder_baseline` (densest first-paint creation).

## What is being measured (not yet optimized)

Whether `IDWriteTextLayout` creation for FolderView item text (label / details /
metadata) is a **material** per-frame and per-layout-pass cost. The only accepted
recent FolderView win came from repainting less; this pilot asks whether *recreating
text layouts less* is the next lever — before building anything.

## Instrumentation seam (verified in source)

There are **three** live UI-thread item-layout creation functions, all creating the
same `item.labelLayout` / `item.detailsLayout` / `item.metadataLayout` objects:

| Function | File:line | When it runs |
| --- | --- | --- |
| `UpdateItemTextLayouts` | [FolderView.Layout.cpp:401](../../../RedSalamander/FolderView.Layout.cpp) (creates at 462/521/565) | Layout pass (visible + predictive buffer), outside the render frame |
| `EnsureItemTextLayout` | [FolderView.Layout.cpp:731](../../../RedSalamander/FolderView.Layout.cpp) (creates at 756/796/832) | Lazily inside `DrawItem`, **inside** the render frame |
| `ProcessIdleLayoutBatch` | [FolderView.Layout.cpp:908](../../../RedSalamander/FolderView.Layout.cpp) (creates at 945/981/1010) | Idle `WM_TIMER` batch (UI thread) |

All three route through a single new helper
`FolderView::CreateInstrumentedItemTextLayout(...)`, which is the exclusive point a
future cache lookup would hook. `UpdateEstimatedMetrics` (Layout.cpp:18/45) creates
sample-text layouts for metric estimation only and is **excluded** (one-off per
DPI/font change, not the per-item path). The overlay/chrome layouts in
FolderView.Rendering.cpp (already counted by `render.empty_state_layout_creates`) are
out of scope.

## New metric keys (counts + timing only — no cache keys yet)

Per the parent backlog: metric-only phase adds **no** `cache_hit`/`miss`/`eviction`/
`bytes` keys (there is no cache).

| Key | Type | Where emitted | Meaning |
| --- | --- | --- | --- |
| `dwrite.text_layout.create_count` | counter | helper, per creation | Total item-layout creations across all three seams. |
| `dwrite.text_layout.create_us` | duration (us) | helper, per creation | Single-creation cost distribution. `value0` = kind (0 label, 1 details, 2 metadata). Clamped to ≥1us so the JSONL sink encodes it as a duration. |
| `dwrite.text_layout.frame_create_count` | counter | render-frame scope-exit ([FolderView.Rendering.cpp:987](../../../RedSalamander/FolderView.Rendering.cpp)) | Item-layout creations during one render frame (render-path only). |
| `dwrite.text_layout.frame_create_us` | duration (us) | render-frame scope-exit, when count > 0 | Summed creation time attributable to one render frame; directly comparable to `folder.frame.total_us`. |

The per-call keys are complete and frame-boundary independent (they count every seam).
The per-frame keys are render-window-scoped, so layout-pass creation cost shows up via
the per-call totals against `render.layout_items_us`, and render-path creation cost
shows up via `frame_create_us` against `folder.frame.total_us`.

Production overhead is zero: timing and emission are gated on
`Debug::Perf::IsCaptureEnabled()` (same gate the `Debug::Perf::Scope` RAII uses), so a
normal build with no JSONL sink and ETW off pays only one branch per creation.

## Selftest wiring

Add metric-presence assertions to the `frameMetricPresence` vector in
`TestFolderViewPerfScrollRenderStress`
([Commands.SelfTest.ViewCommands.cpp:17376](../../../RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp))
via the existing `makeMetricPresence` lambda, for all four new keys. The emitted
`folderView_perf_scroll_render_stress_metrics.json` `allPresent` field then enforces a
red/green metric-presence artifact.

## Build / run / archive

```powershell
# Build (Debug x64 is where selftests compile)
.\build.ps1

# Confirm the case list
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-list-cases --selftest-case=folderView_perf

# Establish the variance band: run the scenario N times (foreground desktop session required)
for ($i = 1; $i -le 5; $i++) {
    .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
}

# Aggregate p50/p95/p99/max per metric
.\Tools\Show-PerfRuns.ps1
```

- Raw output: `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\perf\perf_metrics.jsonl`
  and `...\perf\folderView_perf_scroll_render_stress_metrics.json`.
- Auto-archived (Debug, repo root detected) under
  `Specs/TestRuns/<MachineHash>/Commands/<RunId>/`; confirm via the `ArchiveToRepo:`
  trace line.
- GUI selftests are **only** valid in a foreground-capable desktop session
  (Testing_SelfTests.md). Hidden/background launches are not evidence.

## Variance band and decision rule (defined before running)

1. Run `folderView_perf_scroll_render_stress` at least 5 times same-machine and record
   the run-to-run p95 spread of `folder.frame.total_us` (min, max, coefficient of
   variation). This is the noise floor.
2. **Text-layout creation is material** (cache worth designing) if **either**:
   - `dwrite.text_layout.frame_create_us` p95 ≥ 15% of `folder.frame.total_us` p95,
     by a margin exceeding the noise floor; **or**
   - summed `dwrite.text_layout.create_us` over the layout pass is a dominant fraction
     of `render.layout_items_us` p95.
3. If neither holds, **close as measured no-op**: the instrumentation is the
   deliverable and no cache is built.

## Correctness gates (must pass regardless of the perf verdict)

- All existing `folderView_perf_*` cases still pass (no behavior change — the helper is
  a pure wrapper around the same `CreateTextLayout` call).
- Scroll, overlay, sort, directory churn, icon contention, selection/focus, DPI, theme,
  rename, and teardown unaffected.
- No new cross-thread access (all three seams + reset/emit are UI-thread).

## Open questions resolved into this patch

- Per-kind split: deferred. `value0` carries kind in the raw JSONL; a single
  `create_us` key avoids key sprawl. Split only if the cache decision needs it.
- 10k fixture: not added. The existing 1,600-item scroll stress drives all three seams;
  add a bounded 10k fixture only if the variance band is inconclusive.

## Still open (decided after the data lands)

- Whether text-layout creation or residual dirty area is the material cost.
- Whether a cache is justified, and if so which surface first.

## Spec closeout target

On a durable outcome, fold the lasting contract into `Specs/UI/UI_FolderView.md`
(behavior) and `Specs/Testing/Testing_PerformanceValidation.md` /
`Testing_TestCoverage.md` (metric + validation entrypoint), then move this plan to
`Specs/Plans/Done/`.
