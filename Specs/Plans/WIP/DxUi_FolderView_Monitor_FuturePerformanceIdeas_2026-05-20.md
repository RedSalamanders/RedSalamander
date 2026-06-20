# DxUi, FolderView, And Monitor Future Performance Ideas

**Status:** WIP decision backlog, not an implementation plan.
**Created:** 2026-05-20.
**Last reviewed:** 2026-06-19.
**Scope:** Future performance improvements for DxUi, FolderView, and
RedSalamanderMonitor after the completed frame-performance foundation and
remaining-work closeout.
**Intent:** Preserve candidate ideas, evidence, entry conditions, and gates so a
future patch can only be called a performance improvement when it has
deterministic coverage and archived before/after evidence.

This document is intentionally stricter than a brainstorming note. It records
what is worth investigating, what is already rejected, and what evidence is
required before implementation. Durable behavior, UI contracts, metrics, and
workflow requirements must move into the authoritative subsystem specs when a
chosen implementation plan closes.

---

## Progress Checklist

- [x] Capture measured local baseline and accepted improvement from the
  2026-05-19 / 2026-05-20 frame-performance work.
- [x] Capture external references used during brainstorming.
- [x] Rank potential future improvements by expected payoff, risk, and
  maintainability.
- [x] Define required metrics and scenario gates for each idea.
- [x] Cross-link the existing DxUi, FolderView, Monitor, and performance
  validation contracts that already own parts of this work.
- [x] Add entry conditions, no-op criteria, and correctness gates so speculative
  rendering changes cannot skip evidence.
- [x] Pick the first implementation target: **FolderView** overlay/scroll/
  text-layout performance (over Monitor high-rate log rendering and DxUi
  chrome/animation smoothness). Chosen 2026-06-19.
- [x] Split the chosen target into a separate implementation plan before coding:
  `Specs/Plans/Done/FolderView_TextLayout_MetricPilot_2026-06-19.md`
  (metric-only first patch, no cache).
- [x] Add or reuse deterministic selftests before changing production behavior:
  reused `folderView_perf_scroll_render_stress` and added `dwrite.text_layout.*`
  metric-presence assertions.
- [x] Archive baseline evidence under `Specs/TestRuns/`: six clean runs under
  `Specs/TestRuns/7d3a1247382a/Commands/` (2026-06-19).
- [x] Move the finished implementation plan to `Specs/Plans/Done/` and merge
  lasting contracts into `Specs/UI/`, `Specs/Core/`, `Specs/Testing/`, or
  repo-level guidance as appropriate. Done 2026-06-19: the text-layout pilot
  closed as a measured no-op; contract merged into `UI_FolderView.md`,
  `Testing_PerformanceValidation.md`, and `Testing_TestCoverage.md`.

---

## Authority And Use

This file is a backlog and decision record. It must not be used as the sole
authorization for production code changes.

Before coding any candidate below, create a focused WIP implementation plan
that names:

- the protected scenario,
- the exact user-visible regression being prevented or improved,
- the metric keys that already exist or will be added first,
- the deterministic selftest or harness,
- the baseline archive to compare against,
- the candidate archive expected after implementation,
- the correctness tests that must pass,
- the authoritative spec that will receive any lasting contract.

Measurement-only work can close as green only when it produces reusable
instrumentation or a reproducible capture workflow. It must not be described as
an optimization.

**Caution on process weight.** This governance is deliberately heavy and has
already earned its keep by rejecting three speculative directions as measured
no-ops. But the document is not the work. A candidate that stays fully analyzed
and never piloted has produced zero runtime value. Once a target is chosen, the
next action is the smallest executable slice (metric presence plus a variance
baseline), not another round of backlog curation.

---

## Global Acceptance Contract

Every perf-sensitive candidate in this document inherits the project contract
from `Specs/Testing/Testing_PerformanceValidation.md`,
`Specs/Testing/Testing_SelfTests.md`, and `Specs/TestRuns/README.md`.

Required for every implementation plan:

1. **Scenario first:** name the exact scenario before changing production code.
   Examples: `folderView_perf_scroll_render_stress`, Monitor ETW burst latency,
   Monitor scrollback page-up, or DxUi connected-overlay animation.
2. **Instrumentation first:** add or reuse metric presence checks before
   optimizing. A red/green metric-presence artifact is preferred when new
   metric keys are introduced.
3. **Deterministic coverage:** use an existing selftest when it covers the
   scenario; add a focused deterministic case when it does not.
4. **Same-machine evidence:** compare baseline and candidate from the same
   machine profile whenever possible. Cross-machine or changed-suite
   comparisons are directional only.
5. **Archived artifacts:** keep `results.json`, `trace.txt`, and
   `perf_metrics.jsonl` together under
   `Specs/TestRuns/<MachineHash>/<Area>/<RunId>/`. Keep custom `perf/*.json`
   summaries with the run when the harness writes them.
6. **Quantiles, not anecdotes:** report count, p50, p95, p99, and max for frame
   or latency metrics when sample count allows. Use p95 as the primary decision
   metric, p99 as the regression guard, and max as diagnostic context.
7. **Variance before verdicts:** a single before/after pair is not evidence.
   Frame metrics drift with background scheduling, so establish the noise floor
   first: run the baseline scenario at least 5 times and record the run-to-run
   p95 spread (min, max, and coefficient of variation). Every percentage
   threshold in this document is measured *against that noise floor* — a delta
   counts as an improvement only when it exceeds the baseline's run-to-run p95
   spread by the stated margin. If the noise floor is wider than the threshold,
   the gate cannot be decided: make the scenario more deterministic or sample it
   harder before claiming any verdict. This is why every threshold below is
   expressed as a percentage rather than an absolute microsecond target.
8. **Correctness is a gate:** visual correctness, ordering, focus/selection,
   DPI/theme/language invalidation, accessibility, and teardown safety must pass
   before any faster metric is accepted.
9. **Memory is a gate for caches:** cache candidates need bounded entry and byte
   budgets, repeated navigation/theme/resize evidence, and explicit eviction
   metrics.
10. **No-op is valid:** if baseline pressure is not present, close the candidate
    as measured no-op and do not land speculative production behavior.
11. **Spec closeout:** if implementation establishes a lasting contract, update
    the authoritative subsystem spec before moving the implementation plan to
    `Specs/Plans/Done/`.

Reject or disable a candidate when:

- metric presence is incomplete,
- the deterministic scenario is flaky or depends on background desktop state,
- the measured improvement is within the baseline's run-to-run p95 noise band,
- p95 improves but p99 or correctness regresses without an accepted reason,
- memory grows without an explicit budget and eviction proof,
- the dominant cost is outside the candidate's mechanism,
- the work requires a full rendering rewrite to prove a small isolated idea.

---

## Ground Truth From Completed Work

The only accepted production performance improvement from the recent
frame-performance pass was FolderView overlay invalidation.

Measured same-machine result:

- `folder.frame.overlay_dirty_rect_area_px` p95 improved from `1,161,911px` to
  `721,522px` (`-37.90%`).
- Full-client overlay rows dropped from `124/124` to `4/148` (`100%` to
  `2.70%`).
- Overlay `folder.frame.total_us` p95 improved from `71,825us` to `54,236us`
  (`-24.49%`).
- Overlay `render.frame_us` p95 improved from `73,533us` to `55,859us`
  (`-24.04%`).
- Non-overlay scroll guard did not regress: `folder.frame.total_us` p95
  improved from `45,217us` to `41,145us` (`-9.01%`) and p99 improved from
  `47,086us` to `46,126us` (`-2.04%`).

Implication: the strongest current evidence favors reducing overdraw, dirty
area, layout churn, and text/layout recreation. More speculative GPU, present,
composition, or D3D batching changes must be gated behind evidence.

Measured no-op or rejected areas:

- Monitor ETW scheduling: measured no-op. `monitor.etw.batch_drain_us` and
  `monitor.frame.append_to_visible_us` did not pass the scheduling-change
  threshold.
- DXGI flip-discard full redraw: rejected for current surfaces. Current DxUi,
  FolderView, and Monitor paths depend on dirty-rect and scroll-rect partial
  present.
- Composition pilot: rejected. The allowed connected-overlay surface did not
  show an overlay rendering bottleneck; the observed issue was scheduler/timer
  cadence.

Authoritative historical evidence lives in:

- `Specs/Plans/Done/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/Core/Core_RedSalamanderMonitor.md`

---

## Existing Validation Entrypoints

Use these before adding a new harness. Add a new case only when the existing
entrypoint cannot isolate the scenario or metric family.

| Area | Entrypoint | Current role |
| --- | --- | --- |
| FolderView large folder | `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4` | Enumeration, sort, icon, queue, and first visible render baseline. |
| FolderView scroll/render | `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4` | Real horizontal and vertical scroll across Brief, Detailed, and Extra Detailed modes. |
| FolderView overlay | `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4` | Overlay dirty-area and incremental-search effect guard. |
| FolderView sort | `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_sort_toggle_stress --selftest-timeout-multiplier=4` | Repeated Name, Extension, Time, Size, and None sort toggles with inactive quick-search guard. |
| FolderView directory churn | `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_directory_change_storm --selftest-timeout-multiplier=4` | Deterministic create, rename, delete, and directory churn. |
| FolderView icon contention | `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_iconcache_contention --selftest-timeout-multiplier=4` | IconCache wait/hold diagnostics under dual-pane icon-heavy churn. |
| DxUi frame host | `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` then `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost` | Dirty-frame metric presence and host render/present behavior. |
| DxUi animation | `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` | Connected-overlay animation, frame, render, present, and jitter metrics. |
| Monitor default frame | Test-enabled Monitor build, then `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf` | Default chrome and frame metric presence. |
| Monitor ETW latency | Test-enabled Monitor build, then `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260` | Append-to-visible, ETW drain, queue depth, and repost metrics. |
| Monitor scrollback | Test-enabled Monitor build, then `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-scrollback-selftest` | SCROLL_BACK slice, mode, frame, and present metrics. |
| Monitor document model | `.\.build\x64\Debug\MonitorTest.exe --document-model-selftest` | Non-UI ordering and document model correctness guard. |

Monitor selftests require a test-enabled build:

```powershell
try {
    $env:RSBuildEnableTests = 'true'
    .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
```

GUI selftests that depend on focus, pointer routing, visible presentation, or
Monitor chrome must run in a foreground-capable desktop session. Hidden-window
launches are not valid evidence for those cases.

---

## External References

External references inform candidate shape, but local RedSalamander metrics and
correctness tests decide whether a change lands.

- [DXUT](https://github.com/microsoft/DXUT): useful as structural inspiration
  for deterministic frame phases, centralized device lifecycle, and callback
  sequencing. It is archived and sample-oriented, so do not copy its global
  state or legacy sample patterns directly.
- [DXGI 1.2 flip model, dirty rectangles, and scrolled areas](https://learn.microsoft.com/en-us/windows/win32/direct3ddxgi/dxgi-1-2-presentation-improvements):
  supports dirty rectangles and scroll rectangles to reduce memory bandwidth
  and draw work. It also requires tracking dirty rectangles across frames so
  every dirty pixel is current before `Present1`.
- [Direct3D 11 dynamic resources](https://learn.microsoft.com/en-us/windows/win32/direct3d11/how-to--use-dynamic-resources):
  dynamic vertex/index buffers should use `D3D11_MAP_WRITE_NO_OVERWRITE` while
  appending unused ranges and `D3D11_MAP_WRITE_DISCARD` when wrapping or
  replacing the buffer.
- [Direct2D profiling DirectX apps](https://learn.microsoft.com/en-us/windows/win32/direct2d/profiling-directx-applications):
  WPR/WPA/GPUView-style traces are needed to distinguish UI-thread CPU,
  Direct2D command submission, GPU busy, present wait, DWM composition, and
  synchronization bottlenecks.
- [Direct2D D1111 layer warning](https://learn.microsoft.com/en-us/windows/win32/direct2d/d1111-using-layer-when-clip-is-sufficient):
  rectangular clips should use Push/Pop Clip instead of expensive layers when a
  layer adds no opacity mask or non-rectangular effect.
- [DirectWrite rendering](https://learn.microsoft.com/en-us/windows/win32/directwrite/rendering-directwrite):
  `IDWriteTextLayout` objects are a first-class rendering path for DirectWrite
  text. RedSalamander has repeated `CreateTextLayout` call sites, so layout
  creation must be measured before caching is justified.
- [DirectComposition overview and benefits](https://learn.microsoft.com/en-us/windows/win32/directcomp/why-use-directcomposition-):
  DirectComposition can animate visuals independently of the UI thread, but
  current local evidence does not justify enabling it broadly.
- [PresentMon](https://github.com/GameTechDev/PresentMon): useful for
  high-level frame presentation and latency capture outside app-local JSONL
  metrics.

---

## Guiding Principles

1. Measure before optimizing. Every accepted change needs a deterministic
   scenario and archived before/after evidence.
2. Prefer smaller dirty regions over new rendering technology. The recent 24%
   p95 win came from repainting less, not from a new graphics stack.
3. Keep flip-sequential partial present as the default. Dirty/scroll rects are
   part of the current correctness and performance contract.
4. Add shared abstractions only after the first adopter proves the shape and a
   second adopter needs it. A single-surface pilot may use a shared-shaped local
   adapter instead of a broad framework rewrite.
5. Cache only with explicit invalidation keys, byte/entry budgets, and eviction
   metrics. Text and retained visuals can improve p95, but uncontrolled caches
   become memory leaks.
6. Do not move all UI to composition. Only isolated transform/opacity surfaces
   should be considered, and only after a scenario proves UI-thread rendering is
   the bottleneck.
7. Avoid per-control hot-path telemetry. Use per-frame, per-scenario,
   coalesced, and thresholded slow-path metrics.
8. Preserve accessibility, DPI, theme, high-contrast, reduced-motion, and input
   contracts as first-class gates.

---

## Candidate Priority Matrix

Rows are ordered by rough return-on-investment (payoff per unit effort), which
is why the cheap-to-measure text-layout candidate now leads the table even
though it shares "Highest" priority with the dirty-region model.

| Candidate | Priority | Est. effort | Entry condition | Default decision if entry condition is absent |
| --- | --- | --- | --- | --- |
| DirectWrite text layout cache (metric-only first) | Highest | S to measure, M to cache | `CreateTextLayout` count/time is proven material in a text-heavy scenario. | Add metrics only; do not cache. |
| Unified dirty/scroll region model | Highest | M for one surface, XL if generalized | Dirty area or full redraws dominate a named scenario. | Close as measured no-op; keep existing partial-present path. |
| FolderView Virtualization V2 | High | XL | Visible work, icon hydration, or retained row rebuilds dominate huge-folder scenarios after narrower fixes. | Keep visible-only model and improve smaller bottlenecks first. |
| Monitor tail and scrollback renderer | High | M | ETW append, queue depth, or scrollback p95 pressure is reproduced in focused Monitor drills. | Keep current scheduling and renderer; rerun weekly or before related changes. |
| Perf lab with WPR/GPUView/PresentMon | Required for advanced rendering | M (tooling, no product code) | Any GPU, present-policy, batching, or composition claim is being considered. | Block production graphics-policy changes. |
| D2D/D3D batching layer | Medium, high risk | L to XL | External trace proves draw submission or GPU work dominates after dirty/text fixes. | Reject as speculative complexity. |
| Narrow composition pilot | Low | M, high teardown risk | An allowed transform/opacity surface has rendering-caused jitter that composition can remove. | Keep rejected. Scheduler jitter is not a composition justification. |

Effort legend (rough ROI-sorting estimates, not commitments): **S** ≈ under a
day, **M** ≈ a few days, **L** ≈ one to two weeks, **XL** ≈ multi-week and
touches multiple subsystems.

---

## Candidate 1: Unified Dirty Region And Scroll Region Model

**Recommendation:** Highest priority, but start with one surface and prove the
API before broad sharing.

### Idea

Create a dirty-region model that DxUi, FolderView, NavigationView, and Monitor
can eventually use to express:

- visual dirty regions,
- layout dirty regions,
- scroll regions and scroll offsets,
- retained previous-frame regions,
- fallback-to-full-redraw reasons.

The first patch should not be a large rendering rewrite. Prefer a
FolderView-local or tiny common tracker with a future-compatible shape:
accumulate rectangles, clip to client, coalesce when complexity is too high,
account for area, and prepare `Present1` parameters without changing present
policy for unrelated hosts.

### Entry Conditions

Start only when a named scenario has baseline evidence showing at least one:

- p95 dirty area is high relative to visible changed content,
- full-client redraws appear after the first frame when only small visual state
  changed,
- scroll movement repaints instead of reusing scroll rectangles,
- p95/p99 `frame.total_us` pressure correlates with dirty area or fallback
  reason counts.

Existing first scenarios:

- `folderView_perf_overlay_invalidation_stress`
- `folderView_perf_scroll_render_stress`
- `folderView_perf_directory_change_storm`
- Monitor scrollback drill, only if scroll-rect reuse is the suspected limiter.

### Design Rules

- Keep flip-sequential partial present as the default.
- Track dirty rectangles across swap-chain buffers when partial present is used.
- Include the current client size, DPI generation, buffer/present generation,
  dirty rectangles, scroll rectangle, scroll offset, and fallback reason in the
  per-frame snapshot.
- Do not expose owning `HRGN`, `ID2D1*`, `IDXGI*`, or other Windows resources
  through raw ownership. Use WIL RAII if Win32 regions are used internally.
- Cap region complexity. When the cap is exceeded, fall back to full redraw and
  emit the reason.
- Treat resize, device loss, theme change, DPI change, unknown invalidation, and
  debug forced-full-redraw as explicit fallback reasons.

### Required Metrics

- `*.frame.dirty_rect_count`
- `*.frame.dirty_rect_area_px`
- `*.frame.full_redraw_count`
- `*.frame.full_redraw_reason`
- `*.frame.scroll_rect_area_px`
- `*.frame.scroll_reuse_count`
- `*.frame.input_to_paint_us`
- `*.frame.total_us`
- `*.frame.present_us`

For FolderView, preserve existing compatibility metrics:

- `folder.frame.dirty_rect_area_px`
- `folder.frame.overlay_dirty_rect_area_px`
- `render.dirty_rect_area_px`
- `render.frame_us`

### Acceptance Gates

Accept only if at least one target scenario shows:

- p95 `frame.total_us` improves by at least 15%, or
- p95 dirty area drops by at least 25% with no p95/p99 frame regression, or
- scrollback/high-rate append p95 improves without ordering or correctness
  regressions.

The candidate must also pass:

- deterministic visual or state verification for dirty propagation,
- resize, device-loss/device-recreate, DPI, theme, and occlusion/minimize/restore
  guards when the touched host owns those paths,
- full-redraw fallback correctness,
- teardown without leaked posted work or retained stale surfaces.

### No-op Criteria

Close as measured no-op if:

- dirty area is already small for the scenario,
- present or DWM wait dominates while app dirty work is not material,
- region coalescing overhead offsets repaint savings,
- correctness requires invalidating most of the client for the chosen surface.

### Risks And Mitigations

- **Stale visuals:** use a debug verifier that occasionally forces full redraw
  and compares conservative visual probes or scenario state.
- **Region complexity cost:** cap region count/area and emit fallback metrics.
- **Multi-buffer partial-present bugs:** preserve previous-frame dirty data and
  document the buffer-generation rule in the owning spec before closeout.
- **Over-abstraction:** land one surface first; extract only after a second
  surface uses the same contract.

---

## Candidate 2: DirectWrite Text Layout Cache

> **CLOSED 2026-06-19 — measured no-op.** The metric-only pilot
> (`Specs/Plans/Done/FolderView_TextLayout_MetricPilot_2026-06-19.md`) measured
> FolderView text-layout creation at ~1.2–1.4% of `render.layout_items_us`, present
> in only ~5 of ~108 render frames — not material. The cache is not built; the
> `dwrite.text_layout.*` instrumentation is retained (see
> `Specs/UI/UI_FolderView.md`). Reopen only with new evidence.

**Recommendation:** Highest priority alongside Candidate 1, but metrics must
prove layout creation is material before caching.

### Idea

Add explicit text-layout caching for stable text surfaces. Cache keys must be
structured, not ad hoc concatenated strings, and must include every layout input
that can change measured glyph placement.

Candidate key fields:

- text content or stable item text version,
- text format identity and font generation,
- width and height constraints,
- wrapping, trimming, alignment, max-lines, and ellipsis behavior,
- DPI,
- reading direction, flow direction, and locale,
- view mode or column identity when it changes layout constraints,
- text decoration/style state when it changes layout.

Start with FolderView labels/details or a single DxUi menu path only after
`dwrite.text_layout.*` metric presence is green.

### Entry Conditions

Start only when the baseline shows:

- repeated `CreateTextLayout` calls in a named scenario,
- measurable `dwrite.text_layout.create_us` or equivalent layout time,
- stable text/layout inputs across repeated paints,
- a cacheable surface with clear invalidation ownership.

Existing first scenarios:

- FolderView scroll in Detailed and Extra Detailed modes.
- FolderView rapid rename and sort toggles.
- FolderView icon-heavy folder after enumeration settles.
- DxUi menu open/cascade keyboard and pointer navigation.
- NavigationView path edit suggestion popup with many path segments.

### Design Rules

- Use `wil::com_ptr<IDWriteTextLayout>` for cached layout ownership.
- Keep cached layouts UI-thread owned unless a future plan proves a safe
  cross-thread DirectWrite ownership model.
- Use explicit entry and estimated-byte budgets; emit current size and eviction
  counts.
- Separate stable item/detail caches from transient overlay/status text.
- Invalidate on text change, rename, metadata/detail change, width/height
  constraint change, view mode change, DPI change, theme/font change, locale or
  reading-direction change, and trimming/wrapping policy change.
- Sort/filter/reorder should invalidate row leases and visible record bindings;
  it should not necessarily flush text layouts whose key inputs are unchanged.
- Keep a debug bypass so deterministic tests can compare cache-on/cache-off
  behavior when investigating stale text.

### Required Metrics

- `dwrite.text_layout.create_count`
- `dwrite.text_layout.create_us`
- `dwrite.text_layout.cache_hit_count`
- `dwrite.text_layout.cache_miss_count`
- `dwrite.text_layout.eviction_count`
- `dwrite.text_layout.cache_entry_count`
- `dwrite.text_layout.cache_bytes`
- scenario-level `frame.total_us`, `layout_us`, and `render_us`

### Acceptance Gates

Accept only if:

- one text-heavy scenario shows p95 frame or layout time improvement by at
  least 10%,
- cache hit rate is high enough that lookup overhead is clearly justified,
- p99 frame/layout time does not regress materially,
- memory growth is bounded after repeated navigation, theme, DPI, font, resize,
  rename, and sort runs,
- DPI/theme/language/rename changes invalidate correctly.

### No-op Criteria

Close as measured no-op if:

- layout creation is not a material p95/p99 contributor,
- hit rate is low because text/constraints are too volatile,
- cache lookup/locking overhead offsets layout reuse,
- memory budget needed for useful reuse is unacceptable.

### Risks And Mitigations

- **Stale text:** explicit key fields plus invalidation tests for rename,
  metadata, DPI, theme, locale, and view mode.
- **Memory retention:** strict budgets, eviction metrics, and repeated navigation
  evidence.
- **Tiny-label overhead:** do not cache surfaces whose measured create time is
  lower than lookup/eviction overhead.
- **Accessibility drift:** ensure cached visual text still matches model text
  exposed through UI Automation/debug snapshots.

---

## Candidate 3: FolderView Virtualization V2

**Recommendation:** High priority after dirty-region and text-layout work either
lands or is closed as measured no-op.

### Idea

Move FolderView from "visible-only render with some cached layouts" toward a
retained visible-window model:

- stable render records for visible and near-visible rows,
- item state versions for label, icon, selection, hover, metadata, hidden/filter
  state, and plugin/VFS identity,
- layout only for invalidated rows or mode changes,
- cheap placeholder draw during fast scroll,
- async icon/detail hydration after scroll settles.

### Entry Conditions

Start only when baseline evidence shows:

- visible work or row rebuilds dominate p95/p99 after narrower dirty/text
  candidates,
- icon/detail hydration is causing visible latency or repeated redraws,
- directory-change churn invalidates rows outside the viewport,
- large-folder responsiveness cannot be improved with smaller isolated changes.

### Design Rules

- Key retained records by stable item identity plus visual-state generation, not
  by transient row index alone.
- Treat selection, focus, anchor, hover, quick search, hidden/filter state,
  plugin/VFS source identity, and async icon generation as independent versioned
  inputs.
- Apply async results only when item identity and generation still match.
- Keep UI-thread model reads locked or copied; never read/write shared
  non-atomic state without synchronization.
- Keep a force-slow-path debug toggle for verification and bisection.

### Required Metrics

- `folder.visible_cache.hit_count`
- `folder.visible_cache.miss_count`
- `folder.visible_cache.eviction_count`
- `folder.layout.invalidated_item_count`
- `folder.layout.reused_item_count`
- `folder.icon.placeholder_frame_count`
- `folder.frame.visible_work_count`
- existing `folder.frame.total_us`
- existing `folder.frame.input_to_paint_us`
- existing `folder.frame.dirty_rect_area_px`

### Acceptance Gates

Accept only if:

- 10k-item scroll p95 improves by at least 15%, or
- icon-heavy folder p95 improves by at least 15%, or
- remote/VFS navigation remains responsive while enumeration/icon work
  continues.

All acceptance must also preserve:

- selection/focus/anchor under reorder, rename, filter, and directory changes,
- icon and metadata freshness after async updates,
- accessibility row text and selection state,
- memory budget after repeated navigation and large-folder churn.

### First Scenarios

- 10k local files in Detailed and Extra Detailed mode, or a documented smaller
  fixture if 10k makes the selftest too slow for routine evidence.
- 10k mixed extensions with icon extraction, or reuse
  `folderView_perf_iconcache_contention` if it isolates the decision.
- Directory change storm while scrolled away from changed rows.
- Rapid sort toggle with selection and focus preservation.

### No-op Criteria

Close as measured no-op if:

- dirty-region or text-layout fixes remove the p95/p99 pressure,
- retained record memory exceeds the budget needed for the gain,
- correctness requires broad rebinding on every common mutation.

---

## Candidate 4: Monitor Tail And Scrollback Renderer

**Recommendation:** High priority for RedSalamanderMonitor only when focused
Monitor drills reproduce pressure.

### Idea

Treat the Monitor renderer like a terminal/log renderer:

- append-only line ring with stable line IDs,
- visible-line layout cache,
- append coalescing,
- scrollback slice reuse,
- explicit AUTO_SCROLL versus SCROLL_BACK frame contracts,
- scroll rectangles for line movement where safe.

### Current Evidence

The Monitor ETW latency gate was close enough to keep visible. The final
closeout ETW latency archive reported:

- `append_to_visible` p95 `22,501us`,
- `batch_drain` p95 `1,214us`,
- `frame.total` p95 `18,746us`,
- `present` p95 `4,622us`.

The Monitor scrollback archive reported:

- `scrollback_slice` p95 `35,844us`,
- `frame.total` p95 `128,975us`,
- `present` p95 `867us`.

**Flag — suspected real, not a cold artifact.** A `frame.total` p95 of
`128,975us` is ~129ms, roughly 8 fps during scrollback: visible jank, not a
cosmetic delta. Process discipline still applies (a single-slice sample is too
thin for a p95 verdict), but this is the one Monitor number that smells like a
real user-facing bug rather than a measurement artifact. Triage it as
suspected-real: the *first* action is to harden the scrollback drill to collect
repeated page-up/page-down samples (see Entry Conditions below), and if 129ms
reproduces, this candidate jumps ahead of its "rerun weekly" default decision.

The earlier green latency run had `append_to_visible_us` p95 near the
`50,000us` threshold. Monitor can become a bottleneck when ETW rate increases,
filters change, or scrollback is active.

### Entry Conditions

Start only when a focused Monitor drill shows one:

- `monitor.frame.append_to_visible_us` p95 exceeds the accepted latency budget,
- `monitor.etw.batch_drain_us` p95/p99 exceeds the scheduling gate,
- `monitor.etw.queue_depth` remains high under sustained burst,
- `monitor.frame.scrollback_slice_us` or `monitor.frame.total_us` dominates a
  multi-sample scrollback scenario.

If the current scrollback drill has too few samples to support a p95 decision,
first improve the drill to collect repeated page-up/page-down samples. That
measurement work is not an optimization.

### Design Rules

- Preserve append order and never drop ETW rows silently.
- Keep AUTO_SCROLL and SCROLL_BACK mode transitions explicit.
- Do not self-post paint/batch work when no real work is pending.
- Keep line-layout caches bounded by visible plus near-visible lines, not the
  entire history.
- Keep Monitor test-only helpers behind `ENABLE_TESTS`; default Debug and
  Release Monitor builds must not pollute the live monitor display with self
  diagnostics.
- If new cross-thread payload messages are introduced, use
  `PostMessagePayload(...)`, initialize/drain payload windows, and add teardown
  stress coverage.

### Required Metrics

- `monitor.line_layout.cache_hit_count`
- `monitor.line_layout.cache_miss_count`
- `monitor.line_layout.eviction_count`
- `monitor.frame.append_to_visible_us`
- `monitor.frame.scrollback_slice_us`
- `monitor.frame.visible_line_count`
- `monitor.frame.new_line_count`
- `monitor.frame.reused_line_count`
- `monitor.frame.scroll_rect_area_px`
- `monitor.etw.batch_drain_us`
- `monitor.etw.queue_depth`
- `monitor.etw.batch_repost_count`

### Acceptance Gates

Accept only if:

- high-rate append p95 `append_to_visible_us` improves by at least 20%, or
- scrollback page-up/page-down p95 `frame.total_us` improves by at least 20%, or
- queue depth remains bounded under sustained ETW burst without line ordering
  regressions.

The default chrome selftest, latency burst drill, scrollback drill, and
`MonitorTest --document-model-selftest` must all pass for renderer changes.

### No-op Criteria

Close as measured no-op if:

- append and drain metrics remain below the scheduling threshold,
- scrollback pressure is a cold single-slice artifact and not reproducible,
- present/DWM wait dominates while app-side renderer work is not material,
- ordering or mode correctness requires the current simpler path.

---

## Candidate 5: Perf Lab With WPR, GPUView, And PresentMon

**Recommendation:** Required before any major GPU, present-policy, batching, or
composition work.

### Idea

Add scripts and documentation for repeatable external captures:

- WPR/WPA/GPUView ETL capture for CPU/GPU/DWM analysis.
- PresentMon CSV capture for frame pacing and presentation behavior.
- Correlation markers from app-local JSONL metrics to external traces.

### Why It Matters

App-local JSONL tells us what RedSalamander believes it spent. External traces
tell us whether the bottleneck is:

- UI-thread CPU,
- Direct2D command submission,
- GPU busy,
- present wait,
- DWM composition,
- device/resource creation,
- synchronization.

This workflow must exist before D3D batching, flip-discard experiments,
composition pilots, or any claim that the GPU/present stack is the limiter.

### Required Artifacts

- `Specs/Testing/Testing_PerformanceValidation.md` update with capture recipe.
- Script under `Tools/` or another appropriate tooling folder.
- Archive pointer under `Specs/TestRuns/<MachineHash>/<Area>/<RunId>/`.
- `summary.md` or `external_trace_summary.json` listing:
  - git commit,
  - build configuration and platform,
  - command line,
  - machine hash,
  - OS build,
  - display refresh rate and DPI,
  - app metric archive path,
  - external tool versions,
  - captured ETL/CSV filenames or external storage path,
  - scenario start/end markers.

Large ETL files should normally stay untracked unless explicitly curated. The
repo evidence should still contain a small summary and the app-local metrics
needed to understand the run.

### Acceptance Gates

No production code changes are allowed in this candidate except instrumentation
hooks. Accept when the capture process is reproducible, documented, and has one
known-good capture against a deterministic scenario.

---

## Candidate 6: D2D/D3D Batching Layer

**Recommendation:** Medium priority, high risk. Do not start until Candidate 5
proves Direct2D draw submission or GPU work is the bottleneck after dirty-region
and text-layout opportunities are addressed.

### Idea

For repeated simple primitives, replace many immediate Direct2D calls with
batched geometry:

- row backgrounds,
- selection rectangles,
- hover rectangles,
- separators,
- simple icons/badges if atlas-backed later.

If D3D11 is used directly, dynamic vertex/index buffers must follow
`WRITE_NO_OVERWRITE` while appending unused ranges and `WRITE_DISCARD` when the
buffer wraps or is replaced.

### Entry Conditions

Start only when external trace plus app-local metrics show:

- Direct2D draw submission or GPU work dominates the frame,
- text layout, icon conversion, dirty area, and present wait are not the primary
  bottlenecks,
- the target primitives are numerous and simple enough to batch without changing
  text quality, DPI behavior, hit-testing, or accessibility semantics.

### Required Metrics

- `render.draw_call_count`
- `render.state_change_count`
- `render.batch_count`
- `render.vertex_count`
- `render.map_us`
- `render.gpu_busy_us` when external capture is available
- scenario-level `frame.total_us`, `render_us`, and `present_us`

### Acceptance Gates

Accept only if:

- draw-call or Direct2D submission cost is proven dominant,
- p95 frame time improves by at least 15% in a large visible-work scenario,
- p99 frame time does not materially regress,
- no text quality, DPI, theme, high-contrast, reduced-motion, or accessibility
  regression is introduced.

### No-op Criteria

Reject without production code when:

- text/icons/layout dominate,
- dirty-region improvements remove the pressure,
- batching needs broad renderer ownership changes before a narrow benefit can
  be measured,
- external traces do not support the bottleneck theory.

---

## Candidate 7: Narrow Composition Pilot For Proven Animation Bottlenecks

**Recommendation:** Low priority until evidence changes.

### Idea

Keep composition as an opt-in, isolated pilot only for transform/opacity
animations where the animated content can be snapshotted as a bitmap surface and
moved independently of the UI thread.

Allowed surfaces could include:

- connected overlay transform,
- tooltip fade/translate if it becomes expensive,
- non-text chrome reveal transitions.

Disallowed initial surfaces:

- FolderView item rows,
- Monitor text lines,
- DirectWrite-heavy controls,
- virtualized lists,
- anything requiring live accessibility text updates during animation.

### Current Evidence

The 2026-05-20 gate rejected compositor code. Connected-overlay paint p95 was
`11us`, frame render p95 stayed below the 8.333ms budget, and present p95 was
`381us`; the observed issue was synthetic scheduler/timer jitter, not overlay
rendering cost.

### Entry Conditions

Start only if a future allowed-surface scenario proves:

- UI-thread animation p95 jitter is above budget,
- rendering/paint work is a dominant cause,
- the animated content can be snapshotted without stale visuals or accessibility
  drift,
- reduced-motion behavior remains instant or within the existing contract.

### Acceptance Gates

Accept only if composition reduces jitter or frame p95 without increasing
memory, input latency, stale visuals, teardown complexity, or accessibility
staleness.

### No-op Criteria

Close as measured no-op if:

- timer/scheduler cadence is the bottleneck,
- paint/render p95 is already below budget,
- the target surface requires live text/list updates during animation,
- the implementation would replace broader DirectWrite/list rendering paths.

---

## Recommended Sequencing

1. **Phase 0: Decision And Plan Split**
   - Choose one first target.
   - Create a separate WIP implementation plan.
   - Fill the perf gate template before coding.

2. **Phase A: Measurement Backbone**
   - Add or refine missing metric presence checks.
   - Add text-layout and dirty-region metrics before optimizing.
   - Add the external capture workflow before GPU, present-policy, batching, or
     composition changes.

3. **Phase B: FolderView Dirty Region + Text Layout Pilot**
   - Start with FolderView because it has the only accepted recent runtime win.
   - Use existing overlay and scroll scenarios first.
   - Add one 10k or documented scaled fixture only if existing scenarios cannot
     isolate the decision.
   - Validate normal scroll, overlay, sort, directory churn, icon contention,
     theme, DPI, rename, and metadata scenarios.

4. **Phase C: Monitor Tail/Scrollback**
   - Improve scrollback sampling first if current evidence is too thin.
   - Optimize append and scrollback only if gates show reproducible p95 pressure.

5. **Phase D: Advanced Rendering**
   - Consider D2D/D3D batching only after external traces show draw submission or
     GPU busy is the limiter.
   - Consider composition only for isolated animation surfaces after the
     existing rejected gate changes.

---

## Proposed First Implementation Plan

The best first plan should be:

**FolderView Dirty Region And Text Layout Measurement Pilot**

Reasons:

- FolderView is key to the product.
- The only measured recent runtime performance gain was in FolderView.
- It exercises both dominant ideas: smaller dirty regions and less text layout
  churn.
- It can reuse `folderView_perf_overlay_invalidation_stress`,
  `folderView_perf_scroll_render_stress`, `folderView_perf_sort_toggle_stress`,
  `folderView_perf_directory_change_storm`, and
  `folderView_perf_iconcache_contention` before adding a new fixture.

Minimum deliverables for that future plan:

- Add `dwrite.text_layout.*` metric presence for FolderView before adding a
  cache.
- Re-run baseline FolderView perf scenarios and record archive paths.
- Decide whether a new 10k-item scenario is needed; if not, document why
  existing scenarios isolate the decision.
- Add a small bounded `IDWriteTextLayout` cache only for one stable
  FolderView text surface.
- Add explicit invalidation for rename, detail/metadata change, view mode,
  column width, DPI, theme/font, locale/reading direction, wrapping/trimming
  policy, and model generation changes.
- Add dirty-region work only where the baseline still shows full-client or
  excessive dirty area after the overlay fix.
- Preserve scroll, overlay, sort, directory churn, icon contention, selection,
  focus, accessibility, and teardown correctness.
- Archive before/after evidence and compare same-machine p95/p99 metrics.
- Update `Specs/UI/UI_FolderView.md`,
  `Specs/Testing/Testing_PerformanceValidation.md`, or
  `Specs/Testing/Testing_TestCoverage.md` when the implementation establishes a
  durable metric, validation entrypoint, or behavior rule.

**Honesty note — this pilot bundles two "Highest" candidates.** The title pairs
the dirty-region model (Candidate 1) and the text-layout cache (Candidate 2),
which sits in mild tension with the "land one surface, extract after a second
needs it" principle. The resolution is sequencing, not scope creep: the **first
patch is metric-only** — it ships `dwrite.text_layout.*` presence metrics and a
variance baseline, with **no production cache and no dirty-region change**. Only
after the data shows which cost (text-layout creation versus residual dirty
area) is actually material does a *second* patch add the one mechanism the
evidence justifies. If neither is material, the pilot closes as a measured no-op
and the instrumentation is the deliverable.

Do not combine FolderView virtualization V2 with the first text-cache/dirty-area
pilot. If the pilot proves retained visible records are needed, split
virtualization into a later plan with its own correctness and memory gates.

---

## Non-Goals

- Do not replace Direct2D/DirectWrite wholesale.
- Do not enable flip-discard full redraw for current DxUi, FolderView, or
  Monitor surfaces.
- Do not ship DirectComposition broadly.
- Do not claim performance improvement from measurement-only work.
- Do not add unbounded caches.
- Do not add hot-path per-control telemetry.
- Do not introduce a broad shared rendering framework before one surface proves
  the contract and a second surface needs it.
- Do not weaken visual, input, accessibility, DPI, theme, high-contrast,
  reduced-motion, or teardown contracts to improve a single metric.
- Do not rely on hidden-window GUI selftest runs for focus, pointer, presentation,
  or Monitor chrome evidence.

---

## Decisions Made

First target — **decided 2026-06-19: option 1, FolderView** overlay/scroll/
text-layout performance — chosen over (2) Monitor high-rate ETW/log rendering
and (3) DxUi chrome/menus/animation. Rationale: FolderView holds the only
measured recent runtime win, is central to the product, and reuses five existing
deterministic selftests.

Resolved sub-decisions for this target:

- **First patch is metric-only.** Ship `dwrite.text_layout.*` presence metrics
  plus a variance baseline; no cache and no dirty-region production change in the
  first patch.
- **First cacheable surface (deferred to the data).** FolderView Detailed /
  Extra Detailed row label+detail text is the leading candidate, but the cache
  is not built until metrics prove creation cost is material.
- **Reuse the existing ~1,600-item scroll stress** plus the overlay, sort,
  directory-churn, and icon-contention scenarios first; add a bounded 10k fixture
  only if those cannot isolate the decision.
- **Dirty-region work stays deferred** until the baseline still shows full-client
  or excessive dirty area *after* the already-shipped overlay fix.
- **Authoritative specs to receive any lasting contract:**
  `Specs/UI/UI_FolderView.md` for behavior and
  `Specs/Testing/Testing_PerformanceValidation.md` (plus
  `Testing_TestCoverage.md`) for the metric/validation contract.

Tracking plan (closed): `Specs/Plans/Done/FolderView_TextLayout_MetricPilot_2026-06-19.md`.

**Pilot result (2026-06-19, measured).** The metric-only patch landed and ran. On
machine `7d3a1247382a`, text-layout creation is **not material**: ~1.2–1.4% of
layout-pass time and present in only 5 of ~108 render frames. The `folder.frame
.total_us` p95 noise floor is ~9.7% CoV / 31.5% spread over six runs, so the
~1.3% creation cost is inside the noise. **Candidate 2 (DirectWrite text-layout
cache) is closed as a measured no-op**; the instrumentation is retained as
reusable measurement infrastructure. The data redirects attention to
`render.layout_items_us` (p95 ~200ms vs frame p95 ~50ms) — a separate, larger
layout-pass cost that is the real next lead.

**Layout-pass decomposition + correction (2026-06-19, measured).** The decomposition
pilot (`Specs/Plans/Done/FolderView_LayoutPassDecomposition_MetricPilot_2026-06-19.md`)
*appeared* to show `UpdateItemTextLayouts` owning ~82.6% of layout time. The
optimization follow-up
(`Specs/Plans/Done/FolderView_UpdateItemTextLayouts_Optimization_2026-06-19.md`)
sub-decomposed that and found it was **~96% a measurement artifact**: the perf JSONL
sink opened and closed the file on every metric row, and the per-creation text-layout
emits drove ~3,748 such writes per run inside `UpdateItemTextLayouts`. **Fixing the
sink** to keep the append handle open (`Common/Common/PerfJsonl.cpp`) collapsed
`render.layout_items_us` p95 from ~250ms to ~15ms and `folder.frame.total_us` p95
from ~50ms to ~16ms (Debug), with all selftests still passing. There is **no real
FolderView layout bottleneck**; the layout optimization is closed as a measured
no-op. The valuable, durable fix was the perf sink — which also makes every perf
measurement project-wide accurate (prior absolute figures across these docs were
inflated by sink I/O).

Still open:

- nothing for FolderView layout (the apparent bottleneck was a measurement artifact);
- optionally, coalesce any remaining high-frequency per-creation emits into per-pass
  aggregates per the spec anti-pattern (the sink fix already removed the distortion).
