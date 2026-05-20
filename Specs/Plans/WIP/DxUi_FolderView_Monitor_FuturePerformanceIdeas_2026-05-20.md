# DxUi, FolderView, And Monitor Future Performance Ideas

**Status:** WIP idea/spec backlog.
**Created:** 2026-05-20.
**Scope:** Future performance improvements for DxUi, FolderView, and RedSalamanderMonitor after the completed frame-performance foundation and remaining-work closeout.
**Intent:** Preserve the best candidate ideas, the evidence that motivated them, and the measurement gates required before any implementation is described as a performance improvement.

---

## Progress Checklist

- [x] Capture measured local baseline and accepted improvement from the 2026-05-19 / 2026-05-20 frame-performance work.
- [x] Capture external references used during brainstorming.
- [x] Rank potential future improvements by expected payoff, risk, and maintainability.
- [x] Define required metrics and scenario gates for each idea.
- [ ] Pick the first implementation target: FolderView huge folders, Monitor high-rate log rendering, or DxUi chrome/animation smoothness.
- [ ] Split the chosen target into a separate implementation plan before coding.
- [ ] Add or reuse deterministic selftests before changing production behavior.
- [ ] Archive before/after evidence under `Specs/TestRuns/` before accepting any optimization.

---

## Ground Truth From Completed Work

The only accepted production performance improvement from the recent frame-performance pass was FolderView overlay invalidation.

Measured same-machine result:

- `folder.frame.overlay_dirty_rect_area_px` p95 improved from `1,161,911px` to `721,522px` (`-37.90%`).
- Full-client overlay rows dropped from `124/124` to `4/148` (`100%` to `2.70%`).
- Overlay `folder.frame.total_us` p95 improved from `71,825us` to `54,236us` (`-24.49%`).
- Overlay `render.frame_us` p95 improved from `73,533us` to `55,859us` (`-24.04%`).
- Non-overlay scroll guard did not regress: `folder.frame.total_us` p95 improved from `45,217us` to `41,145us` (`-9.01%`) and p99 improved from `47,086us` to `46,126us` (`-2.04%`).

Implication: the strongest current evidence favors reducing overdraw, dirty area, layout churn, and text/layout recreation. More speculative GPU/present/composition changes must be gated behind evidence.

Measured no-op or rejected areas:

- Monitor ETW scheduling: measured no-op. `monitor.etw.batch_drain_us` and `monitor.frame.append_to_visible_us` did not pass the scheduling-change threshold.
- DXGI flip-discard full redraw: rejected for current surfaces. Current DxUi, FolderView, and Monitor paths depend on dirty-rect and scroll-rect partial present.
- Composition pilot: rejected. The allowed connected-overlay surface did not show an overlay rendering bottleneck; the observed issue was scheduler/timer cadence.

---

## External References

- [DXUT](https://github.com/microsoft/DXUT): useful as structural inspiration for deterministic frame phases, centralized device lifecycle, and callback sequencing. It is archived and sample-oriented, so do not copy its global-state or legacy sample patterns directly.
- [DXGI 1.2 flip model, dirty rectangles, and scrolled areas](https://learn.microsoft.com/en-us/windows/win32/direct3ddxgi/dxgi-1-2-presentation-improvements): supports using dirty rectangles and scroll rectangles to reduce memory bandwidth and draw work. This matches the measured FolderView overlay win and current Monitor scrollback direction.
- [Direct3D 11 dynamic resources](https://learn.microsoft.com/en-us/windows/win32/direct3d11/how-to--use-dynamic-resources): dynamic vertex/index buffers should use `D3D11_MAP_WRITE_NO_OVERWRITE` and `D3D11_MAP_WRITE_DISCARD` patterns when serially appending geometry. Relevant only if profiling proves Direct2D draw submission is a bottleneck.
- [Direct2D profiling DirectX apps](https://learn.microsoft.com/en-us/windows/win32/direct2d/profiling-directx-applications): recommends WPR/WPA/GPUView-style tracing to understand CPU, GPU, and present behavior.
- [Direct2D D1111 layer warning](https://learn.microsoft.com/en-us/windows/win32/direct2d/d1111-using-layer-when-clip-is-sufficient): layers are expensive when a simple rectangular clip is sufficient. This informs future overlay, popup, and clipping audits.
- [DirectWrite rendering](https://learn.microsoft.com/en-us/windows/win32/directwrite/rendering-directwrite): DirectWrite layout/rendering choices matter for text-heavy UI. RedSalamander has many `CreateTextLayout` call sites in FolderView, DxUi controls, menus, NavigationView, and overlays.
- [DirectComposition overview and benefits](https://learn.microsoft.com/en-us/windows/win32/directcomp/why-use-directcomposition-): DirectComposition can animate visuals independently of the UI thread, but current local evidence does not justify enabling it broadly.
- [PresentMon](https://github.com/GameTechDev/PresentMon): useful for high-level frame presentation and latency capture outside app-local JSONL metrics.

---

## Guiding Principles

1. Measure before optimizing. Every accepted change needs a deterministic scenario and archived before/after evidence.
2. Prefer smaller dirty regions over new rendering technology. The recent 24% p95 win came from repainting less, not from a new graphics stack.
3. Keep flip-sequential partial present as the default. Dirty/scroll rects are part of the current correctness and performance contract.
4. Cache only with explicit invalidation keys and budgets. Text and retained visuals can improve p95, but uncontrolled caches can turn into memory leaks.
5. Do not move all UI to composition. Only isolated transform/opacity surfaces should be considered, and only after a scenario proves UI-thread animation is the bottleneck.
6. Avoid per-control hot-path telemetry. Use per-frame, per-scenario, and thresholded slow-path metrics.

---

## Candidate 1: Unified Dirty Region And Scroll Region Engine

**Recommendation:** Highest priority.

### Idea

Create a shared dirty-region model that DxUi, FolderView, NavigationView, and Monitor can use to express:

- visual dirty regions,
- layout dirty regions,
- scroll regions and scroll offsets,
- retained previous-frame regions,
- fallback-to-full-redraw reasons.

This should not be a large abstract rendering rewrite. It should start as a small library around region accumulation, coalescing, clipping, area accounting, and `Present1` parameter preparation.

### Why It Could Be Heavy Impact

The measured FolderView overlay optimization reduced p95 frame time by about 24% by avoiding full-client invalidation after the overlay settled. A shared dirty model can apply the same principle to:

- DxUi popups, menus, connected overlays, and page transitions,
- FolderView selection, hover, quick search, busy/error overlays, and incremental icon arrival,
- Monitor tail append and scrollback,
- NavigationView breadcrumb/edit suggestions.

### Potential Components

- `DirtyRegionTracker`: accumulates rectangles, clips to client, coalesces when region complexity gets too high.
- `ScrollRegionTracker`: tracks scrollable content copies and newly exposed regions.
- `DirtyFrameSnapshot`: immutable per-frame dirty data passed to render/present.
- `FallbackReason` metrics: full redraw due to resize, device loss, theme/DPI change, region complexity, unknown invalidation, or forced scenario.

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

### Gates

Accept only if at least one target scenario shows:

- p95 `frame.total_us` improves by at least 15%, or
- p95 dirty area drops by at least 25% without p95/p99 frame regression, or
- scrollback/high-rate append p95 improves without ordering/correctness regressions.

### Risks

- Stale visuals if dirty propagation misses a dependency.
- Excessive region complexity can cost more than repainting.
- Multi-buffer dirty-rect correctness is subtle; previous/current dirty intersections must be tracked.

### Mitigations

- Debug verifier that occasionally forces full redraw and compares visible checksums or conservative visual probes.
- Region-complexity cap that falls back to full redraw with a metric.
- Per-surface rollout rather than changing every host at once.

---

## Candidate 2: DirectWrite Text Layout Cache

**Recommendation:** Highest priority alongside Candidate 1.

### Idea

Add explicit text-layout caching for stable text surfaces. The cache should be keyed by:

- text content or stable item text version,
- text format identity,
- width and height constraints,
- DPI,
- flow direction and locale,
- theme/text state when it affects layout or decorations.

Start with FolderView and DxUi menus because they are text-heavy and have repeated paint paths.

### Why It Could Be Heavy Impact

Searches show many `CreateTextLayout` call sites in hot or repeated rendering paths, including:

- `RedSalamander/FolderView.Rendering.cpp`,
- `Common/DxUi/DxUi.Controls.cpp`,
- `Common/DxUi/DxUi.Menu.cpp`,
- `RedSalamander/NavigationView.*`,
- `RedSalamander/Ui/AlertOverlay.h`.

DirectWrite layout creation can be expensive relative to drawing an already-created layout. The recent overlay work also showed `render.incremental_search_effect_updates` dropping after tighter invalidation. A layout cache can reduce both CPU work and frame variance.

### Potential Components

- `TextLayoutCache`: small LRU with byte/entry budgets and explicit invalidation.
- `TextLayoutKey`: structured key, no ad hoc string concatenation.
- `TextLayoutLease`: safe wrapper for cached `IDWriteTextLayout`.
- Per-surface invalidation hooks: item rename, width change, DPI change, theme/font change, flow-direction change.

### Required Metrics

- `dwrite.text_layout.create_count`
- `dwrite.text_layout.create_us`
- `dwrite.text_layout.cache_hit_count`
- `dwrite.text_layout.cache_miss_count`
- `dwrite.text_layout.eviction_count`
- scenario-level `frame.total_us`, `layout_us`, and `render_us`

### Gates

Accept only if:

- one text-heavy scenario shows p95 frame or layout time improvement by at least 10%, and
- memory growth is bounded after repeated navigation/theme/resize runs, and
- DPI/theme/language changes invalidate correctly.

### First Scenarios

- FolderView 10k item scroll in Detailed and Extra Detailed modes.
- FolderView rapid rename and sort toggles.
- DxUi menu open/cascade keyboard and pointer navigation.
- NavigationView path edit suggestion popup with many path segments.

### Risks

- Stale text after rename, localization, DPI, or theme changes.
- Cache memory growth with long paths and remote folder churn.
- Cache lookup overhead can erase wins for tiny labels.

### Mitigations

- Strict per-cache budgets.
- Separate caches for stable item labels and transient overlay/status text.
- Metrics required from the first patch.

---

## Candidate 3: FolderView Virtualization V2

**Recommendation:** High priority after Candidates 1 and 2.

### Idea

Move FolderView from "visible-only render with some cached layouts" toward a retained visible-window model:

- stable render records for visible and near-visible rows,
- item state versions for label, icon, selection, hover, metadata, hidden/filter state,
- layout only for invalidated rows or mode changes,
- cheap placeholder draw during fast scroll, with async icon/detail hydration after scroll settles.

### Why It Could Be Heavy Impact

FolderView is central to the product. It has large folders, icons, details, overlays, sorting, filtering, hidden names, plugin/VFS sources, and async icon arrival. The measured overlay result suggests that fine-grained invalidation and bounded work can pay off immediately.

### Required Metrics

- `folder.visible_cache.hit_count`
- `folder.visible_cache.miss_count`
- `folder.visible_cache.eviction_count`
- `folder.layout.invalidated_item_count`
- `folder.layout.reused_item_count`
- `folder.icon.placeholder_frame_count`
- `folder.frame.visible_work_count`
- existing `folder.frame.total_us`, `folder.frame.input_to_paint_us`, `folder.frame.dirty_rect_area_px`

### Gates

Accept only if:

- 10k item scroll p95 improves by at least 15%, or
- icon-heavy folder p95 improves by at least 15%, or
- remote/VFS navigation remains responsive while enumeration/icon work continues.

### First Scenarios

- 10k local files in Detailed and Extra Detailed mode.
- 10k mixed extensions with icon extraction.
- Directory change storm while scrolled away from the changed rows.
- Rapid sort toggle with selection and focus preservation.

### Risks

- Selection/focus state bugs when item indexes change.
- Stale icons or stale metadata after async updates.
- Increased complexity in FolderView state ownership.

### Mitigations

- Version every retained row record by item identity plus visual state.
- Keep a "force slow path" debug toggle for verification.
- Add focused selftests for selection/focus under reorder, rename, filter, and directory-change churn.

---

## Candidate 4: Monitor Tail And Scrollback Renderer

**Recommendation:** High priority for RedSalamanderMonitor.

### Idea

Treat the Monitor renderer like a terminal/log renderer:

- append-only line ring with stable line IDs,
- visible-line layout cache,
- append coalescing,
- scrollback slice reuse,
- explicit tail mode versus scrollback mode frame contracts,
- scroll rectangles for line movement where safe.

### Why It Could Be Heavy Impact

The Monitor ETW latency gate was close enough to keep visible. The final closeout ETW latency archive reported:

- `append_to_visible` p95 `22,501us`,
- `batch_drain` p95 `1,214us`,
- `frame.total` p95 `18,746us`,
- `present` p95 `4,622us`.

The earlier green run had `append_to_visible_us` p95 near the `50,000us` threshold. Monitor can become a major bottleneck when ETW rate increases, filters change, or scrollback is active.

### Required Metrics

- `monitor.line_layout.cache_hit_count`
- `monitor.line_layout.cache_miss_count`
- `monitor.frame.append_to_visible_us`
- `monitor.frame.scrollback_slice_us`
- `monitor.frame.visible_line_count`
- `monitor.frame.new_line_count`
- `monitor.frame.reused_line_count`
- `monitor.frame.scroll_rect_area_px`
- `monitor.etw.batch_drain_us`
- `monitor.etw.queue_depth`

### Gates

Accept only if:

- high-rate append p95 `append_to_visible_us` improves by at least 20%, or
- scrollback page-up p95 `frame.total_us` improves by at least 20%, or
- queue depth remains bounded under sustained ETW burst without line ordering regressions.

### First Scenarios

- 60, 600, and 6,000 event burst modes.
- High-rate filter toggle while events continue.
- Scrollback page-up/page-down with 10k lines.
- Resize while tailing.

### Risks

- Incorrect line order or lost ETW rows.
- Scrollback and tail mode state drift.
- Memory retention from line-layout cache.

### Mitigations

- Stable line sequence numbers and selftest assertions.
- Cache budgets by visible plus near-visible lines, not whole history.
- Separate append ingestion metrics from paint metrics.

---

## Candidate 5: Perf Lab With WPR, GPUView, And PresentMon

**Recommendation:** Required before any major GPU, present-policy, or composition work.

### Idea

Add scripts and documentation for repeatable external captures:

- WPR/WPA/GPUView ETL capture for CPU/GPU/DWM analysis.
- PresentMon CSV capture for frame pacing and presentation behavior.
- Correlation markers from app-local JSONL metrics to external traces.

### Why It Matters

App-local JSONL tells us what RedSalamander believes it spent. External traces tell us whether the bottleneck is:

- UI thread CPU,
- Direct2D command submission,
- GPU busy,
- present wait,
- DWM composition,
- device/resource creation,
- synchronization.

This should be mandatory before D3D batching, flip-discard experiments, or composition pilots.

### Required Artifacts

- `Specs/Testing/Testing_PerformanceValidation.md` update with capture recipe.
- Script under an appropriate tooling folder, if the project has one.
- Archive pointer under `Specs/TestRuns/<machine>/<area>/<runid>/`.
- A small summary file with app metrics plus external trace filenames.

### Gates

No production code changes in this candidate except instrumentation hooks. Accept when the capture process is reproducible and documented.

---

## Candidate 6: D2D/D3D Batching Layer

**Recommendation:** Medium priority, high risk. Do not start until Candidate 5 proves Direct2D draw submission is the bottleneck.

### Idea

For repeated simple primitives, replace many immediate Direct2D calls with batched geometry:

- row backgrounds,
- selection rectangles,
- hover rectangles,
- separators,
- simple icons/badges if atlas-backed later.

If D3D11 is used directly, dynamic vertex/index buffers should follow `WRITE_NO_OVERWRITE` while appending and `WRITE_DISCARD` when the buffer wraps.

### Required Metrics

- `render.draw_call_count`
- `render.state_change_count`
- `render.batch_count`
- `render.vertex_count`
- `render.map_us`
- `render.gpu_busy_us` if external capture is available

### Gates

Accept only if:

- draw-call or Direct2D submission cost is proven dominant, and
- p95 frame time improves by at least 15% in a large visible-work scenario, and
- no text quality, DPI, theme, or accessibility regression is introduced.

### Risks

- Complex renderer ownership.
- Text/icons still dominate, leaving geometry batching with low payoff.
- Harder debugging versus Direct2D primitives.

---

## Candidate 7: Narrow Composition Pilot For Proven Animation Bottlenecks

**Recommendation:** Low priority until evidence changes.

### Idea

Keep composition as an opt-in, isolated pilot only for transform/opacity animations where the animated content can be snapshotted as a bitmap surface and moved independently of the UI thread.

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

The 2026-05-20 gate rejected compositor code. Connected-overlay paint p95 was `11us`, frame render p95 stayed below the 8.333ms budget, and present p95 was `381us`; the observed issue was synthetic scheduler/timer jitter, not overlay rendering cost.

### Gates

Accept only if a future allowed-surface scenario shows:

- UI-thread animation p95 jitter above budget,
- rendering/paint work is a dominant cause,
- composition reduces jitter or frame p95 without increasing memory, input latency, or stale visuals.

---

## Recommended Sequencing

1. **Phase A: Measurement Backbone**
   - Add or refine external capture workflow.
   - Add missing text-layout and dirty-region metrics.
   - Do not optimize yet.

2. **Phase B: Dirty Region + Text Layout Cache**
   - Start with FolderView because it already has a measured dirty-area win.
   - Add text layout cache to a single high-value surface with strict budgets.
   - Validate against normal scroll, overlay, sort, theme, DPI, and rename scenarios.

3. **Phase C: Monitor Tail/Scrollback**
   - Add line-layout and visible-line reuse metrics.
   - Optimize append and scrollback only if gates show p95 pressure.

4. **Phase D: Advanced Rendering**
   - Consider D2D/D3D batching only after external traces show draw submission or GPU busy is the limiter.
   - Consider composition only for isolated animation surfaces after the existing gate changes.

---

## Proposed First Implementation Plan

The best first plan should be:

**FolderView Dirty Region And Text Layout Cache Pilot**

Reasons:

- FolderView is key to the product.
- The only measured recent runtime performance gain was in FolderView.
- It exercises both dominant ideas: smaller dirty regions and less text layout churn.
- It can be validated with existing `folderView_perf_overlay_invalidation_stress` and `folderView_perf_scroll_render_stress`, then expanded to 10k item and icon-heavy scenarios.

Minimum deliverables for that future plan:

- Add `dwrite.text_layout.*` metrics for FolderView.
- Add one deterministic 10k item FolderView scenario.
- Add a small bounded `IDWriteTextLayout` cache for stable FolderView item labels/details.
- Add explicit invalidation on rename, sort, view mode, DPI, theme, font, width, and metadata changes.
- Archive before/after evidence.

---

## Non-Goals

- Do not replace Direct2D/DirectWrite wholesale.
- Do not enable flip-discard full redraw for current surfaces.
- Do not ship DirectComposition broadly.
- Do not claim performance improvement from measurement-only work.
- Do not add unbounded caches.
- Do not add hot-path per-control telemetry.

---

## Open Decision

Before implementation planning, choose the first target:

1. FolderView huge folders and overlay/scroll performance.
2. Monitor high-rate ETW/log rendering and scrollback.
3. DxUi chrome, menus, and animation smoothness.

Recommended choice: **1. FolderView huge folders and overlay/scroll performance**.
