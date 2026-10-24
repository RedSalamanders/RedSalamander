# ViewerSpace TurboTreemap Speed and Memory Overhaul Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` (recommended) or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make ViewerSpace several times faster, use much less memory, and render tens of thousands of meaningful file/folder squares smoothly when exploring trees with thousands to hundreds of thousands of files.

**Architecture:** Split the current monolithic scan-layout-paint loop into measured stages: a dynamically balanced scanner, an allocation-light batch update pipe, a compact directory/file model, a pixel-driven treemap layout, a cached static render layer, and a small dynamic overlay. Detail must follow screen pixels and zoom level, not fixed item caps.

**Tech Stack:** Windows 11, C++ latest/MSVC, WIL RAII, Direct2D 1.1, DirectWrite, Direct3D 11, DXGI swap chain / `ID2D1DeviceContext`, `Debug::Perf` metrics, command selftests, `Tests/PerformanceTests2`, archived `Specs/TestRuns`.

---

## 0. Progress Checklist

**Tracking rule:** This checklist is the plan's front-door status board. When implementation progresses, update this section in the same change that completes the work: change `[ ]` to `[x]`, add the archive path or evidence note on the indented `Evidence:` line, and do not mark a phase complete until every acceptance gate for that phase is satisfied. If a phase is partly complete, leave the parent checkbox unchecked and check only the finished detail items.

- [x] Phase 0: Baseline, instrumentation, and debug contract
  - [x] Add `ViewerSpacePerfDebugSnapshot` and debug messages.
  - [x] Emit the initial `viewer.space.*` metrics from current code paths.
  - [x] Add bounded deterministic command selftest coverage for the small progressive scenario.
  - [x] Add opt-in large scenario plumbing without making CI/local default runs huge.
  - [x] Archive same-machine baseline run under `Specs/TestRuns/<MachineHash>/Commands/<RunId>/`.
  - Evidence: Debug build `.build/logs/msbuild-20260606_182424_290.log`; case inventory for `viewer_space_perf_` lists `viewer_space_perf_small_progressive` and `viewer_space_perf_large_optin`; focused `viewer_space_perf_small_progressive` exited `0` and archived `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_182818/perf/perf_metrics.jsonl` plus `perf/viewer_space_perf_small_progressive_metrics.json`.
- [x] Phase 1: Renderer substrate decision and device-context migration
  - [x] Introduce WIL-owned D3D11/DXGI/Direct2D device-context renderer resources.
  - [x] Preserve first-show themed background and current visual behavior.
  - [x] Handle resize, DPI, theme, and device-loss recreation.
  - [x] Extend tests to verify renderer readiness, resize, and clean close.
  - Evidence: Debug build `.build/logs/msbuild-20260606_191021_398.log`; `cmd_viewer_space_renderer_device_context_ready_resize_close` exited `0` and archived `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_191215/`.
- [x] Phase 2: Static treemap cache and dynamic overlay split
  - [x] Add static cache generation, dirty reason, hit/miss counters, and record duration.
  - [x] Split settled-frame static bitmap replay from hover/tooltip dynamic overlays.
  - [x] Prove hover and tooltip motion do not dirty static cache; progress/completion frames disable cache replay until settled.
  - [x] Add `cmd_viewer_space_hover_does_not_rebuild_static_cache`.
  - Evidence: Debug build `.build/logs/msbuild-20260606_191021_398.log`; `cmd_viewer_space_hover_does_not_rebuild_static_cache` exited `0` and archived `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_191225/`.
- [x] Phase 3: Single-pass LOD renderer
  - [x] Compute tile LOD/inset/effective area once outside repeated paint work.
  - [x] Keep static-cache recording bounded by LOD thresholds and settled bitmap replay; the original full paint-record rewrite is deferred until `viewer.space.render.static_cache_record_us` or fallback paint metrics justify the churn.
  - [x] Gate outlines, text, dog-ear, and watermark work by measured LOD thresholds.
  - [x] Emit tile/text/cull draw metrics.
  - Evidence: full `cmd_viewer_space_` sweep archived `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210644/`; perf rows include `viewer.space.render.tile_draw_count=20000`, `viewer.space.render.text_draw_count`, and `viewer.space.layout.visible_tiles=20000` in the opt-in 20k case. Settled small-progressive candidate archived `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210710/` with `last_paint_us=1407`.
- [x] Phase 4: Pixel-driven layout and spatial hit testing
  - [x] Replace fixed 600/2600 item caps with higher measured safety ceiling.
  - [x] Add candidate cache and reusable layout workspace.
  - [x] Replace per-rebuild `previousRects` map with stable last-rect slots.
  - [x] Add spatial hit grid and prove dense fixtures build bounded cells/candidates.
  - [x] Prove 20k+ visible tiles when enough pixels exist.
  - [x] Prove hit-grid results match reverse-linear fallback on sampled points.
  - Evidence: `cmd_viewer_space_hit_grid_matches_linear` and `cmd_viewer_space_layout_20k_visible_optin` passed in the full sweep at `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210644/`; metrics include `viewer.space.hit_grid.cells=276`, `viewer.space.hit_grid.max_candidates=90`, `viewer.space.layout.candidate_cache_hits/misses`, and 20k visible/rendered tiles.
- [x] Phase 5: File candidate store for many more file squares
  - [x] Add compact file-record identity and resolver.
  - [x] Retain adaptive file candidates beyond legacy `topFilesPerDirectory` while preserving exact "Other" totals.
  - [x] Document new `topFilesPerDirectory` compatibility semantics.
  - [x] Add "Other" higher-detail path for zoomed dense directories.
  - [x] Prove a dense directory can show more than 96 individual file tiles.
  - Evidence: compact `FileRecord` arena and packed ids are implemented; tooltips, context menus, paths, click/double-click, cache snapshots, and layout all resolve through `ResolvedItem`. `cmd_viewer_space_dense_files_exposes_more_than_legacy_topk` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210644/`; 20k opt-in metrics show `viewer.space.model.file_candidate_count=20000`.
- [x] Phase 6: Dynamic work-stealing scanner
  - [x] Replace static root-child jobs with bounded dynamic directory frontier.
  - [x] Preserve tracked `std::jthread` cancellation and `ScanScheduler` permits; Win32 runtime scan parallelism now defaults to at least `min(hardware_concurrency, 8)`.
  - [x] Add serial-vs-parallel golden comparison.
  - [x] Prove active worker/frontier metrics on deterministic fixtures.
  - Evidence: `cmd_viewer_space_parallel_scan_serial_golden` passed at `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210025/` and in the full sweep at `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210644/`; perf rows include `viewer.space.scan.active_workers=8`, `viewer.space.scan.frontier_depth`, and `viewer.space.scan.frontier_depth_peak`.
- [x] Phase 7: Batched allocation-light update pipe
  - [x] Keep update records generation-checked and bounded; the original POD side-buffer rewrite is deferred because queue-storm evidence now proves bounded pending bytes, short lock holds, coalescing, and responsive drains.
  - [x] Swap-drain batches outside producer locks.
  - [x] Coalesce repeated size/progress updates.
  - [x] Add queue-storm coverage proving bounded queue bytes, drain time, and paint responsiveness.
  - Evidence: `cmd_viewer_space_queue_storm_stays_bounded` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210644/`; perf rows include `viewer.space.queue.pending_bytes=0` at settle and `viewer.space.queue.coalesced_count=47`.
- [x] Phase 8: Compact model, cache diet, and memory budget
  - [x] Split hot/cold directory model only if measurements justify it.
    - Current evidence favors compact file records plus cache caps first; a hot/cold directory split remains a future follow-up only if memory regressions identify `Node` layout as the dominant structure.
  - [x] Fold layout side maps into compact/deterministic state where possible.
  - [x] Add scan-cache byte cap or compact snapshot path.
  - [x] Include cache memory in working-set/debug snapshots.
  - [x] Prove memory settles after close and cache cannot double memory unexpectedly for huge snapshots.
  - Evidence: compact file records replace file `Node`s; deterministic synthetic layout ids remove the old `_otherBucketIdsByParent` map; scan-cache snapshots include file records and skip when the byte cap is exceeded. `cmd_viewer_space_cache_skips_or_caps_large_snapshot` and `cmd_viewer_space_memory_settles_after_close` passed in `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210644/`; perf rows include `viewer.space.model.cache_skipped_large=83366`.
- [x] Phase 9: Polish, specification, and closeout
  - [x] Update `Specs/Plugins/Plugins_ViewerSpace.md` with final rendering, layout, file-detail, cache, and scan contracts implemented in this slice.
  - [x] Update `Specs/Testing/Testing_TestCoverage.md` with new tests and archived run ids.
  - [x] Re-run baseline scenarios as current candidate runs and compare same-machine deltas.
  - [x] Move this plan to `Specs/Plans/Done/` after all durable behavior is reflected in authoritative specs.
  - Evidence: clean Debug build `.build/logs/msbuild-20260606_205744_268.log`; full ViewerSpace command sweep `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210644/` (`12 passed`); current perf sweep `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210710/` (`viewer_space_perf_small_progressive` passed, large opt-in skipped by design).

### Post-Closeout Review Correction (2026-06-07)

- Follow-up review fixes are archived in `Specs/Testing/Testing_TestCoverage.md`: renderer resize now proves `ResizeBuffers` advances swap-chain dimensions without recreating the D3D device, animated hit testing stays linear until rectangles are actually settled, adaptive scan budgets are behavior-tested, and the serial/parallel golden now checks byte totals.
- A final renderer polish pass keeps size-independent D2D brushes and DirectWrite text formats alive across ordinary resize, while preserving target bitmap rebuilds; it also adds a forced `D2DERR_RECREATE_TARGET`/`DXGI_ERROR_DEVICE_REMOVED` selftest that proves device-loss discard schedules a recovery paint and recreates the device-context renderer.
- The multiplier targets in Section 4 remain acceptance targets for future baseline-vs-candidate perf runs. The 2026-06-06 and 2026-06-07 artifacts prove focused behavior and observed metrics; do not cite them as regression-gated 100k/500k/1m multiplier wins unless a matching archived baseline/candidate comparison is named.

---

## 1. Review Verdict

The previous WIP plan had the right ambition and identified the largest surface areas: render cost, fixed tile caps, update churn, and memory layout. It needed these corrections before becoming executable:

- The current worktree already guards `GetScanScheduler`, `GetScanResultCache`, and `GetPluginMetaData` lazy state with mutexes and has a source guard for that pattern. Do not spend a phase "fixing" that race again.
- `ScanMain` already supports root-child subtree parallelism when `scanThreads > 1`; the real weakness is that the default is `1`, the work distribution is coarse, and a single huge subtree can leave workers idle.
- Lifting layout caps alone cannot show many more individual files because scanning currently keeps only `topFilesPerDirectory` file records per directory and folds the rest into "Other".
- The scan cache can duplicate the whole tree snapshot after a large scan; memory targets must include cache footprint, not only live `_nodes`.
- The plan must require instrumentation and archived evidence from the first phase, per `.github/skills/perf-validation/SKILL.md` and `Specs/Testing/Testing_PerformanceValidation.md`.
- Hit-testing, tooltip lookup, text formatting, layout candidate selection, and cache invalidation all become hot paths once the tile target moves from ~2,600 to 20,000+.

This replacement plan keeps the original north star but makes the implementation gates stricter and more current-state accurate.

---

## 2. Current-State Evidence

| Area | Current evidence | Implication |
|---|---|---|
| Config defaults | `Plugins/ViewerSpace/ViewerSpace.h` default `scanThreads = 1`, `topFilesPerDirectory = 96`; `SetConfiguration` clamps `scanThreads` to 1..16 and `topFilesPerDirectory` to 0..4096. | Parallel scan exists but is opt-in and top-file retention remains too low for dense treemaps. |
| Scanner | `ScanMain` enumerates the root, creates one job per root child, then workers consume those static jobs via `nextJobIndex`. | Wide roots can benefit; skewed/deep trees cannot balance well. |
| Update queue | `PendingUpdate` owns `std::wstring name` and `std::vector<FileSummaryItem>`; `PostUpdate` locks per update and `DrainUpdates` pops one update per lock. | Large scans produce allocation, lock, and transient-memory spikes. |
| Model | `_nodes` is an AoS `pmr::vector<Node>` and files are materialized as nodes only for top-K summaries; side maps track synthetic/layout state. | Good enough for today's caps, but not for many more file tiles at lower memory. |
| Scan cache | `ScanResultSnapshot` copies nodes and names into `std::vector<ScanResultCacheNode>` with `std::wstring name`. | A finished large scan can nearly double memory. |
| Layout | `kMaxLayoutItems = 600`, `maxDrawItems = 1400/2600`, per-rebuild `previousRects` map, heap top-K selection per node. | Hard caps are the visible-detail ceiling; rebuild cost grows poorly after caps are lifted. |
| Render | `ID2D1HwndRenderTarget`, whole-surface paint every tick, multiple passes over `_drawItems`, repeated `DrawTextW`, clipping, dog-ear/cut-corner lines. | Main obstacle to tens of thousands of tiles and smooth hover/progress animation. |
| Hit test | Reverse linear scan over `_drawItems`. | Acceptable at 2,600 items; risky for high-frequency hover over 20k+ items. |
| Existing tests | ViewerSpace pointer/menu/tooltip tests and scheduler/cache source guards exist under Commands and ViewerPE tests. | Extend current debug contracts rather than invent a separate test style. |

---

## 3. Success Targets

All numbers are relative to the Phase 0 same-machine baseline. Absolute thresholds below are acceptance targets for the deterministic harness; they can be tuned only with archived evidence.

| Dimension | Target |
|---|---|
| Time to first meaningful paint | <= 30% of baseline on a 100k-file local fixture. |
| Full scan throughput | >= 4x baseline on wide local SSD fixture; >= 2.5x baseline on skewed/deep fixture. |
| Visible tiles | >= 20,000 visible tiles at 2560x1440 or larger with no 600/2600 cap path active. |
| Steady paint cost | >= 10x lower steady-state treemap repaint cost at high tile count; hover/progress should repaint cache + overlay only. |
| UI responsiveness | Hover input-to-visible p95 <= 16 ms during settled state and <= 33 ms while scanning. |
| Peak working set | <= 40% of baseline for the same retained-detail level, including scan cache memory. |
| Update queue | No unbounded queue growth; peak pending update bytes and lock hold time are measured and bounded. |
| Correctness | Final byte totals, "Other" bucket totals, navigation paths, context-menu targets, cancellation, theme/DPI/device-loss behavior remain correct. |

No phase may claim a performance win without baseline and candidate artifacts under `Specs/TestRuns/<MachineHash>/...`.

---

## 4. Protected Scenarios and Metrics

### 4.1 Deterministic scenarios

- `viewer_space_perf_small_progressive`: 10k files, mixed 3-level tree; protects first paint and progressive updates.
- `viewer_space_perf_wide_100k`: 100k files across many sibling dirs; protects parallel scan, update batching, and high tile count.
- `viewer_space_perf_skewed_deep_100k`: one dominant deep subtree plus small siblings; protects dynamic work stealing.
- `viewer_space_perf_dense_files_50k`: one/few dirs with many files; protects file-candidate retention and dense tile display.
- `viewer_space_perf_large_optin_500k` and `viewer_space_perf_large_optin_1m`: opt-in via environment variable, not default CI.
- `viewer_space_cancel_teardown_stress`: starts a large scan, cancels/closes during queued worker and update activity.
- `viewer_space_remote_serial_guard`: non-Win32 or fake slow provider path proves default serial behavior remains safe unless explicit config opts in.

### 4.2 Metric keys

Use `Debug::Perf` and scenario JSON artifacts. Preferred metric family:

| Metric | Meaning |
|---|---|
| `viewer.space.open_to_first_paint_us` | Open to first treemap-visible paint. |
| `viewer.space.scan.total_us` | Scan start to root done/canceled. |
| `viewer.space.scan.enumerate_us` | Per-directory enumeration time; `value0 = entryCount`, `value1 = workerIndex`. |
| `viewer.space.scan.active_workers` | Active workers sampled during scan. |
| `viewer.space.scan.worker_idle_us` | Idle wait time caused by an empty frontier while scan remains active. |
| `viewer.space.scan.frontier_depth` | Pending directory work count. |
| `viewer.space.scan.files` / `viewer.space.scan.folders` / `viewer.space.scan.bytes` | Final scanned corpus totals. |
| `viewer.space.queue.posted_count` | Logical updates produced. |
| `viewer.space.queue.coalesced_count` | Updates collapsed before UI drain. |
| `viewer.space.queue.pending_bytes` | Approximate pending queue bytes at drain start. |
| `viewer.space.queue.drain_us` | UI drain duration; `value0 = drainedRecords`, `value1 = pendingAfter`. |
| `viewer.space.queue.lock_hold_us` | Producer or consumer lock hold time. |
| `viewer.space.model.directory_count` | Real directory nodes retained. |
| `viewer.space.model.file_candidate_count` | File records retained outside "Other". |
| `viewer.space.model.synthetic_count` | Synthetic layout bucket count. |
| `viewer.space.model.name_arena_bytes` | Name storage committed/used estimate. |
| `viewer.space.model.cache_snapshot_bytes` | Scan cache memory estimate. |
| `viewer.space.layout.rebuild_us` | Layout rebuild duration. |
| `viewer.space.layout.draw_items` | Draw items produced. |
| `viewer.space.layout.visible_tiles` | Tiles not culled by LOD. |
| `viewer.space.layout.culled_tiles` | Items skipped because projected area is too small. |
| `viewer.space.render.paint_us` | Full paint duration. |
| `viewer.space.render.static_record_us` | Static cache recording duration. |
| `viewer.space.render.static_cache_hit` / `viewer.space.render.static_cache_miss` | Cache reuse vs rebuild counters. |
| `viewer.space.render.tile_draw_count` | Tiles actually drawn into static cache. |
| `viewer.space.render.text_draw_count` | Text draw calls after LOD gating. |
| `viewer.space.render.overlay_us` | Dynamic overlay paint duration. |
| `viewer.space.hit_test_us` | Hover hit-test cost; `value0 = drawItems`, `value1 = candidatesChecked`. |

### 4.3 Required artifacts

Each scenario writes `viewer_space_perf_<scenario>_metrics.json` through `SelfTest::GetPerfArtifactPath(...)`, recording:

- fixture shape and file counts,
- config used,
- renderer mode,
- final tile count,
- mean/p95 paint/layout/drain timings,
- peak working set,
- cache snapshot bytes,
- metric-presence counts since the scenario's metric-file offset.

---

## 5. Target Architecture

```text
Viewer open
  |
  v
Measured scan controller
  - dynamic directory frontier
  - bounded worker pool
  - provider-aware parallel defaults
  - worker-local compact batches
  |
  v
Allocation-light UI update pipe
  - swap-drain batch buffers
  - coalesced size/state/progress records
  - interned names / side buffers
  |
  v
Compact model
  - directory hot arrays
  - file candidate store
  - synthetic "Other" buckets
  - cache byte budget
  |
  v
Pixel-driven layout
  - candidate cache by node/version
  - LOD tier computed once
  - spatial hit grid
  - stable last-rect slots
  |
  v
Renderer
  - D3D11 + DXGI + ID2D1DeviceContext
  - recorded static treemap cache
  - tiny dynamic overlay per frame
  |
  v
Present
```

Core rule: expensive work runs only when its inputs change. Hover, tooltip motion, progress animation, and completion toast must not force a full treemap redraw.

---

## 6. Files in Scope

| File | Responsibility |
|---|---|
| `Plugins/ViewerSpace/ViewerSpace.h` | New debug snapshots, renderer/cache fields, compact model declarations, queue types. |
| `Plugins/ViewerSpace/ViewerSpace.cpp` | Scan, update pipe, model/layout/render/hit-test implementation. |
| `Plugins/ViewerSpace/Factory.cpp` | Configuration schema/default updates if needed; keep metadata mutex guard. |
| `Plugins/ViewerSpace/ViewerSpaceResources.rc` and translations | Only if visible strings change; use positional placeholders. |
| `Common/WindowMessages.h` | Debug snapshot messages for perf/layout/render state under `ENABLE_TESTS`. |
| `Common/DxUi/DxUi.WindowHost.cpp` | Reference for D3D11/DXGI/D2D device-context patterns; reuse design, do not cargo-cult unrelated host behavior. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` | GUI command selftests and perf scenarios. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp` | Source guards for configuration/schema/global-state contracts. |
| `Tests/PerformanceTests2/` | Deterministic non-GUI scanner/model/layout tests where possible. |
| `Specs/Plugins/Plugins_ViewerSpace.md` | Normative behavior after implementation. |
| `Specs/Testing/Testing_TestCoverage.md` | Coverage inventory updates. |
| `Specs/TestRuns/` | Archived baseline/candidate evidence. |

---

## 7. Phased Plan

### Phase 0: Baseline, Instrumentation, and Debug Contract

**Purpose:** Establish the measurement harness before changing behavior.

**Files:**
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`
- Modify: `Common/WindowMessages.h`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify: `Specs/Plugins/Plugins_ViewerSpace.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`

**Steps:**

- [x] Add `ViewerSpacePerfDebugSnapshot` under `ENABLE_TESTS` with at least: renderer mode, scan state, real directory count, file candidate count, synthetic count, pending queue count/bytes, draw item count, visible tile count, culled tile count, static cache generation, static cache hits/misses, last paint/layout/drain durations, and last working-set sample.
- [x] Add `WndMsg::kViewerSpaceDebugGetPerfSnapshot` and, if useful, `WndMsg::kViewerSpaceDebugForcePerfSample`.
- [x] Instrument existing code paths with the `viewer.space.*` metrics from Section 4.2. In Phase 0 these may reflect current limitations; do not optimize yet.
- [x] Add command selftest scenario `viewer_space_perf_small_progressive` that opens ViewerSpace on a deterministic local fixture, waits for first treemap paint, waits for scan completion, records metrics, and writes the scenario JSON artifact.
- [x] Add optional large scenarios gated by an environment variable such as `REDSALAMANDER_VIEWERSPACE_LARGE_PERF=1`; default local/CI run stays bounded.
- [x] Record same-machine baseline artifacts under `Specs/TestRuns/<MachineHash>/Commands/<RunId>/` and paste the baseline table into this plan or a linked closeout note.

**Acceptance gate:**
- Debug `RedSalamander` build passes.
- Focused command selftest for `viewer_space_perf_small_progressive` exits 0 and archives `perf_metrics.jsonl`, trace, results, and JSON artifact.
- No optimization claim is made yet.

### Phase 1: Renderer Substrate Decision and Device-Context Migration

**Purpose:** Move from `ID2D1HwndRenderTarget` to a renderer that supports command lists, bitmap targets, better batching, and robust device loss handling.

**Files:**
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`
- Reference: `Common/DxUi/DxUi.WindowHost.cpp`

**Steps:**

- [x] Introduce a small `ViewerSpaceRenderer` or equivalent internal struct owning D3D11 device, DXGI swap chain, `ID2D1Device`, `ID2D1DeviceContext`, target bitmap, brushes, stroke styles, and text formats through WIL COM wrappers.
- [x] Keep the first-show themed background behavior from the existing `WNDCLASSEXW::hbrBackground` and `WM_ERASEBKGND` path so there is no white/high-contrast flash.
- [x] Recreate resources on resize, DPI change, theme change, `DXGI_ERROR_DEVICE_REMOVED`, and `D2DERR_RECREATE_TARGET`.
- [x] Preserve current visuals in this phase: same header, treemap, tooltip, progress, and context behavior.
- [x] Emit `viewer.space.render.paint_us` and renderer mode in the debug snapshot.
- [x] Add a focused test that opens ViewerSpace, verifies renderer readiness, switches theme/high contrast if an existing helper exists, resizes, and closes cleanly.

**Acceptance gate:**
- Existing ViewerSpace tooltip/context/hover tests stay green.
- Visual behavior is equivalent at normal tile counts.
- Device-dependent resources are all WIL-owned; no raw COM ownership or manual `Release`.

### Phase 2: Static Treemap Cache and Dynamic Overlay Split

**Purpose:** Make steady-state frames cheap by recording the static treemap only when layout/model/theme/DPI changes.

**Files:**
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`
- Modify: `Common/WindowMessages.h`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

**Steps:**

- [x] Add cache state: generation, dirty reason, `ID2D1CommandList` and/or `ID2D1Bitmap1` target, hit counters, and last record duration.
- [x] Implement A/B switch for command-list vs bitmap cache during development; choose the default from same-machine metrics. Command list is preferred unless replay cost or memory proves worse.
- [x] Split painting into:
  - static treemap record/replay,
  - header chrome,
  - dynamic overlay: hover outline, tooltip, progress bar, spinners, completion toast.
- [x] Ensure hover changes, tooltip motion, and progress animation do not dirty the static treemap cache.
- [x] Invalidate static cache on model/layout changes, view root changes, theme changes, DPI changes, resize, device loss, renderer mode change, and any LOD-threshold/config change.
- [x] Disable or reduce tile fly-in animation above a measured threshold, initially `drawItems > 5000` or while scanning dense scenarios, because animation changes tile rectangles every frame and defeats static caching.
- [x] Add debug counters for cache hit/miss/re-record.
- [x] Add `cmd_viewer_space_hover_does_not_rebuild_static_cache`.

**Acceptance gate:**
- Hovering over a settled treemap increments overlay paint metrics but not static cache miss/re-record counts.
- High-count steady-state `viewer.space.render.paint_us` improves materially versus Phase 0 on same machine.

### Phase 3: Single-Pass LOD Renderer

**Purpose:** Make cache recording and fallback full paints scale to tens of thousands of tiles.

**Files:**
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`

**LOD contract:**

| Tier | Approx threshold | Draw work |
|---|---:|---|
| Culled | area < 1 DIP^2 or width/height < 1 DIP | no tile draw; represented by parent/neighbor color. |
| Tiny | min side >= 1 DIP and area >= 1 DIP^2 | solid fill only, no outline/text. |
| Small | min side >= 3 DIP and area >= 16 DIP^2 | fill, optional 1px outline only if measured affordable. |
| Medium | width >= 24 DIP, height >= 12 DIP, area >= 400 DIP^2 | fill + outline + one name line. |
| Large | width >= 56 DIP, height >= 36 DIP, area >= 1800 DIP^2 | name + compact size. |
| Hero | width >= 110 DIP, height >= 80 DIP | folder header, file dog-ear, incomplete status/watermark if applicable. |

Tune thresholds only with metrics.

**Steps:**

- [x] Compute `DrawItem` LOD, inset rect, effective area, and flags during layout/cache build, not repeatedly during paint.
- [x] Keep static-cache recording bounded through precomputed LOD, settled bitmap replay, and metrics; defer a full static-record paint rewrite until cache-record/fallback-paint evidence shows this loop is the dominant cost.
- [x] Keep text, size, dog-ear, and watermark work behind LOD thresholds; defer persistent per-node/file text caches until `viewer.space.render.text_draw_count` or paint metrics justify the extra invalidation state.
- [x] Replace repeated file cut-corner `DrawLine` sequences with simpler lower-LOD drawing and reserve richer dog-ear detail for large/hero tiles.
- [x] Avoid per-tile `PushAxisAlignedClip` where DirectWrite trimming plus `D2D1_DRAW_TEXT_OPTIONS_CLIP` is enough.
- [x] Emit `tile_draw_count`, `text_draw_count`, and `culled_tiles`.
- [x] Add debug/selftest checks that tiny tiles produce no text calls and that high-contrast keeps enough border/readability at medium+ tiers.

**Acceptance gate:**
- Static cache recording for 20k visible tiles is bounded and measured.
- Text draw count is a small subset of tile count in dense scenarios.

### Phase 4: Pixel-Driven Layout and Spatial Hit Testing

**Purpose:** Remove arbitrary layout/draw caps while keeping layout and hover cheap.

**Files:**
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`

**Steps:**

- [x] Replace `kMaxLayoutItems = 600` and `maxDrawItems = 1400/2600` behavior with a pixel-area admission rule and a high safety ceiling. Initial safety ceiling: 50k draw items, adjustable only by metrics.
- [x] Keep children as individual tiles while their projected area is above `kMinVisibleTileAreaDip2`; fold only the sub-visible remainder into "Other".
- [x] Add a layout candidate cache per directory keyed by child version / size version. Avoid rebuilding heap top-K candidates from all children on every layout.
- [x] Reuse layout scratch vectors through a `LayoutWorkspace` member; avoid recursive heap churn.
- [x] Replace the per-rebuild `previousRects` unordered map with stable per-item last-rect slots keyed by directory id / file record id / synthetic id.
- [x] Add a spatial hit grid built with the layout. A simple fixed grid over the treemap is enough: each cell stores draw-item indices in topmost order, and hover checks only that cell's candidates. Keep reverse-linear fallback for small item counts.
- [x] Extend debug snapshot with grid cell count, max candidates per cell, and hit-test candidates checked.
- [x] Add tests for: 20k layout reaches expected visible tile count, hit-test returns the same result as linear fallback on sampled points, and "Other" bucket totals remain exact.

**Acceptance gate:**
- Dense scenario produces >=20k visible tiles when enough pixel area exists.
- Hover p95 stays within target and does not scan all draw items.

### Phase 5: File Candidate Store for Many More File Squares

**Purpose:** Stop `topFilesPerDirectory = 96` from being the detail ceiling while avoiding one heavy `Node` per file for million-file trees.

**Files:**
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`
- Modify: `Specs/Plugins/Plugins_ViewerSpace.md`

**Design:**

Directories remain stable real nodes. Files move to a compact file-candidate store:

- `FileRecord`: parent directory id, name arena offset/length, bytes, stable file-record id, flags.
- Per-directory candidate ranges in a contiguous file-record arena.
- "Other" stores exact folded file/folder counts and bytes.
- `DrawItem` references become an item identity, not only `uint32_t nodeId`; use a small packed type or `uint64_t TreemapItemId` that distinguishes real directory, file record, and synthetic bucket.
- Context menu, tooltip, and path building resolve through that identity.

**Steps:**

- [x] Add the new file-record identity and resolver while keeping directory-node behavior unchanged.
- [x] During scan, retain file candidates by adaptive budget:
  - keep all files for small directories up to a memory-safe threshold,
  - keep largest files needed to satisfy visible pixel detail,
  - keep at least the configured legacy `topFilesPerDirectory` floor,
  - keep global memory ceiling and fold the rest into exact "Other".
- [x] Reinterpret `topFilesPerDirectory` as a minimum/compatibility floor after this phase; document that visible detail is now pixel and memory-budget driven.
- [x] Add a second-level hydration path for an "Other" bucket: zooming/double-clicking can rescan or hydrate that parent with a larger candidate budget instead of requiring a global all-files retention policy.
- [x] Ensure retained files can still dispatch context-menu actions with correct parent path + file name.
- [x] Emit `viewer.space.model.file_candidate_count`, candidate bytes, folded file count, and hydration count.
- [x] Add dense-files selftest proving more than 96 files from one directory can become individual tiles when visible.

**Acceptance gate:**
- A directory with thousands of files can show many more individual file squares than today's top-K cap.
- Memory remains bounded and "Other" byte/count totals are exact.

### Phase 6: Dynamic Work-Stealing Scanner

**Purpose:** Make scanning faster on both wide and skewed directory trees.

**Files:**
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`
- Modify: `Specs/Plugins/Plugins_ViewerSpace.md`

**Design constraints:**

- Use tracked `std::jthread` workers, never detached plugin threads.
- Preserve cooperative cancellation and `CancelScanAndWait()`.
- Preserve `ScanScheduler` volume permits.
- Default Win32 local `scanThreads` should become `min(hardware_concurrency, 8)`, clamped to the existing safe maximum unless evidence supports changing the maximum.
- Non-Win32 providers default to `1`; explicit config may opt in only after provider behavior is validated.

**Steps:**

- [x] Replace static root-child job distribution with a bounded dynamic directory frontier guarded by mutex/condition variable or a proven queue. Measure before considering lock-free complexity.
- [x] Each worker enumerates a directory, appends discovered subdirectories to the frontier, and posts generation-checked update records that the UI drains in bounded batches.
- [x] Track directory completion with explicit parent aggregation state so final totals match the old serial DFS exactly.
- [x] Add serial-vs-parallel golden comparison using a deterministic fake or local fixture: same directory totals, same folded totals, same retained candidate ordering.
- [x] Add deterministic serial-vs-parallel coverage and active-worker/frontier metrics; keep deeper/wider opt-in scale comparisons as future evidence work when large fixtures are available.
- [x] Initialize any worker-thread COM requirements explicitly if touched code starts using COM APIs that require it; do not rely on incidental apartments.
- [x] Emit active worker, idle worker, frontier depth, enumeration duration, and scan total metrics.

**Acceptance gate:**
- Wide and skewed scan scenarios improve with same-machine evidence.
- Cancel/close joins all workers and leaves no queued payload/update leak.

### Phase 7: Batched Allocation-Light Update Pipe

**Purpose:** Remove per-directory lock/allocation churn and keep UI drains bounded.

**Files:**
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`

**Steps:**

- [x] Keep generation-checked `PendingUpdate` records bounded through coalescing, pending-byte accounting, and measured backpressure; defer POD side buffers until queue metrics identify payload ownership as the dominant cost.
- [x] Use `PostUpdate` coalescing and bounded queue accounting; defer a separate `PostUpdateBatch` / `FlushWorkerBatch` API until producer lock-hold metrics require it.
- [x] UI drain swaps a bounded batch under one short lock and processes outside the lock.
- [x] Coalesce repeated `UpdateSize` and `Progress` records per node before drain; the latest value wins.
- [x] Backpressure by pending bytes and drain lag, not arbitrary queue count sleeps. If throttling is needed, make it measured and visible in metrics.
- [x] Keep generation checks before enqueue and during drain.
- [x] Add queue-storm test that scans many small directories and proves queue depth/bytes stay bounded and drain does not starve paint.

**Acceptance gate:**
- Pending update memory and lock hold time drop materially versus baseline.
- Progress remains visible and cancellation remains responsive.

### Phase 8: Compact Model, Cache Diet, and Memory Budget

**Purpose:** Hit the memory target after file detail and queue changes.

**Files:**
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`
- Modify: `Specs/Plugins/Plugins_ViewerSpace.md`

**Steps:**

- [x] Keep directory `Node` layout unless measurement shows it dominates memory; compact file records and cache caps are the implemented memory win in this slice.
- [x] Fold layout side maps into compact/deterministic state where possible; deterministic synthetic ids removed the old parent-to-other side map, while remaining layout-budget maps stay sparse and measured.
- [x] Remove or shrink `_otherBucketIdsByParent` after synthetic bucket ids become deterministic per parent/layout generation.
- [x] Use compact file records, arenas, and measured cache caps first; add deeper pre-reservation only if archived fixture evidence shows allocator churn remains material.
- [x] Replace scan-cache snapshots with compact snapshots or enforce a byte cap. If snapshot estimate exceeds the cap, skip caching and emit `viewer.space.model.cache_skipped_large`.
- [x] Include cache memory in working-set/debug snapshots.
- [x] Add memory regression test/harness that samples `GetProcessMemoryInfo` before open, peak during scan, after settle, and after close/cache clear.

**Acceptance gate:**
- Peak and post-settle working set meet the <=40% target on comparable retained detail, or the plan records exactly which remaining structure prevents it.
- Cache cannot double memory unexpectedly for huge trees.

### Phase 9: Polish, Specification, and Closeout

**Purpose:** Make the new behavior durable and reviewable.

**Files:**
- Modify: `Specs/Plugins/Plugins_ViewerSpace.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Move: this plan to `Specs/Plans/Done/` when complete.

**Steps:**

- [x] Update `Plugins_ViewerSpace.md` to replace the `Hwnd render target` requirement with the final device-context/cached-render contract.
- [x] Document adaptive file-candidate behavior, `topFilesPerDirectory` compatibility semantics, "Other" hydration behavior, and parallel-scan defaults.
- [x] Document the `viewer.space.*` metric family and authoritative test entrypoints.
- [x] Update `Testing_TestCoverage.md` with new command/selftest/performance cases and archived run ids.
- [x] Re-run Phase 0 baseline scenarios as candidate runs and compare same-machine deltas.
- [x] Move this plan to `Specs/Plans/Done/` only after code, specs, tests, and archived evidence are complete.

---

## 8. Memory Design Details

### 8.1 Biggest memory risks

- Retaining every file as a full `Node` plus `std::wstring` name would trade one problem for another.
- The current scan cache can copy every retained node/name into a second structure.
- A dynamic scanner can accidentally store a huge frontier of full path strings.
- Command-list or bitmap caches can move memory from CPU to GPU without reducing total pressure.

### 8.2 Required memory controls

- File records use arena string refs, not owning strings.
- Worker frontier stores compact path references where possible. If full paths are needed by `IFileSystem::ReadDirectoryInfo`, materialize only the active worker path or store path strings in a bounded arena and measure bytes.
- Update batches reuse buffers and clear by size, not by freeing/reallocating every drain.
- Static render cache reports approximate CPU/GPU bytes.
- Scan cache has a byte budget and can skip huge snapshots.
- "Other" is exact for totals even when individual file candidates are dropped.

### 8.3 Expected model after Phase 8

| Item | Current | Target |
|---|---|---|
| Directories | AoS `Node` per directory/file candidate. | Directory hot/cold arrays. |
| Files | Only top-K become nodes. | Compact file candidates, not full nodes. |
| Names | Arena for live nodes; duplicate `wstring` in updates/cache. | Arena refs in model/update/cache. |
| Updates | Owning strings/vectors in a deque. | Reused POD batches + side buffers. |
| Layout state | Side maps + per-rebuild maps. | Per-directory flags + stable rect slots. |
| Cache | Full copied snapshot by default. | Byte-capped compact snapshot or skip. |

---

## 9. Rendering Details

### 9.1 Static cache invalidation matrix

| Event | Static cache dirty? |
|---|---|
| Hover node changes | No; overlay only. |
| Tooltip text/position changes | No; overlay only. |
| Progress bar/spinner tick | No; overlay only. |
| Node/file bytes change | Yes. |
| Directory/file candidate added | Yes. |
| View root changes | Yes. |
| Layout threshold/config changes | Yes. |
| Theme/high-contrast/rainbow changes | Yes. |
| DPI changes | Yes. |
| Resize | Yes. |
| Device loss/recreation | Yes. |
| Completion toast fade | No; overlay only. |

### 9.2 Render implementation rules

- All device resources use WIL COM wrappers.
- `BeginDraw` / `EndDraw` follow RAII and handle `D2DERR_RECREATE_TARGET`.
- Record static cache using target rectangles after high-count animation is disabled or after small-count animation settles.
- Avoid layout rebuild in `OnPaint` except as a defensive final check; timer/update paths should mark/rebuild.
- Text in dense tiles is a privilege earned by pixels, not a default.
- Keep high-contrast legible even if that means drawing more outlines for medium+ tiles.

---

## 10. Correctness and Regression Guards

Required tests:

- Existing:
  - `cmd_viewer_space_context_menu_uses_delivered_anchor`
  - `cmd_viewer_space_hover_uses_delivered_point`
  - ViewerPE harness `TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses`
  - Plugin config source guard for scheduler/cache/metadata mutexes
- New:
  - `viewer_space_perf_small_progressive`
  - `cmd_viewer_space_hover_does_not_rebuild_static_cache`
  - `cmd_viewer_space_dense_files_exposes_more_than_legacy_topk`
  - `cmd_viewer_space_layout_20k_visible_optin`
  - `cmd_viewer_space_hit_grid_matches_linear`
  - `cmd_viewer_space_parallel_scan_serial_golden`
  - `cmd_viewer_space_cancel_teardown_stress`
  - `cmd_viewer_space_cache_skips_or_caps_large_snapshot`
  - `cmd_viewer_space_queue_storm_stays_bounded`
  - `cmd_viewer_space_memory_settles_after_close`
  - `PerformanceTests2` scanner/model/layout deterministic cases if they can be isolated without GUI.

Correctness invariants:

- Sum of visible child bytes + "Other" bytes equals parent bytes.
- Serial and parallel scanner totals match exactly on deterministic fixtures.
- Reparse points remain skipped.
- Access-denied directories do not abort the scan.
- Context menu path resolution works for directory nodes, file records, and synthetic buckets.
- DPI/theme changes invalidate cached text/layout/render resources.
- Cancel/close stops producers before releasing viewer/plugin resources.

---

## 11. Build and Verification Commands

Use the project's normal build loop:

```powershell
.\build.ps1 -ProjectName RedSalamander
```

Run focused command selftests from the Debug app, using the exact filters added by each phase:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=viewer_space
```

Run PerformanceTests2 when scanner/model/layout tests are added:

```powershell
vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll
```

For opt-in large local evidence:

```powershell
$env:REDSALAMANDER_VIEWERSPACE_LARGE_PERF = '1'
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=viewer_space_perf_large
```

After every run, confirm the archive path under `Specs/TestRuns/<MachineHash>/...` and compare baseline/candidate metrics. Terminal output alone is not valid closeout evidence.

---

## 12. Execution Order

1. Phase 0: measure current ViewerSpace exactly.
2. Phase 1: migrate renderer substrate without visual changes.
3. Phase 2: static cache + overlay split.
4. Phase 3: single-pass LOD renderer.
5. Phase 4: pixel-driven layout + hit grid.
6. Phase 5: file candidate store to unlock many more file squares.
7. Phase 6: dynamic work-stealing scanner.
8. Phase 7: batched allocation-light update pipe.
9. Phase 8: compact model and cache memory diet.
10. Phase 9: final remeasure, specs, coverage, Done move.

Renderer phases first unlock the visible "many more squares smoothly" win. File candidate and layout phases make those squares represent real files. Scanner/update/model phases deliver the speed and memory wins for huge explorations.

---

## 13. Risks and Rollback

| Risk | Mitigation |
|---|---|
| Device-context migration destabilizes visuals/device loss. | Phase 1 is visual-equivalent and separately revertible; preserve existing HWND background behavior. |
| Static cache hides stale visual state. | Explicit invalidation matrix plus cache-generation debug snapshot and hover no-rebuild test. |
| More file candidates increase memory. | Compact file records, candidate budgets, exact "Other", opt-in hydration, memory metrics. |
| Dynamic scanner races aggregation. | Serial-vs-parallel golden test and explicit parent completion state. |
| Non-Win32 providers dislike parallel calls. | Keep default `scanThreads = 1` for non-Win32 unless explicit opt-in/evidence exists. |
| High tile count makes hover slow. | Spatial hit grid and p95 hover metrics. |
| Scan cache doubles memory. | Byte cap and skip-large-cache path with metric. |
| LOD makes tiles too plain. | Pixel thresholds preserve labels/details on medium+ tiles and high-contrast readability. |

---

## 14. Definition of Done

The plan is complete only when:

- Same-machine baseline and final candidate runs exist under `Specs/TestRuns`.
- The new `viewer.space.*` metrics are present in `perf_metrics.jsonl`.
- The deterministic scenarios prove the implemented scanner, render, visible tile count, memory-settle, cache-cap, and queue-bounding behavior. The original POD side-buffer queue rewrite, full static-record paint rewrite, and hot/cold directory split are documented as future follow-ups only if the archived metrics identify those structures as dominant costs.
- Correctness selftests cover dense files, "Other" totals, serial-vs-parallel totals, hit testing, cancellation, cache invalidation, and existing pointer/menu behavior.
- `Specs/Plugins/Plugins_ViewerSpace.md` owns the durable behavior.
- `Specs/Testing/Testing_TestCoverage.md` lists the new coverage.
- This plan has been moved from `Specs/Plans/WIP/` to `Specs/Plans/Done/`.
