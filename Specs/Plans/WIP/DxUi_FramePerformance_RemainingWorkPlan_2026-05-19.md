# DxUi Frame Performance Remaining Work Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Finish the remaining DxUi, FolderView, and RedSalamanderMonitor performance work by adding the missing targeted scenarios first, then only implementing optimizations with same-machine before/after evidence.

**Architecture:** Treat the completed frame-runtime work as a prerequisite: shared DxUi frame metrics, FolderView frame metrics, and Monitor ColorTextView metrics must already exist before this plan is executed. Each remaining optimization is gated by a scenario that exercises the exact behavior, records p50/p95/p99 evidence, and archives results under `Specs/TestRuns/`.

**Tech Stack:** C++23, Win32, WIL, Direct2D, DirectWrite, Direct3D 11, DXGI, `Debug::Perf`, `DxUiTests`, `RedSalamander.exe --commands-selftest`, `RedSalamanderMonitor.exe --chrome-selftest --perf`, `MonitorTest`, `build.ps1`, archived JSONL perf metrics.

---

## Progress Checklist

- [x] Plan created in `Specs/Plans/WIP` on `master`.
- [x] Task 0: Verify frame-performance foundation is present.
- [x] Task 1: Add FolderView overlay perf scenario and metrics.
- [x] Task 2: Optimize FolderView overlay invalidation only if the new scenario justifies it.
- [ ] Task 3: Add richer Monitor ETW burst perf scenario.
- [ ] Task 4: Tune Monitor ETW scheduling only if burst metrics justify it.
- [ ] Task 5: Add Monitor SCROLL_BACK perf scenario.
- [ ] Task 6: Design and gate DXGI present-policy experiments with partial-present fallbacks.
- [ ] Task 7: Add targeted DxUi composition-animation scenario before any compositor pilot.
- [ ] Task 8: Run ARM64 Release validation for the frame-performance surfaces.
- [ ] Task 9: Update authoritative specs and move this plan to `Specs/Plans/Done`.

## Execution Notes

- This plan was written on `master`. If `master` does not yet contain the completed frame-performance foundation from `codex/dxui-frame-performance`, do not start Task 1. First merge or port that foundation, then run Task 0.
- Tasks 2, 4, 6, and 7 are optimization gates. If the required evidence does not pass the gate, record a measured no-op decision and do not change production code.
- Do not promote Tasks 9-11 from the previous closeout as performance gains. They were measured no-op decisions: Monitor scheduling, DXGI present policy, and composition pilot were not justified by evidence.
- Release selftest evidence requires explicit test hooks. Use `RSBuildEnableTests=true` for Release selftest builds.

## Task 0: Verify Frame-Performance Foundation Is Present

**Files:**

- Read: `Common/DxUi/DxUi.FrameRuntime.h`
- Read: `Common/DxUi/DxUi.WindowHost.cpp`
- Read: `RedSalamander/FolderView.Rendering.cpp`
- Read: `RedSalamanderMonitor/ColorTextView.cpp`
- Read: `RedSalamanderMonitor/RedSalamanderMonitor.vcxproj`
- Modify only if missing: the implementation branch or merge plan that brings the foundation to `master`

- [x] **Step 1: Verify expected source hooks**

Run:

```powershell
rg -n "DxUi.FrameRuntime|dxui\.frame\.|folder\.frame\.|monitor\.frame\.|monitor\.etw\.batch_drain_us|RSBuildTestDefinitions" Common RedSalamander RedSalamanderMonitor Tests Specs
```

Expected:

- `Common/DxUi/DxUi.FrameRuntime.h` exists.
- `dxui.frame.*` metrics appear in `Common/DxUi/DxUi.WindowHost.cpp` and `Tests/DxUiTests`.
- `folder.frame.*` metrics appear in FolderView render/selftest paths.
- `monitor.frame.*` and `monitor.etw.batch_drain_us` appear in ColorTextView/Monitor selftest paths.
- `RedSalamanderMonitor.vcxproj` Release x64 and ARM64 `ClCompile` definitions include `$(RSBuildTestDefinitions)`.

- [x] **Step 2: Stop if the foundation is absent**

If any expected source hook is missing, stop this plan and first land the completed frame-performance foundation branch. Record the blocker in this file under `## Implementation Notes`.

- [x] **Step 3: Run foundation smoke checks**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=Animation
.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\.build\x64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: all commands exit `0`; Release Monitor archive records `monitorFrameMetricPresence.allPresent=true`.

- [x] **Step 4: Commit only if documentation changed**

If Task 0 only verified the foundation, do not commit. If Task 0 recorded a blocker, run:

```powershell
git add Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "docs(perf): record frame foundation prerequisite"
```

## Task 1: Add FolderView Overlay Perf Scenario And Metrics

**Files:**

- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify: `RedSalamander/FolderView.Rendering.cpp`
- Modify as needed: `RedSalamander/FolderView.ErrorOverlay.cpp`
- Modify as needed: `RedSalamander/FolderViewIncrementalSearch.h`
- Modify: `Specs/UI/UI_FolderView.md`

- [x] **Step 1: Add scenario-local metric presence checks first**

Add a new commands selftest case named:

```text
folderView_perf_overlay_invalidation_stress
```

The scenario must archive a case metrics summary with these required fields:

```json
{
  "folderOverlayMetricPresence": {
    "allPresent": false,
    "metrics": [
      {"metric": "folder.frame.overlay_animation_count", "present": false, "count": 0},
      {"metric": "folder.frame.overlay_dirty_rect_area_px", "present": false, "count": 0},
      {"metric": "render.incremental_search_effect_updates", "present": false, "count": 0}
    ]
  }
}
```

The metric scan must start at the scenario's initial `perf_metrics.jsonl` byte offset so earlier test rows cannot satisfy the case.

- [x] **Step 2: Run red evidence**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4
```

Expected: selftest exits `0`, but `folderOverlayMetricPresence.allPresent=false` and the raw JSONL contains no required overlay rows.

- [x] **Step 3: Instrument overlay metrics without optimizing**

Emit these metrics from the actual overlay and incremental-search paint/update paths:

```text
folder.frame.overlay_animation_count
folder.frame.overlay_dirty_rect_area_px
render.incremental_search_effect_updates
```

Rules:

- Count only overlay work that actually ran during the scenario.
- `folder.frame.overlay_dirty_rect_area_px` must use the dirty/invalidation rectangle area, not the full client area unless the dirty rectangle is truly full-client.
- Do not alter dirty-rect behavior in this task.

- [x] **Step 4: Run green evidence**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4
```

Expected: selftest exits `0`; `folderOverlayMetricPresence.allPresent=true`; archive path and metric counts are recorded in this plan.

- [x] **Step 5: Commit**

Run:

```powershell
git add RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp RedSalamander/FolderView.Rendering.cpp RedSalamander/FolderView.ErrorOverlay.cpp RedSalamander/FolderViewIncrementalSearch.h Specs/UI/UI_FolderView.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "test(folderview): add overlay perf scenario"
```

## Task 2: Optimize FolderView Overlay Invalidation Only If Metrics Justify It

**Files:**

- Modify if justified: `RedSalamander/FolderView.Rendering.cpp`
- Modify if justified: `RedSalamander/FolderView.ErrorOverlay.cpp`
- Modify if justified: `RedSalamander/FolderView.Interaction.cpp`
- Modify: `Specs/UI/UI_FolderView.md`

- [x] **Step 1: Compare overlay scenario metrics**

Compare the Task 1 archive against a fresh candidate baseline:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4
```

Extract p50/p95/p99 and max for:

```text
folder.frame.total_us
folder.frame.present_us
folder.frame.overlay_dirty_rect_area_px
folder.frame.overlay_animation_count
render.incremental_search_effect_updates
render.dirty_rect_area_px
```

- [x] **Step 2: Gate the optimization**

Proceed only if all are true:

- `folder.frame.overlay_animation_count` is greater than `0`.
- Overlay dirty rectangles are full-client for at least 25 percent of overlay frames.
- Overlay frames increase `folder.frame.total_us` or `render.frame_us` p95 by at least 10 percent compared with non-overlay scroll frames from the same machine.

If the gate fails, record a measured no-op decision and skip to Step 6.

- [x] **Step 3: Add minimal overlay dirty-rect invalidation**

Constrain invalidation to the union of:

- previous overlay bounds,
- current overlay bounds,
- caret/search badge bounds if incremental search is active.

Do not invalidate the full FolderView client area unless the union touches every visible row or the overlay covers the full viewport.

- [x] **Step 4: Run before/after validation**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
```

Expected: both selftests exit `0`; overlay p95 dirty area decreases without increasing scroll p95/p99 frame time.

- [x] **Step 5: Commit code change if accepted**

Run:

```powershell
git add RedSalamander/FolderView.h RedSalamander/FolderView.ErrorOverlay.cpp RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp Specs/UI/UI_FolderView.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "perf(folderview): reduce measured overlay invalidation"
```

- [x] **Step 6: Commit measured no-op if rejected**

Not applicable: the evidence gate passed, so the accepted-code Step 5 path was used instead.

Run only if the gate fails:

```powershell
git add Specs/UI/UI_FolderView.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "docs(folderview): record overlay invalidation gate decision"
```

## Task 3: Add Richer Monitor ETW Burst Perf Scenario

**Files:**

- Modify: `RedSalamanderMonitor/RedSalamanderMonitor.cpp`
- Modify as needed: `RedSalamanderMonitor/ColorTextView.cpp`
- Modify: `Specs/Core/Core_RedSalamanderMonitor.md`

- [ ] **Step 1: Add a high-sample ETW burst mode behind selftest only**

Extend `--chrome-selftest --perf` with a selftest-only burst mode that produces at least 60 append-to-visible samples:

```text
--chrome-selftest --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260
```

The scenario must preserve the existing default chrome selftest path. The new mode must emit or summarize:

```text
monitor.etw.batch_drain_us
monitor.etw.selftest_burst_drain_us
monitor.frame.append_to_visible_us
monitor.frame.total_us
monitor.frame.present_us
monitor.frame.tail_layout_us
monitor.etw.queue_depth
monitor.etw.batch_repost_count
```

- [ ] **Step 2: Add red metric-presence evidence**

Before adding any new production scheduling behavior, run:

```powershell
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260
```

Expected: command exits `0`; summary shows the new burst scenario exists. If new queue/repost metrics are not yet present, record that as red evidence.

- [ ] **Step 3: Add metric summary output**

The scenario `results.json` must include p50/p95/p99/max for:

```text
monitor.frame.append_to_visible_us
monitor.etw.batch_drain_us
monitor.frame.total_us
monitor.frame.present_us
```

Use at least 60 append-to-visible samples before evaluating p95/p99.

- [ ] **Step 4: Run green evidence**

Run:

```powershell
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260
```

Expected: command exits `0`; `append_to_visible_us` sample count is at least `60`; archive path and quantiles are recorded in this plan.

- [ ] **Step 5: Commit**

Run:

```powershell
git add RedSalamanderMonitor/RedSalamanderMonitor.cpp RedSalamanderMonitor/ColorTextView.cpp Specs/Core/Core_RedSalamanderMonitor.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "test(monitor): add etw burst latency scenario"
```

## Task 4: Tune Monitor ETW Scheduling Only If Burst Metrics Justify It

**Files:**

- Modify if justified: `RedSalamanderMonitor/ColorTextView.cpp`
- Modify if justified: `RedSalamanderMonitor/ColorTextView.h`
- Modify if justified: `RedSalamanderMonitor/RedSalamanderMonitor.cpp`
- Modify: `Specs/Core/Core_RedSalamanderMonitor.md`

- [ ] **Step 1: Gate ETW scheduling work**

Use the Task 3 archive. Proceed only if one of these is true:

- `monitor.etw.batch_drain_us` p95 is greater than `8,333us`.
- `monitor.etw.batch_drain_us` p99 is greater than `16,667us`.
- More than one ColorTextView present is recorded per burst before all burst events are visible.
- `monitor.frame.append_to_visible_us` p95 exceeds `50,000us` and the dominant cost is ETW drain or repeated paints, not startup/initial paint.

If none are true, record a measured no-op and skip to Step 5.

- [ ] **Step 2: Bound drain work only if the gate passes**

Keep event order stable. Bound each UI-thread drain slice by:

```text
max events per slice: 200
max drain time per slice: 4,000us
```

If events remain, repost the existing zero-payload `WM_APP_ETW_BATCH` message. Do not allocate raw message payloads for this path.

- [ ] **Step 3: Add scheduling metrics**

Emit:

```text
monitor.etw.queue_depth
monitor.etw.batch_drained_count
monitor.etw.batch_repost_count
monitor.etw.batch_time_budget_hit
```

- [ ] **Step 4: Validate before/after**

Run:

```powershell
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260
.\build.ps1 -ProjectName MonitorTest -Configuration Debug
.\.build\x64\Debug\MonitorTest.exe --document-model-selftest
```

Expected: all commands exit `0`; append order is preserved; `append_to_visible_us` p95/p99 do not regress; queue depth drains to zero.

- [ ] **Step 5: Commit accepted code or measured no-op**

For accepted code:

```powershell
git add RedSalamanderMonitor/ColorTextView.cpp RedSalamanderMonitor/ColorTextView.h RedSalamanderMonitor/RedSalamanderMonitor.cpp Specs/Core/Core_RedSalamanderMonitor.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "perf(monitor): bound measured etw drain work"
```

For a measured no-op:

```powershell
git add Specs/Core/Core_RedSalamanderMonitor.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "docs(monitor): record etw scheduling gate decision"
```

## Task 5: Add Monitor SCROLL_BACK Perf Scenario

**Files:**

- Modify: `RedSalamanderMonitor/RedSalamanderMonitor.cpp`
- Modify as needed: `RedSalamanderMonitor/ColorTextView.cpp`
- Modify: `Specs/Core/Core_RedSalamanderMonitor.md`

- [ ] **Step 1: Add real scrollback scenario path**

Add a selftest mode:

```text
--chrome-selftest --perf --monitor-scrollback-selftest
```

The scenario must:

1. Append enough ETW lines to make scrolling meaningful.
2. Toggle auto-scroll off through the actual command path used by the menu (`IDM_OPTION_AUTO_SCROLL` / `WM_COMMAND`), not by directly manufacturing a render-only mode switch for metric presence.
3. Scroll up or page up through the ColorTextView input path.
4. Force one visible render.
5. Restore auto-scroll through the same command path.

- [ ] **Step 2: Require scrollback metric only in this scenario**

The scenario `results.json` must require:

```text
monitor.frame.scrollback_slice_us
monitor.frame.mode
monitor.frame.total_us
monitor.frame.present_us
```

The default chrome selftest must continue not requiring `monitor.frame.scrollback_slice_us`.

- [ ] **Step 3: Run validation**

Run:

```powershell
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-scrollback-selftest
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: both commands exit `0`; scrollback scenario requires `monitor.frame.scrollback_slice_us`; default chrome scenario does not.

- [ ] **Step 4: Commit**

Run:

```powershell
git add RedSalamanderMonitor/RedSalamanderMonitor.cpp RedSalamanderMonitor/ColorTextView.cpp Specs/Core/Core_RedSalamanderMonitor.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "test(monitor): add scrollback frame metric scenario"
```

## Task 6: Design And Gate DXGI Present-Policy Experiments

**Files:**

- Modify if justified: `Common/DxUi/DxUi.WindowHost.cpp`
- Modify if justified: `RedSalamander/FolderView.Rendering.cpp`
- Modify if justified: `RedSalamanderMonitor/ColorTextView.cpp`
- Modify if justified: `Tests/DxUiTests/DxUiTests.WindowHost.cpp`
- Modify: `Specs/UI/UI_DxUiWinUIDesign.md`

- [ ] **Step 1: Add a partial-present fallback design before code**

Document a design in `Specs/UI/UI_DxUiWinUIDesign.md` that answers:

- Which surfaces may use flip-discard.
- Which surfaces must switch to full redraw when flip-discard is active.
- How ColorTextView scroll-rect reuse is disabled or replaced.
- How dirty rectangles are validated under both modes.
- How device-lost, resize, occlusion, minimize, and DPI change paths are tested.

- [ ] **Step 2: Add capability probe behind disabled experiment flag**

If the design is accepted, add a disabled-by-default experiment flag:

```text
RS_DXGI_PRESENT_EXPERIMENT=flip_discard_full_redraw
```

The flag must not affect default builds. When enabled, it must log capability and fallback mode metrics:

```text
dxgi.present.experiment_enabled
dxgi.present.flip_discard_supported
dxgi.present.full_redraw_fallback_count
```

- [ ] **Step 3: Gate with same-machine before/after evidence**

Run baseline and candidate:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected before acceptance:

- All commands exit `0`.
- `present_us` p95/p99 improves or variance decreases on at least one target surface.
- No dirty-rect correctness, scrollback, resize, or device-lost regression appears.

- [ ] **Step 4: Commit accepted experiment or measured rejection**

For accepted code:

```powershell
git add Common/DxUi/DxUi.WindowHost.cpp RedSalamander/FolderView.Rendering.cpp RedSalamanderMonitor/ColorTextView.cpp Tests/DxUiTests/DxUiTests.WindowHost.cpp Specs/UI/UI_DxUiWinUIDesign.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "perf(render): add gated present policy experiment"
```

For rejection:

```powershell
git add Specs/UI/UI_DxUiWinUIDesign.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "docs(render): record present experiment gate decision"
```

## Task 7: Add Targeted DxUi Composition-Animation Scenario Before Any Pilot

**Files:**

- Modify: `Tests/DxUiTests/DxUiTests.Animation.cpp`
- Modify as needed: `Tests/DxUiTests/DxUiTests.WindowHost.cpp`
- Modify if justified: `Common/DxUi/DxUi.WindowHost.cpp`
- Modify if justified: `Common/DxUi/DxUi.CompositionPilot.h`
- Modify if justified: `Common/DxUi/DxUi.CompositionPilot.cpp`
- Modify if justified: `Common/DxUi/DxUi.vcxproj`
- Modify: `Specs/UI/UI_DxUiWinUIDesign.md`

- [ ] **Step 1: Add targeted scenario first**

Add an Animation or WindowHost scenario that exercises one allowed surface:

```text
tooltip fade
popup opacity/translate
dialog smoke opacity
lightweight overlay transform
```

Record:

```text
dxui.animation.tick_delta_us
dxui.animation.jitter_us
dxui.frame.total_us
dxui.frame.render_us
dxui.frame.present_us
dxui.animation.allowed_surface
```

- [ ] **Step 2: Gate compositor pilot**

Proceed only if:

- The targeted scenario shows unresolved CPU-thread jitter p95 greater than `8,333us`.
- The jitter is on an allowed animation surface.
- The issue persists after the shared frame runtime scheduler is enabled.
- The pilot can be disabled by default.

If the gate fails, record a measured no-op.

- [ ] **Step 3: Add disabled-by-default pilot only if accepted**

If accepted, create:

```text
Common/DxUi/DxUi.CompositionPilot.h
Common/DxUi/DxUi.CompositionPilot.cpp
```

Requirements:

- Use WIL RAII wrappers for COM resources.
- Do not introduce global mutable device state.
- Do not touch FolderView item rendering, ColorTextView log rendering, DirectWrite text layout, hit-testing, or file list virtualization.
- Keep the pilot disabled unless an explicit feature flag enables it.

- [ ] **Step 4: Validate**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=Animation
.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost
```

Expected: tests pass with the pilot disabled. If the pilot is enabled manually, archived metrics or PresentMon evidence must show smoother animation before any default enablement.

- [ ] **Step 5: Commit accepted pilot or measured rejection**

For accepted pilot:

```powershell
git add Common/DxUi/DxUi.CompositionPilot.h Common/DxUi/DxUi.CompositionPilot.cpp Common/DxUi/DxUi.vcxproj Common/DxUi/DxUi.WindowHost.cpp Tests/DxUiTests/DxUiTests.Animation.cpp Tests/DxUiTests/DxUiTests.WindowHost.cpp Specs/UI/UI_DxUiWinUIDesign.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "perf(dxui): pilot compositor-backed allowed animation"
```

For rejection:

```powershell
git add Tests/DxUiTests/DxUiTests.Animation.cpp Specs/UI/UI_DxUiWinUIDesign.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "docs(dxui): record composition animation gate decision"
```

## Task 8: Run ARM64 Release Validation

**Files:**

- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Modify: `Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md`

- [ ] **Step 1: Build ARM64 Release targets**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Release -Platform ARM64
.\build.ps1 -ProjectName RedSalamander -Configuration Release -Platform ARM64
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release -Platform ARM64 } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
```

Expected: all builds exit `0`. If ARM64 executables cannot be run on the current machine, record build-only validation and the reason.

- [ ] **Step 2: Run ARM64 selftests only when supported locally**

If the host can execute ARM64 binaries, run:

```powershell
.\.build\ARM64\Release\DxUiTests.exe
.\.build\ARM64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\.build\ARM64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: commands exit `0`; archives are recorded.

- [ ] **Step 3: Commit validation evidence**

Run:

```powershell
git add Specs/Testing/Testing_TestCoverage.md Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "docs(testing): record arm64 frame validation"
```

## Task 9: Specs Closeout For Remaining Work

**Files:**

- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Modify: `Specs/UI/UI_DxUiWinUIDesign.md`
- Modify: `Specs/Core/Core_RedSalamanderMonitor.md`
- Modify: `Specs/UI/UI_FolderView.md`
- Move: `Specs/Plans/WIP/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md`

- [ ] **Step 1: Run final focused validation**

Run:

```powershell
.\build.ps1 -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-scrollback-selftest
```

Expected: every command exits `0`; all archive paths are recorded.

- [ ] **Step 2: Update authoritative specs**

Update:

- `Specs/UI/UI_FolderView.md` with overlay metrics, overlay scenario, and optimization/no-op outcome.
- `Specs/Core/Core_RedSalamanderMonitor.md` with ETW burst scenario, scheduling/no-op outcome, and scrollback scenario.
- `Specs/UI/UI_DxUiWinUIDesign.md` with accepted or rejected DXGI/composition experiment results.
- `Specs/Testing/Testing_TestCoverage.md` with final command list and archive paths.

- [ ] **Step 3: Move this plan to Done**

Run:

```powershell
Move-Item -LiteralPath 'Specs\Plans\WIP\DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md' -Destination 'Specs\Plans\Done\DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md'
```

- [ ] **Step 4: Commit closeout**

Run:

```powershell
git add Specs/Testing/Testing_TestCoverage.md Specs/UI/UI_DxUiWinUIDesign.md Specs/Core/Core_RedSalamanderMonitor.md Specs/UI/UI_FolderView.md Specs/Plans/Done/DxUi_FramePerformance_RemainingWorkPlan_2026-05-19.md
git commit -m "docs: close remaining frame performance plan"
```

## Acceptance Matrix

| Area | Required evidence |
| --- | --- |
| Foundation | `dxui.frame.*`, `folder.frame.*`, and `monitor.frame.*` metrics exist on the execution branch |
| FolderView overlay | overlay-specific scenario archives required overlay metrics |
| FolderView overlay optimization | dirty area and p95/p99 frame metrics improve or measured no-op is recorded |
| Monitor ETW burst | at least 60 append-to-visible samples with p50/p95/p99/max |
| Monitor scheduling | queue drains to zero; append order preserved; p95/p99 do not regress |
| Monitor scrollback | real command/input path emits `monitor.frame.scrollback_slice_us` in scrollback scenario only |
| DXGI present policy | partial-present fallback design exists before any flip-discard or waitable-object experiment |
| Composition pilot | targeted allowed-surface scenario exists before any compositor-backed code |
| ARM64 Release | build-only or run evidence recorded with exact reason if runtime execution is unavailable |

## Implementation Notes

Add dated notes here during execution. Each note must include:

- command or archive path,
- pass/fail result,
- metric gate decision,
- whether code changed or the task was a measured no-op.

### 2026-05-19 - Task 0 Foundation Verification

- Foundation branch integration: `git merge --no-ff codex/dxui-frame-performance -m "merge: bring frame performance foundation to master"` exited `0`; code changed by bringing the completed foundation onto `master`.
- Source hook scan: `rg -n "DxUi.FrameRuntime|dxui\.frame\.|folder\.frame\.|monitor\.frame\.|monitor\.etw\.batch_drain_us|RSBuildTestDefinitions" Common RedSalamander RedSalamanderMonitor Tests Specs` exited `0`; expected `DxUi.FrameRuntime`, `dxui.frame.*`, `folder.frame.*`, `monitor.frame.*`, `monitor.etw.batch_drain_us`, and Release `$(RSBuildTestDefinitions)` hooks were present.
- Smoke build: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_184908_473.log`; diagnostics `0 warning(s), 0 error(s)`.
- Smoke tests: `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` and `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost` both exited `0`.
- Release Monitor test-enabled build: `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; log `.build/logs/msbuild-20260519_185047_485.log`; diagnostics `0 warning(s), 0 error(s)`.
- Release Monitor perf smoke: `.\.build\x64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_185140`; `results.json` reports `status=passed` and `monitorFrameMetricPresence.allPresent=true`.
- Metric gate decision: foundation prerequisite passed. No Task 0 optimization was attempted.

### 2026-05-19 - Task 1 FolderView Overlay Perf Scenario

- Red evidence: after adding the scenario-local `folderOverlayMetricPresence` artifact before overlay metric instrumentation, `.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_191156`; `folderOverlayMetricPresence.allPresent=false` with counts `folder.frame.overlay_animation_count=0`, `folder.frame.overlay_dirty_rect_area_px=0`, and `render.incremental_search_effect_updates=6171`. Follow-up scenario cleanup removed warm/icon-message pumping from the overlay-specific case without changing the required artifact contract.
- Implementation: `folderView_perf_overlay_invalidation_stress` now creates a deterministic 180-item local folder, waits for a scenario-local enumeration-completed callback before input, drives direct FolderView incremental-search characters for `overlay`, shows the busy cancel overlay, paints overlay frames with `InvalidateRect` + `UpdateWindow`, scans `perf_metrics.jsonl` from the case's initial byte offset, and writes `perf/folderView_perf_overlay_invalidation_stress_metrics.json`.
- Metrics added without dirty-rect optimization: `FolderView::DrawErrorOverlay` emits `folder.frame.overlay_animation_count` when a busy/show animation frame is drawn; `FolderView::Render` emits `folder.frame.overlay_dirty_rect_area_px` only when the incremental-search indicator or error overlay actually draws, using the current paint dirty-rect area; existing `render.incremental_search_effect_updates` remains the incremental-search update counter.
- Green build: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_195820_189.log`; diagnostics `0 warning(s), 0 error(s)`.
- Green selftest: waited GUI run `Start-Process .build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4` exited `0` in `2562ms`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_200009`; `commands_results.json` reports 1 passed / 0 failed / 0 skipped.
- Green metric counts: `folderOverlayMetricPresence.allPresent=true`; counts are `folder.frame.overlay_animation_count=8`, `folder.frame.overlay_dirty_rect_area_px=8`, and `render.incremental_search_effect_updates=12`.
- Stability hardening: the first fresh Task 2 overlay gate exposed a scenario flake where the queued command/message path left quick search inactive (`Specs/TestRuns/4cb089111a23/Commands/2026-05-19_200415`, follow-up failures `2026-05-19_200851` and `2026-05-19_201100`). The scenario now activates quick search through `FolderWindow::CommandQuickSearch(...)` and avoids pumping unrelated messages between activation and typed input.
- Stabilization validation: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_201127_450.log`; diagnostics `0 warning(s), 0 error(s)`. Three consecutive waited GUI runs of `folderView_perf_overlay_invalidation_stress` exited `0` with archives `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_201303`, `2026-05-19_201306`, and `2026-05-19_201309`; the latest metric artifact still reports `folderOverlayMetricPresence.allPresent=true` with counts `8`, `8`, and `12`.
- Metric gate decision: Task 1 added scenario and metrics only. No overlay invalidation optimization was attempted; Task 2 must decide separately from same-machine before/after evidence.

### 2026-05-19 - Task 2 FolderView Overlay Invalidation Gate And Optimization

- Gate baseline commands: created isolated baseline worktree `C:\Users\eric\.config\superpowers\worktrees\RedSalamander\dxui-overlay-baseline-limitedpump` at `8c717bffe` and applied only the Task 2 overlay harness change (settled paint flush, limited message pump, and natural 1600ms capture window). `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `C:\Users\eric\.config\superpowers\worktrees\RedSalamander\dxui-overlay-baseline-limitedpump\.build\logs\msbuild-20260519_205250_160.log`; diagnostics `0 warning(s), 0 error(s)`.
- Gate baseline runs: baseline `folderView_perf_overlay_invalidation_stress` exited `0` in `11,529ms`; archive `C:\Users\eric\.config\superpowers\worktrees\RedSalamander\dxui-overlay-baseline-limitedpump\Specs\TestRuns\4cb089111a23\Commands\2026-05-19_205512`; `folderOverlayMetricPresence.allPresent=true` with counts `88`, `88`, and `88`. Baseline `folderView_perf_scroll_render_stress` exited `0` in `8,638ms`; archive `C:\Users\eric\.config\superpowers\worktrees\RedSalamander\dxui-overlay-baseline-limitedpump\Specs\TestRuns\4cb089111a23\Commands\2026-05-19_205526`.
- Gate baseline metrics: overlay animation count was present (`124` raw rows). Overlay dirty rects were full-client on `124/124` raw overlay rows (`100%`, `1,161,911px`). Overlay p95 was slower than non-overlay scroll p95 on the same baseline harness: `folder.frame.total_us` `71,825us` vs `45,217us` (`+58.85%`) and `render.frame_us` `73,533us` vs `46,828us` (`+57.03%`).
- Gate decision: accepted. All Task 2 prerequisites passed, so production overlay invalidation was changed rather than recording a measured no-op.
- Implementation: settled overlay animation ticks now invalidate a bounded rectangle made from current overlay panel bounds, previous overlay invalidation bounds, and the active incremental-search badge bounds. Full-client invalidation remains for initial show/scrim settling, forced full-client transitions, or missing overlay layout. The overlay selftest now waits for initial show settling, flushes one settled paint before taking the metric scan offset, uses a limited message pump, and records a natural 1600ms animation window without draining the animation queue in one loop.
- Validation build: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_205250_115.log`; diagnostics `0 warning(s), 0 error(s)`.
- After validation commands: waited GUI run of `folderView_perf_overlay_invalidation_stress` exited `0` in `10,980ms`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_205543`; `commands_results.json` reports 1 passed / 0 failed / 0 skipped and `folderOverlayMetricPresence.allPresent=true` with counts `108`, `108`, and `108`. Waited GUI rerun of `folderView_perf_scroll_render_stress` exited `0` in `7,873ms`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_205649`; `commands_results.json` reports 1 passed / 0 failed / 0 skipped.
- Overlay quantiles (`count`, `p50`, `p95`, `p99`, `max`):
  - Baseline `folder.frame.total_us`: `128`, `56,777`, `71,825`, `77,556`, `78,538`; after: `152`, `41,865`, `54,236`, `62,337`, `66,246`.
  - Baseline `folder.frame.present_us`: `128`, `104`, `144`, `665`, `758`; after: `152`, `118`, `157`, `564`, `699`.
  - Baseline `folder.frame.overlay_dirty_rect_area_px`: `124`, `1,161,911`, `1,161,911`, `1,161,911`, `1,161,911`; after: `148`, `721,522`, `721,522`, `1,161,911`, `1,161,911`.
  - Baseline `folder.frame.overlay_animation_count`: `124`, `1`, `1`, `1`, `1`; after: `148`, `1`, `1`, `1`, `1`.
  - Baseline `render.incremental_search_effect_updates`: `128`, `147`, `147`, `147`, `147`; after: `152`, `105`, `105`, `147`, `147`.
  - Baseline `render.dirty_rect_area_px`: `128`, `1,161,911`, `1,161,911`, `1,161,911`, `1,188,249`; after: `152`, `721,522`, `1,161,911`, `1,161,911`, `1,188,249`.
  - Baseline `render.frame_us`: `128`, `58,433`, `73,533`, `79,318`, `80,610`; after: `152`, `43,898`, `55,859`, `65,996`, `67,902`.
- After metrics: `folder.frame.overlay_dirty_rect_area_px` p95 improved from `1,161,911px` to `721,522px` (`-37.90%`) and full-client overlay rows dropped from `124/124` to `4/148` (`2.70%`). Overlay `folder.frame.total_us` p95 improved from `71,825us` to `54,236us` (`-24.49%`) and `render.frame_us` p95 improved from `73,533us` to `55,859us` (`-24.04%`).
- Scroll guard metrics: non-overlay scroll `folder.frame.total_us` p95 moved from `45,217us` to `41,145us` (`-9.01%`) and p99 moved from `47,086us` to `46,126us` (`-2.04%`). Compatibility `render.frame_us` p95 moved from `46,828us` to `43,114us` (`-7.93%`); `render.frame_us` p99 moved from `49,418us` to `49,885us` (`+0.94%`), but the primary `folder.frame.total_us` p95/p99 guard improved and no user-visible scroll-frame regression was measured.

### 2026-05-19 - Stop/Handoff Snapshot

- Current `master` state for the next chat: `HEAD=042c07073` (`perf(folderview): reduce measured overlay invalidation`), with Task 2 code, WIP-plan updates, and `Specs/UI/UI_FolderView.md` updates committed.
- Review state: Task 2 spec-compliance re-review returned `APPROVED` from subagent `019e419a-c111-7f21-816b-a50ebfabcee6`; code-quality re-review returned `APPROVED` from subagent `019e419a-c15d-76d2-ba5e-8f8f6e195679`.
- Final committed-tree verification before the stop request: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_205250_115.log`; diagnostics `0 warning(s), 0 error(s)`. `folderView_perf_overlay_invalidation_stress` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_205543`. `folderView_perf_scroll_render_stress` rerun exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_205649`.
- Continuation caveat: the same-harness baseline numbers above came from disposable worktree `C:\Users\eric\.config\superpowers\worktrees\RedSalamander\dxui-overlay-baseline-limitedpump`. After committing Task 2, that worktree was removed once, then recreated at `8c717bffe` with only the overlay selftest harness diff applied. Its replacement build exited `0`; log `C:\Users\eric\.config\superpowers\worktrees\RedSalamander\dxui-overlay-baseline-limitedpump\.build\logs\msbuild-20260519_210918_544.log`; diagnostics `0 warning(s), 0 error(s)`. Because the user asked to stop, baseline overlay/scroll reruns were not repeated in the recreated worktree and its old archive paths may need regeneration before this plan is closed or moved to `Specs/Plans/Done`.
- Recommended next chat start: first run `git status --short --branch` in `Z:\src\RedSalamander`. If Task 2 archival durability is required, rerun the two baseline commands from `C:\Users\eric\.config\superpowers\worktrees\RedSalamander\dxui-overlay-baseline-limitedpump`, copy or record the new baseline archives in this plan, then remove the disposable worktree. Otherwise continue with Task 3 (`Add Richer Monitor ETW Burst Perf Scenario`) from this WIP checklist.
