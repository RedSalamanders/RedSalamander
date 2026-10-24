# DxUi FolderView Monitor Frame Performance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Improve smoothness, frame-time stability, and maintainability for DxUi, FolderView, and RedSalamanderMonitor without replacing the specialized renderers that already make FolderView and ColorTextView fast.

**Architecture:** Add a small shared frame runtime inside `Common/DxUi` because both `RedSalamander` and `RedSalamanderMonitor` already reference `DxUi.vcxproj`. The runtime owns clocking, frame-stage telemetry, animation scheduling policy, and debug stage guards; each surface keeps ownership of its D2D/D3D/DXGI resources, render code, dirty-region policy, and selftests.

**Tech Stack:** C++23, Win32, WIL, Direct2D, DirectWrite, D3D11, DXGI flip-model swap chains, `Debug::Perf`, `DxUiTests`, `RedSalamander.exe --commands-selftest`, `RedSalamanderMonitor.exe --chrome-selftest`, `MonitorTest`, archived `Specs/TestRuns/` evidence, optional PresentMon/GPUView manual profiling.

---

## Progress Checklist

- [x] Plan created in `Specs/Plans/WIP`.
- [x] User selected Subagent-Driven execution.
- [x] Create or switch to implementation branch/workspace.
- [x] Task 1: Capture protected baselines.
- [x] Task 2: Add shared frame runtime skeleton.
- [x] Task 3: Add DxUi frame-stage telemetry.
- [x] Task 4: Replace fixed animation timer policy.
- [x] Task 5: Add debug guard against layout during render.
- [x] Task 6: Adopt frame runtime in FolderView without behavioral optimization.
- [x] Task 7: Optimize FolderView only where metrics justify it.
- [x] Task 8: Adopt frame runtime in ColorTextView without changing renderer modes.
- [x] Task 9: Tune Monitor batch and paint scheduling only where metrics justify it.
- [x] Task 10: Evaluate DXGI present policy improvements.
- [x] Task 11: Evaluate optional composition animation pilot.
- [x] Task 12: Full validation and spec closeout.

## Implementation Notes And New Topics

Add dated notes here while coding and testing. New topics are allowed when evidence discovers them, but each new topic must state:

- why it was added,
- what metric, test, or code path exposed it,
- whether it changes scope or only improves validation,
- which task owns it.

### 2026-05-19

- Subagent-driven execution requested. The first execution change is adding this progress checklist so task state can be checked off directly in the WIP plan as work advances.
- Implementation branch created: `codex/dxui-frame-performance`.
- Worker 1 captured protected baselines on same-machine profile `4cb089111a23`. Initial `git status --short` was clean. Builds, DxUi baselines, FolderView baselines, and Monitor model/gate baselines were executed; `RedSalamanderMonitor.exe --chrome-selftest --perf` produced an archive but its selftest result is [blocked] by missing `ENABLE_TESTS` in the built monitor binary.
- Task 1 spec review found the placeholder-scan result was not recorded. The evidence section now records a reproducible scan and result; no implementation scope changed.
- Task 1 passed spec and quality review. New topic for later validation: identify the correct test-enabled `RedSalamanderMonitor` build path before Task 8 or Task 12 so Monitor chrome perf has a passing baseline.
- Worker 2 Task 2 red evidence: after adding `TestFrameRuntimeClockIsMonotonic`, `TestFrameRuntimeClampsLargeDelta`, and `TestFrameRuntimeReducedMotionPolicy`, `./build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `1` with `DxUiTests.Animation.cpp(2,10): error C1083: Cannot open include file: 'DxUi/DxUi.FrameRuntime.h': No such file or directory`. The follow-up `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` exited `0` against the previous binary, so the compile failure is the authoritative red result.
- Worker 2 Task 2 green evidence: `./build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0` with 0 warnings and 0 errors, log `.build/logs/msbuild-20260519_133632_158.log`; `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` exited `0` with `All DxUi tests passed`.
- Worker 2 Task 2 adjustment: `Common/DxUi/DxUi.vcxproj.filters` is absent in this repo, so Task 2 only updated `Common/DxUi/DxUi.vcxproj` and did not create a filters file.
- Worker 2 Fix corrected Task 2 frame elapsed conversion to avoid multiply-first overflow for large QPC deltas. Red evidence: after adding `TestFrameRuntimeElapsedUsHandlesLargeQpcDelta`, `./build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`, then `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` exited `1` with `FAILED: frame runtime elapsed microseconds handles large QPC deltas without overflow`. Green evidence before final verification: `./build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0` with 0 warnings and 0 errors, log `.build/logs/msbuild-20260519_134429_066.log`; `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` exited `0` with `All DxUi tests passed`.
- Worker 2 Fix 3 corrected Task 2 `FrameClock::ElapsedUs` to compute ordered QPC deltas in `uint64_t`, avoiding signed-subtraction UB for wide cross-sign timestamp pairs. Red evidence: after adding `TestFrameRuntimeElapsedUsHandlesWideCrossSignDelta`, `./build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`, then `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` exited `0`; the test is retained as UB regression coverage because MSVC happened to wrap the overflowing signed subtraction. Green evidence: `./build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0` with 0 warnings and 0 errors, log `.build/logs/msbuild-20260519_135910_869.log`; `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` exited `0` with `All DxUi tests passed`; the non-self-matching placeholder scan exited `1` with no matches.
- Task 3 test seam: no production debug helper was added. `TestWindowHostEmitsFrameStageMetricsForCaptureRender` uses the existing `--perf-jsonl` sink path from `REDSALAMANDER_PERF_JSONL_PATH`, captures the file offset immediately before direct `DebugCaptureBitmap`, and inspects only the appended JSONL rows. When no runner sink is configured, the test creates a local scratch JSONL sink.
- Task 3 red evidence: after adding the focused WindowHost telemetry test, `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_140858_788.log`; diagnostics `1 warning(s), 0 error(s)` from the new array initializer, fixed during green implementation. `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost --perf-jsonl=Specs\TestRuns\local_scratch\dxui_windowhost_stage_metrics_red_20260519.jsonl` exited `1` at `FAILED: window host capture render emits every frame-stage metric`; `rg '"metric":"dxui\.frame\.' Specs\TestRuns\local_scratch\dxui_windowhost_stage_metrics_red_20260519.jsonl` exited `1` with no matches.
- Task 3 implementation: `WindowHost::Render` now uses `FrameClock`, `FrameStage`, and `FrameStageScope` around update/resource preparation, render/capture work, and present in both normal and `ENABLE_TESTS` paths. It keeps `Debug::Perf::Scope paintPerf(L"DxUI::Paint")` and emits `dxui.frame.total_us`, `dxui.frame.update_us`, `dxui.frame.render_us`, `dxui.frame.present_us`, `dxui.frame.dirty_rect_count`, and `dxui.frame.dirty_rect_area_px`. Full-frame or invalid dirty rect renders report dirty count/area as `0`; partial dirty rect renders report count `1` and width-by-height area.
- Task 3 green evidence: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_141052_666.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost --perf-jsonl=Specs\TestRuns\local_scratch\dxui_windowhost_stage_metrics_green_20260519.jsonl` exited `0` with `All DxUi tests passed`; `rg '"metric":"dxui\.frame\.' Specs\TestRuns\local_scratch\dxui_windowhost_stage_metrics_green_20260519.jsonl` found 138 rows, including all six required `dxui.frame.*` metric names.
- Task 3 hygiene evidence: the non-self-matching placeholder scan over the plan and touched source files exited `1` with no matches. `git diff --check` exited `0`; it printed only line-ending normalization warnings for the three touched files.
- Task 3 quality-review fix: review found `TestWindowHostEmitsFrameStageMetricsForCaptureRender` captured the JSONL offset before host/root setup, so setup paint rows could satisfy the assertion. The test now captures the offset immediately before `CaptureAttachedHostWindowBitmapForWindowHostSuite(...)`, searches quoted JSON `metric` fields, and verifies full-frame dirty count/area rows report `"value":0`.
- Task 3 quality-fix evidence: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_141959_761.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost --perf-jsonl=Specs\TestRuns\local_scratch\dxui_windowhost_stage_metrics_qualityfix_20260519.jsonl` exited `0` with `All DxUi tests passed`; quoted-metric and dirty-value `rg` checks found the required rows.
- Task 3 second quality-review fix: review found `CaptureAttachedHostWindowBitmapForWindowHostSuite(...)` itself can show, pump, redraw, and emit paint metrics after the offset. The test now uses that helper only as a warm-up before taking the offset, then calls `window.Host().DebugCaptureBitmap(capture)` directly so appended rows are isolated to `DebugCaptureBitmap -> Render(nullptr, &capture)`.
- Task 3 capture-isolation evidence: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_142333_729.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost --perf-jsonl=Specs\TestRuns\local_scratch\dxui_windowhost_stage_metrics_capturefix_20260519.jsonl` exited `0` with `All DxUi tests passed`.
- Task 4 red evidence: after adding the dispatcher scheduler tests, `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_143035_060.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_scheduler_red_20260519.jsonl` exited `1` with `FAILED: animation dispatcher emits high-resolution tick delta metrics for active subscribers`. The red JSONL contained only the legacy `dxui.animation.tick_gap_ms` and `dxui.animation.tick_overrun` rows, confirming the current dispatcher did not emit high-resolution tick delta, jitter, or active-count metrics.
- Task 4 implementation: `AnimationDispatcher` keeps the existing `bool (*)(void* context, uint64_t nowTickMs) noexcept` subscription API and message-only-window `SetTimer` fallback, but now derives callback time and metrics from `DxUi::FrameClock`, uses an 8 ms fallback cadence with an 8,333 us synthetic 120 Hz target, clamps virtual callback-time hitches through `FrameBudget`, emits `dxui.animation.tick_delta_us`, `dxui.animation.jitter_us`, and `dxui.animation.active_count`, and preserves `dxui.animation.tick_gap_ms` / `dxui.animation.tick_overrun`. Final quality review kept `WindowHost::OnAnimationTick` resumable while hidden/minimized: it skips nonessential root/tooltip ticking and invalidation but keeps the active animation subscription alive for restore.
- Task 4 seam note: initial metric tests drove the real public subscription API and pumped the dispatcher message window, then later spec/quality fixes added narrow `ENABLE_TESTS` hooks only for deterministic scheduler-policy and hidden-host tick assertions.
- Task 4 green evidence: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_143243_560.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_scheduler_green_20260519.jsonl` exited `0` with `All DxUi tests passed`; `rg '"metric":"dxui\.animation\.(tick_delta_us|jitter_us|active_count|tick_gap_ms|tick_overrun)"' Specs\TestRuns\local_scratch\dxui_animation_scheduler_green_20260519.jsonl` found the new high-resolution metrics and retained legacy tick-gap rows. `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_143312_765.log`; diagnostics `0 warning(s), 0 error(s)`.
- Task 4 hygiene evidence: the non-self-matching placeholder scan over the touched files exited `1` with no matches. `git diff --check` exited `0`; it printed only line-ending normalization warnings for the four touched files.
- Task 4 spec-fix red evidence: review found `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` failed without `--perf-jsonl`; reproduced exit `1` with `FAILED: animation dispatcher emits high-resolution tick delta metrics for active subscribers`. After changing the test to call dispatcher-owned scheduler-policy helpers before adding the hook, `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `1`; log `.build/logs/msbuild-20260519_144306_044.log`; diagnostics `0 warning(s), 3 error(s)` for missing dispatcher scheduler-policy test helpers.
- Task 4 spec-fix implementation: `Tests/DxUiTests/DxUiTests.Animation.cpp` now creates `Specs/TestRuns/local_scratch/dxui_animation_scheduler_testlocal_20260519.jsonl` when the runner did not configure `REDSALAMANDER_PERF_JSONL_PATH`, and clears global perf config only when it created that local sink. The 120 Hz and hitch-clamp assertions now query the actual `AnimationDispatcher` scheduler policy through a narrow `ENABLE_TESTS` hook: target frame budget, hitch clamp budget, and clamp operation only.
- Task 4 spec-fix green evidence: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_144333_918.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` exited `0` with `All DxUi tests passed`, proving the fallback sink path. `.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_scheduler_specfix_20260519.jsonl` exited `0`; `rg '"metric":"dxui\.animation\.(tick_delta_us|jitter_us|active_count|tick_gap_ms|tick_overrun)"' Specs\TestRuns\local_scratch\dxui_animation_scheduler_specfix_20260519.jsonl` found the expected dispatcher metrics. `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_144402_765.log`; diagnostics `0 warning(s), 0 error(s)`.
- Task 4 spec-fix hygiene evidence: the non-self-matching placeholder scan over the touched files exited `1` with no matches. `git diff --check` exited `0`; it printed only line-ending normalization warnings for the three touched files.
- Task 4 quality-fix red evidence: after adding focused coverage, `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `1`; log `.build/logs/msbuild-20260519_145619_838.log`; diagnostics `0 warning(s), 4 error(s)` because `AnimationDispatcher::DebugComputeTimingForTest` and `WindowHost::DebugAnimationTickForTest` did not exist. This proved the new tests were aimed at the missing resumability and raw legacy-gap seams before implementation.
- Task 4 quality-fix implementation: hidden/minimized `WindowHost::OnAnimationTick` now returns `true` without clearing `_animationSubscriptionId`, ticking root/tooltip, or invalidating, so active animations remain resumable when the host becomes visible again. `AnimationDispatcher` now computes callback `nowTickMs` from the smoothed/clamped delta while emitting legacy `dxui.animation.tick_gap_ms` and `dxui.animation.tick_overrun` from the raw elapsed delta. Tests cover both paths with narrow `ENABLE_TESTS` helpers.
- Task 4 quality-fix green evidence: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_145914_629.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=Animation` exited `0`; `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost` exited `0`; `.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_scheduler_qualityfix_20260519.jsonl` exited `0`, and the JSONL contains `dxui.animation.tick_delta_us`, `dxui.animation.jitter_us`, `dxui.animation.active_count`, `dxui.animation.tick_gap_ms`, and `dxui.animation.tick_overrun`. `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_145938_746.log`; diagnostics `0 warning(s), 0 error(s)`.
- Task 4 quality-fix hygiene evidence: the non-self-matching placeholder scan over the touched files exited `1` with no matches. `git diff --check` exited `0`; it printed only line-ending normalization warnings for the six touched files.
- Task 6 started in the existing `codex/dxui-frame-performance` workspace with a clean tracked tree. Scope is limited to FolderView source/header, the existing scroll/render selftest artifact, and this WIP plan. The first selftest change records presence booleans for `folder.frame.total_us`, `folder.frame.present_us`, `folder.frame.visible_work_count`, and `folder.frame.input_to_paint_us` without adding thresholds or failures.
- Task 6 red/baseline evidence: after adding only the selftest artifact presence recording, `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_152658_900.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_152856`; `commands_results.json` reports 1 passed / 0 failed / 0 skipped; `perf/folderView_perf_scroll_render_stress_metrics.json` records `folderFrameMetricPresence.allPresent=false` and `present=false` for all four new metrics. `rg '"metric":"folder\.frame\.(total_us|present_us|visible_work_count|input_to_paint_us)"' Specs\TestRuns\4cb089111a23\Commands\2026-05-19_152856\perf\perf_metrics.jsonl` exited `1`, confirming the metrics were absent before production instrumentation.
- Task 6 implementation: `FolderView::Render` now uses `DxUi::FrameClock`/`FrameStageScope` around render and present work while preserving existing `render.frame_us`, `render.present_us`, dirty-rect clipping, `Present1` parameters, visible item counters, and device-lost returns. It emits `folder.frame.total_us`, `folder.frame.present_us`, `folder.frame.visible_work_count`, and `folder.frame.dirty_rect_area_px`. `FolderView.cpp` captures optional user-input timestamps in `WM_MOUSEWHEEL`, `WM_MOUSEHWHEEL`, `WM_HSCROLL`, `WM_KEYDOWN`, and handled `WM_SYSKEYDOWN` only when viewport or focus state changes, then emits `folder.frame.input_to_paint_us` after the next successful present.
- Task 6 green evidence: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_153156_901.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_153359`; `commands_results.json` reports 1 passed / 0 failed / 0 skipped; `perf/folderView_perf_scroll_render_stress_metrics.json` records `folderFrameMetricPresence.allPresent=true`. Raw metric counts in `perf/perf_metrics.jsonl`: `folder.frame.total_us` 100, `folder.frame.present_us` 100, `folder.frame.visible_work_count` 100, `folder.frame.input_to_paint_us` 12, `folder.frame.dirty_rect_area_px` 100.
- Task 6 hygiene evidence: the placeholder scan over the plan and touched files exited `1` with no matches. `git diff --check` exited `0`; it printed only line-ending normalization warnings for the five touched files.
- Task 6 review topic: spec review found the artifact presence block omitted `folder.frame.dirty_rect_area_px` even though Task 6 emits it, and code review found the scroll stress still sent unhandled `WM_VSCROLL` messages. This does not change optimization scope, but it improves validation quality. The Task 6 review fix owns: add `dirty_rect_area_px` to the presence artifact with per-metric counts, bound metric scanning to the case's starting `perf_metrics.jsonl` offset, replace unhandled vertical-scroll steps with handled wheel/key navigation, and move input-to-paint timing wrappers out of `WndProc` into interaction handlers.
- Task 6 review-fix evidence: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_154340_146.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_154534`; `commands_results.json` reports 1 passed / 0 failed / 0 skipped. `perf/folderView_perf_scroll_render_stress_metrics.json` now records `folderFrameMetricPresence.allPresent=true` with counts from the scenario-local JSONL offset: `folder.frame.total_us` 109, `folder.frame.present_us` 109, `folder.frame.visible_work_count` 109, `folder.frame.input_to_paint_us` 23, and `folder.frame.dirty_rect_area_px` 109. `git diff --check` exited `0` with line-ending normalization warnings only, and the placeholder scan over the plan and touched files exited `1` with no matches.
- Task 7 metric-gate decision: no FolderView overlay or idle/distant-state optimization was justified, so no code files were changed. Same-machine comparison used baseline `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_132547` and Task 6 review-fix candidate `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_154534`; both `commands_results.json` files report 1 passed / 0 failed / 0 skipped. Selected p50/p95/p99 rows: `render.frame_us` 28643/48848/52752 baseline vs 27437/50142/53337 candidate; `render.present_us` 344/525/611 vs 455/768/4260; `render.dirty_rect_area_px` p50/p95/p99 1161911/1161911/1188249 in both; `render.items_drawn` 38/62/62 in both; `folder.frame.total_us` candidate-only 25806/47761/50948; `folder.frame.present_us` candidate-only 453/765/4257; `folder.frame.visible_work_count` candidate-only 38/62/62; `folder.frame.input_to_paint_us` candidate-only 28308/48127/52838.
- Task 7 overlay gate failed: `folder.frame.overlay_animation_count` was absent from baseline, early Task 6 candidate, and Task 6 review-fix candidate; `render.incremental_search_effect_updates` was present but summed to 0 in all compared archives; direct `rg -n -i "overlay"` over the three input archives exited `1` with no matches. Dirty rects were mostly full-window-sized during the scroll stress (`1161911` px for 97/100 baseline frames and 106/109 candidate frames, plus `1188249` px for 3 frames in both), but this archive does not exercise overlay animation, so the dirty-rect evidence cannot justify overlay invalidation changes.
- Task 7 distant-state release gate failed: visible/text/icon work remained bounded to visible rows rather than retained distant state. Candidate counts and quantiles showed `render.item_textlayout_label` 109 rows, sum 3740, p50/p95/p99 38/62/62; `render.item_textlayout_details` sum 2148, p50/p95/p99 23/38/38; `render.item_textlayout_metadata` sum 0; `render.layout_items_us` count stayed 30. Icon guardrails stayed essentially unchanged: `icons.batch_update_retrieved` sum 0, `icons.queue_groups_total` sum 3, `icons.queue_visible_groups` sum 1, `icons.ui_apply_count` 1600 once, and `icons.invalidate_full_count` 1 once. No memory or thumbnail-retention metric in the archives showed repeated distant text/icon state cost.
- New validation topic for Task 7: add or identify an overlay-specific FolderView scenario before attempting overlay invalidation optimization. This was added because the available scroll-render archives produce full-window dirty rects but no overlay-animation rows or overlay trace evidence; it improves validation only and does not expand Task 7 code scope.
- Task 8 started on `codex/dxui-frame-performance`. Scope stayed in ColorTextView, Monitor chrome selftest, and the monitor project opt-in test gate; `Tests/MonitorTest/MonitorTest.cpp` did not need changes.
- Task 8 selftest seam: `RedSalamanderMonitor.exe --chrome-selftest --perf` writes a `monitorFrameMetricPresence` block to `results.json` with per-metric counts for required chrome-scenario metrics. After the Task 8 review fix, required presence covers `monitor.frame.total_us`, `monitor.frame.present_us`, `monitor.frame.append_to_visible_us`, `monitor.frame.tail_layout_us`, `monitor.frame.mode`, and `monitor.etw.batch_drain_us`, and missing required rows fail the scenario. `monitor.frame.scrollback_slice_us` remains an optional ColorTextView metric for real SCROLL_BACK rendering and is not required by the chrome scenario.
- Task 8 build-gate correction: `RedSalamanderMonitor.vcxproj` keeps default Debug selftest hooks off, but now lets `RSBuildEnableTests=true` leave the inherited test default enabled. Red evidence confirmed the opt-in build path: `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; log `.build/logs/msbuild-20260519_160142_121.log`; diagnostics `0 warning(s), 0 error(s)`.
- Task 8 red evidence: after adding only selftest metric-presence attribution and the opt-in build-gate/selftest baseline fix, `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_160158`; `results.json` reports `status=passed` and `monitorFrameMetricPresence.allPresent=false` with count `0` for all seven required monitor frame/ETW metrics. Direct `rg` for those metric rows in `perf_metrics.jsonl` exited `1`.
- Task 8 implementation: `ColorTextView` uses the shared `DxUi::FrameClock`/`FrameStageScope` for paint/render/present timing, emits `monitor.frame.total_us`, `monitor.frame.present_us`, `monitor.frame.tail_layout_us`, `monitor.frame.scrollback_slice_us`, and `monitor.frame.mode` while preserving existing `Present1` parameters, dirty rect handling, AUTO_SCROLL tail layout, SCROLL_BACK slice/fallback behavior, and display-row mapping. ETW UI-thread batch processing emits `monitor.etw.batch_drain_us`; AUTO_SCROLL ETW appends mark a pending append-to-visible timestamp before invalidation and emit `monitor.frame.append_to_visible_us` after the next successful tail paint/present. `SwitchToScrollBackMode()` clears pending append-to-visible work so SCROLL_BACK does not emit that metric.
- Task 8 green evidence before final verification: `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; log `.build/logs/msbuild-20260519_160438_988.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_160452`; `results.json` reports `status=passed` and `monitorFrameMetricPresence.allPresent=true` with counts: `monitor.frame.total_us` 7, `monitor.frame.present_us` 7, `monitor.frame.append_to_visible_us` 1, `monitor.frame.tail_layout_us` 16, `monitor.frame.scrollback_slice_us` 1, `monitor.frame.mode` 7, and `monitor.etw.batch_drain_us` 2.
- Task 8 final verification: `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; log `.build/logs/msbuild-20260519_160726_813.log`; diagnostics `0 warning(s), 0 error(s)`. `.\build.ps1 -ProjectName MonitorTest -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_160742_959.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\MonitorTest.exe --document-model-selftest` exited `0`. `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_160755`; `results.json` reports `status=passed` and `monitorFrameMetricPresence.allPresent=true` with counts: `monitor.frame.total_us` 7, `monitor.frame.present_us` 7, `monitor.frame.append_to_visible_us` 1, `monitor.frame.tail_layout_us` 16, `monitor.frame.scrollback_slice_us` 1, `monitor.frame.mode` 7, and `monitor.etw.batch_drain_us` 2.
- Task 8 hygiene evidence: `git diff --check` exited `0` with line-ending normalization warnings only. The non-self-matching placeholder scan over touched source/project/plan files exited `1` with no matches.
- Task 8 review-fix topic: code/spec review found metric attribution issues in the initial Task 8 slice. The follow-up makes `monitor.frame.scrollback_slice_us` emit only when `_renderMode == SCROLL_BACK`; invalid or empty AUTO_SCROLL tail fallback still renders through the existing fallback path but no longer pollutes scrollback evidence. `monitor.frame.append_to_visible_us` now starts before `_document.AppendInfoLines(...)` for AUTO_SCROLL ETW drain work and preserves the earliest pending timestamp until a successful AUTO_SCROLL paint emits it. `monitor.etw.batch_drain_us` no longer emits for stale empty posted batch messages. The chrome perf selftest now treats required metric presence as a pass/fail requirement and no longer toggles auto-scroll solely to manufacture a SCROLL_BACK frame. Scenario contract change: `monitor.frame.scrollback_slice_us` remains a ColorTextView metric for real SCROLL_BACK rendering, but it is not required by `--chrome-selftest --perf` because this chrome scenario does not legitimately exercise user scrollback.
- Task 8 review-fix verification: `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; log `.build/logs/msbuild-20260519_161631_494.log`; diagnostics `0 warning(s), 0 error(s)`. `.\build.ps1 -ProjectName MonitorTest -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_161649_204.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\MonitorTest.exe --document-model-selftest` exited `0`. `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_161700`; `results.json` reports `status=passed`, includes a passed `required monitor frame metrics` check, and `monitorFrameMetricPresence.allPresent=true` with counts: `monitor.frame.total_us` 5, `monitor.frame.present_us` 5, `monitor.frame.append_to_visible_us` 1, `monitor.frame.tail_layout_us` 15, `monitor.frame.mode` 5, and `monitor.etw.batch_drain_us` 2. The chrome scenario produced no `monitor.frame.scrollback_slice_us` row after removing the artificial scrollback toggle. `git diff --check` exited `0` with line-ending normalization warnings only; the non-self-matching placeholder scan exited `1` with no matches.
- Task 9 metric-gate decision before fresh verification: no Monitor batch-drain or paint-coalescing optimization is justified from latest legitimate Task 8 archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_161938`. Required gate metrics: `monitor.etw.selftest_burst_drain_us` count 1 value 59426 us; `monitor.etw.batch_drain_us` count 2 min 2164 us max 6832 us avg 4498 us; `monitor.frame.append_to_visible_us` count 1 value 58001 us; `monitor.frame.tail_layout_us` count 15 p50 2 us max 282 us; `monitor.frame.total_us` count 5 min 418 us p50 1856 us p95 47717 us max 145712 us; `monitor.frame.present_us` count 5 min 218 us p50 303 us max 1537 us. Attribution: the 145712 us `monitor.frame.total_us` row occurs before ETW batch drains and aligns with startup/initial paint, so it is not evidence for Task 9 scheduling work. The two ETW batch drains occur before one later `monitor.frame.present_us`/`monitor.frame.total_us`/`monitor.frame.append_to_visible_us` emission, so the archive does not show repeated ETW batches causing multiple ColorTextView paints before one present. Step 2 gate failed because no single UI-thread batch monopolized a frame. Step 3 gate failed because the measured sequence already coalesced two ETW batches into one ColorTextView present. Code remains unchanged for Task 9.
- Task 9 verification and fresh metric scan: `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; log `.build/logs/msbuild-20260519_162347_170.log`; diagnostics `0 warning(s), 0 error(s)`. `.\build.ps1 -ProjectName MonitorTest -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_162403_557.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\MonitorTest.exe --document-model-selftest` exited `0`. `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_162416`; `results.json` reports `status=passed` with summary `Monitor DxUI toolbar/status strip selftest passed.` Fresh gate metrics: `monitor.etw.selftest_burst_drain_us` count 1 value 74584 us; `monitor.etw.batch_drain_us` count 2 min 3538 us max 7709 us avg 5623.5 us; `monitor.frame.append_to_visible_us` count 1 value 73372 us; `monitor.frame.tail_layout_us` count 15 p50 2 us max 286 us; `monitor.frame.total_us` count 5 min 519 us p50 2810 us p95 61037 us max 159141 us; `monitor.frame.present_us` count 5 min 229 us p50 319 us max 1152 us. The sequence again shows the largest `monitor.frame.total_us` before ETW batch drains, then both ETW batch drains before one ColorTextView present/append-to-visible emission. `git diff --check` exited `0` with the existing LF-to-CRLF normalization warning for this plan file only.
- Task 10 current swap-chain inventory: DxUi `WindowHost` uses `IDXGISwapChain1`, two buffers, `DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL`, `CreateSwapChainForHwnd` for normal hosts and `CreateSwapChainForComposition` for DirectComposition popup hosts, `Present1` dirty rectangles for partial paints and `Present` for full paints, `ResizeBuffers(0, ...)` after detaching the D2D target and flushing D3D, and device-lost handling on `D2DERR_RECREATE_TARGET`, `DXGI_ERROR_DEVICE_REMOVED`, and `DXGI_ERROR_DEVICE_RESET`. FolderView uses an `IDXGISwapChain1` HWND flip-sequential path with `kSwapChainBufferCount`, `Present1` dirty rectangles, `ResizeBuffers(kSwapChainBufferCount, ...)`, and a legacy `IDXGISwapChain`/`DXGI_SWAP_EFFECT_DISCARD` fallback when flip-model creation fails; present or EndDraw failures release and recreate the swap chain. ColorTextView uses an `IDXGISwapChain1` HWND flip-sequential path with two buffers, `Present1` for both full and partial paints, `pDirtyRects` plus `pScrollRect`/`pScrollOffset` for scroll reuse, `ResizeBuffers(0, ...)` or full swap-chain recreation after detaching D2D resources, and device-lost recovery through resource discard plus full redraw. None of the three surfaces currently queries or stores `IDXGISwapChain2`; therefore frame-latency waitable-object support is not established in production code.
- Task 10 flip-discard gate decision: failed. The existing rendering contract depends on flip-sequential buffer preservation: DxUi and FolderView clip drawing to dirty rectangles and present only those dirty rectangles; ColorTextView additionally relies on DXGI scroll-rect partial-present semantics. Switching to `DXGI_SWAP_EFFECT_FLIP_DISCARD` would require an explicit full-redraw or retained-content fallback for partial paints before it could be correct. No such fallback was implemented, and no before/after same-machine candidate metrics were generated because the semantic gate failed first. Current same-machine present metric evidence remains on the existing policy: FolderView Task 6/7 candidate archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_154534` has `folder.frame.present_us` count 109, avg 476.7 us, p50 453 us, p95 746 us, p99 1671 us, max 4553 us; Monitor Task 9 archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_162416` has `monitor.frame.present_us` count 5, avg 475.8 us, p50 319 us, max 1152 us; local DxUi scratch `Specs/TestRuns/local_scratch/dxui_windowhost_stage_metrics_testlocal_20260519.jsonl` has `dxui.frame.present_us` count 3, avg 519.7 us.
- Task 10 frame-latency waitable-object gate decision: failed. The current WM_PAINT-driven surfaces present from paint/render paths and have resize/device-loss paths that synchronously detach D2D/D3D references. Adding `DXGI_SWAP_CHAIN_FLAG_FRAME_LATENCY_WAITABLE_OBJECT` would need `IDXGISwapChain2` capability probing, wait-handle ownership, message-pump integration outside `WM_PAINT`, and explicit minimize/occlusion behavior. Because those prerequisites are not present and no PresentMon or `present_us` variance candidate shows improvement, no waitable-object flag or wait was added.
- Task 10 no-op implementation decision: keep the current flip-sequential present policy for all three surfaces. This is a measured plan-only closeout, not a production code change. It preserves dirty-rect and scroll-rect behavior while leaving a future experiment possible only after adding explicit partial-present fallbacks and waitable-object message-pump tests.
- Task 10 fresh protected render verification: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_163115_205.log`; diagnostics `0 warning(s), 0 error(s)`. `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_163124_826.log`; diagnostics `0 warning(s), 0 error(s)`. `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; log `.build/logs/msbuild-20260519_163316_088.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost` exited `0` with `All DxUi tests passed`; scratch metrics were refreshed at `Specs/TestRuns/local_scratch/dxui_windowhost_stage_metrics_testlocal_20260519.jsonl` and `dxui.frame.present_us` count 3, avg 446.7 us, p50 440 us, max 699 us. `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_163342`; `commands_results.json` reports 1 passed / 0 failed / 0 skipped; `folder.frame.present_us` count 109, avg 255.3 us, p50 272 us, p95 497 us, p99 576 us, max 653 us. `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_163340`; `results.json` reports `status=passed` and required monitor frame metrics present; `monitor.frame.present_us` count 5, avg 539.6 us, p50 424 us, max 1222 us. `git diff --check` exited `0` with the existing LF-to-CRLF normalization warning for this plan file only.
- Task 11 composition pilot gate decision: failed. The existing Task 3/4 evidence shows DxUi frame-stage telemetry and animation scheduler instrumentation are present, but it does not show a remaining production CPU-thread animation jitter problem that cannot be fixed by the frame runtime. The Task 4 jitter rows come from deterministic Animation suite scheduler-policy coverage and include deliberate harness tick gaps/hitches; they are not evidence that tooltip, popup, dialog smoke, or lightweight overlay transforms require compositor-backed animation. Fresh Task 11 metrics from `Specs/TestRuns/local_scratch/dxui_animation_task11_20260519.jsonl` were produced by `.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_task11_20260519.jsonl`; the command exited `0` with `All DxUi tests passed`. The scratch file has 27 rows: `dxui.animation.tick_delta_us` count 4, min 8333 us, max 31773 us, avg 16419.0 us; `dxui.animation.jitter_us` count 4, min 0 us, max 23440 us, avg 8086.0 us; `dxui.animation.active_count` count 4, min 0, max 1; `dxui.animation.tick_gap_ms` count 2, min 17 ms, max 31 ms; `dxui.animation.tick_overrun` count 1. This confirms metric coverage, not an unresolved product jitter scenario. Scope remains plan-only and no `Common/DxUi/DxUi.CompositionPilot.*`, `Common/DxUi/DxUi.vcxproj`, `Common/DxUi/DxUi.WindowHost.cpp`, or `Tests/DxUiTests/DxUiTests.Animation.cpp` changes were made for Task 11.
- Task 11 fresh verification: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_164726_221.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_task11_20260519.jsonl` exited `0` with `All DxUi tests passed`. `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost` exited `0` with `All DxUi tests passed`.

## Decision

Do not port DXUT into RedSalamander. DXUT is useful as a teaching reference for deterministic frame callbacks and centralized device lifecycle, but it is not the best production model for this codebase.

Use this approach instead:

1. Keep DxUi as the shared UI framework.
2. Keep FolderView as a specialized large-list/file-display renderer.
3. Keep ColorTextView as a specialized append-only log renderer with AUTO_SCROLL and SCROLL_BACK modes.
4. Share the frame contract, not the renderer.
5. Use performance evidence to decide every optimization.

This is the best maintainability/performance tradeoff because it removes duplicated timing, animation, and telemetry policy while preserving the highly tuned rendering logic in the two most important displays.

## External References

- [DXUT](https://github.com/microsoft/DXUT) - use for frame-loop and device-lifecycle inspiration only.
- [DirectXTK basic game loop](https://github.com/microsoft/DirectXTK/wiki/The-basic-game-loop) - better modern reference for explicit update/render loop structure.
- [DirectXTK DeviceResources](https://github.com/microsoft/DirectXTK/wiki/DeviceResources) - better modern reference for device and size-dependent resource separation.
- [DXGI flip-model, dirty rectangles, and scrolled areas](https://learn.microsoft.com/en-us/windows/win32/direct3ddxgi/dxgi-1-2-presentation-improvements) - relevant because all three target surfaces already use flip-model and partial present paths.
- [DXGI frame-latency waitable swap chains](https://learn.microsoft.com/en-us/windows/uwp/gaming/reduce-latency-with-dxgi-1-3-swap-chains) - evaluate as an opt-in capability, not as a first change.
- [Direct2D performance guidance](https://learn.microsoft.com/en-us/windows/win32/direct2d/improving-direct2d-performance) - source for resource reuse, batching, and avoiding unnecessary render-target churn.
- [Windows Visual Layer in desktop apps](https://learn.microsoft.com/en-us/windows/apps/desktop/modernize/ui/visual-layer-in-desktop-apps) - useful for optional compositor-driven overlay animation, but not for text/list virtualization.
- [DirectComposition animation](https://learn.microsoft.com/en-us/windows/win32/directcomp/how-to--animate-a-visual) - legacy-compatible animation reference if Visual Layer is rejected.
- [PresentMon](https://github.com/GameTechDev/PresentMon) - optional manual validation of present pacing and frame-time variance.

## Current Code Grounding

DxUi:

- `Common/DxUi/DxUi.WindowHost.cpp:1118` has `WindowHost::Invalidate()`.
- `Common/DxUi/DxUi.WindowHost.cpp:1134` has `WindowHost::RequestAnimation()`.
- `Common/DxUi/DxUi.WindowHost.cpp:1877` handles `WM_PAINT`.
- `Common/DxUi/DxUi.WindowHost.cpp:1885` handles `WM_SIZE`.
- `Common/DxUi/DxUi.WindowHost.cpp:2408` creates device resources.
- `Common/DxUi/DxUi.WindowHost.cpp:2475` creates size-dependent resources.
- `Common/DxUi/DxUi.WindowHost.cpp:2513` uses `DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL`.
- `Common/DxUi/DxUi.WindowHost.cpp:2724` renders with `DxUI::Paint`.
- `Common/DxUi/DxUi.WindowHost.cpp:2811` presents dirty rectangles through `Present1`.
- `Common/DxUi/DxUi.WindowHost.cpp:3355` runs animation ticks.
- `RedSalamander/Ui/AnimationDispatcher.h:114` uses a 16 ms timer interval.
- `RedSalamander/Ui/AnimationDispatcher.h:230` uses `SetTimer`.
- `RedSalamander/Ui/AnimationDispatcher.h:297` uses `GetTickCount64`.
- `RedSalamander/Ui/AnimationDispatcher.h:301` emits animation tick-gap metrics.

FolderView:

- `RedSalamander/FolderView.Rendering.cpp:162` creates D2D/D3D device resources.
- `RedSalamander/FolderView.Rendering.cpp:432` creates the swap chain.
- `RedSalamander/FolderView.Rendering.cpp:942` renders a dirty rectangle.
- `RedSalamander/FolderView.Rendering.cpp:960` emits `render.frame_us`.
- `RedSalamander/FolderView.Rendering.cpp:1808` through `1817` emits visible-work and dirty-area counters.
- `RedSalamander/FolderView.Rendering.cpp:1842` uses `Present1`.
- `RedSalamander/FolderView.Layout.cpp:64` emits `render.layout_items_us`.
- `RedSalamander/FolderView.Layout.cpp:401` updates visible item text layouts.
- `RedSalamander/FolderView.Layout.cpp:647` has distant rendering-state release logic.
- `RedSalamander/FolderView.Layout.cpp:865` and `908` schedule and process idle layout creation.
- Existing command selftests already cover large-folder baseline, sort toggle stress, scroll/render stress, directory-change storm, icon-cache contention, and thumbnail scroll.

RedSalamanderMonitor / ColorTextView:

- `RedSalamanderMonitor/ColorTextView.cpp:1263` creates device resources.
- `RedSalamanderMonitor/ColorTextView.cpp:1373` uses `DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL`.
- `RedSalamanderMonitor/ColorTextView.cpp:1473` creates swap-chain resources.
- `RedSalamanderMonitor/ColorTextView.cpp:1623` matches the backbuffer to the client.
- `RedSalamanderMonitor/ColorTextView.cpp:2201` draws the scene.
- `RedSalamanderMonitor/ColorTextView.cpp:2486` handles paint.
- `RedSalamanderMonitor/ColorTextView.cpp:2749` uses `Present1`.
- `RedSalamanderMonitor/ColorTextView.cpp:2873` invalidates.
- `RedSalamanderMonitor/ColorTextView.cpp:2912` switches to AUTO_SCROLL.
- `RedSalamanderMonitor/ColorTextView.cpp:2933` switches to SCROLL_BACK.
- `RedSalamanderMonitor/ColorTextView.cpp:2969` rebuilds tail layout.
- `RedSalamanderMonitor/ColorTextView.cpp:3191` adapts layout.
- `Specs/Core/Core_RedSalamanderMonitor.md` documents the append-only, high-throughput, two-mode renderer contract.

## Non-Goals

- Do not rewrite FolderView as a generic DxUi grid.
- Do not rewrite ColorTextView as a generic DxUi text control.
- Do not move D3D/DXGI ownership into a global singleton.
- Do not introduce Visual Layer or DirectComposition as the first optimization.
- Do not remove existing `Present1`, dirty-rect, visible-work, or two-mode rendering paths.
- Do not claim better performance without same-machine baseline and candidate evidence.

## Success Criteria

- DxUi animation uses a high-resolution frame clock and stage telemetry instead of depending on a fixed 16 ms `SetTimer` cadence.
- DxUi render-stage code is debug-guarded against layout mutation during paint.
- FolderView retains current visible-work rendering behavior while reporting comparable frame-stage metrics.
- ColorTextView retains AUTO_SCROLL and SCROLL_BACK contracts while reporting append-to-visible and frame-stage metrics.
- Active animation/input renders continuously only while needed; idle, minimized, and occluded windows throttle.
- Existing DxUi, FolderView, and Monitor tests remain green.
- Same-machine candidate evidence shows no p50/p95/p99 regression in protected scenarios; targeted optimizations need measurable improvement in the metric they target.

## Shared Metric Contract

Use these metric families. Keep existing `render.*` metrics, then add surface-prefixed metrics where attribution is missing.

| Surface | Metrics |
| --- | --- |
| Shared runtime | `ui.frame.total_us`, `ui.frame.update_us`, `ui.frame.layout_us`, `ui.frame.render_us`, `ui.frame.present_us`, `ui.frame.dirty_rect_count`, `ui.frame.dirty_rect_area_px`, `ui.frame.active_animation_count`, `ui.frame.over_budget_count`, `ui.frame.mode` |
| DxUi | `dxui.frame.total_us`, `dxui.frame.update_us`, `dxui.frame.layout_us`, `dxui.frame.render_us`, `dxui.frame.present_us`, `dxui.animation.tick_delta_us`, `dxui.animation.jitter_us`, `dxui.animation.active_count`, `dxui.layout.dirty_subtree_count` |
| FolderView | existing `render.*`, plus `folder.frame.total_us`, `folder.frame.present_us`, `folder.frame.visible_work_count`, `folder.frame.overlay_animation_count`, `folder.frame.input_to_paint_us` |
| Monitor | `monitor.frame.total_us`, `monitor.frame.present_us`, `monitor.frame.append_to_visible_us`, `monitor.frame.tail_layout_us`, `monitor.frame.scrollback_slice_us`, `monitor.etw.batch_drain_us`, existing `monitor.etw.selftest_burst_drain_us` |

Budget interpretation:

- 60 Hz target: 16,667 us.
- 120 Hz target: 8,333 us.
- 144 Hz target: 6,944 us.
- Treat hard thresholds as scenario-specific. Use before/after same-machine comparisons as the primary gate because UI test hardware varies.

## Proposed File Structure

Create in `Common/DxUi`:

- `DxUi.FrameRuntime.h` - public lightweight clock, frame-stage, budget, and scheduling contracts usable by app renderers.
- `DxUi.FrameRuntime.cpp` - QPC clock, delta smoothing, metric emission helpers, debug stage state.

Modify:

- `Common/DxUi/DxUi.vcxproj`
- `Common/DxUi/DxUi.Internal.h`
- `Common/DxUi/DxUi.WindowHost.cpp`
- `Common/DxUi/DxUi.h` only if the runtime contract must be public outside the static library.
- `RedSalamander/Ui/AnimationDispatcher.h`
- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.cpp`
- `RedSalamander/FolderView.Rendering.cpp`
- `RedSalamander/FolderView.Layout.cpp`
- `RedSalamander/FolderView.ErrorOverlay.cpp`
- `RedSalamanderMonitor/ColorTextView.h`
- `RedSalamanderMonitor/ColorTextView.cpp`
- `RedSalamanderMonitor/RedSalamanderMonitor.cpp`
- `Tests/DxUiTests/DxUiTests.cpp`
- `Tests/DxUiTests/DxUiTests.Animation.cpp`
- `Tests/DxUiTests/DxUiTests.WindowHost.cpp`
- `Tests/MonitorTest/MonitorTest.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/Core/Core_RedSalamanderMonitor.md`

## Task 1: Capture Protected Baselines

**Files:**

- Read: `Specs/Testing/Testing_PerformanceValidation.md`
- Read: `Specs/TestRuns/README.md`
- Update: `Specs/Plans/WIP/DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md`

- [x] **Step 1: Verify clean status before evidence**

Run:

```powershell
git status --short
```

Expected: either empty output or only unrelated user-owned changes that are recorded in the baseline notes.

- [x] **Step 2: Build protected projects**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug
.\build.ps1 -ProjectName MonitorTest -Configuration Debug
```

Expected: each build exits `0`. Record the `.build/logs/msbuild-*.log` paths in the evidence section of this plan.

- [x] **Step 3: Run DxUi frame and animation baselines**

Run:

```powershell
.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost --perf-jsonl=Specs\TestRuns\local_scratch\dxui_windowhost_frame_runtime_baseline_20260519.jsonl
.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_frame_runtime_baseline_20260519.jsonl
```

Expected: both suites exit `0`, and both JSONL files contain DxUi timing rows.

- [x] **Step 4: Run FolderView protected scenarios**

Run:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_sort_toggle_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_iconcache_contention --selftest-timeout-multiplier=4
```

Expected: each run passes and prints an `ArchiveToRepo:` path under `Specs\TestRuns`.

- [x] **Step 5: Run Monitor protected scenarios**

Run:

```powershell
.\.build\x64\Debug\MonitorTest.exe --diagnostics-gate-selftest
.\.build\x64\Debug\MonitorTest.exe --scrollbar-model-selftest
.\.build\x64\Debug\MonitorTest.exe --document-model-selftest
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: each run exits `0`. The `--chrome-selftest --perf` run archives Monitor selftest evidence under `Specs/TestRuns/.../Monitor/.../`.

- [x] **Step 6: Add baseline evidence section**

Add a `Baseline Evidence` section near the end of this plan with:

- build logs,
- test commands,
- archive paths,
- pass/fail counts,
- metric keys found in each `perf_metrics.jsonl` or local scratch JSONL,
- same-machine caveats.

## Task 2: Add Shared Frame Runtime Skeleton

**Files:**

- Create: `Common/DxUi/DxUi.FrameRuntime.h`
- Create: `Common/DxUi/DxUi.FrameRuntime.cpp`
- Modify: `Common/DxUi/DxUi.vcxproj`
- Note: `Common/DxUi/DxUi.vcxproj.filters` is absent in this repo and was not created for Task 2.
- Test: `Tests/DxUiTests/DxUiTests.Animation.cpp`

- [x] **Step 1: Add failing clock tests**

Add tests in `Tests/DxUiTests/DxUiTests.Animation.cpp` that verify:

- `FrameClock::Now()` is monotonic.
- `FrameClock::ElapsedUs(a, b)` is non-negative when `b >= a`.
- `SmoothDeltaUs()` clamps a 500 ms hitch to the configured maximum.
- Reduced-motion mode can return a non-animated immediate-progress policy.

Expected new test names:

```cpp
TestFrameRuntimeClockIsMonotonic();
TestFrameRuntimeClampsLargeDelta();
TestFrameRuntimeElapsedUsHandlesLargeQpcDelta();
TestFrameRuntimeElapsedUsHandlesWideCrossSignDelta();
TestFrameRuntimeReducedMotionPolicy();
```

- [x] **Step 2: Run test to verify failure**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=Animation
```

Expected: build or test fails because `DxUi.FrameRuntime.h` does not exist yet.

- [x] **Step 3: Add runtime header contract**

Create `Common/DxUi/DxUi.FrameRuntime.h` with these concrete concepts:

```cpp
#pragma once

#include <cstdint>
#include <string_view>

namespace RedSalamander::DxUi
{
enum class FrameStage : uint8_t
{
    Idle,
    Input,
    Update,
    Layout,
    Render,
    Present,
};

enum class FrameMode : uint8_t
{
    Idle,
    EventDriven,
    ActiveAnimation,
    Occluded,
};

struct FrameBudget
{
    uint64_t refreshPeriodUs = 16667;
    uint64_t hitchClampUs = 50000;
};

struct FrameTimestamp
{
    int64_t qpc = 0;
};

class FrameClock final
{
public:
    FrameClock() noexcept;
    [[nodiscard]] FrameTimestamp Now() const noexcept;
    [[nodiscard]] uint64_t ElapsedUs(FrameTimestamp start, FrameTimestamp end) const noexcept;
    [[nodiscard]] uint64_t SmoothDeltaUs(uint64_t rawDeltaUs, const FrameBudget& budget) noexcept;

private:
    int64_t _frequency = 1;
    uint64_t _lastSmoothedDeltaUs = 16667;
};

class FrameStageScope final
{
public:
    FrameStageScope(FrameStage& currentStage, FrameStage nextStage) noexcept;
    ~FrameStageScope() noexcept;
    FrameStageScope(const FrameStageScope&) = delete;
    FrameStageScope& operator=(const FrameStageScope&) = delete;

private:
    FrameStage& _currentStage;
    FrameStage _previousStage = FrameStage::Idle;
};

void EmitFrameMetric(std::wstring_view metric, uint64_t valueUs) noexcept;
}
```

- [x] **Step 4: Add runtime implementation**

Implement with `QueryPerformanceFrequency`, `QueryPerformanceCounter`, WIL-free value types, and `Debug::Perf::EmitValue` / `Debug::Perf::EmitDurationUs` according to the existing metrics helpers available in the tree.

Implementation rules:

- no global mutable singleton,
- no heap allocation in `Now()` or `ElapsedUs()`,
- no exception throwing,
- no raw owning COM or Win32 handles,
- no logging normal timer jitter as errors.

- [x] **Step 5: Add project entries**

Add `DxUi.FrameRuntime.cpp` and `DxUi.FrameRuntime.h` to:

- `Common/DxUi/DxUi.vcxproj`

Do not create `Common/DxUi/DxUi.vcxproj.filters`; it is absent in this repo.

- [x] **Step 6: Run test to verify pass**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=Animation
```

Expected: `Animation` suite exits `0`.

- [x] **Step 7: Commit**

Run:

```powershell
git add Common/DxUi/DxUi.FrameRuntime.h Common/DxUi/DxUi.FrameRuntime.cpp Common/DxUi/DxUi.vcxproj Tests/DxUiTests/DxUiTests.Animation.cpp
git commit -m "feat(dxui): add frame runtime clock"
```

## Task 3: Add DxUi Frame-Stage Telemetry

**Files:**

- Modify: `Common/DxUi/DxUi.WindowHost.cpp`
- Modify: `Common/DxUi/DxUi.Internal.h`
- Modify: `Tests/DxUiTests/DxUiTests.WindowHost.cpp`

- [x] **Step 1: Add failing WindowHost telemetry test**

Add a focused `WindowHost` test that:

- creates a small host,
- sets a root control,
- invalidates once,
- renders through the existing capture path,
- asserts new metrics are emitted for update, render, present, and total frame time.

Expected metric names:

```text
dxui.frame.total_us
dxui.frame.update_us
dxui.frame.render_us
dxui.frame.present_us
```

- [x] **Step 2: Verify the test fails**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost --perf-jsonl=Specs\TestRuns\local_scratch\dxui_windowhost_stage_metrics_red_20260519.jsonl
```

Expected: test fails because the metrics are not emitted.

- [x] **Step 3: Add stage scopes to `WindowHost::Render`**

In `Common/DxUi/DxUi.WindowHost.cpp`, add `FrameClock`, `FrameStage`, and stage scopes around:

- animation/root update before paint,
- layout if a layout-dirty flag is present,
- render,
- present.

Keep the existing `Debug::Perf::Scope paintPerf(L"DxUI::Paint")` metric for compatibility.

- [x] **Step 4: Add present timing**

Wrap the existing `Present1` and fallback `Present` calls in present-stage timing. Emit:

- `dxui.frame.present_us`
- `dxui.frame.total_us`
- `dxui.frame.dirty_rect_count`
- `dxui.frame.dirty_rect_area_px`

Use zero dirty-rect area when rendering the full frame or when no valid dirty rect exists.

- [x] **Step 5: Run focused tests**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost --perf-jsonl=Specs\TestRuns\local_scratch\dxui_windowhost_stage_metrics_green_20260519.jsonl
```

Expected: suite exits `0`, and the JSONL file contains the new `dxui.frame.*` metrics.

- [x] **Step 6: Commit**

Run:

```powershell
git add Common/DxUi/DxUi.WindowHost.cpp Common/DxUi/DxUi.Internal.h Tests/DxUiTests/DxUiTests.WindowHost.cpp
git commit -m "feat(dxui): emit frame stage metrics"
```

## Task 4: Replace Fixed Animation Timer Policy

**Files:**

- Modify: `RedSalamander/Ui/AnimationDispatcher.h`
- Modify: `Common/DxUi/DxUi.WindowHost.cpp`
- Modify: `Tests/DxUiTests/DxUiTests.Animation.cpp`
- Modify: `RedSalamander/RedSalamander.vcxproj` only if a new implementation file is created.

- [x] **Step 1: Add failing jitter and active-subscription tests**

Add tests that verify:

- active animation subscribers receive a monotonic high-resolution timestamp,
- a subscriber returning inactive stops continuous frame requests,
- a synthetic 120 Hz budget uses an 8,333 us target,
- a hitch does not produce unbounded interpolation delta.

- [x] **Step 2: Verify tests fail**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_scheduler_red_20260519.jsonl
```

Expected: tests fail because current `AnimationDispatcher` is fixed at 16 ms and uses `GetTickCount64`.

- [x] **Step 3: Convert timestamp source**

Change `AnimationDispatcher` to use `DxUi::FrameClock` for timestamps and deltas while preserving the current subscription API:

```cpp
using TickCallback = bool (*)(void* context, uint64_t nowTickMs) noexcept;
```

Keep `nowTickMs` for compatibility during this task, but compute it from the high-resolution clock so jitter metrics can use microseconds internally.

- [x] **Step 4: Add active/idle scheduling policy**

Keep the hidden message window and `SetTimer` fallback for compatibility, but change behavior:

- active subscribers request continuous ticks,
- no active subscribers stop the timer,
- minimized/occluded owners can skip nonessential animation ticks,
- reduced-motion policy resolves animations immediately where controls support it.

- [x] **Step 5: Emit improved animation metrics**

Emit:

- `dxui.animation.tick_delta_us`
- `dxui.animation.jitter_us`
- `dxui.animation.active_count`
- keep existing `dxui.animation.tick_gap_ms` and `dxui.animation.tick_overrun` during migration.

- [x] **Step 6: Run focused tests**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_scheduler_green_20260519.jsonl
```

Expected: suite exits `0`, JSONL contains new animation metrics, and existing tests still pass.

- [x] **Step 7: Commit**

Run:

```powershell
git add RedSalamander/Ui/AnimationDispatcher.h Common/DxUi/DxUi.WindowHost.cpp Tests/DxUiTests/DxUiTests.Animation.cpp
git commit -m "feat(dxui): improve animation frame scheduling"
```

## Task 5: Add Debug Guard Against Layout During Render

**Files:**

- Modify: `Common/DxUi/DxUi.FrameRuntime.h`
- Modify: `Common/DxUi/DxUi.FrameRuntime.cpp`
- Modify: `Common/DxUi/DxUi.Internal.h`
- Modify: `Common/DxUi/DxUi.Controls.cpp`
- Modify: `Common/DxUi/DxUi.Grid.cpp`
- Modify: `Common/DxUi/DxUi.Tree.cpp`
- Modify: `Common/DxUi/DxUi.WindowHost.cpp`
- Modify: `Common/DxUi/DxUi.cpp` (Task 5 correction: central `Control::SetBounds` geometry mutation guard lives here)
- Test: `Tests/DxUiTests/DxUiTests.WindowHost.cpp`

- [x] **Step 1: Add failing render-mutation test**

Add a test-only control that attempts to request layout from `Paint()`. The test should expect a debug assertion path or a diagnostic counter in debug builds.

Expected metric/counter:

```text
dxui.frame.render_layout_mutation_blocked
```

Red evidence: after adding `TestWindowHostBlocksLayoutMutationDuringRender`, `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_151341_054.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost` exited `1` at `FAILED: render layout mutation keeps control bounds unchanged`, proving the paint-time `SetBounds` mutation was not yet blocked/counted.

- [x] **Step 2: Add stage query helper**

Expose a debug-only helper through internal DxUi APIs:

```cpp
[[nodiscard]] bool IsDxUiRenderStageActiveForDebug() noexcept;
void EmitDxUiRenderMutationBlockedForDebug() noexcept;
```

- [x] **Step 3: Guard layout mutation entry points**

Add guard checks to layout/invalidating code paths that can mutate tree geometry. The guard should:

- emit `dxui.frame.render_layout_mutation_blocked`,
- use the diagnostic counter path instead of a modal assertion,
- avoid throwing exceptions,
- not block visual-only invalidation.

Implementation note: the narrow guard was placed in `Control::SetBounds` in `Common/DxUi/DxUi.cpp`, the base geometry mutation entry point. `Control::RequestInvalidate` was left unchanged so visual-only invalidation during render remains allowed. No extra guards were added in Controls/Grid/Tree/WindowHost because the central `SetBounds` guard covers geometry mutation without broadening behavior.

- [x] **Step 4: Run focused tests**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost
```

Expected: suite exits `0` with the intentional mutation test passing through the expected debug diagnostic path.

Green evidence: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited `0`; log `.build/logs/msbuild-20260519_151441_057.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost` exited `0`; `TestWindowHostBlocksLayoutMutationDuringRender` passed and the suite reported `All DxUi tests passed.`

- [x] **Step 5: Commit**

Run:

```powershell
git add Common/DxUi/DxUi.FrameRuntime.h Common/DxUi/DxUi.FrameRuntime.cpp Common/DxUi/DxUi.Internal.h Common/DxUi/DxUi.cpp Tests/DxUiTests/DxUiTests.WindowHost.cpp Specs/Plans/WIP/DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md
git commit -m "test(dxui): guard layout mutation during render"
```

## Task 6: Adopt Frame Runtime In FolderView Without Behavioral Optimization

**Files:**

- Modify: `RedSalamander/FolderView.h`
- Modify: `RedSalamander/FolderView.Rendering.cpp`
- Modify: `RedSalamander/FolderView.cpp`
- Modify: `RedSalamander/FolderView.Interaction.cpp` for measured input handler wrappers if review cleanup is needed.
- Modify: `RedSalamander/RedSalamander.vcxproj` only if new files are created.
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

- [x] **Step 1: Add FolderView metric assertions to existing selftest artifacts**

Extend `folderView_perf_scroll_render_stress` to record whether these metrics appeared:

```text
folder.frame.total_us
folder.frame.present_us
folder.frame.visible_work_count
folder.frame.input_to_paint_us
```

Do not add pass/fail thresholds in this task. The first purpose is attribution.

- [x] **Step 2: Verify missing metrics**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
```

Expected: run passes functionally but the artifact records missing new frame metrics.

- [x] **Step 3: Add FolderView frame metric scopes**

In `RedSalamander/FolderView.Rendering.cpp`, wrap the existing render path with `DxUi::FrameClock` timing while preserving:

- current `render.frame_us`,
- current dirty-rect render behavior,
- current `Present1` params,
- current visible item counters,
- current device-lost handling.

Emit:

- `folder.frame.total_us`
- `folder.frame.present_us`
- `folder.frame.visible_work_count`
- `folder.frame.dirty_rect_area_px`

- [x] **Step 4: Add input-to-paint measurement**

In `RedSalamander/FolderView.cpp`, capture a timestamp when scroll or keyboard navigation invalidates the view, then clear it after the next successful render. Emit `folder.frame.input_to_paint_us`.

Keep the timestamp optional so programmatic invalidations do not produce false input-latency rows.

Review cleanup moved timing wrappers into `RedSalamander/FolderView.Interaction.cpp` so `WndProc` remains a small message router.

- [x] **Step 5: Run focused FolderView scenario**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
```

Expected: run passes, archives evidence, and the artifact records all new `folder.frame.*` metrics.

- [x] **Step 6: Commit**

Run:

```powershell
git add RedSalamander/FolderView.h RedSalamander/FolderView.Rendering.cpp RedSalamander/FolderView.cpp RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp
git commit -m "feat(folderview): add frame stage metrics"
```

- [x] **Step 7: Address Task 6 review findings**

Keep Task 6 attribution-only semantics, but tighten evidence:

- track `folder.frame.dirty_rect_area_px` in `folderFrameMetricPresence` because Task 6 emits it,
- include per-metric counts in the artifact,
- scan `perf_metrics.jsonl` only from this scenario's starting file offset,
- replace unhandled `WM_VSCROLL` stress steps with handled wheel and keyboard navigation,
- keep `WndProc` as a small message router by moving input-to-paint timing wrappers into `FolderView.Interaction.cpp`.

## Task 7: Optimize FolderView Overlay And Idle Work Only If Metrics Justify It

**Files:**

- Modify if justified: `RedSalamander/FolderView.ErrorOverlay.cpp`
- Modify if justified: `RedSalamander/FolderView.Layout.cpp`
- Modify if justified: `RedSalamander/FolderView.Rendering.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

- [x] **Step 1: Compare FolderView candidate metrics with baseline**

Use the baseline and Task 6 candidate archives. Extract:

```powershell
$baseline = Read-Host 'Baseline perf_metrics.jsonl path'
$candidate = Read-Host 'Candidate perf_metrics.jsonl path'
Get-Content $baseline,$candidate | ConvertFrom-Json | Group-Object metric | Sort-Object Name | Select-Object Name,Count
```

Expected: the new metrics identify whether overlay animation, idle layout, present, or visible draw work dominates.

- [x] **Step 2: Gate overlay invalidation optimization**

Proceed only if `folder.frame.overlay_animation_count`, `render.incremental_search_effect_updates`, or `folder.frame.dirty_rect_area_px` shows avoidable full-window redraw during overlay animation.

If the gate passes, change overlay animation invalidation to invalidate only the overlay bounds plus blur/shadow padding. Preserve full invalidation for cases where the overlay changes size or theme.

- [x] **Step 3: Gate distant-state release optimization**

Proceed only if `folder.frame.visible_work_count`, text layout counts, or memory evidence shows retained distant text/icon state is a repeated cost.

If the gate passes, re-enable or tune `ReleaseDistantRenderingState()` with:

- a distance threshold based on viewport height,
- no release of visible or buffered visible items,
- no release during active thumbnail generation,
- a metric for released label/details/metadata layouts.

- [x] **Step 4: Run focused cases / no-op verification**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_scroll_stress --selftest-timeout-multiplier=4
```

Expected: all pass. Candidate metrics must show no p95/p99 regression in visible scrolling and no thumbnail guardrail regression.

Task 7 result: skipped the build/selftest trio because neither optimization gate passed and no code changed. Instead, verified the existing archives with targeted metric extraction and overlay scans, then ran `git diff --check`.

- [x] **Step 5: Commit**

Run:

```powershell
git add RedSalamander/FolderView.ErrorOverlay.cpp RedSalamander/FolderView.Layout.cpp RedSalamander/FolderView.Rendering.cpp RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp
git commit -m "perf(folderview): reduce measured frame work"
```

If neither gate passes, do not commit code for this task. Add the measured no-op decision to the evidence section instead.

## Task 8: Adopt Frame Runtime In ColorTextView Without Changing Renderer Modes

**Files:**

- Modify: `RedSalamanderMonitor/ColorTextView.h`
- Modify: `RedSalamanderMonitor/ColorTextView.cpp`
- Modify: `RedSalamanderMonitor/RedSalamanderMonitor.cpp`
- Modify: `Tests/MonitorTest/MonitorTest.cpp`

- [x] **Step 1: Add Monitor metric expectations**

Extend Monitor selftests so `--chrome-selftest --perf` records whether these metrics appeared:

```text
monitor.frame.total_us
monitor.frame.present_us
monitor.frame.append_to_visible_us
monitor.frame.tail_layout_us
monitor.frame.mode
monitor.etw.batch_drain_us
```

- [x] **Step 2: Verify missing metrics**

Run:

```powershell
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\build.ps1 -ProjectName MonitorTest -Configuration Debug
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: functional pass with missing new metrics recorded.

- [x] **Step 3: Add ColorTextView frame metrics**

In `ColorTextView::OnPaint()` and `ColorTextView::DrawScene()`, emit:

- `monitor.frame.total_us`,
- `monitor.frame.present_us`,
- `monitor.frame.tail_layout_us` around `RebuildTailLayout()` and AUTO_SCROLL layout refresh,
- `monitor.frame.scrollback_slice_us` around slice bitmap or fallback scroll-back rendering,
- `monitor.frame.mode` as AUTO_SCROLL or SCROLL_BACK value.

Keep existing `Present1`, slice bitmap, tail layout, and display-row mapping behavior unchanged.

- [x] **Step 4: Add append-to-visible measurement**

When ETW batch processing queues/appends lines in `RedSalamanderMonitor.cpp`, store a timestamp on the ColorTextView before invalidation. After the next successful paint that includes the appended tail in AUTO_SCROLL mode, emit `monitor.frame.append_to_visible_us`.

Do not emit append-to-visible in SCROLL_BACK mode because new lines intentionally do not move the visible viewport.

- [x] **Step 5: Run Monitor checks**

Run:

```powershell
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\build.ps1 -ProjectName MonitorTest -Configuration Debug
.\.build\x64\Debug\MonitorTest.exe --document-model-selftest
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: all pass, and Monitor archive contains `monitor.frame.*` metrics.

- [x] **Step 6: Commit**

Run:

```powershell
git add RedSalamanderMonitor/ColorTextView.h RedSalamanderMonitor/ColorTextView.cpp RedSalamanderMonitor/RedSalamanderMonitor.cpp Tests/MonitorTest/MonitorTest.cpp
git commit -m "feat(monitor): add ColorTextView frame metrics"
```

## Task 9: Tune Monitor Batch And Paint Scheduling Only If Metrics Justify It

**Files:**

- Modify if justified: `RedSalamanderMonitor/RedSalamanderMonitor.cpp`
- Modify if justified: `RedSalamanderMonitor/ColorTextView.cpp`
- Modify if justified: `RedSalamanderMonitor/ColorTextView.h`
- Test: `Tests/MonitorTest/MonitorTest.cpp`

- [x] **Step 1: Compare Monitor baseline and candidate metrics**

Extract:

- `monitor.etw.selftest_burst_drain_us`,
- `monitor.etw.batch_drain_us`,
- `monitor.frame.append_to_visible_us`,
- `monitor.frame.tail_layout_us`,
- `monitor.frame.total_us`,
- `monitor.frame.present_us`.

Expected: one of batch drain, tail layout, present, or append-to-visible latency is the dominant target.

- [x] **Step 2: Gate ETW drain budget optimization**

Proceed only if one UI-thread batch monopolizes the frame. If the gate passes:

- bound per-frame ETW drain work by time and count,
- repost remaining work with `PostMessagePayload(...)` or the existing safe payload pattern,
- keep in-order append semantics,
- emit queued, drained, and reposted counters.

- [x] **Step 3: Gate paint coalescing optimization**

Proceed only if repeated ETW batches cause multiple paints before one present. If the gate passes:

- coalesce invalidation during active AUTO_SCROLL batch processing,
- request one paint at the end of the drain slice,
- keep immediate paint behavior for user scrolling and selection.

- [x] **Step 4: Run Monitor checks**

Run:

```powershell
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\build.ps1 -ProjectName MonitorTest -Configuration Debug
.\.build\x64\Debug\MonitorTest.exe --document-model-selftest
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: all pass. Candidate metrics show no regression in append-to-visible p95/p99 and no loss of in-order display.

- [x] **Step 5: Commit**

Run:

```powershell
git add RedSalamanderMonitor/RedSalamanderMonitor.cpp RedSalamanderMonitor/ColorTextView.cpp RedSalamanderMonitor/ColorTextView.h Tests/MonitorTest/MonitorTest.cpp
git commit -m "perf(monitor): tune measured frame scheduling"
```

If neither gate passes, do not commit code for this task. Add the measured no-op decision to the evidence section instead.

Measured no-op commit path:

```powershell
git add Specs/Plans/WIP/DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md
git commit -m "docs(monitor): record scheduling gate decision"
```

## Task 10: Evaluate DXGI Present Policy Improvements

**Files:**

- Modify if justified: `Common/DxUi/DxUi.WindowHost.cpp`
- Modify if justified: `RedSalamander/FolderView.Rendering.cpp`
- Modify if justified: `RedSalamanderMonitor/ColorTextView.cpp`
- Test: `Tests/DxUiTests/DxUiTests.WindowHost.cpp`

- [x] **Step 1: Inventory current swap-chain capabilities**

Record for each surface:

- current swap effect,
- buffer count,
- `Present1` support,
- composition vs HWND swap chain,
- resize behavior,
- device-lost handling,
- whether `IDXGISwapChain2` is available.

- [x] **Step 2: Gate flip-discard migration**

Evaluate `DXGI_SWAP_EFFECT_FLIP_DISCARD` only after baseline metrics exist.

Proceed only if:

- it is compatible with the surface's Direct2D target creation,
- partial present semantics remain correct or the fallback is explicit,
- selftests show no visual or device-loss regression,
- same-machine present metrics improve or reduce variance.

- [x] **Step 3: Gate frame-latency waitable object**

Evaluate `DXGI_SWAP_CHAIN_FLAG_FRAME_LATENCY_WAITABLE_OBJECT` only as a capability-guarded experiment.

Proceed only if:

- supported for the actual HWND/composition swap chain in use,
- it does not deadlock WM_PAINT or resize paths,
- it reduces present pacing variance in PresentMon or `present_us` metrics,
- occlusion/minimize behavior is correct.

- [x] **Step 4: Run full protected render checks**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }
.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: all pass. If present policy changes are not supported or not faster, keep the current flip-sequential path and record the measured decision.

- [x] **Step 5: Commit supported improvements**

Run only if a present-policy change is measured and accepted:

```powershell
git add Common/DxUi/DxUi.WindowHost.cpp RedSalamander/FolderView.Rendering.cpp RedSalamanderMonitor/ColorTextView.cpp Tests/DxUiTests/DxUiTests.WindowHost.cpp
git commit -m "perf(render): tune dxgi present policy"
```

If neither present-policy gate passes, do not commit code for this task. Add the measured no-op decision to the evidence section instead.

Measured no-op commit path:

```powershell
git add Specs/Plans/WIP/DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md
git commit -m "docs(render): record present policy gate decision"
```

## Task 11: Optional Composition Animation Pilot

**Files:**

- Create if justified: `Common/DxUi/DxUi.CompositionPilot.h`
- Create if justified: `Common/DxUi/DxUi.CompositionPilot.cpp`
- Modify if justified: `Common/DxUi/DxUi.vcxproj`
- Do not create: `Common/DxUi/DxUi.vcxproj.filters` is absent in this repo.
- Modify if justified: `Common/DxUi/DxUi.WindowHost.cpp`
- Test: `Tests/DxUiTests/DxUiTests.Animation.cpp`

- [x] **Step 1: Gate the pilot**

Proceed only if DxUi frame metrics show CPU-thread animation jitter that cannot be fixed by the frame runtime alone.

Task 11 gate result: failed. Current evidence shows scheduler metrics and deterministic tests, but no measured unresolved production CPU-thread animation jitter on an allowed pilot surface.

Allowed pilot surfaces:

- tooltip fade,
- popup opacity/translate,
- dialog smoke opacity,
- lightweight overlay transforms.

Disallowed pilot surfaces:

- FolderView item rendering,
- ColorTextView log rendering,
- DirectWrite text layout,
- hit-testing,
- file list virtualization.

- [x] **Step 2: Add prototype behind a feature flag**

The pilot must be disabled by default and use WIL RAII wrappers for COM resources. It must not introduce global mutable device state.

Not accepted: no prototype was added because Step 1 failed.

- [x] **Step 3: Run visual and perf checks**

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=Animation
.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost
```

Expected: tests pass with the flag disabled. With the flag enabled manually, PresentMon or archived metrics must show smoother animation before enabling by default.

Not accepted: no flag exists. The required disabled-path build and DxUi suites passed, and the Animation suite wrote fresh gate metrics to `Specs\TestRuns\local_scratch\dxui_animation_task11_20260519.jsonl`.

- [x] **Step 4: Commit only if accepted**

Run only if the pilot is measured and accepted:

```powershell
git add Common/DxUi/DxUi.CompositionPilot.h Common/DxUi/DxUi.CompositionPilot.cpp Common/DxUi/DxUi.vcxproj Common/DxUi/DxUi.WindowHost.cpp Tests/DxUiTests/DxUiTests.Animation.cpp
git commit -m "perf(dxui): pilot compositor-backed overlay animation"
```

If not accepted, record the rejection reason and keep the code out of the tree.

Measured no-op commit path:

```powershell
git add Specs/Plans/WIP/DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md
git commit -m "docs(dxui): record composition pilot gate decision"
```

## Task 12: Full Validation And Spec Closeout

**Files:**

- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Modify: `Specs/UI/UI_DxUiWinUIDesign.md`
- Modify: `Specs/Core/Core_RedSalamanderMonitor.md`
- Modify if behavior changed: `Specs/UI/UI_FolderView.md`
- Move when complete: `Specs/Plans/WIP/DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md`

- [x] **Step 1: Run full focused validation**

Run:

```powershell
.\build.ps1 -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_sort_toggle_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_iconcache_contention --selftest-timeout-multiplier=4
.\.build\x64\Debug\MonitorTest.exe
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: all pass and archive or write metric evidence.

- [x] **Step 2: Run release validation for perf-sensitive evidence**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Release
.\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release
.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\.build\x64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf
```

Expected: Release builds pass with selftest hooks enabled by the project configuration. Archive paths are recorded.

- [x] **Step 3: Update authoritative specs**

Update:

- `Specs/UI/UI_DxUiWinUIDesign.md` with the frame runtime, animation clock, and debug render-stage contract.
- `Specs/Core/Core_RedSalamanderMonitor.md` with any ColorTextView frame metrics, append-to-visible measurement, and scheduling behavior that became durable.
- `Specs/UI/UI_FolderView.md` if FolderView invalidation, distant-state release, or frame metrics became durable.
- `Specs/Testing/Testing_TestCoverage.md` with exact run commands and archive paths.

- [x] **Step 4: Move plan to Done**

After all accepted implementation tasks and spec updates are complete:

```powershell
Move-Item -LiteralPath 'Specs\Plans\WIP\DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md' -Destination 'Specs\Plans\Done\DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md'
```

- [x] **Step 5: Commit closeout**

Run:

```powershell
git add Specs/Testing/Testing_TestCoverage.md Specs/UI/UI_DxUiWinUIDesign.md Specs/Core/Core_RedSalamanderMonitor.md Specs/UI/UI_FolderView.md Specs/Plans/Done/DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md
git commit -m "docs: close frame performance plan"
```

Task 12 closeout evidence (2026-05-19):

- Logs were written under `Specs/TestRuns/local_scratch/frame_perf_closeout_20260519_165526/`.
- Debug validation:
  - `.\build.ps1 -Configuration Debug` exited `0`; build log `.build/logs/msbuild-20260519_165527_622.log`; diagnostics `0 warning(s), 0 error(s)`.
  - `.\.build\x64\Debug\DxUiTests.exe` exited `0`; log `Specs/TestRuns/local_scratch/frame_perf_closeout_20260519_165526/02_debug_dxuitests.log`; output ended with `All DxUi tests passed`.
  - `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_165927/`; `1 passed / 0 failed / 0 skipped`.
  - `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_165933/`; `1 passed / 0 failed / 0 skipped`.
  - `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_sort_toggle_stress --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_165945/`; `1 passed / 0 failed / 0 skipped`.
  - `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_iconcache_contention --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_165950/`; `1 passed / 0 failed / 0 skipped`.
  - `.\.build\x64\Debug\MonitorTest.exe` exited `0`; log `Specs/TestRuns/local_scratch/frame_perf_closeout_20260519_165526/07_debug_monitortest.log`; no repo archive expected.
  - `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; build log `.build/logs/msbuild-20260519_165815_580.log`; diagnostics `0 warning(s), 0 error(s)`.
  - `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0` when run as a waited visible GUI process; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_170035/`; required metrics present: `monitor.frame.total_us`, `monitor.frame.present_us`, `monitor.frame.append_to_visible_us`, `monitor.frame.tail_layout_us`, `monitor.frame.mode`, and `monitor.etw.batch_drain_us`.
- Release validation:
  - `.\build.ps1 -ProjectName RedSalamander -Configuration Release` exited `0`; build log `.build/logs/msbuild-20260519_170103_776.log`; diagnostics `0 warning(s), 0 error(s)`.
  - The exact normal Release command `.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4` exited `2` with no archive because normal Release `RedSalamander.exe` omits `ENABLE_TESTS`.
  - Explicit test-enabled Release perf evidence was collected with `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamander -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }`; build exited `0`, log `.build/logs/msbuild-20260519_170412_660.log`, diagnostics `1 warning(s), 0 error(s)` (`C4883` in `FolderWindow.FileOperations.SelfTest.cpp`). The rerun `.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4` exited `0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_170631/`.
  - `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; build log `.build/logs/msbuild-20260519_170241_261.log`; diagnostics `0 warning(s), 0 error(s)`.
  - `.\.build\x64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `1`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_170251/`; blocker: `Monitor DxUI chrome selftest requires ENABLE_TESTS.` This red evidence showed that Release Monitor did not wire `$(RSBuildTestDefinitions)` into Release `ClCompile` definitions.
  - Follow-up fix `fix(monitor): enable release selftest opt-in` updated `RedSalamanderMonitor.vcxproj` Release x64/ARM64 `ClCompile` `PreprocessorDefinitions` to include `$(RSBuildTestDefinitions)`. The rerun `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; build log `.build/logs/msbuild-20260519_171244_915.log`; wrapper log `Specs/TestRuns/local_scratch/frame_perf_release_monitor_fix_20260519_171500/01_release_monitor_test_enabled_build.log`; diagnostics `0 warning(s), 0 error(s)`.
  - Final Release Monitor chrome validation `.\.build\x64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0`; wrapper log `Specs/TestRuns/local_scratch/frame_perf_release_monitor_fix_20260519_171500/02_release_monitor_chrome_perf.log`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_171308/`; `results.json` reports `status=passed` and `monitorFrameMetricPresence.allPresent=true` with required metric counts: `monitor.frame.total_us` 5, `monitor.frame.present_us` 5, `monitor.frame.append_to_visible_us` 1, `monitor.frame.tail_layout_us` 15, `monitor.frame.mode` 5, and `monitor.etw.batch_drain_us` 2.
  - Controller rerun after the same fix: `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }` exited `0`; build log `.build/logs/msbuild-20260519_171503_045.log`; diagnostics `0 warning(s), 0 error(s)`. `.\.build\x64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf` exited `0`; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_171520/`; `results.json` reports `status=passed` and `monitorFrameMetricPresence.allPresent=true`.
- Authoritative spec updates landed in `Specs/UI/UI_DxUiWinUIDesign.md`, `Specs/Core/Core_RedSalamanderMonitor.md`, `Specs/UI/UI_FolderView.md`, and `Specs/Testing/Testing_TestCoverage.md`.
- Closeout move target: `Specs/Plans/Done/DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md`.
- Closeout commit message: `docs: close frame performance plan`.

## Acceptance Matrix

| Area | Required before enabling by default |
| --- | --- |
| DxUi frame runtime | `DxUiTests --suite=Animation` and `DxUiTests --suite=WindowHost` pass with metrics |
| DxUi render-stage guard | debug test proves layout mutation during render is detected |
| FolderView metrics | large-folder, scroll/render, sort-toggle, and thumbnail guardrails pass |
| FolderView optimization | same-machine evidence shows no visible-work or p95/p99 regression |
| Monitor metrics | `MonitorTest` and `RedSalamanderMonitor.exe --chrome-selftest --perf` pass |
| Monitor optimization | append order preserved; append-to-visible and batch-drain metrics do not regress |
| Present policy | capability-gated; device-loss, resize, occlusion, and dirty rect behavior pass |
| Composition pilot | disabled by default until measured; never used for text/list virtualization |
| Documentation | authoritative specs updated before moving this plan to Done |

## Risk Register

| Risk | Mitigation |
| --- | --- |
| Shared runtime becomes a hidden global framework | Keep it value-type and host-owned; no singleton device or swap-chain state. |
| New metrics add noise or overhead | Emit only per-frame aggregates; avoid per-item metric rows in hot loops. |
| Animation scheduler breaks existing controls | Preserve current subscription API first; change internals behind tests. |
| FolderView loses existing optimizations | Adopt metrics first; optimize only gated by FolderView archives. |
| ColorTextView AUTO_SCROLL latency regresses | Preserve two-mode renderer and measure append-to-visible before scheduling changes. |
| Waitable swap-chain path deadlocks WM_PAINT | Capability-gate and keep current flip-sequential fallback. |
| Composition introduces airspace/text problems | Pilot only overlays; keep DirectWrite text and list virtualization in existing renderers. |

## Baseline Evidence

### 2026-05-19 Worker 1 Protected Baseline

Initial tree state:

- `git status --short`: exit `0`, clean output before baseline. No unrelated tracked changes were present.
- Same-machine profile: `4cb089111a23`.
- Branch during archived selftests: `dxui-frame-performance` as recorded in Command perf rows.
- Commit recorded by Command perf rows: `f6602f45432ea6767bfad9a46b8f99fad820ed83`.

Builds:

| Command | Exit | Summary | Log |
| --- | ---: | --- | --- |
| `./build.ps1 -ProjectName DxUiTests -Configuration Debug` | 0 | Passed, 0 warnings, 0 errors | `.build/logs/msbuild-20260519_132118_208.log` |
| `./build.ps1 -ProjectName RedSalamander -Configuration Debug` | 0 | Passed, 0 warnings, 0 errors | `.build/logs/msbuild-20260519_132128_099.log` |
| `./build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug` | 0 | Passed, 0 warnings, 0 errors | `.build/logs/msbuild-20260519_132314_311.log` |
| `./build.ps1 -ProjectName MonitorTest -Configuration Debug` | 0 | Passed, 0 warnings, 0 errors | `.build/logs/msbuild-20260519_132327_191.log` |

DxUi baselines:

| Command | Exit | Output summary | Evidence | Metric keys found |
| --- | ---: | --- | --- | --- |
| `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost --perf-jsonl=Specs\TestRuns\local_scratch\dxui_windowhost_frame_runtime_baseline_20260519.jsonl` | 0 | `All DxUi tests passed`; `WindowHost` suite completed. | `Specs/TestRuns/local_scratch/dxui_windowhost_frame_runtime_baseline_20260519.jsonl` (133 rows) | `DxUI::FocusChange`, `DxUI::Paint`, `dxui.focus.tab_navigation`, `dxui.slider.paint`, `dxui.tabcontrol.paint`, `dxui.tabcontrol.sync_layout_us`, `dxui.tabcontrol.update_visible_pages_us`, `dxui.textinput.activate_us`, `dxui.textinput.hit_test_us`, `dxui.textinput.key_to_state_us`, `dxui.windowhost.dpi_change_us` |
| `.\.build\x64\Debug\DxUiTests.exe --suite=Animation --perf-jsonl=Specs\TestRuns\local_scratch\dxui_animation_frame_runtime_baseline_20260519.jsonl` | 0 | `All DxUi tests passed`; `Animation` suite completed. | `Specs/TestRuns/local_scratch/dxui_animation_frame_runtime_baseline_20260519.jsonl` (12 rows) | `DxUI::FocusChange`, `dxui.textinput.activate_us` |

FolderView baselines:

| Command | Exit | Pass/fail count | Archive | Metric keys found |
| --- | ---: | --- | --- | --- |
| `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4` | 0 | 1 passed / 0 failed / 0 skipped | `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_132404` | 3296 rows; includes `folder.selftest.render_warmup_us`, `render.frame_us`, `render.present_us`, `render.begin_to_enddraw_us`, `render.draw_item_us`, `render.dirty_rect_area_px`, `render.items_drawn`, `render.items_considered`, `render.layout_items_us`, `render.incremental_search_effect_updates`, `icons.*`, `iconcache.*`, `FolderView.ApplyCurrentSort`, `FolderView.ExecuteEnumeration.*`, and startup metrics |
| `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4` | 0 | 1 passed / 0 failed / 0 skipped | `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_132547` | 5711 rows; includes `folder.scroll_frame_count`, `folder.scroll_input_to_paint_us`, `folder.scroll_visible_item_count`, `folder.selftest.render_warmup_us`, `render.*`, `icons.*`, `iconcache.*`, `FolderView.ApplyCurrentSort`, `FolderView.ExecuteEnumeration.*`, and startup metrics |
| `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_sort_toggle_stress --selftest-timeout-multiplier=4` | 0 | 1 passed / 0 failed / 0 skipped | `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_132646` | 3093 rows; includes `folder.sort_toggle_us`, `folder.selftest.render_warmup_us`, `render.*`, `icons.*`, `iconcache.*`, `FolderView.ApplyCurrentSort`, `FolderView.ExecuteEnumeration.*`, and startup metrics |
| `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_iconcache_contention --selftest-timeout-multiplier=4` | 0 | 1 passed / 0 failed / 0 skipped | `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_132411` | 3039 rows; includes `folder.iconcache_contention_cycle_us`, `folder.selftest.render_warmup_us`, `render.*`, `icons.*`, `iconcache.*`, `FolderView.ApplyCurrentSort`, `FolderView.ExecuteEnumeration.*`, and startup metrics |

FolderView caveat:

- The first direct attempts for `folderView_perf_scroll_render_stress` and `folderView_perf_sort_toggle_stress` exited `0` but only later inspection showed no archive for those two runs. They were rerun before documenting this baseline and produced the archives above. No production or test code was changed.

Monitor baselines:

| Command | Exit | Output summary | Evidence | Metric keys found |
| --- | ---: | --- | --- | --- |
| `.\.build\x64\Debug\MonitorTest.exe --diagnostics-gate-selftest` | 0 | Quiet stdout; command completed successfully. | No repo archive expected for this MonitorTest entrypoint. | Not applicable. |
| `.\.build\x64\Debug\MonitorTest.exe --scrollbar-model-selftest` | 0 | Quiet stdout; command completed successfully. | No repo archive expected for this MonitorTest entrypoint. | Not applicable. |
| `.\.build\x64\Debug\MonitorTest.exe --document-model-selftest` | 0 | Quiet stdout; command completed successfully. | No repo archive expected for this MonitorTest entrypoint. | Not applicable. |
| `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf` | 0 | [blocked] Archived `results.json` reports `status: failed` with summary `Monitor DxUI chrome selftest requires ENABLE_TESTS.` | `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_132437` | `perf_metrics.jsonl` has 2 rows, both `DxUI::Paint`; expected passing Monitor chrome baseline metrics are missing because the selftest did not run past the `ENABLE_TESTS` gate. |

Pass/fail summary:

- Protected builds: 4 passed / 0 failed.
- DxUi test commands: 2 passed / 0 failed.
- FolderView command selftests: 4 passed / 0 failed / 0 skipped after rerun of the two initially unarchived cases.
- MonitorTest commands: 3 exited `0`.
- Monitor chrome perf command: exit `0`, but archived selftest status failed; record as [blocked] for passing Monitor chrome baseline evidence. Missing data: passing `results.json`, non-failing `trace.txt`, and Monitor chrome metric rows beyond the two `DxUI::Paint` samples.

Commit/ignore caveat:

- `Specs/TestRuns/...` and `Specs/TestRuns/local_scratch/...` artifacts are intentionally ignored by the repo. They were recorded here by path and were not force-added.

Placeholder scan:

- Worker 1 ran a placeholder scan after editing the plan and reported no stale baseline placeholders.
- Recheck command: `rg -n "T[B]D|T[O]DO|i[m]plement later|f[i]ll in details|a[p]propriate error handling|S[i]milar to Task|<M[a]chineHash>|<R[u]nId>|<B[a]selineRun>|<C[a]andidateRun>" Specs\Plans\WIP\DxUi_FolderView_Monitor_FramePerformancePlan_2026-05-19.md`
- Recheck result: exit `1`, no matches.
