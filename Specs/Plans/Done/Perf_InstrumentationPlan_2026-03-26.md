# Performance Instrumentation Implementation Plan

Last updated: 2026-04-02

Status: Done

References:

- `RedSalamander/FolderView.Rendering.cpp`
- `RedSalamander/FolderView.Icons.cpp`
- `RedSalamander/IconCache.cpp`
- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `Common/LocalSearchIndexCore.cpp`
- `RedSalamander/FindFilesWindow.cpp`
- `README.md`
- `Specs/Testing/Testing_SelfTests.md`
- `Specs/TestRuns/README.md`

## Progress Checklist

Update this checklist as work progresses.

- [x] Add shared perf helper APIs and scoped timer/value helpers.
- [x] Define metric naming conventions and output format contract.
- [x] Add folder rendering instrumentation in `FolderView.Rendering.cpp`.
- [x] Add icon queue/UI conversion instrumentation in `FolderView.Icons.cpp`.
- [x] Add icon cache lock/shell/conversion instrumentation in `IconCache.cpp`.
- [x] Add file-operations scheduler instrumentation in `FolderWindow.FileOperations.State.cpp`.
- [x] Add file-operations queue/progress/pre-calc instrumentation in `FolderWindow.FileOperations.State.cpp`.
- [x] Extend `FileOps.Operation` rollups with new fields.
- [x] Add search backend instrumentation in `LocalSearchIndexCore.cpp`.
- [x] Add search callback bridge instrumentation in `FindFilesWindow.cpp`.
- [x] Add search UI handler instrumentation in `FindFilesWindow.cpp`.
- [x] Add payload enqueue timestamps for search UI latency metrics.
- [x] Build x64 Debug and fix any warnings/regressions in touched files.
- [x] Run focused self-tests for search, windows/modeless flows, and command infrastructure.
- [x] Run folder-pane metric scenarios and archive the baseline report.
- [x] Run file-operations metric scenarios and archive the baseline report.
- [x] Run search metric scenarios and archive the baseline report.
- [x] Compare two consecutive metric runs to verify stability and overhead.
- [x] Document interpretation guidance and next optimization decisions.

## Current Validation Status

Validated on 2026-03-26 after fixing the Commands speed-limit prompt harness and the FileOps aggregate-run family transition/setup issue.

Passing validation runs:

- `.\build.ps1 -ProjectName RedSalamander`
- `RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface`
- `RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=cmd_pane_find_dialog_search_ops`
- `RedSalamander.exe --compare-selftest --selftest-fail-fast --selftest-case=local_search_callback_contract`
- `RedSalamander.exe --fileops-selftest --selftest-fail-fast`

Archived passing runs:

- Commands speed-limit prompt: `Specs/TestRuns/7d3a1247382a/Commands/2026-03-26_155717`
- Commands find dialog search ops: `Specs/TestRuns/7d3a1247382a/Commands/2026-03-26_160517`
- Compare/local-search callback contract: `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-03-26_160515`
- Compare/local-search callback contract rerun: `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-03-26_160904`
- FileOps aggregate suite with perf JSONL: `Specs/TestRuns/7d3a1247382a/FileOps/2026-03-26_160446`
- FileOps aggregate suite rerun with perf JSONL: `Specs/TestRuns/7d3a1247382a/FileOps/2026-03-26_174352`
- Commands/folder large-folder baseline with archived `render.*` and `icons.*`: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-02_182303`

Non-blocking follow-up:

- optional dedicated runner automation for expected fatal-error terminal states during stress/perf scenarios

Additional folder-view validation completed on 2026-03-26:

- Commands/folder watermark render smoke: `Specs/TestRuns/7d3a1247382a/Commands/2026-03-26_175258`
- Commands/folder empty-state smoke: `Specs/TestRuns/7d3a1247382a/Commands/2026-03-26_175304`
- Commands/folder watermark rerun: `Specs/TestRuns/7d3a1247382a/Commands/2026-03-26_175428`

These runs pass and archive perf JSONL correctly, but they mostly emit startup and `iconcache.*` metrics. They were useful initial smoke coverage, not the final folder-pane baseline.

Additional same-machine validation completed on 2026-03-27:

- Commands/find dialog search ops with sorted-result instrumentation: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_140722`
- Commands/find dialog search ops with timer-based deferred refresh coalescing: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_141541`
- FileOps aggregate suite with JSONL capture fix and path-update counters: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-27_141221`
- FileOps aggregate suite with throttled UI-facing progress path updates: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-27_144845`
- FileOps aggregate suite with item-completed path-copy dedup against last callback path: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-27_145536`

These 2026-03-27 runs are the best current basis for optimization decisions because they are same-machine, same-config, same-suite comparisons and they explicitly exercise the newly added `find.ui.deferred_refresh_*` and `FileOps.*PathUpdate*` metric families.

## Current Run-To-Run Notes

Search/compare stability review is now available from the consecutive runs:

- baseline: `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-03-26_160515`
- candidate: `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-03-26_160904`

Observed with `Tools/Show-PerfRuns.ps1`:

- `App.Startup.Metric.TimeToFirstPanePopulated` regressed from about `930917us` to `1170419us` (`+25.7%`)
- `iconcache.shgetfileinfo_us` `p95` regressed from about `12312us` to `13857us` (`+12.5%`)
- `iconcache.lock_wait_us` improved sharply, from `3329` to `3` at `p95`, but the sample counts changed materially (`5` to `27`) so that win is not yet a stable conclusion
- `iconcache.lock_hold_us` increased from `17us` to `58us` at `p95`, again with a much larger sample count on the candidate run

FileOps stability review is now available from the consecutive runs:

- baseline: `Specs/TestRuns/7d3a1247382a/FileOps/2026-03-26_160446`
- candidate: `Specs/TestRuns/7d3a1247382a/FileOps/2026-03-26_174352`

Observed with `Tools/Show-PerfRuns.ps1`:

- `FileOps.Operation` `p95` improved from about `30343000us` to `28344000us` (`-6.6%`)
- `FileOps.Scheduler.WaitForWorkUs` `p95` improved from about `119985562` to `107563595` (`-10.4%`)
- `FileOps.Progress.LockWaitUs` `p95` improved from about `578us` to `239us` (`-58.7%`)
- `FileOps.PreCalc.TotalUs` `p95` regressed from about `33735` to `190689` on the aggregate metric record and `FileOps.PreCalc` moved from about `196110us` to `363019us`, so pre-calc cost is the biggest remaining volatility source
- startup metrics also regressed by about `10-11%` between the two FileOps runs, which means startup noise is still leaking into end-to-end comparisons

Interpretation guidance at the current state:

- treat startup metrics as the most trustworthy regression signal from the compare/local-search scenario because they are single-run, same-scenario values with direct user-facing meaning
- treat icon-cache deltas as provisional until the same scenario is repeated with closer event counts or a more icon-heavy dedicated pane scenario
- treat the FileOps aggregate suite as repeatable enough for run-to-run comparison after the late-family setup reuse fixes, but keep the interpretation focused on broad deltas rather than single-metric micro-claims because pre-calc timings still vary materially
- the next highest-value instrumentation follow-up is a dedicated folder-pane scenario baseline so render/icon metrics can be reviewed without FileOps/startup noise

Same-machine search optimization review is now available from the consecutive runs:

- baseline: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_140722`
- candidate: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_141541`

Observed with `Tools/Show-PerfRuns.ps1`:

- `find.ui.deferred_refresh_post_count` dropped from `2` to `1`
- `find.ui.deferred_refresh_coalesced_count` moved from `0` to `1`, which confirms the timer-based deferral is actually merging adjacent sorted batches
- `find.ui.results_full_rebuild_count` dropped from `5` to `4` (`-20%`)
- `find.ui.results_to_visible_latency_ms` average dropped from about `12104us` to `6225us` (`-48.6%`) on the same scenario
- `find.ui.results_to_visible_latency_ms` max remained noisy (`15470us` to `17486us`), so the main confidence signal is the reduced rebuild / refresh count rather than the single-run latency tail

Interpretation:

- keep the timer-based sorted refresh coalescing; it is the clearest same-machine search win so far
- use the new sorted search scenario when reviewing future `find.ui.sort_results_ms` or `find.ui.deferred_refresh_*` changes because the older `2026-03-26` search run did not exercise those paths at all

Same-machine FileOps optimization review is now available from the consecutive runs:

- baseline: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-27_141221`
- candidate 1: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-27_144845`
- candidate 2: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-27_145536`

Observed with `Tools/Show-PerfRuns.ps1`:

- candidate 1 reduced `FileOps.Progress.PathUpdateBytes` from `1383024` to `717782` (`-48.1%`) and `FileOps.Progress.LockHoldUs` total from `18184us` to `16116us` (`-11.4%`)
- candidate 1 introduced `FileOps.Progress.PathUpdateThrottledCount` with total `3304`, confirming the UI-facing progress path throttle is active
- candidate 1 also shifted path churn into item completion, with `FileOps.ItemCompleted.PathUpdateBytes` moving from `6914` to `50356`; that was treated as a real follow-up issue rather than ignored
- candidate 2 removed that catch-up cost: `FileOps.ItemCompleted.PathUpdateBytes` dropped from `50356` to `0`
- candidate 2 improved `FileOps.Progress.PathUpdateBytes` again, from `717782` to `509382` (`-29.0%` vs candidate 1, `-63.2%` vs baseline)
- combined `FileOps.Progress.PathUpdateBytes + FileOps.ItemCompleted.PathUpdateBytes` moved from `1389938` on the baseline to `509382` on candidate 2 (`-63.4%`)
- `FileOps.Progress.LockHoldUs` improved again on candidate 2, from `16116us` to `15051us` (`-6.6%`)

Interpretation:

- keep both FileOps follow-up optimizations: throttled UI-facing progress path updates and item-completed dedup against the last exact callback path
- the current FileOps path-churn metrics are now actionable enough to use for future micro-optimizations because the same-machine runs show large, repeated reductions rather than one-off noise

Current stable baseline set to use going forward:

- search baseline: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_140722`
- search candidate/current best: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_141541`
- FileOps baseline: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-27_141221`
- FileOps candidate/current best: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-27_145536`
- folder-pane baseline: `Specs/TestRuns/7d3a1247382a/Commands/2026-04-02_182303`

Additional folder-pane baseline work completed on 2026-03-27:

- added a dedicated Commands selftest case: `folderView_perf_large_folder_baseline`
- archived same-machine folder baseline attempts:
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_150721`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_150843`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_150943`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_153032`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_153447`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_154032`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_154242`

Observed:

- these runs now emit repeatable folder enumeration/sort metrics for the synthetic large-folder scenario:
  - `FolderView.ExecuteEnumeration.BuildItems`
  - `FolderView.ExecuteEnumeration.IconIndex.QueryExtensions`
  - `FolderView.ApplyCurrentSort`
- the scenario reaches the folder-pane codepaths more reliably than the older empty-state/watermark smokes and is the right baseline shape for enumeration-heavy regressions
- at that point the render/icon capture remained unresolved, so the scenario was kept as provisional rather than final

Folder-pane baseline closure on 2026-04-02:

- reran the dedicated Commands selftest case on the current workspace build:
  - `RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=folderView_perf_large_folder_baseline`
- archived passing baseline:
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-04-02_182303`
- the archived `perf/perf_metrics.jsonl` now contains the full folder-pane families the checklist was waiting on:
  - enumeration/sort: `FolderView.ExecuteEnumeration.BuildItems`, `FolderView.ExecuteEnumeration.IconIndex.QueryExtensions`, `FolderView.ApplyCurrentSort`
  - render: `render.draw_item_us`, `render.begin_to_enddraw_us`, `render.present_us`, `render.frame_us`, `render.items_drawn`
  - icons: `icons.queue_build_us`, `icons.queue_wait_to_dequeue_us`, `icons.extract_us`, `icons.ui_convert_us`, `icons.ui_apply_count`
- the archived custom artifact `perf/folderView_perf_large_folder_baseline_metrics.json` also records the expected warmup activity:
  - `warmRenderInvocations = 5`
  - `renderCalls = 21`
  - `queueIconLoadingCalls = 7`
- that run closes the original render/icon baseline gap; the old 2026-03-27 conclusion about a partial Commands capture is no longer current

Interpretation update after investigating `FileOps.ItemCompleted.*Us`:

- the suspicious exact powers-of-two values seen during review were not bogus duration samples; they were the `completedBytes` payload carried in `value0` on the raw JSONL rows
- the `FileOps.ItemCompleted.CallbackUs` and `FileOps.ItemCompleted.LockHoldUs` durations in raw JSONL are small and plausible; they are not the current highest-value FileOps optimization target
- the biggest remaining FileOps volatility source is still pre-calc cost, not item-completed callback timing

Next search optimization target:

- the sorted search scenario shows `find.ui.refresh_results_view_ms` dominating `find.ui.sort_results_ms` on the 50-result deferred-refresh batch (`1337us` vs `376us`, with `find.ui.deferred_refresh_handler_ms = 1750us` in `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_141541`)
- the next search optimization should therefore focus on reducing results-view rebuild/refresh cost before chasing sort micro-optimizations

Same-machine follow-up search refresh optimization review on 2026-03-27:

- baseline/current pre-grid-common-case pass: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_162805`
- candidate/current best after grid common-case fast path: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_162913`

Observed:

- the new `find.ui.results_fast_refresh_count` metric confirms the unsorted Find Files fast path is active in both runs (`3` fast refreshes, `4` full grid notifications)
- the first narrow Find Files-only fast path was not enough on its own; it engaged, but the `2026-03-27_162805` run still carried unsorted refreshes around `382us avg`
- the follow-up optimization moved the common-case reduction into `DxUi::Grid::NotifyDataChanged()` and `ClampScrollOffsets()` for the no-group / no-selection path used by the results grid
- that second pass improved same-machine search UI costs in `2026-03-27_162913`:
  - unsorted `find.ui.refresh_results_view_ms`: `382us avg -> 346us avg` (`-9.4%`), max `482us -> 350us`
  - sorted `find.ui.refresh_results_view_ms`: `1335us avg -> 1242.5us avg` (`-6.9%`), max `1606us -> 1509us`
  - `find.ui.deferred_refresh_handler_ms`: `1985us -> 1860us` (`-6.3%`)
- `find.ui.sort_results_ms` stayed roughly flat (`112.5us avg -> 119.75us avg`), which reinforces the earlier conclusion that refresh/reconcile work, not sort itself, is the better optimization target

Interpretation:

- keep the grid common-case fast path; it gives a modest but repeatable same-machine win on both unsorted and deferred sorted Find Files refreshes
- use `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_162913` as the new current-best search baseline when evaluating future `find.ui.refresh_results_view_ms` work

Search refresh instrumentation follow-up on 2026-03-27:

- initial substep capture: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_164259`
- cleaned core-vs-instrumentation capture: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_164613`

Observed:

- raw `find.ui.refresh_results_view_ms` became too noisy for fine-grained decision-making once multiple substep emits were added inside the refresh path
- the corrected `find.ui.refresh_core_us` metric in `2026-03-27_164613` is the better guide for the next optimization pass:
  - overall refresh core: `317us avg`, `798us max`
  - fast unsorted refreshes: `11us` and `15us` in the two visible same-run samples
  - sorted/full-rebuild refresh core samples: `364us`, `461us`, and one `798us` outlier
- the measured synchronous substeps inside the cleaned run are all small:
  - `find.ui.refresh_grid_update_us`: `~10us avg`
  - `find.ui.refresh_action_buttons_us`: `~14us avg`
  - `find.ui.refresh_host_invalidate_us`: `~9us avg`
  - `find.ui.sort_results_ms`: `~113us avg`
- interpretation: the remaining user-visible cost is no longer in the simple refresh plumbing that was just optimized; future search work should target the deferred sorted-refresh path around the 50-result batch and any downstream paint/render cost rather than more micro-tuning in `RefreshResultsView()`

Search progress coalescing follow-up on 2026-03-27:

- before: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_172945`
- after: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_173156`

Observed:

- the latest same-machine Commands run showed a clear queueing pattern: only `5` result messages were reaching the UI, but `52` progress messages were still being handled, and `find.ui.results_to_visible_latency_ms` averaged `14286.8us`
- a UI-thread head-of-queue progress coalescing pass reduced `find.ui.progress_message_count` from `52` to `20`
- the new instrumentation confirms real coalescing work:
  - `find.ui.progress_messages_coalesced_count`: `10`
  - `find.ui.progress_messages_drained`: `32 total`
- that materially improved end-to-end message latency in the same scenario:
  - `find.ui.progress_to_visible_latency_ms`: `11647.31us avg -> 4670.95us avg` (`-59.9%`)
  - `find.ui.results_to_visible_latency_ms`: `14286.8us avg -> 6349.8us avg` (`-55.6%`)
- refresh cost itself stayed roughly flat, which is the expected and useful outcome:
  - `find.ui.refresh_results_view_ms`: `2619.6us avg -> 2554.8us avg`
  - `find.ui.refresh_core_us`: `255.8us avg -> 218.4us avg`

Interpretation:

- keep the progress coalescing path; it attacks queue churn rather than refresh micro-cost, and the measured same-machine improvement is much larger than the recent refresh-only wins
- the next search optimization target should stay in queueing / downstream presentation work, not in sort internals

Search render probe follow-up on 2026-03-27:

- latest passing selftest evidence is currently in `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\commands\trace.txt` rather than the repo archive copy

Observed:

- a direct render probe was added to `cmd_pane_find_dialog_search_ops` so the Commands selftest now records `dxRenderDelta` and `gridPaintDelta` per major Find phase in the archived Commands trace
- the latest passing run currently reports:
  - `Find`: `dxRenderDelta=0`, `gridPaintDelta=0`
  - `Append`: `dxRenderDelta=0`, `gridPaintDelta=0`
  - `sorted Find`: `dxRenderDelta=0`, `gridPaintDelta=0`
- each probe also waited roughly `3s`, which means the blocker is no longer “`dxui.grid.paint_us` is missing from JSONL even though paint definitely happened”

Interpretation:

- for this selftest scenario, the better conclusion is that the current render-count probes are not observing a render transition at all, so paint-side optimization remains blocked on a scenario/probe mismatch rather than a plain JSONL sink bug
- keep the non-blocking trace probe in place so future work can validate any stronger render-driving harness changes without destabilizing the Commands suite

CompareDirectories progress coalescing follow-up on 2026-03-27:

- scan/content progress handlers now coalesce head-of-queue payloads and emit:
  - `compare.ui.scan_progress_message_count`
  - `compare.ui.scan_progress_messages_coalesced_count`
  - `compare.ui.scan_progress_messages_drained`
  - `compare.ui.scan_progress_to_visible_latency_ms`
  - `compare.ui.content_progress_message_count`
  - `compare.ui.content_progress_messages_coalesced_count`
  - `compare.ui.content_progress_messages_drained`
  - `compare.ui.content_progress_to_visible_latency_ms`
- the unrelated Compare engine selftest blocker was partially unblocked in two follow-up fixes:
  - `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-03-27_180231`: `sqlite_index_store_bootstrap_creates_schema` now passes after persisting `auto_vacuum=INCREMENTAL`
  - `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-03-27_180402`: `sqlite_index_store_automatic_checkpoint_truncates_wal` now passes after a final post-metadata truncate checkpoint
- those engine-suite reruns still do not verify `compare.ui.*`, because they do not exercise the `CompareDirectoriesWindow` UI path
- end-to-end UI verification is now covered by the new Commands scenario `cmd_compare_directories_progress_perf`:
  - failing proof run with archived metrics: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_182330`
  - passing current-best run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_182411`
- the passing Commands run confirms the real Compare window path archives the new metrics:
  - `compare.ui.scan_progress_message_count`
  - `compare.ui.scan_progress_messages_coalesced_count`
  - `compare.ui.scan_progress_messages_drained`
  - `compare.ui.scan_progress_to_visible_latency_ms`
  - `compare.ui.content_progress_message_count`
  - `compare.ui.content_progress_messages_coalesced_count`
  - `compare.ui.content_progress_messages_drained`
  - `compare.ui.content_progress_to_visible_latency_ms`
- the archived values from `2026-03-27_182411` show meaningful queue coalescing, not just raw emission:
  - scan progress drained at least one queued payload (`compare.ui.scan_progress_messages_drained = 1`)
  - content progress drained large queued bursts (`compare.ui.content_progress_messages_drained` samples include `120`, `120`, `61`, `30`, `28`)
  - scan/content progress latencies were captured in the `49ms-79ms` range for this synthetic UI scenario

Interpretation:

- keep the CompareDirectories coalescing code; it is the same low-risk queue-churn reduction pattern that already paid off in Find Files
- use `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_182411` as the current reference run for future Compare window progress-path tuning

CompareDirectories progress UI follow-up on 2026-03-27:

- first banner/task-card reduction candidate:
  - cold candidate run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_185305`
  - warmed rerun / current-best keep candidate: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_185337`
- rejected startup task-card deferral candidate:
  - rejected run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_185418`

Observed:

- the first pass added banner/task-card proof metrics:
  - `compare.ui.progress_controls_update_count`
  - `compare.ui.progress_controls_text_applied_count`
  - `compare.ui.progress_controls_show_count`
  - `compare.ui.progress_controls_hide_count`
  - `compare.ui.progress_controls_skipped_count`
  - `compare.ui.progress_controls_progressbar_invalidate_count`
  - `compare.ui.progress_controls_update_us`
  - `compare.ui.taskcard_update_count`
  - `compare.ui.taskcard_update_throttled_count`
  - `compare.ui.taskcard_update_us`
- the code change behind `2026-03-27_185337` caches the last visible banner state and banner text, avoids re-showing / re-hiding controls when state is unchanged, avoids redundant progress-bar invalidation while the spinner timer is already active, and throttles non-finished informational-task updates to roughly banner cadence
- the warmed same-machine rerun `2026-03-27_185337` improved end-to-end Compare UI latency versus the earlier reference run `2026-03-27_182411`:
  - `compare.ui.scan_progress_to_visible_latency_ms`: `71806us avg -> 66056us avg` (`-8.0%`)
  - `compare.ui.content_progress_to_visible_latency_ms`: `61569us avg -> 56881us avg` (`-7.6%`)
- the new metrics confirm the optimization is doing real work in `2026-03-27_185337`:
  - only `7` of `9` banner updates changed text (`compare.ui.progress_controls_text_applied_count = 7`)
  - `2` banner updates were full no-op skips (`compare.ui.progress_controls_skipped_count = 2`)
  - the progress bar was explicitly invalidated only once on show (`compare.ui.progress_controls_progressbar_invalidate_count = 1`)
  - only `3` informational-task updates were actually applied while `6` were throttled (`compare.ui.taskcard_update_count = 3`, `compare.ui.taskcard_update_throttled_count = 6`)
- a follow-up attempt deferred the initial informational-task creation until first progress / completion; that was rejected after `2026-03-27_185418`
- although the rejected pass reduced some internal update cost (`compare.ui.progress_controls_update_us` average fell, and startup banner work stayed cheap), it clearly regressed the user-facing metric we care about:
  - `compare.ui.scan_progress_to_visible_latency_ms`: `66056us avg -> 166991us avg`
  - `compare.ui.content_progress_to_visible_latency_ms`: `56881us avg -> 151618us avg`

Interpretation:

- keep the banner-state caching and non-finished task-card throttling from `2026-03-27_185337`
- explicitly reject the startup task-card deferral candidate from `2026-03-27_185418`; it lowers some internal work but makes progress noticeably less responsive
- use `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_185337` as the new current-best Compare window progress-path baseline

CompareDirectories task-card investigation follow-up on 2026-03-27:

- rejected informational-task state dedup candidate:
  - rejected run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_190016`
- rejected popup create-path async first-paint candidate:
  - rejected run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_190415`
- rejected popup show cleanup candidate:
  - rejected run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_190502`
- diagnostic-only follow-up with task-card substep and lock timings:
  - evidence run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_190636`
- rejected coalesced popup-refresh candidate:
  - rejected run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_190843`

Observed:

- the informational-task state dedup candidate was not exercised by the real Compare scenario:
  - `FileOps.InfoTask.UpdateSkippedCount` never appeared in `2026-03-27_190016`
  - end-to-end latency regressed, so that change was rejected
- the popup create-path async first-paint candidate materially reduced `compare.ui.taskcard_update_us`, but also materially delayed visible progress:
  - `compare.ui.taskcard_update_us`: `29031us avg -> 10387us avg`
  - `compare.ui.scan_progress_to_visible_latency_ms`: `66056us avg -> 95966us avg`
  - `compare.ui.content_progress_to_visible_latency_ms`: `56881us avg -> 101532us avg`
- the narrower popup show cleanup candidate (removing the redundant `ShowWindow` before `SetWindowPos(... SWP_SHOWWINDOW)`) did not hold up:
  - `FileOps.InfoTask.EnsurePopupVisibleUs` increased to `47464us`
  - scan/content progress latency stayed flat-to-worse, so that candidate was also rejected
- the diagnostic-only run `2026-03-27_190636` established a clearer root-cause split:
  - `compare.ui.taskcard_build_us` is negligible (`14us avg`, `20us max`)
  - `compare.ui.taskcard_apply_us` is where the time goes (`25616us avg`, `49689us max`)
  - `FileOps.InfoTask.EnsurePopupVisibleUs` alone accounts for nearly all of the first expensive update (`48300us`)
  - file-ops state mutex contention is not the cause:
    - `FileOps.InfoTask.Update.LockWaitUs`: `0.67us avg`
    - `FileOps.InfoTask.Update.LockHoldUs`: `413us avg`
    - `FileOps.InfoTask.Collect.LockWaitUs`: `1us`
    - `FileOps.InfoTask.Collect.CopyUs`: `551us`
- the coalesced popup-refresh candidate confirmed that later `taskcard_apply_us` can be lowered:
  - in `2026-03-27_190843`, the posted/coalesced refresh path produced one posted refresh and one coalesced refresh
  - two later task-card applies dropped to about `1.1ms`, and the last one was about `30ms`
  - but the user-visible cost was unacceptable:
    - `compare.ui.scan_progress_to_visible_latency_ms`: `66056us avg -> 115703us avg`
    - `compare.ui.content_progress_to_visible_latency_ms`: `56881us avg -> 131266us avg`

Interpretation:

- keep the newer instrumentation:
  - `compare.ui.taskcard_build_us`
  - `compare.ui.taskcard_apply_us`
  - `FileOps.InfoTask.EnsurePopupVisibleUs`
  - `FileOps.InfoTask.Update.LockWaitUs`
  - `FileOps.InfoTask.Update.LockHoldUs`
  - `FileOps.InfoTask.Collect.LockWaitUs`
  - `FileOps.InfoTask.Collect.CopyUs`
- the remaining Compare task-card cost is not in Compare payload construction and not in file-ops state lock contention
- the initial expensive update is dominated by popup creation / first visible paint
- later expensive updates are somewhere in the popup apply/paint path after the state lock is released
- future work should target popup rendering/layout or a lower-latency refresh coalescing strategy inside the popup itself, not Compare payload building or informational-task vector dedup

File Operations popup render-split follow-up on 2026-03-27:

- diagnostic render-split run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_191145`

Observed:

- the popup render path now emits:
  - `FileOps.Popup.Render.BuildSnapshotUs`
  - `FileOps.Popup.Render.CardLayoutUs`
  - `FileOps.Popup.Render.AutoResizeUs`
  - `FileOps.Popup.Render.ScrollLayoutUs`
  - `FileOps.Popup.Render.DrawUs`
  - `FileOps.Popup.Render.TotalUs`
- in the measured Compare scenario from `2026-03-27_191145`, the single captured popup render was modest:
  - `FileOps.Popup.Render.BuildSnapshotUs`: `1394us`
  - `FileOps.Popup.Render.CardLayoutUs`: `8us`
  - `FileOps.Popup.Render.AutoResizeUs`: `5us`
  - `FileOps.Popup.Render.ScrollLayoutUs`: `129us`
  - `FileOps.Popup.Render.DrawUs`: `5094us`
  - `FileOps.Popup.Render.TotalUs`: `9406us`
- the same run still showed expensive task-card apply times:
  - `compare.ui.taskcard_apply_us`: `50484us avg`, `86898us max`
  - `compare.ui.taskcard_update_us`: `51587us avg`, `88305us max`
  - `FileOps.InfoTask.EnsurePopupVisibleUs`: `62117us` for the first expensive update
- the render split therefore rules out several suspects:
  - popup snapshot collection is not dominant for the measured render
  - popup card-height computation and auto-resize are negligible in the measured render
  - raw D2D drawing is measurable but still much smaller than the slowest `taskcard_apply_us` spikes

Interpretation:

- the remaining expensive `taskcard_apply_us` time is not explained by normal popup render body cost alone
- the first expensive update is still strongly tied to popup creation/show
- later expensive updates likely include synchronous windowing work around invalidation/paint dispatch or other popup non-draw work adjacent to `CreateOrUpdateInformationalTask`, not just the measured render body itself
- the next useful instrumentation target is popup window-message timing around `WM_PAINT`, `WM_NCPAINT`, caption-status redraws, and any synchronous frame work triggered by popup visibility changes

Caption-status redraw candidate follow-up on 2026-03-27:

- rejected candidate run: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_192347`

Observed:

- delaying the caption `Ok` state until a finished successful/cancelled task existed was intended to avoid an early non-client frame redraw on popup bring-up
- that candidate did not help the user-visible path:
  - `compare.ui.scan_progress_to_visible_latency_ms`: `66056us avg -> 147408us avg`
  - `compare.ui.content_progress_to_visible_latency_ms`: `56881us avg -> 170409us avg`
- although some internal task-card apply samples were smaller in that run, the end-to-end Compare responsiveness regressed too much to justify keeping the behavior change

Interpretation:

- reject the caption-status state change and keep the earlier Compare baseline behavior
- keep using `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_185337` as the current best Compare reference run

Popup window-message timing follow-up on 2026-03-27:

- added non-behavioral timing emits for:
  - `FileOps.InfoTask.EnsurePopupVisible.CreateUs`
  - `FileOps.InfoTask.EnsurePopupVisible.ShowWindowUs`
  - `FileOps.InfoTask.EnsurePopupVisible.SetWindowPosUs`
  - `FileOps.InfoTask.EnsurePopupVisible.RedrawWindowUs`
  - `FileOps.InfoTask.EnsurePopupVisible.InvalidateUs`
  - `FileOps.Popup.WmPaintUs`
  - `FileOps.Popup.WmNcPaintUs`
  - `FileOps.Popup.WmNcActivateUs`

Current status:

- the code is in place and builds cleanly
- targeted reruns after adding these metrics returned success locally but did not produce a fresh archived Commands run or a completed Compare trace under `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\commands`
- treat this slice as instrumentation-ready but verification-blocked until the Commands harness writes a fresh Compare capture again

Waited Commands rerun follow-up on 2026-03-27:

- fixed the `cmd_compare_directories_progress_perf` harness so dataset creation no longer blocks the UI thread:
  - the compare payload tree is now built through a worker-backed progress path in `Commands.SelfTest.cpp`
  - this removed the mid-case `EXITCODE=-1` failure during test dataset generation
- fixed the selftest perf sink reset path so `last_run/perf/perf_metrics.jsonl` is deleted at run start instead of silently appending across runs
- fresh waited baseline run:
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_212710`

Observed:

- the waited launch produced a clean archived Commands run with fresh run metadata:
  - `commands_results.json` reports `run_started_utc = 2026-03-27T20:27:08.500Z`
  - archived `perf_metrics.jsonl` now carries `runId = 2026-03-27_202709`
- the new popup timing split is now proven in a fresh Compare capture:
  - `FileOps.InfoTask.EnsurePopupVisibleUs`: `66821us`
  - `FileOps.InfoTask.EnsurePopupVisible.CreateUs`: `12936us`
  - `FileOps.InfoTask.EnsurePopupVisible.ShowWindowUs`: `6460us`
  - `FileOps.InfoTask.EnsurePopupVisible.SetWindowPosUs`: `652us`
  - `FileOps.InfoTask.EnsurePopupVisible.RedrawWindowUs`: `43829us`
  - `FileOps.Popup.WmPaintUs`: `42991us`
  - `FileOps.Popup.WmNcPaintUs`: `1278us`
  - `compare.ui.taskcard_apply_us`: first expensive sample `68909us`
- Compare progress-path latency in the fresh waited run:
  - `compare.ui.content_progress_to_visible_latency_ms`: first sampled value `117875us`
  - the run still reaches completion with `contentDone = 120` / `contentTotal = 120`

Interpretation:

- the harness/capture path is unblocked again for this Compare scenario, but only when the Commands selftest is launched as a waited GUI process
- the first expensive task-card update is still dominated by popup visibility and synchronous paint work, not by payload construction
- `EnsurePopupVisible.RedrawWindowUs` plus `FileOps.Popup.WmPaintUs` remains the clearest next hotspot for popup-show optimization
- later repeated reruns can still fail early if another foreground `RedSalamander.exe` instance is already running, so use a clean waited launch environment when collecting Compare baselines

Compare waited-run stability follow-up on 2026-03-27:

- kept harness-stability changes:
  - selftest-only startup simplifications in `RedSalamander.cpp`:
    - skip splash for any selftest launch
    - skip persisted pane-path restore during deterministic Commands startup
    - skip settings hot-reload startup during selftests
    - skip synchronous `UpdateWindow(...)` during selftests
  - deeper startup tracing in `RedSalamander.cpp` and `FolderWindow.cpp`
  - lighter Compare perf dataset in `Commands.SelfTest.cpp`:
    - `10` directories x `8` files with `32 KiB` payloads per side
- fresh consecutive passing waited runs after those changes:
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_233157`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_233202`

Observed:

- these two runs both pass and archive cleanly back-to-back, which is stronger repeatability than the earlier waited Compare sequence
- the consecutive runs land in a tighter range on the popup-show hotspot:
  - `FileOps.InfoTask.EnsurePopupVisibleUs`: `51497us` and `54104us`
  - `FileOps.InfoTask.EnsurePopupVisible.RedrawWindowUs`: `34598us` and `34959us`
  - `FileOps.Popup.WmPaintUs`: `34181us` and `34517us`
  - first `compare.ui.taskcard_apply_us`: `53014us` and `55456us`
- first visible progress latencies in these lighter-dataset runs are also in a tighter band:
  - `compare.ui.content_progress_to_visible_latency_ms`: `75763us` and `77006us`
  - `compare.ui.scan_progress_to_visible_latency_ms`: `78613us` and `79583us`

Interpretation:

- keep the selftest startup simplifications and lighter Compare dataset because they materially improved waited-run repeatability
- treat `2026-03-27_233157` and `2026-03-27_233202` as the current practical Compare harness baseline pair for future popup A/B work
- do not treat these runs as strict apples-to-apples improvements versus `2026-03-27_231703`, because the synthetic Compare dataset is lighter in this newer harness configuration

Rejected Compare popup-show candidate follow-up on 2026-03-27:

- rejected candidate:
  - create the initial Compare informational task card before `_session->StartScan()` so popup creation/first paint is front-loaded before scan/content progress begins
- evidence runs:
  - unstable candidate pass: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_233746`
  - failing rerun: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_233802`

Observed:

- the first passing candidate run was worse than the new stable baseline pair on the main popup metrics:
  - `FileOps.InfoTask.EnsurePopupVisibleUs`: `62966us` vs `51497us` / `54104us`
  - `FileOps.InfoTask.EnsurePopupVisible.RedrawWindowUs`: `38949us` vs `34598us` / `34959us`
  - first `compare.ui.taskcard_apply_us`: `64579us` vs `53014us` / `55456us`
  - first `compare.ui.content_progress_to_visible_latency_ms`: `124112us` vs `75763us` / `77006us`
- the second waited rerun then failed the scenario:
  - `compare started` and `scan progress` were observed with `scanEntries=0` / `contentTotal=0`
  - the case timed out waiting for content progress and completion
  - archived failure: `Compare progress perf test did not observe content progress.`

Interpretation:

- reject the “pre-create task card before `StartScan()`” reorder
- keep using the post-`StartScan()` task-card creation order from the stable baseline
- future popup-show candidates should target `EnsurePopupVisible.RedrawWindowUs` / `WM_PAINT` more directly rather than changing compare-run sequencing

Compare waited-run methodology correction on 2026-03-27:

- kept harness-hardening follow-up:
  - serialized selftest trace-file writes in `SelfTestCommon.cpp`
  - selftest-only top-level SEH trace marker in `RedSalamander.cpp`
  - `SplashScreen::SetOwner(...)` now runs only when a splash window actually exists
- corrected rerun methodology:
  - the earlier `RUN1=-1` / `RUN2=-1` evidence collected immediately after this hardening pass was invalid because the waited selftests were launched in parallel
  - the interleaved `last_run/trace.txt` proved that two startup sequences overlapped, so those failures are not valid product-signal data
- clean sequential waited reruns after correcting that methodology:
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_234921`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_234929`

Observed:

- both sequential waited reruns pass and archive cleanly
- the latest traces show the full expected path through:
  - `InitInstance`
  - `RunApplication: entering message loop`
  - `CommandsSelfTest::Run`
  - `cmd_compare_directories_progress_perf`
  - archive to repo
- the current popup-show metrics in these sequential runs are consistent with the existing popup hotspot diagnosis:
  - `FileOps.InfoTask.EnsurePopupVisibleUs`: `57307us` and `58746us`
  - `FileOps.InfoTask.EnsurePopupVisible.RedrawWindowUs`: `37326us` and `37622us`
  - `FileOps.Popup.WmPaintUs`: `36756us` and `36746us`
  - first `compare.ui.taskcard_apply_us`: `58722us` and `60248us`

Interpretation:

- the immediate blocker is no longer a mysterious fresh startup crash in product code
- the corrected sequential methodology restores a usable waited Compare harness
- keep treating popup-show work as the active optimization target, but use strictly sequential waited launches when comparing future candidates

Folder-view smoke rerun notes:

- baseline: `Specs/TestRuns/7d3a1247382a/Commands/2026-03-26_175258`
- candidate: `Specs/TestRuns/7d3a1247382a/Commands/2026-03-26_175428`

Observed with `Tools/Show-PerfRuns.ps1`:

- `App.Startup.Metric.TimeToFirstPanePopulated` moved from about `1044354us` to `1055145us` (`+1.0%`)
- `iconcache.shgetfileinfo_us` `p95` improved from about `11820us` to `10682us` (`-9.6%`)
- `iconcache.lock_wait_us` `p95` moved from `2` to `1`, but both runs only captured startup/light icon-cache activity

Interpretation:

- these two Commands runs are good proof that JSONL archiving and run-to-run comparison work for folder-view-adjacent cases
- they are not sufficient as the folder-pane baseline because they do not emit the full `render.*` or `icons.*` profile needed for pane rendering review

## Summary

This plan adds low-overhead performance instrumentation to three hotspot areas:

1. folder-pane rendering, icon loading, and icon caching
2. file-operations scheduling, queueing, progress, and pre-calculation
3. search backend cutover, callback bridging, batching, and search UI update cost

The instrumentation must stay behavior-preserving, warning-clean, and compatible with the project’s existing debug/perf logging style. The goal is not to introduce a heavyweight tracing framework. The goal is to produce stable, comparable metrics that allow:

- one-run hotspot identification
- run-to-run regression review
- scenario-based performance evaluation before and after code changes

## Goals

- Measure where wall time is spent in each hotspot path.
- Separate actual work from coordination overhead.
- Measure contention, queue delay, UI wakeup pressure, and batch fragmentation.
- Make results comparable across repeated local runs.
- Keep instrumentation cheap enough for Debug scenario runs.

## Non-Goals

- No production telemetry backend.
- No large tracing dependency.
- No behavior changes to scheduling, rendering, or search logic in this slice.
- No optimization changes in the same patch unless needed to keep builds/tests green.

## Instrumentation Style

Use one minimal shape across all touched files.

### Helper APIs

Implement or extend `Debug::Perf` with a compact set of helpers:

```cpp
namespace Debug::Perf
{
    void EmitCounter(std::wstring_view name, int64_t value = 1);
    void EmitValue(std::wstring_view name, int64_t value);
    void EmitDurationUs(std::wstring_view name, int64_t elapsedUs);
}
```

These helpers can be wrappers around the existing perf/debug output mechanism. The key requirement is a stable metric name and value emission path.

### Scoped Timer Shape

```cpp
class ScopedPerfTimer final
{
public:
    explicit ScopedPerfTimer(std::wstring_view metricName) noexcept;
    ~ScopedPerfTimer() noexcept;

    ScopedPerfTimer(const ScopedPerfTimer&) = delete;
    ScopedPerfTimer& operator=(const ScopedPerfTimer&) = delete;

private:
    std::wstring_view _metricName;
    int64_t _startQpc{};
};
```

Use this for whole-function or block timing where scope-based emission is appropriate.

### Manual Timer / Value Shape

```cpp
class ScopedPerfValue final
{
public:
    explicit ScopedPerfValue(std::wstring_view metricName) noexcept;
    void Stop() noexcept;
    ~ScopedPerfValue() noexcept;

private:
    std::wstring_view _metricName;
    int64_t _startQpc{};
    bool _stopped{};
};
```

Use this for:

- lock wait timing
- lock hold timing
- timing only one branch inside a larger function
- payload enqueue-to-consume latency

### Convenience Macros

```cpp
#define RS_PERF_COUNTER(name) Debug::Perf::EmitCounter(L##name)
#define RS_PERF_COUNTER_V(name, value) Debug::Perf::EmitCounter(L##name, static_cast<int64_t>(value))
#define RS_PERF_VALUE(name, value) Debug::Perf::EmitValue(L##name, static_cast<int64_t>(value))
#define RS_PERF_TIMER(name) ScopedPerfTimer rsPerfTimer_##__LINE__{L##name}
```

### Naming Conventions

- Use subsystem prefixes:
  - `render.*`
  - `icons.*`
  - `iconcache.*`
  - `FileOps.*`
  - `search.*`
  - `find.*`
- Use `_us` for microseconds.
- Use `_ms` only where the metric is naturally end-to-end or already described in milliseconds.
- Use `.count`, `.attempts`, `.hits`, `.failures`, `.successes` as plain counters.
- Use explicit reason suffixes instead of free-form text where possible.

### Output Contract

Every metric emission should be parseable from logs into rows like:

```text
timestamp, process, thread, metric_name, value, unit, scenario, build, branch
```

The stored perf artifact format must be JSONL for easy parsing and append-only collection. Each emitted perf record should serialize as one JSON object per line with stable keys. CSV or markdown summaries can be derived later, but JSONL is the source-of-truth raw format.

Required JSONL fields:

- `timestamp`
- `process`
- `thread`
- `metric`
- `value`
- `unit`
- `scenario`
- `build`
- `branch`
- `commit`
- `machineHash`
- `runId`

If the current perf logger cannot emit JSONL directly, add a thin adapter or post-processing path that converts stable debug/perf output into JSONL before the run is considered complete.

## Patch Order

Patch in this order to keep the schema stable before touching many files:

1. shared perf helper additions
2. `RedSalamander/FolderView.Rendering.cpp`
3. `RedSalamander/FolderView.Icons.cpp`
4. `RedSalamander/IconCache.cpp`
5. `RedSalamander/FolderWindow.FileOperations.State.cpp`
6. `Common/LocalSearchIndexCore.cpp`
7. `RedSalamander/FindFilesWindow.cpp`
8. scenario scripts or parsing helpers if needed
9. baseline runs and archived metric reports

## Implementation Checklist By File

### 1. Shared Perf Helper Layer

Target files:

- existing `Debug::Perf` implementation files and headers in shared code

Tasks:

- [x] Add `EmitCounter`, `EmitValue`, and `EmitDurationUs` APIs if missing.
- [x] Add `ScopedPerfTimer`.
- [x] Add `ScopedPerfValue`.
- [x] Add minimal QPC helper if one does not already exist.
- [x] Keep APIs `noexcept` where possible.
- [x] Ensure no `catch (...)`.
- [x] Ensure helper overhead is low and does not allocate on hot path if avoidable.

Validation:

- [x] Build touched targets warning-clean.
- [x] Emit one sample metric from a harmless path and verify formatting.

### 2. Folder Rendering

Target file:

- `RedSalamander/FolderView.Rendering.cpp`

Target functions:

- `FolderView::Render(const RECT&)`
- `FolderView::DrawItem(FolderItem&)`

Metrics to add:

- `render.frame_us`
- `render.begin_to_enddraw_us`
- `render.present_us`
- `render.dirty_rect_area_px`
- `render.items_considered`
- `render.items_drawn`
- `render.empty_state_layout_creates`
- `render.draw_item_us`
- `render.item_has_icon`
- `render.item_placeholder_icon`
- `render.item_textlayout_label`
- `render.item_textlayout_details`
- `render.item_textlayout_metadata`
- `render.incremental_search_effect_updates`

Patch tasks:

- [x] Time the full render scope.
- [x] Time begin-draw through end-draw separately from present.
- [x] Record dirty rect area once per render.
- [x] Count considered vs drawn items in the visible draw lambda.
- [x] Count `CreateTextLayout` calls in empty-state and watermark branches.
- [x] Time `DrawItem`.
- [x] Count icon and text-layout branches inside `DrawItem`.
- [x] Count highlight drawing-effect mutations.

Interpretation targets:

- full frame spikes
- present blocking
- redraw scope too large
- layout churn
- unusually expensive item draw tail

### 3. Icon Queue and UI Bitmap Application

Target file:

- `RedSalamander/FolderView.Icons.cpp`

Target functions:

- `FolderView::QueueIconLoading()`
- `FolderView::ProcessIconLoadQueue()`
- `FolderView::OnCreateIconBitmap()`
- `FolderView::OnBatchIconUpdate()`
- `FolderView::OnIconLoaded()`

Metrics to add:

- `icons.queue_groups_total`
- `icons.queue_visible_groups`
- `icons.queue_build_us`
- `icons.cache_stamp_count`
- `icons.queue_wait_to_dequeue_us`
- `icons.extract_us`
- `icons.extract_failures`
- `icons.extract_retries`
- `icons.post_message_latency_us`
- `icons.ui_convert_us`
- `icons.ui_apply_count`
- `icons.invalidate_single_count`
- `icons.invalidate_full_count`
- `icons.invalidate_area_px`
- `icons.batch_update_scan_us`
- `icons.batch_update_retrieved`

Patch tasks:

- [x] Stamp each queued request with enqueue timestamp.
- [x] Measure queue build time and group counts.
- [x] Measure request wait time before worker dequeue.
- [x] Measure shell icon extraction duration.
- [x] Count extract failures and retries.
- [x] Timestamp UI payload posting and compute enqueue-to-consume latency.
- [x] Add per-conversion histogram value in `OnCreateIconBitmap`.
- [x] Count how many items one converted icon applies to.
- [x] Count full-view vs single-item invalidations.
- [x] Measure scan time in cache reconciliation handlers.

Interpretation targets:

- worker starvation
- shell extraction latency
- UI-thread conversion pressure
- icon completion redraw storms

### 4. Icon Cache

Target file:

- `RedSalamander/IconCache.cpp`

Target functions:

- `IconCache::GetIconBitmap`
- `IconCache::GetCachedBitmap`
- `IconCache::HasCachedIcon`
- `IconCache::ConvertIconToBitmapOnUIThread`
- `IconCache::QuerySysIconIndexForPath`
- `IconCache::Initialize`
- `IconCache::PrewarmBitmaps`
- `IconCache::ExtractSystemIcon`
- `EvictLRUIfNeeded()`
- `EvictPathQueryBatch()`
- `EvictAssociationQueryBatch()`

Metrics to add:

- `iconcache.lock_wait_us`
- `iconcache.lock_hold_us`
- `iconcache.get_bitmap_hit`
- `iconcache.get_bitmap_miss`
- `iconcache.miss_extract_us`
- `iconcache.miss_convert_us`
- `iconcache.miss_store_us`
- `iconcache.extract_attempt_count`
- `iconcache.extract_selected_size`
- `iconcache.extract_success_size`
- `iconcache.extract_total_us`
- `iconcache.ui_convert_hit_after_race`
- `iconcache.convert_wic_bitmap_us`
- `iconcache.convert_format_converter_us`
- `iconcache.convert_create_d2d_bitmap_us`
- `iconcache.path_query_hit`
- `iconcache.path_query_miss`
- `iconcache.shgetfileinfo_us`
- `iconcache.duplicate_path_query_race`
- `iconcache.device_lru_evict_scan_us`
- `iconcache.path_lru_evict_scan_us`
- `iconcache.association_lru_evict_scan_us`

Patch tasks:

- [x] Add lock wait and lock hold timing at the major `_mutex` entry sites.
- [x] Separate hit vs miss in `GetIconBitmap`.
- [x] Split miss cost into extract, convert, and store phases.
- [x] Time the full shell extraction cascade and record attempt count.
- [x] Record selected and actual successful image-list size.
- [x] Split conversion into WIC bitmap creation, format converter, and D2D bitmap creation.
- [x] Time `SHGetFileInfoW` path queries and count duplicate-race insertions.
- [x] Time LRU eviction scans.

Interpretation targets:

- mutex contention
- shell API cost
- conversion-stage bottlenecks
- eviction spikes

### 5. File Operations

Target file:

- `RedSalamander/FolderWindow.FileOperations.State.cpp`

Target functions:

- `effectiveMaxConcurrencyLocked`
- `tryDequeueWorkLocked`
- `workerMain`
- `hasSchedulableWorkLocked`
- `EnterOperation`
- `LeaveOperation`
- `RemoveFromQueue`
- `NotifyQueueChanged`
- `Task::FileSystemProgress`
- `Task::FileSystemItemCompleted`
- `ReportDirectorySizeProgress`
- pre-calc block around `GetDirectorySize`
- `WaitWhilePaused`
- conflict wait sites
- operation completion block with existing `Debug::Perf::Emit(L"FileOps.Operation", ...)`

Metrics to add:

- `FileOps.Scheduler.DequeueAttempts`
- `FileOps.Scheduler.DequeueSuccess`
- `FileOps.Scheduler.ScanJobsPerAttempt`
- `FileOps.Scheduler.EffectiveConcurrency`
- `FileOps.Scheduler.WaitForWorkUs`
- `FileOps.Scheduler.ProcessIndexUs`
- `FileOps.Scheduler.InFlightPerJob`
- `FileOps.Queue.EnterCount`
- `FileOps.Queue.WaitUs`
- `FileOps.Queue.DepthOnEnter`
- `FileOps.Queue.CancelWhileWaiting`
- `FileOps.Queue.ActiveOperations`
- `FileOps.Queue.NotifyAllCount`
- `FileOps.Progress.CallbackCount`
- `FileOps.Progress.CallbackUs`
- `FileOps.Progress.LockWaitUs`
- `FileOps.Progress.LockHoldUs`
- `FileOps.Progress.BytesPerCallback`
- `FileOps.Progress.InFlightFileCount`
- `FileOps.Progress.InFlightEvictions`
- `FileOps.Progress.PathUpdateBytes`
- `FileOps.ItemCompleted.CallbackUs`
- `FileOps.PreCalc.TotalUs`
- `FileOps.PreCalc.CallbackCount`
- `FileOps.PreCalc.CallbackUs`
- `FileOps.PreCalc.LockWaitUs`
- `FileOps.PreCalc.BytesFound`
- `FileOps.PreCalc.ItemCount`
- `FileOps.Pause.WaitUs`
- `FileOps.Conflict.WaitUs`
- `FileOps.Conflict.PendingCount`

Extend rollups:

- `FileOps.Operation`
  - `queueWaitUs`
  - `preCalcUs`
  - `progressLockWaitUs`
  - `progressLockHoldUs`
  - `schedulerWaitUs`
  - `dequeueAttempts`
  - `dequeueSuccess`
  - `progressCallbacks`
  - `itemCallbacks`
  - `conflictWaitUs`
  - `pauseWaitUs`

Patch tasks:

- [x] Count dequeue attempts and successes.
- [x] Record number of jobs scanned per dequeue attempt.
- [x] Record effective concurrency decisions.
- [x] Time worker wait on `_cv`.
- [x] Time `processIndex` only.
- [x] Record queue wait from enqueue to operation admission.
- [x] Count queue depth and noisy `notify_all` usage.
- [x] Time progress callback cost.
- [x] Time `_progressMutex` wait and hold durations.
- [x] Record bytes-per-callback and in-flight counts.
- [x] Count in-flight table evictions.
- [x] Record path-character churn for progress path updates.
- [x] Time pre-calculation total and callback cost.
- [x] Time pause and conflict waits separately.
- [x] Extend final operation rollups.

Interpretation targets:

- scheduler starvation
- fairness underutilization
- queue serialization
- progress callback flood
- progress lock contention
- pre-calc dominating total task time

### 6. Search Backend

Target file:

- `Common/LocalSearchIndexCore.cpp`

Target functions:

- `LocalSearchIndexCore::Repository::EnumerateNoWait(...)`
- `EmitRepositoryProgress(...)`
- `UpdateRepositoryRuntimeStatus(...)`
- live-fallback candidate lambda passed to `EnumerateLiveFileSystem(...)`

Metrics to add:

- `search.enumerate_nowait.calls`
- `search.enumerate_nowait.total_ms`
- `search.backend.sqlite.attempts`
- `search.backend.sqlite.query_ms`
- `search.backend.sqlite.successes`
- `search.backend.sqlite.soft_fallbacks`
- `search.backend.sqlite.fallback_reason.warmup_running`
- `search.backend.sqlite.fallback_reason.store_stale`
- `search.backend.sqlite.fallback_reason.cutover_blocked`
- `search.backend.sqlite.fallback_reason.sqlite_failure`
- `search.backend.sqlite.fallback_reason.store_missing`
- `search.backend.sqlite.fallback_reason.store_invalid`
- `search.backend.inmemory.attempts`
- `search.backend.inmemory.query_ms`
- `search.backend.live_fallback.attempts`
- `search.backend.live_fallback.query_ms`
- `search.sqlite.partial_emit_count`
- `search.sqlite.partial_emit_set_size`
- `search.dedupe.lookups`
- `search.dedupe.hits`
- `search.dedupe.lookup_ms`
- `search.progress.emit_calls`
- `search.progress.emit_forced`
- `search.progress.emit_suppressed`
- `search.progress.callback_ms`
- `search.runtime_status.updates`

Patch tasks:

- [x] Time `EnumerateNoWait` end-to-end.
- [x] Time sqlite branch, in-memory branch, and live fallback branch separately.
- [x] Count fallback outcomes and reasons.
- [x] Record sqlite partial-emit set size before live fallback.
- [x] Count dedupe lookups and dedupe hits.
- [x] Time dedupe path folding plus set lookup.
- [x] Count progress attempts, forced emits, and suppressed emits.
- [x] Time progress callback execution.
- [x] Count runtime status updates.

Interpretation targets:

- sqlite optimistic failures
- fallback penalty
- dedupe overhead after partial sqlite emission
- overly chatty progress path

### 7. Search Callback Bridge and UI

Target file:

- `RedSalamander/FindFilesWindow.cpp`

Target functions:

- `SearchSessionController::Start(...)`
- `SearchSessionController::Run(...)`
- `SearchCallbacks::FlushResults()`
- `SearchCallbacks::FileSystemSearchMatch(...)`
- `SearchCallbacks::FileSystemSearchProgress(...)`
- result payload handler
- `FindFilesWindow::OnSearchProgress(...)`
- `RefreshResultsView(...)`
- `RebuildResultsList()`
- `SortResults()`
- `FindFilesWindow::OnSearchComplete(...)`

Metrics to add:

- `find.session.start_count`
- `find.session.run_total_ms`
- `find.session.backend.native_search_calls`
- `find.session.backend.native_search_ms`
- `find.session.backend.unsupported_fallbacks`
- `find.session.backend.fallback_engine_ms`
- `find.results.match_callbacks`
- `find.results.match_callback_ms`
- `find.results.batch_size`
- `find.results.flush_count`
- `find.results.flush_trigger.progress`
- `find.results.flush_trigger.capacity`
- `find.progress.callback_count`
- `find.progress.callback_ms`
- `find.progress.current_path_chars`
- `find.service_status.snapshot_count`
- `find.ui.results_message_count`
- `find.ui.results_message_batch_size`
- `find.ui.results_handler_ms`
- `find.ui.results_inserted_count`
- `find.ui.results_updated_count`
- `find.ui.results_full_rebuild_count`
- `find.ui.refresh_results_view_ms`
- `find.ui.sort_results_ms`
- `find.ui.progress_message_count`
- `find.ui.progress_handler_ms`
- `find.ui.status_refresh_count`
- `find.ui.status_refresh_ms`
- `find.ui.progress_to_visible_latency_ms`
- `find.ui.results_to_visible_latency_ms`
- `find.ui.complete_handler_ms`

Payload struct changes:

- add enqueue timestamp field to result payload
- add enqueue timestamp field to progress payload

Patch tasks:

- [x] Count started sessions.
- [x] Time `Run(...)` end-to-end.
- [x] Count native-search vs fallback-engine use.
- [x] Time match callback and progress callback bridge cost.
- [x] Record batch sizes and flush counts.
- [x] Distinguish flush-by-progress vs flush-by-capacity.
- [x] Record current-path payload size histogram.
- [x] Count service status snapshot updates.
- [x] Time result handler, progress handler, refresh, sort, and completion handlers.
- [x] Record insert vs update counts in UI result mutations.
- [x] Add payload timestamps and compute queue latency to visible application.

Interpretation targets:

- bridge overhead
- tiny batches and excessive flushes
- UI sorting and rebuild cost
- UI queue congestion
- progress message over-frequency

## Test Plan

Run tests after each major subsystem patch and again after the full instrumentation slice lands.

Perf scenarios in this plan are intended to run in Debug mode, using the same broad execution framework and archival conventions as the existing `--selftest` / command self-test infrastructure. The target state is a repeatable Debug-only perf scenario flow that behaves like self-test execution, but emits metric JSONL artifacts in addition to the normal test outputs.

### Build Validation

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander
```

If the touched files affect search or monitor-only projects indirectly, also run:

```powershell
.\build.ps1 -ProjectName RedSalamanderSearchService
.\build.ps1 -ProjectName RedSalamanderMonitor
```

If the solution is already stable and time permits, run the full Debug solution build:

```powershell
.\build.ps1
```

### Self-Tests and Focused Functional Coverage

Use the built-in self-test and command test guidance from `README.md` and `Specs/Testing/Testing_SelfTests.md`.

Minimum functional validation set after instrumentation:

```powershell
.\build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=app_windows_open_and_close_modeless
.\build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=app_find_files
.\build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=app_compare_directories
```

If exact self-test names differ in-tree, use the nearest equivalent existing cases that validate:

- modeless window open/close flows
- find/search UI flows
- file operations or compare-directory background flows

Search-service validation if available:

```powershell
.\build\x64\Debug\RedSalamanderSearchService.exe
```

or any existing service self-test/manual smoke command used by the repo.

### Perf Scenario Execution Model

Perf scenarios should be runnable from Debug builds in a way that is close to the existing self-test framework:

- scenarios should reuse the same command-line and harness style where practical
- scenario naming should be explicit and stable
- scenario setup should be deterministic enough for run-to-run comparison
- perf runs should emit JSONL metrics as part of the scenario output

Implementation tasks:

- [x] Define how perf scenarios are invoked from the Debug executable or associated harness.
- [x] Reuse existing self-test scenario setup/teardown where possible instead of inventing a separate runner.
- [x] Ensure each perf scenario records `scenario`, `runId`, `machineHash`, `branch`, and `commit`.
- [x] Ensure failed or aborted scenarios still leave enough metadata to explain incomplete runs.

### Manual Smoke Validation

For each patched area, do one short smoke pass:

- folder navigation between large directories
- rapid scroll in a large folder view
- icon-pop-in scenario on a cold folder
- copy or move a directory tree with many files
- launch Find Files on a warm indexed root
- launch Find Files on a cold or forced-fallback root

## Metric Scenario Suite

Each scenario should be run at least twice:

- Run A: warm-up run, not used for comparison
- Run B: measured run

For baseline or regression comparison, run a third measured pass if variance looks high.

### Scenario 1. Folder Pane Warm Scroll

Purpose:

- measure render and icon steady-state cost after cache warmup

Steps:

1. Open a folder with at least several thousand entries.
2. Let initial enumeration settle.
3. Scroll continuously for 10 to 15 seconds.

Primary metrics:

- `render.frame_us`
- `render.present_us`
- `render.items_drawn`
- `icons.invalidate_full_count`
- `iconcache.lock_wait_us`

Review:

- p50, p95, p99 for render timings
- full invalidation count per 10 seconds
- icon-cache lock tail latency

### Scenario 2. Folder Pane Cold Icon Pop-In

Purpose:

- separate shell extraction, UI conversion, and redraw pressure

Steps:

1. Open a folder not visited recently or after clearing relevant icon caches if there is a safe dev-only path.
2. Record from first visible paint until icon pop-in stabilizes.

Primary metrics:

- `icons.extract_us`
- `icons.ui_convert_us`
- `icons.post_message_latency_us`
- `icons.invalidate_full_count`
- `iconcache.shgetfileinfo_us`
- `iconcache.extract_total_us`

Review:

- extraction p95
- conversion p95
- queue latency tail
- count of full-view redraws

### Scenario 3. File Operations Many Small Files

Purpose:

- expose queue, callback, and lock overhead

Steps:

1. Prepare a tree with many small files and nested directories.
2. Run copy or move to another local directory.
3. Repeat once warm.

Primary metrics:

- `FileOps.Queue.WaitUs`
- `FileOps.Scheduler.DequeueAttempts`
- `FileOps.Scheduler.ProcessIndexUs`
- `FileOps.Progress.CallbackUs`
- `FileOps.Progress.LockWaitUs`
- `FileOps.Progress.BytesPerCallback`
- `FileOps.PreCalc.TotalUs`
- `FileOps.Operation`

Review:

- queue wait
- dequeue success ratio
- progress callback rate
- progress lock contention
- pre-calc share of total duration

### Scenario 4. File Operations Large Files

Purpose:

- separate actual I/O from callback and progress bookkeeping

Steps:

1. Copy a smaller set of large files.
2. Compare against Scenario 3.

Primary metrics:

- `FileOps.Progress.CallbackCount`
- `FileOps.Progress.BytesPerCallback`
- `FileOps.Progress.LockHoldUs`
- `FileOps.Operation`

Review:

- callback count should drop sharply relative to bytes moved
- lock wait should not dominate total time

### Scenario 5. Search Warm Indexed Query

Purpose:

- measure sqlite or in-memory path under expected fast conditions

Steps:

1. Run a query on a root expected to be indexed and current.
2. Use both a low-result and high-result query.

Primary metrics:

- `search.enumerate_nowait.total_ms`
- `search.backend.sqlite.query_ms`
- `search.backend.inmemory.query_ms`
- `find.results.batch_size`
- `find.ui.results_handler_ms`
- `find.ui.sort_results_ms`

Review:

- backend latency
- result batch efficiency
- UI result update cost

### Scenario 6. Search Fallback Query

Purpose:

- expose sqlite-to-live-fallback penalty and dedupe cost

Steps:

1. Run a query on a root likely to force live fallback or use an intentional dev setup that triggers fallback.
2. Capture the full search lifecycle.

Primary metrics:

- `search.backend.sqlite.soft_fallbacks`
- `search.backend.sqlite.fallback_reason.*`
- `search.sqlite.partial_emit_set_size`
- `search.dedupe.lookups`
- `search.dedupe.hits`
- `search.backend.live_fallback.query_ms`
- `find.ui.results_to_visible_latency_ms`

Review:

- fallback reason distribution
- partial sqlite emit size
- dedupe hit ratio
- UI backlog under fallback

### Scenario 7. Search Progress Storm

Purpose:

- verify progress throttling and batch fragmentation behavior

Steps:

1. Run a high-result search that emits frequent progress.
2. Capture progress and result handling cost.

Primary metrics:

- `search.progress.emit_calls`
- `search.progress.emit_forced`
- `search.progress.emit_suppressed`
- `find.results.flush_trigger.progress`
- `find.results.batch_size`
- `find.ui.progress_handler_ms`
- `find.ui.status_refresh_count`

Review:

- whether progress is fragmenting result batches
- whether UI status refresh is too chatty

## Metric Collection and Reporting

### Per-Run Report Requirements

For each measured scenario, archive a short report with:

- date/time
- git branch and commit
- build config
- scenario name
- dataset description
- run number
- whether the run is warm-up or measured
- top metrics table
- short interpretation notes

Recommended report location:

- store perf outputs close to the existing test-run archive structure
- use the same date-folder convention as other tests
- place runs under the machine-hash folder used for other archived test artifacts on that machine

Target shape:

```text
<test-run-root>\<machine-hash>\<yyyy-mm-dd>\perf\
```

Within that folder, store:

- raw JSONL metrics
- one scenario metadata file if needed
- derived summary markdown and/or CSV

Implementation tasks:

- [x] Identify the exact existing archived test-run root used by the repo/machine.
- [x] Add a perf subfolder convention under the same machine-hash and date folder structure.
- [x] Ensure perf scenario outputs are co-located with other test artifacts for the same day.

Recommended raw artifacts:

- raw perf JSONL log
- parsed CSV or TSV
- one markdown summary

### Required Aggregates

For timing metrics, report:

- count
- p50
- p95
- p99
- max

For counters and values, report:

- total
- rate per second when relevant
- ratio against another metric when relevant

Required derived ratios:

- `render.items_drawn / render.items_considered`
- `icons.invalidate_full_count / total icon apply events`
- `FileOps.Scheduler.DequeueSuccess / FileOps.Scheduler.DequeueAttempts`
- `FileOps.Progress.CallbackCount / bytes moved`
- `search.dedupe.hits / search.dedupe.lookups`
- `find.results.flush_count / total results`

### Run-to-Run Comparison Rules

Compare only runs that match on:

- same machine
- same build config
- same dataset
- same scenario steps
- same warm/cold state intent

When comparing two runs, classify:

- Improvement: p95 or total reduced by more than 10 percent with no major regression elsewhere
- Regression: p95 or total increased by more than 10 percent
- Noise: change within 10 percent unless repeated in three runs

For very small absolute timings, prefer percentage and absolute delta together.

## Reporting Script

Add a PowerShell reporting tool to display and compare runs and show performance evolution over time.

Expected responsibilities:

- discover archived perf runs from the machine-hash/date-folder structure
- load JSONL metric files
- summarize one run by scenario and metric
- compare two runs directly
- show metric evolution over time across multiple runs
- display p50, p95, p99, max, totals, and selected derived ratios
- highlight regressions and improvements using the comparison rules in this plan

Suggested script location:

- `Tools/` or another existing scripts location used by test/report helpers

Suggested commands:

```powershell
.\Tools\Show-PerfRuns.ps1
.\Tools\Show-PerfRuns.ps1 -Scenario FolderPaneWarmScroll
.\Tools\Show-PerfRuns.ps1 -CompareRun <baselineRunId> <candidateRunId>
.\Tools\Show-PerfRuns.ps1 -Metric render.frame_us -Trend
```

Implementation tasks:

- [x] Add a PowerShell script that enumerates perf runs from archived test folders.
- [x] Parse JSONL records into grouped scenario/metric summaries.
- [x] Support single-run summary mode.
- [x] Support two-run comparison mode.
- [x] Support time-series trend mode for one scenario or one metric.
- [x] Show baseline vs candidate deltas in both percent and absolute units.
- [x] Make the output readable in plain terminal text.
- [ ] Optionally emit CSV or markdown summaries for archival.

### Review Template

Use this review template per scenario:

```text
Scenario:
Build:
Branch/Commit:
Dataset:
Measured run:

Key metrics:
- metric_name: baseline -> candidate (delta %, delta absolute)

Interpretation:
- hottest path:
- main blocker:
- whether UI thread or worker thread is the bottleneck:
- whether contention is material:
- whether batching is effective:

Decision:
- no action
- investigate deeper
- implement optimization
```

## Acceptance Criteria

The plan is complete when all of the following are true:

- all metrics in scope are implemented or explicitly deferred with a reason
- touched targets build warning-clean
- focused self-tests pass
- at least one measured scenario is archived for each subsystem
- at least one run-to-run comparison is archived
- the resulting metrics are sufficient to identify the dominant cost in each hotspot area

## Deferred / Optional Follow-Ups

- add a small parser script under `Tools/` to convert perf log lines into CSV summaries
- add scenario-launch helper scripts for repeatable cold/warm setup
- add a compact markdown generator for `Specs/TestRuns/`
- add targeted self-test hooks if existing command self-tests are not enough to trigger the desired scenarios

The PowerShell reporting script is not optional for this slice. It is part of the expected deliverable because the instrumentation is only useful if runs can be reviewed and compared over time with low friction.

## Resume Checkpoint

Use this section as the restart point for the next perf session.

Current same-machine baselines to compare against:

- Search current-best baseline: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_173156`
- FileOps current-best baseline: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-27_145536`
- Compare progress/task-card current-best behavioral baseline: `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_185337`
- Compare waited-run harness baseline pair for popup-show work:
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_234921`
  - `Specs/TestRuns/4cb089111a23/Commands/2026-03-27_234929`

Current kept code state relevant to resumed Compare work:

- selftest dataset creation stays worker-backed in `Commands.SelfTest.cpp`
- selftest perf JSONL reset stays enabled in `SelfTestCommon.cpp`
- selftest startup stays simplified in `RedSalamander.cpp`:
  - no splash for selftests
  - no persisted pane-path restore during deterministic Commands startup
  - no settings hot-reload startup during selftests
  - no synchronous `UpdateWindow(...)` during selftests
- selftest trace writes are serialized in `SelfTestCommon.cpp`
- `SplashScreen::SetOwner(...)` is guarded behind `SplashScreen::Exist()` in `RedSalamander.cpp`
- popup timing instrumentation stays enabled:
  - `FileOps.InfoTask.EnsurePopupVisible*`
  - `FileOps.Popup.WmPaintUs`
  - `FileOps.Popup.WmNcPaintUs`
  - `FileOps.Popup.WmNcActivateUs`
- Compare task-card instrumentation stays enabled:
  - `compare.ui.taskcard_build_us`
  - `compare.ui.taskcard_apply_us`
  - `compare.ui.taskcard_update_us`

Current rejected candidates that should not be retried without a new hypothesis:

- defer initial Compare informational task until first progress/completion
- caption-status redraw state change on popup bring-up
- coalesced popup refresh that delayed visible progress
- pre-create initial Compare task card before `_session->StartScan()`

Methodology rules for the next session:

- use only sequential waited launches for Compare Commands perf runs
- never interpret overlapping selftest launches as product-signal data
- compare only runs with the same dataset and harness generation
- treat `2026-03-27_233157` / `2026-03-27_233202` as lighter-dataset harness-stability evidence, not apples-to-apples comparisons with older heavier runs

Recommended next experiment:

- target popup-show cost directly in `FolderWindow.FileOperations.Dialog.cpp` / `FolderWindow.FileOperations.Popup.cpp`
- optimize only if it specifically reduces synchronous `EnsurePopupVisible.RedrawWindowUs` and first `FileOps.Popup.WmPaintUs`
- verify against the sequential waited Compare baseline pair above before keeping any candidate

## Notes

- Keep instrumentation additive and behavior-preserving.
- Avoid measuring every lock site in `FolderWindow.FileOperations.State.cpp`; measure only the lock sites that answer a concrete contention question.
- Prefer explicit reason-coded counters over free-form text payloads.
- If overhead becomes measurable, keep the schema and trim sampling frequency before removing categories entirely.
