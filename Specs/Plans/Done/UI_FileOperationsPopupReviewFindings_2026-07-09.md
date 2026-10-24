# File Operations Popup — Pre-Push Review Findings & Remediation

**Status:** Done - R1-R9 and S1-S6 are fixed and covered; closeout reconciled on 2026-07-10.
**Date:** 2026-07-09
**Provenance:** max-effort local `/code-review` of the unpushed work on `master`
(`origin/master..HEAD`, 8 commits) plus the uncommitted working tree, run as ten
independent finder angles + a gap sweep. Every anchor below was read in the
current working tree, not inferred. Line numbers are from the working-tree copy of
each file and will drift as the file is edited — re-anchor by symbol before fixing.
**Scope reviewed:** the "Operation Clearview" implementation
(`UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md`) — footer segmented Queue/Parallel
control, auto-dismiss + compact-density + footer-only toggles, aggregate footer/taskbar
progress, status-at-a-glance stripe/chip, auto-collapse of completed cards, queued-task
**Start now**, and completed-task destination actions.
**Files:**
`RedSalamander/FolderWindow.FileOperations.Popup.cpp` / `.Popup.h`,
`RedSalamander/FolderWindow.FileOperations.State.Runtime.Part.cpp`,
`RedSalamander/FolderWindow.FileOperations.State.cpp`,
`RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp`,
`RedSalamander/SettingsSave.h`, `Common/Common/SettingsStore.cpp`.
**Cleared (checked, sound):** `RunQueuedTaskNow` pointer lifetime & atomics, both
`ExecuteInPane` call signatures, STA/COM usage for `ITaskbarList3`, the conflict-card
height reservation (5 lines reserved == 5 drawn), the single-source destination-path
resolution guard, and the always-on 100 ms `WM_TIMER` that drives the resize animation.
No CLAUDE.md/AGENTS.md convention violations found (RAII/WIL, no C-style casts,
`.value()` not `*`, all UI strings via `LoadStringResource`).

---

## Summary (severity-ranked)

| ID | Status | Severity | One-line | Anchor |
|----|--------|----------|----------|--------|
| R1 | Fixed | High | Compact-density footer toggle is dead to mouse clicks | `Popup.cpp:7126` |
| R2 | Fixed | High | Compact-density preference is dropped before save (lost on restart) | `SettingsSave.h:46` |
| R3 | Fixed | High | Compact density makes the collapse chevron a no-op and hides completed-card actions | `Popup.cpp:4852` |
| R4 | Fixed | Med-High | Footer/taskbar show a determinate 100% while an unknown-size copy is still active | `Popup.cpp:852` |
| R5 | Fixed | Med | Taskbar sticks on red error / "N need attention" for a completed failed task | `Popup.cpp:826` |
| R6 | Fixed | Med (plausible) | Taskbar progress never retries after a transient COM failure | `Popup.cpp:6903` |
| R7 | Fixed | Low-Med (plausible) | "Start now" click can be silently missed (lost wakeup) | `State.Runtime.Part.cpp:913` |
| R8 | Fixed | Low (latent UB) | Aggregate ETA/speed double→uint64 cast is UB for extreme decayed rates | `Popup.cpp:951`, `:8633` |
| R9 | Fixed | Low (cosmetic) | Completed (Ok) status renders the same accent color as Running | `Popup.cpp:716` |
| S1-S5 | Fixed | Cleanup | Efficiency / reuse / altitude (per-frame waste and duplicated blocks) | see §S |
| S6 | Fixed | Cleanup | Two default-check sources for fileOperations settings | see §S |

Remediation evidence:

- R1-R3 fixed by `56e8dc006` with focused Commands coverage:
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_181048/selftest_run_results.json` and
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_181056/selftest_run_results.json`.
- R4-R5 fixed by the live-summary reducer follow-up. Focused coverage:
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_182713/selftest_run_results.json` and
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_182734/selftest_run_results.json`.
- R6 fixed by the taskbar-list retry/backoff follow-up. Build: `.build/logs/msbuild-20260709_183239_666.log`.
  Focused coverage:
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_064734_fileops_popup_taskbar_retry/run-all-tests-results.json`.
- R7 fixed by synchronizing `Task::SetWaitForOthers` with the queue condition-variable mutex.
  Build: `.build/logs/msbuild-20260710_085158_825.log`. Focused coverage:
  `Specs/TestRuns/00013ba62/FileOps/2026-07-10_065718_fileops_popup_start_now_queue_notify/run-all-tests-results.json`,
  `Specs/TestRuns/00013ba62/FileOps/2026-07-10_065731_fileops_queue_mode_switch_notify/run-all-tests-results.json`, and
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_065648_fileops_popup_conflict_taskbar_retry_after_r7/run-all-tests-results.json`.
- R8 fixed by saturating non-negative double-to-`uint64_t` conversions and flooring decayed silent
  rates below 1 B/s before ETA calculation. Build: `.build/logs/msbuild-20260710_090407_996.log`.
  Focused coverage:
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_070736_fileops_popup_global_summary_eta_clamp/run-all-tests-results.json` and
  `Specs/TestRuns/00013ba62/FileOps/2026-07-10_070747_fileops_popup_rate_smoothing_eta_clamp/run-all-tests-results.json`.
- R9 fixed by adding `fileOperations.successText` and resolving `Ok` status tone through that
  success color instead of the active accent. Build: `.build/logs/msbuild-20260710_090958_935.log`.
  Focused coverage:
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_071344_fileops_popup_success_status_color/run-all-tests-results.json`.
- S6 fixed by centralizing the file-operations non-default predicate in `SettingsSave.h` and using
  it from both save preparation and runtime pruning. Build:
  `.build/logs/msbuild-20260710_091518_777.log`. Focused coverage:
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_071852_fileops_settings_shared_default_predicate/run-all-tests-results.json`.

---

## R1 — Compact-density footer toggle is unclickable  *(High)*

**Anchor:** `Popup.cpp:7126-7127` (`HitTest`).

`HitTest`'s footer-button exemption never gained the new `FooterDensity` kind:

```cpp
const bool isFooterButton = it->hit.kind == PopupHitTest::Kind::FooterCancelAll ||
    it->hit.kind == PopupHitTest::Kind::FooterAutoDismiss ||
    it->hit.kind == PopupHitTest::Kind::FooterQueueMode ||
    it->hit.kind == PopupHitTest::Kind::FooterToggleDetails;   // FooterDensity missing
if (! isFooterButton && ! PointInRectF(_listViewportRect, x, y)) { continue; }
```

The density button is laid out in the footer band (`_footerDensityRect`, `footerBtnY =
footerTop + 26dip`), which is **below** `_listViewportRect` (whose bottom is `footerTop`).
So a click on it has `y > footerTop` → `PointInRectF(_listViewportRect,…)` is false, and
because `FooterDensity` is not exempt the loop `continue`s past it. `HitTest` returns
`None`; no `_pressedHit`; `OnActivatedHit(FooterDensity)` never runs.

**Repro:** widen the popup ≥ 560 dip so the density toggle renders, click it → nothing
happens. **Why tests miss it:** `OnSelfTestInvoke` dispatches `OnActivatedHit(FooterDensity)`
directly and `OnLayoutSnapshotRequest` counts it as a visible footer button, so both
bypass hit-testing.

**Fix:** add `|| it->hit.kind == PopupHitTest::Kind::FooterDensity` to `isFooterButton`.

---

## R2 — Compact-density preference is dropped before save  *(High)*

**Anchor:** `SettingsSave.h:46` (`PrepareForSave`), chain via `SettingsHotReload.cpp:190`.

`PrepareForSave`'s `fileOperations` `hasNonDefault` OR-chain was extended with
`popupFooterOnly` but **not** `popupCompactDensity`. Every write goes through
`SaveSettingsAndSchema → SavePreparedSettingsAndSchema → SettingsSave::PrepareForSave`
(`SettingsHotReload.cpp:190`), which resets the whole `fileOperations` block when
`hasNonDefault` is false.

**Repro:** toggle compact density on with every other file-op setting at default →
`hasNonDefault` evaluates false → `result.fileOperations.reset()` → `SaveSettings` writes
nothing → the value parses back as `false` next launch. The in-session value is correct
(the in-memory setter uses the complete `HasNonDefaultFileOperationsSettings`), so this is
disk-persistence only.

**The other four persistence sites all include the field** — `SaveSettings` gate and
per-field write (`Common/Common/SettingsStore.cpp`), `ParseFileOperations`, and
`HasNonDefaultFileOperationsSettings` (`State.cpp`). `PrepareForSave` is the lone omission.

**Fix (see also S6):** add `|| fileOperations.popupCompactDensity != defaults.popupCompactDensity`,
or better, replace the hand-maintained chain with a call to the shared
`HasNonDefaultFileOperationsSettings`.

---

## R3 — Compact density makes the collapse chevron a no-op and hides completed-card actions  *(High)*

**Anchor:** `Popup.cpp:4852` (`isCollapsedTask = isStoredCollapsedTask || compactDensity`);
card-height branches `:4569` / `:4855`; completed-action emission gated on the expanded
branch (`CompletedTaskHasOverflowActions`, `:5764`).

`compactDensity` is OR-ed onto the stored collapse state, so **every** card is forced
collapsed. Unlike auto-collapse (which a manual expand overrides, because
`ToggleTaskCollapsed` sets `_collapsedTasks[id]=false` and `AutoCollapseCompletedTasks`
skips ids already in the map), `compactDensity` ignores `_collapsedTasks` entirely. The
`TaskToggleCollapse` chevron is still drawn and pushed as a button, but clicking it only
flips the ignored map entry — the card never expands.

Because the completed-card **Dismiss** (the primary action) and the new More overflow
(Reveal item / Open destination / Failed items) are emitted only in the expanded branch,
enabling compact density makes those per-task actions **permanently unreachable**, and the
chevron becomes a dead control.

**Fix:** decide the intended behavior — either (a) let per-card collapse state win over
`compactDensity` (compact = default-collapsed, still individually expandable), or (b) keep
the completed-card action row visible in compact rows. (a) matches the auto-collapse model
and is the smaller change. The self-test scaffolding already calls `expandTaskIfCollapsed`
before checking actions, which quietly encodes this gap.

---

## R4 — Aggregate footer/taskbar progress reads determinate 100% while an unknown-size copy is active  *(Med-High)*

**Anchor:** `Popup.cpp:852-865` (`BuildGlobalStatusSummary`); `HasDeterminateGlobalAggregateProgress`;
`BuildGlobalTaskbarProgressModel`.

`BuildGlobalStatusSummary` accumulates `totalBytes`/`completedBytes` over **all**
FileOperation tasks including finished ones — there is no `!finished` guard:

```cpp
if (task.totalBytes > 0) { summary.totalBytes += task.totalBytes;
                           summary.completedBytes += std::min(task.completedBytes, task.totalBytes); }
...
if (! task.finished && task.totalBytes == 0 && task.totalItems == 0 && ...)
    summary.hasUnknownActiveProgress = true;
```

Scenario: file A finished (`totalBytes == completedBytes == N`) plus file B running in
pre-calc (`totalBytes == 0` until pre-calc completes). Then `summary.totalBytes == N > 0`,
so `HasDeterminateGlobalAggregateProgress` is true and `GlobalAggregateProgressFraction`
returns `N/N = 1.0`; `BuildGlobalTaskbarProgressModel` likewise emits `TBPF_NORMAL` with
`completed == total`. The footer bar fills to 100% and the taskbar shows a full/complete
bar while B is still actively working. The `hasUnknownActiveProgress` flag exists precisely
to force the indeterminate presentation here, but neither determinate check consults it.

**Fix:** exclude finished tasks from the byte/item accumulation (or gate the determinate
path on `!hasUnknownActiveProgress`). This also fixes R5.

---

## R5 — Taskbar sticks on red error / "N need attention" for a completed failed task  *(Med)*

**Anchor:** `Popup.cpp:826-850` (`needAttention` / `hasActiveTaskbarProgress`);
`BuildGlobalTaskbarProgressModel` (`needAttention > 0 → TBPF_ERROR`).

Same missing `!finished` guard as R4: a completed-with-errors task (status resolves to
`Failed` or `Partial`) lingers in the completed list until dismissed. Each render
`StatusNeedsAttention(status)` is true, so `++needAttention` and
`hasActiveTaskbarProgress = true` even with nothing active. The taskbar button then paints a
persistent `TBPF_ERROR` (red) bar and the footer summary keeps reading "N need attention"
with zero running/waiting operations.

Distinct symptom from R4 (this is the status-count path, not the byte-sum path) but the
same root: finished tasks are not distinguished. **Fix:** decide whether a finished-with-
errors task should drive the *live* taskbar/attention state at all; most likely it should
be excluded once `finished`, leaving the per-card status to carry the "needs attention"
signal.

---

## R6 — Taskbar progress never retries after a transient COM failure  *(Med, plausible)*

**Anchor:** `Popup.cpp:6903` (`EnsureTaskbarList`).

`EnsureTaskbarList` sets `_taskbarListAttempted = true` before `CoCreateInstance` / `HrInit`
and never clears it on failure; only a fresh `TaskbarButtonCreated` broadcast resets it
(`WndProc`). If `CoCreateInstance` or `HrInit` fails transiently *after* the button became
ready, `_taskbarListAttempted` stays true and `_taskbarList` stays null, so the method
returns false for the rest of the window's lifetime — taskbar progress silently never
appears, and only an Explorer restart (which re-broadcasts `TaskbarButtonCreated`) recovers
it.

**Fix:** on failure, either leave `_taskbarListAttempted` false (retry next tick, bounded)
or add a small retry budget / backoff so one hiccup does not disable the feature
permanently.

**Fixed:** `EnsureTaskbarList` now retries after transient initialization failure with a short
backoff instead of permanently latching failure for the popup lifetime. The debug snapshot exposes
taskbar button/list retry state, and `cmd_pane_fileops_conflict_prompt_compacts_actions` now covers
the forced-failure/retry path.

---

## R7 — "Start now" click can be silently missed (lost wakeup)  *(Low-Med, plausible)*

**Anchor:** `State.Runtime.Part.cpp:913-928` (`RunQueuedTaskNow`); `NotifyQueueChanged`
`:838-845` (`_queueCv.notify_all()`).

`RunQueuedTaskNow` flips the atomic `_waitForOthers` / `_waitingInQueue` and calls
`NotifyQueueChanged()` (`_queueCv.notify_all()`) **without ever holding `_queueMutex`**,
while the queue worker evaluates its wait-predicate on those fields under `_queueMutex`.
Standard condition-variable rules require the predicate state be mutated under the wait
mutex (or the notify otherwise synchronized) to guarantee the wakeup is observed. If the
"Start now" click lands in the narrow window where the worker has evaluated the predicate
false and is committing to sleep, the notify can be lost and the task stays queued until the
next `NotifyQueueChanged` (another task entering/leaving) re-evaluates.

Self-resolving and it mirrors the pre-existing `ApplyQueueMode` pattern, but the entire
point of **Start now** is immediacy, so a missed wakeup reads as "the button did nothing."

**Fix (also fixes the latent `ApplyQueueMode` case):** take `_queueMutex` around the flag
mutation, or notify with the mutex held, so the worker's predicate re-check and the notify
are ordered.

**Fixed:** `Task::SetWaitForOthers` now stores `_waitForOthers` while holding `_queueMutex`
and notifies while the predicate mutation is synchronized with the wait lock. That covers both
popup `Start now` and `ApplyQueueMode` callers through the shared task API.

---

## R8 — Aggregate ETA/speed double→uint64 cast is UB for extreme decayed rates  *(Low — latent UB, not a functional bug)*

**Anchor:** `Popup.cpp:943` / `:951` (`FormatGlobalStatusSummaryText`), `:4401`
(`DrawBandwidthGraph`), `:8633` (debug snapshot, **no clamp at all**);
`DecayRateForCallbackSilence` `:2103`.

The new saturating casts use
`static_cast<uint64_t>(std::min<double>(x + 0.5, static_cast<double>(UINT64_MAX)))`. But
`(double)UINT64_MAX` rounds up to `2^64`, so the clamp ceiling is one past the max and
`static_cast<uint64_t>(2^64)` is out of range → UB. The codebase's own correct idiom
(compare `>=` then assign `UINT64_MAX`) already exists at `Popup.cpp:1150-1157`.

**Reachability — important context.** You cannot reach `2^64` seconds (≈ 584 billion years)
with a *slow* transfer; a real 1 B/s copy of 1 TB is only ~1e12 s, ~7 orders below `2^64`.
It is reachable only through the **unfloored exponential rate decay**:
`DecayRateForCallbackSilence` returns `smoothedRate · exp(−(silenceMs−600)/900)` with no
floor, so during progress-callback silence the denominator in the aggregate ETA
(`remaining / displayedBytesPerSec`) shrinks toward zero and the quotient blows up. For a
100 MB/s transfer with 1 GB remaining, ~38 s of callback silence drives the decayed rate to
~5e-11 B/s and the aggregate ETA past `2^64`. This is realistic on the remote/cloud
back-ends the app ships (Curl / S3 / GoogleDrive / MicrosoftDrive stalling), rare locally.

**True severity.** On x64 an out-of-range `double→uint64` conversion (`cvttsd2si`) yields a
sentinel, not a crash — you'd see a nonsense ETA string for those frames. It is UB per the
standard (and the debug path `:8633` has no clamp), but benign on the target ISA, and the
ETA is already garbage before the cast overflows. **This is the weakest correctness item.**

**Fix (preferred):** floor `DecayRateForCallbackSilence` (and/or the aggregate denominator)
to zero below a small epsilon — this both removes the phantom ETA shown to the user *and*
prevents the overflow. Secondarily, give `:8633` a clamp and switch all sites to the
saturating idiom at `:1150-1157`.

**Fixed:** aggregate, graph-label, per-task speed, per-task ETA, and debug-snapshot conversions now
use saturating non-negative helpers before casting to `uint64_t`. Silent callback decay floors
sub-byte rates to zero, and aggregate/per-task ETA only uses rates of at least 1 B/s.

---

## R9 — Completed (Ok) status renders the same accent color as Running  *(Low, cosmetic)*

**Anchor:** `Popup.cpp:716` (`StatusVisualColorForTone`).

`StatusVisualColorForTone` maps both `Ok` and `Accent` to `theme.accent`:

```cpp
case PopupStatusVisualTone::Accent: return theme.accent;   // Running / Calculating / Preparing
...
case PopupStatusVisualTone::Ok:     return theme.accent;   // Done — same color
```

So a Done task's status stripe and chip are color-indistinguishable from an in-progress
task. The color-blind-safe encoding still holds (glyph + text differentiate), so this is
cosmetic, but a distinct success/positive color was almost certainly intended for `Ok`.

**Fix:** map `Ok` to a success color (e.g. `theme.folderView.successText` or the theme's
positive/green token) rather than `theme.accent`.

**Fixed:** `FileOperationsTheme` now has a `successText` token. The popup resolves
`PopupStatusVisualTone::Ok` and the matching status brush through that token, while high-contrast
themes fall back to menu text for contrast.

---

## §S — Cleanups (efficiency / reuse / altitude)

These are quality items on the changed code, not bugs. Correctness (R1–R9) outranks them.

- **S1 — `IsReducedMotionEnabled` rebuilds a full palette 4×/frame.** `Popup.cpp:2352` calls
  `MakeAppThemeDxPalette(...)` just to read one bool, invoked per-frame from
  `AutoResizeWindow` (`:3679`), `DrawBandwidthGraph` (`:4092`), `DrawFooterQueueModeControl`
  (`:3858`), and `OnLayoutSnapshotRequest` (`:8612`). The cached `_palette` member (set
  `:1376`) already holds `.reducedMotion`; resolve once per `Render`. Also removes the risk
  of two draw methods reading different values mid-frame.
- **S2 — `Render` sweeps the snapshot twice.** `Popup.cpp:4708-4709`:
  `BuildGlobalStatusSummary` then `EnrichGlobalStatusSummaryWithRates` each iterate the full
  task vector every frame (the second also does a `_rates` lookup per task). Fold the rate
  aggregation into the single accumulation loop.
- **S3 — Per-frame `LoadStringResource` for static footer labels.** `Popup.cpp:3809-3811`
  (`DrawFooterQueueModeControl` loads three strings) + `DrawFooterAutoDismissControl` (one)
  run on every 100 ms paint though the strings never change. Cache, invalidate on
  language / `WM_SETTINGCHANGE`.
- **S4 — Three hand-rolled easing state machines.** Auto-resize (`AutoResizeWindow`, ~`:3718`),
  the queue-mode thumb, and the graph latest-point each re-implement smoothstep + tick
  tracking + reduced-motion gating, duplicating the shared DxUi `EvaluateEasing` /
  `EasingCurve` facility. Centralizing keeps popup motion consistent with the rest of the app
  and gives one place for the reduced-motion policy.
- **S5 — Settings save-and-log block copy-pasted 4×.** `State.Runtime.Part.cpp` in
  `SetQueueNewTasks` (`:817`), `SetAutoDismissSuccess` (`:1202`), `SetPopupFooterOnly`
  (`:1266`), `SetPopupCompactDensity` (`:1300`): identical `SaveSettingsAndSchema` + `FAILED`
  + `GetSettingsPath` + `Debug::Error` blocks. A single `SaveFileOpsSettingsOrLog(settings,
  context)` collapses them. Note `SetPopupFooterOnly` saves **unconditionally** (no
  `previous == new` guard), unlike the other two setters — a redundant disk write on no-op.
- **S6 — Fixed: two sources of truth for "is fileOperations default".** This diff centralized the
  check in `HasNonDefaultFileOperationsSettings` (`State.cpp`) but `PrepareForSave`
  (`SettingsSave.h:45`) still hand-maintains a parallel copy — and that copy already drifted,
  which is exactly bug **R2**. Promoting the shared helper to a header used by both call
  sites was the structural fix that prevents R2 from recurring.

Smaller duplication noted but not itemized: the Start-now/Cancel two-button block
(`Popup.cpp` ~`:6559` & ~`:6763`), the hovered/pressed hit-compare (3 inline copies), the
`rectHasArea` lambda (defined twice), the round-bytes-to-uint64 idiom (3 sites), and the
three-slot footer width solver.

**S1-S5 closeout (2026-07-10):** [UI_FileOperationsPopupCodeReviewRemediation_2026-07-10.md](UI_FileOperationsPopupCodeReviewRemediation_2026-07-10.md)
completed the cleanup set. Reduced motion is frame-cached while explicit theme overrides remain
immediate; global status/rates use one live-task scan; static footer strings refresh on create/theme/
setting changes; equivalent motion paths share `EaseFileOperationsUiMotionFraction`; and settings
saves use `SaveFileOperationsSettingsOrLog` with no-op guards. The shared non-default settings
predicate moved to `Common::Settings`, completing S6 at the serialization boundary as well.

Focused 3-repeat Commands/FileOps coverage and the full 102-case FileOps suite are recorded in the
remediation plan. No cleanup item remains independently executable from this tracker.
