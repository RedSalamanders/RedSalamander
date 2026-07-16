# Operation Clearview — File Operations Progress Popup UX Refinement

**Status:** WIP — A-tier and selected polish/workflow items are implemented; the review-remediation evidence gate is complete, while the remaining product backlog keeps this plan in WIP
**Date:** 2026-07-07
**Scope owner:** the File Operations progress popup surface —
`RedSalamander/FolderWindow.FileOperations.Popup.cpp` (~8156 lines) + `.Popup.h`,
its resource strings (`RedSalamander/RedSalamander.rc` `IDS_FILEOPS_*`, ~1585–1665), and the
narrow slice of the engine it reads (`FolderWindow.FileOperations.State.Runtime.Part.cpp`
queue mode). The copy/move/delete *engine* itself (bytes, verification, scheduling correctness) is
out of scope and owned elsewhere (Causeway, Firebreak, Floodgate).
**Companion docs:** seeded UI spec
`Specs/UI/UI_FileOperationsPopup.md`, plus the authoritative engine/UI contract in
`Specs/FileSystem/FileSystem_FileOperations.md`. DxUi widget facts below are
drawn from a `Common/DxUi/DxUi.h` inventory pass and are cited inline.
**Provenance:** direct code walk of the popup render/hit-test/scheduler paths on `master` at
`cb2478689`, plus a DxUi toolkit survey. Every anchor below was read, not inferred.

**2026-07-11 review-remediation status:** the scoped correctness/architecture review is complete at
[UI_FileOperationsPopupCodeReviewRemediation_2026-07-10.md](../Done/UI_FileOperationsPopupCodeReviewRemediation_2026-07-10.md),
with replacement same-machine baseline/candidate evidence and a classified Full-suite blocker
archived and audited. Its implementation tracks remain in code: conflict metadata no
longer blocks the arbiter lock or reads content; completed destination actions retain qualified
provider identity; minimized taskbar updates, mixed-known/unknown ETA, reduced-motion marquees,
minimum-width footer hit targets, and expanded-placement persistence are covered. Static footer
resources, global summary/rate scans, easing, settings-save blocks, and the Common settings predicate
were consolidated. The real product capture is
`Specs/UI/Images/FileOperationsPopup_Product_Conflict_2026-07-10.png`. The explicit product backlog
below remains open (B1 reorder, B4 group animation, B5
keyboard/UIA, C5 Keep Both engine support, and C7 shared DxUi graph support).

---

## 0. Why this plan exists (the two user asks, in the code)

1. **"The Wait/Parallel button is unclear."** The footer control is a *plain text button* whose
   caption is the **current** mode, not the action. `DrawButton(queueBtn, …, modeText)` with
   `modeText = queueMode ? IDS_FILEOPS_BTN_MODE_QUEUE ("Wait") : IDS_FILEOPS_BTN_MODE_PARALLEL
   ("Parallel")` (`Popup.cpp:4011-4014`); clicking flips it via
   `fileOps->ApplyQueueMode(!queueMode)` (`Popup.cpp:7165-7175`). A lone button reading **"Wait"**
   is genuinely ambiguous — it can be read as *"you are waiting"* (state) or *"click to wait"*
   (action). Neither the current value nor the toggle nature is legible. → **A1**.

2. **"We now show progress near the file name, so the bar under the graph is redundant, and the
   global one should be reviewed."** Confirmed. There are **two** per-file progress renderings:
   - the **mini-bar beside the current file name** (`Popup.cpp:5150-5178`), drawn to the right of
     `currentSourcePath`, filled by `fileCompletedBytes / fileTotalBytes` of the in-flight file;
   - the **item bar under the graph** (`Popup.cpp:5490-5524`), filled by
     `task.itemCompletedBytes / task.itemTotalBytes` — i.e. **the same per-file byte fraction**.
     Directly below it sits a distinct **total bar** (`:5526-5543`, whole-task
     `completedBytes / totalBytes`). → the item bar is pure duplication; the total bar is the only
     unique signal there. → **A2**.
   - The **global** presentation is a footer *text line only* —
     `IDS_FMT_FILEOPS_GLOBAL_STATUS_SUMMARY "{0} running, {1} waiting, {2} need attention"`
     (`Popup.cpp:4016-4026`, built by `BuildGlobalStatusSummary` `:649-675`). There is **no**
     aggregate progress bar and **no** Windows taskbar progress. → **A3**.

---

## 1. Current anatomy (ground truth — read before touching anything)

**Window & chrome.** The popup is a hand-rolled `WndProc` window rendering into an
`ID2D1HwndRenderTarget` on a 100 ms `WM_TIMER` (`kFileOperationsPopupTimerIntervalMs = 100`,
`Popup.cpp:27`); it is **not** a DxUi control tree. ~15 D2D brushes are managed by hand
(`:2365-2473`). Buttons are `PopupButton{ D2D1_RECT_F bounds; PopupHitTest hit; }` (`Popup.h:54-58`)
rebuilt into `_buttons` every frame; clicks dispatch by `PopupHitTest::Kind`
(`Popup.h:15-40`: `FooterCancelAll`, `FooterQueueMode`, `TaskToggleCollapse`, `TaskPause`,
`TaskCancel`, `TaskSkip`, `TaskDestination`, `TaskSpeedLimit`, `TaskShowLog`, `TaskExportIssues`,
`TaskCompletedMore`, `TaskConflict*`, `TaskDismiss`).

**Footer (`:3991-4026`).** 44 dip tall, top border stroke. Holds: the **Cancel all / Clear
completed** button (`_footerCancelAllRect`, label swaps on `HasActiveOperations()`), the
**Wait/Parallel** button (`_footerQueueModeRect`), and the **global summary text**
(`_footerSummaryRect`, small font, sub-text brush). Auto-dismiss is a *preference*
(`IDS_PREFS_FILEOPS_AUTODISMISS_*`), not a footer control.

**Per-task card, running (`:4907-5545`).** Rounded rect + border (`:4043`), then top-to-bottom:
- **Header** (`:4750`): status glyph (`CaptionStatus` Ok/Warning/Error → Fluent check/warning/error,
  `:4705-4747`) + `BuildTaskHeaderText` (`:611-624`) = `"{op}: {done}/{total}"`
  (`IDS_FMT_FILEOPS_OP_COUNTS`) while running, else `"{op}: {status}"`. A **collapse chevron**
  sits at the right (`:4678-4684`, `TaskToggleCollapse`).
- **Speed / size text** (`:5000-5017`): `"{bytes}/s"` + `"{completedBytes} / {totalBytes}"`
  (`IDS_FMT_FILEOPS_SIZE_PROGRESS`) + ETA (`IDS_FMT_FILEOPS_ETA`).
- **Current file line(s)** (`:5100-5178`): `From:` path (middle-truncated) + the **mini progress
  bar** for the in-flight file (or per-file entries from `task.inFlightFiles`).
- **Bandwidth graph** (`:5372-5425` → `DrawBandwidthGraph` `:3372`): sparkline of throughput with a
  speed-limit reference line, plus a status overlay (`GraphOverlayTextForStatus` `:626-647`:
  "Waiting for previous to complete", "Paused", "Calculating…", "Preparing…", "Needs attention").
- **Under-graph bars** (`:5427-5545`): one of three shapes —
  marquee during pre-calc (`:5427-5442`), a single total bar (`:5451-5485`), or the
  **item-bar + total-bar pair** (`:5490-5543`).
- **Button row** (`barsTop`/`barsBottom` layout `:5240-5253`): Pause/Resume, Cancel, Speed-limit
  (menu), destination, etc.

**Per-task card, finished (`:4770-4904`).** Warnings/errors counts, optional HRESULT line, From/To,
a **completion bar** (`ComputeFileOperationsTaskCompleteFractionForDisplay`, `:4846-4865`), and
**Dismiss** / **More…** buttons (the latter only when `warningCount|errorCount > 0`).

**Informational tasks (`:4057-4629`).** Batch-rename / change-attributes progress; own header,
counts, marquee/determinate bar, Dismiss.

**Test-observation surface.** The popup exposes a debug snapshot **without pixels**:
`OnLayoutSnapshotRequest` → `PopupLayoutDebugSnapshot` (`Popup.cpp:7411-7500+`) reports
`globalRunningCount/globalWaitingCount/globalNeedAttentionCount/globalSummaryText/globalSummaryVisible`,
`visibleButtonCount`, `footerVisibleButtonCount` (counts `FooterCancelAll` + `FooterQueueMode`,
`:7491-7492`), conflict action layout, and graph-hue aggregates. **This is the hook every popup
self-test must use.** Any control we add/remove must keep this snapshot meaningful (extend it, do
not orphan it).

**Accessibility today (verified gap).** The popup has **no keyboard handling and no UIA** — zero
`WM_KEYDOWN`/`VK_`/`OnKeyDown` and zero `WM_GETOBJECT`/`IRawElementProvider` in `Popup.cpp`. It is
mouse-only and invisible to Narrator. (Contrast: real DxUi controls get both for free.) → **B5**.

---

## 1.5 Visual reference — before → after

ASCII mockups (monospace). Circled numbers ①②③ are the three progress renderings discussed in §0.2.
Boxes are open on the right so the art stays legible regardless of terminal width.

**CURRENT — one running task + footer**

```text
┌──────────────────────────────────────────────────────────────
│ ⚠ Copy: 3 / 10                                        [ ∧ ]     header: op + item count · collapse chevron
│ 240 MB/s                                                        speed
│ 1.2 GB / 5.0 GB                        Remaining: 3 min         size · ETA
│ From: …\Photos\IMG_0421.CR2          [▓▓▓▓▓▓░░░]  ← ①           current file + PER-FILE mini-bar
│ ┌────────────────────────────────────────────────┐
│ │    ╲  ╱╲╱  ╲  ╱╲╱  ╲╱╲     throughput sparkline │             bandwidth graph
│ └────────────────────────────────────────────────┘
│ [▓▓▓▓▓▓░░░░░░░░░░░░]  ② item bar   ← DUPLICATE of ①             ⟵ REDUNDANT
│ [▓▓▓▓▓░░░░░░░░░░░░░]  ③ total bar  (whole task)
│ [ Pause ]  [ Cancel ]          Speed: [ 50 MB/s ▾ ]            action row (speed limit = menu)
└──────────────────────────────────────────────────────────────
   [ Cancel all ]   [ Wait ]         2 running, 1 waiting, 0 need attention
                      ▲                  ▲
             ambiguous — is this      text only: no aggregate bar,
             my state or a button?    no Windows taskbar progress
```

**PROPOSED — same task, refined (A1·A2·A3·B2·C1·C2)**

```text
┌──────────────────────────────────────────────────────────────
│▎⚠ Copy · 3 / 10                                       [ ∧ ]     ▎ accent stripe by status (B2) + header
│▎▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔ 52%   overall                        ③ total bar → thin header underline (A2)
│▎240 MB/s     1.2 GB / 5.0 GB           Remaining: 3 min         speed · size · ETA (compact, one line)
│▎From: …\Photos\IMG_0421.CR2          [▓▓▓▓▓▓░░░] 68%  ← ①       the ONLY per-file bar now
│▎┌────────────────────────────────────────────────┐
│▎│    ╲  ╱╲╱  ╲  ╱╲╱  ╲╱╲     throughput sparkline │             graph (collapsible — C3)
│▎└────────────────────────────────────────────────┘
│▎[ ⏸ Pause ] [ ✕ Cancel ]     Speed: [ 50 MB/s ▾ ]                 one menu button; menu lists presets/custom (C2)
└──────────────────────────────────────────────────────────────
   [▓▓▓▓▓▓▓▓▓░░░░░░░]   ≈ 3 min · 240 MB/s          aggregate bar (A3) + combined ETA (C1)
   [ Cancel all ]    New tasks: [ Queue │ Parallel ]      2 running · 1 waiting
                                 ▲ segmented — both options shown, selection filled (A1)
   Footer-only saved state:
   [▓▓▓▓▓▓▓▓▓░░░░░░░]   ≈ 3 min · 240 MB/s   [Show details] [Cancel all] New tasks:[Queue│Parallel]
```

---

## 2. Scheduler semantics the UX must respect

- **Queue mode is a single global boolean**, `_queueNewTasks` (default **true**), not per-task.
  `ShouldQueueNewTask()` = `_queueNewTasks && HasActiveOperations()`
  (`State.Runtime.Part.cpp:862-870`). A new task started while another is active begins
  `WaitForOthers`.
- `ApplyQueueMode(queue)` (`State.Runtime.Part.cpp:882-911`): stores the flag, then for every task —
  if switching to **parallel**, clears `WaitForOthers` on all (releases the queue); if switching to
  **queue**, sets `WaitForOthers` on every **not-yet-started** task. Then
  `UpdateQueuePausedTasks()` + `NotifyQueueChanged()`. **Mode is changeable mid-flight.**
- Per-task hold already exists: `Task::SetWaitForOthers(bool)`, `IsWaitingForOthers()`,
  `IsWaitingInQueue()` (`FileOperationsInternal.h:351-352`, `State.cpp:6271-6279`). This is the
  primitive a "Run this next" / "Start now" control (B1) would drive — clearing `WaitForOthers` on
  one task without flipping the global flag.
- Auto-concurrency (`IDS_FMT_FILEOPS_AUTO_CONCURRENCY`) and per-plugin worker counts are resolved
  in the engine; the popup only displays outcomes. **Do not** surface concurrency knobs here — they
  live in Preferences → Plugins → File System (`IDS_PREFS_FILEOPS_PLUGIN_HINT_*`).
- **Constraint for A1:** the toggle governs *new* tasks and *releases/holds unstarted* ones. Its two
  states are "queue new tasks behind the running one" vs "run everything in parallel". Label both
  states in those terms — not the bare words "Wait"/"Parallel".

---

## 3. Work ledger

Priorities: **A** = user-requested (do first, in order). **B** = high-value workflow.
**C** = appeal / polish. Effort: S ≈ hours, M ≈ 1–3 days, L ≈ week+.

State meanings: **Done** = implemented and has at least targeted regression coverage planned or
already present; **In progress** = code exists but verification/closeout or a sub-requirement is
still open; **Not done** = no meaningful implementation in this plan yet.

| State | ID | Item | Current status | Pri | Effort |
|-------|----|------|----------------|-----|--------|
| Done | A1 | Replace the Wait/Parallel text button with a labeled toggle (state-legible) | Footer now uses `New tasks: [ Queue | Parallel ]` segmented control, keeps `FooterQueueMode`, and has an animated subdued accent thumb. | A | M |
| Done | A2 | Drop the redundant under-graph item bar; consolidate to one meaningful bar | Duplicate per-file under-graph item bar removed; snapshot tracks one under-graph task progress bar and `taskDuplicateUnderGraphItemBarVisible=false`. | A | S |
| Done | A3 | Rework the global presentation: aggregate progress + Windows taskbar progress | Footer aggregate progress bar, aggregate ETA/throughput text, debug taskbar model, and `ITaskbarList3` taskbar progress application are implemented. | A | M |
| In progress | B1 | Per-task queue control: "Start now" / reorder waiting tasks; Pause-all/Resume-all | Queued-card Start now and footer Pause all/Resume all are implemented and covered. Reorder remains open because it needs an engine queue-order key. | B | M |
| Done | B2 | Status-at-a-glance: colored accent stripe + consistent state chip per card | File-operation cards now draw a status-tone stripe plus localized status chip, and `PopupLayoutDebugSnapshot` exposes stripe/chip/tone fields for tests. | B | S |
| Done | B3 | Post-completion actions: Open destination / Reveal / jump to Failed Items | Completed copy/move cards expose Open destination and Reveal item from More when destination data is available; completed diagnostic cards also expose Failed Items and open the existing Failed Items pane. | B | M |
| In progress | B4 | Auto-collapse & group completed tasks; compact vs expanded density | Completed cards now auto-collapse on first finished snapshot; the footer exposes a persisted, hit-testable Compact/Expanded density toggle, compact rows can be expanded per card, and `Completed (N)` grouping is implemented. Animated expand/collapse remains open. | B | M |
| Not done | B5 | Keyboard navigation + UIA/LiveRegion for the whole popup | Popup is still hand-drawn mouse-first; no UIA/keyboard model landed. | B | L |
| Done | C1 | Aggregate ETA + throughput in the footer/header | Footer summary now appends aggregate smoothed throughput and aggregate ETA when totals/rates allow it. | C | S |
| Done | C2 | Speed-limit menu button: current value + preset/custom flyout | Kept the one menu button showing current limit; flyout/prompt path remains the intended control shape. | C | S |
| Done | C3 | Collapsible details/footer-only mode, remember popup size/position/density | Footer-only state and Compact/Expanded density are persisted; collapse right-aligns the chevron, shrinks to the footer minimum, and expand restores the captured pre-collapse window rect. | C | M |
| Done | C4 | Reduced-motion + high-contrast + color-blind-safe status encoding | Reduced-motion gates popup motion; high-contrast keeps card stripe/chip/tone semantics while suppressing the caption glyph, and status encoding now has glyph/text snapshot coverage so it is not color-only. | C | S |
| In progress | C5 | Conflict prompt enrichment: "Keep both", stacked source/destination compare | Stacked full-width Source/Destination rows now include right-aligned size/modified metadata when available. Keep Both still needs engine-backed destination rename support. | C | M |
| In progress | C6 | Motion polish: animated throughput transitions, progress easing, collapse/expand motion | Graph latest-point easing, queue-thumb animation, and debounced eased auto-resize are implemented; the same paths now honor reduced motion. Broader animation QA remains open. | C | S |
| In progress | C7 | Rainbow throughput graph mode + shared DxUi graph/gradient support | Local graph supports rainbow/per-stream bands. Shared DxUi graph primitive is not implemented. | C | M |

### 3.0.1 Current checkpoint

**Done in code in this pass**
- Footer details chevron moved to the far right; footer-only collapse now targets the footer minimum
  height and expansion restores the captured pre-collapse rectangle.
- Auto-dismiss is now visible in the popup footer, toggles the existing setting, persists it, and
  removes already auto-dismissable completed entries when enabled.
- Queue/Parallel control uses a calmer animated accent thumb instead of the bright progress fill.
- Auto-resize is debounced and eased to reduce card/window dancing as parallel file rows appear and
  disappear.
- Bandwidth graph now draws a current smoothed effective-throughput marker and text label; the
  configured speed limit stays in the speed-limit menu button, not the graph marker.
- Footer summary now appends aggregate smoothed throughput and aggregate ETA when current task rates
  and totals make that meaningful.
- Windows taskbar progress is modeled from the aggregate footer state and applied through
  `ITaskbarList3` after the taskbar button is available.
- File-operation cards now include a status-tone stripe plus localized header chip, both derived
  from the single resolved `TaskSnapshot::StatusKind`.
- Queued file-operation cards now expose a localized `Start now` action that clears the selected
  task's wait gate without changing the global new-task Queue/Parallel setting.
- Queue admission now clears a task's stale `waitingForOthers` flag when the queued task naturally
  enters operation, so started tasks no longer keep a Waiting status from their former queue gate.
- Reduced-motion now gates auto-resize easing, Queue/Parallel thumb animation, graph status
  animation, and graph latest-point easing through the existing app theme/DxUi palette state.
- High-contrast/color-blind-safe status encoding now has explicit snapshot coverage: the card keeps
  stripe/chip/tone semantics plus glyph/text signals, while high contrast suppresses the non-client
  caption status glyph.
- `PopupLayoutDebugSnapshot` was extended for auto-dismiss, right-aligned footer chevron, animated
  resize, reduced-motion animation enablement, high-contrast/color-blind-safe status signals,
  current-bandwidth marker, aggregate ETA/throughput, taskbar progress, and status stripe/chip/Start
  now assertions.
- `Specs/UI/UI_FileOperationsPopup.md` was seeded, and
  `Specs/FileSystem/FileSystem_FileOperations.md` was updated to match the current footer/card
  contract.
- Debug x64 build passed after this implementation pass.

**Focused verification passed**
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully with MSBuild 18.6.3.
- Commands `settings_file_operations_precalc_roundtrip`: passed, archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-09_145554_fileops_popup_ux_settings_precalc_roundtrip/run-all-tests-results.json`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed, archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-09_145557_fileops_popup_ux_conflict_prompt_popup_ux/run-all-tests-results.json`.
- FileOps `Phase5_PreCalcCancelReleasesSlot`: passed, archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-09_145601_fileops_popup_ux_phase5_precalc_popup_ux/run-all-tests-results.json`.
- FileOps `Phase6_PopupRateSmoothing`: passed, archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-09_145605_fileops_popup_ux_phase6_rate_smoothing_popup_ux/run-all-tests-results.json`.
- FileOps `Phase5_PreCalcCancelReleasesSlot`: passed with queued Start now visibility/activation
  coverage, archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-09_155250_fileops_popup_start_now_phase5_precalc_cancel/run-all-tests-results.json`.
- FileOps `Phase5_PreCalcSkipContinues`: passed after deterministic queue-state assertion cleanup,
  archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-09_155316_fileops_popup_start_now_phase5_precalc_skip/run-all-tests-results.json`.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully with 0 warnings and 0 errors,
  log `.build/logs/msbuild-20260709_160812_803.log`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with reduced-motion popup
  snapshot coverage, archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-09_161248_fileops_popup_reduced_motion_conflict_prompt/run-all-tests-results.json`.
- FileOps `Phase5_PreCalcCancelReleasesSlot`: passed with live pre-calc reduced-motion popup
  snapshot coverage, archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-09_161203_fileops_popup_reduced_motion_phase5_precalc_cancel/run-all-tests-results.json`.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully with 0 warnings and 0 errors,
  log `.build/logs/msbuild-20260709_161614_670.log`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with high-contrast and
  color-blind-safe status snapshot coverage, archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-09_162008_fileops_popup_high_contrast_status/run-all-tests-results.json`.
- Real product screenshots captured from the Debug x64 File Operations popup:
  `Specs/UI/Images/FileOperationsPopup_Product_Active_2026-07-09.png` and
  `Specs/UI/Images/FileOperationsPopup_Product_Partial_2026-07-09.png`.
- Completed diagnostic cards now advertise a Failed Items action in the More menu and route it to
  the real File Operations Failed Items pane.
- Completed copy/move cards now advertise Open destination and Reveal item in the More menu when
  destination data is available; Open destination navigates the destination pane to the completed
  destination folder, and Reveal item navigates to the parent folder and selects the completed item.
- Completed cards now auto-collapse once when they first resolve to a finished state, preserving
  auto-dismiss behavior and manual expand/collapse overrides.
- The footer now exposes a persisted Compact/Expanded density toggle backed by
  `fileOperations.popupCompactDensity`. Compact density renders each task as a one-line row with a
  mini progress meter and localized percent text when progress is known.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully with 0 warnings and 0 errors,
  log `.build/logs/msbuild-20260709_162855_448.log`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with completed-task Failed
  Items menu-action coverage, archived at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_163338/run-all-tests-results.json`.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully with 0 warnings and 0 errors,
  log `.build/logs/msbuild-20260709_164855_068.log`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with completed-task Open
  destination, Reveal item, and Failed Items menu-action coverage, archived at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_165348/run-all-tests-results.json`.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the compact-density
  implementation, log `.build/logs/msbuild-20260709_172523_435.log`.
- Commands `settings_file_operations_precalc_roundtrip`: passed with persisted compact-density settings
  coverage, archived at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_173115/run-all-tests-results.json`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with footer density toggle,
  compact row, completed auto-collapse, and compact-row progress coverage, archived at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_173221/run-all-tests-results.json`.
- Conflict prompts now capture source/destination `IFileSystemIO` metadata when the prompt opens and
  render compact size/modified metadata on the right side of each stacked Source/Destination label row.
- `PopupLayoutDebugSnapshot` now distinguishes global popup button overlap from selected-task button
  overlap, so completed-card assertions are not polluted by unrelated visible controls.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the conflict metadata
  implementation, log `.build/logs/msbuild-20260709_175103_301.log`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with conflict source/destination
  metadata and size/date compare coverage, archived at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_175512/run-all-tests-results.json`.
- Pre-push review follow-up R1-R3: the footer density toggle is included in footer hit testing,
  `popupCompactDensity` survives canonical save when it is the only non-default file-operation
  setting, and compact density now behaves as a default display state instead of forcing permanent
  collapsed cards; per-card chevrons can expand compact rows without suppressing later completed-card
  auto-collapse.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the density
  follow-up fixes, log `.build/logs/msbuild-20260709_180644_072.log`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with footer density
  hit-target and compact-row expandability coverage, archived at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_181048/selftest_run_results.json`.
- Commands `settings_file_operations_precalc_roundtrip`: passed with compact-density-only canonical
  save coverage, archived at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_181056/selftest_run_results.json`.
- Pre-push review follow-up R4-R5: the global footer/taskbar reducer now ignores finished cards for
  live counters, aggregate totals, need-attention totals, and taskbar progress. Unknown live progress
  forces indeterminate aggregate/taskbar state even when another active task has known totals.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the live-summary
  follow-up fixes, log `.build/logs/msbuild-20260709_182310_776.log`.
- Commands `cmd_pane_fileops_popup_global_summary_ignores_finished_tasks`: passed with deterministic
  reducer coverage for completed-only, completed+unknown, known+unknown, and completed+conflict
  scenarios, archived at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_182713/selftest_run_results.json`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with live popup coverage that
  completed diagnostic cards no longer keep live footer/taskbar attention active, archived at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-09_182734/selftest_run_results.json`.
- Archive checkpoint for R6 review follow-up: taskbar-list initialization now uses a retry-after
  backoff instead of a permanent failed latch, `PopupLayoutDebugSnapshot` exposes taskbar
  button/list retry state, and focused retry coverage is drafted in
  `cmd_pane_fileops_conflict_prompt_compacts_actions`. Debug x64 build passed with 0 warnings and
  0 errors, log `.build/logs/msbuild-20260709_183239_666.log`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with forced taskbar-list
  failure/retry coverage, archived at
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_064734_fileops_popup_taskbar_retry/run-all-tests-results.json`.
- Pre-push review follow-up R7: `Task::SetWaitForOthers` now mutates the queue wait predicate under
  `_queueMutex` and notifies through the synchronized queue path, covering both popup Start now and
  `ApplyQueueMode`.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the R7 queue
  synchronization fix with 0 warnings and 0 errors, log `.build/logs/msbuild-20260710_085158_825.log`.
- FileOps `Phase5_PreCalcCancelReleasesSlot`: passed with queued Start now release coverage,
  archived at
  `Specs/TestRuns/00013ba62/FileOps/2026-07-10_065718_fileops_popup_start_now_queue_notify/run-all-tests-results.json`.
- FileOps `Phase5_Switch`: passed with queue-mode wait-to-parallel transition coverage, archived at
  `Specs/TestRuns/00013ba62/FileOps/2026-07-10_065731_fileops_queue_mode_switch_notify/run-all-tests-results.json`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed on the current build after the
  R7 synchronization fix, archived at
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_065648_fileops_popup_conflict_taskbar_retry_after_r7/run-all-tests-results.json`.
- Pre-push review follow-up R8: rate/ETA display paths now use saturating non-negative double-to-uint64
  conversions, silent callback decay floors sub-byte rates to zero, and aggregate/per-task ETA ignores
  rates below 1 B/s.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the R8 rate/ETA
  clamp fix with 0 warnings and 0 errors, log `.build/logs/msbuild-20260710_090407_996.log`.
- Commands `cmd_pane_fileops_popup_global_summary_ignores_finished_tasks`: passed with deterministic
  extreme ETA clamp and silent-rate floor coverage, archived at
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_070736_fileops_popup_global_summary_eta_clamp/run-all-tests-results.json`.
- FileOps `Phase6_PopupRateSmoothing`: passed with live popup rate/ETA smoothing coverage, archived at
  `Specs/TestRuns/00013ba62/FileOps/2026-07-10_070747_fileops_popup_rate_smoothing_eta_clamp/run-all-tests-results.json`.
- Pre-push review follow-up R9: completed-success status now resolves through the scoped
  `fileOperations.successText` theme color instead of reusing the active accent color.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the R9 success-color
  fix with 0 warnings and 0 errors, log `.build/logs/msbuild-20260710_090958_935.log`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with completed-success Ok tone
  color coverage, archived at
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_071344_fileops_popup_success_status_color/run-all-tests-results.json`.
- Pre-push review follow-up S6: file-operation settings default detection is centralized in
  `SettingsSave::HasNonDefaultFileOperationsSettings`, and runtime pruning now uses the same helper
  as canonical save preparation.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the S6 settings
  cleanup with 0 warnings and 0 errors, log `.build/logs/msbuild-20260710_091518_777.log`.
- Commands `settings_file_operations_precalc_roundtrip`: passed with shared non-default predicate
  coverage for file-operation settings save/prune behavior, archived at
  `Specs/TestRuns/00013ba62/Commands/2026-07-10_071852_fileops_settings_shared_default_predicate/run-all-tests-results.json`.
- B1 bulk pause/resume follow-up: the footer now exposes `Pause all` while any started live task is
  unpaused and `Resume all` when the started live tasks are paused. The action calls
  `FileOperationState::SetAllRunningTasksPaused(...)`, updates only started tasks' manual pause
  state, and does not change queue gating or the global Queue/Parallel mode.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the B1 bulk
  pause/resume change with 0 warnings and 0 errors, log `.build/logs/msbuild-20260710_093016_902.log`.
- FileOps `Phase6_PopupSmokeResizeAndPause`: passed with the real footer Pause all / Resume all
  activation path, archived at
  `Specs/TestRuns/8c8d65bd9/FileOps/2026-07-10_073348_fileops_popup_footer_pause_resume_all/run-all-tests-results.json`.
- B4 completed-group follow-up: the popup now inserts a localized `Completed (N)` group when at
  least two completed file-operation cards remain visible. The group defaults expanded, its chevron
  collapses/hides or expands the grouped rows without dismissing them, and its `Clear` action
  dismisses those completed summaries. `PopupLayoutDebugSnapshot` reports group visibility, expanded
  state, counts, toggle/clear visibility, and per-task hidden-by-group state. Animated expand/collapse
  remains open.
- Debug x64 build: `RedSalamander/RedSalamander.vcxproj` built successfully after the completed-group
  change with 0 warnings and 0 errors, log `.build/logs/msbuild-20260710_094127_552.log`.
- Commands `cmd_pane_fileops_conflict_prompt_compacts_actions`: passed with completed-group collapse,
  expand, count, and selected-task hidden-state coverage, archived at
  `Specs/TestRuns/a85cf77e9/Commands/2026-07-10_074513_fileops_completed_grouping/run-all-tests-results.json`.

**In progress before Done closeout**
- Re-check the in-app popup visually after the focused tests, especially minimum-width footer layout and
  the compact density button label/icon treatment.
- Decide whether the remaining B/C backlog stays in this plan or is split into follow-up WIP plans
  before this file can move to Done.

**Still not done in this plan**
- Per-task queue reorder controls; do not fake this in the view until the engine owns an explicit
  queue-order key.
- Animated expand/collapse polish for completed group/card density.
- Keyboard navigation and UIA/LiveRegion support.
- Keep Both engine-backed destination rename support.
- Shared DxUi graph primitive for the local rainbow/per-stream graph behavior.

---

## 3.1 2026-07-09 implementation review follow-up

The first implementation/mockup pass exposed several acceptance gaps that must be closed before
this plan can move to `Specs/Plans/Done/`:

- **Footer-only chevron placement and resize.** The footer details chevron belongs on the far right,
  matching the task-card chevron position. Collapsing all details must resize the popup to the
  minimum footer-only height; expanding must restore the pre-collapse window rectangle when one was
  captured, then fall back to normal auto-size only if no prior rectangle exists. The footer-only
  state remains saved in settings.
- **Footer controls need calmer motion and layout.** `Clear completed` / `Cancel all`,
  auto-dismiss, `New tasks: Queue | Parallel`, the global summary, and the collapse chevron must
  have stable hit targets with no clipped summary text. The Queue/Parallel selected segment should
  animate between states and must not reuse the large green progress fill as its selected thumb.
- **Auto-dismiss must be visible.** The existing auto-dismiss setting is not discoverable in the
  popup. Add a footer control that toggles the same persisted setting and immediately removes
  already auto-dismissable completed entries when enabled.
- **Card/window resize must stop dancing.** When parallel in-flight files appear/disappear and card
  height changes, debounce the auto-resize briefly and ease the height transition instead of
  snapping on every frame.
- **Bandwidth graph marker semantics.** The horizontal graph marker and label must represent the
  current smoothed effective throughput, with animated/smoothed updates. It must not be the
  configured speed limit; the configured limit remains visible through the speed-limit menu button.
- **Conflict paths stay stacked.** Long source/destination conflict paths must remain label + full
  width path rows, not side-by-side columns.
- **Rainbow mode / DxUi follow-up.** The hand-drawn popup may use the app's rainbow mode for graph
  hue bands, but any future shared DxUi graph helper should expose the same primitives: smoothed
  current-rate marker, optional rainbow/per-stream bands, reduced-motion/high-contrast fallbacks,
  and deterministic debug snapshots.

---

## Tier A — user-requested

### A1 — Wait/Parallel → a state-legible toggle

**Problem.** `Popup.cpp:4011-4014` paints one button whose caption is the current mode; the control
neither reads as a toggle nor tells you what a click does. `IDS_FILEOPS_BTN_MODE_QUEUE` = "Wait",
`IDS_FILEOPS_BTN_MODE_PARALLEL` = "Parallel" (`RedSalamander.rc:1593-1594`; satellites at each
`Lang/*/…rc:1300-1301`).

**Proposal.** Present the choice as an explicit two-state control with a **caption + two labeled
states**, so both the current value and the alternative are visible at rest. Two viable shapes:

- **Preferred — segmented pill** `New tasks:  [ Queue │ Parallel ]`, the selected segment filled
  with `accent`/`selectionFill`. This shows *both* options and the current selection at once (least
  ambiguous), matches the WinUI segmented look, and reads left-to-right.
- **Alternative — DxUi ToggleSwitch** with a leading caption `Run new tasks in parallel [○/●]`.
  `DxUi::Toggle` already exists (`DxUi.h:1655-1680`, impl `DxUi.Controls.cpp:2395-2560`) with
  `SetStateLabels(off,on)`, `SetOnToggled`, animated knob, keyboard, and **UIA TogglePattern**
  (`DxUi.Accessibility.cpp:2840,2878`). A boolean switch maps cleanly to the boolean
  `GetQueueNewTasks()`.

```text
   TODAY (ambiguous)              OPTION 1 — segmented pill (preferred)
   ┌────────┐                     New tasks:  ┏━━━━━━━━┓──────────┐
   │  Wait  │  ← one word,                    ┃ Queue  ┃ Parallel │
   └────────┘    no context                   ┗━━━━━━━━┛──────────┘
                                                  ▲ selected segment filled (accent)
   caption flips to "Parallel"    both choices + current selection visible at rest
   after click — you never see
   both, and can't tell the       OPTION 2 — labelled toggle switch
   button from a status label     Run new tasks in parallel   ●━━○   ← off = Queue
                                   Run new tasks in parallel   ○━━●   ← on  = Parallel
```


**Implementation.** The popup is hand-drawn, so there are two integration routes; pick one and note
the trade-off in the UI spec:
- **(a) Hand-drawn chrome (lowest risk, keeps the surface uniform).** Add a
  `PopupHitTest::Kind::FooterQueueMode` **segmented** renderer next to the existing
  `DrawDxUiButtonChrome` bridge (`Popup.cpp:3242-3271`), using `ButtonChromeCustomStyle`
  (`DxUi.h:843-858`, custom `cornerRadiusDip` → pill) and `ResolveToggleVisualStyle`
  (`DxUi.h:3261-3271`) for the selected fill. **Keep the hit kind `FooterQueueMode`** so
  `footerVisibleButtonCount` (`Popup.cpp:7492`) stays valid; extend the debug snapshot with the
  current mode enum for testing.
- **(b) Host a real `DxUi::Toggle`.** Precedent exists *in this file*: the SpeedLimit sub-dialog
  hosts a `DxUi::WindowHost` with `AddChild<…>` (`Popup.cpp:1040-1070`). Brings keyboard+UIA for
  free but injects a control tree into a hand-drawn footer — heavier, and overlaps B5. Prefer (a)
  now; revisit under B5.

Add new resource strings (do **not** repurpose the old two): `IDS_FILEOPS_MODE_CAPTION`
("New tasks:"), `IDS_FILEOPS_MODE_QUEUE_SEG` ("Queue"), `IDS_FILEOPS_MODE_PARALLEL_SEG`
("Parallel"), plus a tooltip `IDS_FILEOPS_MODE_TOOLTIP` explaining "Queue = start each new copy
after the current ones finish; Parallel = run them at the same time." Localize across all
`Lang/*/RedSalamander-*.rc`. Retire the ambiguous "Wait" string.

**Proof.** Extend `PopupLayoutDebugSnapshot` with `footerQueueModeIsParallel` (and, for a segment
control, both segment rects). Self-test: with the popup showing ≥1 active task, assert the control
reports the mode matching `fileOps->GetQueueNewTasks()`; simulate a click on the *Parallel* segment
and assert `GetQueueNewTasks()==false` **and** `ApplyQueueMode` released a waiting task
(reuse the `IsWaitingInQueue()` pattern from
`SelfTest/FileOperations/…Phases05_06.cpp:1508,1586`). Register any new case in
`kFileOpsFamilyDefinitions` (`…SelfTest.cpp:991`) or the Full suite skips it silently.

**Risk.** Localized labels are longer ("Parallèle", "パラレル") — the segmented control must measure
and fit or wrap; budget width like `BatchRenameWindow`'s mode selector (`BatchRenameWindow.cpp:2532`).

---

### A2 — Drop the redundant under-graph item bar; consolidate to one bar

**Problem.** Per-file byte progress is drawn **twice**: the mini-bar beside the file name
(`Popup.cpp:5150-5178`) and the under-graph **item bar** (`:5490-5524`) both render
`itemCompletedBytes / itemTotalBytes`. The under-graph **total bar** (`:5526-5543`) is the only
non-duplicated signal in that block; the single-bar (`:5451-5485`) and marquee (`:5427-5442`)
variants are alternate shapes of the same region.

**Proposal.** Keep the mini-bar as the canonical **per-file** indicator (it is adjacent to the file
it describes — best locality). In the under-graph region:
- **Remove the item bar** (`:5490`, `:5501-5524`) from the two-bar branch; keep **only the total
  (whole-task) bar**, promoted to the single-bar layout (reuse `:5451-5485`).
- Keep the **pre-calc marquee** (`:5427-5442`) — it conveys indeterminate state the mini-bar can't.
- Because per-file progress now lives only at the file name, ensure the mini-bar is **always**
  present for an in-flight determinate file (today it is suppressed at 100% by design, `:5130-5133`
  — keep that; it prevents a stuck-at-100% artifact).

This reclaims one bar's vertical space per card (see B4 density) and removes a genuine "why are there
two identical bars?" confusion.

```text
   BEFORE — 3 bars, ② duplicates ①            AFTER — ① stays, ② gone, ③ relocated
   From: file.CR2      [▓▓▓▓▓▓░░] ①            From: file.CR2   [▓▓▓▓▓▓░░] 68% ①  (per-file)
   ┌──── graph ─────────────────┐             ┌──── graph ─────────────────┐
   └────────────────────────────┘             └────────────────────────────┘
   [▓▓▓▓▓▓░░░░░░░] ② item  ⟵ dup of ①          (removed)
   [▓▓▓▓▓░░░░░░░░] ③ total                     ③ total → header underline ▔▔▔▔ 52%
```


**Decision needed (record in the UI spec).** Should the *whole-task* total bar stay under the graph,
or move up next to the header count `"{op}: 3/10"` (making the header the single home of "overall"
and the file line the home of "current file")? Recommendation: **move the total bar to a thin
underline beneath the header** so vertical scanning is header→overall, file→per-file; but this is a
visual-hierarchy call worth a quick mock before committing.

**Proof.** No self-test asserts on the under-graph bar rects today (tests read the debug snapshot,
not pixels), so removal is low-risk. Add a `PopupLayoutDebugSnapshot` field for the **count of
distinct progress bars rendered per task** and assert it drops from 2→1 (plus the mini-bar) for a
running determinate copy — a RED-able guard against the duplication regressing. Confirm the
fairstream graph-hue test (`PopulateGraphHueDebugSummary`, `Popup.cpp:7458-7461`) is unaffected
(it reads `_rates`, not the bars).

**Risk.** Minimal. Watch card-height math: `barsHeight`/`barsTop` (`:5252-5253`) feed layout; shrink
`barsHeight` when only one bar remains so the card doesn't keep dead space.

---

### A3 — Rework the global presentation

**Problem.** The only global feedback is a footer *sentence* — "N running, M waiting, K need
attention" (`Popup.cpp:4016-4026`). There is no at-a-glance aggregate bar and no Windows taskbar
progress, so a minimized/background transfer is invisible.

**Proposal.**
1. **Footer aggregate progress bar.** Add a slim determinate bar spanning the footer (above or behind
   the summary text) driven by an aggregate fraction = Σ`completedBytes` / Σ`totalBytes` across
   active `FileOperation` tasks (fall back to Σ items when bytes unknown; indeterminate marquee when
   any active task is still pre-calculating). Compute alongside `BuildGlobalStatusSummary`
   (`:649-675`) — add a `globalCompletedBytes/globalTotalBytes` reduction there and expose it on the
   debug snapshot next to the existing global counts.
2. **Windows taskbar progress** via `ITaskbarList3::SetProgressValue`/`SetProgressState`
   (none exists today — no `ITaskbarList3` in the tree). Drive it from the same aggregate:
   `TBPF_NORMAL` while running, `TBPF_PAUSED` when all active are paused, `TBPF_ERROR` when any
   `needAttention`, `TBPF_NOPROGRESS` when idle. Own it on the main `FolderWindow` HWND, RAII the
   `wil::com_ptr` per project rules, and update on the existing 100 ms tick.
3. **Keep the text summary but make it earned** — once a real bar exists, the sentence can shrink to
   just the exceptional part ("2 need attention") and stay silent when everything is nominal, cutting
   footer noise.

```text
   BEFORE — footer is a sentence          AFTER — bar + earned text (+ OS taskbar)
   ─────────────────────────────          ─────────────────────────────────────────
   [Cancel all] [Wait]                    [▓▓▓▓▓▓▓▓▓░░░░░░]  ≈3 min · 240 MB/s
   2 running, 1 waiting, 0 need attn.     [Cancel all]  New tasks:[Queue│Parallel]  2 running · 1 waiting

                                          Windows taskbar button:  ▐▓▓▓▓▓▓░░░▌  (NORMAL/PAUSED/ERROR)
```


**Proof.** Self-test asserts `globalCompletedBytes/globalTotalBytes` on the snapshot for a known
multi-task set; taskbar integration is cover-by-inspection (COM call, no pixel probe) but assert the
state-mapping helper (`running→NORMAL`, `attention→ERROR`, …) with a table-driven unit test.

**Risk.** `ITaskbarList3` must be created after the taskbar-button-created message and released on
shutdown; guard for RDP/session-0. Aggregate fraction must ignore finished/informational tasks to
avoid a bar that jumps when a completed card lingers (auto-dismiss off).

---

## Tier B — workflow

### B1 — Per-task queue control + bulk pause/resume

**Rationale.** Today the only queue lever is the global A1 toggle; a user who wants *this* download
to jump ahead must not have to flip everything to parallel. V1 is implemented: queued cards expose
**"Start now"**, backed by `FileOperationState::RunQueuedTaskNow(taskId)`, which clears that task's
`WaitForOthers` without touching the global flag. V2 is implemented: the footer exposes **Pause all**
when any started live task is unpaused and **Resume all** when started live tasks are paused, backed
by `FileOperationState::SetAllRunningTasksPaused(...)`. Remaining B1 work: **move-up/move-down**
among waiting tasks (needs an engine ordering key — see risk).

**Anchors.** Per-task pause path `TaskPause`, footer `FooterPauseResumeAll`,
`FileOperationState::SetAllRunningTasksPaused(...)`, `ApplyQueueMode`/`SetWaitForOthers`, and
waiting predicates.

**Proof.** Current focused proof: `Phase5_PreCalcCancelReleasesSlot` verifies queued-card Start now
visibility and activation while footer mode remains Queue. `Phase6_PopupSmokeResizeAndPause`
verifies the real footer Pause all / Resume all activation path. Remaining reorder proof should use
three tasks in queue mode after the engine owns an ordering key. Register new cases in
`kFileOpsFamilyDefinitions` when that reorder work lands.

**Risk / decision.** True reordering needs an explicit queue-position field the engine doesn't have
yet (promotion today is implicit). Keep B1's reorder scope deferred to a follow-up that adds an
ordering key. Don't fake reorder in the view.

### B2 — Status-at-a-glance

Add a **left accent stripe** (2–3 dip) on each card colored by `TaskStatusKind`
(running→`accent`, waiting→neutral/sub-text, paused→muted, conflict→`warningText`,
error/failed→`errorText`, done→`accentOk`), plus a consistent small **state chip** in the header so
state never relies on the graph overlay alone. Colors already exist:
`AppTheme.folderView.warningText/errorText` + `accent` (`Popup.cpp:2354-2356`), DxUi `AdornmentTone`
(`DxUi.h:118-124`). Never encode by color alone (see C4 — pair with glyph/text). Effort S; pure
paint, no engine change. Proof: snapshot reports per-card `statusKind`→stripe-color mapping.

```text
   ▎▶ running   (accent)      ▎⏸ paused   (muted)       ▎⚠ conflict (warning)
   ▎  Copy · 3/10             ▎  Move · 1/4              ▎  needs attention
   ┆
   ▎✓ done      (ok green)    ▎✕ failed   (error red)   ▎⌛ waiting  (neutral)
   ▎  Completed               ▎  Failed (0x80070005)    ▎  Waiting for previous
   stripe colour + glyph + text → readable without colour (color-blind safe, C4)
```


### B3 — Post-completion actions

A finished card keeps Dismiss / More... (`Popup.cpp:4884-4899`). Done slice:
completed copy/move tasks with resolved destination data expose **Open destination folder** and
**Reveal item** from the More menu; completed diagnostic tasks also include **Failed items** when
`warningCount|errorCount>0` and route it to the existing File Operations Failed Items pane. Proof:
`PopupLayoutDebugSnapshot` reports the completed action set, and the focused Commands selftest
invokes the More-menu actions to open the real Failed Items pane, navigate the destination pane, and
select the completed destination item.

### B4 — Auto-collapse & group completed; density

Done slice: collapse plumbing now auto-collapses a task once when it first reaches a finished state
and no manual expanded/collapsed override exists. The footer exposes a persisted, normally
hit-testable Compact/Expanded density toggle backed by `fileOperations.popupCompactDensity`; compact
density is a default display state, not a forced stored collapse, so the per-card chevron can expand a
compact row and restore its actions without preventing later completed-card auto-collapse. Compact
rows keep the status glyph/name and draw a mini-bar plus localized percent text when progress is
known. Proof: `PopupLayoutDebugSnapshot` reports the density toggle, its normal hit-test
reachability, compact-density state, compact row, auto-collapse flag, compact progress, and completed
auto-collapse count, and the focused Commands/Settings regressions cover the transition plus
compact-only canonical-save persistence.

Done grouping slice: the popup now inserts a localized **"Completed (N)" group** when at least two
completed file-operation cards remain visible. It defaults expanded; its chevron collapses/hides or
expands the grouped rows without dismissing them; and its `Clear` action dismisses those completed
summaries. `PopupLayoutDebugSnapshot` reports group visibility, expanded state, counts, toggle/clear
visibility, and hidden-by-group state for selected completed tasks.

Remaining slice: animate group/card expand-collapse instead of snapping — reference the `Tree` 320 ms
subtree slide (`DxUi.h:2918-2938`), gated on `reducedMotion`. Effort M.

```text
   EXPANDED (active, full card)     COMPACT density (1 line each)      COMPLETED auto-grouped
   ┌─────────────────────────       ▎▶ report.pdf   ▓▓▓░░░ 40%         ▼ Completed (3)          [Clear]
   │ ▶ Copy · 3/10  …full…           ▎▶ backup.iso   ▓▓▓▓▓░ 82%           ✓ a.zip  ✓ b.txt  ✓ c.mp3
   │ graph, bars, buttons            ▎⏸ song.mp3     paused             ▲ finished cards fold away,
   └─────────────────────────       ▎⌛ notes.md     waiting              one row per group
```


### B5 — Keyboard + UIA (accessibility)

The popup is mouse-only and Narrator-invisible (§1). Add: a focus model over the interactive
`_buttons` (Tab/Shift-Tab order, Space/Enter activate, arrow within a card), shortcuts
(Space=pause/resume focused, Del=cancel, Esc=close), and a UIA provider exposing each task as a
progress element with **LiveRegion announcements** on completion and on needs-attention. Cheapest
correct route is to **migrate the interactive controls to a hosted `DxUi::WindowHost`** (keyboard +
UIA come from the framework; precedent `Popup.cpp:1040-1070`) rather than hand-rolling `WM_GETOBJECT`.
Effort L — sequence **after** A1/A2 settle the control set so the migration targets a stable layout.
Proof: DxUiTests-style focus-order + TogglePattern assertions; announce-on-complete verified via the
UIA event hook.

---

## Tier C — appeal / polish

- **C1 — Aggregate ETA + throughput.** Alongside A3's global bar, show combined "≈ 3 min left ·
  240 MB/s" from the per-task smoothed rate/ETA already computed
  (`SmoothRateForDisplay`/`SmoothEtaSecondsForDisplay`, `Popup.cpp:1743-1755`, `2966-3038`). S.
- **C2 — Speed-limit menu button.** The per-task speed limit is a menu round-trip
  (`TaskSpeedLimit` → `ContextMenu::ShowAsync`, `Popup.cpp:6282`). Keep the current one-button model:
  `Speed: [ 50 MB/s ▾ ]`. The button label is the current limit; clicking opens a flyout with
  Unlimited, common presets (5 / 20 / 50 MB/s), and Custom…. Do **not** replace this with inline
  chips or segmented presets in the card; those add visual noise and read less like the existing
  command. Reuse the preset list already in Preferences (`IDS_PREFS_FILEOPS_BANDWIDTH_*`). S.
- **C3 — Collapsible details + remembered footer-only state.** Explorer-style "hide details" should
  collapse graphs/details to reclaim height, and a stronger **footer-only** state should collapse the
  popup down to just the aggregate progress/footer controls. Persist both the footer-only flag and
  the chosen density in `SettingsStore`; also remember popup **size/position**. Uses existing
  collapse infra + settings persistence. M.
- **C4 — Reduced-motion / high-contrast / color-blind-safe.** Honor `reducedMotion` for the graph
  animation and marquee (freeze to a static fill); ensure every status has a **glyph + text**, not
  color alone (B2 pairs the stripe with the header glyph); verify high-contrast brush swaps
  (`_captionStatus` path, `Popup.cpp:7405`). S.
- **C5 — Conflict prompt enrichment.** Done slice: the inline conflict prompt captures
  source/destination metadata through `IFileSystemIO` when the prompt opens, keeps Source/Destination
  as stacked full-width path rows, and shows compact size/modified metadata on the right side of each
  row when available. Remaining slice: add a real **"Keep both"** action only after the engine can
  choose and retry a unique destination name through copy/move and cross-filesystem bridge paths. Do
  not add a UI-only Keep Both command.
- **C6 — Motion polish.** Throughput graph updates should animate from the previous sample set to the
  next sample set instead of snapping. Pair that with subtle progress-fill easing, aggregate-bar
  easing, and expand/collapse transitions. Avoid decorative motion that competes with the transfer
  status; always respect `reducedMotion` by freezing graph morphs/marquees to a static state. S.
- **C7 — Rainbow throughput graph + DxUi support.** When the app theme is Rainbow
  (`ThemePalette::rainbowMode`) and high contrast is off, the throughput graph may use a restrained
  multi-hue stroke/fill derived from the existing rainbow seed/accent pipeline. Keep the rest of the
  card quiet: rainbow should make the graph feel alive, not turn the whole popup into a decoration.
  Existing DxUi has rainbow-aware grid/tree/menu/combobox styling and HSV helpers, but no reusable
  animated graph primitive. Recommended DxUi improvement: add a small reusable `SparklineGraph` /
  `ThroughputGraph` renderer that accepts samples, optional limit line, `rainbowSeed`, `reducedMotion`,
  `highContrast`, and transition timing; internally cache Direct2D solid/linear-gradient brushes and
  expose debug fields such as `usesRainbowStroke`, `sampleCount`, and `transitionActive`. This keeps
  `FolderWindow.FileOperations.Popup.cpp` from hand-rolling gradient brushes, animation state, and
  theme edge cases. M.

---

## 4. Sequencing & guardrails

1. **A1 → A2 → A3** first (they define the control set and reclaim the space the rest builds on).
2. Then **B2** (cheap, high legibility) and **B4** (density) so the card is compact before adding
   **B1/B3** actions.
3. **B5** last of tier B — it should migrate a *settled* layout, not chase a moving one.
4. Tier C opportunistically.

**Hard rules for every step:**
- Preserve or extend `PopupLayoutDebugSnapshot`; never leave it describing a control that no longer
  exists. It is the only non-pixel test hook.
- Any new self-test case **must** be registered in `kFileOpsFamilyDefinitions`
  (`SelfTest/FileOperations/…SelfTest.cpp:991`) or Full-suite runs skip it silently
  (known trap — see the FileOps family-registration note).
- New user-facing strings go in `.rc` and **all** `Lang/*/RedSalamander-*.rc` satellites; no
  hardcoded literals (CLAUDE.md rule). Retire, don't reuse, the ambiguous "Wait" string.
- RAII all new Win32/COM (`ITaskbarList3` → `wil::com_ptr`); no manual release.
- Respect `reducedMotion`/`highContrast` on every animated or colored addition.
- The full FileOps self-test family runs ~18–20 min and is serialized by a session-global mutex
  (exit 3 if contended) — plan verification runs accordingly.

## 5. Open decisions (resolve in the UI spec before A-tier lands)

1. **A1 shape:** segmented pill (shows both options) vs ToggleSwitch (boolean). Recommend segmented.
2. **A2 total bar home:** keep under the graph, or move to a header underline. Recommend header
   underline; mock first.
3. **B1 scope:** "Start now" only for v1, or invest in an engine queue-ordering key now.
4. **A3 taskbar:** confirm we want OS taskbar progress on the main window (multi-window behavior).
5. **C2 speed control shape:** keep one menu button showing the current limit, e.g. `50 MB/s`, with
   presets/custom in the flyout. Avoid loose chips and inline segmented presets inside the card.
6. **C3 persisted collapse state:** decide whether footer-only mode is global for the popup or
   per-FolderWindow; either way it must round-trip through `SettingsStore`, not just runtime memory.
7. **C7 DxUi graph primitive:** decide whether FileOps should implement rainbow graph drawing locally
   first or wait for a shared DxUi sparkline/throughput graph helper. Recommend shared DxUi support if
   the popup is already moving toward `DxUi::WindowHost` for B5.

## 6. Deliverables

- Code: the A/B/C changes above, each with a registered self-test and updated debug snapshot.
- New spec `Specs/UI/UI_FileOperationsPopup.md` documenting the refreshed anatomy, the mode-toggle
  semantics, the single-progress-bar rule, and the accessibility model (this plan is its seed).
- Update this file's ledger states and the `Specs/Plans/WIP/README.md` index row as items land.
