# Last-Three-Days Code Review — Independent Findings Report (2026-07-11)

## What this document is

An **independent, adversarially-verified** review of every source change committed in the
last three days, produced as a fresh pass (not a re-derivation of the existing
[`CodeReview_Last3Days_Remediation_2026-07-11.md`](CodeReview_Last3Days_Remediation_2026-07-11.md)
checklist). Each finding below was produced by a dimension-scoped reviewer and then
**re-checked by a separate skeptical verifier** that read the actual code at HEAD and tried to
refute it; High/critical candidates got two independent lenses (correctness + reachability).
Severities shown are the **verifier-adjusted** values, with the original finder severity noted
where they diverged.

- **Range reviewed:** `f06b689d4..eaf640798` (34 commits, 2026-07-08 → 2026-07-10).
  `f06b689d4` is the parent of the earliest in-window commit; `eaf640798` is current `HEAD`.
- **Theme of the range:** the *File Operations popup remediation* feature (aggregate progress,
  compact/footer-only density, conflict metadata, completed-group, queued "Start now",
  bulk pause/resume) plus a large *test-suite flake-stabilization* effort.
- **Method:** 9 finder agents → per-finding adversarial verification → this synthesis.
  29 agents total, 0 errors. 20 raw candidates → **18 survived**, **2 refuted**
  (one duplicate across finders → **17 distinct** issues).

### Scope discipline (important)

The working tree is **dirty** with substantial *uncommitted* edits (DxUi, several plugins,
`SearchServiceBroker`, `SqliteIndexStore`, `LocalizationManager`, …). **Those are out of scope**
— this report reviews only what is committed in the range above, reading file context via
`git show eaf640798:<path>` rather than the dirty tree. Two useful facts fell out of that
boundary: the uncommitted tree already contains (a) the locale-aware timestamp fix for the
conflict prompt and (b) a more advanced `started`-gated pause-eligibility model. Where a finding
is "already fixed pending commit," it is flagged as such.

---

## Executive summary

| # | Severity | Area | Finding | Location |
|---|----------|------|---------|----------|
| 1 | **Medium** | correctness | Aggregate meter drops a whole progress dimension when byte-reporting and item-only tasks run together → footer + taskbar can show 100% while work continues | [Popup.cpp:1106](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) |
| 2 | **Medium** | concurrency | `Task::SetPaused` mutates `_paused` + notifies **without `_pauseMutex`** → lost-wakeup can hang a resumed operation indefinitely | [State.cpp:6304](../../../RedSalamander/FolderWindow.FileOperations.State.cpp) |
| 3 | Low | correctness | Copy/move with unknown byte totals renders a **determinate (stuck-at-0%)** aggregate instead of indeterminate | [Popup.cpp:1066](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) |
| 4 | Low *(finder: Med)* | efficiency | Taskbar `ITaskbarList3` progress pushed on **every paint** with no change detection (redundant even when idle-but-open) | [Popup.cpp:7350](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) |
| 5 | Low | efficiency | **Minimized** popup rebuilds the full task snapshot every 100 ms just to push taskbar progress (previously zero work) | [Popup.cpp:7345](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) |
| 6 | Low | efficiency | Footer controls build throwaway DirectWrite layouts + full font-collection glyph probes **every paint** for static text | [Popup.cpp:4074](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) |
| 7 | Low *(finder: Med)* | ui-thread-block | Footer/density toggles do **synchronous multi-file disk writes** (settings + full aggregated schema) on the UI thread | [State.Runtime.Part.cpp:1286](../../../RedSalamander/FolderWindow.FileOperations.State.Runtime.Part.cpp) |
| 8 | Low | correctness | Locale: conflict-metadata timestamp uses a hardcoded `YYYY-MM-DD HH:MM` pattern (already fixed in the *uncommitted* tree) | [Popup.cpp:404](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) |
| 9 | Low | correctness | Consolidated `HasNonDefaultFileOperationsSettings` dropped the `issuesPaneSortDescending` term (latent, doubly-guarded today) | [SettingsStore.cpp:4873](../../../Common/Common/SettingsStore.cpp) |
| 10 | Low *(finder: Med, PLAUSIBLE)* | correctness | Metadata-incomplete traversal seed advertises query-cutover readiness in the idle status badge though every query falls back | [LocalSearchIndexCore.cpp:1397](../../../Common/LocalSearchIndexCore.cpp) |
| 11 | Low *(finder: Med)* | test-coverage | "Mixed known/unknown suppresses aggregate ETA" assertion is **vacuous** (ETA is never computed on that path) | [Commands.SelfTest.FileOps.cpp:4845](../../../RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp) |
| 12 | Low | test-coverage | Quick Search first-match assertion weakened to `alpha.txt OR alpine.log`, and a source-contract test **cements** the weaker check | [Commands.SelfTest.Search.cpp:15813](../../../RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp) |
| 13 | Simplification | duplication | Byte-identical "Start now / Cancel" 50/50 layout block duplicated in two render arms | [Popup.cpp:6956](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) & 7160 |
| 14 | Simplification | duplication | Three-button width-negotiation algorithm copy-pasted between the Start-now and Skip pre-calc arms | [Popup.cpp:6899](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) & 6975 |
| 15 | Simplification | duplication | 12-line member-by-member `ItemMetadata` copy in `BuildSnapshot` (silent-drop hazard on struct growth) | [Popup.cpp:3260](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) |
| 16 | Simplification | duplication | Global-summary → debug-snapshot field mapping duplicated (test builder vs live layout report) | [Popup.cpp:9871](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) & 9109 |
| 17 | Simplification | duplication | `suppressExistsOverwrite` recomputes the `deferLocalExistsOverwrite` predicate instead of reusing it | [State.cpp:3831](../../../RedSalamander/FolderWindow.FileOperations.State.cpp) |
| 18 | Nit | consistency | `SetPopupCompactDensity` mutates the in-memory setting **before** its no-change early-return (asymmetric with `SetPopupFooterOnly`) | [State.Runtime.Part.cpp:1298](../../../RedSalamander/FolderWindow.FileOperations.State.Runtime.Part.cpp) |

**Refuted / not-a-bug:** 2 candidates (see [§ Refuted](#refuted-candidates-not-bugs)).
**Verified clean:** popup memory-safety/RAII (0 findings), localization resource-ID mirroring
(17/17 IDs present in all 4 satellite languages), pure pre-calculation aggregate indeterminacy.

**Headline:** there are **no Critical or High data-safety defects** in this range. The two most
consequential issues are both *reliability/UX* rather than data-loss: a narrow-but-permanent
pause/resume hang (#2) and a misleading "100% / done" aggregate meter (#1). Everything else is
Low, cosmetic, test-fidelity, or maintainability.

---

## Detailed findings

### Medium

#### 1. Aggregate progress meter drops a whole dimension (bytes vs items) — `Popup.cpp`
**Category:** correctness · **Verdict:** CONFIRMED (finder Medium → verifier Medium)
**Where:** `GlobalAggregateProgressFraction` [Popup.cpp:1102-1112](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp),
`HasDeterminateGlobalAggregateProgress` [:1064](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp),
`BuildGlobalTaskbarProgressModel` [:1147-1155](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp),
fed by `BuildGlobalStatusSummary` [:996-1013](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp).

`BuildGlobalStatusSummary` sums `totalBytes` and `totalItems` from **disjoint** subsets of tasks.
`hasUnknownActiveProgress` is set only when a task has *both* `totalBytes==0` **and**
`totalItems==0`, so a task that reports items-but-no-bytes never trips it. The determinacy gate
then treats the mix as determinate, and `GlobalAggregateProgressFraction` prefers bytes whenever
`totalBytes>0`, **ignoring the item-only tasks entirely**.

**Failure scenario (most robust repro):** two concurrent copies in parallel mode — one ordinary
filesystem copy (`totalBytes>0`) plus one size-less-plugin copy (archive / cloud / FTP LIST
without sizes) whose `totalBytes==0` and whose `totalItems` is set to `plannedItems` by the
fallback at [:3341](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp). The aggregate tracks
only the byte copy; when it finishes, both the footer bar **and the Windows taskbar button** show
a determinate green 100% while the size-less copy is still running. The delete-of-1000-files +
copy-of-2GB mix reproduces it the same way.

**Fix:** track whether *every* active (running/waiting) task contributes the chosen dimension
(`allActiveHaveByteTotals` / `allActiveHaveItemTotals`). Choose bytes only when all active tasks
report bytes, else items when all report items, otherwise render **indeterminate**. Apply the
same gate in all three functions so the footer bar and the taskbar never read 100% while an
uncounted task is working. **Severity kept Medium:** display-only and self-correcting (no data
impact), but it defeats the aggregate meter's purpose in a common parallel-mode mix.

#### 2. `Task::SetPaused` resume path has a condition-variable lost-wakeup — `State.cpp`
**Category:** concurrency · **Verdict:** CONFIRMED (finder Medium → verifier Medium; prior plan CR-H1 rated High)
**Where:** `Task::SetPaused` [State.cpp:6296-6312](../../../RedSalamander/FolderWindow.FileOperations.State.cpp),
waiters `WaitWhilePaused` [:6513-6558](../../../RedSalamander/FolderWindow.FileOperations.State.cpp) and
`WaitWhilePreCalcPaused` [:6560-6577](../../../RedSalamander/FolderWindow.FileOperations.State.cpp).

`SetPaused` does `_paused.store(false, release)` (line 6304) then `_pauseCv.notify_all()`
(line 6308) **without ever acquiring `_pauseMutex`**. The waiters block on `_pauseCv` under
`_pauseMutex` with a **timeout-free** predicate that reads the same atomic. This is the textbook
condition-variable lost-wakeup: even with an atomic flag, the predicate must be mutated under the
wait mutex — otherwise `notify_all()` can fire in the window between a waiter's under-lock
`pred()` check and its internal enqueue, and the notification is lost. Because the wait has no
timeout and there is no periodic wakeup, the worker then sleeps until some *unrelated*
`_pauseCv.notify_all()` occurs (a later pause/resume toggle, `SetQueuePaused`, `RequestCancel`,
or shutdown).

**Proof by contrast (same file):** `RequestCancel`
[:6272-6277](../../../RedSalamander/FolderWindow.FileOperations.State.cpp) does exactly the right thing —
`{ std::scoped_lock lock(_pauseMutex); _paused.store(false, release); } _pauseCv.notify_all();`.
`SetPaused` is the odd one out.

**Failure scenario:** user clicks *Resume* / *Resume all* (the new `SetAllRunningTasksPaused`
at [State.Runtime.Part.cpp:939](../../../RedSalamander/FolderWindow.FileOperations.State.Runtime.Part.cpp)
fans this to N tasks via `SetPaused`) exactly as a worker is entering the pause wait. That task's
progress stalls **indefinitely** even though the popup still shows it as running.

**Fix (trivial, load-bearing):** publish the state change under the wait mutex —
`{ std::scoped_lock lk(_pauseMutex); _paused.store(paused, release); } _pauseCv.notify_all();`
— mirroring `RequestCancel`. As defense-in-depth, consider a bounded wait timeout on the pause
CVs. **Severity note:** finder and verifier both landed on Medium (narrow race window → low
per-click odds); the prior plan rated it High because the *consequence* is a permanent hang with
no auto-recovery. Treat it as **top priority regardless of the label** — the fix is one line and
the failure mode is a stuck operation. `SetQueuePaused`
[:6373-6389](../../../RedSalamander/FolderWindow.FileOperations.State.cpp) shares the identical defect but
is **pre-existing** (not touched by this diff); fold it into the same fix.

---

### Low

#### 3. Determinate "stuck-at-0%" aggregate when byte totals are unknown — `Popup.cpp`
**Category:** correctness · **Verdict:** CONFIRMED (Low)
`HasDeterminateGlobalAggregateProgress` [:1066](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)
ignores the already-computed `summary.hasUnknownActiveTransferBytes`
(set at [:1013](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) for COPY/MOVE with
`totalBytes==0`). A single large-folder copy whose byte totals aren't known yet falls back to
`totalItems = plannedItems = 1`, so the gate returns determinate and the fraction is
`completedItems/totalItems = 0/1 = 0%`. The footer bar sits at a solid determinate 0% for the
whole multi-GB copy while the per-task card correctly shows an indeterminate/marquee bar — the
two contradict each other. The author already uses that flag to gate the aggregate **ETA**
([:1048](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)); the determinacy gate simply
omits it. **Fix:** add `&& ! summary.hasUnknownActiveTransferBytes` to the determinacy gate (this
finding is closely related to #1 — a shared "is this dimension trustworthy" helper resolves both).

#### 4. Taskbar progress RPC on every paint, no change detection — `Popup.cpp`
**Category:** efficiency · **Verdict:** CONFIRMED (finder Medium → verifier Low)
`ApplyTaskbarProgress` [:7350](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) keeps no memo
of the last-applied `(state, completed, total)`; it re-issues `ITaskbarList3::SetProgressState`
(and `SetProgressValue`) on every 100 ms timer tick for the popup's lifetime, and the call sits
inside the `BeginDraw`/`EndDraw` scope ([:4987-5001](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)).
When idle-but-open it re-sends `TBPF_NOPROGRESS` 10×/s forever. **Downgraded to Low** because
these are lightweight window-message-based shell calls (not heavy DCOM round-trips) and during an
active copy `SetProgressValue` legitimately changes each tick; the redundant part is the state
re-send and the idle case. **Fix:** cache last-applied values in `ApplyTaskbarProgress` and
early-return when unchanged (the memo must live in the function itself — the hidden/iconic path
at [:7347](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) also calls it). Moving the call
outside `BeginDraw`/`EndDraw` is cosmetic (same thread, same stall) — the memo is the real fix.

#### 5. Minimized popup rebuilds the full snapshot every 100 ms — `Popup.cpp`
**Category:** efficiency · **Verdict:** CONFIRMED (Low)
The minimized/hidden timer branch previously did **zero** work (`return 0;`); this diff inserts
`UpdateTaskbarProgress(hwnd)` ([:8557](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)),
which calls `BuildGlobalStatusSummary(BuildSnapshot())`
([:7345](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)). `BuildSnapshot`
([:3150](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)) allocates 3 vectors + 2
unordered_maps and deep-copies every task record (each embedding a 16-slot in-flight-files array
plus conflict-metadata strings) — 10×/s, while minimized, purely to feed a taskbar value.
**Fix:** gate on a cheap dirty-flag / version counter so no snapshot is built when the published
progress counters are unchanged (pairs naturally with #4's memo).

#### 6. Footer controls do per-paint DirectWrite layout + glyph probes for static text — `Popup.cpp`
**Category:** efficiency · **Verdict:** CONFIRMED (Low)
`DrawFooterAutoDismissControl` calls `MeasureTextWidth`
([:4182](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)) — allocating a throwaway
`IDWriteTextLayout` each frame — and the details chevron calls `DirectWriteFormatHasGlyph`
([:5102](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)), which re-fetches the system font
collection + re-resolves the family/font every frame, all at 10 Hz for content that only changes
on DPI/theme/locale change. (`DrawFooterQueueModeControl`'s measure at
[:4074](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) and the density chevron are
conditional.) `RefreshLocalizedFooterText` already caches the *strings* but not their widths.
**Fix:** memoize `DirectWriteFormatHasGlyph` keyed by `(format, glyph)` after `EnsureTextFormats`,
and cache footer label widths in `RefreshLocalizedFooterText`, invalidating on DPI/theme/locale
change. Note the glyph-probe pattern is pre-existing at ~5 other call sites; only the incremental
per-frame cost is attributable to this diff.

#### 7. Footer/density toggles do synchronous multi-file disk writes on the UI thread — `State.Runtime.Part.cpp`
**Category:** ui-thread-block · **Verdict:** CONFIRMED (finder Medium → verifier Low)
`SetPopupFooterOnly` ([:1286](../../../RedSalamander/FolderWindow.FileOperations.State.Runtime.Part.cpp))
and `SetPopupCompactDensity` ([:1313](../../../RedSalamander/FolderWindow.FileOperations.State.Runtime.Part.cpp)),
invoked from the popup's `WM_LBUTTONUP` handler, call
`SaveFileOperationsSettingsOrLog` → `SettingsHotReload::SaveSettingsAndSchema` →
`SavePreparedSettingsAndSchema`, which synchronously (a) writes the settings JSON and (b) rebuilds
+ writes the **entire aggregated settings schema** (iterating every FileSystem + Viewer plugin).
Each write is an atomic temp+rename with `FlushFileBuffers` + `MOVEFILE_WRITE_THROUGH`; the
verifier found up to **three** such writes per toggle, and the schema doesn't change when a
boolean flips — pure write amplification. This violates the project's "no synchronous I/O in
input handlers" rule. **Downgraded to Low** because it's a rare user-initiated toggle (not a hot
path); the perceptible-hitch case needs a slow/AV-scanned/roaming settings dir. **Fix:** skip the
schema rewrite on value-only changes, or debounce/offload persistence off the input handler. The
same helper is used by the other new footer actions (auto-dismiss, queue mode), so fix it once in
`SaveFileOperationsSettingsOrLog`'s callers.

#### 8. Conflict-metadata timestamp uses a hardcoded ISO pattern, not the user locale — `Popup.cpp`
**Category:** localization · **Verdict:** CONFIRMED (Low) — *found by two independent finders (deduplicated)*
`FormatFileTimeLocalCompact` [:404](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) ends with
`std::format(L"{:04d}-{:02d}-{:02d} {:02d}:{:02d}", …)` — fixed field order and separators from
`SYSTEMTIME`, no `GetDateFormatEx`/`GetTimeFormatEx`. The UTC→local conversion itself is correct.
Since the feature ships cs/fr/ja/sk satellites, a ja-JP/fr-FR user sees ISO ordering regardless
of regional settings. **Two mitigating facts:** (a) this exact pattern is the app-wide house
style for every user-facing timestamp (FindFiles Modified column, CompareDirectories, StatusBar,
ItemProperties), so a lone fix here would create UI inconsistency — proper localization is an
app-wide policy change; and (b) the **uncommitted working tree already replaces this** with
`GetDateFormatEx`/`GetTimeFormatEx` plus a matching self-test, so it is effectively remediated
pending commit. **Action:** commit the existing working-tree fix; decide separately whether to
localize timestamps app-wide.

#### 9. `HasNonDefaultFileOperationsSettings` dropped the `issuesPaneSortDescending` term — `SettingsStore.cpp`
**Category:** correctness (latent) · **Verdict:** CONFIRMED (Low)
The dirty/non-default predicate consolidated out of `SettingsSave.h` into
[SettingsStore.cpp:4861-4874](../../../Common/Common/SettingsStore.cpp) no longer compares
`issuesPaneSortDescending` against its default, unlike the removed version. **No live bug today:**
the serializer only emits `issuesPaneSortDescending` inside the `if (! issuesPaneSortColumnId.empty())`
block, and a non-empty column id already makes the predicate return true; the setter at
[Diagnostics.Part.cpp:365](../../../RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp)
also forbids `descending==true` with an empty column id. It is a **latent trap**: relax either
far-away invariant and a settings object whose only non-default field is
`issuesPaneSortDescending` would be silently reset to default on save. **Fix:** add
`|| fileOperations.issuesPaneSortDescending != defaults.issuesPaneSortDescending` so the single
dirty predicate covers every serializable field. (The consolidation was otherwise a net
improvement — it *added* the new `popupFooterOnly`/`popupCompactDensity` terms.)

#### 10. Metadata-incomplete traversal seed advertises query-cutover readiness — `LocalSearchIndexCore.cpp`
**Category:** correctness · **Verdict:** PLAUSIBLE (finder Medium → verifier Low)
`ResolveSqliteVolumeStateForMetadataCompleteness` [:1397](../../../Common/LocalSearchIndexCore.cpp) now
preserves `kVolumeStateCurrentnessUnproven` instead of demoting to
`kVolumeStateImportedLegacySnapshot` when file metadata can't be fully re-stat'd during the SQLite
mirror. Because `readyForQueryCutover` is derived purely from the *legacy* volume count, such a
volume reports ready even though every query is still rejected (`FallbackReason::StoreStale`) and
served by the slow enumeration path. **The verifier partially refuted the finding's mechanism:**
the cited query-path sink (`StoreState::Ready`/`SyncPhase::Watching`) reaches those states in both
pre- and post-diff code (via `inspectionSucceeded`), so there is **no new dishonesty there**. The
only place `readyForQueryCutover` flips a user-visible outcome is the *idle, pre-query* status
inference in [FindFilesWindow.cpp:2123](../../../RedSalamander/FindFilesWindow.cpp) — and that state is
**transient** (NTFS self-heals to Ready on the next `EnsureVolumeCurrent`, or is reclassified
Legacy on re-mirror) and cosmetic (query results stay correct; live queries honestly report their
fallback). **Downgraded to Low, kept as PLAUSIBLE.** **Action:** low priority — if pursued, don't
derive the idle "ready" badge from `readyForQueryCutover` for a `CurrentnessUnproven` volume;
surface a distinct "proving currentness" phase.

#### 11. Vacuous "mixed known/unknown suppresses aggregate ETA" assertion — `Commands.SelfTest.FileOps.cpp`
**Category:** test-coverage · **Verdict:** CONFIRMED (finder Medium → verifier Low)
In `TestFileOperationsPopupGlobalSummaryIgnoresFinishedTasks` the assertion at
[:4845](../../../RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp)
(`Require(! mixedRateSummary.footerAggregateEtaVisible, …)`) passes **tautologically**:
`DebugBuildFileOperationsPopupGlobalSummarySnapshot` calls `BuildGlobalStatusSummary` with **no
rate map**, and `hasAggregateEta` is only ever set inside the `if (rates)` block or by a
`>= 0.0` override — the test passes `-1.0` and no rates, so the ETA is *never computed*. Deleting
the production suppression guard `! summary.hasUnknownActiveTransferBytes`
([Popup.cpp:1048](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)) would **not** flip this
assertion. (The sibling `footerAggregateBytesPerSecond == 4096.0` assertion at :4843 *is*
meaningful — only the ETA-suppression assertion is vacuous.) **Fix:** extend the debug entry
point to accept a synthetic per-task rate map (so the real `if (rates)` + `hasUnknownActiveTransferBytes`
path runs), or drive a real running task; otherwise remove the misleading assertion.

#### 12. Quick Search first-match assertion weakened and cemented by a source-contract test — `Commands.SelfTest.Search.cpp`
**Category:** test-coverage · **Verdict:** CONFIRMED (Low)
Commit `3be0626a8` relaxed the first-match check at
[:15813](../../../RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp) from
`focusedDisplayName == L"alpha.txt"` to `== L"alpha.txt" || == L"alpine.log"`, and the Down-nav
expectations now float with whatever anchored first. The verifier traced the **unchanged product
code**: with pre-search focus pinned to `gamma.txt` under Name-Ascending sort, typing `a` then `l`
deterministically lands on `alpha.txt` — there is *no* correct-code path to `alpine.log`, so the
OR is strictly looser than the real behavior and a regression anchoring on the second match would
pass. Worse, `TestHarnessSourceContracts.Tests.ps1` (lines 304-305) now *requires* the loose
message and *forbids* the strict one, locking in the reduced coverage. **Fix:** restore
`initialQuickSearchFocus == alMatchDisplayOrder.front()` (pre-search focus is already pinned, so
the fallback rationale doesn't apply) and relax the Pester contract accordingly. If a real flake
produced `alpine.log`, that indicates product-side nondeterminism the OR now masks — worth a look.

---

### Simplification (quality-only; no runtime defect)

All five were confirmed as byte-identical/near-identical duplication introduced by this diff.
None changes behavior today; each is a maintainability hazard where a one-sided future edit
diverges silently.

- **13.** "Start now / Cancel" 50/50 layout block duplicated verbatim at
  [Popup.cpp:6956-6974](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) and 7160-7178
  → hoist a single `layoutStartNowCancel` lambda.
- **14.** Three-column width-negotiation (constants 68/140/72 dip, side split) copy-pasted between
  the Start-now arm [Popup.cpp:6899-6955](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) and
  the Skip arm 6975-7053 → extract a helper parameterized by the leading button's base width.
- **15.** 12-line member-by-member `ItemMetadata` copy in `BuildSnapshot`
  [Popup.cpp:3260-3272](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) → add a
  `ToSnapshotMetadata(...)` converter (or hoist one shared struct); prevents a field silently
  dropping from the conflict prompt if the struct grows.
- **16.** Global-summary → `PopupLayoutDebugSnapshot` mapping duplicated between the `ENABLE_TESTS`
  builder [Popup.cpp:9871](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) and the live
  `OnLayoutSnapshotReport` at 9109/9146-9165 → factor `PopulateGlobalSummaryDebugFields(...)` so
  test and live paths can't diverge (test-fidelity risk).
- **17.** `suppressExistsOverwrite` [State.cpp:3831](../../../RedSalamander/FolderWindow.FileOperations.State.cpp)
  re-types the exact predicate already held in `deferLocalExistsOverwrite` (:3759) →
  `suppressExistsOverwrite = deferLocalExistsOverwrite && (…metadata clause…)`.

### Nit

- **18.** `SetPopupCompactDensity`
  [State.Runtime.Part.cpp:1298-1313](../../../RedSalamander/FolderWindow.FileOperations.State.Runtime.Part.cpp)
  calls `SetPopupCompactDensityInSettings(...)` **before** its `if (previous == compactDensity) return;`
  early-return, whereas the sibling `SetPopupFooterOnly` checks first. Harmless today (writing the
  same value is idempotent) but asymmetric; reorder the check above the setter to match. *(Found
  during independent scouting; not surfaced as a separate fleet finding.)*

---

## Refuted candidates (not bugs)

Recorded so they are not re-flagged in future reviews.

1. **"`ExecuteInPaneLocation` ignores `SetFolderPath`'s `updatingPath` no-op → spurious
   `ERROR_NOT_SUPPORTED`"** ([FolderWindow.FileSystem.cpp:4661](../../../RedSalamander/FolderWindow.FileSystem.cpp)) —
   **REFUTED.** `updatingPath` is a *synchronous* reentrancy guard set/cleared within one
   `UpdateViews` block (set at :5218, cleared at :5239), never true during async enumeration. The
   only caller (`SubmitCompletedOverflowAction`) is a UI-thread-serialized menu action that cannot
   enter while the UI thread is inside that synchronous block. The `ERROR_NOT_SUPPORTED` return is
   the *intended* result for a genuine provider-identity mismatch (asserted by a self-test at
   FileOps.cpp:6037). Even in the impossible reentrant case, `restorePreviousLocation()` is
   likewise a guarded no-op, so no corruption occurs.

2. **"Pause-all eligibility (started-gating + mixed paused/pre-calculating) has no deterministic
   coverage"** — **REFUTED for this review target.** The finding was generated against the **dirty
   working tree**, which contains a `pauseEligibleRunning`/`pauseEligiblePaused` + `if (task.started)`
   model. That model **does not exist** at `eaf640798`: the committed struct has plain
   `activeRunning`/`paused` with no started-gating (verified — zero `pauseEligible` hits at HEAD).
   The specific regressions it worried about reference code paths not in this diff. A much weaker
   true statement remains — the two new pure functions `HasFooterPauseResumeAllControl` /
   `FooterPauseResumeAllShouldPause` and their exported debug fields have no committed self-test —
   but that is a minor coverage gap, not the claimed defect.

---

## Verified-clean areas (no findings)

- **Popup memory safety / RAII / lifetime (0 findings).** The dedicated finder found no
  use-after-free, no dangling references into resized vectors, no raw `new`/`delete`, no
  hand-rolled handle cleanup, and no `std::optional` `operator*` misuse across the +2.4k-line
  `Popup.cpp` change. COM/D2D resources use `wil::com_ptr`; handles use WIL wrappers.
- **Localization resource IDs.** All 17 new `IDS_FILEOPS_*` / `IDS_FILEOP_*` string IDs are
  present in the base `RedSalamander.rc` **and** mirrored in all four satellites
  (cs-CZ, fr-FR, ja-JP, sk-SK) with matching identifiers — no missing-ID gaps or collisions.
  (The only localization defect is the *timestamp format* — finding #8 — not a missing string.)
- **Pure pre-calculation indeterminacy.** A task that is only pre-calculating (status
  `Calculating`, `totalBytes==0 && totalItems==0`) is counted as running and correctly trips
  `hasUnknownActiveProgress`, so a pure pre-calc aggregate renders **indeterminate** — the prior
  plan's CR-M2 concern is handled for that case. Findings #1/#3 are the residual gaps where a
  task reports *one* dimension but not the other.

---

## Reconciliation with the existing remediation checklist

Mapping this independent pass onto
[`CodeReview_Last3Days_Remediation_2026-07-11.md`](CodeReview_Last3Days_Remediation_2026-07-11.md):

| Prior plan item | This review | Notes |
|---|---|---|
| **CR-H1** pause/resume lost wakeups | **Confirmed** (finding #2) | Precise mechanism identified; `SetPaused` vs `RequestCancel` contrast. Rated Medium here (narrow race) vs High there — treat as top-priority regardless (1-line fix, permanent-hang consequence). |
| CR-M2 indeterminate during pre-calc | **Partially already handled** | Pure pre-calc is indeterminate at HEAD; residual gaps are the mixed-dimension cases (#1) and unknown-bytes (#3). |
| CR-M4 shared known-progress predicate | Aligns with #1/#3 | A single "is this dimension trustworthy" helper resolves both the compact-row and aggregate cases. |
| CR-M5 synchronous settings persistence in input handling | **Confirmed** (finding #7) | Verified up to 3 atomic writes + full schema rebuild per toggle; downgraded to Low (rare user action). |
| CR-M1 query-cutover honesty (`CURRENTNESS_UNPROVEN`) | **Downgraded/PLAUSIBLE** (finding #10) | Real effect is narrower and transient than stated; only the idle status badge, not query correctness. |
| CR-L1 exhaustive non-default predicate (`issuesPaneSortDescending`) | **Confirmed** (finding #9) | Latent, doubly-guarded; add the missing term. |
| CR-L3 localize conflict-metadata timestamps | **Confirmed** (finding #8) | Already fixed in the uncommitted tree; commit it. |
| CR-M7 monolithic popup Commands case / isolation; CR-L5 English-only assertions | Related to #11/#12 | This pass confirms a *vacuous* assertion (#11) and a *weakened+cemented* assertion (#12) as the concrete instances. |
| **CR-H2** Search-service test lifetime; **CR-H3** UIA helper deadlines; **CR-H4** perf-closeout evidence | **Out of this pass's file scope** | These concern the foreground-search-service test harness, DxUi UIA helpers, and performance-evidence process. The UIA/broker code lives largely in the **uncommitted** tree; the perf-closeout item is process, not code. Not independently re-verified here — defer to the remediation plan. |

**Net:** the checklist's data-safety framing holds up, but this verification pass **downgrades
several Mediums to Low** and reclassifies CR-M1 as a transient cosmetic issue. The one item worth
elevating in priority (despite its Medium label) is CR-H1 / finding #2.

---

## Recommended remediation order

1. **Finding #2 (`SetPaused` lost-wakeup)** — one-line lock fix mirroring `RequestCancel`; fold in
   the pre-existing `SetQueuePaused` twin. Add a deterministic pause→resume-progresses regression
   test (a `notify` right at the wait boundary). *Highest value: trivial fix, permanent-hang risk.*
2. **Findings #1 + #3 (aggregate dimension/determinacy)** — introduce one shared
   "dimension-trustworthy" predicate and apply it in `HasDeterminateGlobalAggregateProgress`,
   `GlobalAggregateProgressFraction`, and `BuildGlobalTaskbarProgressModel`; add a debug-snapshot
   test for the byte+item-only mix. Resolves the misleading 100% and stuck-0%.
3. **Findings #11 + #12 (test fidelity)** — de-vacuum the ETA assertion and restore the strict
   Quick Search first-match check (+ relax the Pester contract). Do this alongside #1 so the new
   aggregate tests are meaningful.
4. **Finding #7 (sync save on UI thread)** — skip schema rewrite on value-only changes / debounce.
5. **Findings #4, #5, #6 (per-paint efficiency)** — one change-detection memo + dirty-flag covers
   #4 and #5; cache footer widths/glyphs for #6.
6. **Finding #8 (timestamp locale)** — commit the existing working-tree fix.
7. **Findings #9, #18 (latent/nit)** and **#13-#17 (simplification)** — batch as a cleanup commit.
8. **Finding #10 (search idle badge)** — lowest priority; transient cosmetic.

---

## Appendix — commits reviewed (`f06b689d4..eaf640798`)

- `eaf640798` 2026-07-10 Merge branch 'codex/test-suite-stabilization-flake-convergence'
- `886feadde` 2026-07-10 Stabilize test suite flake convergence
- `9b264c763` 2026-07-10 Complete file operations popup remediation
- `8fa8954fe` 2026-07-10 Checkpoint Shortcuts live search stabilization
- `58fe20177` 2026-07-10 Checkpoint Preferences broad-order stabilization
- `f94429cee` 2026-07-10 Stabilize Preferences broad-order readiness checks
- `9b8b801eb` 2026-07-10 Checkpoint Commands broad-order stabilization
- `388a5d1d4` 2026-07-10 Stabilize foreground search service tests
- `efef2ac4a` 2026-07-10 Make foreground search readiness explicit
- `0a62cd231` 2026-07-10 Stabilize Commands find UIA actions
- `7b1ce1b31` 2026-07-10 Stabilize search service foreground status probes
- `97870a0c6` 2026-07-10 Stabilize remaining suite flake contracts
- `d5c5819af` 2026-07-10 Stabilize menu cursor and compare maintenance tests
- `78366858d` 2026-07-10 Add file operations completed group
- `a85cf77e9` 2026-07-10 Add file operations footer bulk pause
- `8c8d65bd9` 2026-07-10 Fix file operations popup review followups
- `5f18b4055` 2026-07-10 Update Tools Pester inventory contract
- `00013ba62` 2026-07-09 Checkpoint file operations taskbar retry
- `a4e27e5ad` 2026-07-09 Fix file operations live summary totals
- `56e8dc006` 2026-07-09 Fix compact file operations popup density
- `79bb08bcb` 2026-07-09 Add conflict metadata to file operations popup
- `b9865d3e3` 2026-07-09 Add compact file operations popup density
- `e684acde5` 2026-07-09 Add completed file operations destination actions
- `33102676f` 2026-07-09 Add completed file operations failed items action
- `e6931d234` 2026-07-09 Cover file operations popup high contrast status
- `86c01d22d` 2026-07-09 Honor reduced motion in file operations popup
- `42e9af6a5` 2026-07-09 Add file operations queued Start now action
- `32abea501` 2026-07-09 Add file operations popup status indicators
- `34efca2f7` 2026-07-09 Add file operations aggregate progress
- `eee0a87bc` 2026-07-09 Refine file operations popup UX
- `3be0626a8` 2026-07-08 Stabilize self-test convergence checkpoint
- `f82fd7ca1` 2026-07-08 Save stabilization progress and WIP spec updates
- `44ea59c23` 2026-07-08 Save stabilization progress and gap plan
- `188f40bbf` 2026-07-08 Stabilize command suite menu hover checks

### Method detail

9 finder agents (popup-correctness, popup-memory-raii, popup-render-perf, state-concurrency,
settings-prefs, search-fs-theme, localization, simplification, test-coverage) reviewed the diff
independently; each candidate finding was then re-checked by a separate verifier that read the
code at HEAD and attempted to refute it (two lenses — correctness + reachability — for
High/Critical candidates). 29 agents, 0 errors. Findings kept unless *every* verifier refuted
them. All line numbers are at `eaf640798`.
