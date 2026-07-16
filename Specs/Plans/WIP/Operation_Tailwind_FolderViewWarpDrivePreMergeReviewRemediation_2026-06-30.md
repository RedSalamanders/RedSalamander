# Operation Tailwind — FolderView WarpDrive Pre-Merge Review Remediation (2026-06-30)

Tracking plan for findings from a max-effort multi-agent review of the **`codex/folderview-warpdrive`**
branch's own production deltas (`git diff master..HEAD`, 26 commits: cached-first thumbnails, draw-item
brush reuse, icon-cache failure-TTL + recall avoidance, render device-loss recovery, async paste-shortcut
worker, refresh→paint instrumentation, and the test/perf-budget harness). This is the **self-review of the
WarpDrive branch before it lands** — distinct from the sibling plans, which review *other* targets:
[[Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28]] (the perf execution plan that
produced this branch), `Operation_IronLedger...` (FolderView data-safety from the audit), and
`Operation_Clearwater...` (the `master` 2-day review; its **CW-1** cross-FS-move fix lives on this branch and
was re-confirmed correct here — see "Verified correct" below).

Method: 8 cluster reviewers (thumbnail concurrency · icon cache · rendering · paste/clipboard ·
enumeration/metrics · cross-FS move · menu/nav/prefs · cross-cutting harness), each covering correctness +
architecture + simplification, → per-finding adversarial verification (independent skeptic instructed to
*refute*, with full-file context) → synthesis. 19 raw findings → **10 confirmed, 1 uncertain, 9 refuted**.
Every item below survived its skeptic. Severity is the verifier-adjusted severity.

> **Headline:** No High-severity bugs. No data-loss bugs. The single pre-merge fix that mattered was
> **TW-1** (a self-test that silently protected nothing; **DONE 2026-07-05 via Granite GR-24**). Everything else is Low-severity dead/test-only
> state plus two product decisions (**TW-D1**, **TW-D2**) that want an explicit author "yes, intended."
>
> **⚠️ Superseded 2026-07-01 — see "Second pass" below.** A wider-scope ultra review (full branch diff incl.
> uncommitted WIP, not just this pass's committed-only diff) found this branch's first **HIGH**-severity
> bug (**TW-8**) and proved this pass's own **TW-7** conclusion backwards. Read the second-pass section
> before treating this headline as current.

> **Branch topology note.** Review target is `codex/folderview-warpdrive` (HEAD). All anchors are
> **current-file** line numbers on HEAD (not the diff), verified 2026-06-30. Re-run
> `git diff master..HEAD -- <file>` before editing — other in-flight work on this branch may have shifted
> lines. There are no `master`-vs-branch topology concerns here: every finding is a property of this
> branch's own new code.

## Second pass — ultra review (2026-07-01, full branch diff incl. uncommitted WIP)

Wider scope than the first pass: `master`'s HEAD equals the merge-base with this branch (master has not
moved), so `git diff master` (no triple-dot) captured both the 26 committed commits **and** the
then-current uncommitted working-tree changes in one diff — 97 files, ~14,000 changed lines. Method:
10 cluster reviewers (icons/enumeration/cache · rendering/interaction · fileops/clipboard/paste ·
DxUi accessibility core · DxUi controls · ConnectionManager+Preferences · CompareDirectories/NavigationView ·
SelfTest Commands harness A · SelfTest Preferences-panes+harness-common · FileSystemDummy+build/perf tooling),
each re-verifying the relevant TW items above against the current tree plus hunting new findings, → one
independent adversarial verifier per candidate (28 candidates) → a fresh gap-sweep pass → verify-the-sweep.
**26/28 confirmed, 2 refuted, +2 new from the sweep** (one of which *overturns* the original TW-7 analysis
below — see the corrected TW-7 row).

> **Headline of pass 2:** still no data-loss bugs, but this pass found the first **HIGH**-severity item on
> this branch (**TW-8**, silent sticky-inverted file selection) and proved **TW-7's original conclusion was
> backwards** — the "dead" field it said to delete is the fix for a real staleness bug. IDs **TW-8..TW-29**
> below are new; TW-1/TW-T1 were later fixed via Granite GR-24, TW-5/TW-6 remain open, and TW-4 was later fixed via Granite GR-18; TW-2/TW-3/TW-D1/TW-D2 were
> not re-examined in this pass (no new evidence either way — still open per the first pass). All anchors are
> **current-file** line numbers verified 2026-07-01 against the working tree at the time; re-run
> `git diff master -- <file>` before editing, since this file's own state is a moving target while WIP is
> uncommitted.

**Re-verification of existing items (no ID change, status confirmed unchanged):**
- **TW-1** — **DONE 2026-07-05 via Granite GR-24.** It had been independently re-confirmed open by three separate cluster reviewers in this pass
  (rendering cluster, icons cluster cross-check, and the SelfTest-harness cluster with exact grep evidence).
  The live guard now keeps normal selected/hovered renders at `drawItemTransientBrushCreateCount == 0`
  and uses an `ENABLE_TESTS` positive-control seam to prove the same counter can increment.
- **TW-4** — **DONE 2026-07-05 via Granite GR-18.** The abandoned-callback late-post
  (`FolderView.Icons.cpp`) now routes off the pending-accounting path with `countsPending=false`, while
  normal thumbnail worker posts still own exactly one pending increment/decrement. Covered by
  `folderView_thumbnail_stale_bitmap_messages_preserve_pending_count` together with GR-18's stale-message sibling.
- **TW-5** — re-confirmed **still open** and its cost is now **worse**: this pass's own uncommitted delta adds
  two *more* unthrottled full-accessibility-snapshot-rebuild call sites (see new **TW-23**, **TW-24** below)
  that funnel through the same missing-throttle `RefreshWindowHostAccessibilitySnapshot` path as the original
  `OnSize` finding. Recommend fixing TW-5/TW-23/TW-24 together as one throttling pass rather than three
  separate patches.
- **TW-6** — re-confirmed **still open**, unchanged mechanism (`IconCache.cpp:1153`, `PathQueryKey` bakes in
  raw `fileAttributes` even when `useFileAttributes==false`). Still LOW/efficiency-only, not correctness.
- **TW-7 — CORRECTED, was WRONG.** ~~Original text: "Dead `PasteShortcutResult::generation`... harmless."~~
  A sweep reviewer in this pass proved the opposite: `OnPasteShortcutComplete` (`FolderView.FileOps.cpp:945-952`)
  checks **only** `OrdinalString::EqualsNoCasePath(result.targetFolder, _currentFolder)` + `_fileSystem`
  presence. That is **not** sufficient staleness protection — unlike every other async-completion path in this
  class (`ProcessEnumerationResult` at `FolderView.Enumeration.cpp:1491`, `OnCreateIconBitmap` at
  `FolderView.Icons.cpp:1362-1366`, the busy-overlay path at `FolderView.ErrorOverlay.cpp:543`), all of which
  gate on the generation token *because* navigating away and back to the same folder (F→G→F) leaves
  `_currentFolder == F` unchanged while `_enumerationGeneration` has been bumped (`EnumerateFolder()`
  `fetch_add`s it on every call, including re-navigation to the same path — `FolderView.Enumeration.cpp:989`).
  A slow `CreateShellShortcut` worker whose completion lands after such a round-trip passes the folder-path
  check and unconditionally calls `RememberFocusedItemForFolder` + `EnumerateFolder()` — stomping the freshly
  re-navigated view's items/scroll/focus and firing a redundant, superseded re-enumeration. **Severity raised
  LOW → MEDIUM. Do not delete `PasteShortcutResult::generation` — wire it up instead**: add
  `if (result.generation != _enumerationGeneration.load(std::memory_order_acquire)) { return; }` to
  `OnPasteShortcutComplete`, matching the sibling paths above. (Ironically the first pass's own cluster
  reviewer for this file re-confirmed the *original, wrong* TW-7 conclusion — a second independent read is
  what caught it. Lesson: re-verification against current code beats trusting a prior "verified correct.")

### Routed — correctness / test integrity (new, pass 2)

| ID | Sev | Item (current-file anchor) | Fix direction |
|----|-----|--------------------------|----------------|
| **TW-8** | **HIGH — DONE 2026-07-05 via Granite GR-21** | **`CompareDirectoriesWindow`'s "Invert Differences Selection" became a silent, sticky per-session mode instead of a one-shot action.** `IDM_COMPARE_INVERT_DIFFERENCES_SELECTION` persisted `_selectionMode = Inverted`, and compare-window refresh could preserve the inverted selection. | Fixed in Granite GR-21: Invert now applies the inverted decision predicate only to the current panes and immediately resets `_selectionMode` to `Default`; compare-window left/right refresh commands re-apply the default compare selection before forcing pane refresh. RED archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_143810/`; GREEN focused archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_144738/`; GREEN adjacent `cmd_compare_directories_` archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_145147/`. |
| **TW-9** | MEDIUM | **`IGridItemProvider::get_IsSelected` contradicts `ISelectionProvider::GetSelection()` for offscreen-selected grid rows.** This pass's diff widens `GetSelection()` (`DxUi.Accessibility.cpp` ~5730-5744) to include selected rows scrolled off-screen by adding them to `record.selectedGridRowIds`/`record.gridRows` without adding them to `record.gridVisibleRowIds`. `get_IsSelected` (`:5974`) still gates purely on `SnapshotContainsGridRow` (visible-only), unlike sibling call sites `GetPatternProvider` (`:4326`) and the `UIA_SelectionItemIsSelectedPropertyId` property getter (`:4576`), which both correctly OR in `SnapshotGridRowIsSelected(...)` for exactly this case. An AT client that calls `GetSelection()`, gets a provider for an offscreen-selected row, then calls `get_IsSelected()` on it receives `UIA_E_NOTSUPPORTED` instead of `TRUE` — self-contradictory. | Add the same `SnapshotContainsGridRow(*record, _gridRowId) \|\| SnapshotGridRowIsSelected(*record, _gridRowId)` fallback used at `:4326`/`:4576` to `get_IsSelected` at `:5974`. |
| **TW-10** | MEDIUM | **`get_SelectionContainer` has the identical gating gap as TW-9**, same root cause, same offscreen-selected-row scenario. (`DxUi.Accessibility.cpp:5998`, `gridRowSupported` computed from `SnapshotContainsGridRow` alone.) | Same fix as TW-9, applied at `:5998`. Consider factoring both into one shared `IsGridRowReachable(record, rowId)` helper so the fallback can't be forgotten a third time. |
| **TW-11** | MEDIUM | **New stationary-cursor hover-priming regresses submenu auto-open-on-hover.** A new block in `CreateMenuPopupWindow` (`DxUi.Menu.cpp:3173-3193`) directly sets `registeredPopup->hoveredIndex` when a popup opens under an already-stationary cursor. This defeats `RouteMenuPointerHover`'s existing `initializeStationaryPopupHover` mechanism (`:3889-3896`, added 2026-05-20 specifically to auto-open a submenu under a resting cursor): because `hoveredIndex` is now pre-populated, the guard's `!targetPopup->hoveredIndex.has_value()` conjunct is false, so the first real hover-check short-circuits to `Ignored` and `ScheduleSubmenuOpenTimer` never fires. The user must move the mouse off the item and back on to trigger the auto-expand — a real regression of the 2026-05-20 feature for the common mouse-invoked-menu case. | Route the new priming through the same `RouteMenuPointerHover`/`HandleMenuMouseMoveAtPointDip` path instead of setting `hoveredIndex` directly, so the submenu-open-timer side effect isn't dropped. |
| **TW-12** | MEDIUM | **→ ROUTED 2026-07-02 to Granite GR-1 (single owner; GR-1 escalates this same mechanism to a HIGH confirmed focus trap with more arming sites).** **Path-edit blur-suppression (TW-D2) can fight focus restoration to *other applications*, not just in-app click-away.** `RefreshActiveEditHostAfterParentPaint` (`NavigationView.Edit.cpp:2429-2437`) computes `focusStillBelongsToEdit`/`focusStillBelongsToRoot` but the suppression-renewal block ignores both and re-arms the 2000ms window unconditionally (gated only on `!_embeddedDestinationMode`) on every NavigationView repaint while `_editMode` is true — including after focus has genuinely left to a foreign top-level window/process. Neither `OnKillFocus` (`NavigationView.Interaction.cpp:807-820`) nor the `SetOnBlur` handler (`NavigationView.Edit.cpp:759-765`) checks where focus actually went before calling `SetFocus(_pathEdit->hwnd)` — only the suppression timer. A stray repaint (window uncover, DWM composition, theme change) after an alt-tab away renews the window and yanks focus back into this app. Goes beyond the already-documented TW-D2 UX tradeoff (which only covers in-app click-away). | Use the already-computed `focusStillBelongsToEdit`/`focusStillBelongsToRoot` to gate the renewal: only re-arm suppression if focus is still somewhere in this app's window tree. Fold into the TW-D2 decision — this is a correctness bug on top of the UX question, not a new UX question. |
| **TW-13** | MEDIUM | **`Toggle::ApplyCheckedState` (real keyboard/mnemonic activation) doesn't refresh the accessibility snapshot, unlike the programmatic `SetChecked`.** This pass's diff added `RefreshWindowHostAccessibilitySnapshot` to `Toggle::SetChecked` (`DxUi.Controls.cpp:2410-2421`) but not to `ApplyCheckedState` (`:2565-2582`), which is what `OnKeyDown`/`OnMnemonic` call. The mouse-click path happens to be covered by a separate new guard in `WindowHost`'s `WM_LBUTTONUP` handler, but `WM_KEYDOWN`/`WM_KEYUP`/mnemonic dispatch have no equivalent — confirmed no refresh call anywhere in those paths. A UIA client reading `ToggleState` after `VK_SPACE`/`VK_RETURN`/mnemonic activation of any Toggle (including the credential prompt's "Show secret" toggle) sees the stale pre-activation value until an unrelated refresh fires. | Add the same `RefreshWindowHostAccessibilitySnapshot` call to `ApplyCheckedState` directly (the single shared activation path), rather than patching each dispatcher separately. |
| **TW-14** | MEDIUM | **`ComboBox::CommitSelection`/`SyncEditableSelectionFromText` (real click/Enter/typeahead paths) don't refresh the accessibility snapshot, unlike the programmatic `SetSelectedIndex`.** Same pattern as TW-13: `SetSelectedIndex` (`DxUi.ComboBox.cpp:595-609`) was patched, `CommitSelection` (`:2608-2645`) and `SyncEditableSelectionFromText` (`:2647-2664`) were not. Typeahead-driven selection on a non-editable ComboBox never reaches `NotifyTextChanged` at all, so it never refreshes. | Add `RefreshAccessibilitySnapshot()` to `CommitSelection` and `SyncEditableSelectionFromText` directly. |
| **TW-15** | MEDIUM | **`Preferences.Dialog.cpp`'s Plugins deferred-action handler calls the wrong pane's focus-restore method — a copy-paste bug that silently drops a queued focus restoration.** `HandleDeferredPaneAction`'s Plugins case (`:5216`) unconditionally calls `hostState._compareDirectoriesPane.RestoreDeferredFocusAfterLayout()` — confirmed absent from `master`, and Plugins has no equivalent method. `SetFocusControl`'s ancestor-visibility check prevents the originally-feared "focus steal to a hidden control," but `RestoreDeferredFocusAfterLayout()` unconditionally clears `_compareDirectoriesPane._deferredFocusAfterLayout` back to `None` as a side effect even when the restore silently no-ops. Repro: toggle Compare Directories' "Ignore files" control while it has focus (arms a deferred restore), switch to Plugins, search/configure/test a plugin (fires the buggy call, silently discards the pending restore), switch back to Compare Directories — the previously-queued focus restoration is gone. | Remove the wrong-pane call from the Plugins case entirely (Plugins has nothing analogous to restore). If Plugins ever needs its own post-layout focus restore, add a dedicated `_pluginsPane` method, following the `_fileOperationsPane.PostLayoutFocusCustomBandwidthEdit()` pattern already used two cases below it. |
| **TW-16** | MEDIUM · TEST-INTEGRITY | **`TestSelectionSelectAllKeepsNavigationShellStable` lost its core assertion.** This pass's diff added a pre-check settle-wait but dropped the `g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount` clause from the actual pass/fail predicate (`Commands.SelfTest.Navigation.cpp:818-826`) — the refresh count is now read only for the failure diagnostic message (`:842`), never checked as part of success. This test exists specifically to catch Select-All/Unselect-All triggering spurious pane refreshes; that regression class is no longer caught. | Restore the `DebugGetForceRefreshCount(...) == baselineRefreshCount` conjunct in the actual predicate (not just the diagnostic string). |
| **TW-17** | MEDIUM · TEST-INTEGRITY | **`FocusFindRootNavigationForKeyboard`'s settle predicate dropped `value.isForegroundWindow`.** (`Commands.SelfTest.Search.cpp:1199`, confirmed via diff — the clause was removed, leaving only the weaker `value.hasWin32Focus`.) `isForegroundWindow` (OS-level `GetForegroundWindow()` root match) and `hasWin32Focus` (tracked-focus-HWND match) are genuinely decoupled Win32 concepts (`FindFilesWindow.cpp:6635-6640`) that can diverge under `SetForegroundWindow` races/lockouts. A regression where keyboard input routes to the RootCombo's win32-focus state without the window truly holding OS foreground would previously fail this wait; now it can pass. | Restore the `value.isForegroundWindow` conjunct in the settle predicate. |

### Routed — performance / efficiency (new, pass 2)

| ID | Sev | Item (current-file anchor) | Fix direction |
|----|-----|--------------------------|----------------|
| **TW-18** | MEDIUM · ARCH | *(Execute with Granite GR-8 + GR-P3 as one perf-gate tooling cluster — 2026-07-02.)* **The FolderView perf-regression hard gate is wired into zero automated runners.** `--selftest-perf-budget=PATH` is parsed and functional (`RedSalamander.cpp:7249-7251` → `CheckFolderViewPerfBudgets`), but neither `Tools/Run-AllTests.ps1`/`Tools/TestRunPlan.ps1`'s `Get-RSSelfTestArguments` nor `.github/workflows/ci.yml` (lines 113/119) ever pass it — the only place the flag appears anywhere in the repo is a documented manual copy-paste command in `Specs/Testing/Testing_PerformanceValidation.md:294`. `Run-AllTests.ps1 -Suite Full` and CI both go green with **zero** perf-regression protection. This sharpens the original synthesis's "watch item" (which framed it as maybe-environment-scoped) into a confirmed fact: the gate runs for nobody automatically. | Wire `--selftest-perf-budget=Specs\Testing\FolderViewPerfBudgets.json5` into the Release-suite CI job (or `Run-AllTests.ps1 -Suite Full` when building Release with `ENABLE_TESTS`), or explicitly document + accept that it's a human pre-merge manual step and add a PR-template checkbox for it. |
| **TW-23** | LOW · PERF | **→ Execute via the DxUi_Uia_ContinuationBaton coalescing pass (GR-A1), with TW-5/TW-24 (2026-07-02).** **`WindowHost`'s `WM_LBUTTONUP`/`WM_RBUTTONUP` handling now unconditionally rebuilds the full accessibility snapshot on every handled click, with no AT-listening gate anywhere in the file (or codebase — `UiaClientsAreListening`/equivalent has zero hits repo-wide).** (`DxUi.WindowHost.cpp:2566-2568`.) Compounds TW-5: every accepted click in grid/tree-heavy dialogs (Preferences, Compare Directories, Connection Manager) now pays a full O(controls) tree walk under a mutex, even with no assistive technology attached. Some call sites (e.g. `Grid`'s column-reorder-commit at `DxUi.Grid.cpp:3114-3115`) already call `RefreshAccessibilitySnapshot()` themselves, making the new WindowHost-level call doubly redundant for those cases. | Fix alongside TW-5/TW-24 as one coalescing/throttling pass: add a dirty flag or a single post-message-loop "refresh once if anything changed" instead of N independent unconditional call sites. |
| **TW-24** | LOW · PERF | **→ Execute via the DxUi_Uia_ContinuationBaton coalescing pass (GR-A1), with TW-5/TW-23 (2026-07-02).** **`Control::SetVisible`/`SetEnabled`/`SetFocusable`/`SetAccessibleName`/`SetAccessibleHelpText` each independently trigger a full snapshot rebuild, with no coalescing against each other or against TW-23's click-driven refresh.** (`DxUi.cpp:145` and sibling setters.) A single click whose handler mutates several sibling controls' state (a common "dependent control enablement" UI pattern) triggers M+1 full tree rebuilds for one user gesture. | Same fix as TW-23 — a single coalesced refresh point instead of per-setter + per-dispatch refreshes. |
| **TW-28** | MEDIUM | **→ ROUTED 2026-07-02 to Granite GR-8 (single owner; same Show-PerfRuns.ps1:448-452 latch — GR-8 is the broader P0 diagnosis incl. compare-mode baseline latch and gauge/distribution taxonomy; the per-metric `minimumSamples` lookup below is part of GR-8's fix).** **`Tools/Show-PerfRuns.ps1 -FailOnQuality` spuriously fails fully-passing perf archives.** `Register-SampleQuality`/`-MinimumSamplesForP95` (default 200, `:24`/`:448-452`) applies one global sample-count threshold to every metric, but `Specs/Testing/FolderViewPerfBudgets.json5` explicitly sets `minimumSamples: 1`-`3` for several deterministic "max"-stat, event-scoped metrics (e.g. `folder.slow_provider.enumeration_us`, `folder.sort_toggle_us`) — the script has no lookup into that file at all. A fully-passing archive with a legitimately-1-sample metric gets marked `P95Quality=fail` and, under `-FailOnQuality`, exits 2 — undermining the one automated quality gate this tool provides and training operators to distrust/ignore its exit code. | Read per-metric `minimumSamples` from `FolderViewPerfBudgets.json5` (or a similar manifest) instead of one hardcoded global default; fall back to the global default only for metrics with no explicit entry. |

### Routed — simplicity / dead code (new, pass 2)

| ID | Sev | Item (current-file anchor) | Fix direction |
|----|-----|--------------------------|----------------|
| **TW-19** | ~~LOW~~ **DONE 2026-07-02** | **Verified complete: `RunArchivePackPromptModalCycle` (`Commands.SelfTest.Dialogs.cpp:196-226`) now contains zero `Trace()` calls, at HEAD and in the working tree** (the file's remaining 36 Trace calls belong to pane_filter tests). Original finding for the record: **Leftover "Pack prompt lifecycle traces" — confirms the continuation baton's own cleanup note.** `RunArchivePackPromptModalCycle` and its worker (`Commands.SelfTest.Dialogs.cpp:198` onward, ~43 unconditional `Trace()` calls total) bracket every micro-step, unlike every sibling `Run*PromptModalCycle` helper in the same file (Rename/CreateDirectory/EditNew/ChangeAttributes/MakeFileList/ArchiveUnpack), which have zero such calls. `Trace()` does real unconditional disk I/O (`AppendUtf16LogLine`), no verbosity guard. | Remove per the baton's own instruction, or convert to a `#ifdef`/verbosity-gated diagnostic if the tracing is still needed for the open Full-suite blocker investigation. |
| **TW-20** | LOW · TEST-ONLY | **Leftover "speed-limit ValuePattern diagnostics" — the other half of the same baton note.** Three new helpers (`DescribeFileOpsValuePatternState`/`...ControlValueStates`/`...PatternStats`, `Commands.SelfTest.FileOps.cpp:84-144`) plus an unconditional `Trace()` + extra UIA descendant-tree walk (`:6263`) fire on every cycle of the speed-limit prompt test. | Same as TW-19 — remove or gate before closeout. |
| **TW-21** | LOW · TEST-ONLY | **`SelfTestLatency::Consume` calls pass a default (never-signals) `stop_token` instead of the real cancellation token, in three places: `FolderView.Icons.cpp:1012` (`ProviderAllowedWork` threadpool callback) and `IconCache.cpp:872`/`:1213` (`ExtractSystemIcon`/`QuerySysIconIndexForPath`, whose signatures have no `stop_token` parameter at all).** A selftest that injects latency via `SetNextDelay` and then triggers teardown/navigation-away expects fast cancellation; instead the full injected delay always elapses on these paths, since the real stop-source (available one layer up) is never threaded through. | For `FolderView.Icons.cpp:1012`: add a `stop_token` field to `ProviderAllowedWork` and capture/forward it. For `IconCache.cpp`: either add an optional `stop_token` parameter to `ExtractSystemIcon`/`QuerySysIconIndexForPath`, or accept that these two call sites are intentionally best-effort (not on a cancelable worker) and document why. |
| **TW-22** | LOW — **DONE 2026-07-05 via Granite GR-12** | **→ ROUTED 2026-07-02 to Granite GR-12 (single owner; same `IconCache.cpp:1235` line — GR-12's insert-only fix, preventing a failure entry from downgrading a concurrent fresh success, subsumes the telemetry ask below).** **`IconCache`'s new failure-cache store path has no duplicate-race telemetry, unlike the pre-existing success-store path.** The success path captures `emplace`'s `inserted` bool and fires `iconcache.duplicate_path_query_race` on a losing race (`IconCache.cpp:1247-1258`); the analogous failure-store `insert_or_assign` (`:1235`) discarded its return pair entirely. | Done in Granite GR-12: the failure store now uses insert-only `emplace(...)` semantics and emits `iconcache.duplicate_path_query_race` when another thread wins the store race. |
| **TW-25** | LOW | **A ~20-line "scroll row into view with 12dip margin" block is duplicated verbatim between `Preferences.FileActions.cpp:2697` (`DebugSelectAssociationRow`) and `Preferences.Themes.cpp:3555-3577` (`ThemesPane::DebugSelectListRow`).** Same variable names, same clamp logic, same `12.0f` margin literal; both call the shared `PrefsPageHost::ScrollTo` helper, so a shared "scroll rect into view" helper clearly belongs in the same `PrefsPageHost` module. | Factor into `PrefsUi::ScrollRectIntoView(hostWindow, state, pageHost, rectDip, marginDip)` or similar, called from both panes. |
| **TW-26** | LOW-MEDIUM | **Diagnostic-only helper functions are eagerly evaluated on every call because C++ has no short-circuit for function arguments.** `describeEditFailure(...)` (`Commands.SelfTest.CompareOptions.cpp:1301`, spins up a `std::jthread` + extra UIA calls) and the analogous helper in `Commands.SelfTest.FileOps.cpp:6270` are passed directly as arguments to `std::format(...)` inside `state.Require(cond, std::format(..., describeEditFailure(...)))` — so they run on every passing assertion too, not just failures. Beyond wasted CPU/thread spin-up, UIA calls can perturb focus state as a side effect, making this a latent source of interference with subsequent same-test assertions. | Guard with `cond ? std::wstring{} : std::format(..., describeEditFailure(...))`, or restructure as `if (!cond) { state.Require(false, std::format(...)); return; }` so the diagnostic collection is genuinely lazy. |
| **TW-27** | LOW | **`focusCategoryTreeHost` diagnostic lambda is duplicated verbatim three times in one file.** Same body/captures/7-field diagnostic format string at `Commands.SelfTest.Preferences.ChromeAndPlugins.cpp:1098` (pre-existing), `:4482`, `:5084` (both added this pass). No shared helper exists anywhere in `RedSalamander/SelfTest/`. | Factor into a shared `SelfTestCommon` helper (e.g. `FocusPreferencesCategoryTreeHost(state, categoryTreeHost, context)`) so a future diagnostic-message fix only needs one edit. |
| **TW-29** | LOW | **`Specs/Testing/FolderViewPerfBudgets.json5` carries a dead `"field"` key in all 14 entries, read by nothing.** `LoadFolderViewPerfBudgetFile` (`Commands.SelfTest.ViewCommands.cpp:16615+`) only reads `case/metric/stat/max/minimumSamples/build/detail/severity/sourceArchive/measured`. Already-observed drift: entries with a `detail` suffix (e.g. `thumbnails.shell_provider_allowed_count`) don't reflect it in their `field` string, proving the two representations can silently diverge with nothing to catch it. | Delete the `field` key from all entries (pure redundant metadata), or if it's meant as a human-readable label, derive it from `metric`+`stat`+`detail` at load time instead of hand-duplicating it per entry. |

### Considered and refuted (pass 2 — recorded so nobody re-investigates)

- **`Preferences.Dialog.cpp:5216` "steals focus to a hidden CompareDirectories control while Plugins is showing"** — the copy-paste wrong-pane call is real (see **TW-15** for the corrected, confirmed mechanism), but this specific claimed *consequence* is refuted: `WindowHost::SetFocusControl`'s ancestor-visibility check (`DxUi.WindowHost.cpp:1528`, via `ResolveControlInteractionState`) ANDs visibility down the whole ancestor chain, so a control under a hidden pane wrapper resolves as not interactive and `SetFocusControl` no-ops. No focus steal occurs; the real bug is the silently-dropped deferred-restore state (TW-15).
- **`Commands.SelfTest.PluginConfig.cpp:3890` "886b3fb71 hang-fix timeout is sized for one inner wait, not the cumulative worst-case chain"** — refuted. `EditPluginConfigurationDialog` is a genuine Win32 **modal** dialog (`DialogBoxParamResourceNoThrow`) that blocks the main thread until `EndDialog`; by the time the main thread resumes at the cited line, the entire preceding worker-phase chain has already elapsed inside the OS modal loop (not raced against the 3000ms gate), and `readyForReopenedDialog` is already `true`. The 3000ms budget only needs to cover the residual gap between the worker's store and the main thread noticing it, which it does.
- **`Commands.SelfTest.Connections.cpp:6625` "asymmetric first/second-click rect-capture retry logic causes the known-failing credential-prompt test"** — refuted. The toggle button's bounds are a fixed layout constant (`kSecretToggleWidthDip = 88.0f`), and `SetSecretVisibility()` never calls `Layout()` — only `OnCreate`/`WM_SIZE`/`WM_DPICHANGED` do, none of which fire on a toggle click. There is no mid-relayout window for the second click's rect to race against. **The actual root cause of the known-failing `cmd_connection_credential_prompt_pointer_click_toggles_secret_visibility` test is still unidentified** — worth a dedicated investigation pass, since two plausible mechanisms (this one, and the DxUi-controls cluster's Toggle/ComboBox accessibility-staleness findings TW-13/TW-14) were both checked and neither explains the specific "focus dropped to None" symptom from the continuation baton.

## How to use this file

- This is a **triage ledger**, not a one-shot executor handoff. Before dispatching an item, expand it into a
  self-contained sub-plan: planned-at SHA, drift check over the in-scope files, in/out-of-scope file list,
  exact verification command, expected output, STOP conditions. Green-check commands are
  `.\Tools\Run-AllTests.ps1 -Suite Full` / `-SkipBuild` (`AGENTS.md:134`, `README.md:204-214`).
- **TW-5/TW-6 are perf-sensitive** and are not complete with a correctness test alone. `AGENTS.md:48-55`
  and `Specs/Testing/Testing_PerformanceValidation.md:55-64,156-166` require scenario definition,
  instrumentation, a deterministic selftest, and archived evidence under `Specs/TestRuns/` (or a documented
  blocked reason) for any UI hot-path change.
- **TW-D1 and TW-D2 are product/UX decisions, not defects.** Do not "fix" them silently. Route to the
  owner, record the decision in the relevant `Specs/<Domain>/` doc, then either WONTFIX or open a scoped
  task. **TW-D1 must be resolved jointly with TW-2** — they are the two halves of one wiring choice.
- Closeout: when finished, move this file to `Specs/Plans/Done/` and fold any durable contract (e.g. the
  cached-first thumbnail behavior, if TW-D1 is accepted) into the authoritative spec, per `AGENTS.md:57-58`.

## Verified correct — do not re-review (skeptic could not break these)

- **CW-1 cross-FS MOVE source-cleanup** (`DeleteCopiedSourceEntryForMove`,
  `FolderWindow.FileOperations.State.cpp:7302-7498`). Captures `sourceDirectoryStillMatches` **once before**
  any deletion (`:7346`), recurses to delete only manifest-matching children, preserves the directory shell
  and new/changed children when the source drifted, and early-skips entirely only for a drifted *reparse*
  link. No data-loss, no duplicate-subtree. (CW-1's outstanding work — reach `master` + add regression
  **CW-T1** — is owned by Clearwater, not here.)
- **Clipboard `Preferred DropEffect` read** (`ReadGlobalDropEffect`, `FolderView.FileOps.cpp`): the branch
  **adds** the `GlobalSize(handle) < sizeof(DWORD)` guard *before* `*effect`, fixing a pre-existing master
  OOB read. The new code is correct.
- **`OnCreateThumbnailBitmap` staleness gating** (`FolderView.Icons.cpp:1535-1546,1622-1626`): re-validates
  `batchId` + `enumerationGeneration` + `itemIndex < _items.size()`, so even a late/abandoned provider
  bitmap can never be applied to the wrong item.
- **`SelfTestLatencyHooks`** (`SelfTest/Common/SelfTestLatencyHooks.cpp`): atomic, `exchange(0)` single-shot
  delay is race-free across concurrent consumers; `Consume` honors the `stop_token`; `ENABLE_TESTS`-only.
- **Icon-cache failure TTL, recall-avoidance key rewrite, `insert_or_assign` failure store, null-brush
  hover/selection skip, refresh→paint metric overwrite-on-second-refresh** — all reviewed and **refuted**
  as non-defects (see "Considered and refuted").

## Routed — correctness / test integrity

| ID | Sev | Item (HEAD anchor) | Fix direction | Proof |
|----|-----|--------------------|---------------|-------|
| **TW-1** | **MEDIUM - DONE 2026-07-05 via Granite GR-24** | **Brush-reuse guard tests were vacuous — the counter had no producer.** `_debugDrawItemTransientBrushCreateCount` (`FolderView.h:1479`) was only zero-init, reset (`FolderView.cpp:1159`), and read into the snapshot (`FolderView.cpp:1108`). Exhaustive grep found no increment anywhere, so the brush-reuse assertions could pass tautologically. | Fixed in Granite GR-24: normal selected/hovered `DrawItem` rendering still uses cached member brushes and must leave the counter at zero, while an `ENABLE_TESTS` forced transient-brush positive control creates a RAII-owned transient brush and increments the same counter. RED build `.build\logs\msbuild-20260705_153748_205.log`; GREEN build `.build\logs\msbuild-20260705_153949_588.log`; GREEN focused archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_154145/`; GREEN adjacent archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_154806/`. | TW-T1 DONE |

## Routed — performance / efficiency

| ID | Sev | Item (HEAD anchor) | Fix direction |
|----|-----|--------------------|---------------|
| **TW-5** | LOW | **→ Execute via the DxUi_Uia_ContinuationBaton coalescing pass (Granite GR-A1: dirty-flag/debounce + `UiaClientsAreListening` gate) together with TW-23/TW-24 — one throttling pass, not three patches (2026-07-02).** **`WindowHost::OnSize` rebuilds the full accessibility snapshot on every WM_SIZE, unthrottled** (new on this branch). `RefreshWindowHostAccessibilitySnapshot(_hwnd, this)` (`Common/DxUi/DxUi.WindowHost.cpp:3697`) `make_shared`s a fresh snapshot and walks the retained control tree twice (`AppendAccessibilitySnapshotNavigation` + `...PointHits`) + `GetWindowTextW`, with no dirty-check and **outside** any bounds-changed guard. WM_SIZE fires continuously during interactive drag-resize; the a11y target is registered eagerly (not lazily on first `WM_GETOBJECT`), so this runs for every window on every resize tick even with no AT client attached — one extra constant-factor O(tree) UI-thread pass, on a *perf* branch. | Coalesce. Minimum: skip when bounds are unchanged. Better: debounce on a short timer or refresh on a settle (`WM_EXITSIZEMOVE`-style) / lazily on next UIA query. Confirm the other 5 call sites (`:1318,1451,1564,1625,2568`) are event-discrete (not per-frame) and need no change. **Perf validation:** add/reuse a `dwrite`/`a11y` snapshot-rebuild count metric; archive a drag-resize selftest before/after under `Specs/TestRuns/` or document a blocker. |
| **TW-6** | LOW | **Per-file icon-cache key fragments by incidental `fileAttributes`.** `PerFileWork` now passes real `work->fileAttributes` (was `0`) with `useFileAttributes=false` (`FolderView.Enumeration.cpp:848`). The returned **icon is unchanged** (the shell ignores attrs without `SHGFI_USEFILEATTRIBUTES`) and this correctly enables the recall-avoidance branch — *not* a correctness bug. Side effect: `fileAttributes` is a `PathQueryKey` member (`IconCache.h:242`), so attribute-bit churn for the same path mints distinct keys → orphan entries + occasional redundant live lookups instead of a cache hit. | Optional efficiency-only: when a lookup resolves live (`useFileAttributes` ends up false and recall is **not** avoided), normalize the stored key's `fileAttributes` to `0` so all live lookups of one path collapse to a single entry. Leave the recall-avoided key (normalized DIRECTORY/NORMAL, `useFileAttributes=true`) as-is. Working-as-intended; fix only if cache-miss telemetry shows churn. |

## Routed — simplicity / dead code

| ID | Sev | Item (HEAD anchor) | Fix direction |
|----|-----|--------------------|---------------|
| **TW-2** | LOW · ARCH | **Provider-allowed-with-deadline thumbnail path is dead code in Release.** `ExtractProviderAllowedThumbnailWithDeadline` (`FolderView.Icons.cpp:941-1084`, ~140 lines of threadpool + deadline/abandoned-delivery concurrency) has exactly one caller — the `ForceProviderAllowedProbe` branch under `#ifdef ENABLE_TESTS` (`:1191-1217`, call at `:1197`). The enum value (`FolderView.h:394`), member, and setter are all `ENABLE_TESTS`-only, but the function **definition** and its declaration (`FolderView.h:1533`) are **not** under `#ifdef`, so they compile into Release as an unreferenced member (stripped only by `/OPT:REF`). `ShellThumbnailLookupMode::ProviderAllowed` is used solely at `:1012` inside this dead function. It is the root of **TW-3**; TW-4's pending-accounting sibling is done. | **Decide intent (jointly with TW-D1).** If provider-allowed-with-deadline is meant to bridge the production cold-cache gap, **wire it** into the production else-branch (`:1219-1251`) behind `allowFileExtraction` + the 50 ms budget (`kProviderAllowedThumbnailWaitMs`). If it is test scaffolding, **move the function + its `ProviderAllowed` enumerant + the `ExtractShellThumbnailBitmap` `ProviderAllowed` handling under `#ifdef ENABLE_TESTS`**. Either choice collapses TW-3. |
| **TW-3** | LOW · TEST-ONLY | **Deadline loop can abandon a just-completed bitmap → needless WIC/fallback re-extract.** The waiter breaks on `now >= deadline` (`FolderView.Icons.cpp:1061-1063`) *before* a final zero-timeout event poll. If the threadpool callback `SetEvent`s and evaluates `if (!abandoned) return;` in that window, the waiter then sets `abandoned=true` too late and returns `WAIT_TIMEOUT`; the extracted bitmap is neither delivered nor late-posted. No UAF (shared `state` is `shared_ptr`; `OnCreateThumbnailBitmap` re-validates). Reachable only via test-only `ForceProviderAllowedProbe`; the one exercising test injects 1500 ms (far from the 50 ms edge) so no observable test impact. | Before breaking, do a final `WaitForSingleObject(state->completed.get(), 0)` and deliver synchronously if signaled; **or** set `abandoned=true` *before* the callback's `if (!abandoned)` check so callback and waiter agree on ownership of a boundary-completed bitmap. Dissolves if TW-2 is gated/rewired. |
| **TW-4** | LOW · TEST-ONLY — **DONE 2026-07-05** | **Abandoned late-delivery over-decremented `pendingBitmapCreates`.** The synchronous path incremented the counter once before posting; the abandoned callback posted a second `kFolderViewCreateThumbnailBitmap` without incrementing, and `OnCreateThumbnailBitmap` decremented for both. Skewed only the `ENABLE_TESTS` thumbnail pending diagnostic. | Done with Granite GR-18: `ThumbnailBitmapRequest::countsPending` defaults true for normal worker posts, abandoned late provider payloads set it false, and the UI-thread apply path decrements only counted payloads. RED/GREEN coverage lives in `folderView_thumbnail_stale_bitmap_messages_preserve_pending_count`. |
| ~~TW-7~~ | — | **ROW RETIRED 2026-07-02 — this pass-1 analysis was proven WRONG by pass 2. Do NOT delete the field.** The folder-path check is *not* sufficient staleness protection (navigate-away-and-back leaves the path equal while `_enumerationGeneration` bumps). See the corrected TW-7 entry in the "Second pass" re-verification section above. Execution owner: **Granite GR-3** (one pass reworks `OnPasteShortcutComplete`: unconditional invalidation/error-reporting + the generation gate). | — |

## Open product / UX decisions — need explicit author sign-off (not defects)

| ID | Disposition | Item (HEAD anchor) | Decision required |
|----|-------------|--------------------|-------------------|
| **TW-D1** | SIGNED OFF 2026-07-04 | **Cached-first thumbnails drop thumbnails for cold-cache non-image files.** Production now calls `ExtractShellThumbnailBitmap(..., CachedOnly)` → ORs `SIIGBF_INCACHEONLY` (`FolderView.Icons.cpp:112-114,1225`); the WIC fallback fires only for raster images (`HasLikelyWicImageExtension`, `:1253`). So `.mp4/.mov/.pdf/.docx/.pptx/.xlsx` in a **cold** directory resolve to a generic icon and are **not** re-queued (`!usedFallback` retry guard) → downgrade persists for the session. Real, visible change vs master (which block-extracted via `GetImage` without `INCACHEONLY`). The only recovery path (TW-2) is unreachable in production. **This is deliberate & spec-mandated:** commit `02e8d0ed5` updated `Specs/UI/UI_FolderView.md` to require cached-only on the visible path and forbid provider-allowed lookups. | Ship cached-only visible thumbnails. Background provider-allowed enrichment is now a required WIP follow-up in `Specs/Plans/WIP/FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md`; TW-2/TW-3 cleanup routes through that plan, while TW-4 is done via Granite GR-18. |
| **TW-D2** | ~~BORDERLINE~~ **SUPERSEDED** | **→ Dissolved 2026-07-02 into Granite GR-1: the "intended UX?" question is moot — GR-1 (HIGH, CONFIRMED) reclassifies the mechanism as a focus trap requiring one-shot restore. No separate sign-off needed.** Original record: **Path-edit blur-suppression widened to a paint-renewed 2000 ms window.** Was one-shot/250 ms; now armed for 2000 ms and **re-armed on every parent paint** while in edit mode (`NavigationView.cpp:1471-1475` `UpdatePathEditHostLayout`; `NavigationView.Edit.cpp:2433-2437` `RefreshActiveEditHostAfterParentPaint`), and the inner self-disarm (`_pathEditBlurSuppressActive=false`) was removed from the kill-focus consumers (`NavigationView.Interaction.cpp:807-819`, `NavigationView.Edit.cpp:759-766`). Correctly cleared on edit-mode exit (`NavigationView.Edit.cpp:567-568,909-910`), so **no permanent stuck-focus**. But it disables click-away-to-dismiss for the whole edit session (snaps focus back to the path edit). Skeptic judged it intentional → refuted as a bug. | Confirm "click-away does not dismiss the path editor; dismiss is Esc/Enter only" is the intended UX. If yes → WONTFIX + a one-line code comment at the arm sites explaining the paint-renewed window. If no → make the suppression decay (don't re-arm on routine paints; arm only on the async host-relayout that motivated it). |

## Tests to add / fix (each must go RED on the un-fixed code)

| ID | For | Required proof |
|----|-----|----------------|
| **TW-T1** | TW-1 | **DONE 2026-07-05 via Granite GR-24.** `folderView_draw_item_brush_reuse_guard` drives selected/hovered FolderView rendering, snapshots `drawItemTransientBrushCreateCount == 0u`, then enables the forced transient-brush positive control and requires the same counter to become nonzero. |
| TW-T5 | TW-5 | (perf) Deterministic drag-resize selftest dispatching N `WM_SIZE` ticks; assert the a11y-snapshot rebuild count is coalesced (≤ small constant, not N). Archive before/after under `Specs/TestRuns/`. |
| TW-T3/T4 | TW-3, TW-4 | TW-T4 is covered by Granite GR-T17's `folderView_thumbnail_stale_bitmap_messages_preserve_pending_count` late-delivery branch. TW-T3 remains only if TW-2 keeps the path reachable in tests: a `ForceProviderAllowedProbe` case injecting latency **at the 50 ms boundary** (e.g. 45–55 ms) asserting a boundary-completed bitmap is delivered instead of falling back. If TW-2 gates the path under `ENABLE_TESTS` only, TW-T3 stays test-scoped. |

## Considered and refuted (skeptic confirmed non-defect — titles only)

- Transient icon-lookup failures suppressed for up to 5 s (intentional, TTL-bounded)
- `insert_or_assign` on the failure path clobbering a concurrent successful icon store (cannot happen)
- Recall-avoidance key rewrite poisoning the live (non-recall) cache entry (distinct keys)
- Hover/selected-text highlight silently skipped when the cached brush is null (cosmetic; never null in practice)
- Clipboard `Preferred DropEffect` OOB read (the branch **fixes** the pre-existing master bug)
- In-flight refresh→paint metric dropped when a second refresh precedes paint (intentional, generation-keyed)
- Refresh→paint metric attributing device-loss-recovery latency to the refresh (intentional)
- Edit-mode blur-suppression 250 ms→2000 ms (real change; judged intended UX → tracked as **TW-D2**, not a bug)
- Release-only perf budgets never applying in the default Debug/ASan selftest build (intentional,
  environment-scoped; **watch item:** ensure CI runs the Release-suite perf gates so they are not silently
  skipped everywhere — confirm in `Specs/Testing/Testing_PerformanceValidation.md`)

## Suggested sequencing

**Updated 2026-07-01 after the second (ultra) pass — TW-8 and the corrected TW-7 now outrank everything
below them; renumber your local worklist accordingly.**

**Routing update (2026-07-02 folder review, deduplicated against Operation_Granite):** Tailwind still owns
TW-8, TW-2/TW-3/TW-D1 (routed here BY Granite), TW-6,
TW-9/TW-10, TW-11, TW-13/TW-14, TW-15, TW-16/TW-17, TW-20, TW-21, TW-25/TW-26/TW-27, TW-29.
Routed AWAY (single owner elsewhere): TW-12+TW-D2 → Granite GR-1; TW-28 → Granite GR-8; TW-22 → Granite
GR-12; TW-7's execution → Granite GR-3 (generation wiring done in GR-3's pass); TW-5/TW-23/TW-24 → the
DxUi_Uia_ContinuationBaton coalescing pass (GR-A1); TW-18 → executed with the Granite GR-8/GR-P3 tooling
cluster. TW-1/TW-T1 is DONE via Granite GR-24; TW-4 is DONE via Granite GR-18; TW-19 is DONE (verified in tree). Steps 9, 10 and 12 below are
therefore executed under their new owners, not from this file.

1. **TW-8 — DONE 2026-07-05 via Granite GR-21**: "Invert Differences Selection" is one-shot again and
   compare-window refresh restores the default decision-model selection.
2. **TW-1 — DONE 2026-07-05 via Granite GR-24**: the draw-item brush-reuse counter is live, normal selected/hovered rendering stays at zero, and the forced positive control proves the guard can fail.
3. **TW-7** (pre-merge, corrected: was miscategorized as trivial dead-field delete — it is not): wire
   `result.generation` into `OnPasteShortcutComplete`'s staleness check. A real async-completion-race bug,
   not cleanup.
4. **TW-9 + TW-10** together (one fix, same root cause): add the `SnapshotGridRowIsSelected` fallback to
   `get_IsSelected`/`get_SelectionContainer`.
5. **TW-13 + TW-14** together (one pattern): move the accessibility-refresh call into the shared
   activation/commit function (`ApplyCheckedState`, `CommitSelection`) instead of the outer dispatcher, so
   keyboard/mnemonic/typeahead paths stop missing it.
6. **TW-11** (menu hover-priming regression) and **TW-15** (Preferences wrong-pane copy-paste) — independent,
   both small, both real UX regressions worth fixing before merge.
7. **TW-16 + TW-17** (test-integrity): restore the two dropped assertion clauses so the selftests catch what
   they were built to catch.
8. **TW-D1 + TW-2** together (one decision): keep cached-first (gate the dead machinery under `ENABLE_TESTS`)
   **or** wire provider-allowed-with-deadline into production. Resolving this collapses **TW-3**; **TW-4** is already done via Granite GR-18.
9. **TW-5 + TW-12 + TW-23 + TW-24** together (one perf/focus-correctness sweep): coalesce all the new
   unthrottled `RefreshWindowHostAccessibilitySnapshot` call sites into a single dirty-flag/debounced refresh,
   and gate blur-suppression renewal on `focusStillBelongsToRoot` while in there. Do under perf-validation
   rules (`AGENTS.md`, `Testing_PerformanceValidation.md`).
10. **TW-18** (perf gate wired into no runner): either wire `--selftest-perf-budget=` into CI/Release-suite
    `Run-AllTests.ps1`, or explicitly document it as a manual pre-merge step.
11. **TW-19 + TW-20** (leftover diagnostic tracing): remove or gate per the continuation baton's own
    instruction — blocks calling the WIP "closed out."
12. **TW-D2** sign-off (folds in TW-12's correctness fix, above); **TW-6, TW-21, TW-22, TW-25, TW-26, TW-27,
    TW-28, TW-29** are all independent low-severity cleanup/efficiency items — batch opportunistically, no
    fixed order.
13. **Unresolved:** the actual root cause of the known-failing
    `cmd_connection_credential_prompt_pointer_click_toggles_secret_visibility` selftest is still not found —
    two plausible mechanisms were checked and refuted in pass 2 (see "Considered and refuted" above). Needs a
    dedicated investigation, not a guess-and-patch.
