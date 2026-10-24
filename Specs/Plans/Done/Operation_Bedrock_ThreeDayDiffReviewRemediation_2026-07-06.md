# Operation Bedrock — Three-Day Diff Review Remediation (2026-07-06)

Remediation ledger for the max-effort review of everything changed **2026-07-03 → 2026-07-06**:
base `275c04034` → master `58b611ece` (Granite remediation `e9e259be5`, Granite archive
`205736167`, Blueprint spec remediation `58b611ece`) **plus the uncommitted working tree**
(perf-measurement-contract Tools slice). Code delta reviewed: 138 files, +15,047/−3,109
(`Specs/` and `*.md` excluded). Full evidence for every item lives in
**`Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md`** (referenced below as *F#n* using that
report's 1–35 numbering); this plan is the execution ledger — work items, fixes, gates, sequencing.

Method: 10 finder angles (5 correctness + reuse/simplification/efficiency/altitude/conventions)
→ 53 raw candidates → 37 after dedup → 37 adversarial verifiers (1-vote, 3-state) → gap sweep
(+3, verified) = **35 surviving findings** (32 CONFIRMED, 3 PLAUSIBLE); 5 refuted (recorded in
the findings doc appendix — **do not re-report or "fix" them**, see the Do-not-touch list at the
bottom). All 64 findings from the 2026-07-05 review were excluded up front; everything here is
new, or documents an incomplete/wrong remediation of a prior finding.

> **Planned-at anchors.** Master `58b611ece`, review base `275c04034`, working tree of
> 2026-07-06 (dirty: `Tools/Run-AllTests.ps1`, `Tools/TestRunPlan.ps1`,
> `Tools/Tests/RunAllTestsPlan.Tests.ps1` — the PerfMeasurementContract slice). All `file:line`
> anchors are verifier-corrected against that tree. Re-check drift with `git diff <anchor-file>`
> before executing any item — this ledger does not auto-track the tree.

> **Coordination caution.**
> - **Track D (MTP)** touches `FileSystemMtp.Core.cpp`/`Device.cpp`, which Evergreen Track A
>   (EV-A1..A5, WPD cache viability) also owns. Execute BR-D items **together with or after**
>   Evergreen Track A on those files; do not create competing partial rewrites of the cache layer.
> - **Track B (DxUi accessibility)** touches files owned by
>   `DxUi_Uia_ContinuationBaton_2026-06-29.md` (coalescing pass). BR-B1/B2 are localized to the
>   new dispatch machinery and safe to land independently; BR-B4 must be coordinated (same
>   rebuild-path lines the baton's gate will touch).
> - **Track G** deletes/rewrites tests added this window; sync with
>   `Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` conventions (behavioral
>   over source-scraping) but the defects are owned here.

## Routing table (single-owner rule)

| Bedrock item | Routed to | Why |
|---|---|---|
| F#19 (Show-PerfRuns flattens `machines[]` without machineHash filter, PLAUSIBLE-LOW) | `Operation_PerfMeasurementContract_2026-07-06.md` | That plan owns analyzer quality gates and the budget-file schema contract; the PS-vs-native reader divergence is a schema-contract question (define whether minimumSamples is per-machine or global) before it is a code fix. Hand the evidence over; do not patch Show-PerfRuns here. |
| F#22 second half (all-columns a11y extraction cost) | `DxUi_Uia_ContinuationBaton_2026-06-29.md` | All-columns extraction is the deliberate, regression-tested Granite GR-9 Narrator fix; any cost reduction (lazy hidden-column synthesis) belongs to the baton's coalescing/gating redesign. Bedrock owns only the identity-vector removal (BR-B4). |
| Systemic flake-harness policy (retry/quarantine) | `Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` | Bedrock owns only defects in code added this window (BR-G items). |

Everything else is owned here. Verified against the WIP README routing map of 2026-07-02: no
other plan tracks these findings. Note BR-D2 modifies the same journal code Evergreen F#12
remediation added — that is intentional (the remediation itself is the defect, F#8 here).

---

## Track A — Plugin module lifecycle (P0: crash class)

Four findings are facets of one missing abstraction: plugin managers are UI-thread-owned,
mutex-less, and now reachable from worker threads and self-referential unload paths. A1 and A2
are the two HIGH crashes; A4 is the structural fix that keeps A3-class divergence from recurring.

### BR-A1 [HIGH][bug] MTP picker worker races the unsynchronized FileSystemPluginManager (F#1)
`RedSalamander/ConnectionManagerWindow.cpp:938` (`MtpPickerWorkerCallback`, queued at `:2386`,
`:2432` via `TrySubmitThreadpoolCallback` — fire-and-forget, no join),
`RedSalamander/FileSystemPluginManager.cpp:1005` (`EnumerateConnectionBrowseDevices` →
`FindPluginById:1105` raw `PluginEntry*`), `:335-407` (`CallConnectionBrowseExport`, unpinned
`entry.module.get()` at `:356`), `FileSystemPluginManager.h:132-136` (no mutex),
`RedSalamander.cpp:8394` (`RefreshRunningPluginsFromSettings` → `Refresh` → `Discover` →
`UnloadAll(FreeLibrary); _plugins.clear()` at `FileSystemPluginManager.cpp:1129-1131`).
- **Problem.** The picker worker reads `_plugins` and executes the browse export through an
  unpinned HMODULE on a threadpool thread while the UI thread can rebuild the vector and
  FreeLibrary the DLL (settings hot-reload, Manage Plugins apply, app exit). The deferred-unload
  guard does not cover browse: `RedSalamanderBrowseConnectionTargets` increments no live counter,
  so `CanUnloadNow` returns TRUE mid-browse. WPD enumeration takes seconds → wide window. UAF +
  execution in unmapped module.
- **Fix (keep the manager single-threaded; don't add a mutex).**
  1. Declare and enforce the threading contract: `FileSystemPluginManager` is UI-thread-only.
     Add `ASSERT_UI_THREAD()`-style debug asserts to its public entry points.
  2. On the UI thread, **before** queueing the worker: resolve the entry once, copy
     `pluginId`/`path`, and create the worker's own module pin —
     `wil::unique_hmodule` from `LoadLibraryExW(entry->path, ..., LOAD_WITH_ALTERED_SEARCH_PATH)`
     (same pattern `CallConnectionBrowseExport` already uses for its module-null transient path
     at `:359`). Resolve the `RedSalamanderBrowseConnectionTargets` proc address up front.
  3. The worker calls the export through its own pin only — it never touches
     `FileSystemPluginManager`. Move the JSON-parse of the browse result to the worker (pure),
     post results back to the window as today.
  4. Teardown: `ConnectionManagerWindow` destruction must not leave workers running against its
     result buffers — either switch `TrySubmitThreadpoolCallback` → `CreateThreadpoolWork` +
     `WaitForThreadpoolWorkCallbacks(work, TRUE)` in the window teardown path, or keep the
     existing heap-owned result contract and verify the posted-message drain already covers it
     (document which).
  5. Hardening (optional, pairs with BR-A2): browse export increments a live-call counter that
     `CanUnloadFileSystemMtpModule` consults, so even a pinned-but-deferred unload stays honest.
- **Verify.** RED-first selftest: start a picker refresh against the selftest fake backend with
  an injected enumeration delay (BR-D4 makes the fixture configurable — do D4 first or use the
  `_DEBUG` fake backend), then drive `RefreshRunningPluginsFromSettings` on the UI thread while
  the worker is in flight; assert no crash and a coherent picker result or clean failure. ASan
  run of `cmd_connection_manager_window_mtp_picker_populates_profile` + plugin refresh.

### BR-A2 [HIGH][bug] MtpBackendCancelRequest keep-alive FreeLibrarys its own DLL (F#2)
`Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:1714-1724` (`backend;` then `moduleKeepAlive;`
declaration order; destructor calls `backend->RequestCancel()`), `Common/Helpers.h:1915`
(`SubmitOwnedThreadpoolCallback` lambda discards `PTP_CALLBACK_INSTANCE`),
`FileSystemMtp.Core.cpp:4050-4056` (`CanUnloadFileSystemMtpModule` checks only quarantine list),
queue sites `:2244` (watchdog), `:3604` (Disconnect), fallback `:1760-1764`.
- **Problem.** Reverse-declaration-order destruction runs `~moduleKeepAlive` (plain
  `FreeLibrary`, executing inside FileSystemMtp.dll) before `~shared_ptr<IMtpBackend>`; when the
  keep-alive holds the last OS reference the return from `FreeLibrary` lands in unmapped memory.
  Member reordering alone does NOT fix it — the lambda epilogue is also module code.
- **Fix.**
  1. Add a `PTP_CALLBACK_INSTANCE`-aware variant of `SubmitOwnedThreadpoolCallback` in
     `Common/Helpers.h` that passes the instance into the callback.
  2. In the cancel-request callback: run `RequestCancel`, destroy/reset `backend` explicitly,
     then as the **last** action `FreeLibraryWhenCallbackReturns(instance,
     request->moduleKeepAlive.release())` and let the request destruct with an empty pin.
  3. Submit-failure fallback (`:1760`): the calling thread is also plugin code, so a plain
     FreeLibrary has the same hazard if it holds the last ref. Deliberately leak the pin there
     (`moduleKeepAlive.release()` + debug log) — matches the ProcessShutdown leave-mapped
     philosophy; a leaked module ref beats UB on an already-failing path.
  4. Hardening: `pendingCancelRequests` atomic counted in `CanUnloadFileSystemMtpModule`, so the
     host does not FreeLibrary while a cancel is queued at all (makes the keep-alive a belt).
- **Verify.** RED-first is impractical to schedule deterministically; instead: (a) unit-level
  fixture that queues a cancel against the fake backend, releases the host module ref, and runs
  under ASan/AppVerifier; (b) assert via the new counter that unload is refused while a cancel is
  pending; (c) existing `mtp_*` cancel selftests stay green.

### BR-A3 [MEDIUM][bug] ViewerPluginManager::Shutdown FreeLibrarys deferred-busy modules (F#4)
`RedSalamander/ViewerPluginManager.cpp:198-200` (`UnloadAll(ProcessShutdown); _plugins.clear();
_deferredUnloadEntries.clear();`), `:1100` (`UnloadAll` iterates `_plugins` only). Model:
`FileSystemPluginManager.cpp:526` (`Shutdown` runs `SweepDeferredUnloadEntries(ProcessShutdown)`).
- **Problem.** Deferred entries (modules that answered `RedSalamanderPluginCanUnloadNow()==FALSE`)
  are destroyed via `wil::unique_hmodule` → FreeLibrary of a busy module at exit. Latent (no
  viewer exports the hook yet) but self-defeating machinery.
- **Fix.** Insert `SweepDeferredUnloadEntries(ModuleUnloadMode::ProcessShutdown);` before the
  clear (the `Unload` ProcessShutdown branch already `release()`s still-busy modules). Subsumed
  by BR-A4 if executed together — do not do it twice.
- **Verify.** Unit-ish selftest with a fake viewer entry whose `canUnloadNow` returns FALSE:
  Shutdown must leave the module mapped (`GetModuleHandleExW` probe), not freed.

### BR-A4 [MEDIUM][reuse/altitude] Extract one shared plugin-module lifecycle component (F#27, F#31)
Copies: `RedSalamander/ViewerPluginManager.cpp:1118-1276` vs
`RedSalamander/FileSystemPluginManager.cpp:1561-1697` (`SweepDeferredUnloadEntries` and
`AddDeferredPlaceholder` token-identical; `Unload`/`IsPluginPathDeferred` near-identical).
Behavioral divergence already shipped: FS manager re-sweeps on demand at lookup paths
(`FileSystemPluginManager.cpp:788-791`); viewer manager sweeps only during `Discover`
(`ViewerPluginManager.cpp:629`) → a deferred viewer DLL stays `ERROR_BUSY` until full
rediscovery. Also F#31: `ViewerPluginManager.cpp:1120-1143` dead always-true `if (entry.module)`
after the `if (!entry.module) return true;` early-out.
- **Fix.** Extract `PluginModuleLifecycle` (header-only or Common/) owning: defer/sweep/
  placeholder semantics, `RedSalamanderPluginCanUnloadNow` probing, `ModuleUnloadMode` handling,
  and the on-demand re-sweep hook; parameterize by entry type or take a minimal `PluginModuleRef`
  view. Both managers call it. Port the FS manager's on-demand sweep to the viewer manager as
  part of the unification (fixes the ERROR_BUSY-until-rediscovery divergence). Delete the dead
  `if (entry.module)` wrapper while touching `Unload`.
- **Verify.** Existing plugin refresh/unload selftests green in both managers; add one viewer-side
  deferred-then-healed case mirroring the FS-side one. The `mtp_identity_helpers_are_shared`-style
  source-contract tests over this area get deleted by BR-G4, so behavior tests carry the load.

## Track B — DxUi accessibility dispatch (assistive-tech correctness)

### BR-B1 [MEDIUM][bug] Signal completion when a UIA dispatch payload is drained undispatched (F#5)
`Common/DxUi/DxUi.Accessibility.cpp:4269` (5s `WaitForSingleObject` on `completedEvent`;
timeout constant `:32`), only `SetEvent` in the handler scope_exit `:7260-7266`,
`Common/DxUi/DxUi.WindowHost.cpp:2905` (`DrainPostedPayloadsForWindow` on WM_NCDESTROY deletes
payload without signaling).
- **Problem.** Window destroyed with a posted dispatch → caller burns the full 5s per action;
  the old `SendMessageTimeoutW` failed fast on dead windows.
- **Fix.** Implement together with BR-B2 as one state machine (see below): the posted-payload
  wrapper that owns a `shared_ptr<AccessibilityUiActionDispatch>` signals `completedEvent` (and
  records `ERROR_CANCELLED`) from its destructor when it is destroyed without having been taken
  by the handler. Do NOT signal from `~AccessibilityUiActionDispatch` itself — the waiter holds a
  shared_ptr ref, so the dispatch destructor can never run while anyone still waits.
- **Verify.** Extend the existing behavioral test
  `TestAccessibilityTextRangeBoundingRectanglesTimeoutKeepsLateHandlerStorageAlive` family with a
  destroy-window-while-pending case asserting the wait returns promptly (<5s, with
  ERROR_CANCELLED), not at the timeout.

### BR-B2 [MEDIUM][bug] Timed-out UIA actions must not execute later (F#6)
`Common/DxUi/DxUi.Accessibility.cpp:7274` (handler executes `ExecuteUiThreadAction` /
`ExecuteSelectOnWindowThread` whenever the message is dispatched), dispatch struct `:82-94`
(no abandoned flag).
- **Problem.** After the dispatcher returns `ERROR_TIMEOUT` the queued payload still executes on
  the UI thread → caller-visible failure with silent later success; retries double-apply.
- **Fix.** Add `std::atomic<uint32_t> state{Pending}` (Pending/Taken/Abandoned) to
  `AccessibilityUiActionDispatch`:
  - Waiter on timeout: CAS Pending→Abandoned; if CAS fails (handler already Took it), do a final
    0-timeout wait on `completedEvent` and return the real result instead of ERROR_TIMEOUT.
  - Handler before executing: CAS Pending→Taken; on failure (Abandoned) skip execution, just
    release storage (scope_exit still signals — harmless, nobody waits).
  - Drain path (BR-B1): CAS Pending→Abandoned + SetEvent from the wrapper destructor.
  This one mechanism resolves F#5 + F#6 and closes the same-shape hole in any future dispatch
  user.
- **Verify.** RED-first: stall the UI thread (existing event-hook trick from the
  bounding-rectangles test) past the timeout with a *mutating* Select action; assert the
  selection did NOT change after the UI thread drains; then a taken-just-in-time case asserting
  no double-execution and a non-timeout return.

### BR-B3 [LOW][reuse] Use ScrollPanel::GetViewportRect instead of the re-implementation (F#33)
`Common/DxUi/DxUi.Accessibility.cpp:814-830` (`ResolveScrollPanelViewportRect`) vs
`Common/DxUi/DxUi.Controls.cpp:6530-6545` (`ScrollPanel::GetViewportRect`, currently private —
`DxUi.h:2722`). Bit-identical today.
- **Fix.** Make `GetViewportRect()` public (it is a pure geometry accessor), delete the helper,
  call the member. No behavior change.
- **Verify.** Build + existing a11y point-hit tests.

### BR-B4 [LOW][simplify] Replace gridAccessibleColumns identity vector with index arithmetic (F#22, Bedrock-owned half)
`Common/DxUi/DxUi.Accessibility.cpp:1220-1224` (iota build), uses `:1263`, `:1319` (loops),
`:1516/:1807/:1855` (`FindSizeValueIndex` linear scans), front/back `:1706/:1744`.
- **Constraint.** The all-columns *extraction* is the deliberate GR-9 Narrator fix
  (regression-tested GR-T9) — do **not** revert to visible-only columns. Only the redundant
  vector goes.
- **Fix.** Delete the member; loops become `for (size_t col = 0; col < record.gridColumnCount;
  ++col)`; `FindSizeValueIndex(gridAccessibleColumns, col)` → `col < gridColumnCount ?
  std::optional(col) : std::nullopt`; `.front()` → `0`; `.back()` → `gridColumnCount - 1`;
  ordinal ±1 → col ±1. Byte-identical behavior.
- **Coordinate** with the DxUi UIA baton (same lines its coalescing gate touches); land as a
  standalone mechanical commit the baton can rebase over.
- **Verify.** Existing GR-T9 + grid a11y tests green.

## Track C — Settings & themes (data loss)

Execute C2 → C1 (C2 creates the extraction seam C1's lenient mode lives in). C3 independent.

### BR-C1 [MEDIUM][bug] One bad color must not delete the whole inline theme (F#3)
`Common/Common/SettingsStore.cpp:1191-1201` (whole-theme `continue` on
`ParseThemeDefinitionJson5` failure), `Common/Common/ThemeDefinitionIo.cpp:326-329` (one bad
color → `InvalidData` for the entire theme). Old inline parser (deleted side of the hunk in the
review diff) skipped only the bad entry.
- **Problem.** One typo'd color value, unknown color key (forward-compat), >64-char name,
  non-builtin `baseThemeId`, or invalid id silently drops the whole `ThemeDefinition`; the next
  settings save serializes the in-memory list and permanently erases the theme from disk.
- **Fix.** Add a parse mode to the shared extractor from BR-C2: `Strict` (theme files — current
  behavior) and `LenientInline` (settings.json), where:
  - invalid/unknown color entries are skipped with one aggregated `Debug::Warning` naming theme +
    keys, theme kept;
  - over-long names are clamped, not fatal;
  - unknown `baseThemeId` keeps the theme and falls back to the default base at apply time;
  - only a structurally unusable theme (no id) is dropped — and a dropped theme must NOT be
    silently re-persisted away: keep unparseable raw entries in an opaque preserved list that the
    settings serializer round-trips verbatim.
- **Verify.** RED-first LocalizationTests/SettingsStore test: settings.json with a theme carrying
  one bad color → loads minus that color; save; reload; theme still present with all other
  colors. Second case: structurally broken theme survives a load→save round-trip untouched.

### BR-C2 [LOW][efficiency] Parse inline themes from the parsed DOM — no serialize/re-parse round-trip (F#23)
`Common/Common/SettingsStore.cpp:1177/1183/1191` (ConvertYyjsonToJsonValue → SerializeJsonValue →
ParseThemeDefinitionJson5), `ThemeDefinitionIo` re-tokenizes into a third DOM.
- **Fix.** Split `ParseThemeDefinitionJson5` into text-front-end + `ParseThemeDefinitionFromValue
  (const Common::Settings::JsonValue&, ThemeDefinition&, mode, errors...)`. SettingsStore calls
  the value-based extractor directly on the converted `JsonValue`; the theme-file path keeps the
  text front-end. This is the seam BR-C1's mode flag lives on.
- **Verify.** Existing theme-file tests green (Strict path unchanged); startup/hot-reload behavior
  identical for valid themes.

### BR-C3 [MEDIUM][bug] Restore a hot-reload readiness guarantee without re-blocking Start (F#7)
`RedSalamander/SettingsHotReload.cpp:561-564` (Start returns immediately after spawning
`WatchSettingsDirectoryThread`), `:521` (readyEvent created), `:140` (signaled — nobody waits),
watcher retry backoff ~`WatchSettingsDirectoryThread:130` (1000ms stop-event wait on transient
arm failure).
- **Problem.** The prior 2s-blocking-Start finding was remediated by deleting the readiness wait
  entirely: writes landing before `FindFirstChangeNotificationW` arms (startup latency or the 1s
  retry window) are silently missed; `readyEvent` is dead write-only machinery. The new selftest
  passes only because it sleeps `Scale(1200ms)` first.
- **Fix (async catch-up, not a blocking wait).**
  1. In `Start`, capture the settings file's `(lastWriteTime, size)` stamp before spawning the
     thread and pass it in.
  2. In the watcher, immediately **after each successful arm** (first arm and every re-arm after
     a transient failure), re-stat the file; if the stamp differs from the last-delivered stamp,
     post `WndMsg::kSettingsFileChanged` once. This closes both the startup and the retry-backoff
     windows with no blocking.
  3. `readyEvent`: either repurpose it as a test-only `DebugWaitForSettingsWatcherReady()` hook
     (preferred — lets the selftest drop its `Scale(1200ms)` sleep) or delete the plumbing.
- **Verify.** RED-first: rewrite `settings_hot_reload_transient_arm_failure_is_async` to save the
  file immediately after `Start` returns (no sleep) and assert the change is delivered; keep a
  transient-arm-failure case asserting delivery after re-arm via the catch-up stamp.

## Track D — MTP plugin correctness

Coordinate with Evergreen Track A (same files). D1/D3/D4 are independent of the cache work; D2
touches the journal layer that Evergreen's F#12-remediation added.

### BR-D1 [MEDIUM][bug] devicePuid must fall back to pnpId on WPD read failure (F#9)
`Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:2316` (`devicePuid.clear()` first line; all
failure paths return empty), caller `:~2385` (sets `devicePuid = descriptor.pnpId` before the
call, ignores the HRESULT).
- **Fix.** Make the callee write to a local and assign only on success, or have the caller apply
  the fallback after the call: `if (FAILED(hr) || devicePuid.empty()) devicePuid =
  descriptor.pnpId;`. Pick one (recommend the caller-side check — it also covers a success that
  returns empty), delete the callee's pre-clear of caller state.
- **Verify.** Fixture case: PUID read fails → browse JSON `devicePuid` equals pnpId; profile
  round-trip keys on it.

### BR-D2 [MEDIUM][bug/altitude] Fix the overwrite-journal absent-cache at the journal layer (F#8)
`Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:2190` area; `MarkOverwriteJournalAbsent:1352`
(check-then-act), `RecordOverwriteJournalIntent:1075-1112` (invalidate+write), plumbing
`FileSystemMtp.h:243-245` (`RunBackendCommand(..., bool cacheOverwriteJournalAbsence = true)`) →
`Submit:1810-1832` → `QueuedCommand:1886` → `OverwriteJournalContext.useAbsentCache`; six
mutating call sites pass `false` (e.g. `:2947`, `:3203`).
- **Problem.** Three defects in one remediation: (1) the cross-instance race survives — a reader
  can mark the identity absent after another instance recorded a live journal → replays skip the
  live journal (the original F#12-2026-07-05 bug); (2) the per-call-site positional `false` is
  fragile — any future mutating call site that forgets it silently regresses; (3) mutating
  commands bypass the absent cache entirely → per-item journal-path resolve +
  `EnsureDirectoryExists` + failed-open probe (1000-item delete = 1000 probes).
- **Fix.**
  1. **Generation counter inside the journal layer**: a per-device-identity atomic generation in
     the shared journal registry. `RecordOverwriteJournalIntent` bumps it (before invalidate).
     `MarkOverwriteJournalAbsent(identity, generationObservedBeforeRead)` only stores absent if
     the generation is unchanged — turning check-then-act into a validated publish. (A
     per-identity mutex around read-probe+mark is an acceptable alternative; the generation form
     avoids holding a lock across file IO.)
  2. Delete the `cacheOverwriteJournalAbsence` parameter end-to-end. The queue derives policy
     from the command itself: add `bool isMutating` to the command descriptor (it already knows
     its kind) — and with (1) in place the absent cache becomes safe for mutating commands too
     (`WriteOverwriteJournalEntry` already calls `InvalidateOverwriteJournalAbsent`), which fixes
     (3) with no per-site flags at all.
- **Verify.** RED-first race test on the fixture: instance A reads (journal absent), instance B
  records intent + writes journal, A marks absent with its stale generation → assert the mark is
  rejected and a subsequent replay sees B's journal. Perf check: bulk fake-backend delete of N
  items performs O(1) journal probes (counter assert), not O(N).

### BR-D3 [PLAUSIBLE-LOW][bug] Validate sizeBytes before writing through the result pointer (F#17)
`Plugins/FileSystemMtp/Factory.cpp:273-277` (`result->jsonUtf8 = nullptr;` precedes the
`sizeBytes != sizeof(...)` check; `jsonUtf8` at offset 8 — `Common/PlugInterfaces/Factory.h:56-70`).
- **Fix.** Reorder: null-check `result`, validate `result->sizeBytes`, only then touch other
  fields. Audit the other exports in `Factory.cpp` for the same pattern while there.
- **Verify.** Compile + existing browse selftests; this is a contract fix, no behavior change for
  well-formed callers.

### BR-D4 [LOW][bug] Stop discarding the selftest WPD-cache fixture options (F#18)
`Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:2539-2543` (`static_cast<void>(optionsJsonUtf8);
return std::make_unique<WpdMtpBackend>(true);`), export forwards the JSON
(`Factory.cpp:372/386`), fake-backend twin parses it (`FileSystemMtp.FakeBackend.cpp:331`,
e.g. `readFileDelayMs`).
- **Fix.** Parse the same option keys the fake backend honors where they apply to the memory
  fixture (`readFileDelayMs`, tree-shape/failure knobs as available) and plumb them into
  `WpdMemoryBackend`; for any unsupported non-empty option, fail creation with `E_INVALIDARG` +
  debug log rather than silently ignoring. (BR-A1's RED test wants the delay knob.)
- **Verify.** Fixture test: create via export with `readFileDelayMs`; measure the read takes ≥
  the delay; unsupported key → E_INVALIDARG.

## Track E — FolderView icons, rendering, metrics

### BR-E1 [MEDIUM][bug] pendingBitmapCreates leaks on stale-generation messages (F#10)
`RedSalamander/FolderView.Icons.cpp:1315` (post-side batch-only guard), `:1330` (increment,
`countsPending` defaults true — `FolderView.h:797`), `:1515-1518` (handler stale-generation
early-return fires before the scope_exit decrement is armed).
- **Fix.** In the handler, order as: (1) stale-**batch** check first (no decrement — the counter
  was reset on batch bump); (2) arm the `countsPending` scope_exit decrement; (3) stale-generation
  check and all later returns now decrement correctly. One-line move of the scope_exit above the
  generation check.
- **Verify.** RED-first: seed a pending create (existing `DebugSeedThumbnailPending...` hook),
  bump the enumeration generation without a batch bump, deliver the message, assert
  `pendingBitmapCreates == 0` in the debug snapshot.

### BR-E2 [MEDIUM][bug] Clear the input_to_paint metric on failed frames, like refresh_to_paint (F#11)
`RedSalamander/FolderView.Rendering.cpp:1930-1932` (EndDraw fail), `:1972` (Present1 fail),
`:2012` (legacy Present fail), `:979` (no-render-target return) — all call
`ClearReadyPendingRefreshToPaintMetric()` (resets only `_pendingRefreshToPaintMetric`,
`FolderView.cpp:156-162`) while `_pendingInputToPaintMetric` stays armed and emits at the next
successful present (`FolderView.cpp:112-122`, emit sites `:1985/:2025`).
- **Fix.** Generalize: `ClearPendingPaintMetricsOnFailedFrame()` clearing both pendings; call it
  at all four sites (replaces the refresh-only call). Keep the two metrics' arm/emit sites
  otherwise untouched.
- **Verify.** RED-first twin of `folderView_refresh_to_paint_metric_clears_after_failed_render`
  for `folder.frame.input_to_paint_us`: force EndDraw failure after arming input metric → next
  present must NOT emit the inflated sample.

### BR-E3 [MEDIUM][efficiency] Reinstate bounded negative caching for live icon-path failures (F#20)
`RedSalamander/IconCache.cpp:1239-1245` (live failures return `std::nullopt` uncached, counter
`iconcache.path_live_lookup_failed_uncached`); attribute-mode failures still use the bounded
negative cache (`:1247-1264`, `kPathIconFailureTtl{5}` at `:40`). Hot caller:
`FolderView.Enumeration.cpp:848` (threadpool workers, per file, per re-enumeration).
- **Problem.** The 5s-failure-cache remediation over-corrected to *no* caching: every
  directory-change refresh re-issues blocking `SHGetFileInfoW` per persistently-failing path
  (dead links / offline shares — seconds each).
- **Fix.** Cache live-lookup failures with per-path escalating backoff: first failure short
  (e.g. 2s), doubling to a cap (e.g. 30s), reset on success; keep the existing counter, add a
  `_cached` variant. Do NOT restore the flat always-5s behavior the prior finding objected to.
- **Verify.** Selftest with a fixture path whose live lookup fails: N refreshes inside the
  backoff window perform 1 shell call (counter assert); after the window a retry occurs; success
  clears the negative entry.

### BR-E4 [LOW][bug] Close the thumbnail abandoned-handshake window (F#16)
`RedSalamander/FolderView.Icons.cpp:1006-1015` (worker: store result → SetEvent → load
`abandoned`), `:1074-1089` (waiter: timeout → deadline break → store `abandoned=true` → return),
test-only ForceProviderAllowedProbe path.
- **Fix.** After the waiter sets `abandoned=true`, do one final `WaitForSingleObject(completed,
  0)`; if signaled, un-abandon (store false) and consume the result normally. The worker side
  stays as-is (SetEvent-then-check is then safe: either the waiter's final check sees the event,
  or the worker's `abandoned` load sees true and posts the late upgrade).
- **Verify.** Existing probe selftest; the interleaving itself is timing-bound — assert via the
  handshake's step counters (both-sides-skipped becomes impossible by construction).

### BR-E5 [MEDIUM][altitude/test] Replace the self-fulfilling DrawItem brush-count test hook (F#26)
`RedSalamander/FolderView.Rendering.cpp:2300-2314` (ENABLE_TESTS lambda force-creates
`ID2D1SolidColorBrush` when `_debugForceDrawItemTransientBrushCreateForSelfTest` is set; `:2311`
is the counter's ONLY increment site).
- **Problem.** The metric exists to catch per-item brush allocation regressions, but its only
  incrementer is the test's own forced path — it can never catch a real regression, and DrawItem
  carries permanent scaffolding.
- **Fix.** Invert the assertion: delete the force-create hook and its flag; instrument the real
  brush-creation seam (the D2D brush factory/cache wrapper used by rendering) to increment the
  counter on any transient (non-cached) brush creation during item draw; the selftest asserts the
  counter stays **0** across a scroll/redraw pass. That is the regression guard the metric was
  meant to be.
- **Verify.** New assertion green on current code; hand-inject a temporary `CreateSolidColorBrush`
  into DrawItem locally to confirm the test goes red (do not commit the injection).

## Track F — Search service & broker

### BR-F1 [MEDIUM][efficiency] Stop re-opening SQLite per keystroke for the generation probe (F#15)
`Common/LocalSearchIndexCore.cpp:4385` (`ReadStoreGeneration` on the steady-state path; guard at
`:4381` has no TTL/mtime check; called from `:4694` and `:4771` per `Enumerate`, `:4889` per
`EnumerateNoWait`), `Common/SqliteIndexStore.cpp:2353-2354` (fresh `OpenReadOnlyReadyConnection`
per call → `EnsureReadableSchema` ~11 probes, `:1494`).
- **Fix (layered, cheapest-first).**
  1. Stat pre-check: capture `(mtime,size)` of the DB and `-wal` at last validation; if both are
     unchanged, skip `ReadStoreGeneration` entirely.
  2. Within one `Enumerate` call, probe at most once (pass the result from `:4694` down to
     `:4771`).
  3. Optional: cache a read-only probe connection in the per-store cache entry, invalidated on
     open failure or generation change (schema re-validation only on reopen).
- **Verify.** Counter/trace assert in the search selftests: steady-state repeated queries perform
  0 sqlite opens after warm-up; generation bump (index rebuild) is still detected within one
  query.

### BR-F2 [MEDIUM][efficiency] Cache transient authorization failures for the batch (F#21)
`Common/SearchServiceBroker.cpp:2332-2343` (durable-only caching; durable list `:2285-2296` =
ACCESS_DENIED/FILE_NOT_FOUND/PATH_NOT_FOUND; transient path logs + returns uncached).
- **Fix.** `state.clientDirectoryAccessCache.emplace(cacheKey, false)` (or a parallel
  transient-verdict set if the distinction matters for stats) before the transient return, and
  demote the per-candidate `Debug::ErrorWithLastError` to once-per-parent-per-batch. Cache
  lifetime is already a single query batch, so nothing is lost.
- **Verify.** Broker selftest: N candidates under one ERROR_BAD_NETPATH parent → exactly 1
  impersonate+CreateFileW attempt and 1 log line (counter assert).

### BR-F3 [MEDIUM][altitude] Move impersonation fault injection behind the test seam (F#25)
`Common/SearchServiceBroker.cpp:2310-2316` (inline `#if _DEBUG || !NDEBUG || ASAN` inside
`CheckClientCanListDirectory`), same gate copy-pasted at `SearchServiceBroker.h:190`,
`SearchServiceBroker.cpp:544`, `:2801`, `RedSalamanderSearchService/Main.cpp:3179` (CLI flag).
- **Fix.** (1) One named macro (e.g. `RS_SEARCH_TEST_HOOKS`, defined from the existing
  ENABLE_TESTS family in the service/broker projects) replaces the five ad-hoc gate copies.
  (2) Replace the inline branch with an injectable failure hook owned by `ServerOptions`
  (`std::function<HRESULT(FaultPoint)>` or a small enum-dispatched struct), set only by the
  test-flag parsing path — the authorization function body stays clean, and future fault points
  register instead of adding `#if`s. Keep the CLI flag under the same macro.
  (3) Sequence with BR-G2 (the selftest that passes the flag must share the same gate).
- **Verify.** `search_service_candidate_impersonation_failure_is_incomplete_warning` green in
  test-hook builds; release binary contains neither flag string nor hook (string search in the
  build output as a one-off check).

### BR-F4 [LOW][simplify] One accumulator for ACCESS_DENIED_SKIPPED (F#30)
`Common/SearchServiceBroker.cpp:2011` (direct `state.stats->warningFlags` write — dead on the
completion path because `*outStats = stats` wholesale at `LocalSearchIndexCore.cpp:4719/4812/5143`
precedes the `:2838` merge), `:2013` (`state.warningFlags`), `:2838` (merge), `:1960` (mid-run
QUERY_BATCH_SENT reads stats).
- **Fix.** Delete the `:2011` write; at the flush site (`FlushServerCandidates` /`:1960` batch
  header) emit `QueryStatsWarningFlags(*state.stats) | state.warningFlags` so mid-run events keep
  seeing the flag. One accumulator (`state.warningFlags`), same wire bytes.
- **Verify.** Existing warning-flag selftests (incomplete-warning case from BR-F3's test) green;
  add an assert that the batch header carries the flag mid-run.

## Track G — Test harness & release-build gates

### BR-G1 [MEDIUM][test-bug] MTP picker selftest: skip on missing debug export (F#12)
`RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp:2135-2137` (`state.Require
(SUCCEEDED(fakeBrowseHr))`), helper `:55-57` returns `ERROR_PROC_NOT_FOUND`, export `_DEBUG`-only
(`Plugins/FileSystemMtp/Factory.cpp:327-328`). Skip pattern model:
`CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp:393`.
- **Fix.** `if (fakeBrowseHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND)) { state.Skip(...); }`
  before the Require, matching the established pattern. Longer term the export gate should join
  the BR-F3 macro family — note it there, do not fork the gate here.
- **Verify.** Release-flavor suite run: case reports SKIP, not FAIL.

### BR-G2 [MEDIUM][test-bug] Impersonation selftest: gate parity with the service flag (F#13)
`RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp:10012-10013`
(unconditional `--test-fail-client-auth-impersonation-once` + `state.Require(service.Start)`),
service parse gate `RedSalamanderSearchService/Main.cpp:3179`.
- **Fix.** Guard the case body with the same compile-time family the service uses (after BR-F3:
  `RS_SEARCH_TEST_HOOKS`) → `state.Skip` otherwise. Test binary and service build in the same
  configuration in this repo's runs; note that assumption in a comment.
- **Verify.** Release-flavor suite run: SKIP, not service-start hard-fail.

### BR-G3 [PLAUSIBLE-MEDIUM][test-bug] Bounded context-menu debug probe (F#14)
`Common/DxUi/DxUi.Menu.cpp:5291-5293` (unbounded cross-thread `SendMessageW` for
`kMenuDebugGetStateMessage`; replaced `SendMessageTimeoutW(1000ms, SMTO_ABORTIFHUNG|SMTO_BLOCK)`
which had a dangling stack-output-write hazard — GR-17).
- **Fix.** Port the BR-B1/B2 pattern (it exists precisely for this): heap-owned
  `shared_ptr` request `{outState copy, completedEvent, atomic state}` posted (or sent) to the
  menu thread; probe thread waits on the event with a bounded timeout scaled by the selftest
  multiplier; Taken/Abandoned CAS prevents the handler writing into abandoned storage — the
  GR-17 hazard stays fixed AND the bound returns. Update the pinning source-contract test
  (`TestContextMenuDebugStateCrossThreadQueryDoesNotUseTimedOutStackStorage`) — or delete it per
  BR-G4 and let the behavioral test carry it.
- **Verify.** RED-first: wedge the menu thread (existing wedge hook from the menu tests), probe →
  returns failure within the bound instead of hanging; unwedged probe returns correct state.

### BR-G4 [MEDIUM][test-quality] Remove the source-scraping contract tests; keep behavioral twins (F#24)
`Tools/Tests/TestHarnessSourceContracts.Tests.ps1` (+382 lines / 18 It blocks this window; exact
`#pragma` wording pins at `:548-549`), `Tests/DxUiTests/DxUiTests.WindowHost.cpp:2535`
(`TestDxUiDeviceLossAndD3dCreationAreShared`), `DxUiTests.Menu.cpp:393`
(`TestContextMenuDebugStateCrossThreadQueryDoesNotUseTimedOutStackStorage`),
`DxUiTests.Accessibility.cpp:72` (`TestAccessibilityUiActionDispatchOwnsTimedOutRequestStorage`),
`CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp` (`mtp_identity_helpers_are_shared`).
- **Fix.** Inventory every source-text assertion added this window; for each: (a) if a behavioral
  twin exists (the a11y dispatch storage test has
  `TestAccessibilityTextRangeBoundingRectanglesTimeoutKeepsLateHandlerStorageAlive`; BR-B2/G3 add
  more), delete the scrape; (b) if the invariant is real but untested behaviorally, write the
  behavioral test first (BR-B2, BR-G3, BR-E5 supply the main ones), then delete the scrape;
  (c) stylistic pins (pragma comment wording, member ordering, function ordering via IndexOf)
  are deleted outright — and normalize the "unreferenced"/"unused" pragma wording divergence the
  test currently enshrines (`FileSystemMtp.h:21` vs `Factory.cpp:17`) to one spelling.
  Coordinate deletions with items here that must first update a pinned form (BR-G3, BR-G5,
  BR-D2's flag removal is pinned by a worker_shutdown-adjacent contract — check before deleting).
- **Verify.** Pester suite green; grep confirms no `Get-Content` of `.cpp` product sources
  remains in tests added this window (pre-existing scrapes are out of scope — the stabilization
  plan owns the policy ratchet).

### BR-G5 [LOW][reuse] One truthy-env parser for ViewerSpace flags (F#32)
`RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp:19492-19507`
(`IsViewerSpaceEnvFlagEnabled`, exact-match accept-set) vs `Common/Helpers.h:278-292`
(`EnvironmentVariables::IsTruthyFlagSet`, first-char accept-set) — divergent semantics
(`=y` works for FolderView flags, not ViewerSpace flags), duplicate pinned by a
TestHarnessSourceContracts block.
- **Fix.** Point `IsViewerSpaceLargePerfEnabled`/`IsViewerSpace20kLayoutEnabled` at
  `EnvironmentVariables::IsTruthyFlagSet`; delete the local helper; update/delete the pinning
  contract test (BR-G4).
- **Verify.** ViewerSpace-gated selftests still trigger with `=1` and now also `=y` (parity with
  FolderView flags).

### BR-G6 [MEDIUM][reuse] One repo-root walker for the selftest binary (F#29)
New 4th copy `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp:3874`
(markers: `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp` + `RedSalamander/RedSalamander.vcxproj`);
existing: `SelfTestCommon.cpp:343` (`TryFindRepoRoot`, anonymous namespace, `.git` +
`RedSalamander.sln` markers), `Commands.SelfTest.PluginConfig.cpp:4723`,
`Commands.SelfTest.Connections.cpp:185/215`.
- **Fix.** Export `SelfTest::TryFindRepoRoot` from `SelfTestCommon.h` with the union marker rule
  (`.git` OR `.sln` OR `RedSalamander.vcxproj` — decide once, comment why) and the env override;
  delete the three per-TU walkers and point call sites at it.
- **Verify.** Selftests that read repo files pass from (a) normal checkout, (b) a `git worktree`
  layout (the divergence trigger), (c) `REDSALAMANDER_REPO_ROOT` override.

### BR-G7 [LOW][reuse] Shared MTP selftest-export loader (F#35)
`Commands.SelfTest.Connections.cpp:33-57` (new copy) vs
`CompareDirectoriesEngine.SelfTest.cpp:~995-1105` (two copies) vs
`FileSystemPluginManager.cpp:335-408` (production sibling with the richer guard set —
unloadDeferred/ERROR_BUSY, module reuse — that the selftest copies lack).
- **Fix.** One `SelfTest::CallMtpPluginExport<Fn>(exportName, ...)` helper in SelfTest common:
  entry lookup (disabled/loadable/path guards + `IsPluginPathDeferred` check like production),
  `LoadLibraryExW` pin, pragma-4191 `GetProcAddress`, typed call. Replace the three copies.
- **Verify.** MTP selftests green in debug flavor; release flavor hits the BR-G1 skip path.

## Track H — Shared-helper hoists (production reuse)

### BR-H1 [MEDIUM][reuse] JsonValue member accessors in Common::Settings (F#28)
New suite `RedSalamander/FileSystemPluginManager.cpp:87-158`; duplicates:
`Commands.SelfTest.Connections.cpp:562+`, `ConnectionProfileUtils.cpp:66-110`
(`ExtraGetBool/ExtraGetUInt32`), `ConnectionManagerWindow.cpp:444` (`ExtraGetString`),
`HostServices.cpp:1399` (`ExtraGetString`). Root cause: `Common::Settings` exposes no member
lookup next to `ParseJsonValue` (`SettingsStore.h:864`).
- **Fix.** Add to `Common/Common/SettingsStore.h` (namespace `Common::Settings`):
  `FindMember(const JsonValue&, string_view) -> const JsonValue*`, plus typed
  `GetString/GetWString/GetBool/GetUInt32/GetArray` accessors with the overflow/UTF-16 behavior
  of the FileSystemPluginManager copy (it is the newest and strictest). Migrate the five sites;
  keep behavior identical (the selftest copy's laxer bool handling — verify no test depends on
  it before unifying).
- **Verify.** Connection browse + profile round-trip selftests green; RedConfigure/settings tests
  green (header is widely included — build-time check).

### BR-H2 [LOW][reuse] Hoist Utf16FromUtf8 (F#34)
New copy `RedSalamander/FileSystemPluginManager.cpp:64`; siblings at
`ConnectionManagerWindow.cpp:411`, `ManagePluginsDialog.cpp:226`, `Preferences.Keyboard.cpp:1641`,
`Preferences.Plugin.Configuration.cpp:70`, `FolderWindow.ItemProperties.cpp:97`,
`FolderWindow.FileOperations.cpp:35`, `Common/Common/SettingsStore.cpp:93`, …(~11 total).
- **Fix.** One `Common::Strings::Utf16FromUtf8(std::string_view)` in `Common/Helpers.h` (flags:
  0 — note two copies differ on `MB_ERR_INVALID_CHARS`; standardize on flags 0 with the
  documented replacement-char behavior unless a call site provably needs strictness). Migrate at
  minimum the two copies added this window (FileSystemPluginManager, ConnectionManagerWindow);
  sweep the rest opportunistically per file touched.
- **Verify.** Build + spot selftests over connection browse and item properties.

---

## Sequencing

1. **P0 crash class first — Track A** (BR-A2 is self-contained; BR-A1 needs BR-D4's fixture knob
   for its RED test, or use the `_DEBUG` fake backend and land the knob after). BR-A3/A4 ride
   behind as one refactor commit.
2. **Data loss — BR-C2 → BR-C1**, then **BR-C3** (independent).
3. **Release-build suite gates — BR-G1, BR-G2** (two small skip guards; unblocks release-flavor
   suite runs immediately; BR-F3's seam consolidation can follow without blocking them).
4. **Metrics/counters before more perf hardening lands — BR-E1, BR-E2** (they corrupt the perf
   evidence the active PerfMeasurementContract work depends on), then **BR-E3**.
5. **A11y dispatch — BR-B1+B2** as one commit; **BR-G3** reuses the pattern immediately after;
   then **BR-B3/B4** (B4 coordinated with the UIA baton).
6. **MTP — BR-D1..D4** aligned with Evergreen Track A scheduling on the same files.
7. **Search — BR-F1..F4.**
8. **Cleanup tail — BR-G4** (after B2/G3/E5 supply behavioral replacements), **BR-G5..G7,
   BR-H1..H2** — mechanical, land opportunistically per file touched by earlier tracks.

Quick wins if a short session wants standalone items: BR-G1, BR-G2, BR-D1, BR-D3, BR-E1, BR-E2,
BR-B3, BR-F4.

## Verification conventions

- RED-first where the ledger says so: write the failing test against current code, land fix,
  keep the test. Use the selftest timeout multiplier for anything timing-adjacent; no bare
  sleeps (BR-C3 explicitly removes one).
- No new source-scraping tests anywhere in this plan — behavioral or nothing (BR-G4 is the
  ratchet; do not regress it while fixing).
- Items touching WarpDrive-hot files (`FolderView.*`, `DxUi.*`) re-run the FolderView perf
  smoke (`Tools/Show-PerfRuns.ps1` gate) after landing — BR-E2/E3 change metric emission
  semantics deliberately; expect and annotate the evidence delta.

## Do-not-touch list (refuted in review — resist "fixing" these)

- `Plugins/FileSystemS3/FileSystemS3.h:343` `write:false` — that block is `kCapabilitiesJsonS3Table`
  (read-only S3 Table), intentional Blueprint alignment; main `kCapabilitiesJsonS3` still
  advertises write (`:309/:318`).
- MTP picker storage names — disambiguation lives in `EnumerateObjectItems`
  (`FileSystemMtp.Device.cpp:845/:863`); nothing was dropped.
- `FileSystem.Search.cpp` E_INVALIDARG fallback — intentional, documented GR-13 behavior
  (already refuted in the 07-05 review appendix; refuted again here).
- Destination bridge IO decorator no-op (`FolderWindow.FileOperations.State.cpp:6479`) —
  contract-pinned scaffolding owned by `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`
  FIR-1 / Floodgate FG-A1. Leave it.
- `QueueBackendCancel` fallback double-cancel — does not exist (`request->backend.reset()` at
  `FileSystemMtp.Core.cpp:1762` guards the destructor path). BR-A2 touches this code; keep that
  property.

## Closeout

All 30 owned items checked off (or explicitly re-routed with a pointer), the three routed items
handed over, findings doc cross-referenced from each landed commit message, and this file moved
to `Specs/Plans/Done/` with a closeout note. Update the WIP README row on every session that
lands a track.
