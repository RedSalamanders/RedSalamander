> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# Three-Day Diff Review — 2026-07-06 Findings

- **Scope**: `git diff 275c0403..working-tree` (2026-07-03 → 2026-07-06), code only (`Specs/` and `*.md` excluded): 138 files, ~15,047 insertions / 3,109 deletions. Covers the Granite remediation commits (`e9e259be5`, `205736167`), the Blueprint spec-remediation commit (`58b611ece`, landed mid-review), and the uncommitted perf-measurement-contract work in `Tools/`.
- **Dedup**: all 64 findings in [ThreeDayDiff-2026-07-05-Findings.md](ThreeDayDiff-2026-07-05-Findings.md) were excluded up front; findings below are new, or document an incomplete/wrong remediation of a prior finding.
- **Method**: 10 independent finder angles (5 correctness, 3 cleanup, altitude, conventions) → 53 candidates → 37 after dedup → 37 adversarial verifiers (1-vote, 3-state) → gap sweep (+3 candidates, verified) → **35 survived** (32 CONFIRMED, 3 PLAUSIBLE), 5 refuted.
- Conventions angle (CLAUDE.md/AGENTS.md) found zero violations — all hard rules (catch(...) ban, RAII cleanup, PostMessage payload rules, .rc placeholder rules) hold across the diff.

## Top 15 (ranked)

### 1. [HIGH][bug] MTP picker worker races the unsynchronized FileSystemPluginManager singleton — UAF / execution in unloaded DLL
`RedSalamander/ConnectionManagerWindow.cpp:938` (flagged independently by 5 finder angles)
`MtpPickerWorkerCallback` (queued fire-and-forget via `TrySubmitThreadpoolCallback`, no join possible) calls `FileSystemPluginManager::GetInstance().EnumerateConnectionBrowseDevices/Storages` on a threadpool thread. `FileSystemPluginManager` has **no mutex anywhere**; `FindPluginById` returns a raw `PluginEntry*` into `_plugins`, and `CallConnectionBrowseExport` uses the module handle unpinned. UI-thread `Refresh`/`Discover` (settings hot-reload, Manage Plugins apply — `RedSalamander.cpp:8394`) does `UnloadAll(FreeLibrary); _plugins.clear()` with no quiesce of picker workers, and the deferred-unload guard doesn't help: `RedSalamanderBrowseConnectionTargets` increments no live counter, so `CanUnloadNow` returns TRUE mid-browse. WPD enumeration takes seconds, making the window wide. The replaced picker code was self-contained in-process WPD with no shared state — this hazard is new.

### 2. [HIGH][bug] MtpBackendCancelRequest's module keep-alive FreeLibrarys the DLL from inside itself before the backend is released
`Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:1717`
Declaration order (`:1714-1715`) is `backend; moduleKeepAlive;` — reverse destruction runs `~moduleKeepAlive` (plain `FreeLibrary`, executing *inside* FileSystemMtp.dll) **before** `~shared_ptr<IMtpBackend>`, and the return from `FreeLibrary` lands in unmapped memory when the keep-alive held the last reference. The generic threadpool lambda (`Common/Helpers.h:1915`) discards `PTP_CALLBACK_INSTANCE`, so `FreeLibraryWhenCallbackReturns` cannot currently be used — which is the required fix (member reorder alone doesn't help; the lambda epilogue is also module code). Reachable: watchdog trip (`:2244`) or Disconnect (`:3604`) queues a cancel; `CanUnloadFileSystemMtpModule` only checks quarantined commands, so the host FreeLibrarys the DLL while the request is pending.

### 3. [MEDIUM][bug] One bad color key now silently and permanently deletes a user's entire inline theme
`Common/Common/SettingsStore.cpp:1191`
`ParseTheme` routes inline settings.json themes through the strict theme-file parser `ParseThemeDefinitionJson5`; one invalid color value/key (`ThemeDefinitionIo.cpp:326-329`), >64-char name, non-builtin `baseThemeId`, or invalid id drops the **whole** `ThemeDefinition` with only a `Debug::Warning`. The old inline parser skipped just the bad entry. Since the next settings save serializes the in-memory list, the theme is then erased from disk — silent, permanent data loss on upgrade/downgrade or a single typo.

### 4. [MEDIUM][bug] ViewerPluginManager::Shutdown FreeLibrarys exactly the deferred-busy modules the deferral machinery exists to protect
`RedSalamander/ViewerPluginManager.cpp:200`
`Shutdown` runs `UnloadAll(ProcessShutdown); _plugins.clear(); _deferredUnloadEntries.clear();` — the deferred entries never get the ProcessShutdown `release()` (leave-mapped) path that `FileSystemPluginManager::Shutdown` applies via `SweepDeferredUnloadEntries(ProcessShutdown)`. `wil::unique_hmodule` destructors FreeLibrary modules that deferred because `RedSalamanderPluginCanUnloadNow` returned FALSE → their live threads execute unmapped code at exit. Latent until a viewer exports the hook, but the machinery is self-defeating as shipped.

### 5. [MEDIUM][bug] UIA action dispatch never signals completion when the payload is drained — every in-flight action stalls assistive tech 5s
`Common/DxUi/DxUi.Accessibility.cpp:4269`
The new posted-payload dispatch waits on `completedEvent` (5000ms). The only `SetEvent` is in the window-thread handler's scope_exit, reached only if the message is actually dispatched. `WM_NCDESTROY`'s `DrainPostedPayloadsForWindow` (`DxUi.WindowHost.cpp:2905`) deletes the `AccessibilityUiActionDispatch` without signaling → a screen-reader action issued during window close blocks the full 5s per action (old `SendMessageTimeoutW` failed fast on a dead window). Fix: signal `completedEvent` from the dispatch destructor.

### 6. [MEDIUM][bug] Timed-out UIA actions still execute later — caller told "failed", mutation happens anyway (double-execution on retry)
`Common/DxUi/DxUi.Accessibility.cpp:7274`
After `DispatchAccessibilityUiActionToWindowThread` returns `ERROR_TIMEOUT`, the posted dispatch stays queued and `TryHandleWindowHostAccessibilityMessage` executes the mutating action (`ExecuteUiThreadAction`/`ExecuteSelectOnWindowThread`) when the UI thread drains. No abandoned/timed-out flag exists in the dispatch struct. A UIA client that retries after the timeout mutates selection/caret twice.

### 7. [MEDIUM][bug] SettingsHotReload::Start no longer guarantees the watcher is armed — early writes are silently missed; readyEvent is now dead machinery
`RedSalamander/SettingsHotReload.cpp:561` (gap-sweep find)
The remediation of the prior 2s-blocking-Start finding deleted the readiness wait entirely instead of making it async-robust: `Start` returns `S_OK` immediately after spawning the watcher thread. Any settings.json write landing before `FindFirstChangeNotificationW` arms (thread-startup latency, or the 1000ms retry backoff after a transient arm failure) is missed with no post-arm catch-up stamp check. The `readyEvent` plumbing (`:521` created, `:140` signaled) is write-only dead code. The new selftest only passes because it sleeps `Scale(1200ms)` before the external save — timing-based evidence of exactly this window.

### 8. [MEDIUM][bug] Overwrite-journal absent-cache remediation: cross-instance race still live, fragile per-call-site opt-out, and per-item disk probes on bulk ops
`Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:2190`
The prior-finding-12 fix threads `bool cacheOverwriteJournalAbsence` through `RunBackendCommand→Submit→QueuedCommand→OverwriteJournalContext` (default `true`; six mutating call sites must each remember `false`). But `MarkOverwriteJournalAbsent` (`:1352`) is still check-then-act with no synchronization against `RecordOverwriteJournalIntent`'s invalidate+write (`:1075-1112`) — a reader can mark the identity absent *after* another instance recorded a live journal, reproducing the original bug. Side effect: every mutating command now pays journal-path resolve + `EnsureDirectoryExists` + failed-open probe per item (1000-item delete = 1000 probes). Right altitude: per-identity lock or generation counter inside the journal layer; derive cache policy from whether the queued command mutates.

### 9. [MEDIUM][bug] Browse can publish an empty devicePuid — the pnpId fallback is wiped by the callee and the HRESULT ignored
`Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:2316`
`ReadConnectionBrowseDevicePersistentId` does `devicePuid.clear()` first; every failure path returns with it empty; the caller sets `devicePuid = descriptor.pnpId` *before* the call and discards the HRESULT. A busy/locked phone whose WPD open or PUID read fails → `devicePuid:""` in browse JSON, persisted into the connection profile; downstream journal identity silently falls back from PUID to host.

### 10. [MEDIUM][bug] pendingBitmapCreates leaks on stale-generation thumbnail messages — phantom in-flight work forever
`RedSalamander/FolderView.Icons.cpp:1515`
The worker increments `pendingBitmapCreates` at post time guarded only by batchId (`:1315`, `:1330`; `countsPending` defaults true), but the UI-thread handler's stale-**generation** early-return (`:1515-1518`) fires before the scope_exit decrement is armed. Generation change without a batch bump → counter stays >0 for the rest of the batch; the thumbnail debug snapshot and any drain-waiting diagnostic/test reports phantom pending work.

### 11. [MEDIUM][bug] input_to_paint metric survives failed frames — corrupted samples feed the p95 perf gates the same delta is hardening
`RedSalamander/FolderView.Rendering.cpp:1932` (gap-sweep find)
The failed-render fix clears `_pendingRefreshToPaintMetric` on EndDraw/Present failure and the no-render-target return (`:1930/:1972/:2012/:979`) but leaves the sibling `_pendingInputToPaintMetric` armed on those same paths. After device-loss recovery, the *next* successful present emits `folder.frame.input_to_paint_us` spanning the failed frame + recovery + idle gap — feeding `perf_metrics.jsonl` / FolderViewPerfBudgets gates with inflated samples. The delta fixed exactly this defect class for the refresh metric only.

### 12. [MEDIUM][test-bug] MTP picker selftest hard-fails on release plugin builds — no ERROR_PROC_NOT_FOUND skip guard (prior finding #15's class, re-introduced)
`RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp:2136`
`RedSalamanderMtpSetPickerFakeBackendForSelfTest` exists only under `#ifdef _DEBUG` (`Factory.cpp:327`); the helper returns `ERROR_PROC_NOT_FOUND` (`:55-57`) and the case does `state.Require(SUCCEEDED(...))` instead of the Skip pattern used at `CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp:393`. The replaced host-side ENABLE_TESTS fixture worked with any plugin flavor.

### 13. [MEDIUM][test-bug] Impersonation-failure selftest passes a debug-only CLI flag to the search service — hard-fails every Release+ENABLE_TESTS run
`RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp:10012`
The test unconditionally passes `--test-fail-client-auth-impersonation-once`, parsed only under `#if defined(_DEBUG) || !defined(NDEBUG) || __SANITIZE_ADDRESS__` (`Main.cpp:3179`). A release service exits with unknown-option → `service.Start` Require fails; the case has no release skip.

### 14. [PLAUSIBLE-MEDIUM][bug] Context-menu debug probe is now an unbounded cross-thread SendMessageW — a wedged menu thread hangs the whole suite
`Common/DxUi/DxUi.Menu.cpp:5293`
The GR-17 fix for the timed-out form's dangling stack-output write replaced `SendMessageTimeoutW(1000ms, SMTO_ABORTIFHUNG|SMTO_BLOCK)` with plain `SendMessageW` (pinned by a source-scraping test). Probing a hung thread now blocks until the outer job timeout kills the process, masking the failure being diagnosed. The robust fix is heap-owned request storage + event timeout, as this same delta did for the UIA action path — not the raw timeout, and not unbounded blocking.

### 15. [MEDIUM][efficiency] Every search keystroke opens + schema-validates a fresh SQLite connection (up to 2× per query) just to read store_generation
`Common/LocalSearchIndexCore.cpp:4385`
`GetValidatedCachedPersistentStoreInfoForQuery` runs unconditionally on the steady-state path (no TTL, no mtime pre-check) and calls `SqliteIndexStore::ReadStoreGeneration` → fresh `OpenReadOnlyReadyConnection` + `EnsureReadableSchema` (~11 sqlite_master/PRAGMA probes) + meta SELECT + close (`SqliteIndexStore.cpp:2353`), executed at `:4694` and `:4771` within one `Enumerate` and once per `EnumerateNoWait`. Cheaper: cached read-only probe connection, mtime/size pre-check, or short TTL.

## Below the cut — 20 more confirmed findings

### Bugs (LOW)
16. **[LOW][bug]** `RedSalamander/FolderView.Icons.cpp:1088` — thumbnail abandoned-handshake check-then-act hole (test-only ForceProviderAllowedProbe path): worker SetEvent-then-load(abandoned), waiter timeout-then-store(abandoned), no final recheck → a completed thumbnail is consumed by neither side; the late-upgrade machinery never fires for that interleaving.
17. **[PLAUSIBLE-LOW][bug]** `Plugins/FileSystemMtp/Factory.cpp:273` — `RedSalamanderBrowseConnectionTargets` writes `result->jsonUtf8 = nullptr` (offset 8) *before* validating `result->sizeBytes`, defeating the size-versioning handshake for any smaller-struct caller.
18. **[LOW][bug]** `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:2541` — `CreateSelfTestWpdMtpBackend` discards `optionsJsonUtf8` (`static_cast<void>`) that the export `RedSalamanderMtpCreateWpdCacheForSelfTest` accepts and forwards; fixture options (delays, fault injection) silently ignored.
19. **[PLAUSIBLE-LOW][bug]** `Tools/Show-PerfRuns.ps1:470` — `Get-FolderViewBudgetMinimumSamples` flattens `machines[]` with no machineHash filter; native reader (`FindFolderViewPerfBudgetMachine`) filters by machine. Latent divergence; bites when a multi-machine budget file declares machine-specific minimums.

### Perf / efficiency
20. **[MEDIUM][efficiency]** `RedSalamander/IconCache.cpp:1241` — the 5s-failure-cache remediation removed negative caching entirely for live path lookups (`path_live_lookup_failed_uncached`); every re-enumeration re-issues blocking `SHGetFileInfoW` per persistently-failing path (dead links / offline shares, seconds each). Wanted: short-TTL or bounded-retry negative cache, not none.
21. **[MEDIUM][efficiency]** `Common/SearchServiceBroker.cpp:2342` — transient authorization failures (ERROR_BAD_NETPATH, ERROR_SEM_TIMEOUT…) are never cached in the per-query `clientDirectoryAccessCache`; thousands of candidates under one dead parent repeat impersonate+CreateFileW+log serially. Cache lifetime is already one batch — caching the transient verdict loses nothing.
22. **[LOW][efficiency/simplify]** `Common/DxUi/DxUi.Accessibility.cpp:1263` — `gridAccessibleColumns` is a materialized identity vector `[0..gridColumnCount)`; every use reduces to arithmetic (linear `FindSizeValueIndex` scans at `:1516/:1807/:1855`). NOTE: the all-columns *extraction* it drives is the deliberate, regression-tested GR-9 Narrator fix — do NOT revert to visible-only columns; only the vector→arithmetic simplification is safely actionable.
23. **[LOW][efficiency]** `Common/Common/SettingsStore.cpp:1183` — ParseTheme round-trips each inline theme DOM→JSON-text→re-parse (3 traversals, 2 intermediates) at startup and every hot-reload; factor the field extraction to accept the parsed value. (Same mechanism enables fixing #3 properly.)

### Test quality / altitude
24. **[MEDIUM][test-quality]** `Tools/Tests/TestHarnessSourceContracts.Tests.ps1:216` — ~440 lines of new source-text-scraping tests (18 It blocks + 3 DxUiTests source-scrape twins + mtp_identity_helpers_are_shared) pin helper names, function order, struct shapes, and exact `#pragma` comment wording (`:548-549` enshrines the "unreferenced" vs "unused" wording divergence forever) — prior finding #36's fragility ×20. Keep the behavioral tests added alongside; delete the substring guards.
25. **[MEDIUM][altitude]** `Common/SearchServiceBroker.cpp:2311` — test-only impersonation fault injection compiled inline into the production authorization path `CheckClientCanListDirectory`, gated by ad-hoc `_DEBUG || !NDEBUG || ASAN` (copy-pasted in 5 places) instead of the ENABLE_TESTS seam; ships a deliberate-auth-failure CLI switch in any non-NDEBUG build config.
26. **[MEDIUM][altitude]** `RedSalamander/FolderView.Rendering.cpp:2301` — `DrawItem` carries an ENABLE_TESTS branch force-creating throwaway D2D brushes so a selftest can watch `drawItemTransientBrushCreateCount`; that forced path is the counter's **only** increment site, so the test is self-fulfilling and can never catch a real per-item brush regression.

### Reuse / simplification
27. **[MEDIUM][reuse]** `RedSalamander/ViewerPluginManager.cpp:1118` — ~150-line copy of FileSystemPluginManager's deferred-unload machinery (Sweep/IsDeferred/AddPlaceholder token-identical) already diverging: FS manager re-sweeps on demand at lookups (`:788-791`), viewer manager only during Discover (`:629`) → deferred viewer DLL stays ERROR_BUSY until rediscovery. Extract one shared PluginModuleLifecycle helper (also the right place to fix #4).
28. **[MEDIUM][reuse]** `RedSalamander/FileSystemPluginManager.cpp:87` — new private JsonValue accessor suite; repo now has 6 hand-rolled member-walkers over the same connection/extra JSON shapes (Connections selftest `:562`, ConnectionProfileUtils, ConnectionManagerWindow `:444`, HostServices `:1399`) because Common::Settings exposes no member-lookup API next to ParseJsonValue. Hoist one set.
29. **[MEDIUM][reuse]** `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp:3874` — fourth repo-root walker in the selftest binary, each with different markers/depth/env handling; they can resolve different roots in worktree layouts. Export `SelfTest::TryFindRepoRoot` from SelfTestCommon.
30. **[LOW][simplify]** `Common/SearchServiceBroker.cpp:2011` — ACCESS_DENIED_SKIPPED tracked via three mechanisms; the direct `state.stats->warningFlags` write is dead on the completion path (`*outStats = stats` wholesale at `LocalSearchIndexCore.cpp:4719/4812/5143` precedes the `:2838` merge) but still feeds mid-run QUERY_BATCH_SENT events — consolidate by OR-ing `state.warningFlags` at the flush site.
31. **[LOW][simplify]** `RedSalamander/ViewerPluginManager.cpp:1127` — `if (entry.module)` immediately after the `if (!entry.module) return true;` early-out is always true; delete and unindent.
32. **[LOW][reuse]** `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp:19492` — `IsViewerSpaceEnvFlagEnabled` duplicates the just-hoisted `EnvironmentVariables::IsTruthyFlagSet` with a divergent accept-set (`=y` works for FolderView flags, not ViewerSpace flags), and a contract test pins the duplicate.
33. **[LOW][reuse]** `Common/DxUi/DxUi.Accessibility.cpp:814` — `ResolveScrollPanelViewportRect` re-implements `ScrollPanel::GetViewportRect()` (currently private) bit-identically; make it public and call it.
34. **[LOW][reuse]** `RedSalamander/FileSystemPluginManager.cpp:64` — ~11th file-local `Utf16FromUtf8`; the delta already touched Common/Helpers.h twice — hoist one converter.
35. **[LOW][reuse]** `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp:33` — third selftest copy of the find-entry + LoadLibraryExW + pragma-4191 GetProcAddress export loader, with a weaker guard set than `CallConnectionBrowseExport`'s (no unloadDeferred/ERROR_BUSY handling).

## Refuted (appendix)
- **S3 write:false capability flip** (`Plugins/FileSystemS3/FileSystemS3.h:343`) — the flipped block is `kCapabilitiesJsonS3Table` (read-only S3 Table sub-plugin), intentional Blueprint spec alignment; main `kCapabilitiesJsonS3` still has `write:true` + wildcard import (`:309/:318`).
- **MTP picker storage disambiguation dropped** (`FileSystemMtp.Device.cpp:2439`) — not dropped; `EnumerateObjectItems` applies `DisambiguateDuplicateNames`/`MtpDuplicateObjectSuffix` internally on both exit paths (`:845/:863`), so picker initialPath matches runtime paths.
- **E_INVALIDARG search fallback masking** (`FileSystem.Search.cpp:1962`) — re-report of a finding already refuted in the 07-05 review (GR-13: intentional, documented, RED/GREEN-tested; fallback logs a Warning with hr).
- **No-op destination bridge decorator** (`FolderWindow.FileOperations.State.cpp:6479`) — deliberate contract-pinned scaffolding for the open FIR-1/Floodgate FG-A1 destination fault-injection slice; deleting it would break a contract test and conflict with two tracked WIP plans.
- **QueueBackendCancel double-cancel** (`FileSystemMtp.Core.cpp:1762`) — fallback resets `request->backend` after the explicit cancel, so the destructor's guarded cancel is a no-op.

## Cross-cutting observations
- The remediation pattern "replace a timeout/cache/sync mechanism by deleting it" appears three times (#7 SettingsHotReload readiness, #14 menu-probe timeout, #20 icon negative cache) — remediations of prior findings that removed the safety margin instead of re-establishing it robustly.
- The plugin-manager work (V1/V2/#4/#27) needs one shared, thread-aware plugin-module-lifecycle component; four of the top findings are facets of that missing abstraction.
- Source-scraping contract tests grew ~20× in this window (#24) while the flake-stabilization plan is trying to reduce exactly this class of fragility.
