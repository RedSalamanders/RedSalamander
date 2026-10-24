# Operation Firebreak - Consolidated Review Remediation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace four overlapping review ledgers with one single-owner, ready-to-implement
remediation plan that fixes the live correctness, crash, data-safety, performance, and test-quality
findings while preserving the original ledgers in `Specs/Plans/Done/` as references.

**Architecture:** Firebreak is the executable owner; the archived source ledgers are evidence, not
parallel work queues. Work is sequenced by blast radius: red gates and crash classes first, then MTP
cache/lifecycle, search/FileOps/data-loss issues, UI/accessibility, perf-sensitive hot paths, and
finally simplification cleanup. Items routed to other live plans remain routed and are listed here
only to prevent duplicate implementation.

**Tech Stack:** Windows C++ (`stdcpplatest`), WIL RAII, Direct2D/DirectWrite/DxUi, vcpkg, MSBuild via
`.\build.ps1`, selftests via `.\Tools\Run-AllTests.ps1`, Pester for tool tests, archived performance
evidence under `Specs/TestRuns/`.

---

## Source Ledgers

The four ledgers below have been challenged, consolidated, and archived as references:

- `Specs/Plans/Done/Operation_Bedrock_ThreeDayDiffReviewRemediation_2026-07-06.md`
- `Specs/Plans/Done/Operation_Evergreen_ThreeDayDiffReviewRemediation_2026-07-05.md`
- `Specs/Plans/Done/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md`
- `Specs/Plans/Done/Operation_Keystone_ThreeDayReviewRemediation_2026-06-20.md`

Evidence documents remain authoritative for finding provenance:

- `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md` for Bedrock F#1..F#35.
- `Specs/Reviews/ThreeDayDiff-2026-07-05-Findings.md` for Evergreen F#1..F#64.
- Clearwater and Keystone findings are embedded in their archived ledgers.

Firebreak owns all live work from those ledgers except the explicit routed/blocked items below.
Do not re-open the archived ledgers as WIP queues.

## Challenge Decisions

These are the consolidation decisions made while reviewing the four inputs.

- **MTP owns one unified track.** Evergreen Track A and Bedrock Tracks A/D touch the same MTP worker,
  WPD cache, journal, fixture, and unload-lifetime files. Execute them as one MTP track in this plan.
  Do not split cache invalidation, COM lifetime, backend recreation, browse worker pinning, and journal
  absent-cache fixes across independent sessions.
- **Search junction aliases use the coverage-only contract first.** Clearwater CW-6 is not a one-line
  gate fix. Firebreak chooses the bounded implementation: recurse directory aliases with a visited set
  so alias-only descendants are indexed; duplicate canonical+alias result paths are out of scope until
  a separate alias-path model is designed.
- **Noexcept layout-cache work is a policy clarification, not a code churn item.** Clearwater CW-S2
  conflicts with the repo rule that `std::bad_alloc` is fatal. Do not mechanically remove `noexcept`.
  Add a short policy comment or mark WONTFIX in the closeout unless a product owner explicitly requests
  recoverable OOM UI layout.
- **Destructive repo cleanup is blocked.** Evergreen EV-G1 deletes branches/snapshot refs and runs git
  GC. It remains `[blocked-human]` until the human explicitly approves the exact ref deletions.
- **UIA coalescing remains routed.** Evergreen EV-B2/EV-B3 and Bedrock's all-columns extraction-cost
  half are owned by `Specs/Plans/WIP/DxUi_Uia_ContinuationBaton_2026-06-29.md`. Firebreak owns only
  bounded dispatch correctness, offscreen selected-row cost, TSF lifetime, and mechanical UIA cleanup.
- **Perf Measurement Contract remains routed.** Bedrock F#19 (`Show-PerfRuns` machine filtering /
  `minimumSamples`) stays with `Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md`.
  Firebreak still owns Evergreen EV-F2/EV-F3/EV-F4 tool-quality fixes that are not schema-contract
  decisions.
- **Done means done.** Clearwater CW-7 and CW-10 are not implementation items here. Keep their test and
  spec evidence intact; do not duplicate them.
- **Source-scraping tests are a regression trap.** Firebreak bans new source-text assertions for product
  behavior. Existing source-scraping tests added by the reviewed windows must be replaced by behavioral
  coverage or deleted when they pin style only.

## Global Execution Rules

- [x] Before touching a task's code, run `git status --short` and inspect any dirty files in the task
  area. Work with user changes; do not revert unrelated changes.
- [x] For every file anchor copied from a source ledger, re-locate by symbol/text before editing. The
  ledgers were anchored across several moving commits.
- [x] Build with `.\build.ps1 -Configuration Debug -Platform x64` unless a narrower project build is
  called out. Use Release builds when validating release-skip behavior.
- [x] Use RED-first tests where the task specifies a failing proof. If the RED step cannot be made
  deterministic, write the closest deterministic counter/assertion first and document why an exact race
  proof is not feasible.
- [x] Perf-sensitive work must define the scenario, metric key, deterministic selftest, and archived
  baseline/candidate runs under `Specs/TestRuns/` before claiming the item complete.
- [x] Closeout for a completed task must include targeted tests, a green broader gate when practical,
  and authoritative spec updates for durable behavior.

Expected green gates:

```powershell
.\build.ps1 -Configuration Debug -Platform x64
.\Tools\Run-AllTests.ps1 -SkipBuild
.\Tools\Run-AllTests.ps1 -Suite Full
```

Expected result for the full suite: build succeeds, all deterministic selftests pass or report documented
SKIP for configuration-gated debug exports, and any perf-sensitive task archives a new run under
`Specs/TestRuns/`.

## Task 0: Red Gates, Release-Skip Parity, and Drift Control

**Files:**

- Modify: `Tools/Tests/TestInventory.Tests.ps1`
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`
- Modify: `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp`
- Modify: `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- Modify: `RedSalamanderSearchService/Main.cpp`
- Modify: `Common/SearchServiceBroker.h`
- Modify: `Common/SearchServiceBroker.cpp`

- [x] **Step 0.1: Recompute the current test inventory count [EV-F1].**

  Use the existing inventory helper instead of trusting snapshot counts from Evergreen. Update the Pester
  expectation and coverage doc to the count produced by the current tree. Do not import the orphan
  snapshot `31a0b6b83`; it is reference evidence only.

  Run:

  ```powershell
  .\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
  pwsh -NoProfile -File .\Tools\Tests\TestInventory.Tests.ps1
  ```

  Expected: the exact-count and doc-vs-source checks pass against the current checkout.

- [x] **Step 0.2: Add release-skip guards for debug-only MTP exports [BR-G1, EV-TA1].**

  In `Commands.SelfTest.Connections.cpp`, if the fake browse export returns
  `HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND)`, call `state.Skip(...)` before `state.Require`.
  Mirror the established MTP skip pattern already used by sibling MTP selftests.

  In `CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp`, guard
  `mtp_wpd_session_and_path_cache_reuse` the same way when the WPD-cache fixture factory export is absent.

  Run:

  ```powershell
  .\build.ps1 -Configuration Release -Platform x64
  .\Tools\Run-AllTests.ps1 -SkipBuild
  ```

  Expected: release-flavor MTP debug-export cases report SKIP instead of FAIL.

- [x] **Step 0.3: Consolidate search test-hook gates [BR-F3, BR-G2].**

  Replace the ad-hoc `#if _DEBUG || !NDEBUG || ASAN` copies in `SearchServiceBroker.h`,
  `SearchServiceBroker.cpp`, and `RedSalamanderSearchService/Main.cpp` with one named build gate
  such as `RS_SEARCH_TEST_HOOKS`, wired to the existing test-enabled project configuration.

  Move the impersonation failure injection out of `CheckClientCanListDirectory` into a small
  enum-dispatched test hook owned by `ServerOptions`. Keep the CLI flag available only under
  `RS_SEARCH_TEST_HOOKS`.

  Guard the selftest that passes `--test-fail-client-auth-impersonation-once` with the same macro and
  call `state.Skip(...)` when hooks are not compiled.

  Run:

  ```powershell
  .\build.ps1 -Configuration Debug -Platform x64
  .\Tools\Run-AllTests.ps1 -SkipBuild
  .\build.ps1 -Configuration Release -Platform x64
  rg --fixed-strings "--test-fail-client-auth-impersonation-once" ".build\x64\Release"
  ```

  Expected: debug selftest still passes; release selftest skips; release output search finds no test flag
  string.

## Task 1: P0 Plugin Module Lifecycle and MTP Browse Worker Safety

### Task 0 closeout evidence — 2026-07-13

- RED inventory proof found 126 FileOps phases versus 121 expected and 270 Pester cases versus 269 expected; the exact-count tests and both coverage inventories now agree and `TestInventory.Tests.ps1` passes 5/5.
- Full Debug x64 rebuild passed with 0 warnings/errors: `.build/logs/msbuild-20260713_123208_618.log`.
- Debug focused proof passed: MTP picker 1/1 and MTP WPD-cache plus search impersonation-hook cases 2/2.
- Test-enabled Release build passed with 0 warnings/errors: `.build/logs/msbuild-20260713_123519_948.log`.
- Release focused proof intentionally skipped MTP picker 1/1 and Compare MTP/search hooks 2/2; the Release SearchService binary contains no `--test-fail-client-auth-impersonation-once` UTF-16 string.

**Files:**

- Modify: `Common/Helpers.h`
- Modify: `RedSalamander/ConnectionManagerWindow.cpp`
- Modify: `RedSalamander/FileSystemPluginManager.h`
- Modify: `RedSalamander/FileSystemPluginManager.cpp`
- Modify: `RedSalamander/ViewerPluginManager.cpp`
- Create: `RedSalamander/PluginModuleLifecycle.h`
- Create: `RedSalamander/PluginModuleLifecycle.cpp`
- Modify: `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp`
- Modify: `Plugins/FileSystemMtp/FileSystemMtp.h`
- Modify tests under `RedSalamander/SelfTest/Commands/` and plugin lifecycle selftests

- [x] **Step 1.1: Fix cancel-request keep-alive unload ordering [BR-A2].**

  Add a threadpool helper variant that passes `PTP_CALLBACK_INSTANCE` into the callback. In the MTP
  cancel request callback, execute `RequestCancel`, reset the backend, and make the module pin the last
  action via `FreeLibraryWhenCallbackReturns(instance, request->moduleKeepAlive.release())`.

  On submit failure from inside FileSystemMtp code, release the module pin deliberately and log once
  instead of calling `FreeLibrary` on the failing path. Add a pending-cancel counter consulted by
  `CanUnloadFileSystemMtpModule`.

  Verify with a fake-backend cancel selftest plus ASan/AppVerifier when available. Expected: unload is
  refused while cancel is pending, and no return occurs into an unmapped DLL.

- [x] **Step 1.2: Keep `FileSystemPluginManager` UI-thread-only and pin MTP browse exports [BR-A1].**

  Add debug UI-thread assertions to public `FileSystemPluginManager` entry points. On the UI thread,
  resolve the MTP browse plugin, copy `pluginId` and path into the request, load a worker-owned
  `wil::unique_hmodule` with `LoadLibraryExW`, and resolve `RedSalamanderBrowseConnectionTargets`
  before queuing work. The worker must not touch `FileSystemPluginManager`.

  Convert the picker refresh to owned threadpool work and wait/cancel it during `ConnectionManagerWindow`
  teardown. Keep heap-owned posted results for normal completion, but teardown must stop producers before
  the window drains payloads.

  RED proof: delay the fake browse backend, refresh plugins on the UI thread while the worker is in
  flight, and assert clean failure or coherent picker results without UAF.

- [x] **Step 1.3: Unify plugin module deferred-unload lifecycle [BR-A3, BR-A4].**

  Extract shared defer/sweep/placeholder logic from `FileSystemPluginManager` and `ViewerPluginManager`
  into a single lifecycle helper. The helper owns `RedSalamanderPluginCanUnloadNow` probing,
  `ModuleUnloadMode`, deferred placeholders, on-demand sweep, and process-shutdown release semantics.

  Port the FileSystem manager's on-demand sweep behavior to the Viewer manager. In Viewer shutdown,
  sweep deferred entries with `ProcessShutdown` before clearing the vector so busy modules are released
  rather than `FreeLibrary`-destroyed.

  Delete the dead always-true `if (entry.module)` wrapper in viewer unload while touching the code.

  Verify with fake viewer and file-system plugin cases: a busy module remains mapped during shutdown,
  and a healed deferred viewer DLL becomes loadable without full rediscovery.

## Task 2: Unified MTP Cache, COM, Backend, and Journal Reliability

**Files:**

- Modify: `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp`
- Modify: `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp`
- Modify: `Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp`
- Modify: `Plugins/FileSystemMtp/FileSystemMtp.Internal.h`
- Modify: `Plugins/FileSystemMtp/FileSystemMtp.h`
- Modify: `Plugins/FileSystemMtp/Factory.cpp`
- Modify MTP selftests under `RedSalamander/SelfTest/CompareDirectories/`

- [x] **Step 2.1: Fix WPD cache invalidation and metadata freshness [EV-A1, EV-A2].**

  Route every operation-phase failure through `FailAndMaybeInvalidateCaches`. On non-whitelisted HRESULTs,
  evict the session entry and the PnP subtree. When a missing-object error comes from a cache-resolved
  object id, evict that entry plus descendants and retry once with a fresh resolve.

  Refresh `_pathCache` child entries on every successful `EnumerateDirectory`. For size-sensitive reads,
  re-fetch object properties before transfer or drop the cache entry before use. On stream size mismatch,
  invalidate the file entry and retry with fresh metadata.

  RED proofs: session-death fixture reopens a fresh session on retry; cached file whose size changes reads
  successfully with the new size instead of returning permanent `ERROR_CRC`.

- [x] **Step 2.2: Initialize COM once for the backend worker lifetime [EV-A3].**

  Move MTA initialization to `MtpBackendCommandQueue::WorkerMain` with
  `wil::CoInitializeEx(COINIT_MULTITHREADED)`. Remove per-command teardown that can invalidate cached WPD
  interfaces; keep per-call guards only as debug assertions that the worker already owns MTA lifetime.

  Add debug assertions around WPD calls and reader stream calls so tests prove COM is initialized.

- [x] **Step 2.3: Recreate backend workers and generation-bind readers [EV-A4, EV-A5, EV-A8].**

  In `Initialize`, recreate `_backendWorker` when it is null and `_backend` is set. Make queue creation
  non-terminating under thread creation failure: report `E_OUTOFMEMORY` or `ERROR_NOT_ENOUGH_MEMORY`
  through `RunBackendCommand` instead of starting a `std::jthread` inside a `noexcept` constructor.

  Bind `MtpBackendReader` to the backend generation it came from. If the owner's current backend
  generation differs, return `ERROR_DEVICE_NOT_CONNECTED` instead of submitting old stream work to a new
  queue.

  RED proof: watchdog trip, `Initialize`, and next command succeeds; stale pre-trip reader fails safely.

- [x] **Step 2.4: Repair overwrite-journal absent-cache correctness and cost [EV-A6, BR-D2].**

  Move the absent-cache race fix into the shared journal layer. Add a per-device-identity generation
  counter. `RecordOverwriteJournalIntent` increments before invalidation/write. `MarkOverwriteJournalAbsent`
  stores absent only when the observed generation is unchanged.

  Delete the positional `cacheOverwriteJournalAbsence` plumbing through `RunBackendCommand`, `Submit`, and
  `QueuedCommand`. Derive mutating behavior from the command descriptor, and allow safe absent-cache use
  for mutating commands once generation validation exists.

  Normalize journal identity keys with the same case folding as `StableDeviceHash`.

  RED proof: instance A observes absent, instance B records a journal, A's stale mark is rejected, and
  replay sees B's journal. Perf proof: N fake-backend deletes perform O(1) journal probes, not O(N).

- [x] **Step 2.5: Finish MTP correctness edge cases [EV-A7, EV-A9, EV-A10, EV-A11, EV-A12, BR-D1, BR-D3, BR-D4].**

  Apply these in one MTP pass after the cache/worker shape is stable:

  - `devicePuid` fallback: do not clear caller state before WPD PUID read; after failure or empty success,
    set `devicePuid = descriptor.pnpId`.
  - Rename/copy/move overwrite guard: fail fast for directory sources overwriting files; centralize in
    the shared overwrite orchestration helper.
  - Selftest WPD-cache fixture: extract the fixture behind a device operations seam, remove production
    hot-path fixture branches, and report `liveWpd=false` for the fixture.
  - Fixture options: parse supported fake-WPD options such as `readFileDelayMs`; reject unsupported
    non-empty options with `E_INVALIDARG`.
  - `_backendThreadIdsOverflow`: read and write under the same thread-stats lock.
  - `Factory.cpp` exports: validate `result` and `result->sizeBytes` before writing fields such as
    `jsonUtf8`.
  - Simplify duplicated copy/move/rename overwrite orchestration; delete `deviceIoMutex`; inline the
    one-line journal replay wrapper; merge the two memory readers into one `MemoryBackendFileReader`.

  Verify debug and release MTP selftests. Expected: debug fixture options are honored; release cases skip
  missing exports; no null derefs in fixture mutators.

### Tasks 1-2 closeout evidence — 2026-07-13

- MTP cancel requests now hold a module pin through backend cancellation and transfer the pin to
  `FreeLibraryWhenCallbackReturns` as the callback's final action; pending requests block unload. The
  Connection Manager prepares a module-pinned browse export on the UI thread, workers never touch
  `FileSystemPluginManager`, and window teardown cancels and waits owned threadpool work. The shared
  plugin lifecycle helper now supplies deferred placeholders, on-demand healing sweeps, and
  process-shutdown release semantics for both FileSystem and Viewer managers. Durable contracts are in
  `Specs/Core/Core_ConnectionManager.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, and
  `Specs/Plugins/Plugins_ViewerPlugins.md`.
- WPD failures now invalidate the affected session/path subtree and retry stale resolutions once;
  enumeration and size-sensitive reads refresh metadata. The backend queue owns MTA initialization for
  its worker lifetime, reports thread-creation failure without terminating, recreates after quarantine,
  and rejects readers from an older backend generation. The shared journal generation prevents stale
  absent-cache commits and permits constant-cost mutating-command probes. PUID fallback, fixture
  validation, statistics locking, export result validation, shared memory readers, and centralized
  file-only overwrite orchestration are recorded in `Specs/FileSystem/FileSystem_Mtp.md`.
- The final Debug x64 solution build passed with 0 errors in
  `.build/logs/msbuild-20260713_142145_838.log`; after removing the one new C5246 test warning, the
  current RedSalamander rebuild passed with 0 errors and only 40 existing WIL C4625/C4626 diagnostics in
  `.build/logs/msbuild-20260713_142404_374.log`. `PerformanceTests2` passed 14/14, including behavioral
  fake-Viewer proofs for deferred healing and busy process-shutdown release.
- Focused lifecycle/MTP proofs passed: pending-cancel unload
  (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_142703/`), quarantined runtime refresh
  (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_142709/`), WPD session reopen plus fresh
  size retry (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_142711/`), worker recreation plus
  stale-reader rejection (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_142713/`), and the
  copy/move/rename directory-over-file no-mutation guard
  (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_142713_001/`). The delayed MTP picker proof
  passed in `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_142728/`.
- Perf scenario: repeated mutating fake-backend commands for one normalized device identity after an
  absent-journal observation. The same Debug run records 16 baseline filesystem probes versus 1
  generation-validated candidate probe in
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_142711_001/perf/perf_metrics.jsonl`, proving
  O(1) absent-cache behavior while the stale-generation replay case remains green. Task 0's test-enabled
  Release proof confirms the debug-only MTP fixture exports skip cleanly and are absent from production
  behavior.

## Task 3: Search, Index, Broker, and Service Trust Boundary

**Files:**

- Modify: `Common/LocalSearchIndexCore.cpp`
- Modify: `Common/SqliteIndexStore.cpp`
- Modify: `Common/SearchServiceBroker.cpp`
- Modify: `Common/SearchServiceBroker.h`
- Modify: `Common/PlugInterfaces/FileSystem.h`
- Modify: `Plugins/FileSystem/FileSystem.Search.cpp`
- Modify: `RedSalamanderSearchService/Main.cpp`
- Modify: `RedSalamander/FindFilesWindow.cpp`
- Modify: `RedSalamander/Resource.h`, `RedSalamander/RedSalamander.rc`, and satellite `.rc` files
- Modify search/index selftests under `RedSalamander/SelfTest/CompareDirectories/`
- Modify: `Specs/Core/Core_Search.md`

- [x] **Step 3.1: Prevent service root rejections from poisoning global cooldown [EV-C5].**

  In `FileSystem.Search.cpp`, arm the 5s service cooldown only for transport/connect failures such as
  missing pipe or timeout. Do not arm it for healthy-service per-root validation rejections. Add a distinct
  warning flag for root rejections so callers can distinguish "service healthy but refused this root" from
  "service unavailable".

  RED proof: service rejects root A; immediate search on valid root B still uses the service backend.

- [x] **Step 3.2: Make missing-pipe connection retry cancellable [EV-C6].**

  Thread the caller's cancel check or stop event into the missing-pipe retry loop in
  `ConnectClientPipe`. Replace uncancellable `Sleep(10)` polling with a loop that exits promptly with
  `ERROR_CANCELLED` when cancellation is requested.

  Verify with an env override stretching retry to a long duration; cancellation returns promptly.

- [x] **Step 3.3: Clean trust-boundary dead code and warning accumulation [EV-C7, BR-F4].**

  Rewrite `AuthorizeClientRootRebuildAccess` as an intentional `for (;;)` walk-up and delete the
  unreachable trailing return.

  In `SearchServiceBroker`, remove the direct `state.stats->warningFlags` write for
  `ACCESS_DENIED_SKIPPED`. Use `state.warningFlags` as the single accumulator and merge it into mid-run
  batch headers and completion stats.

  Verify that warning flags appear in both `QUERY_BATCH_SENT` and completion results.

- [x] **Step 3.4: Cache transient authorization failures for a single batch [BR-F2].**

  For transient parent access failures inside one broker query, cache the negative verdict for the batch
  and demote logging to once per parent per batch. Keep durable negative caching semantics unchanged.

  RED proof: N candidates under one `ERROR_BAD_NETPATH` parent produce one impersonation/open attempt and
  one log event.

- [x] **Step 3.5: Reduce SQLite generation probe churn [BR-F1].**

  Capture database and WAL `(mtime,size)` after a successful generation validation. If both are unchanged,
  skip `ReadStoreGeneration`. Within a single `Enumerate`, probe at most once and pass that result through
  the second validation site. Add a cached read-only probe connection only if the stat and per-call
  reduction does not meet the selftest budget.

  Perf proof: repeated steady-state queries perform zero SQLite opens after warm-up; generation bump is
  detected within one query. Archive baseline and candidate `search.*` evidence under `Specs/TestRuns/`.

- [x] **Step 3.6: Implement the coverage-only junction alias contract [CW-6].**

  In `LocalSearchIndexCore`, when hydrating a directory subtree and a `NodeId` already exists because a
  junction/mountpoint follows the same physical directory, recurse alias-only descendants with a visited
  set instead of dropping the subtree. Do not emit duplicate canonical+alias result paths in this pass.

  Document the chosen contract in `Specs/Core/Core_Search.md`: alias-only descendants must be indexed;
  duplicate path emission requires a future alias-path representation.

  RED proof: a junction alias containing unique `needle_alias.txt` is indexed. Skip only when creating the
  junction fails with `ERROR_PRIVILEGE_NOT_HELD`.

- [x] **Step 3.7: Fix SQLite maintenance and skip-count observability [CW-12, CW-14].**

  Move `kMetaLastCheckpointUtc` update to after a successful final checkpoint. If checkpoint returns busy
  and `RunAutomaticMaintenance` returns `S_FALSE`, leave the timestamp unchanged.

  In `EnumerateVolume`, check `hr == LocalSearchIndexCore::kSkipCandidateHr` before incrementing
  `emittedRows`.

  RED proofs: busy checkpoint leaves timestamp empty; callback skip plus one emitted row reports
  `emittedRows == 1`.

### Task 3 closeout evidence — 2026-07-13

- Service fallback now distinguishes transport failures from request-scoped root rejection, arms the
  cooldown only for the former, carries `SERVICE_ROOT_REJECTED` through the fallback result, and renders a
  localized Find warning. Missing-pipe retries poll the caller cancellation callback. Broker query batches
  and completion share one warning accumulator, and transient parent-open failures are attempted/logged
  once within the query without poisoning later queries.
- SQLite query validation caches database/WAL stamps, normalizes only empty read-only WAL sidecar
  creation/retirement with an unchanged database, reuses one validated store result within `Enumerate`, and
  still refreshes an externally bumped generation in the next query. Junction hydration uses a physical-ID
  visited set to discover alias-only descendants once. Automatic checkpoint timestamps and enumeration
  emitted-row counts now describe only successfully completed/emitted work. Durable contracts and the six
  focused selftests are recorded in `Specs/Core/Core_Search.md`.
- Focused Task 3 Compare proofs passed: missing-pipe cancellation
  (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_133420/`), root rejection without cooldown
  (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_133428/`), transient-parent cache plus
  batch/completion warning parity (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_133437/`),
  junction alias hydration (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_135345/`),
  busy-checkpoint and callback-skip observability
  (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_135355/`), and generation-stamp validation
  (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_135810/`).
- Focused regressions passed for cross-query transient authorization
  (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_135823/`), candidate-impersonation warning
  behavior (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_135828/`), and external SQLite
  rotation (`Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_135834/`). Resource localization
  contracts passed 4/4. The clean current Debug x64
  RedSalamander build passed with 0 warnings/errors in `.build/logs/msbuild-20260713_135435_810.log`.
- Perf scenario: repeated SQLite-authoritative search queries after warm-up, followed by an external store
  generation bump. This is an optimization/stabilization of query responsiveness; the guarded risk is an
  unnecessary SQLite open on every query without delaying rotation detection. New count metrics are
  `search.backend.sqlite.store_generation_probe_opens`, `.store_generation_probe_skips`, and
  `.store_generation_refreshes`, emitted once per validation decision. The deterministic command was
  `Tools/Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter local_search_sqlite_generation_probe_skips_steady_state_and_detects_bump`
  on Debug x64, machine `4cb089111a23`, workspace-default NTFS test root, with no special load assumptions.
- Same-machine/same-suite diagnostic baseline `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_133517/`
  recorded two probe opens (one unnecessary steady-state open plus the required bump probe) and failed the
  zero-open budget. Candidate `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-13_135810/` passed in
  50 ms: the steady query recorded one skip
  and zero opens; the bump query recorded exactly one open and one refresh, reducing total opens from 2 to
  1 and steady-state opens from 1 to 0 while preserving next-query rotation detection. `CompareTestRuns.ps1`
  reported the case transition from failed to passed and duration 51 ms to 50 ms. This is deterministic
  count evidence, not a percentile claim, so `Show-PerfRuns -FailOnQuality` sample-quality gating is not
  applicable.

## Task 4: Settings, Themes, Connections, and Secret Durability

**Files:**

- Modify: `Common/Common/SettingsStore.cpp`
- Modify: `Common/Common/SettingsStore.h`
- Modify: `Common/Common/ThemeDefinitionIo.cpp`
- Modify: `RedSalamander/SettingsHotReload.cpp`
- Modify: `RedSalamander/ConnectionManagerWindow.cpp`
- Modify: `RedSalamander/HostServices.cpp`
- Modify: `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp`
- Modify: `Common/DxUi/DxUi.Theme.cpp`
- Modify settings/theme/connection selftests

- [x] **Step 4.1: Preserve rotated secrets when persistence fails [CW-2].**

  In `HostServices::SetConnectionSecretOnUiThread`, update the in-memory `SessionSecretEntry` first for
  set and delete paths, securely clearing old values. Attempt `SaveGenericCredential` or
  `DeleteGenericCredential` second. A persistence failure must log and return the failure code while the
  session cache already contains the new secret state.

  Add a minimal fault seam for `SaveGenericCredential`/`DeleteGenericCredential`.

  RED proof: seed OAuth token, rotate under injected CredMan failure, then
  `GetConnectionSecretOnUiThread` returns the new token.

- [x] **Step 4.2: Parse inline themes leniently without serialize/re-parse [BR-C2, BR-C1].**

  Split `ParseThemeDefinitionJson5` into a text front-end and
  `ParseThemeDefinitionFromValue(const Common::Settings::JsonValue&, ThemeDefinition&, ParseMode, errors)`.
  Theme files use strict mode. Settings inline themes use lenient mode.

  Lenient mode must skip invalid/unknown color entries with one aggregated warning, clamp over-long names,
  preserve unknown base theme ids for fallback at apply time, and keep structurally unusable raw theme
  entries in an opaque round-trip list so save/reload does not delete them.

  RED proofs: one bad color does not drop the theme; structurally broken inline theme survives load/save
  byte-equivalent enough for round-trip preservation.

- [x] **Step 4.3: Make hot reload non-blocking and not lossy [EV-C1, EV-C2, BR-C3].**

  Preserve the no-blocking `Start` behavior, but capture the settings file stamp before spawning the
  watcher. The watcher must signal readiness as "alive and retrying", create or tolerate a missing settings
  directory, and perform a catch-up comparison against the initial stamp once notifications are armed.

  In the `WAIT_FAILED` path, capture `GetLastError()` immediately before any `Stop()` call can clobber it.

  RED proofs: first run without settings directory returns fast and later hot reload works; a write landing
  before notification arm is detected by the catch-up pass.

- [x] **Step 4.4: Keep edited connection selection coherent [EV-C3].**

  After `TryValidateAndNormalizeConnectionProfiles`, refresh `_selectedConnectionName` from the resolved
  edited model index even when the grid selection is empty. Apply the same validation-before-result rule
  on Connect and Close: Connect returns the normalized selected profile name; Close saves normalized profile
  edits but returns no connection result.

  RED proof: clear grid selection, rename in editor, Connect produces saved settings and modal result with
  the same connection name.

- [x] **Step 4.5: Correct theme accent pressed contrast [CW-9].**

  Change `RefreshAccentVariants` to accept `darkBase` and choose the pressed shift from the base theme,
  not `darkMode`. Pass the correct base flag from default and viewer-theme palette builders.

  RED proof: viewer theme with `darkMode=true`, `darkBase=false`, and mid-luminance accent produces
  `accentPressed` darker than `accent`.

## Task 5: DxUi Accessibility, Text Input, and Animation Correctness

**Files:**

- Modify: `Common/DxUi/DxUi.Accessibility.cpp`
- Modify: `Common/DxUi/DxUi.Controls.cpp`
- Modify: `Common/DxUi/DxUi.h`
- Modify: `Common/DxUi/DxUi.Menu.cpp`
- Modify: `Common/DxUi/DxUi.NativeTextInput.cpp`
- Modify: `Common/DxUi/DxUi.SingleLineTextEditing.cpp`
- Modify: `Common/DxUi/DxUi.TextInput.cpp`
- Modify: `Common/DxUi/DxUi.WindowHost.cpp`
- Modify DxUi tests under `Tests/DxUiTests/`

- [x] **Step 5.1: Bound UIA dispatch lifetime and timeout semantics [BR-B1, BR-B2].**

  Add `Pending/Taken/Abandoned` atomic state to `AccessibilityUiActionDispatch`. Waiter timeout CASes
  `Pending -> Abandoned`; handler CASes `Pending -> Taken` before execution; drained payload wrapper CASes
  `Pending -> Abandoned`, records `ERROR_CANCELLED`, and signals completion from the wrapper destructor.

  If timeout races with a handler that already took the dispatch, perform a zero-timeout final wait and
  return the real result instead of `ERROR_TIMEOUT`.

  RED proofs: destroying a window with pending dispatch returns promptly with `ERROR_CANCELLED`; a timed
  out mutating Select action does not execute later; a just-in-time taken action does not double-execute.

- [x] **Step 5.2: Bound context-menu debug state cross-thread probes [BR-G3].**

  Reuse the UIA dispatch state-machine pattern for `kMenuDebugGetStateMessage`. Use heap-owned request
  storage, a completion event, and `Pending/Taken/Abandoned` state so timeout cannot write into abandoned
  stack storage and cannot hang unbounded.

  RED proof: wedged menu thread returns failure within the selftest-scaled timeout; unwedged probe returns
  the correct state.

- [x] **Step 5.3: Reduce high-cost grid accessibility work without duplicating UIA baton scope [EV-B1, BR-B3, BR-B4].**

  Cap materialized offscreen selected-row records, build a visible-row id set once, and defer offscreen row
  names if a provider can resolve lazily. Keep all-columns extraction intact; only remove the redundant
  `gridAccessibleColumns` identity vector by replacing it with index arithmetic.

  Make `ScrollPanel::GetViewportRect()` public, delete the duplicate accessibility helper, and call the
  member.

  Perf proof: Ctrl+A over a 10k-row grid meets the snapshot rebuild budget; existing GR-T9/grid a11y tests
  stay green.

- [x] **Step 5.4: Fix TSF deactivation lifetime and clean dead teardown scaffolding [EV-B4, EV-B5, EV-B6].**

  Add a `weak_ptr` lifetime token for text input controls and check it in every text store entry point that
  dereferences `_control`. Before `Pop(TF_POPF_ALL)`, verify host/control liveness through that token and
  tree membership. Use the keep-attached ordering only for live controls; otherwise disconnect/detach before
  pop.

  Delete the no-op `shouldReleasePreviousFocusDocumentAfterPop` and collapse duplicate Disconnect/Detach
  calls. Remove the temporary per-line trace-file reopen scaffolding after the lifetime test passes.

  RED proof: destroy-during-deactivate test no longer dereferences stale host/control.

- [x] **Step 5.5: Restore secure text and masked-count contracts [EV-B7, EV-F6, CW-15].**

  When `secureCacheText` is true, securely clear the existing layout-cache text before assignment and copy
  into an exact-size buffer so capacity tails do not retain revealed password text.

  Strengthen the secure-clear contract assertion to match the meaningful call tail instead of a bare
  `"true)"` substring.

  Restore grapheme counting for exact password mask length via a `CountTextElements(std::wstring_view)`
  helper. Treat the previous UTF-16 code-unit behavior as a bug unless a separate perf record proves and
  accepts the regression.

  RED proofs: secure cache replacement wipes old text; one BMP character plus one astral character masks
  as two dots, not three.

- [x] **Step 5.6: Resume animations after restore-from-minimize [CW-8].**

  In the `WM_SIZE` handler, after `OnSize()`, if size is restored or maximized, the host is effectively
  visible, and `_animationSuspendedWhileHidden` is true, clear the flag and call `RequestAnimation()`.

  RED proof: iconic host latches suspension, `WM_SIZE/SIZE_RESTORED` re-arms animation.

- [x] **Step 5.7: Close low-risk DxUi cleanup [CW-S1, CW-S2, CW-S3, EV-C4, EV-D7, EV-E11].**

  Delete dead `FindSemanticControlAtPoint` and its source-text pin if no behavioral use remains.
  For CW-S2, add a short comment or closeout note that OOM remains fatal by repo policy; no code change.
  Change TabItem title-width cache format ownership to `wil::com_ptr<IDWriteTextFormat>` and reset it at
  the existing invalidation sites.

  Delete the unreachable `case WM_NCDESTROY` in `ConnectionManagerWindow.cpp`, remove the misplaced
  `RestoreDeferredFocusAfterLayout()` call in Preferences, clear stale deferred focus during detach, and
  remove unreachable `continue` statements after fatal `Require` calls.

### Tasks 4-5 closeout evidence — 2026-07-13

- **Task 4:** Session-secret mutation is now authoritative before CredMan persistence, with injected
  save/delete failures covered by `connection_secret_persistence_failure_keeps_session_rotation`.
  Inline themes parse directly in lenient mode with opaque malformed-entry round trips; settings hot
  reload captures its initial stamp, tolerates a missing directory, performs catch-up after arming, and
  preserves `WAIT_FAILED` diagnostics. Connection validation now refreshes the edited selection before
  Connect/Close results, and pressed-accent derivation follows the base theme. Focused settings,
  connection, theme, and hot-reload cases passed inside the green Commands and full-suite aggregates.
- **Task 5:** UIA actions and menu debug probes use heap-owned `Pending/Taken/Abandoned` dispatches;
  grid accessibility bounds eager offscreen materialization and reuses the public viewport; TSF entry
  points use a weak lifetime token; secure cached text is wiped before exact-size replacement; password
  masks count grapheme clusters; restore-from-minimize re-arms animation; dead cleanup scaffolding was
  removed and TabItem text formats now use `wil::com_ptr`. DxUi behavioral coverage and the archived
  10k-row accessibility perf proof are green. The candidate snapshot rebuilt in 3,166 us with 256
  materialized offscreen rows versus the 61,938 us / 9,995-row baseline under
  `Specs/TestRuns/4cb089111a23/DxUiTests/2026-07-13_1412` (about 19.6x faster).

## Task 6: FileOps, Providers, MOVE Verification, and Data-Safety Degrade Paths

**Files:**

- Modify: `RedSalamander/FolderWindow.FileOperations.State.cpp`
- Modify: `RedSalamander/FolderView.FileOps.cpp`
- Modify: `RedSalamander/FolderViewInternal.h`
- Modify: `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp`
- Modify: `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp`
- Modify: `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- Modify: `RedSalamander/FolderWindow.FileSystem.Navigation.Part.cpp`
- Modify: `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- Modify: `Plugins/FileSystem/FileSystem.FileOps.cpp`
- Modify: `Plugins/FileSystem/FileSystem.Path.cpp`
- Modify: `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`
- Modify: `Plugins/FileSystem7z/FileSystem7z.cpp`
- Modify: `Plugins/FileSystemS3/FileSystemS3.IO.cpp`
- Modify fileops/provider selftests

- [x] **Step 6.1: Land the cross-FS MOVE safety regression test [CW-1/CW-T1].**

  Keep the existing fixed implementation if present, but add the missing regression test: copy a two-level
  tree, mutate the source root directory basic info before cleanup, and assert matching children are
  deleted while the source shell remains with `ERROR_PARTIAL_COPY`.

  Run the focused FileOps selftest and preserve the RED note if the current tree unexpectedly fails.

- [x] **Step 6.2: Add copy-time content hash and shrink MOVE manifests [CW-P1, CW-P2].**

  **2026-07-13 split status:** Causeway CW-9 completed the shared CW-P1 implementation in both the host
  bridge and native local cross-volume MOVE: the bridge uses copy-time FNV-1a plus destination-only cleanup
  hashing, while the native path uses successful-copy plus stable-snapshot proof with zero post-copy content
  rereads. The old dual-reader cleanup compare is removed. Firebreak retains only the distinct CW-P2
  manifest-retention/extract-and-erase work and its manifest-size evidence; do not reimplement the hash.

  Add `hasContentHash` and `contentHash` to `CopiedEntry`. Compute a streaming hash over bytes already
  flowing through both serial and pipeline copy loops; store it with `RecordCopiedFile`. In
  `CopiedFileStillMatchesDestination`, after size match, re-hash only the destination when a copy-time hash
  exists; fall back to full `ReadersHaveEqualContent` only for legacy/no-hash entries.

  Gate manifest recording to MOVE operations. Change lookup in `DeleteCopiedSourceEntryForMove` to
  extract-and-erase entries under `copiedEntriesMutex` so memory is released as cleanup progresses.
  Remove misleading `noexcept` from internal manifest helpers that perform recoverable bookkeeping
  allocations, while keeping `std::bad_alloc` fatal by repo policy. Do not catch or swallow
  `std::bad_alloc`.

  Perf proof: archive FileOps baseline/candidate read-byte/time metrics and manifest-size metrics under
  `Specs/TestRuns/`. RED proof: source re-open fault during cleanup still deletes source after destination
  hash verification; manifest size strictly decreases to zero.

- [x] **Step 6.3: Fix staged promote and symlink overwrite error quality [CW-13, CW-11].**

  In `PromoteStagedTempIntoFinalPath`, if final `SetFileAttributesW` fails while attributes differ,
  return `HRESULT_FROM_WIN32(lastError ? lastError : ERROR_ACCESS_DENIED)` instead of falling through to
  `S_OK`.

  In `CopyReparsePointInternal`, restore the `!allowOverwriteEffective` guard around remapping
  `ERROR_INVALID_PARAMETER` to `ERROR_NOT_SUPPORTED`.

  RED proofs: injected final attribute failure returns FAILED; staged overwrite injected
  `ERROR_INVALID_PARAMETER` is not mislabeled.

- [x] **Step 6.4: Make directory-size and archive indexing degrade instead of abort [CW-5, CW-4].**

  Add `isNonFatalChildError(DWORD)` for sharing/lock/network/device-not-ready child errors. In both
  `pushDirectory` and `advanceFrame`, mark partial and continue siblings for those errors; reserve fatal
  status for global failures.

  In 7z `BuildIndex`, when a file shadows a needed parent directory, skip creating the synthetic directory
  and warn. On duplicate normalized keys, keep the first indexed entry and continue. Only add a child key
  to a parent when insertion actually occurred.

  RED proofs: directory with one sharing-violation child still includes later sibling sizes and returns
  partial; archives with `foo` plus `foo/bar` or `foo/` open and list root.

- [x] **Step 6.5: Let S3 ranged reads survive HEAD-denied buckets [CW-3].**

  Parse `GetObjectResult::GetContentRange()` in `FillBufferFrom` to discover total size. Make
  `EnsureSizeKnown` non-fatal on HEAD-denied by logging and leaving `_sizeKnown=false`; reads and size
  discovery must proceed through a ranged GET and validate against parsed total.

  RED proof: `HeadObject=403` plus successful ranged GET yields `GetSize()==S_OK` and full read.

- [x] **Step 6.6: Fix FolderView/FileOps wedge and wrong-context notifications [EV-D1, EV-D2, EV-D3].**

  If paste-shortcut completion `PostMessagePayload` fails, post a sentinel when possible, reset
  `_pasteShortcutInFlight` from the UI side, and add a stale timeout/`WM_NCDESTROY` reset as a backstop.
  Log the failure once. Capture the original filesystem/provider identity in paste shortcut request/result
  and notify `DirectoryInfoCache` for that context, not the pane's current filesystem.

  Stop freeing reserved temp-file names before `IPersistFile::Save`; keep the placeholder or use a GUID
  same-volume temp path.

  RED proofs: lost completion does not wedge future paste shortcuts; navigation mid-work notifies the
  provider that changed; concurrent temp saves cannot steal the path.

- [x] **Step 6.7: Close FileOps/UI simplification defects [EV-D4, EV-D5, EV-D6, EV-D8, KS-S5].**

  Reuse one issues-pane close path that saves view state, hides the pane, and restores focus. Hoist one
  guarded prompt-close debug helper that keeps the `message == 0` guard. Extract
  `NavigationView::RestoreFolderViewFocusAfterDropdown()`. Lock FileSystemDummy instance count updates
  and root cleanup with the same mutex. Remove dead `SkipAll` file-op branches.

  Verify affected selftests and one manual-style FileOps close/hide scenario.

  **2026-07-13 completion:** Task 6 is complete. The bridge records manifests only for MOVE,
  extract-and-erases entries during cleanup, emits peak/remaining counters, and retains the Causeway
  copy-time hash proof without reimplementing it. The two-level root-basic-info regression passed with
  manifest count `3 -> 0` under `Specs/TestRuns/4cb089111a23/FileOps/2026-07-13_134907`.
  Staged promotion/error fidelity, directory-size partial continuation, and Skip + All similar passed in
  `FileOps/2026-07-13_135843`, `FileOps/2026-07-13_135851`, and `FileOps/2026-07-13_135858`.
  Paste Shortcut completion recovery/provider-context coverage passed in
  `Commands/2026-07-13_140247` and `Commands/2026-07-13_141006`; the issues-pane close/hide scenario
  passed in `Commands/2026-07-13_141046`. `PluginContractTests.exe` passed, including local filesystem
  `52/52`, 7z `20/20`, and S3 `133/133` debug assertions. Full evidence and the successful full Debug x64
  rebuild (`.build/logs/msbuild-20260713_140629_521.log`) are recorded in
  `Specs/TestRuns/4cb089111a23/Continuation/2026-07-13_141100_firebreak_task6_closeout/README.md`.

## Task 7: FolderView Icons, Rendering Metrics, and IconCache Resilience

**Files:**

- Modify: `RedSalamander/FolderView.Icons.cpp`
- Modify: `RedSalamander/FolderView.Rendering.cpp`
- Modify: `RedSalamander/FolderView.cpp`
- Modify: `RedSalamander/FolderView.h`
- Modify: `RedSalamander/IconCache.cpp`
- Modify: `RedSalamander/FolderView.Enumeration.cpp`
- Modify FolderView selftests and perf records

- [x] **Step 7.1: Fix pending bitmap and thumbnail handshakes [BR-E1, BR-E4].**

  In the thumbnail result handler, perform stale-batch check first, arm the pending-count decrement
  scope_exit second, then perform stale-generation checks. After a provider-probe waiter marks
  `abandoned=true`, perform a final zero-timeout wait; if signaled, clear abandoned and consume the result.

  RED proofs: stale-generation message decrements `pendingBitmapCreates`; both-sides-skipped provider
  probe interleaving is impossible by counter assertion.

- [x] **Step 7.2: Clear all pending paint metrics after failed frames [BR-E2].**

  Replace `ClearReadyPendingRefreshToPaintMetric()` calls on render target absence, `EndDraw` failure,
  `Present1` failure, and legacy present failure with `ClearPendingPaintMetricsOnFailedFrame()` that clears
  both refresh-to-paint and input-to-paint pending samples.

  RED proof: forced `EndDraw` failure after input metric arm does not emit inflated
  `folder.frame.input_to_paint_us` on the next successful present.

- [x] **Step 7.3: Reinstate bounded live icon negative caching [BR-E3].**

  Cache live path icon lookup failures with per-path backoff: short first TTL, doubling to a cap, reset on
  success. Preserve the uncached failure counter and add a cached-hit counter.

  Perf proof: N refreshes inside the backoff window issue one shell lookup; retry occurs after the window;
  success clears the negative entry.

- [x] **Step 7.4: Replace self-fulfilling DrawItem brush-count instrumentation [BR-E5].**

  Delete the force-create flag/hook in `DrawItem`. Instrument the real D2D brush factory/cache seam to
  increment a counter for transient non-cached brush creation during item draw. The selftest asserts zero
  transient brushes during scroll/redraw. Locally hand-inject a temporary brush allocation to prove the test
  goes red, then remove the injection before committing.

  Perf proof: archive rendering/FolderView metric evidence when this touches hot draw paths.

### Task 7 closeout evidence — 2026-07-13

- **7.1:** The provider-probe waiter and worker now share an atomic result claim, including the final
  zero-timeout handoff after abandonment, so exactly one side consumes or posts each result. Current-batch
  stale-generation bitmap messages now decrement their pending count after the stale-batch gate. Focused
  evidence passed in `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_141933` and
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_141939`; the handoff run records exactly one
  `thumbnails.provider_probe_result_claimed` row.
- **7.2:** Render-target absence plus `EndDraw`, `Present1`, and legacy `Present` failures now clear both
  input-to-paint and refresh-to-paint pending samples. The behavioral guard arms both paths, forces failed
  frames, and proves no stale sample appears on later successful presents in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_141944`.
- **7.3:** Live path failures now use a per-key `250 ms` exponential backoff capped at `4 s`; successful
  retry resets the failure history and installs the positive cache result. Same-run evidence in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_141950` records shell lookups `8 -> 1` for eight forced
  refreshes without/with the negative cache, plus two failed retries, nine negative-cache hits, and a final
  successful lookup followed by a positive-cache hit.
- **7.4:** The self-fulfilling force-allocation hook is gone. Actual FolderView solid-brush creation routes
  through the cached/transient lifetime seam, which counts only successful transient brush creation while
  `DrawItem` is active. A temporary local injection made the guard fail with count `12` in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_141658`; after removing it, the same guard passed with
  `folder.draw_item.transient_brush_create_count=0` in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_141929`. The adjacent 10,000-item scale/perf guard passed
  in `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_142008`.
- **Build/specs:** Debug x64 `RedSalamander` built with 0 errors in
  `.build/logs/msbuild-20260713_141713_811.log`; its two C5245 warnings are confined to concurrently edited
  Compare Directories selftest helpers outside Task 7. Durable behavior and coverage are merged into
  `Specs/UI/UI_FolderView.md` and `Specs/Testing/Testing_TestCoverage.md`.

## Task 8: Provider Capabilities, Batch Rename, and Path Identity

**Files:**

- Modify: `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp`
- Modify: `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp`
- Modify: plugin capability JSON/specs
- Modify: `RedSalamander/BatchRenameWindow.cpp`
- Modify: `RedSalamander/BatchRenameEngine.cpp`
- Modify: `RedSalamander/FileSystemPathIdentity.cpp`
- Modify: `RedSalamander/FileSystemPathIdentity.h`
- Modify: `Plugins/FileSystem/FileSystem.cpp`
- Modify batch rename/provider tests

- [x] **Step 8.1: Make advertised case-only rename capabilities honest [KS-2, KS-3].**

  Curl: drive self-rename detection from `pathIdentity.componentComparison`. For protocols with a rename
  operation, execute the case-only rename instead of treating it as a case-insensitive no-op. For protocols
  without a reliable rename operation, downgrade `caseOnlyRename` to `noOp`.

  Microsoft Drive: when same item id differs only by leaf case, issue the Graph name PATCH. If the live
  API cannot make the change, report `caseOnlyRename:"noOp"` instead of `"supported"`.

  Proof: case-only rename changes the remote leaf or capability explicitly says it will not.

- [x] **Step 8.2: Keep plugin debug selftests honest [KS-5].**

  In `PluginContractTests.cpp`, treat a missing `RunDebugSelfTests` export as failure for plugins that
  are expected to ship the export in the current configuration, or as explicit xfail/skip when the plugin
  configuration is not expected to expose it. Removing an expected export must turn the step red.

- [x] **Step 8.3: Resolve byte-identical overwrite and FindFiles partial-result semantics [KS-6, KS-7].**

  For byte-identical existing-file copies, report a host-visible skip and document that metadata/ACL/ADS are
  not synchronized by this optimization. Do not silently report a full overwrite.

  For partial FindFiles move/delete completion, surface per-source results and remove only rows that are
  known completed. Preserve the fail-safe behavior for uncertain rows.

- [x] **Step 8.4: Move BatchRename hot-path work off the UI thread [KS-P1, KS-P2].**

  Debounce local destination conflict validation, cache one directory listing per preview refresh, and move
  remaining synchronous stat calls off the UI thread. Collect provider selection targets off-thread and show
  progress instead of running network/archive I/O in `SetContext`/`Create`.

  Perf proof: define scenarios for per-keystroke preview and provider selection collection, add metrics,
  and archive baseline/candidate runs.

- [x] **Step 8.5: Unify path identity folding and storage-medium probing [KS-P4, KS-P6, KS-A1, KS-A2].**

  Make the key builder fold non-ASCII consistently with `CompareStringOrdinal` where possible so duplicate
  detection retains the O(n) hashed path. Document `nullopt` as "caller must fall back to
  `EquivalentPath`" in `FileSystemPathIdentity.h`. Document the rare NTFS `$UpCase` approximation or
  add an existence check before treating rare-codepoint names as the same object.

  Cache `ProbeLocalStorageMedium` results outside the concurrency-resolution path and invalidate the cache
  only on provider/context changes.

  Perf proof: non-ASCII large batch does not fall into O(n^2) duplicate detection; storage-medium probe no
  longer appears on the hot path.

### Task 8 closeout evidence — 2026-07-13

- **8.1:** Curl self-rename checks now follow the advertised case-sensitive path identity, and Microsoft
  Drive same-item case-only leaf changes continue to the Graph name PATCH instead of becoming a no-op.
  Focused Debug selftests passed Curl 78/78 and Microsoft Drive 79/79.
- **8.2:** Debug plugin configurations now require `RunDebugSelfTests`; intentional non-Debug absence is an
  explicit skip. Direct `PluginContractTests.exe` validation passed FileSystem 52/52, 7z 20/20,
  Microsoft Drive 79/79, S3 133/133, and Curl 78/78. The durable rule is in
  `Specs/Testing/Testing_SelfTests.md`.
- **8.3:** A byte-identical destination is reported as per-item `S_FALSE` without rewriting destination
  metadata/ACL/ADS, directory-copy aggregation preserves that partial result, and Find Files removes only
  source rows with known `S_OK` outcomes. Evidence: FileOps rerun/resume archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-13_135004` (3/3) and Find partial-completion archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_134952` (1/1).
- **8.4:** Provider collection and preview planning/validation run on module-pinned MTA workers; create,
  context changes, and execution no longer perform provider enumeration or preview rebuilds on the UI
  thread. Destination validation caches one listing per parent per preview. Metrics include
  `batchrename.collect.us`, `batchrename.preview.destination_validation.us`, and
  `batchrename.preview.destination_directory_listings`. The gated-provider scenario is a deterministic
  pre-fix substitute: synchronous collection would fail the under-one-second show/close assertions. The
  final teardown stress passed 5/5 and safely accepted late worker completion after window destruction in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_140944`. Folder-scope candidate evidence in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_134935` records 530–1,092 us provider collection and at
  most two destination listings for the fixture's two parents.
- **8.5:** `FileSystemPathIdentity` now uses invariant uppercase component folding, rejects expanding folds
  so callers fall back to `EquivalentPath`, and documents the NTFS `$UpCase` approximation. The 10,000-row
  non-ASCII scenario passed in 643 ms with `batchrename.preview.duplicate_fallback_rows=0` in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-13_134947`; the prior ASCII archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-10_161852` took 696 ms. The candidate is deliberately
  stronger input, so the proof is bounded completion with zero quadratic fallback rather than a raw
  build-plan speedup. Storage-medium evidence in
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-13_135004` records one probe-cache miss followed by eight
  hits, proving the physical probe is no longer repeated on the concurrency hot path.
- **Build/teardown:** Task 8's final source and teardown test are present in the zero-warning
  `RedSalamander` build `.build/logs/msbuild-20260713_135906_910.log`. The earlier mixed-build teardown
  crash, dump pointers, failed archive, and clean rerun are preserved in
  `Specs/TestRuns/4cb089111a23/Continuation/2026-07-13_134451_firebreak_task8_batchrename_stale_build_crash/README.md`.
  Durable behavior is merged into the FileSystem, plugin, Batch Rename, Find Files, and testing specs.

## Task 9: Test Harness, Tooling, and Shared Helper Cleanup

**Files:**

- Modify: `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- Modify: `Tools/Tests/ShowPerfRuns.Tests.ps1`
- Modify: `Tools/Tests/ResourceLocalizationContracts.Tests.ps1`
- Modify: `Tools/Show-PerfRuns.ps1`
- Modify: `Common/Helpers.h`
- Modify: `Common/Common/SettingsStore.h`
- Modify: `RedSalamander/SelfTest/SelfTestCommon.*`
- Modify: `RedSalamander/SelfTest/Commands/*.cpp`
- Modify: `RedSalamander/SelfTest/CompareDirectories/*.cpp`
- Modify plugin files touched by helper hoists

- [x] **Step 9.1: Remove source-scraping tests after behavioral twins exist [BR-G4, EV-E7].**

  Inventory source-text assertions added by the reviewed windows. Delete style/wording/member-order pins.
  For real behavior, first add or keep behavioral tests from Tasks 5, 7, and 9, then delete the scrape.
  Normalize the "unused"/"unreferenced" pragma wording divergence if the pinned text caused drift.

  Verify with Pester and grep: no new `Get-Content`-based product `.cpp` source assertions remain from
  these windows.

- [x] **Step 9.2: Fix UIA/raw-provider and compare-options selftest traps [EV-E1, EV-E2, EV-E3, EV-E4, EV-E5, EV-E6].**

  Make raw UIA provider fallback reachable by marshalling provider creation to the window thread. Keep the
  fallback, enforce the visibility contract, and avoid double-invoking when UIA `Invoke()` reports failure
  after performing the action.

  Trace or assert when OK/Cancel tests use the keyboard backdoor so InvokePattern regressions remain
  visible. Shorten the inner edit-value retry timeout so retries can occur. Remove illusory cancellation in
  diagnostics or mirror the detach-on-timeout pattern used by `RunUiaActionWithMessagePump`.

- [x] **Step 9.3: Harden test filesystem helpers [EV-E8, EV-E9, EV-E10, BR-G6, BR-G7].**

  Use the `\\?\`-aware shortcut path query on failure-path shortcut checks. Hoist JSON integer extraction,
  stable device hash, repo-root walking, local app data path, directory creation, UTF-8 file writing, and
  MTP selftest export loading into shared selftest helpers.

  Shared MTP export loading must use production-like guards: disabled/loadable/path checks,
  `IsPluginPathDeferred`, `LoadLibraryExW` pin, pragma-4191 `GetProcAddress`, and typed call.

  Verify from normal checkout, git worktree, and `REDSALAMANDER_REPO_ROOT` override.

- [x] **Step 9.4: Clean remaining selftest/tool quality issues [EV-E12, EV-F2, EV-F3, EV-F4, EV-F5, BR-G5].**

  Reword ViewerVLC focus checks so they assert what product routing really guarantees and record fallback
  use. In `Show-PerfRuns.ps1`, respect explicit `-MinimumSamplesForP95`, warn and fail quality when an
  existing budget file cannot parse, extract `Invoke-RSPwsh`, and add a `-BudgetPath` override to tests.

  Add a localization contract guard that fails when a new top-level `*.rc` root is not allowlisted. Point
  ViewerSpace env flags at `EnvironmentVariables::IsTruthyFlagSet` and delete the local truthy parser.

- [x] **Step 9.5: Hoist production shared helpers [BR-H1, BR-H2, KS-S1, KS-S2, KS-S3, KS-S4].**

  Add `Common::Settings::FindMember` plus typed `GetString`, `GetWString`, `GetBool`, `GetUInt32`, and
  `GetArray` accessors using the strictest copied behavior. Migrate FileSystemPluginManager,
  connection-profile utilities, ConnectionManagerWindow, HostServices, and command selftests.

  Add one `Common::Strings::Utf16FromUtf8(std::string_view)` helper with documented replacement-character
  behavior. Migrate at least the copies added by the reviewed windows, and opportunistically migrate files
  already open.

  Add shared crypto-random/hex and unique sibling/temp-name helpers for S3, Curl, MSDrive, and BatchRename.
  Extract duplicated plugin debug harness code behind `_DEBUG`/test-only boundaries. Route BatchRename
  separator/path decomposition through `FileSystemPathIdentity`.

  Verify full build because these headers are widely included.

### Task 9 closeout evidence — 2026-07-13

- **9.1:** Redundant source-text checks were removed after retaining or adding compiled behavioral twins,
  including the stale DxUi title-width implementation scrape. The focused tooling set
  (`ShowPerfRuns`, resource-localization contracts, source contracts, and inventory) passed 135/135.
- **9.2:** Test raw-provider creation now marshals to the owning window thread, UIA visibility requires
  `IsOffscreen == false`, failed post-action `Invoke()` calls are not invoked twice, and Compare Options
  reports keyboard fallback while using bounded message-pump retries. The complete DxUi Control suite
  passed after the stale source guard was removed.
- **9.3:** Commands and Compare now share repository/Local AppData/filesystem/JSON/device-hash/MTP-export
  helpers. The MTP loader follows disabled, loadable, deferred, module-pin, and typed-export guards; the
  repository-root helper accepts both checkout/worktree metadata and the explicit environment override.
- **9.4:** ViewerVLC focus coverage now measures the product routing contract and records fallback. The
  perf analyzer honors an explicit p95 minimum, accepts a test budget path, and fails quality on an
  unparseable existing budget. Localization rejects new unallowlisted top-level resource roots, and
  ViewerSpace uses the shared truthy environment parser. The ViewerVLC HUD/loading/teardown contract
  passed three consecutive focused runs; its forced-parent proof uses a 1,000 ms injected cleanup delay
  and a 250 ms nonblocking ceiling to remain stable under suite load.
- **9.5:** Shared strict Settings JSON accessors, replacement-policy UTF-8 conversion, cryptographic
  random/hex and unique sibling-name helpers, debug result checks, and provider-aware path helpers replaced
  the reviewed copies. Batch Rename debug rule setters now wait for asynchronous preview settlement, and
  all six previously order-sensitive focused cases pass.
- **Build/tool evidence:** the full Debug x64 solution rebuild passed with 0 warnings/errors in
  `.build/logs/msbuild-20260713_144107_919.log`; subsequent RedSalamander, DxUiTests, and ViewerPETests
  rebuilds passed in `.build/logs/msbuild-20260713_152810_735.log`,
  `.build/logs/msbuild-20260713_153208_621.log`, and
  `.build/logs/msbuild-20260713_153244_930.log`. `git diff --check` is clean. The earlier full runner exposed
  the now-fixed Task 9 Batch Rename, DxUi source-guard, and ViewerVLC timing failures; remaining full-run
  failures belong to other Firebreak tasks and are not claimed as Task 9 evidence.

## Task 10: Blocked or Routed Items

These items are intentionally not executable inside Firebreak without a prerequisite.

- [blocked-human] **EV-G1 repo cleanup:** requires explicit user approval to delete local branch
  `codex/folderview-warpdrive`, prune snapshot refs reaching artifact commits, and run `git gc`. Before
  approval, only re-run verification diffs that prove clean twins differ solely in artifact files.
- [routed] **EV-B2/EV-B3 and Bedrock all-columns extraction cost:** owned by
  `Specs/Plans/WIP/DxUi_Uia_ContinuationBaton_2026-06-29.md`.
- [routed] **Bedrock F#19 / perf measurement schema-machine filtering:** owned by
  `Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md`.
- [done-reference] **CW-7 and CW-10:** already fixed with evidence; keep the durable spec updates.
- [policy-closed] **CW-S2:** keep OOM fatal unless a future product decision requests recoverable layout
  cache allocation.

## Closeout Checklist

- [x] Every checked item has a RED/GREEN proof or a documented deterministic substitute.
- [x] Perf-sensitive items list scenario, metric key, baseline archive, candidate archive, and comparison.
- [x] `.\build.ps1 -Configuration Debug -Platform x64` succeeds.
- [x] `.\Tools\Run-AllTests.ps1 -SkipBuild` succeeds.
- [x] `.\Tools\Run-AllTests.ps1 -Suite Full` succeeds or every failure is tied to an external blocker.
- [x] Durable behavior is merged into the authoritative spec under `Specs/<Domain>/` or repo-level
  guidance.
- [x] Firebreak is moved to `Specs/Plans/Done/` only after all executable tasks are done or explicitly
  closed as routed/blocked with owner and missing input.

### Continuation checkpoint — 2026-07-13

- All executable Tasks 0-9 are implemented and checked. Task 10 remains explicitly routed,
  human-blocked, done-reference, or policy-closed as recorded above.
- Latest targeted Debug x64 build: green with zero diagnostics; log
  `.build/logs/msbuild-20260713_223617_913.log`.
- Focused SearchService readiness stress after replacing the legacy 10-second default with the shared,
  scaled 30-second bound: `search_service_sqlite_legacy_auto_vacuum_queues_idle_maintenance` and
  `search_service_sqlite_status_reports_maintenance_history` passed five repetitions each (10/10).
- Tooling Pester: 254 passed, 0 failed after updating the readiness source contract.
- Final-state no-build aggregate: run
  `20260713T200325Z-8564-41dffa892a5a4c46930546394e298ab6`; 1,132 passed, 0 failed, 53 documented
  skips; all three suites passed with zero flaky, regression, isolation-suspect, or unclassified
  failures.
- A complete full gate before the shared readiness-default follow-up was green: run
  `20260713T184605Z-63640-08c95360966a43a5bea08c0148451d78`; 1,146 passed, 0 failed, 53 documented
  skips; all 17 suites passed and the build log had zero diagnostics.

### Final closeout — 2026-07-14

- Final source-stable Debug x64 build passed with 0 warnings and 0 errors in
  `.build/logs/msbuild-20260714_165926_087.log`.
- The complete tooling inventory passed: 260/260 non-build Pester cases plus the 1/1 serialized
  targeted-deployment build case. Inventory and authoritative documentation now agree on 699 static
  Commands registrations and 261 tooling cases.
- Final `Run-AllTests.ps1 -Suite Full` run
  `20260714T150136Z-77668-a61447b2b01640a19c32d02ed81a2c74` passed all 17 suites: 1,147 passed,
  0 failed, and 53 documented skips out of 1,200 aggregate cases. Classifications recorded 17 passed,
  0 flaky, 0 regression, 0 isolation-suspect, and 0 unclassified failures; quarantine was valid with no
  blocking entries.
- The pre/post SHA-256 manifest of every modified or untracked source/specification file matched, so the
  final Full result covered one stable source state. Aggregate evidence is at
  `.build/TestSandbox/runs/20260714T150136Z-77668-a61447b2b01640a19c32d02ed81a2c74/artifacts/selftest/last_run/run-all-tests-results.json`.
- The prior Commands hang was diagnosed as a keyboard-switched native menu accepting an orphaned initial
  button release. The deterministic regression passed 10/10, the final Commands suite passed 800/800
  with two documented skips, and durable modal/menu selftest rules are in `Specs/Testing/Testing_SelfTests.md`.
  Full forensics and closeout evidence are preserved under
  `Specs/TestRuns/SINON/Continuation/theme-v2-preexisting-selftest-hang-20260714-150354/`.
- All executable Firebreak tasks are complete. Task 10 dispositions remain intentionally routed,
  human-blocked, done-reference, or policy-closed with their owner/missing input recorded above. Firebreak
  is therefore closed and moved to `Specs/Plans/Done/`.
