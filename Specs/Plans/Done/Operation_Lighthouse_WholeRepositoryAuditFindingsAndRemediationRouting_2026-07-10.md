# Operation Lighthouse — Whole-Repository Audit Findings and Remediation Routing

> **Executor instructions:** This document is both the durable finding ledger for the 2026-07-10
> whole-repository review and the single routing authority for those findings. Do not implement a
> finding marked `ROUTED` from this document; open and execute the named owner plan instead. For a
> finding marked `LIGHTHOUSE OWNER`, execute only one Lighthouse track per branch and preserve the
> verification gates and STOP conditions below. Do not combine the loader, settings, session-end, and
> CI tracks into one change. The 2026-07-14 duplication addendum is also Lighthouse-owned, but it is a
> sequence of characterization-first refactors rather than permission for a repository-wide mechanical
> replacement. Execute one duplication track at a time and preserve every named semantic variant.
>
> **Drift check (run first):**
>
> ```powershell
> git diff --stat 78366858d..HEAD -- `
>   Common/Common/SettingsStore.cpp Common/SettingsStore.h Common/Helpers.h `
>   Plugins/FileSystem/FileSystem.FileOps.cpp Plugins/FileSystemMtp/FileSystemMtp.Core.cpp `
>   Plugins/ViewerVLC/ViewerVLC.cpp RedLauncher/Main.cpp `
>   RedSalamander/ConnectionManagerWindow.cpp RedSalamander/FileActionLauncher.cpp `
>   RedSalamander/FileSystemPluginManager.cpp RedSalamander/FolderWindow.FileOperations.State.cpp `
>   RedSalamander/RedSalamander.cpp RedSalamander/ViewerPluginManager.cpp `
>   RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp `
>   Tools/TestRunPlan.ps1 .github/workflows/ci.yml
> ```
>
> If a cited symbol or behavior has changed, re-open the live code and update this ledger before
> implementation. Do not apply line-number-based edits blindly.
>
> **Duplication-addendum drift check (run before Tracks E-I):**
>
> ```powershell
> git diff --stat 0bee16269..HEAD -- Common Plugins RedConfigure RedSalamander `
>   RedSalamanderMonitor Tests
> ```
>
> Re-run the definition census and the focused `rg` gates in the applicable track. Line counts and names
> are discovery aids only; a helper is removed only after its policy and output are characterized.

## Status

- **Status:** DONE — routed prerequisites and Lighthouse Tracks A-I are implemented; the 2026-07-16 closeout
  build is clean, Lighthouse-owned test surfaces are green, and unrelated Full-suite blockers are accepted and
  routed to Observatory under the maintainer's convergence decision
- **Priority:** P0–P3
- **Overall effort:** XL across all owners; each Lighthouse-owned track is M–L
- **Overall risk:** HIGH because the set includes data-path security, plugin unload, settings migration,
  shutdown, and CI policy
- **Category:** security, correctness, reliability, performance, tests, architecture, maintainability, DX
- **Planned at:** commit `78366858d`, 2026-07-10
- **Duplication reconciliation:** commit `0bee16269`, 2026-07-14; no production source changed
- **Audit mode:** standard, hotspot-weighted production audit reconciled against the existing
  whole-application review campaign and current WIP owners
- **No source mutation:** the audit changed no production code and performed no commit or push

## Why this matters

The current tree builds cleanly, has broad deterministic test infrastructure, and follows the repository's
RAII/error-handling rules well. The remaining high-leverage risks are concentrated at trust and lifetime
boundaries: remote names entering local paths, plugin-reported sizes entering heap-buffer loops, DLL unload
while callbacks are active, nested scheduler waits, launcher/action executable resolution, and lossy settings
recovery. Fixing these boundaries first prevents arbitrary-path writes, memory disclosure, crashes,
deadlocks, configuration loss, and avoidable cloud-transfer cost.

This ledger also prevents duplicate work. Several findings already have detailed execution owners under
`Specs/Plans/WIP/`; Lighthouse records and routes them but directly owns only the gaps that did not have a
current WIP owner.

## Audit and verification baseline

Repository facts at the planned commit:

- Windows-native C++23/MSVC application, 1,344 tracked files, including 309 `.cpp` and 211 `.h` files.
- Canonical build: `./build.ps1 -Configuration Debug -Platform x64 -ProjectName RedSalamander`.
- Canonical PR-style test gate: `./Tools/Run-AllTests.ps1 -Suite CI`.
- Canonical closeout gate: `./Tools/Run-AllTests.ps1 -Suite Full`.
- Fresh audit build at `78366858d`: the scoped x64 Debug `RedSalamander` build completed in 3:26 with
  **0 warnings and 0 errors**.
- The newest committed Full-suite artifact found during the audit was not green: 1,098 passed, 2 failed,
  52 skipped. The failures were one CompareDirectories search-service case and one Commands navigation
  focus case. This is evidence of a verification-baseline gap, not evidence that the findings below caused
  those failures.
- Recent current-tree evidence is predominantly focused per-case coverage for File Operations popup work;
  it is not a substitute for a green current-head Full run.
- The 2026-07-14 duplication sweep covered production and first-party test C++ under `Common/`,
  `Plugins/`, `RedConfigure/`, `RedSalamander/`, `RedSalamanderMonitor/`, and `Tests/`. It combined a
  repeated-definition census with semantic review of each retained candidate. Representative census
  totals included 29 `Utf16FromUtf8`, 22 `Utf8FromUtf16`, 18 `EqualsNoCase`, 13 `BlendColor`, 11
  `StableHash32`, 10 `ColorFromHSV`, 9 `ColorRefFromArgb`, and 5 Windows absolute-path classifier
  declaration/definition matches. These counts include declarations and intentional variants; the
  findings below record only the consolidation families that survived source review.

## Finding index and ownership

| ID | Finding | Severity | Effort | Fix risk | Confidence | Owner/status |
|----|---------|----------|--------|----------|------------|--------------|
| LH-1 | Reject untrusted cross-filesystem child names before destination path composition | CRITICAL | M | MED | HIGH | `DONE`: Causeway CW-1/CW-1A |
| LH-2 | Clamp plugin-reported `bytesRead` before any buffer traversal or write | HIGH | S–M | LOW | HIGH | `DONE`: Causeway CW-2/CW-2A |
| LH-3 | Make MTP picker/cancel work module-safe across refresh and unload | HIGH | M–L | HIGH | HIGH | `DONE`: Firebreak Task 1 |
| LH-4 | Remove nested scheduler waits that can deadlock all File Operations | HIGH | M–L | HIGH | HIGH | `DONE`: Causeway CW-3 |
| LH-5 | Eliminate unsafe executable/DLL search-path behavior | HIGH | M | MED | HIGH | `DONE`: ViewerVLC dependency loading by Farsight; launcher/actions by Lighthouse Track A |
| LH-6 | Make settings recovery section-scoped, non-destructive, and forward-safe | HIGH | L | HIGH | HIGH | `DONE`: inline themes by Firebreak 4.2; core recovery by Lighthouse Track B |
| LH-7 | Persist runtime state during Windows session end without running modal teardown | MED | M | MED | HIGH | `DONE`: Lighthouse Track C |
| LH-8 | Stop re-reading both endpoints in full after every MOVE | HIGH | L | MED | HIGH | `DONE`: Causeway CW-9 / Firebreak 6.2 |
| LH-9 | Put critical contract/crash tests and an ARM64 build into the PR gate | MED | M | LOW | HIGH | `DONE`: Lighthouse Track D |

## Detailed findings

### LH-1 — Reject untrusted cross-filesystem child names before destination path composition

- **Evidence:** `RedSalamander/FolderWindow.FileOperations.State.cpp:1961-1988` validates the `FileInfo`
  buffer extent and produces a `wstring_view`, but performs no leaf-name policy validation.
- **Evidence:** sequential tree copy filters only empty/exact `.`/`..` and then calls
  `JoinFolderAndLeaf` at `State.cpp:8848-8854`; the parallel producer repeats the pattern at
  `State.cpp:9208-9214`.
- **Evidence:** S3 returns a key segment as the entry name at
  `Plugins/FileSystemS3/FileSystemS3.S3.cpp:398-411`; Curl uses the raw parsed listing remainder at
  `Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp:869-899`.
- **Impact:** a remote/archive namespace can cause a bridge copy to address a path outside the selected
  destination root, producing an arbitrary write wherever the current user has permission.
- **Fix sketch:** enforce one host-side leaf-name validator before composing either source or destination
  child paths; reject separators, traversal components, drive/stream syntax, reserved local names, and
  destination collisions according to the destination path-identity contract. Continue valid siblings and
  finish with `ERROR_PARTIAL_COPY` when continue-on-error is enabled.
- **Owner:** completed by
  `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`
  **CW-1/CW-1A**. Do not create a second validator or a second fake-provider seam here.

### LH-2 — Clamp plugin-reported `bytesRead` before any buffer traversal or write

- **Evidence:** the serial pump calls `Read(bufferIn, bufferBytesIn, &bytesRead)` at
  `State.cpp:8315-8318`, then loops to `bytesRead` at `:8329-8341` without checking
  `bytesRead <= bufferBytesIn`.
- **Evidence:** the pipeline reader stores the unchecked value at `State.cpp:8439-8449`; the writer later
  traverses that count at `:8524-8537`.
- **Impact:** a buggy or hostile plugin can turn an ABI contract violation into a heap over-read and copy
  process memory into another backend.
- **Fix sketch:** fail with `ERROR_INVALID_DATA` immediately after every reader call that reports more
  bytes than requested. The failure path must abort/delete staging output and must be clean under ASan.
- **Owner:** completed by Causeway **CW-2/CW-2A**.

### LH-3 — Make MTP picker/cancel work module-safe across refresh and unload

- **Evidence:** `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:1714-1724` stores `backend` before
  `moduleKeepAlive`; after the destructor body, reverse member destruction releases the module before the
  backend, while callback epilogue code may still execute inside the plugin.
- **Evidence:** `Common/Helpers.h:1907-1921` discards `PTP_CALLBACK_INSTANCE`, so the callback cannot use
  `FreeLibraryWhenCallbackReturns`.
- **Evidence:** `RedSalamander/ConnectionManagerWindow.cpp:923-943` calls the singleton
  `FileSystemPluginManager` from a fire-and-forget worker; `FileSystemPluginManager.cpp:1127-1132`
  unloads and clears its vector during refresh, and the manager has no synchronization protecting those
  worker reads.
- **Impact:** refresh, disconnect, watchdog, or shutdown can execute through freed plugin entries or return
  into an unmapped DLL.
- **Fix sketch:** pin exports before queueing; keep manager access on the UI thread; pass
  `PTP_CALLBACK_INSTANCE`; defer the last module release with `FreeLibraryWhenCallbackReturns`; include
  pending cancel/browse work in `CanUnloadNow`.
- **Owner:** completed by
  `Specs/Plans/Done/Operation_Firebreak_ConsolidatedReviewRemediation_2026-07-06.md` **Task 1**.

### LH-4 — Remove nested scheduler waits that can deadlock all File Operations

- **Evidence:** `PerItemTaskScheduler::WaitJob` is a bare condition-variable wait at
  `State.cpp:3987-3996` and does not help execute queued work.
- **Evidence:** `CopyDirectoryParallel` waits for a nested file-copy job at `State.cpp:9290`.
- **Evidence:** the enclosing top-level `processIndex` already runs inside a scheduler job created at
  `State.cpp:10022-10037`.
- **Impact:** enough concurrent directory producers can occupy all workers while waiting for jobs that no
  worker remains available to run. Cancel and application shutdown can remain stuck.
- **Fix sketch:** remove the nested blocking job shape or make waits participate in dequeue/work-stealing
  with cancellation and shutdown predicates.
- **Owner:** completed by Causeway **CW-3**.

### LH-5 — Eliminate unsafe executable/DLL search-path behavior

- **Evidence:** `RedLauncher/Main.cpp:332-358` computes the trusted package directory, then passes
  `nullptr` as `CreateProcessW`'s current directory, so the child inherits the launcher's working
  directory.
- **Evidence:** `RedSalamander/FileActionLauncher.cpp:557-565` accepts any non-empty expanded executable
  path; `:574-584` defaults the working directory to the current browsed context; `:609-619` passes both
  to `ShellExecuteExW`.
- **Resolved ViewerVLC evidence:** Farsight **F-S1-2** replaced ambient dependency discovery and
  process-global DLL-directory mutation. `AutoDetectVlcInstallDir` at
  `Plugins/ViewerVLC/ViewerVLC.cpp:660-710` resolves validated registry/known-folder candidates;
  `LoadVlcState` at `:5267-5284` loads the explicit `libvlc.dll` path with
  `LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS`. The installed-VLC runtime test
  verifies sibling dependency loading and playback while the process DLL directory remains unchanged.
- **Remaining impact:** untrusted or mutable working directories can still enter launcher and external
  action executable resolution. The former ViewerVLC process-global DLL-search-state risk is resolved.
- **Remaining fix sketch:** set safe process DLL search defaults at startup, launch RedSalamander with the
  package directory as `lpCurrentDirectory`, and require external actions to resolve to explicit absolute
  executables.
- **Ownership split:** the ViewerVLC leg is complete under
  `Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md`
  **F-S1-2**. Lighthouse Track A below still owns RedLauncher, early RedSalamander process hardening, and
  FileActionLauncher/Preferences validation.

### LH-6 — Make settings recovery section-scoped, non-destructive, and forward-safe

- **Evidence:** `RecoverSettingsLoadFailure` assigns `out = Settings{}` at
  `Common/Common/SettingsStore.cpp:5117`.
- **Evidence:** any schema version other than exactly 16 routes to that recovery at
  `SettingsStore.cpp:5173-5184`.
- **Evidence:** one invalid file-action, user-menu, or shortcut section routes the entire store to recovery
  at `SettingsStore.cpp:5204-5219`.
- **Evidence:** invalid inline themes are skipped at `SettingsStore.cpp:1274-1284`; the next serialization
  writes only the surviving in-memory theme list at `:5516-5524`.
- **Impact:** a local typo, a partially corrupt section, or opening settings with a different application
  version can reset unrelated windows, folders, connections, themes, and shortcuts. A backup may exist,
  but normal operation proceeds from defaults and can replace the primary file.
- **Fix sketch:** parse independent sections into temporary values and default only the invalid section;
  preserve invalid/unknown raw members for round-trip where possible; never automatically overwrite an
  unsupported future-schema file; report partial recovery explicitly.
- **Ownership split:** Firebreak **Step 4.2** owns inline-theme parse/round-trip behavior. Lighthouse Track
  B owns core section recovery, unsupported-schema save protection, and authoritative SettingsStore spec
  updates. The two owners must agree on one raw-JSON preservation representation.

### LH-7 — Persist runtime state during Windows session end without running modal teardown

- **Evidence:** normal shutdown captures settings and saves them only through `SaveAppSettings` at
  `RedSalamander/RedSalamander.cpp:2894-2919`, called from `OnMainWindowDestroy` at `:11419`.
- **Evidence:** the main WndProc switch at `RedSalamander.cpp:11511-11585` handles `WM_CLOSE` and
  `WM_DESTROY` but has no `WM_QUERYENDSESSION` or `WM_ENDSESSION` cases.
- **Impact:** logoff, restart, or OS shutdown can terminate the process without the normal destroy path,
  losing pane paths, history, view state, menu state, and window placement accumulated since the last
  explicit save.
- **Fix sketch:** return promptly from `WM_QUERYENDSESSION`; on confirmed `WM_ENDSESSION`, run an idempotent,
  non-modal snapshot/save path that reuses `CaptureRuntimeSettings` but does not prompt, close viewers, or
  unload plugins. Save settings only; leave the schema sidecar unchanged.
- **Owner:** Lighthouse Track C.

### LH-8 — Stop re-reading both endpoints in full after every MOVE

- **Evidence:** bridge cleanup opens both readers and performs a complete dual-reader compare through
  `ReadersHaveEqualContent` at `State.cpp:7521-7613`, called from
  `CopiedFileStillMatchesDestination` at `:7616-7672`.
- **Evidence:** the local cross-volume fallback repeats the pattern in
  `Plugins/FileSystem/FileSystem.FileOps.cpp:3364-3420` and `:3434-3461` before source deletion.
- **Impact:** a large local-to-cloud MOVE can upload the file, then download/read both endpoints again in
  a serial cleanup pass. This multiplies wall time and may incur cloud egress charges.
- **Fix sketch:** hash bytes during the copy, retain that digest in the MOVE manifest, and verify only the
  destination before deletion. Use one hashing scheme for bridge and native cross-volume paths; preserve
  the existing fail-safe rule that an unverifiable source is not deleted.
- **Owner:** Causeway **CW-9** and Firebreak **Step 6.2** share this requirement. Firebreak's copy-time hash
  was implemented once by Causeway on 2026-07-13. The bridge now uses copy-time FNV-1a and destination-only
  cleanup verification; the native local cross-volume path uses successful-copy plus stable-snapshot proof
  and performs zero post-copy content rereads. Firebreak retains only its separate manifest-retention subtask;
  do not reopen the bridge hash under Lighthouse.

### LH-9 — Put critical contract/crash tests and an ARM64 build into the PR gate

- **Evidence:** `.github/workflows/ci.yml:80-85` builds the self-test artifact only for x64.
- **Evidence:** `ci.yml:115-118` explicitly keeps PluginContract, SettingsSchema, and CrashHandling out of
  the PR gate.
- **Evidence:** `Tools/TestRunPlan.ps1:1871-1880` registers those suites only for `Full`.
- **Impact:** changes to plugin factories/unload contracts, settings recovery, and crash handling can merge
  without running their dedicated deterministic suites. ARM64-only compile failures are deferred to release.
- **Fix sketch:** make those deterministic suites CI-capable and add them to `Suite CI`; add an ARM64
  build-only PR job; retain Full-only classification only for genuinely intrusive or environment-sensitive
  lanes such as ETW latency.
- **Owner:** Lighthouse Track D. This supersedes the still-useful but stale current-tree assumptions in
  `plans/011-ci-coverage-gaps.md`. Do not edit or execute that old plan independently.

## Lighthouse-owned execution tracks

### Track A — Harden launcher and external-action resolution (`LH-5`, excluding ViewerVLC)

**Priority:** P0 security
**Effort:** M
**Risk:** MED
**Depends on:** none. Farsight completed the ViewerVLC dependency-discovery and isolated-load leg; the
remaining launcher/action changes are independently owned and executable here.

**Files initially in scope:**

- `RedLauncher/Main.cpp`
- `RedSalamander/RedSalamander.cpp`
- `RedSalamander/FileActionLauncher.cpp`
- `RedSalamander/FileActionLauncher.h`
- `RedSalamander/Preferences.FileActions.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
- relevant resource strings and authoritative app-shell/settings specs

**Out of scope:** `Plugins/ViewerVLC/**`, general plugin discovery/loading, shell verbs used to open user
documents, and replacement of `ShellExecuteExW` for already-absolute configured executables.

1. Add RED-first tests proving a relative/bare executable is rejected after macro expansion and an
   absolute executable remains launchable. Extend the existing `settings_store_file_actions_v16_*` family.
2. Add a small shared non-throwing validator for external-action executable paths. Enforce it in both
   settings/Preferences validation and `BuildExternalLaunchPlan`, so edited JSON cannot bypass the UI.
3. Pass `packageDirectory.c_str()` as RedLauncher's `lpCurrentDirectory`.
4. Call `SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32 | LOAD_LIBRARY_SEARCH_APPLICATION_DIR)`
   before optional/runtime loads in both launcher and main executable. Log and fail startup if the call
   fails; do not silently continue with the unsafe search path.
5. Update the authoritative app-shell and file-action settings contracts, including the migration behavior
   for existing relative actions: preserve them but mark them invalid/disabled with an actionable UI error;
   do not rewrite them to a guessed path.

**Verify:**

```powershell
./build.ps1 -Configuration Debug -Platform x64 -ProjectName RedSalamander
./Tools/Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter settings_store_file_actions_v16_
./Tools/Run-AllTests.ps1 -Suite Full -SkipBuild
```

Expected: build reports 0 errors; all file-action cases pass; Full is green or every unrelated external
blocker is recorded with a rerun/classification artifact.

**Completion — 2026-07-15: DONE**

- Added `Common::Paths::IsExplicitAbsoluteExecutablePath` for explicit drive, UNC, and Win32 extended
  executable paths while rejecting relative, drive-relative, rooted-without-drive, device, and bare names.
- Settings loading preserves legacy relative executable text but forces the action disabled; Preferences
  blocks newly saved relative actions with a localized repair message; launch-plan construction validates the
  fully expanded executable again and `LaunchExternalPlan` retains a final defense-in-depth check.
- `RedLauncher.exe` and `RedSalamander.exe` install
  `LOAD_LIBRARY_SEARCH_SYSTEM32 | LOAD_LIBRARY_SEARCH_APPLICATION_DIR` before optional/runtime loads and fail
  startup if the policy cannot be installed. RedLauncher starts the main executable with the resolved package
  directory as `lpCurrentDirectory`.
- Authoritative settings, schema, startup, and user settings documentation now records the path, migration,
  working-directory, and DLL-search contracts.
- Verification:
  - resource localization contracts: passed;
  - clean full-solution x64 Debug rebuild: 0 warnings, 0 errors;
  - `settings_store_file_actions_v16_`: 5 passed, 0 failed;
  - `file_action_external_launch_`: 2 passed, 0 failed;
  - Full run `20260715T063912Z-68848-0c7538cb385a4701ad263d91ce76ee93`: 1,138 passed,
    10 failed, 53 skipped. No loader/action/plugin-load failure occurred. The ten unrelated Commands UI,
    FolderView perf, monitor ETW, and Tools Pester blockers are preserved with classifications at
    `Specs/TestRuns/SINON/Continuation/2026-07-15_lighthouse_track_a_full/`.

### Track B — Make SettingsStore recovery lossless by section (`LH-6`, excluding Firebreak theme work)

**Status (2026-07-15): COMPLETE on `codex/lighthouse-settings-recovery`.** The owned
`JsonValue` preservation model, section-scoped File Actions/User Menu/Shortcuts recovery, future-schema
save permission, low-level/serialized save guards, explicit-replacement backup entry point, and the four
named regression cases are implemented. The authoritative spec and localized startup decision are updated;
Debug x64 builds with 0 warnings/errors, the resource-localization contract passes 5/5, the inventory
contract passes 5/5, the complete tooling Pester suite passes 262/262, and the focused `settings_store_`
Commands slice passes 16/16.

The first Full run after implementation completed in 38m56s with 1,146 passed, 6 failed, and 53
skipped. No SettingsStore/HotReload case failed. It exposed stale Commands inventory documentation and
Pester assertions (699/800 versus the then-current 704 static/805 listed cases); those counts and the Settings
family total are now reconciled. The remaining Full failures are outside Track B and are under focused
classification: one command-line focus case passed immediately on rerun; ViewerPE's outer harness,
LocalizationTests, and RedConfigureTests reproduce independently; Monitor ETW remains the known baseline
blocker. Raw Full artifacts were removed by the runner's stale-run cleanup when the focused rerun began, so
the durable closeout record explicitly preserves the observed manifest summary rather than claiming the raw
directory was archived. Closeout details are at
`Specs/TestRuns/SINON/Continuation/2026-07-15_lighthouse_track_b_full/`.

**Priority:** P1 data integrity
**Effort:** L
**Risk:** HIGH
**Depends on:** coordinate the raw JSON preservation type with Firebreak Step 4.2 before either branch
changes `ThemeDefinition` or SettingsStore JSON value ownership.

**Files initially in scope:**

- `Common/SettingsStore.h`
- `Common/Common/SettingsStore.cpp`
- `RedSalamander/SettingsHotReload.cpp`
- `RedSalamander/SettingsHotReload.h`
- `RedSalamander/RedSalamander.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/SettingsStore.schema.json` only if the persisted contract changes

**Out of scope:** rewriting the theme-file parser, changing credential persistence, changing unrelated
settings defaults, or silently migrating unsupported schema versions.

1. [x] Add RED-first cases:
   - `settings_store_invalid_file_actions_preserves_unrelated_sections`;
   - `settings_store_invalid_shortcuts_preserves_unrelated_sections`;
   - `settings_store_unsupported_future_schema_never_overwrites_source`;
   - a round-trip case proving unknown top-level members survive when the settings file is saved without
     editing them.
2. [x] Parse into a temporary `Settings` value and commit each section independently. A malformed optional
   section must reset only that section and add structured recovery evidence; it must not call the
   whole-store `RecoverSettingsLoadFailure` path.
3. [x] Preserve unknown/invalid JSON members in an opaque round-trip collection owned by the parsed settings
   document. Do not borrow yyjson pointers past document lifetime; use copied `JsonValue` data.
4. [x] For an unsupported future schema, leave the source file byte-for-byte in place, mark automatic save as
   blocked, and surface an actionable alert. Centralize the block in `SettingsHotReload` so shutdown,
   Preferences, plugin configuration, and other save call sites cannot bypass it. An explicit user-approved
   replacement may clear the block only after a backup succeeds.
5. [x] Update `Core_SettingsStore.md` with section recovery, opaque round-trip, future-schema, backup, and save
   permission rules.

Implementation note: the low-level `Common::Settings` writer also rejects blocked snapshots as defense in
depth, while `SettingsHotReload` rejects synchronous, asynchronous, and shutdown submissions before they
enter the serialized queue. Explicit replacement is a separate API whose caller must obtain user approval;
it cannot clear the block until the original file has moved to the timestamped backup path. The startup
warning defaults to preservation (`No`) and offers `Yes` only as the deliberate backup-then-replace action.

**Verify:**

```powershell
./build.ps1 -Configuration Debug -Platform x64 -ProjectName RedSalamander
./Tools/Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter settings_store_
./Tools/Run-AllTests.ps1 -Suite Full -SkipBuild
```

Expected: every new RED case passes; unsupported-schema fixture bytes remain unchanged; unrelated settings
survive each invalid-section case; Full meets the closeout rule.

**Result:** PASS for Track B-owned behavior. The four named RED cases plus the complete `settings_store_`
slice pass; future-schema automatic saves are byte-preserving and explicit replacement backs up exact bytes
before clearing the block. Full completed without a SettingsStore/HotReload failure. Its unrelated blockers
and focused classifications are recorded in the closeout evidence above.

### Track C — Save runtime state on Windows session end (`LH-7`)

**Status (2026-07-15): COMPLETE on `codex/lighthouse-session-end`.**
`CaptureAndSaveRuntimeSettingsForSessionEnd` performs one prepared settings-only
write behind an atomic save owner, `WM_QUERYENDSESSION` returns immediately, canceled `WM_ENDSESSION` is a
no-op, and later `WM_DESTROY` retains teardown while skipping a duplicate settings/schema save. The test seam
records the written snapshot and normal-teardown entry count; the Commands case changes the active pane and
menu/function-bar state, sends canceled and repeated confirmed messages, asserts one writer/no teardown, and
writes `session_end_settings_metrics.json`. `Core_StartupBootstrap.md` now owns the durable Windows session-end
and metric contract (and its stale schema-version note was reconciled from 10 to the current version 16).
Debug x64 build `msbuild-20260715_103152_565.log` is GREEN with 0 warnings and 0 errors; focused and Full
runtime validation remain pending. The first focused run reached the new case and wrote its perf artifact but
failed its active-pane assertion because the test changed the logical pane without moving live folder-view
focus. A first fixture correction then used the production restore helper, but that helper intentionally
refuses to run while either folder view already owns focus. The fixture now focuses the target folder-view
HWND directly so it deterministically drives the production `GetFocusedPane()` capture contract; the focused
rerun passed 1/1 and emitted the archived perf artifact. The current source inventory is 705 static Commands
registrations with 98 in the Settings family; the rebuilt executable authoritatively lists 808 Commands cases,
and the 5/5 inventory contract is green after reconciling those counts. The final focused-build log is
`msbuild-20260715_103907_803.log` (0 warnings/errors); focused run
`20260715T084139Z-24412-aff83ed8f2e444dcbad9d511d7247008` is archived under
`Specs/TestRuns/SINON/Continuation/2026-07-15_lighthouse_track_c_session_end/` with a 1 us query response,
1,533 us capture/injected-write duration, one writer call, zero normal-teardown entries, and `S_OK`. Full is
complete: run `20260715T084345Z-73644-5868df5e04cd4cb5b00536ef772b6cb6` took 38m50s and reported
1,147 passed, 5 failed, and 54 skipped. The new session-end case passed in the complete Commands run, as did
FileOperations, PluginContractTests, SettingsSchemaTests, CrashHandlingTests, PerformanceTests2 (14/14), and
the inventory contract. The five failures are outside Track C: Preferences reverse-keyboard navigation;
the already-known LocalizationTests and RedConfigureTests access violations; the existing Monitor ETW latency
baseline; and a Tools Pester source-contract `ArgumentNullException`. Full manifests/output were preserved
under the Track C evidence directory before any focused classification rerun. The exact Commands navigation
case then passed 1/1 in run `20260715T092433Z-65140-f5dd821388d24fa9828440da758eea18`, and the complete
`TestHarnessSourceContracts.Tests.ps1` file passed 118/118 including the previously failing payload assertion.
The remaining LocalizationTests, RedConfigureTests, and Monitor ETW failures match the independently reproduced
Track B external blockers. Classification details are archived beside the focused and Full evidence.

**Priority:** P1 reliability
**Effort:** M
**Risk:** MED
**Depends on:** Track B only if it introduces the automatic-save permission guard first; otherwise none.

**Files initially in scope:**

- `RedSalamander/RedSalamander.cpp`
- a focused selftest file under `RedSalamander/SelfTest/Commands/`
- resource/spec files only if user-visible failure policy changes
- `Specs/Core/Core_StartupBootstrap.md`

**Out of scope:** cancelling active File Operations, running normal `WM_CLOSE` prompts, unloading plugins,
closing viewers, or updating the plugin schema sidecar during session termination.

1. [x] Extract an idempotent `CaptureAndSaveRuntimeSettingsForSessionEnd(HWND)` helper that calls
   `CaptureRuntimeSettings` and `Common::Settings::SaveSettings`, logs once on failure, and performs no UI
   prompts or teardown.
2. [x] Handle `WM_QUERYENDSESSION` by returning `TRUE` promptly. Handle `WM_ENDSESSION` only when `wParam != 0`;
   guard the save with an atomic/idempotent state so repeated delivery and later `WM_DESTROY` do not race or
   write twice.
3. [x] Add a test seam around the settings writer and a Commands selftest that sends the two messages, proves
   the snapshot contains changed pane/menu state, proves cancel (`wParam == 0`) does not save, and proves no
   plugin shutdown/viewer close path ran.
4. [x] Record the session-end durability and non-modal constraints in `Core_StartupBootstrap.md`.

**Verify:**

```powershell
./build.ps1 -Configuration Debug -Platform x64 -ProjectName RedSalamander
./Tools/Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_app_session_end_
./Tools/Run-AllTests.ps1 -Suite Full -SkipBuild
```

Expected: focused cases pass; no modal prompt appears; normal `WM_CLOSE` shutdown remains green.

**Result:** PASS for Track C-owned behavior. Debug x64 builds with 0 warnings/errors; the focused session-end
case passes with one settings write, no normal teardown, and a bounded 1,533 us injected-write path. The same
case passes in Full. Full's two fresh unrelated failures passed focused reruns, while its three established
external blockers retain their existing owners and independent reproduction evidence.

### Track D — Close the PR verification gap (`LH-9`)

**Status (2026-07-15): COMPLETE on `codex/lighthouse-ci-gate`.** The test-stabilization dependency is
closed in Done. Drift review confirms Suite CI still excludes PluginContractTests, SettingsSchemaTests, and
CrashHandlingTests and the PR workflow still lacks an ARM64 lane. The reusable build workflow already accepts
ARM64, installs the ARM64 MSVC component, selects `arm64-windows`, and keys the vcpkg cache by triplet, so no
reusable-workflow change is required. On the current Debug x64 artifact, all three candidate executables passed
twice (6/6 invocations) before gate edits. Implementation now adds them to Suite CI, adds a Debug ARM64
build-only PR lane, protects both contracts in `RunAllTestsPlan.Tests.ps1`, updates the authoritative testing
specs, and retires historical `plans/011-ci-coverage-gaps.md`. The local Debug ARM64 build completed in 10m33s
with 0 errors and 80 existing C4324 alignment-padding warnings; captured log
`msbuild-20260715_113113_607.log`. The first valid CI run built x64 Debug with 0 warnings/errors and executed
all three new adapters successfully, but remained red because the existing ViewerImgRaw latest-wins harness
failed both its initial and runner retry attempts. The second complete CI run reused the same binaries and was
green: run `20260715T102731Z-69604-cc848739f552462a82eaf5775d54c352`, 1,169 passed, 0 failed, 53 skipped,
and 0 flaky/regression/isolation-suspect/unclassified classifications. The ViewerImgRaw failure is the same
independently reproduced Farsight contract regression recorded during Track B; it was not quarantined,
weakened, or assigned to Track D. Bounded evidence is archived under
`Specs/TestRuns/SINON/Continuation/2026-07-15_lighthouse_track_d_ci_gate/`.

**Priority:** P1 tests/CI
**Effort:** M
**Risk:** LOW
**Depends on:** `Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` must be handed off or
closed first. It is currently owned by another task and must not be edited concurrently.

**Files initially in scope:**

- `.github/workflows/ci.yml`
- `.github/workflows/build-reusable.yml` only if a reusable ARM64 build input/output fix is required
- `Tools/TestRunPlan.ps1`
- `Tools/Tests/RunAllTestsPlan.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_SelfTests.md`
- `Specs/Testing/Testing_TestCoverage.md`

**Out of scope:** making ETW latency a blocking PR test, running ARM64 binaries on an x64 runner, changing
test product behavior merely to make CI green, or modifying the active test-stabilization plan without
handoff.

1. [x] Prove PluginContractTests, SettingsSchemaTests, and CrashHandlingTests are deterministic on a clean x64
   CI-style build. Add explicit adapter/timeout/classification rules where needed.
2. [x] Add those three executables to `Suite CI` and update `RunAllTestsPlan.Tests.ps1` so removing one makes the
   plan test fail.
3. [x] Add an ARM64 build-only PR job through `build-reusable.yml`; do not attempt to execute ARM64 output on
   the x64 hosted runner.
4. [x] Run CI twice from clean artifacts. Any pass-on-rerun classification remains blocking under the existing
   flaky-test contract.
5. [x] Reconcile and retire the remaining current-tree intent in `plans/011-ci-coverage-gaps.md`; do not maintain
   two CI plans.

**Verify:**

```powershell
./Tools/Run-AllTests.ps1 -Suite CI
./Tools/Run-AllTests.ps1 -Suite CI -SkipBuild
./build.ps1 -Configuration Debug -Platform ARM64
```

Expected: both CI runs pass with all three added suites present in `run-all-tests-results.json`; ARM64 build
exits 0; workflow plan tests assert the new lanes.

**Result:** PASS for Track D-owned behavior. PluginContractTests, SettingsSchemaTests, and CrashHandlingTests
passed twice before integration and in both complete CI compositions; the second composition was wholly green.
The ARM64 build-only lane is source-guarded and locally cross-built. The first composition's sole ViewerImgRaw
regression remains a blocking external Farsight test-contract finding rather than being hidden to make CI green.

## Simplification backlog

These are architecture/DX improvements, not ranked above the correctness and security work.

### LH-S1 — Unify plugin unload state machines

`FileSystemPluginManager` and `ViewerPluginManager` independently implement deferred unload, quiet-point
probing, resource-owner registration, and process-shutdown behavior and have drifted. Firebreak Step 1.3
owns a shared lifecycle helper. Execute it only after LH-3's immediate MTP lifetime fixes are green.

### LH-S2 — Make SettingsStore section-oriented

Track B should leave parsing, validation, recovery, and serialization organized around independently
testable sections instead of one 6,000+ line monolith with duplicated reader/writer contracts. Do not split
files mechanically before the lossless behavior is characterized.

### LH-S3 — Decompose included-`.cpp` mega translation units

`RedSalamander/Commands.SelfTest.cpp:469-481` includes thirteen large implementation files;
`FolderWindow.FileOperations.State.cpp:10808-10810` includes three `.Part.cpp` files. Create a design spike
after the P0/P1 work to introduce internal headers and real compilation units without exposing test-only
symbols to production Release builds. The current warmed scoped build still took 3:26, so measure build-time
and peak-memory improvement before accepting the refactor.

### LH-S4 — Finish File Operations popup cleanup

`UI_FileOperationsPopupReviewFindings_2026-07-09.md` S1–S5 owns the remaining single-pass summary reducer,
cached resource/palette state, shared easing, and settings-save helper cleanup. Keep it opportunistic and do
not mix it into Causeway's bridge safety changes.

## 2026-07-14 duplication addendum

This addendum records consolidation opportunities that survived semantic source review. A matching helper
name is not enough to justify extraction: each item below names the policy that must become explicit and the
variants that must remain distinct. Generated files, third-party code, and PoCs are excluded. All rows are
`LIGHTHOUSE OWNER`; execute them through Tracks E-I rather than as opportunistic edits in unrelated work.

| ID | Duplicated policy or mechanism | Priority | Effort | Risk | Confidence |
|----|--------------------------------|----------|--------|------|------------|
| LH-D1 | AppTheme resolution and palette projection | P2 | M | HIGH | HIGH |
| LH-D2 | ARGB/COLORREF/D2D color conversion and blending | P2 | M | MED | HIGH |
| LH-D3 | Stable visual hashing, HSV conversion, luminance, and contrast | P2 | M | MED | HIGH |
| LH-D4 | DPI, modifier-key, and popup geometry primitives | P3 | M | MED | HIGH |
| LH-D5 | HWND Direct2D render-target resource lifecycle | P3 | L | HIGH | MED |
| LH-D6 | Plugin-configuration schema parsing and serialization | P2 | M | MED | HIGH |
| LH-D7 | Strict yyjson accessors and document ownership | P2 | M | MED | HIGH |
| LH-D8 | SettingsStore grid-layout parsing | P3 | S | LOW | HIGH |
| LH-D9 | UTF conversion, ordinal string, trim, and environment helpers | P2 | M | MED | HIGH |
| LH-D10 | Windows path classification, absolute-path construction, and unique sibling names | P1 | M | HIGH | HIGH |
| LH-D11 | Cloud protocol encoding and local handle-I/O primitives | P1 | M | MED | HIGH |
| LH-D12 | Plugin callback, COM-worker, and pinned-threadpool lifetime envelopes | P1 | L | HIGH | HIGH |
| LH-D13 | Packed `IFilesInformation` storage and bounds-safe construction | P1 | L | HIGH | HIGH |
| LH-D14 | Viewer chrome and file-combo scaffolding | P2 | M | MED | HIGH |
| LH-D15 | Unicode clipboard ownership and retry behavior | P2 | S-M | MED | HIGH |
| LH-D16 | Contiguous posted-payload coalescing | P2 | M | MED | HIGH |
| LH-D17 | Modal DxUi shell and owner-window centering | P3 | M | MED | MED |
| LH-D18 | Test sandbox, message-pump polling, snapshots, and child-process execution | P2 | M | MED | HIGH |
| LH-D19 | RedConfigure binary-file reader | P3 | S | LOW | HIGH |
| LH-D20 | Throughput and byte-size parsing | P3 | S | MED | HIGH |
| LH-D21 | File metadata display formatting | P3 | S | LOW-MED | HIGH |
| LH-D22 | Canonical filesystem-plugin lookup | P3 | S | LOW | HIGH |

### LH-D1 — Canonicalize AppTheme resolution and palette projection

- **Evidence:** full theme override/resolution paths exist in `RedSalamander/RedSalamander.cpp` near
  `ApplyThemeOverrides` and configured-theme resolution, in `RedSalamander/SettingsHotReload.cpp`, and in
  `RedSalamander/Preferences.Themes.cpp`. Palette projection already has a partial canonical entry point in
  `RedSalamander/DxUiThemePalette.h`, while `ConnectionCredentialPromptDialog.cpp`, `FindFilesWindow.cpp`,
  and `FolderWindow.FileOperations.IssuesPane.cpp` rebuild subsets locally.
- **Required simplification:** introduce one side-effect-free resolved-theme pipeline and named palette
  projections for the supported control families. Startup, hot reload, and preview must consume the same
  resolved model; storage, file watching, and UI invalidation remain at their current boundaries.
- **Parity gate:** characterize missing keys, invalid expressions, inheritance/fallback order, high contrast,
  preview-only edits, and hot-reload failure. Snapshot resolved colors before migrating callers.

### LH-D2 — Share representation-explicit color conversion and blending

- **Evidence:** `Common/DxUi/DxUi.Theme.cpp` already defines `ColorFromArgb` and its inverse, but equivalent
  conversion code remains in `RedConfigure/RedConfigureUiHelpers.cpp`,
  `RedSalamanderMonitor/RedSalamanderMonitor.cpp`, and ViewerSpace. `ColorRefFromArgb` and bit-identical
  `BlendColor(COLORREF, COLORREF, ...)` implementations recur in ViewerText, ViewerImgRaw, ViewerVLC,
  ViewerWeb, and ViewerSpace; `RedSalamander.cpp` also packs ARGB locally.
- **Required simplification:** provide dependency-appropriate helpers whose names state both source and
  destination representations and whose blend parameter documents endpoint/rounding semantics. Migrate only
  bit-identical implementations; keep adapters at dependency boundaries when a plugin cannot depend on DxUi.
- **Parity gate:** exhaustive channel-edge tests (`0`, `1`, `254`, `255`), alpha round-trips, blend endpoints,
  midpoint rounding, and COLORREF BGR ordering.

### LH-D3 — Version visual hashes and share color mathematics without merging policies

- **Evidence:** `StableHash32`/`ColorFromHSV` pairs recur in `Common/DxUi/DxUi.Theme.cpp` and ViewerPE,
  ViewerVLC, ViewerText, ViewerImgRaw, and ViewerWeb. Luminance calculations recur in
  `Common/Common/ThemeExpression.cpp`, `RedConfigureWorkflow.cpp`, and ViewerText Hex. Contrast decisions
  also recur, but with intentionally different thresholds.
- **Required simplification:** centralize the bit-identical byte hash, HSV conversion, and luminance math;
  expose contrast as named policies rather than a universal threshold. Preserve persisted/visible output.
- **Intentional variant:** `RedSalamander/AppTheme.cpp` hashes both bytes of each UTF-16 code unit and is not
  bit-compatible with the plugin/DxUi hash. Keep it separate or give both algorithms explicit versioned names.
- **Parity gate:** golden hashes for ASCII and non-ASCII labels, hue-wrap/boundary cases, exact generated
  colors, and luminance/contrast threshold boundaries.

### LH-D4 — Consolidate pure DPI, input-modifier, and popup geometry helpers

- **Evidence:** `Common/WindowSizing.h` provides shared DPI primitives, while `RedSalamander/UiMetrics.cpp`,
  both splash screens, and several viewers retain local conversions. Modifier extraction is duplicated in
  `DxUi.WindowHost.cpp`, `DxUi.TextInput.cpp`, `DxUi.Grid.cpp`, and `DxUi.ComboBox.cpp`; popup-height and work-
  area arithmetic also recur across viewer file combos.
- **Required simplification:** extend the existing low-level geometry surface with explicitly rounded
  `DipToPixel`/`PixelToDip`, key-modifier decoding, and bounded-popup helpers. Keep control policy outside the
  primitives.
- **Parity gate:** DPI values 96/120/144/192, negative coordinates, overflow-safe multiplication, fractional
  rounding, mixed-monitor work areas, and every modifier combination.

### LH-D5 — Reuse a narrow HWND Direct2D render-resource host

- **Evidence:** `RedSalamander/FunctionBar.cpp` and `FolderWindow.StatusBar.cpp` independently manage nearly
  identical factory/device-context/render-target creation, resize, device-loss, and resource-discard paths.
- **Required simplification:** after rendering behavior is measured, extract only the shared HWND target and
  device-loss lifecycle; drawing, text layouts, brushes tied to component policy, and invalidation stay local.
- **Boundary:** do not create an all-viewer rendering framework. FunctionBar and StatusBar are the first and
  only required adopters; expand later only with separately reviewed evidence.
- **Parity/perf gate:** resize/DPI/device-loss stress, no retained stale COM resources, unchanged pixel output,
  and archived render responsiveness evidence when the hot path changes.

### LH-D6 — Use one plugin-configuration schema model and codec

- **Evidence:** `RedSalamander/ManagePluginsDialog.cpp` and
  `RedSalamander/Preferences.Plugin.Configuration.cpp` independently parse, validate, display, mutate, and
  serialize the same plugin configuration schema.
- **Required simplification:** create a non-UI schema/model/codec with lossless unknown-field handling and
  explicit validation results. Both dialogs keep their own layout and interaction state but use one data
  contract and serializer.
- **Parity gate:** golden parse/serialize cases for every field type, invalid constraints, unknown/future
  fields, duplicate IDs, ordering, cancel/no-save, and semantically identical output from both entry points.

### LH-D7 — Establish strict yyjson accessors and RAII ownership

- **Evidence:** local `TryGet...`/typed-member accessors and document cleanup recur in
  `Common/FileSystemPathIdentity.cpp`, GoogleDrive, 7z, Preferences plugin configuration, Curl, S3,
  MicrosoftDrive, and tests.
- **Required simplification:** add dependency-light const accessors with explicit required/optional and strict/
  coercing policies, plus the existing approved RAII ownership shape. Migrate exact policy matches only.
- **Boundary:** provider protocols may intentionally accept a numeric string or tolerate a missing member;
  those choices remain visible at the call site. Mutable builders must continue using copying APIs for dynamic
  strings and keys.
- **Parity gate:** missing, null, wrong type, overflow, malformed UTF-8, numeric-string, and unknown-member
  cases, with leak checks on every early return.

### LH-D8 — Parameterize SettingsStore's repeated grid-layout parser

- **Evidence:** `Common/Common/SettingsStore.cpp` repeats the same grid column/order/width/visibility parsing
  block for multiple settings sections.
- **Required simplification:** replace the repeated blocks with one section parser driven by field descriptors
  or a narrow callback. Preserve section-specific defaults, validation messages, migration versions, and
  serialization order.
- **Parity gate:** existing SettingsSchema golden inputs plus missing/duplicate/invalid columns, unknown future
  fields, and section-scoped recovery.

### LH-D9 — Separate strict UTF conversion from replacement conversion and reuse string policies

- **Evidence:** the public helpers in `Common/Helpers.h` use replacement-oriented UTF conversion, while strict
  local variants recur in SettingsStore, ThemeDefinitionIo, Curl, VLC, and other plugins. Case-insensitive
  equality, trim, and environment reads also have many wrappers; `NavigationLocation.h` additionally uses a
  locale-sensitive `towlower` path unlike the ordinal helpers.
- **Required simplification:** expose named strict and replacement conversion APIs, retain ordinal Windows
  comparison as the canonical identifier policy, reuse shared trim, and add an environment reader whose
  empty/missing/truthy semantics are explicit.
- **Parity gate:** invalid and truncated UTF, embedded NUL, supplementary characters, Turkish-I/ordinal cases,
  whitespace sets, missing versus empty environment variables, and no silent replacement in strict parsers.

### LH-D10 — Make Windows path and unique-name policies explicit

- **Evidence:** drive/UNC/extended/absolute classifiers recur in `NavigationLocation.h`,
  `FolderView.Enumeration.cpp`, `NavigationViewInternal.h`, and `FolderWindow.FileSystem.cpp`; absolute-path
  builders recur in LocalSearchIndexCore and the FileSystem plugin. Separator/dot-entry predicates recur in
  rename and path code. `Common/Helpers.h` already has `BuildUniqueSiblingName`, while LocalSearchIndexCore,
  ViewerWeb, MTP, and File Operations still generate sibling names independently.
- **Required simplification:** define separate `DriveQualified`, `DriveAbsolute`, UNC, extended, device, and
  fully-absolute classifications; centralize normalization-free absolute construction and simple leaf
  predicates. Add an atomic `CreateUniqueSiblingFile` operation that preserves `CREATE_NEW` retry semantics.
- **Semantic warning:** current callers disagree about whether `C:` is “drive-like” or absolute. Do not replace
  those predicates until the caller's intended policy is named and tested. Never weaken traversal, ADS,
  reserved-name, or cross-filesystem leaf validation.
- **Parity/security gate:** drive-relative/rooted, UNC, extended/device, slash variants, dot segments, trailing
  spaces/dots, reserved names, long paths, symlink/reparse parents, collision races, and cleanup after failure.

### LH-D11 — Share cloud protocol encoding and total-write handle primitives

- **Evidence:** percent encoding recurs in Curl, MicrosoftDrive, and GoogleDrive with a meaningful
  preserve-slash variant. Total-write loops recur in FileActionLauncher, SettingsStore, ViewerWeb, and test
  support. The SettingsStore loop does not reject a successful zero-byte write and can spin indefinitely;
  file-size/rewind/read loops also recur.
- **Required simplification:** provide a byte-oriented percent encoder with an explicit `preserveSlash` policy,
  and handle-I/O helpers for bounded size, rewind, exact read, and total write. Treat zero forward progress as
  failure and retain cancellation/maximum-size checks at domain boundaries.
- **Parity gate:** RFC3986 byte cases, UTF-8 input, slash-preserving and slash-encoding golden URLs, partial
  writes, zero-progress writes, disk-full/error propagation, large-file bounds, and file-position restoration.

### LH-D12 — Consolidate narrow plugin callback and worker lifetime envelopes

- **Evidence:** `Common/Helpers.h` already contains `RegistrationCallbackState` and an owned-threadpool helper,
  while GoogleDrive, MicrosoftDrive, MTP, ViewerWeb, ViewerSqlite, ViewerPE, ViewerImgRaw, and other plugins
  retain local callback counters, module pins, callback submission/transfer, and MTA initialization patterns.
- **Required simplification:** reuse one callback registration/in-flight/drain primitive; add only a narrow
  module-pin envelope that transfers ownership with `FreeLibraryWhenCallbackReturns` at callback entry; and
  reuse a fail-fast MTA worker-entry helper where the COM policy is identical. Keep producer stop, cancellation,
  provider sessions, progress counters, watchdogs, and result delivery domain-specific.
- **Boundary:** do not revive the rejected generic `std::function` threadpool scheduler from Farsight plan 030.
  Do not merge STA/OLE UI initialization with MTA worker initialization. Plugin quiet-point order remains:
  stop producers, cancel, stop posting, clear callback, release instances, then unload.
- **Parity/stress gate:** callback replacement and clearing, concurrent in-flight drain, submission failure,
  cancellation, unload during work, callback-return module-pin transfer, COM initialization failure, and focused
  teardown stress under every migrated plugin.

### LH-D13 — Build packed `IFilesInformation` results through one bounds-safe owner

- **Evidence:** MTP, S3, GoogleDrive, MicrosoftDrive, Curl, and 7z independently allocate and lay out packed
  `IFilesInformation`/`FileInfo` blocks, calculate aligned offsets, copy names, and expose enumeration results.
- **Required simplification:** introduce an ABI-neutral owning builder that performs checked size arithmetic,
  alignment, name storage, and final immutable interface exposure. Provider-specific metadata extraction and
  enumeration stay local.
- **Boundary:** do not change the plugin ABI or reinterpret provider-reported lengths before validation. Builder
  failure must leave no partially published interface and must preserve the current allocator/deallocator pair.
- **Parity/security gate:** empty and maximum-size sets, long/non-BMP names, alignment on x64/ARM64, arithmetic
  overflow, malformed provider lengths, allocation failure, iteration order, and consumer compatibility in
  PluginContractTests.

### LH-D14 — Share viewer chrome and pure file-combo behavior

- **Evidence:** ViewerWeb, ViewerPE, ViewerVLC, ViewerText, and ViewerImgRaw repeat DWM title-bar application;
  four or five of those viewers also repeat `MessageMayOpenWindowComboPopup`, popup-height calculation,
  file-combo refresh, and related input geometry.
- **Required simplification:** extract a small viewer-chrome helper and pure popup/input functions. Keep each
  viewer's `WndProc`, control ownership, model population, activation policy, and error handling local.
- **Boundary:** ViewerSpace has materially different chrome and is excluded until separately adapted. Do not
  create a shared base window class or cross-plugin mutable singleton.
- **Parity gate:** dark/light/high-contrast title bars, older DWM attribute fallback, Alt/F4/mouse activation,
  mixed-DPI work areas, empty and long file lists, focus restoration, and no popup-open regression.

### LH-D15 — Provide one dependency-light Unicode clipboard helper

- **Evidence:** clipboard-open/allocation/copy/ownership-transfer sequences recur in
  `Common/DxUi/DxUi.WindowHost.cpp`, `FolderWindow.FileSystem.cpp`, ViewerText Text/Hex, and ViewerWeb.
- **Required simplification:** expose a Common helper with explicit owner HWND, Unicode-only allocation,
  retry policy, empty-text behavior, and ownership transfer only after successful `SetClipboardData`.
- **Parity gate:** busy clipboard retries, empty text, embedded NUL policy, large text, allocation and API
  failures, owner destruction, and proof that the global-memory owner is released exactly once.

### LH-D16 — Reuse typed, contiguous queue-head payload coalescing

- **Evidence:** Compare Directories independently coalesces scan and content progress payloads; Find Files has
  a similar queue-head helper and separate result/progress drains.
- **Required simplification:** add a typed helper that consumes the current payload with
  `TakeMessagePayload<T>`, then drains only immediately contiguous messages matching HWND/message/key. The
  reducer is supplied by the caller and must preserve cancellation and completion ordering.
- **Boundary:** never search/reorder the full queue, consume an unrelated message, or bypass
  `PostMessagePayload` registry/drain rules. Capture message metadata before moving any payload.
- **Parity/stress gate:** interleaved message types, multiple operation IDs, completion behind progress,
  destroy-with-queued-payload, cancellation, producer flood, and no lost final snapshot.

### LH-D17 — Factor the repeated modal DxUi shell and owner centering

- **Evidence:** About and Fatal Error dialogs in `RedSalamander.cpp` duplicate modal host creation, owner
  disable/reenable, nested loop, close/result handling, and WndProc plumbing. Owner-centering arithmetic also
  recurs in credential prompts, folder UI, app dialogs, and both splash implementations.
- **Required simplification:** create a narrow modal-shell utility with explicit result, exception boundary,
  owner lifetime, and teardown rules, plus a pure monitor-aware centering helper. Dialog content and commands
  remain separate.
- **Boundary:** do not use the modal shell for session-end handling, worker callbacks, or any path that requires
  non-modal teardown. Preserve accessibility focus and disabled-owner restoration on every exit.
- **Parity gate:** owner destroyed mid-dialog, nested dialog, Escape/close/result, DPI/monitor changes,
  off-screen owners, taskbar work areas, initialization failure, and fatal-dialog fallback.

### LH-D18 — Consolidate first-party test infrastructure

- **Evidence:** sandbox/environment setup is repeated in ViewerPE, ViewerSqlite, RedConfigure,
  CrashHandling, DxUi, and Performance tests. Message pumping and asynchronous snapshot polling recur across
  viewer and Commands tests. Child-process launch, timeout, pipe capture, and cleanup recur in ViewerPE,
  Commands Settings, RedConfigure, and Compare selftests.
- **Required simplification:** add a test-support surface in this order: sandbox/environment scope; bounded
  message-pump/snapshot polling; then a child-process runner with kill-on-close job ownership and concurrent,
  bounded stdout/stderr drains. Migrate one harness at a time without changing its oracle.
- **Correctness finding:** the Compare selftest waits for process completion before reading a captured pipe;
  a sufficiently verbose child can fill the pipe and deadlock. The shared runner must drain concurrently.
- **Parity gate:** timeout, cancellation, process-tree kill, large stdout/stderr, inherited-handle closure,
  quoted Unicode arguments, nonzero exit, UI message starvation, sandbox cleanup, and parallel-test isolation.

### LH-D19 — Reuse RedConfigure's binary-file reader

- **Evidence:** `RedConfigure/Workspace/WorkspaceDiscovery.cpp` and
  `RedConfigure/RedConfigureSession.cpp` contain equivalent open/size/read/bounds/error handling.
- **Required simplification:** create one RedConfigure-local bounded binary-reader returning the existing error
  form. Preserve sharing flags, maximum size, empty-file behavior, and logging ownership.
- **Parity gate:** empty/small/maximum/oversize files, short read, locked/disappearing file, and Unicode path.

### LH-D20 — Consolidate throughput and byte-size parsers behind explicit options

- **Evidence:** File Operations popup and Preferences duplicate the same wide fractional throughput grammar,
  unit aliases, rounding, and saturation, but the final census found one boundary-policy difference: the popup
  trims the six ASCII whitespace characters while Preferences trims every C0/control value through U+0020.
  Dummy and SettingsStore contain related integer/narrow byte-size parsers with different unit/default
  semantics.
- **Required simplification:** share the wide parser engine with a named boundary-whitespace policy so both
  existing callers retain their exact behavior. A broader parser is allowed only with options that state
  fractional input, binary/decimal units, default unit, whitespace, suffix case, zero, and bounds.
- **Parity gate:** golden current inputs, fractions, overflow, unit aliases, missing suffix, whitespace, locale-
  independent decimal parsing, and round-trip with the corresponding formatter.

### LH-D21 — Share file metadata display formatting

- **Evidence:** size/date/attribute display formatting recurs in `FolderViewInternal.h`, Compare Directories,
  StatusBar, and performance-related views.
- **Required simplification:** provide pure formatting functions over normalized metadata and a named display
  profile. Keep localization/resource selection and context-specific omission policy explicit.
- **Parity gate:** unknown and zero sizes, directories, sparse/compressed attributes, pre-epoch/invalid times,
  locale/time-zone behavior, and exact existing resource output.

### LH-D22 — Expose canonical filesystem-plugin lookup

- **Evidence:** `FileSystemPluginManager.cpp` has the authoritative manager lookup, while Compare, Find Files,
  and selftests reproduce identifier matching and fallback traversal.
- **Required simplification:** expose a const, case-insensitive lookup/query API from the manager and remove
  external scans. Preserve plugin ordering, disabled/invalid handling, and lifetime/lock guarantees.
- **Parity gate:** casing variants, duplicate-case-invalid IDs, disabled and unloaded plugins, missing ID,
  concurrent refresh/unload, and no returned pointer outliving its manager lease.

## Duplication remediation tracks

Each track starts with characterization tests against the current implementations, introduces the narrowest
shared API, migrates one caller, and only then removes the remaining exact copies. After each caller migration,
run its focused tests before continuing. A reduced line count is useful evidence, but is never sufficient proof
that policies are equivalent.

### Track E — Common pure primitives (`LH-D2`, `LH-D3`, `LH-D4`, `LH-D7`, `LH-D9`–`LH-D11`, `LH-D20`, `LH-D21`)

**Priority:** P1-P3 foundation
**Effort:** L across multiple commits
**Risk:** MED-HIGH
**Dependencies:** none; complete policy characterization before defining APIs.

**Status (2026-07-15): COMPLETE on `codex/lighthouse-common-primitives`.** Tracks A-D are committed;
Track E was delivered as independently characterized sub-batches so the shared API does not
silently merge incompatible policies:

- **E1 visual/input primitives — complete:** `Common::Colors`, `DxUi`, and `Common::WindowSizing` now own the
  characterized ARGB-to-D2D/COLORREF conversions, truncating and weighted COLORREF blends, clamped D2D blend,
  versioned UTF-16-code-unit visual hash, both negative-hue HSV policies, linearized and encoded-channel
  luminance/contrast math, integer/fractional DPI conversion, bounded popup geometry, work-area clamping, and
  modifier composition/queries. Exact-equivalent callers in ViewerPE/VLC/ImgRaw/Text/Web/Space,
  RedConfigure, Monitor, both splash screens, UiMetrics, FunctionBar, StatusBar, and DxUi menus now delegate to
  those primitives; compatibility member/adaptor names remain only where their owning class API needs them.
  Golden tests cover channel and blend endpoints, ASCII/non-ASCII/astral hashes, distinct negative-hue results,
  DPI 96/120/144/192, negative coordinates, fractional and saturating conversion, bounded/oversized popup
  placement, and all 256 modifier combinations. The final clean Debug x64 full rebuild completed in 2:42 with
  0 warnings/0 errors (`.build/logs/msbuild-20260715_135414_602.log`), followed by a passing
  `DxUiTests.exe --suite=Theme`. AppTheme's byte-wise UTF-16 hash and D2D HSV surface,
  SettingsSchemaExport's UTF-8 byte hash, and product-specific contrast thresholds remain explicit tested
  policy variants rather than false consolidation targets.
- **E2 JSON/UTF/string primitives — complete:** added separately named strict/optional,
  strict-or-empty compatibility, and replacement UTF-8/UTF-16 conversions; dependency-light yyjson immutable/
  mutable ownership; typed member results for required/optional missing, null, wrong type, range, invalid UTF-8,
  numeric-string coercion, integer-to-bool coercion, and unsigned-storage policy; canonical ordinal comparison,
  `iswspace` trimming, and an environment reader that distinguishes missing from explicitly empty values.
  Goldens cover malformed/truncated UTF, embedded NUL, supplementary characters, unpaired surrogates, yyjson
  presence/type/range/coercion states, Turkish-I ordinal behavior, whitespace boundaries, and missing/empty/
  truthy environment values. Exact converter engines now delegate in core settings/theme/schema, app and monitor
  console/file paths, File Operations, RedConfigure, ViewerPE/VLC/Web/Sqlite, and FileSystem, 7z, Curl, Dummy,
  GoogleDrive, MicrosoftDrive, MTP, and S3. Exact typed accessors and the approved yyjson ownership shape are
  shared across `FileSystemPathIdentity` and the characterized provider/test boundaries; MicrosoftDrive's
  legacy numeric-string parser, plugin-configuration real-number coercion, ViewerImgRaw/7z ACP fallback,
  ViewerWeb's chunked encoder, and navigation's byte-wise malformed-escape fallback remain explicit policy
  variants. The current-head full Debug x64 rebuild completed in 2:44 with 0 warnings/errors
  (`msbuild-20260715_144326_398.log`), followed by passing Theme, SettingsSchema, FileSystemCurlTests, and
  complete PluginContractTests runs. The final definition census leaves only thin compatibility delegates and
  the named variants above; there is no second strict/replacement UTF engine in production. The independently
  reproduced ViewerImgRaw latest-wins failure remains the existing Track B/D blocker and is not an E2
  regression.
- **E3 path/I/O/protocol primitives — complete:** E2 is committed as `a7f116e8`; E3 is committed as
  `dab70c70`. Added explicit relative,
  rooted, drive-relative, drive-absolute, UNC, extended-drive, extended-UNC, other-extended, and device path
  classes plus a stricter complete-drive/UNC `IsFullyAbsoluteWindowsPath` policy. Navigation callers retain
  their characterized drive-qualified versus drive-absolute behavior through explicit class sets. Shared
  normalization-free extended-path conversion and atomic `CREATE_NEW` unique-file/sibling ownership now back
  ViewerWeb, ViewerImgRaw, ViewerText, ViewerSqlite, LocalSearch snapshot persistence, and local filesystem
  overwrite writers. Remote MTP overwrite leaves now reuse the bounded 128-bit sibling-name primitive. Added
  RFC3986 byte encoding with explicit encode/preserve-slash
  policy and migrated GoogleDrive and MicrosoftDrive; S3's AWS-SDK encoder remains an intentional SDK-owned
  variant pending its own golden. Added bounded size, rewind, exact-read, and total-write handle operations;
  SettingsStore's successful-zero-byte infinite loop is removed, and exact callers in FileActionLauncher,
  ViewerWeb/Text/PE, Curl, MicrosoftDrive, S3, and MTP now delegate. Goldens cover path class/slash/namespace
  boundaries, malformed UTF, percent bytes, simulated partial/zero/error transfer, actual bounded rewind/read,
  two simultaneous atomic sibling reservations, and failure output cleanup. The current-head full Debug x64
  build completed in 2:14 with 0 warnings/errors (`.build/logs/msbuild-20260715_151331_339.log`), followed by
  passing Theme, ViewerSqliteTests, complete PluginContractTests/provider debug selftests,
  `search_low_hardening_smoke`, and the complete `mtp_` CompareDirectories family. The closeout census leaves
  only thin domain-name delegates, operational non-overwrite `CREATE_NEW` calls, ConnectionManager identity
  GUIDs, test-fixture GUIDs, and S3's explicitly SDK-owned `Aws::Utils::StringUtils::URLEncode` policy.
- **E4 parsing/metadata and closeout census — complete:** the File Operations popup and Preferences
  wide-string throughput parser bodies share locale-independent `.`/`,` decimals, optional `/s`, binary aliases
  from bytes through PiB, bare values as KiB, nearest-byte rounding, and saturation at `uint64_t`. The census
  found and routed one previously undocumented boundary difference: the popup trims the six ASCII whitespace
  characters, while Preferences trims every code unit through U+0020. The shared engine will name and preserve
  that policy. The Dummy plugin's narrow integer parser and SettingsStore's narrow byte-size parser remain
  intentionally separate input policies. The shared parser/formatter and both named boundary policies are now
  implemented with goldens for empty/unlimited, aliases, fractions, rounding, overflow, invalid syntax, the
  control-character boundary difference, and formatter round-trip. The four exact local-time/compact-attribute
  implementations in Folder View, its performance-access mirror, Compare Directories, and the status bar now
  use one normalized-metadata `CompactDetails` profile. Tests cover zero/negative/invalid and pre-Unix times,
  Windows local-time conversion with locale-independent fixed output, exact `RHSACETOP` order, dash fallback,
  and the intentionally omitted directory/sparse flags. Find Files, Item Properties, File Operations conflict
  metadata, and provider property documents retain their different seconds, size, attribute, output-encoding,
  or omission policies. The durable grammar/profile contracts are recorded in the File Operations popup,
  Preferences, Folder View, and Compare Directories authoritative specs. A clean Debug x64 full rebuild passed
  in 2:48 with 0 warnings/errors (`.build/logs/msbuild-20260715_152635_507.log`), and the focused Theme suite
  passed. The focused definition census leaves the Dummy narrow integer throughput parser and SettingsStore
  narrow byte-size parser; AppTheme/SettingsSchemaExport versioned hash/HSV policies; Find Files, Item
  Properties, File Operations conflict, and provider-property metadata profiles; S3's SDK URL encoder; thin
  domain conversion adapters; and test-only fixture/mirror helpers as the explicit allowlist. None is an
  unclassified second implementation of the newly shared contracts.

  The required Full run `20260715T133041Z-77752-85058bb7ce164687a5bbc7bf211e4d8c` completed in 38:10
  with 1,146 passed, 5 failed, and 55 skipped. The only Track E-owned failure was the new Theme atomic-sibling
  test's raw `GetTempPathW` use; it now uses the unified DxUi TestSandbox artifact helper, and the owning
  118-case `TestHarnessSourceContracts.Tests.ps1` file passes together with Theme after a 0-warning rebuild
  (`.build/logs/msbuild-20260715_161429_734.log`). Both Commands UI-focus failures, ViewerSpace's capped-scan
  aggregate race, and the monitor ETW latency harness passed their exact isolated reruns and touch none of the
  E4 parser/metadata paths. The complete Pester aggregate had 261/262 passing before the E4-owned
  source-contract repair; the repaired failing file is green. Track E is therefore PASS with all owned
  behavior green and the four composition-only failures classified by focused reproduction. Track E is
  committed as `5bfeea7b` (`Consolidate throughput and metadata formatting`).

No shared primitive has been declared complete from name similarity alone. Each sub-batch must first capture
current edge behavior in focused tests, migrate exact-equivalent callers, and rerun the owning projects.

**Initial scope:** `Common/Helpers.h`, `Common/Common`, `Common/DxUi`, focused path/string/JSON users in
`Plugins/` and `RedSalamander/`, and their existing test projects. Do not add a UI or plugin-manager dependency
to Common.

1. Add golden tests around every retained implementation and record incompatible variants in the test names.
2. Land color conversion/blend, versioned visual hash/HSV/luminance, and DPI/input/geometry primitives.
3. Land strict yyjson accessors and strict/replacement UTF conversion as separately named APIs; migrate strict
   parsers before convenience/UI callers.
4. Land explicit Windows path classifications and atomic unique-sibling creation. Migrate security-sensitive
   path callers only after their current policy is named.
5. Land percent encoding and total-write/read helpers; fix the SettingsStore zero-progress loop through the
   shared primitive and retain bounded-size checks.
6. Share only the exact throughput parser first, then metadata display helpers. Broader option-driven parsing
   can follow when all existing input policies have golden coverage.
7. Remove private helpers only when the final census shows no non-allowlisted definition and the owning tests
   are green.

**Focused definition gate:**

```powershell
rg -n "ColorF(rom)?Argb|ColorRefFromArgb|BlendColor|StableHash32|ColorFromHSV|Utf16FromUtf8|Utf8FromUtf16|TryGet(Json|String|Int)|LooksLikeWindowsAbsolutePath|BuildUniqueSiblingName|WriteAll(Bytes)?|Parse.*(Throughput|ByteSize)" Common Plugins RedConfigure RedSalamander RedSalamanderMonitor Tests
```

**Verify:** build affected projects, run their focused DxUi/SettingsSchema/plugin tests, then:

```powershell
./build.ps1 -Configuration Debug -Platform x64
./Tools/Run-AllTests.ps1 -Suite Full -SkipBuild
```

Expected: golden outputs are unchanged except for the explicitly fixed zero-progress failure; every remaining
local definition is either a tested intentional policy variant or an allowlisted dependency adapter.

### Track F — Theme and viewer/UI consolidation (`LH-D1`, `LH-D5`, `LH-D14`, `LH-D15`, `LH-D17`)

**Priority:** P2-P3
**Effort:** L
**Risk:** HIGH
**Depends on:** Track E color, geometry, and string primitives.

**Status (2026-07-15): COMPLETE on `codex/lighthouse-theme-viewer-dedup`.** The opening definition census
confirmed four exact standalone file-combo popup-input predicates in ViewerWeb, ViewerImgRaw, ViewerText, and
ViewerPE alongside the existing `Common/ViewerFileComboHost.h` keyboard host. It also found three exact-looking
viewer Unicode-text clipboard writers (ViewerWeb plus the ViewerText text/hex surfaces), while DxUi and the main
application have related retrying clipboard code whose retry/logging contracts must be characterized before
reuse. Owner-centering implementations remain in credential, File Operations, filesystem-command, splash, and
performance-mirror code even though `Common/WindowSizing.h` already owns the pure centered-rectangle primitive.
The FunctionBar and StatusBar retain separate HWND render-target lifecycles. F1-F5 are complete; F6 is now the
only open Track F batch and remains gated on archived current-head evidence before its shared lifecycle is
introduced.

- **F1 viewer chrome, Unicode clipboard, and owner centering — implemented, verification complete:**
  `ViewerFileComboHost` now owns the exact mouse/keyboard popup-opening predicate and the standalone eight-row
  popup-height profile used by ViewerPE, ViewerWeb, ViewerImgRaw, and ViewerText. `UnicodeClipboard.h` owns
  `CF_UNICODETEXT` allocation, lock, terminator, transfer, and release semantics; ViewerText preserves its
  allow-empty behavior, while ViewerWeb explicitly preserves its reject-empty boundary. `WindowSizing.h` now
  owns pure owner-centering geometry plus unbounded and owner-work-area-clamped placement; the credential
  prompt, File Operations popup, filesystem command dialogs, two app-hosted dialogs, and the performance mirror
  delegate without changing z-order or activation. Splash centering remains a distinct topmost/fallback-monitor
  policy and is not folded into this helper. Golden tests cover every popup-opening message, rejected messages,
  row caps, DPI scaling, negative monitor coordinates, and odd midpoint rounding; Pester source contracts reject
  reintroduced viewer-local popup/clipboard bodies. Theme, all five viewer-chrome contracts, PluginContractTests,
  the complete ViewerPETests integration executable, the credential-prompt DX validation case, and the File
  Operations popup progress-contract case pass. A full Debug x64 build completed in 2:14 with 0 warnings/errors
  (`.build/logs/msbuild-20260715_162248_534.log`). This batch does not touch a render hot path and therefore does
  not consume Track F's required FunctionBar/StatusBar before/after performance gate. F1 is committed as
  `950d049d` (`Consolidate viewer chrome and centering helpers`).
- **F2 resolved-theme pipeline — implemented, verification complete:** `AppTheme.cpp` now owns the single
  built-in theme-id mapping, full static color-override application, accent-before-palette sequencing, custom
  definition resolution, invalid-definition fallback, high-contrast bypass, and compiled dynamic-override
  attachment through `ResolveAppThemeSelection`. Startup, configured-theme validation, settings hot reload,
  dialog resolution, the Preferences live dialog, Themes grid/swatch preview, and Themes edit validation all
  delegate to that result. The local `ApplyThemeOverrides`, `ApplyAppThemeOverrides`,
  `ApplyDialogThemeOverrides`, accent extractors, and three theme-id parsers are removed. Settings ownership,
  watcher invalidation, editor dirtiness, and UI application remain at their previous call sites. New snapshot
  tests cover all built-in IDs plus system fallback, custom ARGB alpha/channel propagation, GDI/D2D navigation
  synchronization, file-operation derived colors, and invalid-reference fallback without partial colors. Theme,
  all eight theme-distribution source contracts, `theme_v2_runtime_resolution_and_dynamic_perf`, hot-reload
  preference merging, Preferences theme-cycle, and live set-theme interaction pass. The RedSalamander Debug x64
  build completed in 2:16 with 0 warnings/errors (`.build/logs/msbuild-20260715_163626_923.log`). The durable
  resolver ordering and high-contrast contract are merged into `Specs/UI/UI_VisualStyle.md`. F2 is committed as
  `26cbdedb` (`Unify app theme selection resolution`).
- **F3 viewer title-bar policy — implemented, verification complete:**
  `Common/ViewerTitleBarTheme.h` now owns first-party viewer dark/light/high-contrast DWM attributes, deterministic
  rainbow accent projection, inactive-caption blending, and caption-text contrast selection. ViewerPE,
  ViewerWeb, ViewerImgRaw, ViewerText, ViewerVLC, and ViewerSqlite are thin delegates; their window procedures,
  theme state, activation decisions, and viewer-specific accent seed ownership remain local. ViewerSpace stays
  outside this batch because its dynamically loaded DWM/topmost host contract is materially different. Golden
  theme tests and a source contract cover the shared policy and reject reintroduced local DWM bodies. The focused
  DxUi Theme suite, six ViewerChrome source contracts, and complete ViewerPETests integration executable pass;
  the full Debug x64 build completed with 0 warnings/errors
  (`.build/logs/msbuild-20260715_164357_196.log`). The durable policy and ViewerSpace exclusion are synchronized
  into `Specs/Plugins/Plugins_ViewerPlugins.md`. F3 is committed as `741e8506`
  (`Consolidate viewer title bar theming`).
- **F4 modal DxUi shell — implemented, verification complete:** the pre-change self-test inventory covers About
  and Fatal Error owner assignment, owner-disabled modality, repeat open/close, Enter/Escape and UIA close paths,
  Fatal Error theme/scroll behavior, and restoration of folder-view focus after About closes. This satisfies the
  prerequisite owner-lifetime/focus gate. The implementation batch is restricted to shared client-DIP window
  creation, owner disable/reenable, owner centering, show/foreground ordering, and `WM_QUIT`-preserving nested
  message-loop behavior. Dialog control trees, class registration, WndProc dispatch, theme/layout, commands, and
  result ownership remain local, and session-end handling is explicitly outside the helper. `ModalWindowShell`
  now owns those shared host operations and delegates pumping to the existing `RunDxUiModalLoop`, which preserves
  the `WM_QUIT` code and now also preserves `GetLastError` across its shared failure diagnostic. Both dialogs use
  `_hWnd.reset()` for owned-window close. All three modal-shell source contracts and the 13 focused About/Fatal
  Error cases pass, including repeated churn, owner modality, keyboard/UIA close, Fatal Error scrolling/theme,
  and About navigation-shell focus restoration. The Debug x64 application build completed with 0 warnings/errors
  (`.build/logs/msbuild-20260715_170022_184.log`). The durable contract is synchronized into
  `Specs/UI/UI_DxUiSharedGrid.md`, and the new source guard is inventoried in
  `Specs/Testing/Testing_TestCoverage.md`. F4 is committed as `cef0d142`
  (`Consolidate app modal window shell`).
- **F5 palette-adapter reconciliation — implemented, verification complete:** the final census found that Find
  Files and the File Operations Issues pane carried byte-for-byte identical folder-content palette builders, while Compare
  Directories, Manage Plugins, and the Preferences namespace added exact one-line pass-through adapters around
  `MakeAppThemeDxPalette`. The duplicate folder profile moved to the named
  `MakeFolderContentDxPalette`; exact pass-throughs were removed. The credential-prompt palette remains a
  named local policy because it intentionally differs in surface/header/overlay token projection. The new folder
  profile golden test, complete DxUi Theme suite, and all nine theme-distribution source contracts pass. Focused
  runtime coverage passes for Find Files/modeless ownership, File Operations Issues theme cycling, the credential
  prompt, Compare Options, Manage Plugins navigation stability, and Preferences theme cycling. The Debug x64
  application build completed with 0 warnings/errors (`.build/logs/msbuild-20260715_170913_738.log`). The durable
  profile/allowlist contract is synchronized into `Specs/UI/UI_DxUiSharedGrid.md`, and the source guard is
  inventoried in `Specs/Testing/Testing_TestCoverage.md`. F5 is committed as `9f6b7149`
  (`Consolidate app theme palette profiles`).
- **F6 FunctionBar/StatusBar render-resource lifecycle — implemented, verification complete:** protect repeated
  FunctionBar visibility/layout churn and active/inactive pane StatusBar refreshes at the existing synchronous
  paint boundary. The stable metrics are `render.function_bar.paint_us` and `render.status_bar.paint_us`; the
  behavioral sentinels are `cmd_app_toggleUiChrome`,
  `cmd_app_toggleUiChrome_keeps_navigation_shell_stable`, and
  `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable`. The test-enabled Release x64 build completed
  with 0 warnings/errors (`.build/logs/msbuild-20260715_171521_606.log`). Current-head commit `9f6b7149` baseline
  archives are `Specs/TestRuns/4cb089111a23/Commands/2026-07-15_171844/` (`cmd_app_toggleUiChrome`, 10/10),
  `2026-07-15_171853/` (StatusBar navigation-shell, 10/10), and `2026-07-15_171932/` (chrome
  navigation-shell, 10/10). For direct chrome churn, FunctionBar paint is n=40/p50=1,694us/p95=12,409us and
  StatusBar paint is n=60/p50=688us/p95=2,016us. For the two navigation-shell sentinels, StatusBar paint is
  n=91/p50=543us/p95=4,484us and n=100/p50=636us/p95=1,267us; chrome navigation-shell FunctionBar paint is
  n=31/p50=1,501us/p95=4,190us. Cold-start outliers are retained in the archives and are not removed from the
  comparison. The candidate must
  preserve paint count/coverage, retain factories and typography across ordinary paints, release only
  target-dependent resources on device loss, release all HWND-owned resources on the surface's established HWND
  teardown message, and show no
  material median/tail regression in the two paint metrics. Archive baseline and candidate artifacts under
  `Specs/TestRuns/<machine>/Commands/`; do not close F6 on source-shape evidence alone. Implementation now uses
  `Common::Rendering::HwndRenderTargetResources` as the sole owner of the shared D2D factory, HWND target, and
  target-dependent brush lifecycle. FunctionBar/StatusBar keep their DirectWrite factories, text formats,
  trimming, DPI policy, drawing, and window-property ownership. Target reset retains factory and typography;
  the respective HWND teardown (`WM_DESTROY` for FunctionBar, `WM_NCDESTROY` for StatusBar) releases the full
  aggregate. The new three-case source contract passes, all three focused Debug behavior sentinels pass, and the
  Debug x64 build is warning-clean (`.build/logs/msbuild-20260715_172221_610.log`). The optimized test-enabled
  Release x64 candidate build is also warning-clean (`.build/logs/msbuild-20260715_172617_472.log`), and the
  exact-commit (`30a2b3f3`) 10-repeat candidate archives are
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-15_173025/`, `2026-07-15_173030/`, and
  `2026-07-15_173036/`; all 30 cases pass. Paint counts are identical before/after
  for each paired case (40/60, 1/91, and 31/100 FunctionBar/StatusBar samples respectively). Direct chrome
  FunctionBar p50 improves while p95 moves only +340us (1,694/12,409us -> 1,566/12,749us); direct chrome
  StatusBar median is flat and p95 improves (688/2,016us -> 691/1,516us). StatusBar navigation-shell p50 moves only +56us while p95
  improves (543/4,484us -> 599/4,143us). Chrome navigation-shell FunctionBar is effectively flat at
  1,501/4,190us -> 1,493/4,640us, and its StatusBar p50 improves while p95 moves only +151us
  (636/1,267us -> 574/1,418us). These sub-millisecond tail movements are normal run-to-run variance and not a
  material responsiveness regression. Retained-resource characterization is also unchanged: baseline and candidate both
  retain D2D/DirectWrite factories and typography across ordinary paint, reset only brush/target on size or
  device loss, and destroy the complete per-HWND aggregate at teardown; the new source contract locks that
  boundary. Track F is complete.

**Initial scope:** AppTheme startup/hot reload/Preferences preview, DxUi palette adapters, FunctionBar,
StatusBar, About/Fatal Error dialogs, credential/folder centering users, and first-party viewers except
ViewerSpace where explicitly excluded.

1. Snapshot resolved themes and all named palette profiles, then route startup, hot reload, and preview through
   one pure resolver without changing invalidation or file-watcher ownership.
2. Adopt shared title-bar, clipboard, centering, popup input, and popup-height helpers one viewer/dialog at a
   time. Keep window procedures and mutable state local.
3. Introduce the modal shell only after owner-lifetime and focus-restoration tests exist; prohibit its use in
   session-end paths.
4. Extract the FunctionBar/StatusBar HWND render-resource lifecycle last. Capture responsiveness and retained-
   resource evidence before and after because this touches a render hot path.

**Focused definition gate:**

```powershell
rg -n "ApplyThemeOverrides|Make(AppTheme)?DxPalette|ApplyTitleBarTheme|MessageMayOpenWindowComboPopup|ComputeWindowComboPopupHeightPx|CopyUnicodeTextToClipboard|Center.*(Owner|Window)|CreateHwndRenderTarget" Common Plugins RedSalamander RedConfigure Tests
```

**Verify:** run Commands, DxUi, affected viewer tests, mixed-DPI/device-loss UI stress, and the Full suite. Archive
the required before/after performance run under `Specs/TestRuns/` for the render-host change.

### Track G — Plugin contract and lifetime consolidation (`LH-D12`, `LH-D13`, `LH-D22`)

**Priority:** P1
**Effort:** L
**Risk:** HIGH
**Depends on:** Firebreak Task 1 must be complete or formally handed off before overlapping MTP unload work.

**Status (2026-07-16): DONE.** Firebreak Task 1 is complete in
`Specs/Plans/Done/Operation_Firebreak_ConsolidatedReviewRemediation_2026-07-06.md`; Lighthouse will not reopen its
MTP cancellation/browse-worker changes. The fresh definition census found: (1) Google Drive and Microsoft
Drive retain byte-for-byte equivalent navigation callback generation/in-flight/drain state while
`EmbeddedViewerBase` already uses `RegistrationCallbackState`; (2) six buffered providers (7z, Curl, Google
Drive, Microsoft Drive, MTP, and S3) duplicate packed `FileInfo` sizing/layout/query logic, while local
`FilesInformation` is a resumable streaming enumerator and Dummy adopts a prebuilt fixture buffer and therefore
remain separate; (3) module-pinned viewer/threadpool callbacks repeat the exact callback-return transfer, but
MTA/STA setup is not globally interchangeable—WIC workers tolerate `RPC_E_CHANGED_MODE`, MTP lifetime workers
require an MTA, and UI/export paths intentionally use STA; and (4) Compare, Find Files, and selftests repeat the
manager's private case-insensitive ID lookup. Implementation is restricted to those exact families: migrate the
two callback states, add one checked packed-buffer owner beneath the six existing COM facades, expose the const
UI-thread manager lookup with a pointer lease valid only until the next manager mutation, and centralize exact
module-pin transfer. Provider sorting, metadata extraction, cancellation, task accounting, COM-apartment policy,
and local/Dummy enumeration remain explicit.

**Work ledger (2026-07-15):** `LH-D12` callback/module-pin migration and `LH-D22` manager lookup migration are
implemented. Google Drive and Microsoft Drive now use `RegistrationCallbackState`; exact threadpool module-pin
transfers route through `TransferModulePinToCallbackReturn`; and Compare, Find Files, and their selftest helpers
delegate to the manager's public const case-insensitive lookup. The first Debug build exposed private mutable
overload shadowing (`.build/logs/msbuild-20260715_173730_269.log`, four compile errors); separating that internal
path as `FindMutablePluginById` fixed the public lookup, and the corrected x64 Debug build is clean
(`.build/logs/msbuild-20260715_174040_890.log`, zero warnings/errors). The focused source contracts pass 3/3.

`LH-D13` is implemented. `Common/PackedFileInfoBuffer.h` owns checked aggregate/name
sizing, `alignof(FileInfo)` layout, zero/single/multi-entry construction, common buffer queries, and bounded
offset traversal beneath the 7z, Curl, Google Drive, Microsoft Drive, MTP, and S3 COM facades. Sorting and
provider metadata mapping remain local. Local streaming enumeration and Dummy's prebuilt fixture buffer remain
excluded. Direct runtime contracts were added to `PluginContractTests`, with a source-contract suite covering
the exact six-owner boundary.

Track G focused validation is green: the two source-contract files pass 6/6; the rebuilt x64 Debug solution is
clean (`.build/logs/msbuild-20260715_174918_425.log`, zero warnings/errors); `PluginContractTests.exe` passes its
new empty/multi-entry/alignment/malformed-buffer cases and every loaded provider/debug-selftest contract; and
`google_drive_navigation_menu_callback_clear_drains` passes 1/1 through `Run-AllTests.ps1` (TestSandbox run
`20260715T155320Z-68320-8a7c159fd6f94071979e3c641d5cf6e3`). The ARM64 Debug solution also builds with zero
errors (`.build/logs/msbuild-20260715_175327_575.log`); its 80 C4324 warnings are the existing
`FolderWindow::PaneState`/`FolderWindow` ARM64 padding diagnostics, not touched-code diagnostics. Authoritative
callback, module-pin, packed-result, plugin-lookup, and test-coverage contracts are merged into
`Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Plugins/Plugins_ViewerPlugins.md`, and
`Specs/Testing/Testing_TestCoverage.md`. The repository-wide Full suite remains the final Lighthouse closeout
gate after Tracks H and I, so this track is not yet declared DONE in isolation.

**Initial scope:** `Common/Helpers.h`, filesystem/viewer plugins that retain exact callback or module-pin
envelopes, packed `IFilesInformation` implementations, `FileSystemPluginManager`, Compare, Find Files, and
PluginContractTests.

1. Characterize callback registration/replacement/clear and quiet-point ordering, then migrate exact callback
   states to `RegistrationCallbackState` one provider at a time.
2. Add the narrow callback-owned module-pin transfer and identical MTA entry envelope. Do not absorb provider
   cancellation, watchdogs, result queues, or task accounting.
3. Add an ABI-neutral packed-result builder with checked arithmetic. Migrate one provider and run consumer
   compatibility tests before each subsequent provider.
4. Expose the manager's const plugin lookup with an explicit lease/lifetime contract; remove external scans.
5. Run teardown/unload stress after every provider migration and on x64 plus ARM64 builds.

**Focused definition gate:**

```powershell
rg -n "RegistrationCallbackState|CallbacksInFlight|TrySubmitThreadpoolCallback|FreeLibraryWhenCallbackReturns|CoInitializeEx\(COINIT_MULTITHREADED\)|class FilesInformation|IFilesInformation|Find.*FileSystemPlugin" Common Plugins RedSalamander Tests
```

**Verify:** focused PluginContract/provider tests, repeated cancel/unload/refresh stress, Full suite, and Debug
x64 plus ARM64 builds. Any callback after quiet point, leaked packed block, or unload hang is blocking.

### Track H — Application/settings model consolidation (`LH-D6`, `LH-D8`, `LH-D16`, `LH-D19`)

**Priority:** P2-P3
**Effort:** M-L
**Risk:** MED
**Depends on:** Track E strict JSON/string/I/O helpers.

**Status (2026-07-15): DONE.** The fresh census confirmed all four routed families were actionable.
Manage Plugins and Preferences retain independent plugin-schema field models/parsers and configuration codecs;
SettingsStore repeats the same grid-layout array parser four times, with Batch Rename's control-character
normalizer as the only intentional per-section difference; Compare retains two latest-snapshot progress drains
and Find retains result-merge plus latest-progress drains, while Find's existing queue-head probe already proves
that unrelated older HWND input must not be skipped; and WorkspaceDiscovery/RedConfigureSession retain equivalent
`std::ifstream` binary readers. Track H landed the low-risk grid/binary-reader primitives, then the
plugin-schema model/codec, then the keyed contiguous posted-payload helper. The payload migration places the
operation/run key in `wParam` so the helper stops before consuming a same-message payload for another run.
Provider-specific/UI layout state, grid column-ID normalization, result-drain budgets, progress metrics, and
RedConfigure error ownership remain local.

**Work ledger (2026-07-15):** `LH-D8` is DONE: the four SettingsStore grid-layout readers now use one
callback-parameterized parser, Batch Rename keeps its stricter column-ID sanitizer, and a focused Commands case
covers missing/wrong-type entries, duplicates/order, width clamps, unknown fields, and section-scoped recovery.
`LH-D19` is DONE: both RedConfigure call sites use one local reader backed
by the shared bounded-size/exact-read primitives; it preserves read/write/delete sharing and empty files, caps
inputs at 64 MiB, clears output on failure, returns the originating HRESULT without logging, and has direct
empty/exact-bound/oversize/missing/locked/Unicode-path tests. The durable contracts are recorded in
`Core_SettingsStore.md` and `Core_RedConfigure.md`. `LH-D6` and `LH-D16` are also complete as detailed below.

The first `RedConfigureTests` build exposed an over-broad include cleanup: `RedConfigureSession.cpp` still
needs `<fstream>` for its export writers after removing the local input reader. The four compile errors are
captured in `.build/logs/msbuild-20260715_180324_990.log`; the include is restored before the clean rerun.

Focused validation is green. `RedConfigureTests` rebuilt cleanly with zero warnings/errors
(`.build/logs/msbuild-20260715_180350_089.log`) and the complete executable passed, including the new reader
case (repo-sized scan 219 ms, validation 3 ms, preview 222 us). `SettingsSchemaTests.exe` passed its complete
real/synthetic schema set. The new `settings_store_grid_layout_parsing_parity` Commands case passed 1/1 in
TestSandbox run `20260715T160918Z-45972-c8c095cec171463cb534f1531c32c34a`. An initial 120-second local command
limit interrupted the broad RedSalamander dependency build and correctly raised the artifact-contamination
marker; the required serialized full Debug rebuild then cleared it and completed with zero warnings/errors in
`.build/logs/msbuild-20260715_180624_643.log`. These rows do not change a product hot path and require no new
performance archive beyond the already-running RedConfigure repo-sized acceptance scenario.

`LH-D6` implementation has started with the dependency-layer model/codec and its golden tests. The first
`SettingsSchemaTests` build found a missing public `HRESULT` type include plus two `/Wall` local-shadow warnings;
the diagnostic is retained at `.build/logs/msbuild-20260715_181434_688.log` and the header/warnings are corrected
before migration of either UI entry point.

The next build compiled Common cleanly but exposed that the result-type convenience methods were not exported
from the DLL (`.build/logs/msbuild-20260715_181457_128.log`). They are now dependency-free inline observers on
the public result structs; the codec functions remain the only exported implementation surface.

The clean codec build is `.build/logs/msbuild-20260715_181522_888.log` (zero warnings/errors). Its first direct
run passed every model, validation, coercion, unknown-member, and serialization behavior; the sole red assertion
mistook yyjson's positive-integer reparse category (`uint64_t`) for a signed value. The expectation now accepts
the same numeric value in either JSON integer storage category before the UI migrations proceed.

`LH-D6` is DONE. `Common::PluginConfiguration` owns the
field/choice/value model, explicit schema/configuration validation results, default/coercion behavior, and
unknown-member-preserving serializer. Manage Plugins and Preferences both use that surface for schema parsing,
configuration parsing, and serialization; their local code now only maps shared values to UI controls and back.
Both serializers overlay edits on the original JSON object, so unknown plugin members survive either entry
point, while cancel remains a no-save path. The migration also exposed a real divergence: Preferences rendered
every option with at least two choices as an on/off toggle and left both choice indices at zero, while Manage
Plugins required exactly two recognizable boolean choices. Boolean-option recognition is now shared, accepts
boolean tokens in labels with values as fallback, and sends every non-boolean or multi-choice option to the
normal choice control on both surfaces. Golden tests cover all five field types, ordering and UI metadata,
coercion, wrong-typed selection members, validation codes, unknown-member retention, numeric clamping, and the
shared toggle decision. The durable behavior contract is in `Core_SettingsStore.md`. `SettingsSchemaTests`
rebuilt with zero warnings/errors (`.build/logs/msbuild-20260715_182241_815.log`) and its complete synthetic,
golden plugin-codec, and real-schema suite passed. The first RedSalamander migration build was functionally
successful but `/Wall` found two unused local UTF conversion wrappers left by deleting the old parser
(`.build/logs/msbuild-20260715_182300_556.log`); they and the stale yyjson include were removed. The clean
application rebuild then completed with zero warnings/errors
(`.build/logs/msbuild-20260715_182604_418.log`). Closeout review found that Manage Plugins still restored a
schema default when a user intentionally cleared a text edit, while Preferences retained the empty string; the
Manage mapper now preserves the empty edit. The shared serializer also propagates `E_OUTOFMEMORY` instead of
discarding the original forward-data object after that specific parse failure. The final D6 application rebuild
is `.build/logs/msbuild-20260715_182931_220.log` (zero warnings/errors); the complete schema/golden executable
and the anti-drift source contract pass again. Manage Plugins live DX interaction, edit/apply, reopen, and cancel
behavior passed in run `20260715T163154Z-4748-39ee847e9f4843efa314900579c6c521`; the Preferences plugin
configuration interaction passed in `20260715T163201Z-91296-9a5bd90c4fa14abe84fb0b24d002261c`.

`LH-D16` is DONE. `Common/Helpers.h` now provides a typed contiguous posted-payload
drain that takes the currently dispatched payload, inspects the unfiltered thread-queue head, and removes only
an exact HWND/message/`wParam` operation-key match. Caller-supplied predicates retain cancellation and drain
budgets; caller reducers retain Find result merging and latest-wins Compare/Find progress semantics. Compare
scan/content posts now key by compare run id, and Find result/progress/completion posts key by search epoch; all
keys are captured before moving their payloads. Compare and Find have been migrated without changing their
existing queue/latency counters. A deterministic DxUi test now covers contiguous reduction, different operation
ids, completion/interleaved-message boundaries, caller budget, cancellation, final-snapshot retention, and
exactly-once payload destruction; the existing destroy-with-64-queued-payload stress remains adjacent. The
authoritative queue-order contracts are updated in `Core_CompareDirectories.md` and `UI_FindFilesWindow.md`.
A new Pester source contract prevents local queue-drain loops and unkeyed Compare/Find posts from returning;
the adjacent test-harness module-pin contract was updated to recognize the Track G shared transfer helper.

The first `DxUiTests` build compiled the helper and test but `/Wall` reported the result template's implicitly
deleted copy constructor/assignment at its first `unique_ptr` instantiation; the diagnostic is retained in
`.build/logs/msbuild-20260715_183706_459.log`. The result type now explicitly deletes copying and defaults
noexcept moves before the clean rerun.

The final `DxUiTests` build is `.build/logs/msbuild-20260715_183927_043.log` and the complete executable passed.
The application build is `.build/logs/msbuild-20260715_184300_357.log`; both builds report zero warnings and
zero errors. The stale-epoch and child-input-order cases passed in committed archives
`Specs/TestRuns/4cb089111a23/Commands/2026-07-15_184536` and `2026-07-15_184542`; active-search cancellation and
window teardown passed in `2026-07-15_185046`. The large incremental Find run passed in
`2026-07-15_184550`: 32 result handlers coalesced eight additional queued result messages while preserving the
256-record/8-message budget, and three progress handlers delivered the final state. Compare progress passed in
`2026-07-15_184557`: four content handlers coalesced 238 queued updates (maximum 132 in one handler), while
three scan handlers preserved their final snapshot. These are same-machine Debug acceptance runs, not a
cross-machine percentile claim. All five curated run folders pass `Tools/Test-TestRunArchive.ps1`; the two
coalescing source-contract files pass 121/121 assertions together. Track H has no remaining implementation item.

1. Build the plugin-configuration model/codec from golden documents, then migrate Manage Plugins and
   Preferences separately. Preserve unknown fields and cancel/no-save behavior.
2. Replace repeated SettingsStore grid blocks with one descriptor/callback parser only after section-scoped
   recovery tests cover each current section.
3. Add typed contiguous payload coalescing and migrate one Compare/Find queue at a time under interleaving and
   destroy stress.
4. Replace the two RedConfigure binary readers with one RedConfigure-local bounded helper.

**Focused definition gate:**

```powershell
rg -n "PluginConfiguration|Grid.*(Columns|Layout)|PeekMessageW|TakeMessagePayload|Read(Binary)?File|GetFileSizeEx" Common RedConfigure RedSalamander Tests
```

**Verify:** SettingsSchema, Commands, Compare, RedConfigure, plugin-configuration UI selftests, payload teardown
stress, and the Full suite.

### Track I — Shared test support (`LH-D18`)

**Priority:** P2 test reliability
**Effort:** M
**Risk:** MED
**Depends on:** Track E environment/string and handle-I/O helpers.

**Status (2026-07-16): DONE.** The fresh census reconfirmed seven copies of the unified TestSandbox
root/run-id/segment logic across ViewerPE, ViewerSqlite, RedConfigure, CrashHandling, DxUi, and Performance
tests (DxUi uses the artifact branch; the others use scratch). The copies have two intentional policies that
the shared surface must name: the empty leaf fallback (`case` versus Performance's `default`) and whether an
existing case directory is removed before creation (only ViewerPE and Performance clean; ViewerSqlite,
RedConfigure, CrashHandling, and DxUi retain it). ViewerPE and
ViewerSqlite also retain equivalent message-pump loops and typed snapshot polling. Six first-party
`CreateProcessW` call sites remain: ViewerPE's suspended loader probe and the RC compiler have specialized
launch semantics, while Compare, Commands Settings, and plugin-configuration cases repeat ordinary launch/
wait/capture work. The shared implementation will be a test-only header so standalone test executables and
the in-process RedSalamander selftests can consume it without adding a production DLL ABI. Work is starting
with the sandbox/environment primitives and their characterization tests; message pumping follows, then the
bounded concurrent child-process runner and Compare migration.

The sandbox/environment substep is **DONE**. `Tests/TestSupport/TestSupport.h` now owns environment reads and
exact scoped restoration plus TestSandbox root/run-id/segment resolution; ViewerPE,
ViewerSqlite, RedConfigure, CrashHandling, PerformanceTests2, and DxUi use thin policy adapters over it. A DxUi
characterization test covers missing/value environment restoration, temporary removal, sanitization,
clean-versus-retain acquisition, caller-selected empty-leaf fallback, and the artifact directory branch. The
first DxUi build compiled the implementation but `/Wall` rejected the relative `..` include spelling in all 18
translation units (`C4464`); the diagnostic is retained in
`.build/logs/msbuild-20260715_185947_541.log`. The first conditional test include-root spelling did not match
MSBuild's normalized repo-relative project directory and produced 18 missing-header errors in
`.build/logs/msbuild-20260715_190226_375.log`; the corrected test-project-only include root then built without
suppression. Clean/focused evidence is green: DxUi
(`.build/logs/msbuild-20260715_190339_207.log` plus the full direct executable), ViewerPE
(`.build/logs/msbuild-20260715_190647_238.log` plus `TestViewerWebSecurityPolicyAndBounds`), ViewerSqlite
(`.build/logs/msbuild-20260715_190700_780.log` plus `TestListTables`), RedConfigure
(`.build/logs/msbuild-20260715_190714_855.log` plus the full direct executable), CrashHandling
(`.build/logs/msbuild-20260715_190756_829.log` plus the full direct executable), and PerformanceTests2
(`.build/logs/msbuild-20260715_190814_479.log` plus VSTest
`LargeFolderIconEnumeration_MixedItems`). `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` is also fully
green after moving the six harness contracts from local implementation tokens to shared-ownership and
per-adapter policy checks. The next active substep is bounded message-pump/snapshot polling.

The bounded polling substep is now **DONE** in the same test-only header. It caps each message
batch, reports operation/budget/elapsed/message-count timeout context, and preserves the last typed snapshot
on timeout. ViewerSqlite and ViewerPE thin adapters have been migrated without changing their 20 ms cadence.
The first characterization probe incorrectly attempted asynchronous `WM_SETTEXT`; Windows rejected that
pointer-bearing system message, so the full DxUi run failed at the test setup rather than the helper. The
probe now posts a pointer-free `WM_APP` message and the focused WindowHost suite passes after a zero-warning
build (`.build/logs/msbuild-20260715_191715_945.log`). Two integration diagnostics remain recorded for the
current rerun: ViewerSqlite exposed one now-removed unreferenced local adapter (`C5245` in
`.build/logs/msbuild-20260715_191743_462.log`), and ViewerPE exposed the Win32 `min` macro at the shared
`std::min` call (`C2589` in `.build/logs/msbuild-20260715_191806_437.log`); the shared call now uses the
macro-safe `(std::min)` spelling. Corrected ViewerPE and ViewerSqlite builds are zero-warning
(`.build/logs/msbuild-20260715_191836_451.log` and
`.build/logs/msbuild-20260715_191847_208.log`); their focused `TestViewerWebSecurityPolicyAndBounds` and
`TestListTables` cases pass. The focused WindowHost suite and a subsequent full DxUi executable pass, and the
source-contract suite is fully green with shared-ownership guards. The child-process runner is now the active
and final Track I implementation substep.

The child-process substep is **DONE**. `Tests/TestSupport/ChildProcess.h` now launches a suspended root,
allowlists only NUL/stdout/stderr inheritance through `PROC_THREAD_ATTRIBUTE_HANDLE_LIST`, assigns the root to
a WIL-owned kill-on-close JobObject before resume, drains stdout and stderr concurrently while retaining
bounded prefixes, and exposes timeout/cancellation/nonzero-exit diagnostics plus Windows argument quoting.
Compare's former wait-before-read helper now delegates to this runner and uses structured argument vectors.
A Debug-only SearchService probe and the new `test_support_child_process_runner_contract` cover output beyond
pipe capacity, bounded truncation, quoted Unicode/space/quote/trailing-backslash arguments, independent
nonzero streams, unrelated inheritable-handle exclusion, timeout process-tree kill, cancellation, and launch
failure. The probe executable builds with zero warnings
(`.build/logs/msbuild-20260715_192728_202.log`). The first full RedSalamander project build exceeded the
120-second command budget while still compiling and exposed only WIL template copy-operation `/Wall`
diagnostics from the new header; the header-local WIL warning scope now covers those instantiations before the
rerun. No compiler error was reported before the external timeout. The repository contamination guard then
required the full-solution recovery rebuild; Debug/x64 completed with zero warnings and zero errors in
`.build/logs/msbuild-20260715_193016_564.log`. The new Compare case passed 1 / 0 / 0 through
`Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter test_support_child_process_runner_contract` under run
id `20260715T173411Z-93784-77ccd2958a9645ee9839834eb972ef31`, and the updated source-contract suite is fully
green. The durable test-support ownership and containment rules are merged into
`Specs/Testing/Testing_TestCoverage.md` and `Tests/README.md`. All Track I implementation substeps are now
complete; repository-wide closeout gates remain before the track and operation can be declared DONE.

Closeout attempt 1 started with a clean Debug x64 build (0 warnings / 0 errors,
`.build/logs/msbuild-20260715_193640_905.log`) under run id
`20260715T173639Z-32712-2c3774bfef76442f854fb8ae8090abc8`. Its first CompareDirectories pass reported
225 passed, 1 failed, and 31 skipped; the only failure was
`search_service_sqlite_delete_burst_maintenance_preserves_query_parity`, which timed out after 17.979 seconds
waiting for maintenance or the first compaction window (`lastHr=0x80070002`). The classifier's complete retry
then passed all 257 cases (226 passed, 0 failed, 31 environment-dependent skips). The outer 15-minute command
budget expired while the classifier was continuing isolated confirmation, so this interrupted attempt is
diagnostic evidence only and does **not** count as either required clean CI pass. Focused reproduction and two
fresh uninterrupted CI passes remain in progress. The first focused-reproduction command then found the stale
owner lock from the interrupted CI and, by design, created `.build/artifact-operation-contaminated.json` and
refused `-SkipBuild`; a full Debug x64 rebuild is therefore the mandatory next gate before reproduction.
That mandatory rebuild completed in 2m38s with 0 warnings and 0 errors
(`.build/logs/msbuild-20260715_195407_768.log`), clearing the contamination gate; focused reproduction is now
running against the rebuilt artifact set. The exact case then passed 1 / 0 / 0 in 4.7 seconds under run id
`20260715T175652Z-83860-5b3a95e1f2c447c4aed50709e2267c04`; no Lighthouse code change was needed. Fresh
CI pass 1 is now running with an outer budget large enough to complete failure classification if a recurrence
occurs. Live artifact inspection later showed that this attempt entered Commands classification and began
re-executing a smaller Commands slice; it therefore will not count as a required clean CI pass even if all
retries pass. The invocation's canonical build completed in 2m14s with 0 warnings and 0 errors
(`.build/logs/msbuild-20260715_195715_549.log`). The initial Commands failure was
`folderView_thumbnail_size_change_while_pending`, proven by the classifier's saved focused `case_filter`; its
isolated retry passed in 459 ms. Classification then entered its three complete shuffled Commands triage runs.
The 30-minute outer budget expired during the first recorded shuffle after 187 passes and four unrelated
UI-state failures: `cmd_pane_navigation_hot_paths_keeps_navigation_shell_stable`,
`cmd_preferences_dialog_viewers_reordered_resized_copy_follows_visible_columns_after_sort_cycles_and_search_roundtrip`,
`cmd_preferences_dialog_keyboard_tab_traversal_live_dx_interaction`, and
`cmd_pane_fileops_issues_pane_restores_combined_view_state_after_recreate`. Because that shuffle was
interrupted (its result retained `duration_ms=0`), these four are partial order/isolation diagnostics rather
than completed classifications. The operation remains short of a clean CI pass, and the interruption again
requires a full rebuild before any focused or CI execution.
The second recovery rebuild completed in 2m45s with 0 warnings and 0 errors
(`.build/logs/msbuild-20260715_202905_415.log`). A standalone default-order Commands reproduction, with CI
classification disabled so its original result cannot be overwritten by shuffled retries, completed under
run id `20260715T183207Z-96644-fa533f09f48449cd87382ae429ca8e99`: 802 passed, 5 failed, and 2 opt-in cases
skipped in 15m24.5s. The earlier `folderView_thumbnail_size_change_while_pending` case passed in this run, so
its CI failure did not reproduce. The five preserved failures were instead focus/accessibility UI-state checks:
`cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable`,
`cmd_pane_selection_mask_prompt_live_dx_interaction`,
`cmd_pane_selection_mask_prompt_long_run_open_close_stays_stable`, `cmd_app_openDriveMenus`, and
`cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane`. The navigation-shell, drive-menu, and
unfocused-pane diagnostics reported a null focused HWND/folder view; the two selection-mask prompt checks
reported an unstable accessible name. Focused isolation of all five cases is now in progress before deciding
whether this is product behavior, test order leakage, or an unavailable interactive desktop condition.
Focused isolation passed the navigation-shell case (run
`20260715T184826Z-99540-df7c55100e6d4234bdac425bf7e89e45`), drive-menu case (run
`20260715T184850Z-101812-74f280313b3f4475ad2b2117c6f789e4`), and unfocused-pane case (run
`20260715T184858Z-98064-d833fb2483254ffda5f210489b951a6c`). Both selection-mask cases reproduced
independently: live interaction failed under run `20260715T184836Z-13716-53dbf2e1b6cc4dbbb98f75d2dcf0f756`,
and long-run open/close failed under run `20260715T184844Z-44640-23cef5de3cb548cd8f9019739dd0b247`.
Source inspection found the concrete defect in `FolderViewSelectionMaskPromptWindow::BuildUi`: the visible label
owns `_labelText`, but the associated DxUi `TextField` never receives that localized label through
`SetAccessibleName`, leaving UIA Name empty. The targeted accessibility fix and focused regression reruns are
now in progress; `BuildUi` assigns `_labelText` to the field's explicit accessible name, and the durable
selection-mask prompt contract plus both validation anchors are recorded in `Specs/UI/UI_FolderWindow.md`.
The other three failures remain classified as order/desktop-sensitive diagnostics, not independent product
reproductions. The scoped Debug x64 `RedSalamander` build completed in 2m12s with 0 warnings and 0 errors
(`.build/logs/msbuild-20260715_205007_553.log`). Both formerly reproducible RED cases are now GREEN:
`cmd_pane_selection_mask_prompt_live_dx_interaction` passed under run
`20260715T185223Z-97072-ae4ced131a2341ca850a587fd9b77f8e`, and
`cmd_pane_selection_mask_prompt_long_run_open_close_stays_stable` passed under run
`20260715T185230Z-102396-5e62467850714fe2a02ce24b033ca7e3`. A complete default-order Commands rerun is
now the next widening gate before returning to the two clean CI passes. That run completed under id
`20260715T185247Z-96492-6c93d0251aff4664babcffb5c6240178`: 799 passed, 7 failed, and 3 skipped in
14m59.5s. Both repaired selection-mask accessibility cases passed. The failures changed again and were
`cmd_preferences_dialog_scroll_host_preserves_retained_page_state` (off-screen/empty host rect),
`cmd_pane_clipboardPasteShortcut_failure_after_navigate_shows_alert` (delayed-navigation expectation),
`cmd_app_shortcuts_long_run_open_close_stays_stable` (search field reopened with leaked `8` input),
`cmd_app_shortcuts_restores_reordered_sorted_grid_layout` (baseline grouped surface unavailable),
`cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics` (two baseline present
failures), `cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable` (null focused HWND), and
`cmd_app_menuBar_mouse_opened_popup_processes_keyboard_before_mouse_move` (could not post `VK_DOWN`). The same
run skipped `cmd_preferences_dialog_category_switches_do_not_churn_tree_host` because the desktop foreground
changed during measurement. Together with the changing failure set and successful isolated reruns, this points
to an unstable foreground/reset boundary in the long Commands process rather than another selection-mask
regression. The shared case/desktop reset and focus-sensitive skip contract are now under inspection before a
third Commands widening run.
A fresh canonical CI attempt is now running with a two-hour outer envelope so the runner's individual-case and
three-seed shuffle classification can complete if the foreground/order instability recurs. It will count only
if the initial CI pass is clean; retry-only success remains blocking under the existing policy.
Live process/manifest inspection shows this attempt's initial Compare pass reported 225 passed, 1 failed, and
31 skipped; the sole failure was again
`search_service_sqlite_delete_burst_maintenance_preserves_query_parity`, timing out after 18.052 seconds with
`lastHr=0x80070002`. Compare classification completed before Commands. Commands then completed without entering
classification after the selection-mask accessibility repair, and the runner advanced to its initial
`--fileops-selftest` process at 21:40 local time. This is strong widening evidence for the repair, but the CI
attempt is already ineligible as a clean pass because Compare required retry classification; the uninterrupted
runner remains active to capture the complete result.
The uninterrupted attempt completed under run id
`20260715T191031Z-85080-eea1492193ea4bf0829f0e65918b6e0b` after 47m02.8s. Its build was clean (0
warnings / 0 errors, `.build/logs/msbuild-20260715_211037_306.log`), and the aggregate was 1,169 passed, 2
failed, and 53 skipped. Commands passed its initial 809-case manifest with 807 passed / 0 failed / 2 opt-in
skips, confirming the selection-mask accessibility repair under full canonical load. File Operations passed
108 / 0 / 20. Compare was classified **FLAKY**: its single compaction-timeout failure passed alone in 4.376s,
then the complete suite passed all three shuffle seeds `1434081272`, `795752390`, and `1684631478` in
193-196s each. `ToolsPesterTests` is the second blocker and classified **REGRESSION** because both its initial
66.327s invocation and 29.643s retry exited 1. Every other CI entry passed its initial attempt. Exact Pester
failure extraction and a deterministic repair for the recurring Compare maintenance wait are now in progress;
this attempt does not count as a clean CI pass.

Direct Pester execution localized `ToolsPesterTests` to four stale inventory assertions in
`Tools/Tests/TestInventory.Tests.ps1`: the source-derived Commands registration count is 706 rather than the
hard-coded 705, and the current tooling Pester count is 287 rather than the hard-coded 262. The same stale
262 count remained in `Tests/README.md`, while the authoritative coverage summary already reported 706/287;
its detailed Commands section still carried 705. The inventory constants and both current documentation
surfaces are now aligned to 706/287. Focused inventory and complete non-toolchain Pester reruns are in
progress; no product behavior or test oracle was weakened.

The first focused inventory rerun passed the Commands/Pester corrections but exposed a fifth stale contract:
CompareDirectories has 249 static registrations while the test constant and the current README/detailed
coverage sections still said 248. Those three current-count surfaces are now aligned to 249; the focused gate
is being repeated from a clean invocation.

The second focused inventory rerun reached the README family reconciliation and found its Settings row still
reported 98 cases while source now contains 99. That final current inventory row is aligned to 99 and the
entire five-case inventory contract is being rerun again.

The third focused rerun confirmed that the numeric reconciliation is green, then the exhaustive documentation
lint reached six Lighthouse source-contract files that were present under `Tools/Tests` but absent from the
README tooling table: HWND render-target resources, modal window shell, packed FileInfo buffer, plugin
configuration, plugin lifetime, and posted-payload coalescing. All six three-case files are now listed with
their enforced ownership/policy boundaries; the inventory rerun continues.

A complete source-versus-table count comparison found the only remaining per-file drifts:
`ThemeDistributionContracts.Tests.ps1` is 9 rather than 7 cases and
`ViewerChromeSourceContracts.Tests.ps1` is 6 rather than 4. Both README rows are aligned; no other tooling
test file is missing or count-mismatched.

The final focused `TestInventory.Tests.ps1` rerun passed 5 / 0 / 0 in 0.874s. This closes the stale inventory
contract exposed by canonical CI. The complete `Tools/Tests` Pester aggregate excluding the single
`RequiresBuildToolchain` case is now running as the widening gate.

The complete non-toolchain tooling aggregate passed 286 / 0 / 0 in 31.552s; the source-derived total of 287
includes the one intentionally excluded `RequiresBuildToolchain` deployment case. `ToolsPesterTests` is no
longer a closeout blocker. The remaining canonical blocker is the recurring broad-order Compare timeout in
`search_service_sqlite_delete_burst_maintenance_preserves_query_parity`; its maintenance/idle-window
synchronization is now under deterministic root-cause analysis.

Root cause is confirmed in the test harness rather than product compaction. The delete-burst case called
`ForegroundSearchServiceProcess::Start(..., waitForReady=true)`, whose successful readiness status request
can release the already queued scheduler into automatic maintenance. While that bounded compaction runs, the
single-threaded service intentionally has no listening pipe; the case immediately issued another status
request but allowed only 15 seconds, half the shared 30-second cold-service readiness budget. Under focused
load the compaction finished in roughly four seconds, while broad initial order could keep the pipe absent
past 15 seconds and reported `0x80070002`. The case now launches without the redundant generic readiness
request, lets its own predicate observe either queued state or an already completed first window, uses the
shared 30-second budget for maintenance-window polling, captures service output, and fails early with process
diagnostics if the child exits. A Pester source contract pins the launch mode, budget, and diagnostics; focused
RED/GREEN validation is in progress.

The complete `TestHarnessSourceContracts.Tests.ps1` file passes 121 / 0 / 0 in 2.071s with the new
delete-burst launch/budget/diagnostic contract. A Debug x64 `RedSalamander` rebuild is now running before
runtime repeat and canonical-order Compare widening.

The current-tree Debug x64 `RedSalamander` build passed with 0 warnings / 0 errors in
`.build/logs/msbuild-20260715_220810_245.log`. The repaired delete-burst case is now running ten native
in-process repetitions before a complete canonical-order Compare run.

The repaired case passed 10 / 0 / 0 native in-process repetitions in 48.0s under run id
`20260715T201036Z-45744-50553b7f278a453b8d100f5862b756dc`. The complete 256-case Compare manifest is now
running in canonical order, matching the context that repeatedly exposed the old 15-second missing-pipe
failure.

The complete canonical-order Compare widening passed on its initial attempt with 226 passed / 0 failed / 31
expected environment skips in 3m19.9s, run id
`20260715T201146Z-102468-f049ffb2ddb243dd94e6ddc5a0dab592`. No retry, shuffle triage, or classifier was
entered. Together with the ten-repeat focused run, this closes the recurring delete-burst maintenance timeout
as a deterministic harness repair. Lighthouse closeout reconciliation and the first clean canonical CI pass
are next.

Canonical CI closeout pass 1 (`20260715T201606Z-68024-f5faab4d8bcb485690f63065fb8beb21`) completed in
44m28.6s with 1,169 passed / 2 failed / 53 skipped and therefore does **not** count as clean. The repaired
Compare suite, File Operations, all standalone entries, and `ToolsPesterTests` passed their initial attempts.
Commands failed `cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup` and
`cmd_app_menuBar_submenu_placement_matches_spec`; both passed their isolated reruns, but shuffle seeds
`884396155`, `1652104485`, and `1380075548` each reproduced a Commands failure, so the runner correctly
classified the suite as a blocking order regression rather than flaky. Exact shuffle case extraction and the
shared menu/pointer reset boundary are now under diagnosis; CI pass 2 and Full remain blocked until this is
deterministically repaired.

Standalone replay of shuffle seed `884396155` reproduced a different order failure after 32 passes:
`folderView_thumbnail_sort_popup_slider` reported that the pane bottom-right sort popup did not expose the
right-aligned Ctrl+F2 through Ctrl+F6 shortcuts (run
`20260715T210455Z-95456-74fa8a6af672435b98179f36545f0f00`). Thus the CI classifier did not merely relabel the
two initial menu failures: Commands has a broader cross-case popup/menu state leak. The shared popup dismissal,
cursor/synthetic-pointer routing, and retained menu-model reset boundaries are being compared across the
initial failures and this seed before replaying the remaining seeds.

Standalone replay of shuffle seed `1652104485` also reproduced the order regression, this time after 75
passing cases: `cmd_app_menuBar_hover_switches_top_level_popup` failed because the temporary keyboard-opened
menu bar did not keep the expected first top-level menu selected. Run
`20260715T210933Z-85820-86b39a81c6154e8cb264011d1b3fba7f` completed with 75 passed / 1 failed / 733
fail-fast skips. The immediately preceding cases were unrelated dialogs, settings, clipboard, viewer-text,
properties, theme-row, and shortcuts tests, further supporting a shared foreground/menu reset leak rather than
a direct dependency on one neighboring menu case. Replay of seed `1380075548` was intentionally stopped when
the operation was paused; it has no completion evidence and must be rerun from the beginning.

**Saved continuation point (2026-07-15; superseded 2026-07-16):** this checkpoint originally required three
shuffle seeds, two clean CI passes, and a clean Full pass. Subsequent maintainer review reduced that duplicated
validation to the proportional closeout policy recorded below, and the common explicit-order isolation plus
focused menu/pane teardown repairs were completed without weakening their behavioral oracles. The final build,
Full attempt, external-blocker classification, authoritative-spec reconciliation, and Done decision later in
this ledger replace the old resume instructions. `Operation_Observatory_...2026-07-15.md` remains separate
concurrent user work and was not edited or staged by Lighthouse.

The durable anti-regression contract is now explicit rather than being distributed only across individual
domain specs and Pester guards. `Specs/Core/Core_SharedHelpers.md` is the authoritative canonical-owner catalog
for the Lighthouse extractions, including `ColorFromArgb`, conversion/path/I/O policies, viewer/app primitives,
plugin lifecycle/configuration/packed buffers, posted payloads, and test infrastructure. `AGENTS.md` now makes
reuse-or-extension mandatory and requires any intentional local variant to name and test its semantic or
dependency difference. The `cpp-modern-style` skill links the same pre-implementation check, and
`Specs/README.md` exposes the catalog in the Core start-here list. Existing source-contract tests remain the
mechanical guards for the high-risk migrated families.

Closeout resumed on 2026-07-16. The first attempt to replay shuffle seed `1380075548` was correctly rejected
before test execution because the intentionally interrupted prior run left
`.build/artifact-operation-contaminated.json`. Per the shared build/test ownership contract, a serialized
`build.ps1 -Rebuild` is required before any existing binary can provide trustworthy evidence; the replay and
all later gates remain pending that clean rebuild.

The required x64 Debug full rebuild then completed successfully in 2m16s with 0 warnings and 0 errors
(`.build/logs/msbuild-20260716_082005_282.log`). The contamination marker is cleared, so the saved
`1380075548` shuffle replay can now provide valid diagnostic evidence.

Replay of the final recorded shuffle seed `1380075548` completed under run
`20260716T062234Z-67412-e22152442f5f4be5b6a29ce19023bede` and failed after 36 passing cases:
`settings_uia_blocked_operation_deadline_is_bounded` reported that the deliberately blocked UIA operation was
not classified as timed out (36 passed / 1 failed / 772 fail-fast skips). Because all three saved seeds fail in
different UI/popup/UIA cases, the closeout blocker is now characterized as a broader Commands cross-case
foreground/worker-state isolation defect, not a single menu implementation regression. The exact case and its
preceding order are under focused inspection before changing shared reset behavior.

Inspection narrowed the third-seed failure to its own timeout-multiplier arithmetic rather than foreground
state: the case scaled an 800 ms sleep to 2400 ms, while the helper's 3x total budget used a capped 500 ms
cancellation reserve and therefore allowed a 2500 ms operation window. The supposedly blocked action completed
100 ms before that window. The test now derives its blocking duration from `MakeUiaDeadlineBudget`, placing
completion halfway inside the actual cancellation reserve for every supported multiplier; the Pester source
contract pins that relationship. Focused normal/3x executions and the saved seed replay are pending rebuild.

The focused RedSalamander build completed successfully with no warning/error diagnostics in
`.build/logs/msbuild-20260716_082546_605.log`. Validation has advanced to the focused Pester contract and exact
native UIA case at both default and 3x timeout multipliers.

Focused validation is green: the complete `TestHarnessSourceContracts.Tests.ps1` manifest passed 121 / 0 / 0
in 3.66s; the exact UIA deadline case passed at the default multiplier under run
`20260716T063719Z-20236-01d84b0613034e408b75d57f57e34236` and at 3x under run
`20260716T063727Z-70644-157991b06b3c4a2eacfd489ba39af73a`. Seed `1380075548` is now being replayed from
the beginning to prove the fix in its original order context.

The repaired seed advanced from 36 to 48 passing cases, then failed
`cmd_pane_filter_prompt_uses_dxui_surface` because the initial prompt contained no history (run
`20260716T063747Z-47956-6c394a6e74714bf09f339c7b63b67117`). The test asserted a non-empty history but
relied on canonical-order predecessors to populate `g_settings.selectionMasks.filterHistory`. It now seeds a
deterministic two-entry history and restores the prior optional settings at case exit; the existing broad-order
dialog source contract pins both setup and restoration. Focused validation and another seed continuation are
pending rebuild.

The filter-history repair built cleanly with 0 warnings/errors
(`.build/logs/msbuild-20260716_084048_127.log`). Its exact 3x case passed under run
`20260716T064308Z-72240-484eb8e1a28943ccb72f28f521de449a`, and the complete test-harness source-contract
manifest again passed 121 / 0 / 0 in 2.25s. Seed `1380075548` is continuing from a fresh invocation.

The seed next advanced to 72 passing cases and failed `folderView_perf_huge_folder_scale`: after quick-search
activation, synthetic `WM_CHAR('s')` left the query empty (run
`20260716T064340Z-77212-d9888cfbbcc44bf3a47de9d4ed5834ab`). The exact 3x case passed alone under run
`20260716T064758Z-73900-d541e6daec9049ecbe98f6fe7856475c`, confirming an inherited input-state dependency.
`FolderView::OnCharMessage` rejects input only when Ctrl/Alt is down before appending; the perf helper previously
silently returned when keyboard-state capture failed and ignored keyboard-state apply failure. Modifier clearing
is now verified, `GetKeyboardState` failure falls back to the already zeroed test-process state, every quick-search
send reports failure explicitly, and the existing quick-search source contract pins the verified boundary.

The verified modifier repair built with 0 warnings/errors
(`.build/logs/msbuild-20260716_084850_110.log`). The exact huge-folder 3x case passed under run
`20260716T065107Z-81668-7bf946069f1e4d8d8d45093ca499a468`, and the complete harness source contract passed
121 / 0 / 0 in 2.53s. The original seed is being rerun again; it must complete all 809 registered cases before
this order is accepted as clean.

The next replay of seed `1380075548` advanced to 464 passing cases before `modeless_window_ownership` failed
because the Shortcuts window did not open (run
`20260716T065154Z-75384-899ba7b57564460c93b64f97de5efc22`: 464 passed / 1 failed / 344 fail-fast skips).
The Preferences and Connection Manager ownership checks immediately before it succeeded, and the Shortcuts
command only creates the window when `g_settings.shortcuts` is present. The case currently consumes inherited
settings and modeless-window state instead of establishing and restoring its own Shortcuts fixture; that
order-isolation boundary is now under focused repair before the seed is replayed again.

The modeless ownership case passed an isolated pre-repair 3x replay under run
`20260716T071111Z-86940-9ec0b3c427124456a3ea5275de92588d`, confirming order dependence. It now closes
inherited non-main top-level windows through the shared isolated-UI preparation boundary, seeds and reloads
default Shortcuts settings before command dispatch, then closes any test-created Shortcuts window and restores
the exact prior optional settings. The scoped x64 Debug build passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_091214_352.log`); the repaired case passed 3x under run
`20260716T071503Z-77156-2f501531405847a5a84c63cc6de0216c`, and the complete test-harness source-contract
manifest passed 121 / 0 / 0 in 2.7s. Seed `1380075548` must now be replayed from the beginning again.

That replay advanced beyond the repaired modeless boundary to 509 passing cases, then
`cmd_pane_find_dialog_restores_reordered_sorted_grid_layout` failed with zero visible results after an index
search completed as `service unavailable` (run
`20260716T071608Z-76948-086b89b284184d8ba4ffd254785d7360`: 509 passed / 1 failed / 299 fail-fast skips,
18m14s). The case creates three local files and validates grid reorder/sort persistence, but it currently inherits
the default/index preference instead of choosing a deterministic local scan backend. Focused isolation and the
neighboring Find fixture contracts are now being checked before changing that backend setup.

The Find case passed an isolated pre-repair 3x replay under run
`20260716T073512Z-89268-ba5854ee1fe34159a7c2fe1a1a586357`, confirming inherited-order state. Its fixture
now starts from a fresh `SearchDialogSettings`, explicitly selects local scan (`preferIndex=false`), includes
files, excludes directories, and fixes the remaining traversal/matching/result-limit inputs before exercising
the unchanged reorder/sort oracle. The scoped x64 Debug build passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_093640_566.log`); the repaired case passed 3x under run
`20260716T073902Z-94676-803bda0327c94592b889571ab4ffa695`, and the complete harness source-contract
manifest passed 121 / 0 / 0 in 3.7s. Seed `1380075548` is ready for another complete replay.

The next replay stopped earlier after 280 passes at
`cmd_preferences_dialog_category_tree_uses_dxui_host_without_visible_legacy_treeview`: the fresh Preferences
dialog's logical baseline was General, but a new UIA client still observed the prior destroyed host's selected
plugin child, `OneDrive Business` (run
`20260716T073920Z-92988-828e6e96a6b84996903519c7e101028b`: 280 passed / 1 failed / 528 fail-fast skips,
7m34s). The case passed 3x alone under run `20260716T074834Z-83644-fd1bc82821f0465a9bf949fcd2de6af6`.
Source review found that DxUi publishes an empty accessibility snapshot and unregisters its HWND target but does
not issue the platform-required destroyed-window notification after returning raw providers. The teardown path
is being repaired to release the target mutex and then call
`UiaReturnRawElementProvider(hwnd, 0, 0, nullptr)`; the native Accessibility contract and authoritative DxUi
spec will pin the lifetime boundary before the seed is replayed.

The DxUi teardown repair is now implemented: it publishes the empty snapshot and removes the host target while
holding the accessibility mutex, releases that mutex, and then notifies UI Automation with
`UiaReturnRawElementProvider(hwnd, 0, 0, nullptr)` so a recycled HWND cannot retain the destroyed host's raw
provider. `Tests/DxUiTests/DxUiTests.Accessibility.cpp` pins the notification ordering, and
`Specs/UI/UI_DxUiWinUIDesign.md` now owns the durable lifetime rule. The focused DxUi Accessibility suite passed.
An initially targeted application build exceeded its command timeout and set the repository artifact
contamination marker; no result from that attempt is accepted. The required full-solution x64 Debug rebuild
then passed with 0 warnings/errors (`.build/logs/msbuild-20260716_095531_517.log`) and cleared the marker. The
focused Preferences category-tree case passed 3x under run
`20260716T075826Z-82388-40bbe731d9c842168cc9104b6d37cba0`, and the complete harness source-contract
manifest passed 121 / 0 / 0 in 2.1s. Seed `1380075548` is ready for its next complete replay.

That replay advanced through the repaired Preferences boundary but stopped after 72 passes at
`folderView_perf_huge_folder_scale` because its synthetic modifier-clear helper returned false (run
`20260716T075901Z-5576-304c43a05ff143c9889eb1e42ebb7c68`: 72 passed / 1 failed / 736 fail-fast skips,
2m14s). The same seed had previously passed this case, and the immediately relevant keyboard-capture plus
huge-folder pair passed together under run `20260716T080223Z-94220-f3519bf41eca467b81481fe452d1abeb`.
Microsoft's Win32 contract says `SetKeyboardState`, `GetKeyboardState`, and `GetKeyState` access the calling
thread's same input-state table; the helper is being instrumented to distinguish an API failure from a
post-set Ctrl/Alt observation before selecting a repair.

The instrumented rebuild passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_100330_067.log`), but the next replay exposed the upstream failure before it
reached the modifier helper: `cmd_preferences_dialog_keyboard_capture_preview_and_assign_live_dx_interaction`
could not invoke the reopened page's Assign action after its own shell-Cancel cycle (run
`20260716T080543Z-43904-771f31939afe4ae9ac9c23c07830fc40`: 41 passed / 1 failed / 767 fail-fast skips,
1m32s). Platform documentation recommends retiring the raw-provider HWND map from `WM_DESTROY`; the first
repair only reached the notification through DxUi's later `WM_NCDESTROY` detachment. The provider-map
notification is being added at `WM_DESTROY`, while unregister retains the same call as the fallback for owners
that explicitly detach before destruction.

That destruction-boundary repair is implemented and pinned by the native Accessibility source contract; the
authoritative DxUi spec now requires the `WM_DESTROY` notification plus the explicit-detach fallback. The
scoped DxUiTests and application builds passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_100915_398.log` and
`.build/logs/msbuild-20260716_101131_490.log`), the focused DxUi Accessibility suite passed, the capture-preview
case passed five in-process repeats under run
`20260716T081350Z-93280-a2529e9db59b477c8cd2d6e6cda2c7e0`, and the complete harness source-contract
manifest passed 121 / 0 / 0 in 2.1s. Seed `1380075548` is ready for another full replay.

The next replay stopped after 22 passes at
`cmd_compare_directories_options_theme_cycle_keeps_surface_legible` because its generic UIA edit lookup could
not resolve the visible Ignore-files edit even though the preceding pattern census observed ValuePattern
support (run `20260716T081419Z-81928-ad9da2b5e1974d9789f3307eb2c88fbe`: 22 passed / 1 failed / 786 fail-fast
skips, 1m9s). A sibling Compare Options test already used the stable accessible name, direct raw-provider
ValuePattern state, bounded message-pumped UIA fallback, and detailed diagnostics for the same control. That
logic is being extracted once at file scope and reused by both the live-interaction and theme-cycle cases; the
control identity and ValuePattern oracle remain unchanged.

The shared named-edit helper is now in place and the theme-cycle contract rejects regression to the generic
edit lookup. The scoped application build passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_101736_624.log`), the repaired theme-cycle case passed five in-process repeats
under run `20260716T081952Z-63064-6dad29e4d00a4e0281e510f80b84a4e9`, and the complete harness
source-contract manifest passed 121 / 0 / 0 in 2.3s. Seed `1380075548` is ready for another replay.

After a concurrent Release x64 solution build released the shared artifact lock, the replay advanced through
all earlier repaired boundaries to 125 passes and then stopped at
`cmd_pane_createDirectory_prompt_long_run_open_close_stays_stable`: cycle 8 captured the prompt before its
initial full-text selection had settled (run `20260716T082529Z-76156-e4aeedd6a3a84908acfbaa5657a5890f`).
The same Dialogs family already polled the create-directory prompt's DxUi host, path, text, focus, selection,
and validation state for unique-name coverage, but only to a weaker local readiness condition. That condition
is being consolidated into one named helper that waits for the full selected-text contract and is reused by
both unique-default and long-run churn cases.

The shared selected-text readiness helper is implemented and source-contract guarded. The scoped application
build passed with 0 warnings/errors (`.build/logs/msbuild-20260716_104645_552.log`), the repaired 12-cycle
churn case passed five in-process repeats under run
`20260716T084914Z-95440-e999cc3c576146fa8c4498319b1ac2a2`, and the complete harness source-contract
manifest passed 121 / 0 / 0 in 3.8s. Seed `1380075548` is ready for its next replay.

That replay stopped after three passes at `folderView_perf_sort_toggle_stress`: the first Name/Ascending
transition used a one-shot warm-render probe while the FolderView could still have a zero client area or be
recreating its Direct2D target (run `20260716T084954Z-75116-73483d6fa693433fabc5236b054fa0db`: 3 passed /
1 failed / 805 fail-fast skips, 39.2s). The product sort assertions remained valid and the same case had passed
in earlier identical-seed runs, so this is a fixture-readiness race. The case is being changed to wait within a
bounded deadline for the requested sort, direction, unchanged 5,000-item population, valid window, and a
successful warm render. Its existing layout, inactive-search-effect, render-count, elapsed-time, and artifact
oracles remain intact.

The bounded sort-ready warm-render wait is now implemented and source-contract guarded. An interrupted
focused build correctly tripped the shared-artifact contamination guard; the required full x64 Debug solution
rebuild then recovered the outputs and passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_105508_999.log`). The repaired sort-toggle case passed five in-process repeats
under run `20260716T085808Z-75628-0be35f7c48da40cab5fbb2d9616f1824`, and the complete harness
source-contract manifest passed 121 / 0 / 0 in 2.44s. Seed `1380075548` is ready for another full replay.

That replay advanced through 195 cases before
`cmd_app_compare_keeps_navigation_shell_stable` timed out waiting for the left FolderView to regain Win32
focus after the Compare window closed (run
`20260716T085913Z-53632-fe950778a3c24a509924d55292119a3f`: 195 passed / 1 failed / 613 fail-fast
skips, 6m2s). The captured navigation state, pane path, history, refresh count, 2-item population, selection,
and focused `b.log` model item were all correct; the fixture's diagnostic omitted the actual/expected focus
HWNDs, so the remaining condition is being reproduced in isolation before choosing between a fixture focus
reset and a product close-boundary repair.

The exact App Compare navigation-shell case then passed five in-process repeats under run
`20260716T103537Z-65576-d8576b8f89244e8dad311543bf878152`, classifying the failure as inherited UI/focus
state rather than a standalone product close defect. Unlike the already-stabilized focus-sensitive fixtures,
both equivalent Compare navigation-shell cases bypassed the shared helper that closes inherited non-main
top-level windows and reactivates the main window, and neither required a stable FolderView focus precondition
before opening Compare. Both fixtures are being aligned with those existing isolation contracts; the App case
also gains actual/expected folder-view, Win32-focus, and foreground HWND diagnostics without weakening its
post-close model, path, refresh, selection, or focus oracle.

The shared isolation and bounded pane-focus preconditions are now implemented and guarded by the harness
source-contract manifest (122 / 0 / 0 in 3.44s). The scoped application build passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_123932_023.log`), the Navigation-family registration passed five repeats under
run `20260716T104218Z-92616-9f8370b7bd5d493c8db9aad009c83c60`, and the App-family registration passed five
repeats under run `20260716T104221Z-92616-fa532709b50f47b495e280c101aaa794`. Seed `1380075548` is ready for
another full replay.

The next replay stopped after 72 passes at the already-instrumented
`folderView_perf_huge_folder_scale` Quick Search activation boundary (run
`20260716T104242Z-90672-48da411c1980408db2683912304d3306`: 72 passed / 1 failed / 736 fail-fast skips,
2m7s). The helper successfully cleared synthetic Ctrl/Alt state and the case had established stable FolderView
focus, but `CommandQuickSearch` was followed by an inactive empty search snapshot. The immediate predecessor
was a Preferences Viewers Reset interaction; the perf case currently does not use the shared inherited
top-level cleanup/main-window reactivation step. An isolated repeat is being used to confirm that this is
queued UI/focus state rather than a standalone Quick Search activation defect before applying that existing
isolation contract.

The huge-folder case passed five isolated repeats under run
`20260716T104632Z-74456-1ffebcdbe9e241588867b8a619c6825b`, confirming the Quick Search product path is
stable without inherited UI state. The fixture now invokes the shared non-main-window cleanup/main-window
reactivation helper before loading the dummy provider or starting any measurement, and its activation failure
reports actual/expected folder-view, Win32-focus, and foreground HWNDs. The performance scenario and budgets
remain unchanged.

The harness source-contract manifest passed 122 / 0 / 0 in 2.33s, the scoped application build passed with
0 warnings/errors (`.build/logs/msbuild-20260716_124832_516.log`), and the repaired huge-folder case passed
five further repeats under run `20260716T105050Z-50376-6b75c0c3793943db80751c30daa8e8ce`. Seed
`1380075548` is ready for its next full replay.

That replay stopped after 17 passes before exercising any Compare Directories tail toggle because
`cmd_preferences_dialog_compare_directories_tail_toggles_live_dx_interaction` asserted that
`SetFocus(treeHost)` returned `treeHost` (run
`20260716T105210Z-12628-d8cdb73f44354e4cbf62f057cccce438`: 17 passed / 1 failed / 791 fail-fast skips,
57s). Win32 returns the previous focus HWND, so this assertion only passes when the target happened to be
focused already. A census found 141 copies of the same invalid comparison across the six Preferences Commands
self-test chunks, despite the existing bounded `FocusWindowAndWait(...)` helper and the authoritative self-test
contract requiring observation of resulting focus. All 141 copies are being migrated mechanically to that
helper with a 1-second scaled bound, and the harness source contract will reject any recurrence across the
entire six-file family.

The 141-call migration is complete. A whole-family search finds no remaining `SetFocus(target) == target`
comparison, the harness source-contract manifest passed 122 / 0 / 0 in 2.33s, the scoped application build
passed with 0 warnings/errors (`.build/logs/msbuild-20260716_125455_814.log`), and the originally failing
tail-toggle case passed five repeats under run `20260716T105714Z-82080-aeda5b6edfc44fc98c973c1dd4ee416a`.
Seed `1380075548` is ready for another full replay.

The next replay advanced to 83 passes before
`cmd_app_menuBar_mouse_opened_popup_processes_keyboard_before_mouse_move` failed while posting `VK_DOWN` to
the mouse-opened DxUi popup (run `20260716T105735Z-75404-f7eb8cfb28034e34b4a4c8fafdb23d0f`: 83 passed /
1 failed / 725 fail-fast skips, 2m24s). The popup had opened and exposed its initial readable debug state, but
the fixture records the combined key-down/key-up post only as a boolean, so it cannot distinguish a stale
popup handle, a key-down failure, or closure between key-down and key-up. Its immediate predecessor was a
Find dialog live-UIA case and this menu fixture does not use the shared inherited-top-level cleanup/main-window
reactivation helper. The exact case is being repeated before selecting the repair; any change will preserve the
keyboard-before-mouse ordering and repaint oracle.

The exact mouse-opened popup case passed five repeats under run
`20260716T110039Z-50500-554f1427c5c246a68437466a4154348a`, classifying the failure as inherited UI/focus
state. The fixture now uses the shared non-main-window cleanup/main-window reactivation helper before opening
the menu, and records target validity plus key-down and key-up post results separately. The popup initial-state,
keyboard-index, render-count, no-mouse-move, and dismissal assertions remain unchanged.
The source-contract manifest passed 123 / 0 / 0, the scoped application build passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_130358_891.log`), and the repaired exact case passed five repeats under run
`20260716T110623Z-2012-21de83cca873406c819c554008f88815`. Seed `1380075548` is ready for another full
replay.

The next seed replay stopped after four passes at
`cmd_pane_find_dialog_action_buttons_activate_expected_commands` (run
`20260716T110645Z-50604-1d55081c5ad842b884fb010281baf9a6`: 4 passed / 1 failed / 804 fail-fast skips,
58.3s). The Find result Open button delivered its click, but the focused pane remained at the fixture root
instead of navigating into the selected `sub` directory. This exact case is being repeated before changing
the fixture; the selected-result identity, delivered-click proof, active-pane identity, and final navigation
oracle remain required.

The exact Find action-button case passed five repeats under run
`20260716T110804Z-69448-094714a76f8e435bac16a01944bc6134`, classifying the failure as inherited
top-level/input state. The fixture now runs the shared non-main-window cleanup/main-window reactivation barrier
before capturing or changing pane state. Its live button clicks, selected-result checks, active-pane selection,
and Open/Parent navigation assertions remain unchanged; the source-contract manifest now guards the isolation
barrier.
The source-contract manifest passed 123 / 0 / 0, the scoped application build passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_131017_103.log`), and the repaired exact case passed five repeats under run
`20260716T111237Z-31276-ed79e75d641c42e0ae210ee05d4a4234`. Seed `1380075548` is ready for another full
replay.

The next replay advanced to 511 passes before `cmd_app_showShortcuts_keeps_navigation_shell_stable` could not
open the Shortcuts window (run `20260716T111305Z-79356-838f5d7c32c5463999762da3653e43dc`: 511 passed /
1 failed / 297 fail-fast skips, 18m17s). The exact case is being repeated before selecting the repair. The
fixture must continue to prove that the Shortcuts window opens from the real app command, closes cleanly, and
preserves both navigation shells, pane identity, and focused-folder-view state.

The exact app Shortcuts shell-stability case passed five repeats under run
`20260716T113139Z-38248-578becc51ec54987ac091a5df14ae583`. The real command intentionally opens no
window when `g_settings.shortcuts` is absent, but this fixture did not own that optional precondition; its result
therefore depended on prior cases. The fixture now saves the optional shortcut settings, installs and reloads
canonical defaults, runs the shared top-level/main-window isolation barrier, and restores both settings and any
opened Shortcuts window on every exit. Its real `WM_COMMAND`, UIA-provider, child-window, close, navigation-shell,
pane, and focus assertions are unchanged, and the source-contract manifest guards the new isolation contract.
The source-contract manifest passed 124 / 0 / 0, the scoped application build passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_133256_413.log`), and the repaired exact case passed five repeats under run
`20260716T113512Z-92044-cebd243ce84047cf84ad2f822cf82bc6`. Seed `1380075548` is ready for another full
replay.

The next replay advanced to 556 passes before
`cmd_preferences_dialog_keyboard_reset_live_dx_interaction` could not expose live UIA InvokePattern on the
Preferences shell Cancel action during reset-discard validation (run
`20260716T113538Z-67432-3760c1601cfa4c71991e31ce9551513e`: 556 passed / 1 failed / 252 fail-fast skips,
19m02s). The exact case is being repeated before selecting the repair. The reset-state change, discard through
the live Cancel action, restored shortcut model, and dialog teardown assertions remain required.

The exact Keyboard reset live-interaction case passed five repeats under run
`20260716T115455Z-56888-9eb7c755cb9f456c9d451ee96de3ad60`, including repeated Preferences host
recreation against the same UIA client context. This classifies the seed failure as inherited top-level/input
state rather than a stale UIA-client cache. The fixture now runs the shared non-main-window cleanup/main-window
reactivation barrier before opening Preferences. The real Reset Defaults and shell Cancel InvokePattern calls,
filtered-row shortcut assertions, discard/reopen checks, and teardown waits remain unchanged; a source-contract
test guards the isolation boundary.
The source-contract manifest passed 125 / 0 / 0, the scoped application build passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_135709_583.log`), and the repaired exact case passed five repeats under run
`20260716T115928Z-91800-d3d9ae6eec604420b8050c6f32b6c5f1`. Seed `1380075548` is ready for another full
replay.

The next replay stopped after 48 passes when `cmd_pane_filter_prompt_uses_dxui_surface` could not capture the
expanded pane-filter help snapshot during the empty-confirm pass (run
`20260716T115954Z-18360-d86cd0acae9b47a79215f4719417a4d8`: 48 passed / 1 failed / 760 fail-fast skips,
1m18s). The exact case is being repeated before selecting the repair. The empty-confirm result, real DxUi
prompt surface, help expansion/content, filter application, and cancel/teardown assertions remain required.

The exact pane-filter DxUi surface case passed five repeats under run
`20260716T120144Z-94348-626153dbb2094024bf7df3578c9db716`. The cycle worker captures the intended prompt
HWND, but the existing debug expand/snapshot helpers rediscover a same-process prompt globally by class; an
inherited top-level can therefore redirect the immediate expanded-help snapshot. The fixture now runs the
shared non-main-window cleanup/main-window reactivation barrier before any prompt cycle. The three real modal
cycles, ownership/UIA checks, help expansion, filter mutation, confirm/cancel outcomes, and teardown assertions
remain unchanged; a source-contract test guards the isolation boundary.
The source-contract manifest passed 126 / 0 / 0, the scoped application build passed with 0 warnings/errors
(`.build/logs/msbuild-20260716_140256_780.log`), and the repaired exact case passed five repeats under run
`20260716T120509Z-72684-85405a91a7d9468ca0eb24801d33c4b0`. Seed `1380075548` is ready for another full
replay.

The next replay advanced to 111 passes before
`cmd_pane_navigationView_full_path_popup_ancestor_click_navigates_to_ancestor` could not expose a clickable
full-path popup ancestor segment (run `20260716T120603Z-34972-18f2b213ed4b4bb9b8ca813dd5a4a5c5`: 111 passed /
1 failed / 697 fail-fast skips, 3m18s). The exact case is being repeated before selecting between a local fixture
repair and a Commands-wide per-case UI isolation boundary. The popup opening, ancestor hit target, navigation,
focus restoration, and clean teardown assertions remain required.

The exact full-path-popup ancestor-click case passed five repeats under run
`20260716T120940Z-84312-24d9a9a865b6495aaeda122f1479f1ce`, confirming another inherited-order failure.
The repeated pattern now warrants one systematic boundary instead of more fixture-local cleanup: Commands
seeded/repeated execution already dispatches one exact case at a time, so every executing case now runs the
shared non-main-window cleanup/main-window enabled/reactivation barrier immediately before dispatch. If that
barrier fails, the harness records the failure against the intended case and does not run its fixture against
dirty UI state; after fail-fast has recorded a failure, remaining skipped cases avoid redundant cleanup. The
authoritative self-test specification and source-contract manifest now require this boundary. Existing local
fixture barriers remain useful for non-explicit canonical execution and for clearer fixture diagnostics.

The Commands-wide isolation boundary passed the source-contract manifest (127 passed / 0 failed / 0 skipped),
the scoped application build completed with 0 warnings/errors
(`.build/logs/msbuild-20260716_141226_440.log`), and the previously shuffle-only full-path-popup case passed five
repeated executions from the rebuilt binary under run
`20260716T121553Z-77120-fb19ce5bb43b4bd2b565b9dff03b858b`. Replay of seed `1380075548` is now the first
remaining closeout gate.

That replay advanced through 437 passing cases before
`cmd_app_menuBar_submenu_placement_matches_spec` failed to open its second DxUi popup (run
`20260716T121632Z-18960-aacf30c82724406e83e5407bd3b7bc5c`: 437 passed / 1 failed / 371 fail-fast skips,
14m30s). The per-case boundary removed inherited non-main top-level windows, so this remaining failure is being
classified as menu-fixture state rather than weakened or waived. The root popup, cascading submenu, attachment
geometry, parent-return behavior, Escape closure, and temporary menu-bar teardown assertions remain required.
The exact case then passed five repeats under run
`20260716T123124Z-24488-5a86425f29e14e25bff55761de7bcba1`, confirming an order-only retained menu/input
dependency. Diagnosis is now focused on resetting the main window's internal menu-selection/input state and on
making the cascade-target acquisition report its observed keyboard path before the three seed gates resume.

The fixture review found a narrower deterministic defect: it selected the cascade child from the native menu
before opening the root, while the product runs `OnInitMenuPopup` during open and may refresh dynamic entries.
The repair now reactivates the directed-input target, requires the main window to own the foreground, reacquires
the first enabled cascade index from the initialized root menu, and reports the pre/post target, item count,
popup validity, observed keyboard/hover state, and separate key-message delivery results on failure. The shared
helper used to discover cascade items now serves both the pre-open top-level selection and post-open refresh.
The authoritative self-test contract and Pester source contract have been updated. The source-contract manifest
passed 128 / 0 / 0, the scoped application build completed with 0 warnings/errors
(`.build/logs/msbuild-20260716_143832_181.log`), and the repaired exact case passed five repeats under run
`20260716T124051Z-10472-0bf41c23238c48cdb80dc8103ab8b47b`. Seed `1380075548` is being replayed from the
beginning against this binary.

The replay passed the repaired submenu case but stopped after 109 passes at
`cmd_pane_navigation_rename_prompt_keeps_navigation_shell_stable` (run
`20260716T124122Z-7432-be8974ede28c4e12bb4b6985955df135`: 109 passed / 1 failed / 699 fail-fast skips,
3m22s). Rename completed, the refreshed item/path/navigation model was correct, and the renamed item remained
focused in the model, but no FolderView owned Win32 focus after confirm. The preserved trace identifies
`cmd_app_swapPanes_keeps_navigation_shell_stable` as the immediate predecessor. The exact rename case then
passed five repeats under run `20260716T124615Z-41168-f157ac9ba747439b99c4e94f7e6a525d`, confirming order
leakage. The predecessor restores both pane plugins and paths from a scope-exit without waiting for the
asynchronous restoration enumerations to complete, allowing that work to overlap the next fixture. Its teardown
is being changed to wait for both restored paths/enumerations and the original active-pane focus before return;
the rename confirm focus oracle remains unchanged.

The Swap Panes fixture now keeps scope-exit restoration only as an early-failure fallback. Its successful path
restores both providers, waits for both original paths and their matching enumeration-complete callbacks, then
restores and verifies the original active-pane FolderView focus before returning. The authoritative self-test
contract and source-contract manifest require this handoff; the manifest passed 128 / 0 / 0, and the scoped build
completed with 0 warnings/errors (`.build/logs/msbuild-20260716_144817_152.log`). The focused two-case shuffle
ran Swap Panes immediately before rename and passed both cases under run
`20260716T125048Z-29304-d3a5108da3f84d3886d4528d83c496e8`. Seed `1380075548` is ready for another full replay.

That replay advanced to 200 passes before the unrelated
`cmd_preferences_dialog_general_live_dx_interaction` retained-category dependency surfaced (run
`20260716T125111Z-89400-a6e5fdf240ff4c98bab58e124ccb2d9b`: 200 passed / 1 failed / 608 fail-fast skips,
7m25s). The exact case passed five repeats under run
`20260716T125919Z-56440-198ae26f17584efbb61f9620a904850c`. Maintainer review classified this pre-existing
Preferences shuffle-order dependency as outside the shipped Lighthouse helper/product changes. Its experimental
repair was dropped, and the diagnostic shuffle seeds are no longer Lighthouse closeout gates. The observation is
retained here as a non-blocking test-reliability follow-up rather than expanding this operation again.

Closeout validation was reviewed with the maintainer on 2026-07-16 and reduced to non-duplicative evidence:
one clean current-head x64 Debug build and one clean Full suite. The diagnostic Commands shuffle seeds are useful
for future test-reliability work but are not Lighthouse gates. An additional architecture build is required only
if the final repair changes architecture-sensitive code; two separate CI suite passes are not required because
Full already subsumes that coverage.

The converged closeout candidate passed the complete harness source-contract manifest (128 passed / 0 failed /
0 skipped in 2.47s). The first scoped build command was interrupted by its two-minute command window while the
main application was still compiling, so the repository's shared-artifact contamination guard correctly
required a full recovery rebuild. That current-head x64 Debug solution rebuild passed with 0 warnings/errors in
2m44s (`.build/logs/msbuild-20260716_150845_700.log`) and serves as the single required clean build.

The one Full run against those recovered artifacts completed under run
`20260716T131155Z-7948-bdf657ae03db41d0b7f87ad6a08317cd` in 38m08s with 1,151 passed / 3 failed / 54
skipped. All 809 Commands cases passed, including the Lighthouse explicit-order isolation and teardown repairs.
One failure was Lighthouse bookkeeping: the source-derived inventory had not yet ratcheted the added Commands
orchestrator registration and seven added source-contract cases. The inventory now models the 706 family
registrations plus one orchestrator isolation-failure registration explicitly, the authoritative coverage
documents report 809 runner-listed Commands cases / 707 static registrations / 294 tooling cases, the focused
inventory manifest passes 5 / 0 / 0, and the complete tooling set passes 293 / 0 / 0 with the single
`RequiresBuildToolchain` source case intentionally excluded.

The two remaining Full failures are external blockers accepted under the maintainer's convergence decision.
`RedSalamanderMonitorEtwLatency` exited 1 with an empty log in 403ms, matching the already-archived pre-closeout
Track C failure (374ms) at
`Specs/TestRuns/SINON/Continuation/2026-07-15_lighthouse_track_c_session_end/full/` and remaining owned by
Observatory Track 0/Track 6. `ViewerPETests` exposed the pre-existing
`TestViewerSpaceBoundedHostileProviderScanning` roll-up race: a focused five-run diagnostic produced four passes
and one terminal snapshot with the expected 6 folders / 6 files but only 170 of 210 root bytes after an extra
bounded second. Lighthouse changed ViewerSpace color conversion/blending and DIP rounding only; it did not touch
scan accounting. Observatory Track 0 already owns current Full-failure classification, with Track 17 owning test
reliability architecture. The experimental ViewerSpace repair was dropped. Per the maintainer's explicit
instruction to stop expanding Lighthouse for unrelated failures, these blocking external classifications and
focused rerun evidence satisfy closeout without another 38-minute Full replay.

1. Establish a test-only library/header and migrate sandbox/environment scopes first.
2. Add bounded message-pump and snapshot polling with an explicit timeout diagnostic; migrate one viewer test
   harness at a time.
3. Implement the child-process runner using WIL-owned process/thread/pipe/job handles, a kill-on-close job, and
   concurrent bounded pipe drains. Migrate the Compare case early to close the pipe-full deadlock.
4. Preserve each suite's existing assertions, timeout classification, and artifact names. A helper migration
   must not turn a product failure into a skip or retry-only pass.

**Focused definition gate:**

```powershell
rg -n "GetEnvironment(String|Variable)|SanitizeTestSandboxSegment|PumpUntil|Snapshot|CreateProcessW|CreateJobObjectW|WaitForSingleObject|ReadFile" Tests Common Plugins RedConfigure RedSalamander
```

**Verify:** run every migrated executable directly with its focused cases, exercise a child that writes more
than a pipe buffer to both streams, then use one clean current-head build and one `Suite Full` attempt. A Full
failure outside Lighthouse may close only with a blocking classification, named owner, focused rerun evidence,
and maintainer acceptance; duplicate CI/Full passes are not required.

## Repository conventions that every executor must preserve

- Use WIL RAII for Windows handles and COM ownership; no manual `Release`, `DestroyIcon`, `DeleteObject`,
  or owned-`HWND` destruction through `.get()`.
- No `catch (...)`. At mandatory ABI/noexcept boundaries catch only named exception types and treat
  `std::bad_alloc` as fatal.
- Cross-thread window payloads use `PostMessagePayload`/`TakeMessagePayload`, with create-time registry
  initialization and `WM_NCDESTROY` drain.
- User-visible text lives in localized resources and uses positional `{0}` placeholders.
- Any hot-path or transfer change requires scenario, metrics, deterministic selftest, and archived before/
  after evidence under `Specs/TestRuns/`.
- Duplication removal is characterization-first. Shared APIs name semantic choices such as strict/replacement,
  absolute/drive-qualified, slash preservation, rounding, output version, ownership, and queue ordering.
- A shared helper must live at the lowest dependency layer that can own the policy without introducing a cycle,
  plugin ABI change, mutable singleton, or hidden lifetime extension.
- Completed WIP work moves to `Specs/Plans/Done/`, and durable behavior moves into the authoritative domain
  spec before closeout.

## Git workflow

- Use one branch per Lighthouse-owned track: `codex/lighthouse-loader-hardening`,
  `codex/lighthouse-settings-recovery`, `codex/lighthouse-session-end`, or
  `codex/lighthouse-ci-gate`. Duplication tracks use `codex/lighthouse-common-primitives`,
  `codex/lighthouse-theme-viewer-dedup`, `codex/lighthouse-plugin-contract-dedup`,
  `codex/lighthouse-app-model-dedup`, or `codex/lighthouse-test-support`.
- Commit by logical step with an imperative message matching current history, for example
  `Harden external action executable resolution`.
- Do not push or open a pull request unless the operator explicitly requests it.
- Do not edit files owned by an active HOLD plan without a handoff.

## Global done criteria

- [x] Every LH-1 through LH-9 row is either DONE in its owner plan or explicitly REJECTED with fresh
  evidence and rationale.
- [x] Every LH-D1 through LH-D22 row is DONE or explicitly REJECTED with fresh evidence and rationale; no row
  is closed merely because a helper was renamed or moved.
- [x] Causeway CW-1/CW-2/CW-3 are complete before bridge redesign work begins.
- [x] Firebreak Task 1 is complete before plugin lifecycle consolidation begins.
- [x] Loader hardening has behavioral tests and does not rely on process-global `SetDllDirectoryW`.
- [x] Invalid settings sections no longer reset unrelated settings, and future-schema files cannot be
  overwritten automatically.
- [x] Windows session-end messages persist runtime state through a non-modal, idempotent path.
- [x] CI runs PluginContract, SettingsSchema, and CrashHandling tests and has an ARM64 build-only lane.
- [x] Each duplication family has golden behavior tests, all exact copies have migrated, and every remaining
  local implementation appears in a reviewed allowlist as an intentional policy or dependency variant.
- [x] Shared path, UTF/JSON, callback, COM, packed-buffer, I/O, clipboard, payload, and child-process helpers
  pass their malformed-input, teardown, zero-progress, overflow, or timeout stress gates.
- [x] Track F has archived before/after evidence for any render-hot-path change; no measurable regression is
  accepted as a cosmetic simplification.
- [x] The current-head x64 Debug build reports 0 errors.
- [x] `./Tools/Run-AllTests.ps1 -Suite Full` is green, or every external blocker has a blocking
  classification, owner, and rerun evidence accepted by the maintainer.
- [x] Durable contracts are merged into authoritative specs.
- [x] This plan is moved to `Specs/Plans/Done/` only after all routed and Lighthouse-owned rows are closed.

## STOP conditions

Stop and report instead of improvising when any of these occurs:

- A routed owner plan no longer contains the cited requirement or another active plan claims the same
  files/behavior.
- The live code no longer matches the current-state evidence and the finding cannot be re-confirmed.
- Loader hardening breaks first-party plugin dependency resolution; report the exact failing DLL/load site
  before changing search policy.
- Lossless settings recovery would require silently discarding or rewriting unknown future-schema data.
- Session-end persistence requires a modal prompt, plugin unload, viewer close, or unbounded wait.
- CI enablement exposes a reproducible product failure; do not quarantine or skip it merely to turn the
  workflow green.
- Two candidate implementations differ on malformed UTF/JSON, rounding, hash output, path classification,
  percent-encoding, queue order, cancellation, ownership transfer, COM apartment, or unload behavior. Record
  the variants and refine the API before migrating either caller.
- A proposed abstraction requires a generic plugin scheduler, an all-viewer base window/rendering framework,
  a plugin ABI change, a new mutable singleton, or a Common-to-UI/plugin dependency.
- A shared helper obscures a security boundary or can weaken leaf-name/traversal/ADS checks, size bounds,
  checked arithmetic, atomic `CREATE_NEW`, zero-progress failure, payload teardown, or plugin quiet-point order.
- A test-support migration changes the product oracle, artifact/classification contract, or converts a failure
  into a skip/retry-only pass.
- A track requires files outside its listed scope; refine the track and get review before expanding it.

## Findings considered and rejected

Future reconsideration of the twelve scenarios below is `ROUTED` to
`Operation_Parallax_DeferredArchitectureSecurityAndSimplificationDecisionReview_2026-07-14.md`.
This section remains the audit-time snapshot; update decision status and any resulting implementation owner in
Parallax so Lighthouse does not become a parallel decision queue.

- **Search service lacks client impersonation:** rejected as stale. Current
  `Common/SearchServiceBroker.cpp:2318-2323` and `:2522-2529` use
  `ImpersonateNamedPipeClient` with `RevertToSelf` for authorization.
- **ViewerWeb external navigation is automatically a vulnerability:** not recorded as a defect because the
  current ViewerWeb authoritative spec deliberately defines browser-like HTML behavior. Any policy change
  requires a separate product/security decision. The already-fixed inline-script injection remains distinct.
- **Every `DestroyWindow(_hWnd.get())` is a confirmed double-destroy:** rejected as an immediate bug where
  `WM_NCDESTROY` synchronously releases the wrapper. The pattern remains a repo-policy cleanup owned by
  `Win32_Inventory.md` CHK-1.
- **General exception/CRT/payload hygiene is broken:** rejected. The production sweep found no `catch (...)`,
  forbidden `sprintf_s`/`swprintf_s`, or raw `PostMessageW(...release()/new)` payload transfer.
- **Committed production credentials:** none found by the defensive filename/pattern sweep. No secret values
  were copied into this plan.
- **Every similarly named hash is interchangeable:** rejected. `AppTheme.cpp` hashes both bytes of each UTF-16
  code unit, unlike the byte hash used by DxUi and several viewers. Visible/persisted outputs require named,
  versioned algorithms.
- **One “absolute path” predicate can replace every current predicate:** rejected. Current call sites distinguish
  drive-qualified, drive-relative, rooted, UNC, extended, device, and fully absolute paths inconsistently;
  consolidation requires policy-specific APIs and tests.
- **One UTF converter can serve strict parsers and display fallbacks:** rejected. Strict settings/protocol parsers
  must fail malformed input, while selected UI/convenience paths intentionally replace invalid sequences.
- **One contrast threshold should replace all color-choice code:** rejected. Luminance math can be shared, but
  viewer, theme-expression, and application contrast thresholds are product policies.
- **A generic threadpool scheduler should replace plugin worker code:** rejected. Only callback-state, module-pin
  transfer, and identical MTA entry envelopes are shared; cancellation, sessions, watchdogs, counters, and quiet
  points remain explicit per domain.
- **A common viewer base class/rendering framework is required:** rejected. Lighthouse extracts pure chrome,
  popup, clipboard, geometry, and the proven FunctionBar/StatusBar render-target lifecycle only.
- **Every local JSON accessor should coerce the same inputs:** rejected. Provider protocols and settings schemas
  differ on missing, null, numeric-string, and wrong-type behavior; the shared surface exposes these policies.

## Audit limits

The review covered production subsystems at repository level, current high-risk/hot-churn code, build and
CI configuration, specifications, active remediation ownership, and a definition/pattern census followed by
semantic review of retained duplication candidates. It did not run a whole-program clone detector and cannot
prove that no structurally rewritten or template-instantiated duplicate remains. It did not exhaustively line-
review PoC projects, generated localization resources, images/binary assets, third-party source, live cloud/
provider behavior, dependency-CVE reachability, or every manual UI interaction. Fuzzing, ASan/AppVerifier
stress, power-loss testing, live hostile-provider tests, and per-track runtime/performance validation remain
implementation work.
