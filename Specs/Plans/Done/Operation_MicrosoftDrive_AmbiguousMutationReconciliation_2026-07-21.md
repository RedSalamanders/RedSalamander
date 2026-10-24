# Operation Microsoft Drive Ambiguous Mutation Reconciliation

**Status:** Done - implementation, durable contracts, focused proof, and classified broad closeout complete
**Date:** 2026-07-21
**Owner rows:** `FG-P2-1` and `FG-P2-2` in `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`

## Objective

Make same-drive Microsoft Drive `PATCH` and item-ID `DELETE` mutations truthful and idempotent when Graph commits the
request but the client loses the response. An ambiguous committed mutation must reconcile by stable item ID before the
caller rolls back, reports failure, or retries. The ordinary successful path must not gain another Graph request.

## Last Finding and Current Spec/Code State

### Final closeout checkpoint - 2026-07-21

The Microsoft Drive change is complete. Debug and test-enabled Release both pass the complete `PluginContractTests`
executable with the provider at 165/165. The fake Graph ledger proves exact stable-ID/name/parent PATCH reconciliation,
idempotent direct/retried DELETE not-found, one-probe bounds on ambiguous transports, and zero extra GETs on ordinary
success or definitive HTTP conflicts. Production Release exposes no debug-selftest export. Compact provider evidence
is archived at
`Specs/TestRuns/7d3a1247382a/FileOps/2026-07-21_194100_msdrive_ambiguous_reconciliation/`.

The initial Full attempt exposed two independent test-state defects behind the reported `0x8007051A` dialogs, not a
SettingsStore CAS defect. The search roundtrip now reuses its mutable stamped snapshot; the real `WM_ENDSESSION` child
uses a private unified test root/run ID; and queued settings-watcher messages carry their originating `Start(...)`
generation so an old session cannot load the successor app's file. The formerly failing ordered sequence passes 3/3,
and the full 100-case Settings block plus the new regression and MTP consumer passes 102/102. The durable watcher rule
is in `Specs/Core/Core_SettingsStore.md`; testing coverage owns the child-isolation rule.

Broad closeout is trustworthy with explicit unrelated classification. CompareDirectories passes 226/0 and
FileOperations passes 111/0. Twelve Full-plan standalone/native harnesses pass, the rebuilt test-enabled Monitor latency
drill passes all required metrics with 60 append-to-visible samples, task-relevant source contracts pass 149/149, and
inventory passes 7/7. The complete Commands attempt reached 360/0 before a same-process UIA Invoke/D3D wait; the exact
case reproduces from a clean process, and a remainder attempt passed another 12 UI-heavy cases before a second UIA
Invoke input-idle stall. Time-separated stack captures classify this desktop limitation outside the changed code. Full
Tools Pester passes 345 with only two pre-existing workspace/archive failures: an October-2025 top-level generated
`vcpkg_installed` RC root and the tracked oversized July-20 FileOps archive from commit `148f270de`.

The authoritative Microsoft Drive and testing specs contain the lasting behavior, Floodgate FG-P2-1/FG-P2-2 are
closed, and the WIP routing index no longer points at already-completed FG-A1 work. The Done gate's alternative for a
non-green broad run—durable classification of every unrelated failure—is satisfied.

### Original implementation finding

The previous Microsoft Drive merge-metadata optimization is present on this branch at commit `5242378b7`; this work is
based on that implementation rather than the older `master` snapshot. Its Release-capable fake Graph transport and
exact request counters are reusable for this fault-injection slice.

The remaining failure is real, not merely missing test depth:

- `SendHttpRequest` returns a failed transport `HRESULT` immediately. It retries only completed Graph HTTP responses
  with status `429`, `503`, or `504`; it cannot know whether a failed receive occurred before or after the service
  committed a mutation.
- `MoveItemById` returns that transport failure without querying the stable item ID. In an overwrite move,
  `MoveOrRenameItem` interprets the failure as "primary move did not commit" and attempts to restore the rollback item.
  If the primary `PATCH` actually committed, that rollback decision is wrong and can leave a hidden
  `.redsalamander-rollback-*` item plus a partial-result report.
- `DeleteItemById` accepts only `204`. A committed first DELETE followed by retry/reconciliation `404/itemNotFound`
  is therefore reported as failure even though the requested postcondition is already true.
- `DebugGraphDrive` currently injects throttles before mutation. It cannot yet commit a PATCH or DELETE and then return
  a transport failure, so neither ambiguous outcome is RED-able.
- `Specs/FileSystem/FileSystem_MicrosoftDrive.md` documents commit-aware overwrite cleanup but does not yet define
  post-transport reconciliation or item-ID DELETE idempotence.
- The Floodgate parent queue on this branch already closes FG-P2-5 and FG-P2-6, but FG-P2-1 and FG-P2-2 remain open.

## Required Behavior

1. Add an item-ID metadata GET path using the same authenticated Graph transport, error mapping, parsing, and bounded
   retry policy as path metadata GETs.
2. After a PATCH returns an ambiguous transport failure, query the mutated item ID once:
   - accept the PATCH as committed only when the item still has that ID, its exact requested name, and, when the PATCH
     requested a parent change, the exact requested parent ID;
   - otherwise preserve the original PATCH failure. A failed reconciliation must not replace the original mutation
     diagnostic with a less useful probe failure.
3. Item-ID DELETE is idempotent:
   - direct `404/itemNotFound` is success;
   - after an ambiguous transport failure, one by-ID GET that reports not-found proves
     the delete committed and returns success;
   - a still-present item or inconclusive GET preserves the original DELETE failure.
4. Reconciliation is error-path-only. A normal successful PATCH remains one PATCH and zero reconciliation GETs; a
   normal successful DELETE remains one DELETE and zero reconciliation GETs.
5. Do not recursively delete a drained merge source folder. The existing Graph safety rule remains unchanged.

## Deterministic Proof and Performance Contract

- **Protected scenario:** same-drive Microsoft Drive rename/move/overwrite/delete when Graph commits the mutation but
  the response transport fails.
- **Change type:** reliability stabilization with bounded error-path plugin I/O.
- **User-visible risk:** false failures, incorrect overwrite rollback, orphaned recoverable backups, and destructive or
  confusing user retries after the remote state already changed.
- **Measurement:** fake Graph exact request counts. Successful mutations must remain `1 PATCH/DELETE + 0 reconciliation
  GET`; each injected ambiguous mutation may add exactly one item-ID reconciliation GET. No production per-request perf
  row will be added because it would instrument every remote operation to protect a rare failure path; the deterministic
  request counter is the less-distorting measurement.
- **Harness:** `RedSalamanderMicrosoftDriveDebugSelfTests` through `PluginContractTests`, compiled in Debug and
  test-enabled Release.
- **Required RED cases:**
  - overwrite backup PATCH commits then loses its response and is recognized by ID, so the primary move completes and
    the backup is deleted;
  - primary source PATCH commits then loses its response and does not trigger rollback;
  - DELETE commits then loses its response and returns success after by-ID not-found;
  - first DELETE attempt commits but returns a retryable status, the automatic retry receives `404/itemNotFound`, and
    the operation returns success without a separate GET;
  - PATCH transport failure before commit remains failure and preserves source/destination identities;
  - DELETE transport failure before commit remains failure and preserves the item.
- **Archive:** retain compact Debug/test-enabled Release provider results and request-count evidence under
  `Specs/TestRuns/<MachineHash>/FileOps/<RunId>/`.

## Progress Log

- [x] Audited the live branch and confirmed the preceding metadata-reuse implementation and fake Graph harness are
  available here.
- [x] Confirmed PATCH and DELETE currently return ambiguous transport failures without stable-ID reconciliation.
- [x] Confirmed DELETE currently rejects direct `404/itemNotFound` even though absence is the requested item-ID delete
  postcondition.
- [x] Defined the bounded error-path-only I/O contract, RED cases, authoritative-spec migration, and Done gate before
  changing product code.
- [x] Extended fake Graph with attempt-targeted pre-commit and committed-then-transport-failure injection for PATCH and
  DELETE, by-ID GET routing/counting, and `parentReference.id` in item metadata. Failed requests remain visible in the
  request ledger, so each proof can distinguish the mutation attempt from its reconciliation probe.
- [x] Added RED assertions for ambiguous overwrite-backup PATCH, ambiguous primary PATCH, pre-commit PATCH failure,
  committed ambiguous DELETE, direct already-gone DELETE, and pre-commit DELETE failure. Existing successful rename
  and throttled-DELETE cases now also require zero reconciliation GETs.
- [x] Captured the RED provider contract. Debug `PluginContractTests` and Microsoft Drive built with zero warnings/errors
  (`msbuild-20260721_135956_710.log`, `msbuild-20260721_140044_705.log`). The provider reported 147 passed / 11 failed,
  exactly matching the new unmet contract: three backup-PATCH assertions, three primary-PATCH assertions, one
  pre-commit PATCH probe-count assertion, two committed-DELETE assertions, one direct already-gone DELETE assertion,
  and one pre-commit DELETE probe-count assertion. Existing provider checks stayed green. Captured console:
  `.build/msdrive-ambiguous-red-20260721.txt`.
- [x] Implemented item-ID metadata GET through a shared provider-local URL request/parser path. `ItemMetadata` now
  copies `parentReference.id` while the yyjson document is alive, enabling exact requested-parent verification without
  retaining borrowed JSON storage.
- [x] Implemented bounded mutation reconciliation. PATCH preserves its original failure unless one by-ID GET proves the
  same ID has the exact requested name and requested parent (when present). DELETE accepts direct `itemNotFound` and,
  after a transport failure, accepts success only when one by-ID GET proves absence. Normal `2xx`/`204` paths
  return before the probe.
- [x] The first GREEN attempt passed 157/158 provider assertions. Isolation proved the remaining failure was the existing
  late-destination `409 nameAlreadyExists` exact-request-count contract: the initial implementation reconciled every
  terminal HTTP error, adding a useless by-ID GET to a definitive conflict. Narrowed reconciliation to failed transports
  only; completed HTTP failures retain their existing Graph error mapping, while direct DELETE `404` remains success.
  Temporary isolation counters were removed immediately after diagnosis.
- [x] Final focused Debug Microsoft Drive provider contract is green at 158 passed / 0 failed, and the complete
  `PluginContractTests` executable passes. The final provider rebuild has zero warnings/errors
  (`msbuild-20260721_141209_580.log`); captured console: `.build/msdrive-ambiguous-green-debug-20260721.txt`.
- [x] Added a second pre-commit PATCH guard where the stable ID already has the requested leaf name but remains under
  the wrong parent. Reconciliation must preserve the transport failure and source identity, proving that name-only
  matching cannot falsely commit a parent-changing move. This raises the provider contract to 160 assertions pending
  the next verification run.
- [x] Final Debug rerun passes Microsoft Drive 160/160 and complete `PluginContractTests`; provider rebuild is
  warning/error-free (`msbuild-20260721_141427_319.log`). Captured console:
  `.build/msdrive-ambiguous-green-debug-final-20260721.txt`.
- [x] Test-enabled Release provider and `PluginContractTests` builds are warning/error-free
  (`msbuild-20260721_141526_160.log`, `msbuild-20260721_141617_813.log`). Optimized provider coverage passes
  Microsoft Drive 160/160 and the complete plugin contract executable; captured console:
  `.build/msdrive-ambiguous-green-release-20260721.txt`.
- [x] Existing `TestHarnessSourceContracts.Tests.ps1` remains green at 149 passed / 0 failed. No new source-shape
  assertion was added because the compiled provider contract observes the exact mutation behavior and request counts,
  consistent with the repository rule preferring behavioral proof over private-helper spelling.
- [x] Production tests-disabled Release Microsoft Drive rebuilt with zero warnings/errors
  (`msbuild-20260721_141833_535.log`). `dumpbin /exports` reports zero
  `RedSalamanderMicrosoftDriveDebugSelfTests` exports, so the fake Graph and fault controls remain absent from the
  shipped DLL.
- [x] Final acceptance review found the FG-P2-2 proof was only implied by separate throttle and direct-not-found cases.
  Added an exact fake Graph sequence where DELETE attempt 1 removes the item but returns `503`, the built-in retry issues
  attempt 2 and receives `404/itemNotFound`, and the operation succeeds with exactly two DELETEs and zero by-ID GETs.
  The focused Debug/test-enabled Release gates below must be rerun at the new 165-assertion total before closeout.
- [x] Exact committed-DELETE retry coverage is green in Debug and test-enabled Release at Microsoft Drive 165/165;
  complete `PluginContractTests` passes in both flavors. Final provider build logs are
  `msbuild-20260721_142041_783.log` (Debug) and `msbuild-20260721_142147_870.log` (test-enabled Release); captured
  consoles are `.build/msdrive-ambiguous-green-debug-final-20260721.txt` and
  `.build/msdrive-ambiguous-green-release-final-20260721.txt`.
- [x] Production tests-disabled Release was restored after the final test edit with zero warnings/errors
  (`msbuild-20260721_142254_835.log`), and `dumpbin /exports` again reports zero Microsoft Drive debug-selftest
  exports.
- [x] Started the Full closeout gate under TestSandbox run
  `20260721T122429Z-19172-3aadd909e4dd4b1ca03c9e9035cb9909`, preserved both settings revision-mismatch dialogs,
  and captured the complete Commands result (526 passed / 18 failed / 0 skipped). Identified the first failure as an
  invalid fresh-snapshot second save in `settings_store_search_roundtrip`; production CAS behaved correctly.
- [x] Saved a pause-safe continuation archive at
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-21_151853_msdrive_full_settings_revision_mismatch/`, including the
  complete Commands result/trace and explicitly partial File Operations result/trace. The Full wrapper was stopped at
  the user's request only after capture.
- [x] **Continuation resumed:** repaired `settings_store_search_roundtrip` by making the first prepared snapshot mutable
  and reusing it for the dirty second save, so the successful first CAS commit advances the exact snapshot's
  `expectedFileStamp`. Production CAS remains unchanged.
- [x] The intentionally interrupted Full wrapper left a stale artifact-operation owner record. Recovered through the
  required serialized full-solution Debug `-Rebuild`; `msbuild-20260721_164254_446.log` records zero warnings and zero
  errors, and both the owner and contamination markers are cleared.
- [x] Exact post-repair Commands gate passed `settings_store_search_roundtrip` 1/1 with no dialog or failure under
  TestSandbox run `20260721T145808Z-39860-3f83052c38c34a5eb41a968c7e50af19`.
- [x] Started the complete ordered Commands suite under run
  `20260721T150024Z-63056-ba7d0fabad60401b97e3bd7404b01877`. The repaired settings case passed in sequence; the
  partial result reached 169 passed / 1 failed, with only the unrelated theme-resolution perf guard red
  (`380873 us` versus `100000 us`). The Codex tool host then exited and its kill-on-close job stopped the run before
  Connection Manager. Preserved the partial result, trace, and abandoned owner metadata under
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-21_170914_commands_tool_host_interrupt/`; no application crash dump
  was created.
- [x] Recovered the tool-host interruption through a second serialized full-solution Debug rebuild. Log
  `msbuild-20260721_171129_421.log` records zero warnings and zero errors; the owner and contamination markers are
  cleared.
- [x] Exact previously failing cases all pass from a clean process with no modal dialog: Connection Manager MTP save
  (`20260721T152443Z-12616-55cc20433261485fa4906aa7f5619618`), main Preferences Apply
  (`20260721T152501Z-36012-9ac85b3e9b5f43c49127c91482303aa7`), Monitor Apply
  (`20260721T152514Z-60948-99403d2b18104df095791d31b92c384f`), main settings-file link
  (`20260721T152605Z-68168-04318f8f4b044f0ab02137fe944014af`), Monitor settings-file link
  (`20260721T152628Z-53452-7b2063d4fb4146c080f9ac25443ecd94`), and viewer/editor Apply
  (`20260721T152652Z-69112-664ea0030d434453aec38f9447f13be4`). This proves the old `0x8007051A` dialogs are
  ordered-suite state leakage, not standalone startup/save defects.
- [x] The ordered `cmd_connection_manager_window_` prefix passes 35/35 under run
  `20260721T152735Z-64624-ef8ecc3aea2d4b4a8fa7ab150530dd39`, with no modal dialog. The Connection Manager family
  does not corrupt its own settings lineage; the state leak is earlier in Commands, inside or across the Settings,
  Batch Rename, or Plugin Config families.
- [x] Ordered comma-list isolation eliminated Plugin Config (20/20, run
  `20260721T153312Z-36668-a9e3628e1d974b59bcd29debed53ae1d`) and reproduced the failure with the Settings family
  alone (100 setup cases passed before the MTP save failed, run
  `20260721T153444Z-71064-c580081397ba468b8775905d6af3d4dc`). Successive halves reduced the predecessor to the
  single `settings_save_queue_serializes_coalesces_and_flushes` case: paired with the MTP save it reproduces under
  `20260721T153923Z-69240-8aee73baff3148d89c57d66bdcd5b9a9`, while
  `settings_hot_reload_self_save_suppression` plus MTP passes under
  `20260721T153905Z-31624-d03db1d13c4f4dab9671a1e907b26df7`.
- [x] Fixed the isolated root cause: the save-queue case launches a real `WM_ENDSESSION` child with an isolated
  `REDSALAMANDER_SELFTEST_ROOT` for artifacts but inherits the parent's unified `REDSALAMANDER_TEST_ROOT` and
  `REDSALAMANDER_TEST_RUN_ID`. The child therefore writes the parent's real `RedSalamander-debug.settings.json`,
  leaving the parent's `g_settings.expectedFileStamp` stale. The child now receives a private unified test root and
  safe run ID beneath its existing disposable root; the cleanup removes both artifacts and settings. Production CAS
  is unchanged. Debug build `msbuild-20260721_174137_993.log` completed with zero warnings/errors, and the exact
  predecessor/consumer pair now passes 2/2 under run
  `20260721T154624Z-47076-b1a03ea2f72e41d186a2a173318b9812`; no `settings-save-child-*` directory remains.
- [x] Captured the second independent Settings-family state leak: the repaired 100-case slice still left the final MTP settings save red under run
  `20260721T154709Z-37432-e28723bd848241fbaa56a042861eebc4` (100 setup passes / 1 consumer failure, no modal).
  This is evidence of a second independent Settings-family state leak. The new bisection reduces it to the ordered
  interaction `settings_hot_reload_invalid_external_file` then `settings_hot_reload_transient_arm_failure_is_async`:
  the pair plus MTP fails under `20260721T155137Z-69248-5cc01db098ea4ec5b9986d453f0b7083`, while each predecessor
  plus MTP passes independently (`20260721T155200Z-65056-2be4e379ad664539ab5a949200b69ff9` and
  `20260721T155233Z-37624-5baed5aea2c541fa96c7bac2adeaff4b`).
- [x] Fixed the second root cause in production hot reload: `Start()` changes the singleton app/session, but
  `SettingsFileChangedPayload` carries only a tick count. A queued directory-change notification from the prior main
  watcher can therefore reach the main window after the next test watcher starts, and `TryLoadChangedSettings()` then
  loads the new test app into main `g_settings`. Every posted payload now carries the watcher generation captured by
  `Start()` (including deferred internal-save and retry posts), and the main-window notification path rejects a stale
  generation under the watcher lock before reading any app ID or settings file. Direct same-session loads retain the
  existing API for controlled callers. Added the deterministic
  `settings_hot_reload_stale_notification_cannot_cross_session` regression, which captures a first-session payload,
  starts a second app session, and requires the stale notification to be a no-op before the current session loads.
  Debug RedSalamander rebuilt successfully with no warnings/errors in
  `msbuild-20260721_175855_051.log`. The new exact stale-generation regression passes 1/1 under TestSandbox run
  `20260721T160640Z-70592-f959e4cd7f3245af95ee43f681d584c4`. The previously failing ordered interaction—invalid
  external settings, transient watcher-arm failure, then the MTP settings save—now passes 3/3 under run
  `20260721T160657Z-72892-54c4799460ec45b49499d7ffc220cd96`. The original complete 100-case Settings block, the
  new stale-session regression, and the MTP consumer now pass together 102/102 under run
  `20260721T160758Z-18840-f52e382c845e4a818463a82567c1a7f9`, closing both independently isolated settings-state
  leaks before the full ordered suite.
- [x] Merged the lasting settings findings into authoritative specs. `Core_SettingsStore.md` now requires every queued
  watcher notification to carry its originating `Start(...)` generation and requires the main consumer to reject a
  stale generation before app-ID/file access. `Testing_TestCoverage.md` records that regression and the real
  `WM_ENDSESSION` child-process unified-root/run-ID isolation contract, with the focused 1/1, interaction 3/3, and
  ordered 102/102 evidence.
- [x] The complete ordered Commands rerun `20260721T161024Z-68464-d834d39efd6946229109c6f8eb0254a5` reached 360
  passed / 0 failed, including the repaired Settings block and subsequent Connection Manager/Preferences consumers,
  then stopped making progress inside unrelated
  `cmd_preferences_dialog_plugins_live_search_dx_interaction`. Two time-separated `StackDump.exe` captures had the
  identical main-thread path: `InvokeVisibleDescendantByName` -> UI Automation core -> Preferences DxUi window proc ->
  `DxUi::WindowHost::Render` -> DXGI/D3D11 `WaitForSingleObjectEx`. The responsive process was terminated after the
  repeated fingerprint so the wrapper could preserve results; its synthetic coverage failure is therefore not a
  product assertion failure. Compact result/trace evidence is archived at
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-21_1832_commands_plugins_uia_hang/`. The exact single case also
  hangs from a clean process under run `20260721T163208Z-63256-96e7d0de22bf47f4b6de3cbf779a117d`; its stack again
  has the same UIA -> Preferences DxUi render -> graphics wait chain. The standalone wrapper and trace are archived
  under that directory's `standalone/` child. This classifies the interruption as an independent repeatable UIA/D3D
  environment/test hang, not ordered settings state leakage or a Microsoft Drive regression.
- [x] Attempted and classified the 184-case Commands remainder after excluding the Plugins live-search hang. The next 12
  UI-heavy Plugins/Themes cases passed, then `cmd_preferences_dialog_themes_set_live_dx_interaction` also stalled in
  same-process `InvokeVisibleDescendantByName`, with the main thread stationary in UI Automation's input-idle wait.
  Run `20260721T163522Z-59632-0a920729ad834f36a4301a88b88306ff` was stopped after stack capture; its 12/0 partial
  result/trace is archived under the same continuation directory's `remainder/` child. This desktop session cannot
  currently provide a trustworthy complete UIA-heavy Commands gate; continue with non-UIA suites and classify the
  Commands limitation explicitly at closeout.
- [x] Unaffected CompareDirectories closeout gate passes under run
  `20260721T164614Z-72720-8c788e63aed3440fb2849e525b5d78a5`: 226 passed / 0 failed / 31 environment-dependent
  skips. The skips are the declared live-credential/device/UI and unavailable-journal/NTFS capabilities; there are no
  regression or isolation classifications.
- [x] Unaffected FileOperations closeout gate passes under run
  `20260721T165057Z-53244-b9ba54e7418f436cbe28d68a9ccc59c1`: 111 passed / 0 failed / 20 declared skips.
  The skips are 7z cases owned by Compare plus live remote profiles not configured on this machine; all deterministic
  local/dummy/bridge/queue/watch/performance phases are green with no regression or isolation classifications.
- [x] Full-plan standalone lane run `20260721T170824Z-73372-1b56dc1adc5540a8a024742aec0466e8` passes 12 harnesses:
  DxUiTests, FileSystemCurlTests, ViewerPETests, ViewerSqliteTests, MonitorTest, LocalizationTests, RedConfigureTests,
  PluginContractTests, SettingsSchemaTests, CrashHandlingTests, PerformanceTests2, and VcpkgMergeSynthetic. The three
  in-product zero-case failures are an intentional filter/coverage artifact because those suites were handled
  separately. Two real standalone results remain to classify: `RedSalamanderMonitorEtwLatency` exited 1 with no
  output, and `ToolsPesterTests` exited 1 after its full aggregate. Compact aggregate plus relevant native outputs are
  archived at `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-21_1921_full_standalone/`.
- [x] Classified the two standalone-lane failures. The Monitor archive
  `Specs/TestRuns/7d3a1247382a/Monitor/2026-07-21_191300/results.json` explicitly reports that the skip-build binary
  lacked `ENABLE_TESTS`; a subsequent test-enabled direct run at `2026-07-21_192157` passed all required metrics with
  60 append-to-visible samples. A final serialized test-enabled build/rerun remains required because a second
  concurrent launch at `192211` interfered with the single-instance UI and failed.
- [x] Focused Pester inventory and task-relevant source contracts pass 7/7 and 149/149. The complete quiet Pester rerun
  passes 345 and has exactly two unrelated failures: `ResourceLocalizationContracts.Tests.ps1` sees the generated
  top-level `vcpkg_installed` tree (created October 2025) as an unallowlisted RC root, and
  `TestRunArchive.Tests.ps1` rejects the pre-existing tracked July 20 FileOps archive at commit `148f270de` because its
  7,009,000-byte perf JSONL exceeds both file/run size caps. Neither path was created or changed by this work.
- [x] Rebuilt RedSalamanderMonitor Debug with test hooks, zero warnings/errors
  (`msbuild-20260721_193215_677.log`), then ran the ETW latency drill once with no competing Monitor instance and a
  visible interactive window. Archive `Specs/TestRuns/7d3a1247382a/Monitor/2026-07-21_193252/` passes: 60
  append-to-visible samples, 61 batch-drain/queue-depth/repost rows, 122 frame total/present rows, and every required
  metric present. This closes the earlier skip-build artifact mismatch.
- [x] Production tests-disabled Release RedSalamander builds successfully (`msbuild-20260721_193332_845.log`, zero
  errors). Its three C5245 diagnostics are pre-existing unrelated tests-disabled-only helpers in
  `BatchRenameWindow.cpp` and `FindFilesWindow.cpp`; no changed file warns. The same build regenerated the production
  Microsoft Drive DLL, and `dumpbin /exports` again finds zero
  `RedSalamanderMicrosoftDriveDebugSelfTests` exports.
- [x] Implement item-ID GET plus PATCH/DELETE reconciliation. Product code and focused GREEN verification are complete.
- [x] Pass focused Debug and test-enabled Release provider contracts with exact request counts.
- [x] Pass the existing source-contract suite and production tests-disabled Release gate.
- [x] Repair the invalid settings selftest CAS lineage and pass its exact case plus the previously failing standalone
  settings-save cases. The remaining ordered-suite leak is a separate earlier-case isolation task.
- [x] Complete a trustworthy full-suite gate (or durably classify every remaining unrelated failure) and archive compact
  final Microsoft Drive evidence.
- [x] Merge the durable contract into `Specs/FileSystem/FileSystem_MicrosoftDrive.md` and testing guidance.
- [x] Close FG-P2-1 and FG-P2-2, reconcile the WIP index, and move this plan to `Specs/Plans/Done/`.

## Done Gate

This plan moves to `Specs/Plans/Done/` only when all required behavior is implemented, every adversarial fake Graph
case is green in Debug and test-enabled Release, normal-path request counts prove zero regression, production Release
contains no test export, relevant source contracts and the full suite are green (or every unrelated failure is
classified with durable evidence), compact run evidence is archived, the authoritative Microsoft Drive/testing specs
own the lasting rules, and both Floodgate rows are marked complete.
