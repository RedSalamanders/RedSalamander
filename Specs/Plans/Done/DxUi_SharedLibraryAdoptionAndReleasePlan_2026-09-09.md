# DxUi shared-library adoption and release plan

## Completion checklist

**Completed locally on 2026-09-13** for the accepted x64 Debug/Release scope.
Native candidate `1232d181` consumes exact DxUi `13788e95`; subsequent closeout
changes are documentation and retained evidence only. No hosted CI was used.
The user's RedSalamander master checkout has not been merged.
The migration was rebased onto `0dd0bd6d` and incorporates latest master
`e8a9f857` through merge `2f4fb03f`; the final candidate is `1232d181`. Its five WinGet release-tooling files
match master Git blobs exactly; both final Full runs include those changes.

- [x] Rebase onto `0dd0bd6d` and incorporate latest master `e8a9f857`, preserving
  its build/Terminal and WinGet fixes.
- [x] Link exact-pinned canonical `DxUi.lib` in production consumers; remove all 27
  legacy implementation and 28 legacy test/project/baseline files. Keep product adapters.
- [x] Implement G4 clipboard/font corrections and use the accepted standalone
  diagnostics/Slider behavior. RedSalamander pins `13788e95`; RedXe pins `ecad707`.
- [x] Reconcile NavigationView, Find and menu diagnostics with the optional borrowed
  library sink; legacy file/flag compatibility is not required by the accepted G5 decision.
- [x] Support the six build configurations, exact-pin restore, module identity checks,
  advisory updates and the ordinary upgrade/fix/retest/rollback workflow.
- [x] Complete the library and RedXe adoption records, including their tests,
  normative contracts, accepted resource measurements and explicit platform deferrals.
- [x] Preserve original migration/rebase and fixture failures with the corrected
  regressions. Pending-content, borrowed-name lifetime and bandwidth pass ten Debug
  repeats; queue-after and don't-start each pass ten Debug repeats at `f06f779f`
  and ten final Release repeats at `1232d181`.
- [x] Correct the diagnostic source guard: 187/1 before, 217/0 with the corrected
  source-contract, inventory and governance checks. Native assertions remain unchanged.
- [x] Correct the ancestor-click fixture to require stable ellipsis/window geometry
  before input. The original popup predicate, all nine scaled three-second waits and
  ancestor navigation/cleanup suffix remain unchanged. All 188 source checks pass;
  the final thirteen-case history/ancestor context passes 130/130 in Debug.
- [x] Pass the final x64 Debug root build and Fresh Full: **2,122 passed, 0 failed, 52 skipped**,
  20/20 entries, with all 14 final module identities verified.
- [x] Pass the corrected thirteen-case Refresh context ten times at `a8350f4b`
  (130/130), and the history context ten times at `1232d181` (130/130), with
  all protected post-action assertions unchanged and all 14 module identities verified.
- [x] Pass the final x64 Release root build and Fresh Full: **2,118 passed, 0 failed, 56 skipped**,
  20/20 entries, with all 14 final module identities verified.
- [x] Pass the final Release portable A/B/A rollback: application startup, plugin
  contracts, runtime closure, all 11 packaged DxUi identities and identical baseline
  payloads before/after. ZIP growth is 415,535 bytes
  (1.37%); expanded growth is
  1,140,347 bytes (1.45%).
- [x] Record the user's explicit acceptance of the measured resource trade-off,
  retaining original results, thresholds, baselines and measurement limitations.
- [x] Transfer further ARM64/ASan qualification to user-owned
  [H4](../WIP/DxUi_DeferredPlatformQualification_2026-09-13.md). Original failures remain retained; skipped cases and
  unexecuted platform/hardware coverage are not passes.
- [x] Prepare release handoff: publish the tested library revision and verify public
  exact-pin restore without a local Git URL rewrite before consumer merge/release.
  Publication and master merges are separate lifecycle actions, not performed here.
- [x] Test exact admission of this plan and its two supporting Done records: all
  48 focused specification/documentation checks pass; unreviewed siblings are rejected.
- [x] Finish the final archive, link/index and normative-consistency validation:
  48 focused checks, specification/tooling inventories and all 4,620 archived files pass.
  The ten-file final receipt also passes archive admission; its seven original logs/results
  match Git bytes. [Final documentation evidence](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-FinalDocumentation/README.md).

Final receipts: [Debug](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Debug1232d181Isolated/README.md),
[Release](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Release1232d181Isolated/README.md),
[portable rollback](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Packaging1232d181/README.md).
The [qualification index](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-LocalQualification/README.md)
retains earlier failures and their follow-up. Only this opening checklist states the
completed scope; earlier checkpoints below retain historical wording and source identities.
The earlier `f06f779f` Debug Full has two retained navigation failures. Reduced and
expanded context runs did not reproduce them; ordered diagnostics and the Refresh
quiet-baseline correction are recorded separately. The `a8350f4b` Full history-popup
failure is also retained. The thirteen-case history context passes ten repetitions
both unchanged and with diagnostics; those focused passes do not explain the failure.
The `2f4fb03f` Full history case passes, while its source guard and preceding
ancestor-popup fixture fail. The guard correction retains the bounded success condition;
the ancestor correction addresses the existing stable-geometry precondition. The original
failed popup predicate remains unknown. Final passing Full runs qualify this candidate
without assigning an unproven cause to the earlier native failures.
The intermediate `e61d5e34` build fails on an extra parenthesis in the new
diagnostic call; `1232d181` corrects that single character. Its failed build and
188 passing source checks remain retained, separate from final native qualification.

## Retained qualification checkpoints

### Earlier qualification checklist

Current checkpoint: **local qualification; no hosted CI**. The migration is implemented and
rebased. The user's RedSalamander `master` checkout has not been merged. I19 remains WIP while
the acceptance gates below are open; previous successful profiles do not qualify changed code.
The corrected candidate is frozen at `f06f779f`. On 2026-09-13 the user explicitly
deferred ARM64 and ASan and requested completion of the remaining work. Those gates
are now owned by [the deferred platform qualification](../WIP/DxUi_DeferredPlatformQualification_2026-09-13.md).
Local x64 Debug and Release qualification proceeds independently; no ASan or ARM64 pass is inferred.

- [x] Rebase RedSalamander onto master `0dd0bd6d`, preserving its Terminal/build/dependency fixes.
  DxUi and RedXe already contain their reviewed default-branch bases.
- [x] Use canonical `DxUi.lib` and public headers throughout the migration branch; remove all 27
  legacy implementation files and 28 legacy test/project/baseline files. Keep product adapters.
- [x] Implement G4 clipboard semantics and font refresh upstream; accept standalone diagnostics
  and Slider behavior. RedSalamander pins `13788e95`; RedXe pins `ecad707`.
- [x] Implement exact-pin restore, Debug/Release/ASan Debug on x64/ARM64, module identity checks,
  advisory updates and the ordinary upgrade/fix/retest workflow. Retain all test dispositions.
- [x] Fix rebased header, localization, checkbox repaint, Preferences geometry, clipboard fixture
  and ViewerImgRaw fixture failures. The corrected ImgRaw case passes ten Debug and ten ASan Debug
  repeats without changing product code, assertions or deadlines. Menu hover passes ten isolated repeats.
- [x] Build all six RedSalamander profiles at `751437d2`; retain compiler-host/PDB failures and
  successful retries. The later x64-host default is implemented and tested through real MSBuild.
- [x] Reconcile the remaining G4/G6 spec contradictions: native DxUi clipboard transport opens
  once with checked large-selection support; Slider geometry and events match the accepted library.
  Keep the original gap audit as a dated source comparison with a current implementation pointer.
- [x] Qualify canonical DxUi's current code with local x64 suites and ARM64 cross-builds. The latest
  Release run has 18 successful suites, with nine Menu desktop-capability skips explicitly retained.
- [x] Collect matched DxUi/RedXe resource runs and 40 passing RedSalamander Preferences round trips.
  Retain raw comparisons, baseline variability and earlier failed performance comparisons.
- [x] Qualify the preliminary Release portable ZIP at `751437d2`: clean-extraction startup/plugin
  smoke passes; all 11 packaged DxUi sidecars match actual module bytes and the exact pin. Keep the
  tested baseline ZIP. The migration package grows 1.33% compressed / 1.41% expanded.
- [x] Pass the earlier x64 Release build and Fresh Full on frozen `a20f6680`: zero build warnings/errors,
  **2,113 passed, zero failed, 56 skipped; 20/20 entries pass**. All 770 general tooling and
  12 build-tooling checks pass in the isolated process. Compare is 239/0/33, Commands 898/0/2
  and File Operations 179/0/21, including all five corrected ViewerText cases. All 14 post-Full
  module identities match. [Reviewed Release receipts](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Releasea20f6680Isolated/README.md).
- [x] Retain and diagnose the earlier failed x64 ASan Debug qualification on frozen
  `a20f6680`. Its root build passes in 3m45s
  with zero warnings/errors; the seeded heap overflow is detected with expected exit code 1.
  Fresh Full finishes with **1,831 passed, eight failed, 63 skipped; 16/20 entries pass**.
  All 770 general tooling and 12 build-tooling checks pass, and all 14 post-Full module hashes
  match. `content_pending_elided` observes an already-completed initial snapshot; the bounded
  callback gate correction awaits validation. These failures remain evidence and are not waived
  by the Release pass. [Reviewed ASan receipts](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-ASanDebuga20f6680Isolated/README.md).
  Commands also exits after 632 passing cases with incomplete coverage. Its next registered
  case exercises removal-focus contracts; a borrowed-name lifetime bug in that fixture is
  corrected locally. Isolated ASan reproduction confirms heap-use-after-free; corrected validation is pending.
  RedConfigure also fails its 512-token theme mass-preview timing assertion: 216,196 us
  against 100 ms under ASan. The owning UI contract specifies a deterministic Debug budget;
  five isolated probes reproduce the failure on both master and migration (215-225 ms).
  The candidate retains normal Debug/Release limits and records ASan timings diagnostically
  while preserving correctness checks. Corrected native validation remains open.
  File Operations also reports a bandwidth-throttle duration lead of 469 ms against its
  250 ms allowance. The fixture starts its clock when the UI first observes the worker;
  the correction measures the published worker start and separately records observation delay.
  Its final aggregate is 162 passed, five failed and 33 skipped. The other failures are the
  R0d capability benchmark (2.405 s versus 2 s), fake-MTP discovery/cancellation overlap,
  Curl serial wide-copy deadline (180 s; 3,280/4,097 publications), and the queue-after wait
  witness (both operations succeed, but zero waits recorded). Master reproduces R0d and Curl;
  queue-after and properly isolated MTP probes pass on both sources. Broad MTP validation remains open.
- [x] Transfer final ASan and ARM64 qualification to the explicit user-owned deferral.
  On `d404fc90`, the ASan build and seeded detection pass; pending-content and removal-focus
  each pass ten repeats. Bandwidth repeat 10 fails cancellation latency (765 ms versus
  250 ms), stopping the driver. Original failed evidence remains; ASan Full is unqualified.
- [ ] Pass final x64 Debug and Release builds, focused regressions and Fresh Full on
  `f06f779f`, each in a separate clean process. Debug qualification has restarted after
  repairing the repeated queue fixture's stale Dummy destination.
- [x] Retire affected-test manifest references to deleted DxUi projects and map `ProductUiTests`
  to its live project. Exact pins and shared adapters retain conservative Full coverage.
  The ownership test fails before; all nine impact tests pass after. Tool/spec inventory passes.
  [Reviewed mapping regression](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-ProductUiImpact/README.md).
- [x] Reproduce the removal-focus case on frozen `a20f6680` with an ASan dump. The debugger
  confirms heap-use-after-free; exact executable/PDB/dump identities and the source comparison
  are retained. [Reviewed reproduction](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-RemovalFocusAsanBefore/README.md).
- [x] Validate the bounded pending-content and borrowed-name fixture corrections on frozen
  `d404fc90`: each passes ten Debug and ten ASan Debug repeats. The Debug bandwidth case also
  passes ten repeats with delayed UI observation and unchanged throughput/cancellation limits.
- [ ] Validate queue-after and don't-start repeated isolation on `f06f779f`, then Fresh Full.
  Validate the measured RedConfigure/R0d profile corrections, the bounded ASan Curl deadline
  and the queue-after publication gate. Release performance limits remain unchanged.
  Validate the corrected bandwidth worker-start measurement with delayed UI observation;
  retain the existing throughput and cancellation limits.
  Recheck the MTP discovery/cancellation overlap in the broad run; isolated passes do not waive it.
- [ ] Requalify the x64 Release portable package/rollback on `f06f779f`;
  the successful `a20f6680` receipts remain historical. Refreshed ARM64 builds are deferred.
- [x] Retain master/migration diagnosis: both fail the RedConfigure ASan budget, master also
  reproduces R0d and Curl deadlines, and fresh-journal MTP plus queue-after pass in isolation.
  [RedConfigure profile evidence](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-RedConfigureAsanBudgetBefore/README.md),
  [File Operations baseline evidence](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-FileOpsAsanBaselines/README.md).
- [x] Cross-build frozen `a20f6680` in all three ARM64 configurations: Debug 4m00s, Release 8m19s,
  ASan Debug 3m43s; zero errors, 44/60/46 warnings retained. All 42 module hashes match their
  actual binaries, exact pin and x64 compiler host. [Reviewed receipts](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-ARM64a20f6680/README.md).
  Prior `ad7f7735` builds pass: Debug 3m51s, Release 4m58s and ASan Debug 3m40s, all with zero errors.
  Retain 44/60/46 warnings respectively; these receipts precede the test-diagnostic correction.
  All 14 consumer module identities per profile match their bytes, pin and x64 compiler host.
  Restore and the library project reference now use the selected first-party host consistently.
  Both regression checks fail on the old implementation; all eight dependency and 29 governance
  tests pass after correction. The earlier 13 host-mismatch errors and the incorrect probe-path
  retry remain preserved in the compiler-host qualification record.
- [x] Cross-build RedXe `f727932` in all three ARM64 configurations with the x64 compiler host:
  Debug 1m45s, Release 1m21s, ASan Debug 1m08s; zero warnings/errors. All three archive consumers
  per profile match their recorded module bytes and pin. Native execution is a separate gate.
- [x] Record native ARM64 runtime qualification as explicitly deferred to the user.
  Cross-builds do not establish native execution or sanitizer detection; no local ARM64 host is identified.
- [x] Resolve measured resource acceptance: on 2026-09-13 the user explicitly answered
  "Accept the measured trade-off and finish." The accepted costs and limitations are persisted
  in each repository's resource contract; original receipts and thresholds remain unchanged.
  RedXe's controlled Release FPS
  changes range from -1.76% to +0.17%, while 96-DPI dirty composition adds 2.6-5.2 microseconds.
  DxUi Release clean peak private memory rises 731,136 bytes (3.25%) in the longer fixture;
  baseline-repeat comparisons also exceed bands. No baseline or threshold has been changed.
- [x] Qualify the corrected `a20f6680` x64 Release portable package and A/B/A rollback rehearsal.
  All three clean-extraction startup/plugin/runtime-closure checks pass; all 11 candidate DxUi
  sidecars match actual module bytes, and baseline payload hashes match before/after.
  ZIP growth is 412,255 bytes (1.36%); expanded growth is 1,124,475 bytes (1.43%).
  [Reviewed package receipts](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Packaginga20f6680/README.md).
- [x] Prepare the release handoff: the tested library revision must be published before consumer merge/release,
  then restored without the local Git URL rewrite. Publication and product merge/release are separate
  lifecycle actions; no push or merge is performed by this local implementation closeout.
- [ ] Complete applicable product/manual acceptance, reconcile durable contracts and indexes,
  then move I19 and its supporting documents to Done. Existing out-of-scope AV capabilities retain
  their owners and remain explicitly unqualified.

Reviewed local receipts are indexed in
[the local qualification record](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-LocalQualification/README.md).
The following dated checkpoints preserve earlier results; only the checklist above reports current acceptance.

## Implementation and validation record

### Queue-after repeated fixture isolation

At `d404fc90`, x64 Debug builds with zero warnings/errors, and the pending-content,
removal-focus and bandwidth cases each pass ten repeats. Queue-after repeat 1 passes
in 1,610 ms. Repeat 2 reuses the same Dummy destination and waits for an unrelated
overwrite decision; a noninvasive debugger snapshot shows `WaitForConflictDecision`
under `CrossFileSystemBridge::PromptDestinationCollision`. The owned diagnostic
process is interrupted and its incomplete coverage retained; Full never starts.

The fixture now deletes only its own Dummy destination before each repetition,
accepts an absent directory, and reports any other cleanup failure. It keeps all
200 files, the bounded publication gate, and the queue/cancel/source assertions.
Both queue-after and don't-start require repeated validation before Fresh Full.
Production file-operation behavior and the DxUi pin are unchanged.

### ASan bandwidth duration fixture correction

The frozen ASan File Operations run reports `Phase6_LocalBandwidthThrottle` at 3,531 ms
for a four-second ideal copy, giving 469 ms lead against a 250 ms allowance. The old case
is byte-identical on master `0dd0bd6d` and migration `a20f6680`. It records the first UI
observation of `HasStarted()` as the start, although the worker already publishes
`_operationStartTick` before executing the operation. Late UI observation therefore clips
the measured interval and can falsely report excess throughput.

The candidate reads that worker timestamp and deliberately defers first observation by
at least 500 ms while continuing to pump messages. It logs the observation delay separately,
resets the per-run progress quantum, and preserves all duration, sampled-throughput and
cancellation assertions. Production copy/throttle code is unchanged. Focused native and
final Full validation remain pending; the original failed result remains evidence.

### ASan pending-content fixture correction

Frozen `a20f6680` ASan Debug run `20260913T054610Z-95756-605d08f0a71e45ac91d7d81bac52cdab`
reports `content_pending_elided` failing in 657 ms: expected aggregate pending count 1..200,
observed zero. `GetOrComputeDecision` applies completed worker updates before returning its
initial snapshot, so a fast comparison can legitimately finish first. This is a fixture race,
not an AddressSanitizer memory diagnostic. Full has completed; its eight failures remain recorded above.

The candidate fixture uses the existing worker progress callback to hold real content workers
while checking the pending snapshot. Queue notifications remain unblocked. It retains all
existing pending/final assertions and the scaled 20-second completion budget, reports gate
timeouts explicitly, and joins workers before destroying captured callback state. Production
comparison code and the DxUi pin are unchanged. Focused and final validation remain open.

### ASan Commands interruption and borrowed fixture names

The same frozen ASan run completes 632 Commands cases, then exits with incomplete inventory;
the terminal record reports count/coverage mismatch and nonzero exit. The last saved passing
case is `cmd_pane_fileops_popup_global_summary_ignores_finished_tasks`. The next registered
case, `cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs`, calls the
test-only `FolderView::DebugValidateRemovalFocusContractsForSelfTest` helper.

That helper retains `FolderItem::displayName` views from temporary name vectors in two
fixtures. Those vectors die before later view operations. The candidate supplies named
vectors that outlive each view and preserves the large-folder fixture's owner through
view teardown. No production ownership or removal-focus behavior is changed. The isolated
`a20f6680` case reproduces an ASan heap-use-after-free (sanitizer exception `e073616e`).
The dump and matching executable/PDB hashes are retained; the debugger's generic stack
does not resolve the C++ allocation/free sites. The corrected repeated case still needs
validation. Full has completed; its original partial Commands results remain evidence.

### Release diagnostic correction and local driver isolation

Frozen `3a43186a` Release Full completes with **2,108 passed, five failed and 56 skipped**;
19/20 entries pass. Commands is the only failed entry (893/5/2); Compare passes 239/0/33,
File Operations 179/0/21, all 770 tooling tests and 12 build-tooling checks pass, and both
viewer entries and performance entries pass. All 14 post-Full module identities match.
The [failed Release record](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Release3a43186a/README.md)
retains every case and skip reason.

The five ViewerText missing-snapshot failures reproduce in isolation (0/5/0). Snapshot
handlers/counters and the host preview bridge were behind `_DEBUG`, although the cases
run in test-enabled Release. Correction `836f4f97` uses the established `ENABLE_TESTS`
gate. Its Release build passes in 4m57s with zero warnings/errors; the same five cases
pass (5/0/0) with byte-identical fixture sources. All 14 post-test module identities match.
The [before/after record](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-ViewerTextReleaseDiagnostics/README.md)
retains the proof; no assertion or deadline changed.

Frozen `a20f6680` then passes its Release root build in 4m48s with zero warnings/errors
and the portable A/B/A rehearsal. Its first Full finishes with **2,105 passed, eight failed
and 56 skipped**; 19/20 entries pass. Only `tooling.pester` fails (762/8/0). The combined
local driver built ARM64 in `RsI19`, then ran x64 in `RsI19Closeout` in the same PowerShell
process; same-named modules from both checkouts make Pester's module-scoped checks ambiguous.
All 12 build-tooling checks pass. Compare completes 239/0/33, Commands 898/0/2 and File
Operations 179/0/21; both viewer entries and performance pass. All five corrected ViewerText
cases pass in Commands, and all 14 post-Full module identities match their actual bytes.
The [complete failed run](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Releasea20f6680/README.md)
remains failed. The [isolated probe and driver snapshots](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-LocalDriverIsolation/README.md)
reproduce the ambiguity with both checkout modules and verify lookup succeeds with one.

The replacement driver starts each x64 profile in a separate clean PowerShell process,
rejects preloaded repository modules, and preserves the frozen source and separate receipts.
Release, ASan Debug with detection probe, and Debug run serially. No product code or test
assertion changed for this retry. Current completion is tracked only in the opening checklist.

### Final x64 qualification: retained attempts before 3a43186a

Completed `b2be60e4` Full is retained as failed: **2,116 passed, one failed, 52 skipped**;
19/20 entries pass. Debug builds in 3m57s with zero warnings/errors. Commands passes 898/0/2,
Compare 242/0/30 and File Operations 180/0/20; both interactive viewers and performance pass.
ViewerWeb close passes at 29,192 microseconds with zero provider timeouts; direct menu hover
passes in 157 ms. All 14 post-Full module identities match their bytes and exact pin.
The sole failure is the new compiler-host fixture's Git clone exceeding MAX_PATH in Full's
sandbox. The corrected fixture uses the existing local long-path clone policy, explicitly tests
a greater-than-260-character hook path, and passes all eight dependency tests. The following
`3a43186a` Release build, portable A/B/A rehearsal and Full attempt are recorded above;
its failed Full stops that sequence before ASan Debug or Debug. Its only executable change
after `b2be60e4` is this test fixture. Archive validation,
impact governance and specification inventory pass. The failed run is not waived.
`f12acdba` builds Debug with zero warnings/errors. Completed Full run
`20260912T224851Z-84920-8e116d4c1cd6429aa91567f3fb44e182` passes 19/20 entries:
**2,115 cases pass, one fails, 52 skip**. Both interactive viewer entries and performance entries pass.
Commands passes 898/0/2 (passed/failed/skipped), including persistent menu hover. File Operations
passes 180/0/20. The one failed ViewerPE entry is ViewerWeb's sub-500 ms blocked-provider close
assertion. The exact-duration/provider-timeout diagnostic at `cc7595de` compiles and passes ten
isolated repetitions: close takes 27-145 ms, with zero provider timeouts. This does not explain
the original failure. The complete noninteractive viewer group also passes; its ViewerWeb close
takes 32 ms with zero provider timeouts. The full failed attempt is retained under `f12acdba-final-x64-Debug-Full`
in I19 evidence. The `ad7f7735` Debug root build passes with zero warnings/errors; Full exits 3
before tests because an archived `.props` receipt is not an active build input. Its bytes are
preserved with a `.props.txt` suffix, and the unchanged impact validator is checked before retry.
The impact check passes with no unmapped build inputs. The subsequent `b2be60e4` Full result
and its test-only long-path correction are recorded above. All three final x64 profiles remain open.
Earlier re-attested receipt: `2f9c32dea9013838f71ab2be42f5d5791c0200e265c8f6e318e0ce38ade6ed27`.

### Retained validation checkpoints

- Adopt one canonical `DxUi.lib`, public headers and product adapters; remove all 27 old
  library files and 28 old test/project/baseline files from the migration branch.

- Rebase onto RedSalamander master `0dd0bd6d`; preserve its Terminal, build and dependency fixes.

- Implement accepted G4 clipboard semantics and font refresh in canonical DxUi; retain the
  library regression tests and six-profile build evidence. RedSalamander pins `13788e95`;
  RedXe retains `ecad707`.

- Preserve dependency identity checks, explicit test-retirement dispositions, product source
  guards and durable ownership documentation. All 14 built Debug consumer module hashes match
  their pin sidecars; package and loaded-module proof remain separate.

- Prepare header guards, 16 satellite translations, the public checkbox repaint regression
  and live Preferences row geometry in `codex/dxui-i19-closeout`. Header fixtures pass all three
  x64 profiles; nine localization contracts pass. These fixes are integrated at `751437d2`.

- Complete and retain the `f7b2bce5` Fresh Full Debug run as a failed attempt.
  Compare: 241/1/30; Commands: 893/5/2; FileOps: 178/1/21 (passed/failed/skipped).
  FileOps stopped at `R4A19_DiscoveryIndependentVolumes` because the runner omitted
  `-AllowAlternateVolumeTestRoot`; the existing machine manifest already selects marked
  `D:\RedSalamander.Perf`. The focused repeat with that root enabled passes 6 checks, zero failures
  and zero skips. All 15 other Full entries pass, including both interactive viewers and performance suites.

- Repeat the four clipboard/menu cases three times with shuffle seed 190912: 8 pass, 4 fail,
  zero skips. Address-bar copy fails all three attempts; Find fails one destination-hover check.
  Batch Rename and Issues-pane copy each pass all three attempts. Raw diagnostic archives are
  retained under `f7b2-commands-clipboard-menu-repeat3-artifacts` in the I19 evidence directory.

- Integrate prepared fixes at `3d8bdd54`; full Debug compilation succeeds. The public checkbox
  regression passes twice. One array-initializer warning has a prepared brace correction.

- Qualify clipboard fixture quiescence at `751437d2`: the full Debug build has zero warnings/errors
  (3m22s), address-bar copy/paste/cut passes all 10 repeats, and the combined five-case clipboard/
  Preferences group passes all 15 attempts with zero skips. All five observed desktop clipboard
  readers remain running. The prior failed attempts and `clipboard-fixture-quiet-comparison.json`
  remain retained; product actions/assertions and library clipboard code are unchanged.

- Focused harness/localization checks: 197 pass, zero failures; specification inventory:
  zero blocking findings. The final WIL scope-guard cleanup and exact ViewerImgRaw close-latency
  diagnostic compile in the full `310cee2c` x64 Debug rebuild; runtime qualification follows.

- Build frozen `751437d2` in all three x64 configurations with zero warnings/errors: Debug 3m22s, Release 5m11s, ASan Debug 4m06s. The native ASan heap-overflow probe exits 1 with the expected AddressSanitizer diagnostic; its executable hash and log are retained.

- Complete Fresh Full Debug on frozen `751437d2`, retaining the failed result:
  `20260912T203600Z-34704-cf9cbc3084704e7a86f61bbeda368d87`. **18/20 entries pass;
  2,114 cases pass, two fail, 52 skip**. All raw archives and final evidence are preserved under
  `completed-751437d2-Full-Debug`. Commands: 897/1/2; FileOps: 180/0/20; Compare: 242/0/30.
  Both interactive viewers, standalone suites and performance entries pass. All earlier clipboard,
  Preferences, header and localization failures pass.

- Resolve the two retained failures in isolated checks: persistent menu-bar direct hover
  passes ten repetitions; the corrected `TestViewerImgRawLatestWinsExactReaderAndCloseSafety`
  fixture passes ten Debug and ten ASan Debug repetitions at `f12acdba`. The failed Full receipt
  remains retained; final Fresh Full qualification is still required below.

- Integrate the diagnostic/WIL cleanup and ARM64 host default into `codex/dxui-i19` at
  `50d2c4e5`. Its x64 Debug (4m04s) and Release (5m11s) full builds pass with zero warnings/errors;
  ASan Debug (3m56s) also passes with zero warnings/errors, and its seeded heap overflow is detected.

- Qualify the targeted ViewerImgRaw fixture correction in x64 Debug. The unchanged plugin
  DLL fails eight of ten attempts with the original fixture: reads reach their 10-second emergency
  timeout before state inspection, and one failed attempt also records a 676 ms close. At `f12acdba`,
  the canonical one-message pump and explicit progress-delivery checks pass all ten attempts with
  zero read timeouts; closes take 2-7 ms. Median fixture duration falls from 30.70 s to 1.12 s.
  All scheduler/provider/close assertions and deadlines remain unchanged. This is a fixture correction,
  not a claim of faster product code. ASan Debug also passes all ten repetitions; the standalone
  viewer test target compiles in Debug, Release and ASan Debug.

- Persistent menu-bar direct hover passes all ten isolated repetitions at `50d2c4e5`, with
  no failures or skips. The earlier failed Full attempt remains preserved.

- Matched x64 Release Preferences resource runs at master `0dd0bd6d` and candidate `50d2c4e5`
  pass all 40 repetitions of the byte-identical Plugins-page round-trip fixture. Median observed peak
  private bytes rise 0.384% (about 1 MiB); working set falls 1.071%. Last-sampled CPU changes go in
  opposite directions across the pairs; no CPU performance gain or regression is established.
  Raw 100 ms nominal samples, all four runner receipts, fixture hash and limitations are retained in
  `rs-preferences-resource-pairs` under I19 evidence. GPU/idle/frame claims remain separate.

- Final code checkpoint `f12acdba` passes its full x64 Debug build in 4m16s with zero
  warnings/errors. Fresh Full started with verified receipt
  `5c7a36d01c220d172d77936c5d732f101b8bab2b83596c35304cf3d2b2c0f148`, run
  `20260912T224851Z-84920-8e116d4c1cd6429aa91567f3fb44e182`. Release and ASan Debug
  root builds and Fresh Full runs follow serially only if the preceding profile passes.

- ARM64 Debug full build succeeds in 4m52s with zero errors and 44 warnings (29 C4324, 15 C4746); these warnings remain visible for review.

- Diagnose ARM64 Release x86-host heap exhaustion. Real MSBuild evaluation confirms all three
  ARM64 profiles now default to the x64 host on x64 Windows; all six profiles preserve explicit
  overrides. Dependency/tooling checks pass 36 tests. The x64-host retry clears the heap failure
  but encounters PDB-server errors, also seen in the concurrent closeout Release build. Sequential
  Release retry now passes in 4m37s with 60 warnings and zero errors; its module sidecars confirm
  the x64 compiler host and exact library pin. ASan Debug passes in 3m54s with 46 warnings and zero
  errors. All six profiles have now built at `751437d2`; the failed attempts remain retained.

- Complete RedSalamander's three ARM64 cross-builds at `751437d2`; final compiler-host and
  diagnostic changes still require their own build receipts.

- Repeat the complete canonical DxUi x64 Release suite without concurrent build or windowed
  test activity: all 18 suites exit successfully on 2026-09-13, with nine interactive-desktop
  capability skips in Menu and zero skips in the other 17 suites. Reports and logs are retained
  under `dxui-quiet-Release-20260913` in I19 evidence. The skipped paths are not qualified by
  this run, and functional success does not replace paired performance acceptance.

- Frozen `751437d2` x64 Release portable ZIP passes clean-extraction application startup and
  built-in plugin smoke. Its 11 packaged DxUi module sidecars match the exact ZIP entry hashes and
  `13788e95` pin. Test-enabled package: 30,692,818 compressed bytes, 79,757,185 expanded bytes;
  `751437d2-Release-package-provenance.json` and the ZIP are retained in I19 evidence.

These dated/checkpointed observations retain the original failures and then-current next steps.
Only the checklist above reports current acceptance; historical pending wording is not an additional
open task or a requirement to repeat already-qualified work.

- `8d12f6c3` targeted Debug ViewerPETests/ViewerImgRaw builds pass. Fifteen fresh-process
  scheduler cases retain 14 passes and one first-run close-latency failure. All fifteen report the
  expected one active/one pending/replaced-once state and zero fixture read timeouts. The first two
  attempts take 5.98/5.21 seconds; later attempts take 1.20-1.71 seconds. These overlap the Full
  Commands run and do not establish quiet latency acceptance. A diagnostic showing the exact close
  duration is prepared; the 500 ms limit is unchanged. Run subsequent windowed diagnosis serially.
- Master `0dd0bd6d` Debug scheduler case passes all three isolated attempts; frozen `751437d2`
  ASan Debug passes all three. These differing profiles are separate observations, not a matched
  performance comparison or dismissal of the original Full Debug failure.

- `3d8bdd54` Debug rebuild: zero errors, one C5246 warning in the new checkbox initializer
  (brace correction prepared). Native checkbox replacement: two passes. Preferences/navigation
  repeat three: 3 pass / 3 fail / zero skips. Retained logs and raw archives use the `3d8bdd54-`
  prefix under I19 evidence. Observer run `20260912T202339Z-99736-c28d9f2b6f1548fdbd8f1a129426c01c`
  records 41 ownership transitions with no dropped samples; it never opens or reads clipboard data.
  Observed readers include RadeonSoftware, logioptionsplus_agent, JDownloader2, explorer and svchost.
  The nominal observer polling interval is 1 ms; timestamps retain actual scheduling intervals.
- Qualify the fixture quiet-point correction by repeating the same command assertions; no
  library source or product copy implementation changes as part of this fixture correction.


- First post-rebase Full Debug finishes with five failed entries and fifteen passing entries.
  Authoritative per-entry results and identities remain under the run
  `20260912T190303Z-38780-be28e4e68ee241ee8bf0ce63971c925c` in `evidence/runs`.
  The runner removes disposable runtime directories when their owning process exits; copy raw
  diagnostic archives into retained I19 evidence before starting a new runner process.


- Full Debug Commands completes 893 passed / 5 failed / 2 skipped in 21m19s.
  Failures: the category-churn coordinate defect (correction prepared), Batch Rename preview copy,
  Issues-pane Ctrl+C, address-bar edit copy, and Find result menu-state inspection. Retain the original
  results under run `20260912T190303Z-38780-be28e4e68ee241ee8bf0ce63971c925c`, Commands archive
  `2026-09-12_214527`; reproduce the unresolved cases before attribution or acceptance.

- Diagnose the category-churn failure from the retained Full trace: after the tree scrolls
  21.33 DIP, the fixed Viewers point selects Editors. Prepare a test-only live row-rectangle
  adapter and use clipped client bounds for actual clicks; retain all focus/repaint assertions.
- Compile and qualify the category-click correction. Clipboard-copy and Find result
  context-menu failures remain under investigation and visible in the original receipt.

- Full Debug Compare completes 241 passed / 1 failed / 30 skipped; the MTP runtime cases pass.
  The sole failure is a mixed native source guard reading retired `DxUi.Controls.cpp`.
- Qualify the prepared replacement in `search_low_hardening_smoke`: preserve all product source
  checks and exercise checkbox keyboard repaint through the pinned public API. Both directional
  keys must clear a mixed glyph without changing its Boolean value and must request repaint.
- All 14 built Debug consumer module hashes match their exact DxUi-pin sidecars after Full's
  tooling re-attestation. Receipt: `f7b2-module-identities.json` under I19 evidence. This proves
  on-disk build identity; package and loaded-module qualification remain separate.

- Restore the upstream font/configuration messages in all four shipped satellite languages:
  16 translations across Czech, French, Japanese and Slovak. All nine resource-localization
  contract checks pass, including complete IDs and exact placeholder equivalence. Native satellite
  compilation follows when the prepared closeout commits are integrated.
- Full Debug also retains the original satellite-parity failure; the prepared fix changes
  resources, not assertions. Receipt: `font-config-satellite-contracts.log` under I19 evidence.

- Fix the rebased standalone package-fixture compilation: shared sandbox/test headers now guard
  existing `WIN32_LEAN_AND_MEAN` and `NOMINMAX` definitions. The unchanged relocated native fixture
  passes x64 Debug, Release and instrumented ASan Debug with warnings treated as errors. Receipts:
  `header-guards-relocation-{Debug,Release,ASan-Debug}.log` under the I19 evidence directory.
- Integrate the prepared header correction after the active Full run finishes. Its first tooling
  lane retains the original C4005 failure (11 passed / 1 failed); remaining product suites continue.
  That Full run cannot qualify the candidate even if its runtime suites pass.

- Candidate `f7b2bce5` Debug build passes with zero warnings/errors in 4m09s. The R4A19 case passes
  twice with shuffle seed 190912 while all 65 caller journal files retain their hashes. Receipt:
  `Z:\RedSalamander.Perf\evidence\I19-20260909\f7b2-R4A19-journal-isolation.json`.
- Fresh Full x64 Debug completed (failed) at `f7b2bce5`, following verified build receipt
  `5d6226b606fca8b5a915dfe84070d5122508eda929701de84d5688778a5b0d47`.
  The candidate stayed frozen throughout. Remaining profiles, resources and package/rollback follow.


- Reproduce FileOperations R4A19 journal contamination on the rebased binary: default profile fails
  directory creation with `0x800700B7`; the identical executable passes with fresh local-data state.
  Receipts are `rebase-ba72-R4A19.log` and `rebase-ba72-R4A19-clean-profile.log` under I19 evidence.
- Qualify the focused runner correction: each FileOperations child and classification retry gets fresh
  governed LOCALAPPDATA for its entire lifetime. Parent environment/journals remain untouched.
  A behavioral process-launch test covers independent retries, other suites and exit-code propagation;
  88 runner, fingerprint and impact checks pass. Fresh product runtime verification follows at the
  committed candidate.


- Prepare C7 source retirement in the isolated closeout checkout: remove all 27 legacy implementation
  files and 28 old test/project/baseline files. Production projects already import the external archive.
  Product source guards and runtime cases remain with their product owners; shared-only guards have
  explicit [retirement dispositions](DxUi_SharedLibrary/RetirementDispositions_2026-09-12.md).
  Add ownership guards and update audits, impact rules, active specs and developer guidance.
- C7 focused validation passes 261 checks with zero failures; tool inventory passes and the
  specification inventory reports zero blocking findings. No historical Done evidence was changed.
- Integrate C7 at `e7244091` into `codex/dxui-i19`. The migration checkout now has no in-tree DxUi
  implementation or old control-test project. Fresh behavior, performance, packaging and native gates remain open.

- Rebase RedSalamander migration onto master `0dd0bd6d` after the user reports compilation repaired.
  Rebase completed at `7d0b095e`; current local candidate is `ba72f601`. The prior head is retained in `codex/dxui-i19-before-rebase-20260912`;
  prepared reconciliation and evidence are checkpointed at `2d68f526`. DxUi and RedXe already contain main.
  Upstream Terminal, dependency, header and build fixes are preserved. The cancelling automated-format/revert
  pair has no net C++ changes and was omitted while retaining its workflow changes.
- Local qualification only, per the user's instruction: do not launch or depend on CI for this pass.
  Run x64 Debug/Release/ASan Debug tests and all ARM64 builds; distinguish native ARM64 execution from
  cross-compilation. No new hosted validation has been launched. Preserve earlier CI receipts as history.
- Candidate pin advances to local DxUi `13788e95f1c9a16210a87988d742d9dfc6f66eef` for the upstream
  Terminal font-refresh API. Restore the exact clean commit from the canonical local repository through
  a process-scoped Git transport mapping; public repository identity remains in the lock. Publication and
  release qualification remain open; this local restore does not prove remote availability of the new pin.
- Rebased dependency/workflow/inventory/impact checks pass: 53 tests, zero failures. The exact local
  font-refresh commit restores cleanly; the temporary Git transport configuration is restored afterward.
  All 144 changed production/test C++ files were checked with clang-format 22.1.3; four needed formatting.
- Build the rebased test-enabled Debug product locally: zero warnings/errors in 4m20s at ba72f601.
  Its newer vcpkg baseline/tool `1391150` is restored
  in an isolated owned checkout; the developer's vcpkg checkout is untouched. Fresh full x64 qualification
  and ARM64 builds follow. Keep the candidate source frozen while receipts are being produced.
- MTP isolation at b97bbc1a builds in the complete Debug solution with zero warnings/errors and passes
  the saturated-journal reproduction without changing any of the caller's 65 journal files. Two repeats
  and a separate two-repeat shuffled run (seed 190912) each pass 118 checks with only two physical-device
  skips. Evidence: `Specs/TestRuns/Local-x64/CompareDirectories/2026-09-12_083303-DxUiMtpIsolation/`.
- Implement the incoming Terminal font-refresh helper in canonical `DxUi/Typography.h`, with tests
  for stale negative answers and per-factory/all-factory invalidation. All 18 x64 Debug and ASan Debug
  suites pass. Release passes 17 suites; Menu fails once, then passes separately, as does the retained
  baseline. Five further consecutive Menu attempts pass with the same binary, including that focus case.
  The initial focus failure remains recorded and unclassified. All six gallery images reviewed;
  only indeterminate progress animation pixels differ. ARM64 Debug/Release/ASan Debug cross-builds pass
  with zero warnings/errors. Source hashes and all local receipts are retained in DxUi
  `Measurements/SharedLibrary/2026-09-12-font-refresh/`; native ARM64 execution remains unqualified for this change.
  Library checkpoint `13788e9` is committed locally and remains unpublished; RedSalamander's local candidate
  pins it and RedXe remains at ecad707.
- Qualify and publish the new library revision before product release; complete native ARM64,
  product regressions, package/rollback and paired resource gates remain open.

Scope: **all I19 implementation, tests and specs authorized on 2026-09-09**. Current work is local qualification
of the rebased product migration, legacy-source retirement, packaging/rollback and resource acceptance.
Historical CI receipts below retain their original scope; the user requested no CI for this pass.

- Preserve completed ecad707 Full results (2026-09-12 review): Debug 2,073 passed / 40 failed /
  53 skipped; ASan Debug 1,723 passed / 80 failed / 97 skipped. Raw results and traces are retained in
  `Z:\RedSalamander.Perf\evidence\I19-20260909\completed-ecad-Full-{Debug,ASan}`. Neither run qualifies.
- Remaining failures include MTP fixture state, clipboard/focus checks, Terminal/ViewerSpace readiness,
  and FileOperations. ASAN ran under severe memory pressure; reproduce cases in isolation before attribution.
- Native CI failure causes identified: ambient Pester 5 rejects current test semantics, re-attestation
  loses the prior step's vcpkg integration, ARM64 ASAN selects Release MSIX inputs, and evidence upload mixes
  external and repository roots. Prepared corrections pass all 24 release/workflow policy checks.
- Qualify the corrected product on all six profiles locally, per the user's no-CI instruction;
  MTP/Monitor compile and targeted Debug runs pass. Native ARM64 execution requires a suitable local host.
- RedXe a812cec at DxUi ecad707 passes all six native CI profiles (34374884161), including the repeated
  Process Viewer lifecycle correction and native ARM64 sanitizer detection.
- Isolated repair checkout `Z:\RsI19Fix`: formatter policy and MTP case-isolation source guards pass
  222 focused checks. All sixty MTP registrations now enter a fresh journal sandbox per attempt and restore
  the prior environment afterward. C++ compilation, saturated-journal proof and repeated/shuffled MTP coverage now pass in x64 Debug.
- The unchanged original Debug executable fails `mtp_property_fetch_is_batched` with the copied 64-slot
  saturated journal and passes with fresh scratch state. Both receipts are retained as `mtp-baseline-poisoned`
  and `mtp-baseline-clean` under the I19 evidence folder. This establishes a pre-existing isolation failure.
- Correct MonitorTest progress fields to seconds. Native Debug execution completes successfully in 4.59s;
  displayed elapsed times and rates use seconds. ETW transport/Monitor resource qualification remains separate.
- New candidate qualification proceeds locally at the user's request. Prior RedXe CI at 47096dd was
  refused before execution because of account limits; those historical failed starts do not count as tests.
  The preceding six-profile receipt applies only to a812cec. Native ARM64 execution requires an ARM64 host.
- Resume quiet resource measurements on 2026-09-12 with no product processes and about 25 GiB free.
- Run eight sequential RedXe AV measurements with identical original/candidate fixtures in Debug and
  Release. Raw reports, identity checks and comparison script are in RedXe `Measurements/DxUiAdoption/2026-09-12`.
- Accept paired resources: all runs retain bounded 4,608,000-byte surfaces and zero hidden work, but
  Release dirty throughput is 11.9-14.4% lower in both candidate pairs. Baseline timing also varies widely;
  investigate before attributing or accepting the signal. No thresholds or baselines were relaxed.

### Original acceptance targets

These targets describe the intended finished migration. Their current completion status is tracked
only in the opening checklist; the list below does not claim that all acceptance gates have passed.

- Source gap analysis: origin records, live direct-header/project consumers, post-extraction changes, named test dispositions and baseline PNG comparison documented.
- C0 runtime/module closure, targeted reproductions, test replacement mapping, and retained baselines complete.
- User decisions recorded: use `DxUi.lib`; no G5 legacy compatibility work; standalone Slider behavior is authoritative; six build configurations are mandatory; G4 separates native clipboard capacity while retaining validation and responsiveness.
- One canonical library implementation and public consumption path; all product modules/adapters accounted for.
- G4 native large-selection copy/paste and failure cases pass while the embedded text ceiling remains enforced; no UI-thread retry sleeps.
- Every project supports Debug/Release/ASAN Debug on x64/ARM64, with matching native instrumentation, working sanitizer detection and native ARM64 qualification.
- DxUi standalone matrix and both product candidate gates provide current, traceable evidence.
- External dependency changes invalidate build/test receipt reuse and are visible in shipped provenance.
- RedSalamander migration preserves required product interaction, accessibility, localization, settings and resource behavior while adopting the accepted standalone Slider and diagnostics changes.
- Advisory notices and the manual upgrade/fix/retest loop work; consumer CI and rollback are rehearsed.
- Required hardware/manual/native-platform gates pass; any out-of-scope capabilities remain explicitly unqualified under their existing owner.
- Legacy implementation removed only after complete migration; docs/specs/tooling/index closeout complete.

### Earlier library implementation checkpoints

- G4 behavior approved: native large selections, validated memory/Unicode, no retry sleeps, separate embedded ceiling.
- Existing implementation, tests and owning contracts inspected.
- Retain matching Debug/Release baselines before library edits (including retained noisy and repeated Release runs).
- Add a regression that fails on the current native clipboard ceiling.
- Implement native capacity separation and failure-safe handling.
- Verify native boundary/large selections and unchanged embedded rejection in x64 Debug NativeTextInput; remaining configurations stay open.
- Update library specs, usage docs and focused regression evidence; final matrix and paired acceptance remain open.
- G4 additionally passes x64 ASan Debug NativeTextInput with a working sanitizer detection probe.
- G1 tooltip fix and G7 offscreen Grid selection fix pass focused x64 Debug suites; Embedded remains green.
- G2 neutral public helper headers compile/link in the standalone consumer test; exact-pin relocation and product adapters remain open.
- Native module collision reproduced: DLL copies lost animation and used the executable's menu class. Module-owned registration at DxUi ecad707 passes all 12 external EXE/two-DLL probes, rendering and ten negative pin/build checks in all six native CI profiles (34371134464), including both ASAN annotation policies. Local x64 Debug/Release/ASan Debug suites pass. Consumer candidates now pin this correction; product qualification is in progress.
- G3 x64 Debug localized empty/no-match rendering and geometry regressions pass.
- G8 standalone six-configuration mapping and native sanitizer detection qualify at 3dae073 (CI 34355763655); consumer matrices remain open.
- Complete required matched performance acceptance. All fourteen quiet repeats and 32 alternating CPU-subset repeats are retained. Repeated-baseline comparisons still cross thresholds; the prior Debug composition signal drops to +0.61% across the eight paired medians. Results remain inconclusive; no threshold or baseline was relaxed.

### Earlier consumer implementation checkpoints

- Qualify both consumer candidates at DxUi ecad707, which corrects native menu/animation ownership in statically linked plugin modules. Previous consumer pins and receipts remain retained.
- RedXe ecad707 passes complete local x64 Debug/Release/ASan Debug suites. A cumulative Process Viewer counter assertion exposed in duplicate CI is reproduced by repeating the lifecycle in-process; corrected delivery-delta checks pass all three local suites at a812cec. The six-profile a812cec receipt is retained; later fixture CI is blocked by account limits.
- RedSalamander Fresh Full Debug and ASAN completed with the failures recorded at the beginning of this checklist. The corrected ViewerSqlite helpers and CI exit propagation are included; complete product acceptance remains open.
- RedXe unchanged x64 Debug and Release suites pass; previous Release package retained before adoption.
- RedXe candidate 8c548fe passes full x64 Debug, Release and ASan Debug suites, including sanitizer detection and module provenance.
- RedXe candidate upgraded to DxUi 52da33d passes full local x64 Debug, Release and ASan Debug suites with zero build warnings/errors and all six native CI profiles at RedXe 1bf6662 (34368430957), including ARM64 sanitizer detection.
- RedXe native x64/ARM64 Debug, Release and ASan Debug CI passes at 2ef0bcb (run 34359693390), including product tests, real sanitizer detection and provenance. Manual/hardware and paired resource acceptance remain open.
- DxUi native x64/ARM64 suites, external fixtures, both ASAN annotation policies and gallery pass all six configurations at 3dae073 (CI 34355763655).
- Relocated ASAN consumer with RedSalamander's STL compatibility policy passes rendering and ten negative pin/build checks.
- RedSalamander original Debug and Release Full/Fresh baselines captured. Neither is green: Debug has one tooling failure, three Commands failures and one FileOperations failure; Release has one tooling failure, ten Commands failures and one FileOperations failure. Full reasons are retained in `Z:\RedSalamander.Perf\evidence\I19-20260909\baseline-results.json`; diagnosis and clean comparison remain open.
- RedConfigureTests Debug links the exact external archive, with zero warnings/errors and an attested module sidecar.
- Focused dependency tests reject malformed pins, missing/dirty source, wrong archive inputs and incompatible identities.
- RedConfigureTests x64 Debug passes all 42 cases, including the actual Japanese satellite and all twelve combo-box/TagPicker inputs. App and remaining matrix/resource qualification stay open.
- Monitor x64 Debug builds with the external archive and zero warnings/errors; runtime/resources remain open.
- Full RedSalamander x64 Debug solution builds against 3dae073 with zero warnings/errors, including main, every viewer, Terminal, Monitor, RedConfigure and PerformanceTests2.
- Remove obsolete ARM64 ASAN rejection from runner/deployment preflights and make native Commands inventory use its selected profile. All six accepted profiles and invalid-name rejection pass in 54 focused runner-plan tests. Native ASAN execution remains a separate CI gate.
- CI failure diagnosis: e8e5e606 ASAN compilation failed in ViewerSqliteTests because shared theme-test UIA helpers were hidden from the sanitizer build. A subsequent matrix command incorrectly replaced the build exit code. Corrected helpers compile in x64 ASAN with zero warnings/errors; 19 build-evidence/workflow checks pass, including a real failing-script exit propagation test. Fresh native Full qualification remains open.
- Full product regressions and the remaining five native profiles pass. Candidate Debug exposed 53 Commands failures before a modal-test stall. A focused Preferences run reported 137 pass/39 fail; debugger snapshots prove retained dialog state and destroyed HWND registrations. The Preferences wheel subclass removes its saved procedure before forwarding WM_NCDESTROY, bypassing owner cleanup. Fix and regression qualification are in progress; the 128-window library bound is unchanged.
- Reproduce the Preferences leak with a 48-cycle regression: the unfixed build retains 10 hosts/attachments after its first close versus a baseline of 7. Preserve the failing trace in `Z:\RedSalamander.Perf\evidence\I19-20260909\preferences-lifetime-before-context.json`.
- Corrected Preferences subclass forwarding and page-host teardown build in the full Debug solution with zero warnings/errors. The 48-cycle regression passes: both counters return to 7 after every close, versus 10 after the first unfixed close. Paired retention evidence: `Z:\RedSalamander.Perf\evidence\I19-20260909\preferences-lifetime-retention.json`.
- Correct the second teardown defect: pane models/delegates now disconnect before page-host Detach destroys their controls. The expanded 48-cycle regression rotates through all fourteen categories and returns hosts/attachments to 7/7 after every close. Preferences family: 176 passed, zero failed, one foreground-dependent skip; focused PowerShell 7 harness checks: 202 passed. Retained evidence: `Specs/TestRuns/Local-x64/Commands/2026-09-09_154457-DxUiPreferences/`.
- Qualify the teardown fixes in Fresh Full and the other five native profiles; family diagnostics do not replace those gates.
- Initial candidate Full run passes FileOperations (180/0/20 skipped), CompareDirectories (242/0/30 skipped), ProductUi, plugin contracts, both noninteractive viewers and other standalone suites. Two long-path dependency fixtures and the interactive ViewerPE class-name expectation require harness corrections; Commands and final Full acceptance remain open.
- ProductUiTests x64 Debug builds and passes all 23 retained product cases plus two viewer adapter regressions. CI/Full runner mapping and 48 runner-plan, seven inventory, and 40 build-evidence/workflow tests pass.
- Sandbox cleanup is scoped to its fixture IDs; 48 plan-helper and four dependency/profile tests pass before the product runner migration.
- Product source-policy guards rehomed: NavigationView input and credential modal-quit, child-pointer teardown and joined UIA workers; 193 source-policy checks and 19 tooling-governance/impact checks pass. The remaining 61 library dispositions are recorded in DxUi Specs/Testing/SourcePolicyDispositions.json at 76e3d48: runtime replacements, retired helper spelling/G5 diagnostics and explicit ownership reviews. Four mixed runtime cases and the corrected Menu interactive fixtures pass all six native configurations at DxUi 52da33d (CI 34363702073).
- Relocated native sandbox regression passes Debug/Release/ASan Debug on x64 and reproduces the original package-helper failure with the old header. Complete package smoke remains open.
- CI managed-checkout root cause identified: d0b3067 native jobs all stop before compilation because root `vcpkg/` violates source snapshots. Stage the tool/cache in `.build/vcpkg-tool`; 57 fingerprint/build/dependency tests pass, including a real nested-Git fixture. New native CI remains pending.
- Existing receipt reuse checks reject dirty/missing dependency source; linked-module sidecars are attested and packaged. Full Debug build emits all selected module sidecars; package and full-suite qualification remain open.

Work proceeds in `Z:\RsI19` on `codex/dxui-i19` while the original checkout supplied the completed baseline runs. This checklist is synchronized to the
requested original WIP folder; bring the implementation back after candidate validation.
DxUi anonymous Git and Actions API access is verified: neither consumer needs `DXUI_READ_TOKEN`.
The six-job RedXe workflow uses the automatic job token for advisory API requests. Native CI at RedXe 2ef0bcb passes all six configurations (34359693390) after correcting a Windows code-integrity test assumption and replacing a personal Codex-dependent validator with repository-owned validation. No native or hardware qualification is inferred from cross-compilation.

Status: **ACTIVE, I19 — RedSalamander migration rebased onto master; all legacy DxUi source/tests retired
on the migration branch. The opening checklist reports the current final qualification gates;
retained earlier passes do not qualify later corrections.**

Priority: **P1**. Effort: **L**, delivered in bounded changes. Risk: **High** for RedSalamander migration; **Low–Medium** for the advisory version check and consumer upgrade tooling.

Owner: this plan owns the cross-repository adoption sequence, dependency qualification, and RedSalamander migration. DxUi owns shared controls and library contracts; each product owns its adapters, behavior, packaging, and release decision. The user authorized all I19 implementation, tests and spec updates on 2026-09-09. G4 is implemented; current work qualifies the rebased migration locally. Publication remains a separate lifecycle step.

Planned on 2026-09-09 against these clean local checkouts. These are audit identities, not new dependency pins:

| Repository | Local checkout | Audited commit |
|---|---|---|
| [DxUi](https://github.com/RedSalamanders/DxUi) | `Z:\src\DxUi` | `d192e474e540adc2656e69b8e8150b0f7a05d63f` |
| RedXe | `Z:\src\RedXe` | `75545887d002b5075d170d98b7bfd3a77d8d7f61` |
| RedSalamander | `Z:\src\RedSalamander` | `3f0df2f2c28239eb658a68f85b9c7f69c2356af4` |

This is a **non-normative proposal**. Existing contracts remain in force until implementation, tests, and authoritative updates land. Local paths are discovery context; build automation must support relocated checkouts.

## 1. Accepted direction

1. Keep **one canonical DxUi implementation**, built from its repository. Extend RedXe's existing exact-pin consumption; migrate RedSalamander from its in-tree implementation.
2. Use **`DxUi.lib`**. The user selected static linking on 2026-09-09 after reviewing the sizing estimates. DLL implementation and experimentation are outside this plan.
3. Use the simple consumer-driven update loop: merge DxUi with green library tests, show an advisory new-version notice in consumer builds/PR checks, then deliberately update that consumer's pin, build and run regressions locally, and push/merge its PR after consumer CI passes. If DxUi causes a failure, fix and test it upstream, merge it, then update the consumer pin and repeat.
4. Use **two independent compatibility gates**: library behavior tests in DxUi, and candidate-versus-current product tests in RedXe and RedSalamander. Neither replaces the other.
5. Preserve a tested previous pin and product package for rollback. Track which DxUi revision each built and shipped module actually contains; an updated source lock alone is insufficient evidence of adoption.
6. **Accepted after the gap audit:** adopt current DxUi diagnostics/scheduling without legacy compatibility work (G5). The standalone DxUi Slider is the correct behavior (G6); adapt consumers and test expectations to it.
7. **Accepted G4 clipboard policy:** native copy/paste supports large selections independently of the embedded editor's 65,536-unit ceiling. Keep allocation-bounded reads, Unicode/size validation, explicit allocation/clipboard failure and the single open attempt; no UI-thread retry sleeps. Preserve the separate embedded edit/snapshot ceiling.
8. **Mandatory for every project:** Debug, Release and ASAN Debug builds on both x64 and ARM64 across DxUi, RedXe and RedSalamander. Deliver missing support; ASAN is not optional and must not silently map to unsanitized Debug.

The first useful milestone is one RedXe upgrade through this loop plus a RedSalamander consumer/drift inventory and retained baseline. Each product adopts on its own schedule; normal builds continue using its existing exact pin.

### 1.1 Required build matrix

| Configuration | x64 | ARM64 |
|---|---|---|
| Debug | Required | Required |
| Release | Required | Required |
| ASAN Debug | Required | Required |

Apply this matrix to every library, application, plugin, sample and test project in the three repositories, with consistent mappings for resource/package projects. Preserve the established MSBuild name `ASan Debug` where used. ASAN builds must instrument native project code and link the matching instrumented DxUi archive; a configuration label or ordinary Debug fallback does not satisfy the requirement.

Implement and verify project/solution mappings, compiler/CRT/sanitizer runtime, build/test CLI accepted values, dependency restore/imports, isolated outputs/fingerprints and CI/test coverage for all six combinations. Native ARM64 execution qualifies runtime behavior. Missing toolchain/runtime/runner support is a delivery dependency to resolve, not an optional skip or a waived platform. This is the accepted target; the audited current code does not yet establish it.

## 2. Initial audit and gap routing

The September 9 [codebase gap analysis](DxUi_SharedLibrary/GapAnalysis_2026-09-09.md) accounts for all 58 origin records, 79 live consumer source files with 96 direct DxUi includes, and the inherited test cases. The table below preserves that initial audit and its original follow-up routing. The tooltip, public-header, localization, clipboard and Grid gaps are now implemented; G5 needs no legacy preservation and G6 uses the accepted standalone Slider. The opening checklist reports current qualification, including the completed six-configuration builds and still-unverified native ARM64 execution on the final local revisions.

| Evidence inspected on September 9 | Finding at that revision | Original follow-up |
|---|---|---|
| DxUi [consumption contract][dx-build], [project][dx-project], and [capabilities][dx-capabilities] | One static target, public headers, exact API/target validation, consumer MSBuild imports, isolated outputs. Native and embedded mechanisms already exist. | Reuse these entrypoints; do not repeat extraction or introduce consumer-maintained source lists. |
| RedXe [lock][rx-lock], [integration contract][rx-integration], [restore][rx-restore], and `Build/RedXe.DxUi.props` / `.targets` | RedXe already pins the audited DxUi HEAD. `RedXe`, `AVControl`, and `AVControlTests` reference the archive. Host and plugin keep DxUi C++ ownership inside their respective modules. | Add an advisory version check and rehearse a deliberate pin-update PR using existing tests. Preserve host/plugin ownership; no initial integration rewrite is needed. |
| RedXe [AV plan][rx-av-plan] | Synthetic text/UIA and WARP tests exist. Real IME, screen-reader/touch, matched text/UIA resource acceptance, and AV release gates remain open. No checked-in `.github/workflows` YAML was found in this checkout. | Establish product CI around existing tests; confirm actual repository check configuration. Keep AV release work with its existing owner. |
| RedSalamander [legacy project](https://github.com/DualTail/RedSalamander/blob/3f0df2f2c28239eb658a68f85b9c7f69c2356af4/Common/DxUi/DxUi.vcxproj) and its project references | Production references found in RedSalamander, RedSalamanderMonitor, RedConfigure, ViewerText, ViewerSqlite, ViewerSpace, ViewerImgRaw, ViewerPE, ViewerWeb, and Terminal; test references in DxUiTests and RedConfigureTests. | The source audit adds ViewerVLC as a header-only typography consumer and maps all direct DxUi includes. Qualify the transitive build/runtime closure and actual loaded modules during C0/C4/C5. |
| RedSalamander [legacy public header](https://github.com/DualTail/RedSalamander/blob/3f0df2f2c28239eb658a68f85b9c7f69c2356af4/Common/DxUi/DxUi.h) | Includes the product's `PlugInterfaces/Viewer.h`; implementation also uses consumer `WindowMessages.h`. | Replace consumer coupling with adapters; never move viewer/plugin contracts into DxUi. |
| DxUi [origin record][dx-origin] and RedSalamander Git history | Extraction records a September 5 source baseline. Subsequent RedSalamander changes include tooltip clock handling and test updates in `817b7e2d`, `f0710f9b`, and `3f0df2f2`. | Audit disposition: 817b7e2d tooltip host/dispatcher fix and regression are missing upstream; f0710f9b adds product UTF utility coverage; 3f0df2f2 adds the product runner input warning. Resolve the shared fix in DxUi and retain the product cases/runner policy. |
| DxUi [validation contract][dx-testing] and [CI][dx-ci] | Foundation, inherited controls, embedded WARP, public consumer fixtures, galleries, and native x64/ARM64 Debug/Release jobs are defined. | Inspect actual receipts when qualifying a candidate. A workflow definition or historical pass is not a fresh execution result. |
| RedSalamander [test coverage](../../Testing/Testing_TestCoverage.md), [suite plan](../../../Tools/Modules/Testing/TestSuitePlan.psm1), and [validation evidence](../../Testing/Testing_ValidationEvidence.md) | Full already includes DxUi families and product integration tests; build/source/runtime receipts govern reuse. | Extend this machinery for external DxUi inputs; do not create a competing test runner or count old-library tests as new-library proof. |
| DxUi [migration HOLD][dx-rs-hold] and consumption prose | RedSalamander migration remains on HOLD. A trailing paragraph still says RedXe needs its first pin/bridges, contrary to the inspected consumer code and newer integration contract. | At implementation activation, reconcile stale prose and make the HOLD record route to this consumer owner. Do not claim real IME/AT completion while correcting the pin status. |

Before each implementation slice, repeat `git status --short --branch` and compare its owning paths against the audited identities:

```powershell
git -C Z:\src\DxUi diff d192e474e540adc2656e69b8e8150b0f7a05d63f..HEAD -- include src Build Tools Specs
git -C Z:\src\RedXe diff 75545887d002b5075d170d98b7bfd3a77d8d7f61..HEAD -- Dependencies Build RedXe Plugins Tests Specs
git -C Z:\src\RedSalamander diff 3f0df2f2c28239eb658a68f85b9c7f69c2356af4..HEAD -- Common RedSalamander RedConfigure RedSalamanderMonitor Plugins Tests Tools Specs
```

Inspect worktree changes separately. Use focused branches; never reset another checkout. Keep the human-managed scratch note outside this work.

## 3. Packaging decision: DxUi.lib

### 3.1 Decision and scope

**Accepted by the user on 2026-09-09: use the static library.** Build `DxUi.lib` from an exact source pin with matching public headers and compiler/CRT settings, then link the dependent EXEs/plugins. This matches the existing supported consumption path and avoids adding an exported binary interface and runtime dependency.

The sizing inspection below informs the choice: Release static duplication is a few MiB across the product, while most control/model/graphics allocations would remain with either packaging. Retain those estimates as decision context, not a future DLL work queue. A later packaging change would require a separate decision.

### 3.2 Validation still required for static adoption

Retain a baseline of the current in-tree/static implementation and compare the shared-library candidate using the same product scenarios and fixtures. Preserve existing performance/resource gates for startup, input/frame latency, preparation/composition, allocations, cache/surface retention and idle work. Check cold/warm build behavior and unintended rebuilds during integration; add optimization work only for an observed problem. This is migration validation, not a prerequisite to reconsider the accepted packaging choice.

### 3.3 Initial estimate from current RedSalamander binaries (2026-09-09)

This is a sizing estimate for the current **in-tree RedSalamander DxUi**, using x64 Release binaries from the September 9 build at the audited RedSalamander commit. It is not a measured static/DLL comparison, a measurement of the standalone DxUi candidate, or completion of C0/C1. No product was launched and no DLL was built for this inspection.

The existing build receipt records `tests_enabled=false`, compiler `19.51.36256.0`, and receipt ID `3a4434b69e1af54e524fdddbddf549e6ad4acbbaecaafc973861886c4674cf80`. All ten inspected PE file hashes match that receipt, and each PE CodeView GUID matches its PDB. The library compiler log contains `/O2 /GL /Gy /MD`; the inspected ViewerText linker log contains `/OPT:REF /OPT:ICF /LTCG:incremental`. Dead-code removal is already active in that consumer.

Read-only method: use `llvm-pdbutil dump --modules --section-contribs` on each matching PDB, select object modules under `Common/DxUi`, and sum the union of their contribution ranges in each section. Parse only the final section-contribution table, excluding the module-summary contributions to avoid double-counting. Read PE file sizes separately. Attribution covers `.text`, `.rdata`, and `.data`; linker-created unwind/relocation records, alignment, cross-module inlining, and shared template attribution prevent treating these sums as exact removable bytes. See [LLVM's PDB inspection tool](https://llvm.org/docs/CommandGuide/llvm-pdbutil.html).

| Current Release module | Whole file, MiB | PDB-attributed DxUi code/data, MiB |
|---|---:|---:|
| RedSalamander.exe | 6.561 | 0.644 |
| RedSalamanderMonitor.exe | 1.267 | 0.353 |
| RedConfigure.exe | 1.381 | 0.537 |
| ViewerText.dll | 1.270 | 0.523 |
| ViewerSqlite.dll | 0.945 | 0.506 |
| ViewerSpace.dll | 1.045 | 0.418 |
| ViewerImgRaw.dll | 1.046 | 0.463 |
| ViewerPE.dll | 0.924 | 0.463 |
| ViewerWeb.dll | 1.635 | 0.462 |
| Terminal.dll | 0.606 | 0.009 |
| **Total, these ten files only** | **16.680** | **4.380** |

Exact totals are 17,489,920 file bytes and 4,592,672 attributed bytes. The Release archive itself is **121.461 MiB**, a build artifact containing compiler/link inputs; it is not shipped as 121 MiB in every consumer and does not represent RAM use. Debug is a separate, non-shipping comparison: its archive is 95.046 MiB and the main executable is 77.724 MiB. Debug sizes should not drive the Release packaging decision.

For a first DLL estimate, taking the largest observed contribution of each DxUi object/section across consumers gives a **0.685 MiB proxy for one copy**. This is neither a true union of required functions nor a DLL size prediction: consumers retain different subsets, and exporting the public API changes optimization. Allow roughly **0.8–1.5 MiB for a Release DLL**, then remaining inline/template code and adapters in consumers. The resulting decision ranges are:

| Cost | Current static delivery | DLL scenario estimate | Expected difference |
|---|---|---|---|
| Installed DxUi-related code/data across these modules | Approximately 4.4–5 MiB, including a rough allowance beyond direct attribution | Approximately 1–2 MiB across the DLL and retained consumer code | **About 2.5–4 MiB less disk space**; medium confidence. This is not the total installed application or compressed package size. |
| Resident code/read-only memory in ordinary use | Depends on pages touched, not archive size or all mapped image bytes | One copy can serve modules in the process | **Budget around 0–1 MiB saved**, low confidence; main-only use may be neutral or slightly worse. |
| Resident code/read-only memory after heavily exercising several viewers, optionally Monitor/RedConfigure | More duplicated implementation pages can become resident | Same-version DLL image pages can be shared where eligible | **Perhaps 1–3 MiB saved**, low confidence; not an expected reduction in private heap or GPU memory. |
| Direct writable global data | 23,817 attributed bytes across these ten modules, excluding heap allocations | Fewer copies within one process; separate data per process | Only tens of KiB of direct globals; heap/resource retention needs separate measurement. |
| Control/model heap, per-window textures and swap chains | Determined by active controls/windows | Substantially the same for unchanged ownership | **Assume zero saving** from packaging alone. |
| CPU/FPS, threads, timers | Current rendering/input work | Same work; DLL requires no new worker/timer | Expect broadly neutral steady-state cost; do not promise an FPS gain. |
| Implementation-only development edit | Can relink ten dependent production modules | Potentially relink only the DLL if the interface/import library stays stable | Potential developer-time benefit; no timing delta measured. Header/API changes still rebuild consumers. |

The memory ranges assume that only part of each roughly 0.4–0.65 MiB implementation copy is resident in ordinary use, and more overlapping code is exercised in a heavy session. They are illustrative scenario budgets, not confidence intervals from live sampling. DLL image sharing does not share ordinary heap objects or globals across processes. [Microsoft DLL data rules](https://learn.microsoft.com/en-us/windows/win32/dlls/dynamic-link-library-data).

One resource observation retained from the comparison: [WindowHost graphics sharing](https://github.com/DualTail/RedSalamander/blob/3f0df2f2c28239eb658a68f85b9c7f69c2356af4/Common/DxUi/DxUi.WindowHost.cpp) stores D3D11/D2D/DWrite resources in a static map keyed by thread ID. Static linkage gives each module its own map. A common DLL could consolidate compatible same-thread buckets across host/plugins, avoiding some duplicated device/context/factory resources. It would not share buckets across different threads/processes or remove per-window back buffers. No driver-memory saving is priced into the estimates above. As scale context only, two 1920×1080 BGRA8 buffers contain **15.82 MiB** of pixels before driver overhead, regardless of library linkage; this is not a claim that every current control allocates full-screen buffers.

**Decision following this estimate:** use `DxUi.lib`. The estimated DLL savings do not justify additional packaging and ABI work for this adoption. The resource estimates remain unpaired and do not waive migration performance validation.

### 3.4 Measured migration package footprint (2026-09-12)

Matched **test-enabled x64 Release** packages of master `0dd0bd6d` and migration
`751437d2` both pass clean-extraction startup and built-in plugin smoke. The migration
adds **1,109,627 expanded bytes (1.06 MiB, 1.41%)** and **403,353 compressed bytes
(0.38 MiB, 1.33%)**. Its eleven packaged DxUi modules match their pin sidecars.
The unchanged master ZIP is retained as a locally tested rollback artifact.

[The retained comparison](../../TestRuns/Local-x64/DxUiAdoption/2026-09-12-ReleasePackaging/README.md)
binds both ZIP hashes, module sizes and provenance. This is the complete migration
footprint, not a static/DLL experiment or a measurement of ordinary non-test shipping
builds. It does not establish live memory or timing acceptance.

### 3.5 Final portable footprint and rollback (2026-09-13)

Frozen `3a43186a` passes its x64 Release build in 4m48s with zero warnings/errors and the
portable A/B/A rehearsal against master `0dd0bd6d`. All three clean-extraction startup,
plugin-contract and runtime-closure checks pass. All 11 candidate sidecars match the
packaged module bytes and exact DxUi pin. A1/A2 payloads are byte-identical per file.

| Test-enabled x64 Release package | Baseline | Candidate | Change |
| --- | ---: | ---: | ---: |
| Compressed bytes | 30,289,465 | 30,694,026 | +404,561 (+1.34%) |
| Expanded bytes | 78,647,558 | 79,758,721 | +1,111,163 (+1.41%) |
| Main EXE bytes | 28,911,616 | 29,184,000 | +272,384 (+0.94%) |

The [final packaging record](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-FinalPackaging/README.md)
retains original logs, comparison and hashes; the ZIPs remain in the local I19 evidence folder.
This measures static-library migration, not an isolated DLL comparison. It does not qualify
installed update/settings, minimum-OS or native ARM64 behavior. Final Full/runtime and resource
gates remain in the opening checklist.

### 3.6 Corrected final portable footprint (2026-09-13)

The final test-enabled Release package at `a20f6680` includes the ViewerText diagnostic
gate correction verified by the five unchanged regression cases. Against baseline `0dd0bd6d`:

| Measure | Baseline | Candidate | Increase |
|---|---:|---:|---:|
| ZIP bytes | 30,289,465 | 30,701,720 | 412,255 (1.36%) |
| Expanded file bytes | 78,647,558 | 79,772,033 | 1,124,475 (1.43%) |
| Main EXE bytes | 28,911,616 | 29,184,512 | 272,896 (0.94%) |
| Files | 171 | 182 | 11 module-provenance sidecars |

All three A/B/A clean-extraction checks pass; baseline payload hashes match before/after,
and all 11 candidate module identities match their packaged bytes. The
[reviewed receipts](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Packaginga20f6680/README.md)
retain the exact ZIP identity, comparison and original logs. This measures the test-enabled
static migration; it is not a DLL experiment, ordinary shipping-build or runtime-memory result.

## 4. Dependency lifecycle and release workflow

**Accepted approach, 2026-09-09:** library PR, advisory availability notice, deliberate consumer upgrade, local build/regressions, consumer PR/CI, merge. Product maintainers drive adoption. The tooling described here is planned, not implemented by editing this document.

### 4.1 The update loop

1. **Merge DxUi:** implement the change and its library regression tests in a DxUi PR. Run all required DxUi checks, including applicable performance/docs/gallery checks, and merge. An available upgrade means an exact merged `main` commit with successful required CI for that revision; a pending or failed merge build is not advertised as ready. A commit SHA is enough to identify the version; release tags are optional.
2. **Notify in the consumer:** RedSalamander/RedXe builds and existing PR checks compare their pinned SHA with the available DxUi revision and show an advisory notice. They still compile the pinned revision. No dependency file changes automatically.
3. **Upgrade locally:** create or use a focused consumer branch, change `Dependencies/DxUi.lock.json` to the selected exact SHA, then use the normal restore/build/test commands. Restore into that consumer's isolated dependency output; do not reset the developer's sibling DxUi checkout. Include any necessary adapter changes in the same consumer branch.
4. **Run regressions:** build and run the consumer's required regression and performance gates. Use the current pin as the retained baseline when comparison is required. Library-green means the library tests passed; the product tests establish whether this product can adopt it.
5. **If green:** push the consumer change and open/update its normal PR. Record old/new DxUi SHA and local validation results in the PR. Consumer CI validates the actual PR change before merge; if the tested inputs change, rerun the affected required checks. Merge and ship through the normal product process.
6. **If red:** identify whether the defect belongs to the consumer adapter or DxUi. Fix an adapter defect in the consumer branch. For a DxUi defect, add a reproducing regression test and fix in a DxUi PR; after green tests and merge, update the consumer branch to the new merged SHA and repeat steps 3–5. Keep the product regression too. The consumer default branch remains on its previous working pin throughout.
7. **If a regression appears after adoption:** revert the consumer upgrade, including paired adapter edits if necessary, and rebuild/test the previous pin. Repair the issue through the same loop. Keep the previous shipped package available for normal product rollback.

Once an upgrade starts, test a fixed candidate SHA. A newer DxUi merge may produce another notice but does not invalidate a passing upgrade or require chasing a moving target. RedXe and RedSalamander can adopt independently; one product's failed upgrade does not change the other product's pin or block unrelated consumer work.

### 4.2 Where the notice belongs

| Place | Behavior |
|---|---|
| Root consumer build/restore entrypoint | Check once per invocation, not once per project/compiler step. Show current and available short SHAs, a compare link, and the lock path to update. Use a short bounded lookup and cached result where useful; show when the result was last checked. |
| Existing PR validation | Run when a PR opens or is updated, so the maintainer sees the notice before deciding to merge. Put it in the build log/check summary; no automatic PR comments are needed. |
| Existing default-branch build after merge | Reuse the same check if that build already runs. It is a useful reminder but should not be the first place an upgrade is noticed. |

Example: `DxUi update available: pinned <old-sha>, available <new-sha>. Update Dependencies/DxUi.lock.json on a branch and run the product regression suite.`

This is an **advisory build notice**, not a compiler warning or failing MSBuild diagnostic; `/WX` must not turn it into a failure. An unavailable network/API, missing read access, or a rate limit reports **update check unavailable**, not “up to date”, and does not fail a build whose pinned dependencies are already available. A real failure to restore/validate the pinned dependency still fails normally. The check uses existing read access and needs no cross-repository write permissions.

Verify that the candidate is newer along the normal `main` history before calling it an upgrade. A same pin produces no notice; a pending/failed candidate is not recommended; a divergent pin is reported for review rather than silently replaced. Test these cases plus offline behavior, fixed-pin builds, and one notice per invocation. Keep lookup behavior in one reusable helper rather than duplicating it in build scripts and CI.

The deliberate tradeoff is that an idle product may not notice a new DxUi version until its next build or PR. That is acceptable for this workflow. A maintainer can still initiate an upgrade immediately for an urgent fix. There is no periodic monitoring service or adoption deadline in this plan.

### 4.3 Small set of rules to keep

- **Exact pin and normal source build:** keep the existing lock shape and library MSBuild imports. Build matching headers and `DxUi.lib` from that pin. No source enumeration or floating branch dependency.
- **Isolated, compatible outputs:** preserve per-consumer/configuration/architecture outputs and existing source/toolchain/CRT/build-flag identity checks. Changing the pin must invalidate stale artifacts and test receipts. All three repositories require the six configurations in section 1.1, including matching instrumented ASAN Debug libraries on x64 and ARM64; never silently substitute unsanitized Debug.
- **Existing tests and evidence:** keep the library and product suites, required supported-platform coverage, and current result/receipt formats. Add the external DxUi identity to missing build/test inputs rather than introducing a second qualification system.
- **Compatibility notes and rollback:** describe observable/API changes and any required adapter edits in the DxUi PR. Preserve minimum supported Windows behavior and normal notices/symbols/packaging requirements. Each product can retain its previous tested pin while an upgrade is repaired.

Automatic lock-update PRs, cross-repository pre-merge testing/dispatch, a per-product qualification catalogue, new receipt schemas, scheduled reconciliation, a dedicated binary-package cache, and a new version/support-line program are not part of this implementation. Revisit only a demonstrated maintenance problem; ordinary source restore, build, test and PR review are sufficient to start.

## 5. Tests owned by DxUi

Extend the existing suites and fixtures; do not replace them with a new framework. Map each supported capability and every inherited test disposition to an owner. Shared behavior discovered through a consumer defect should gain a minimal synthetic reproduction in DxUi while the product retains its end-to-end regression.

| Layer | Required evidence |
|---|---|
| Public API/build boundary | Compile/link/run consumers using public headers only; exact-pin negative cases; no application includes, consumer source lists, or incompatible flags. Preserve the standalone public samples and relocated fixture. |
| Control semantics | State/layout, event order and counts, preview/commit/cancel, validation/read-only/disabled behavior, keyboard traversal, pointer capture/cancel, nested/reentrant callbacks, deletion during callbacks, Unicode editing, grids/trees and virtualization. |
| Native hosting/input/UIA | HWND lifetime, nested modal/popups, focus restoration, DPI, IME/TSF text-store revisions, clipboard abstraction and G4 large native selections (including 100,000 UTF-16 units), independent embedded overflow rejection, failure-without-text-loss and no retry sleeps, stale/disconnected UIA providers, selected offscreen items, destruction and queued callbacks. |
| Embedded rendering | Supplied device/context, hostile pipeline state, clean/dirty/hidden/zero-size states, clipping/alpha, multiple views, failed preparation, device loss/recovery, independent pools and generation teardown. |
| Visual contract | Existing test baselines plus gallery coverage of relevant themes, high contrast, reduced motion, DPI and states. Review intended pixel changes; semantic assertions remain mandatory. |
| Resource contract | Paired complex-UI baseline/candidate, preparation versus composition, zero clean-composition allocations, cache/surface/queue bounds, hidden idle work, failure/recovery peaks and retention. Hardware presentation remains a separate gate. |
| Test/tool integrity | Case registration/disposition accounting, result/skip parsing, negative fixtures, test-owned data paths, and detection of missing suites or forged/stale receipts. Run required ASAN Debug coverage on x64 and ARM64, prove sanitizer detection with a controlled fixture, and retain applicable fault/stress coverage. |

Library CI must build and test Debug, Release and ASAN Debug on x64 and ARM64. The existing external-consumer fixture currently selects x64; extend it to the required six-configuration matrix, with native ARM64 execution instead of reporting an unexecuted fixture as green. Required unavailable capabilities produce a blocked gate with a reason. Optional skips remain visible and cannot erase a formerly failing test.

For native tests that need an interactive desktop, reserve that resource and keep deterministic fixtures separate from human acceptance. Cover real IME, screen readers, touch, system clipboard, and hardware/device behavior in a named release checklist when affected. Synthetic messages and WARP cannot establish these results. Do not alter normal user settings, clipboard, audio defaults, camera state, or application data through unattended fixtures.

When a test harness changes, rerun the previous implementation with that same harness before comparing. Use the existing independent [performance contract][dx-perf]; standalone reviewed evidence belongs in DxUi `Measurements`, while product measurements stay in the respective product repository.

## 6. Tests owned by each product

### 6.1 RedXe

Use [test.ps1][rx-test], `HostPluginTests`, `AVControlTests`, and `PluginContractTests` as the integration gate. Every RedXe project must support the six configurations in section 1.1. Build the host and every affected plugin/test against the same candidate pin and matching configuration, and verify that the launched binaries came from those outputs. Adopt the standalone Slider contract directly; product tests verify its wiring rather than restoring legacy semantics.

Protect tile and raised-overlay lifecycle; preview/commit/cancel; mouse/keyboard/touch routing; focus transfer; host TSF/clipboard/UIA bridges; clipping and DPI; page changes; resize and device loss; hidden/zero-size surface release; multiple widget instances; shutdown with outstanding callbacks/providers; and unchanged non-DxUi widgets. Keep host-owned scheduling and the COM/POD ABI intact. Test settings read/write and packaging using synthetic AV services and isolated data.

Compare the current pin and candidate with the same product and fixture inputs. Record real presented latency, total preparation plus composition, allocation/resource peaks, idle wake-ups, and module retention after UIA publication. Candidate checks must neither silently waive existing AV HOLD gates nor turn new unrelated AV features into an adoption prerequisite. Current unqualified capabilities stay explicitly unqualified.

### 6.2 RedSalamander

Reuse [Run-AllTests.ps1](../../../Tools/Run-AllTests.ps1), its stable suite/case identities, [test-enabled build evidence](../../Build/Build_Toolchain.md), and the current [validation evidence contract](../../Testing/Testing_ValidationEvidence.md). The external lock, imported build tools, source/content and artifact hashes must enter build attestation, `source.dxui`, runtime closure, affected-set selection, and receipt invalidation. An external source change must not reuse receipts for the old in-tree library.

| Product surface | Protected behavior |
|---|---|
| RedConfigure and Preferences | Draft/apply/cancel/persistence, validation, localized strings, theme preview, control events, keyboard traversal, real focus after modal close, UIA close/reopen. Adapt the alpha slider to the accepted standalone Slider contract; intended Slider visual/event differences are not migration regressions. |
| Main window / FolderView | Selection, navigation, rename/edit, menus/popups, drag/drop and cancellation, Find/Compare dialogs, offscreen selected-grid UIA semantics, large/virtualized models and existing performance budgets. |
| File Operations UI | Existing confirmation, pause/cancel and prompt focus, progress updates, popup lifetime and teardown. Keep current file/provider outcomes and identity rules unchanged. |
| ViewerText, Sqlite, Space, ImgRaw, PE, Web | Open/close and repeated reopen, shared controls/chrome, search/navigation, focus, menus, theme/DPI, malformed/large input and cancellation using existing isolated fixtures. |
| Terminal | Prompt/tool-window focus, launch/close, plugin lifetime and existing input behavior. Preserve the admitted engine pin and current Terminal ownership. |
| Monitor | Control/host lifecycle, theme/DPI, text/menu input, resource retention, and declared selftest availability in test-enabled builds. |
| Packages | Fresh portable extraction plus MSI/MSIX install/update checks as applicable, all dependent EXEs/plugins present, correct architecture and dependency identity, startup on minimum supported OS. |

During migration, keep each inherited case accounted for: shared runtime cases in DxUi, product/adapter cases in RedSalamander, or a reviewed exclusion with reason. Preserve case identity/provenance when moving tests. Convert old private-header/test-macro assumptions into library-owned tests or legitimate public diagnostics; do not enable consumer-only library flags to manufacture compatibility.

Use focused cases while developing, then a **Fresh Full** for admission/closeout, plus paired test-enabled Release performance with enforced budgets. Existing Debug Full is not Release performance proof. Require Debug, Release and ASAN Debug builds for every project on x64 and ARM64, with matching instrumented dependencies for ASAN and native ARM64 execution before claiming runtime qualification. Full with documented environment skips remains limited to those tested capabilities.

All product test data and runtime evidence must use the governed, already initialized, marker-owned `X:\RedSalamander.Perf` sandbox; `.build` holds build outputs and tool-generated inventories, not runtime test data. Archive reviewed receipts through the existing `Specs/TestRuns` workflow. Preserve the runner's profile leases, no-activation behavior, directed-input warning, and prohibition on killing independently launched applications.

### 6.3 Evidence in the existing consumer workflow

Use each repository's existing build/test receipts and the consumer PR description. Record the old/new DxUi SHA, tested consumer change, selected configurations, regression/performance results, and any required manual or environment gates. Add DxUi identity to existing build/test inputs where it is missing; no shared cross-repository receipt schema is needed.

Verify that a pin change rebuilds the library and affected consumer modules, invalidates incompatible prior results, and appears in existing build/package provenance. A green test of the old archive or a generic sample cannot qualify the updated product. Keep this check in the current runner/build machinery.

## 7. Implementation sequence and deliverables

The table records the implementation slices and their exit criteria. Implementation and source retirement are complete on the migration branch; the opening checklist identifies the remaining qualification and release gates. Planning observations and sizing estimates do not establish a test pass.

| Slice / owner | Work and concrete deliverables | Exit gate / dependency |
|---|---|---|
| **C0 — consumer and library maintainers** | Use the completed source gap audit; retain its 58-file/test dispositions and extend them to runtime/module evidence. Route G1–G3/G7/G8 gaps and the accepted G4 implementation; G5 requires no legacy preservation and G6's standalone Slider is accepted. Characterize tooltip/localization/clipboard and existing Grid UIA gaps, name current consumer owners, and retain in-tree product/library baselines and first cost receipt. | No unmapped consumer or unexplained behavior difference. Baselines reproducible. Required before changing binaries. |
| **C1 — build owners across the three repositories** | Deliver Debug/Release/ASAN Debug × x64/ARM64 for every project, with matching instrumented DxUi dependencies, solution/CLI/CI support and isolated configuration identity. Validate exact pins, public imports, toolchain/CRT and rebuild behavior through existing entrypoints. | All six build configurations and sanitizer detection qualify; public consumer fixtures pass and reject mismatches/stale outputs. No DLL experiment or new binary package system. |
| **C2 — DxUi test owner** | Fill the tooltip-clock, public-helper and localized-text gaps; implement accepted G4 native capacity separation with bounded validation, explicit failure, one open attempt and unchanged embedded text limits; coordinate the shared Grid UIA fix and public consumer fixtures. Retain native/embedded resource gates, existing receipts and the library PR record. | Required standalone matrix and paired evidence pass; no consumer checkout dependency. Depends on C0 gap mapping. |
| **C3 — RedXe owner** | Establish/reuse product CI and current-versus-candidate tests. Change the lock on a consumer branch and rehearse an upgrade/rollback with existing artifact provenance, preserving AV release claims. | RedXe gate and rollback receipt; actual pin matches linked and packaged modules. Can proceed with C2 after C0. |
| **C4 — RedSalamander build/adapter owner** | Add product lock, isolated restore/imports, attested dependency closure, mandatory sanitizer matrix, theme/message/viewer adapters, and direct adoption of current DxUi diagnostics without legacy compatibility layers. Pilot the separately launched RedConfigure process plus RedConfigureTests; keep a retained legacy baseline. | Pilot behavior and performance pass, clean/relocated build works, legacy and external definitions cannot enter the same PE. Depends on C0/C2. |
| **C5 — RedSalamander UI/plugin owners** | Migrate remaining process families in bounded changes: Monitor, then the main host with its viewer/Terminal plugins, header-only ViewerVLC typography, and integration tests. Update imports/includes/test dispositions and verify complete runtime module identity. | Fresh product validation, paired Release/hardware evidence, packaging and applicable manual acceptance. Depends on C4; coordinate active UI owners before touching their surfaces. |
| **C6 — consumer build owners** | Add the single advisory availability check to normal builds/PR validation and document the manual lock-update/fix/retest loop. Rehearse same-pin, newer-green, pending/failed, divergent and offline notices. | Each integrated product can notice an update, adopt a fixed SHA after local tests/CI, reject a regression, and revert. Can begin with RedXe during C3; RedSalamander follows C5. |
| **C7 — RedSalamander and DxUi owners** | Remove the obsolete in-tree implementation/project only after every consumer has moved; reconcile durable specs/docs, case ledger and HOLD routing, archive this plan through repository closeout rules. | No duplicate shared implementation, all supported consumers accounted for, required evidence current, deferred items assigned explicitly. |

Migration granularity is **process/ownership compatible**, not an arbitrary file count. Do not mix old/new DxUi definitions in one PE. Before allowing old/new modules in one process, prove window-class registration, message/token registries, COM/provider lifetime and callback ownership cannot cross incompatible implementations. If that cannot be proved, switch the host and all its DxUi-using plugins atomically as one process family. A temporary compatibility header may forward to public APIs, but must have an owner/removal gate and contain no second control implementation.

Until removal, a necessary fix to the legacy tree must have an upstream disposition in the ledger. Shared code is maintained in DxUi; application adapters stay in the consumer. No ongoing bidirectional source-copy workflow is introduced.

Coordinate with RedSalamander I4 for test ownership, I5 for performance infrastructure, I9/I10 for CI/tooling, I15 for Terminal, I18 for Preferences/File Operations, and H0 for confirmed shared accessibility residuals. This plan does not reopen retired plans or take over their unrelated work. At activation, the DxUi migration HOLD record should become a routing record to this owner rather than a second execution queue.

## 8. Verification and closeout

### 8.1 Existing entrypoints to reuse

These are reference entrypoints, not execution receipts. Current results are linked from the opening checklist. Run from the respective repository root with governed fixture roots and retained baseline paths; inspect current help before execution. The current pass uses local execution only, as requested by the user.

```powershell
# DxUi: run matched baseline/candidate measurements as described in its performance contract.
.\validate-skills.ps1
.\validate-specs.ps1
.\validate-dependencies.ps1
.\format.ps1 -Check
.\test.ps1 -Configuration Debug -Platform x64 -PerformanceBaseline '<retained-debug-receipt>'
.\test.ps1 -Configuration Release -Platform x64 -PerformanceBaseline '<retained-release-receipt>'
.\test.ps1 -Configuration 'ASan Debug' -Platform x64 -PerformanceBaseline '<retained-asan-receipt>'
.\build.ps1 -Configuration Debug -Platform ARM64
.\build.ps1 -Configuration Release -Platform ARM64
.\build.ps1 -Configuration 'ASan Debug' -Platform ARM64
foreach ($configuration in 'Debug', 'Release', 'ASan Debug') {
    .\test-consumer.ps1 -Configuration $configuration -Platform x64
}
.\test-consumer.ps1 -Configuration 'ASan Debug' -Platform x64 -DisableStlAnnotations
# For changed visuals: .\gallery.ps1 -PublishDocs, with visual review.

# RedXe: existing restore/build/test entrypoints, per selected candidate checkout.
.\restore-dxui.ps1 -Platform x64
.\test.ps1 -Configuration Debug -Platform x64
.\test.ps1 -Configuration Release -Platform x64
.\test.ps1 -Configuration 'ASan Debug' -Platform x64
foreach ($configuration in 'Debug', 'Release', 'ASan Debug') {
    .\build.ps1 -Configuration $configuration -Platform ARM64
}

# RedSalamander: full product admission, including test-enabled build evidence.
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Get-ToolInventory.ps1 -Validate
.\Tools\Test-TestRunArchive.ps1 -Inventory
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -Platform x64
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Release -Platform x64
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration 'ASan Debug' -Platform x64
foreach ($configuration in 'Debug', 'Release', 'ASan Debug') {
    .\build.ps1 -Configuration $configuration -Platform ARM64
}
```

Replace the quoted performance-receipt placeholders with matching retained baseline paths.
Run each repository serially and stop on a failed command. On a native ARM64 Windows host,
repeat all three test configurations and the DxUi external-consumer checks with `-Platform ARM64`;
cross-builds above are compilation evidence only. DxUi and RedXe test entrypoints run their
seeded sanitizer checks in ASan Debug. RedSalamander also needs its isolated
`PluginContractTests.exe --asan-seed-heap-overflow` detection receipt under the governed
artifact-operation lease, as demonstrated by the retained earlier local ASan probes.
Enable alternate-volume cases only when the governed local fixture roots support them.

The entrypoints now support all six configurations. Current implementation closeout requires the
section 1.1 matrix, Fresh Full product results, matched resource evidence and package/rollback proof.
This pass uses local execution at the user's request; cross-builds do not replace native ARM64 runs.
Validate links, WIP ownership and normative consistency with the specification inventory as changes land.

### 8.2 Durable authorities to update during implementation

- **DxUi:** [build/consumption][dx-build], [testing][dx-testing], [performance][dx-perf], input/accessibility, controls/layout and hosting contracts where behavior changes; `capabilities.json`, public docs and gallery as required. Update supported/pending status only with evidence.
- **RedXe:** [DxUi integration][rx-integration], performance/resources, plugin ABI and AV contracts, build/test guidance, and end-user docs if behavior changes. Retain the existing AV owner for unmet hardware/human gates.
- **RedSalamander:** [toolchain](../../Build/Build_Toolchain.md), [shared helpers](../../Core/Core_SharedHelpers.md), [DxUi design](../../UI/UI_DxUiWinUIDesign.md), [shared grid](../../UI/UI_DxUiSharedGrid.md), [test coverage](../../Testing/Testing_TestCoverage.md), [selftests](../../Testing/Testing_SelfTests.md), [performance](../../Testing/Testing_PerformanceValidation.md), [validation evidence](../../Testing/Testing_ValidationEvidence.md), [tooling governance](../../Testing/Testing_ToolingGovernance.md), affected UI/plugin and installer contracts, `AGENTS.md`, and developer documentation. Add a dedicated consumption contract if needed, with consistency-ledger anchors.

Any new `Tools/` command/module requires inventory, help, caller, focused-test, cache-closure and normative updates. On final closeout, update `Specs/NormativeConsistency.json`, move this plan to Done, remove its active row, and add the exact new Done path and positive/negative admission test required by [plan lifecycle policy](../../README.md). Do not modify frozen Done history.

### 8.3 Completion checklist and stop conditions

The authoritative progress checklist is at the beginning of this document.

Stop promotion of the affected slice for an unexplained behavior difference, missing/changed baseline, confirmed resource regression, untested required capability, stale/mismatched evidence, incompatible module ownership/ABI, or an occupied output owned by an independently launched application. Report the exact missing evidence and preserve the previous pin. Optimize, narrow the slice, repair the regression, or obtain advice for a measured tradeoff; never silently rebaseline or classify a failing required test as optional. Independent planning and unaffected work can continue.

[dx-build]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Build/Build_ToolchainAndConsumption.md
[dx-project]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/src/DxUi.vcxproj
[dx-capabilities]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/capabilities.json
[dx-origin]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Done/SourceImport/source-origin.json
[dx-testing]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Testing/Testing_Validation.md
[dx-ci]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/.github/workflows/ci.yml
[dx-perf]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Core/Core_PerformanceAndResources.md
[dx-rs-hold]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Plans/WIP/RedSalamanderMigration.md
[rx-lock]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/Dependencies/DxUi.lock.json
[rx-integration]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/Specs/Core/Core_DxUiIntegration.md
[rx-restore]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/restore-dxui.ps1
[rx-av-plan]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/Specs/Plans/WIP/RFC_Plugins_AVControl.md
[rx-test]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/test.ps1

Implementation checkpoint: the pilot's root build and module attestation pass, and its full focused
runtime suite passes. The first Japanese test correctly exposed two nested TagPicker inputs; the
product now supplies localization through the public logical control tree. All native project
configuration declarations are present; instrumentation and native execution are qualified separately.
The original Debug baseline recorded three Preferences page-settling failures and one Pester failure;
those precede the dependency switch and are retained for candidate comparison.

Baseline package check: clean-extraction PluginContractTests failed for the attempted original x64 Release
package. Publication correctly discarded the failed staging archive and preserved the existing July package.
That existing archive is not the current baseline; do not use its size or contents as adoption evidence.
The failure log is retained; diagnose the original assertion before claiming rollback qualification.

First RedSalamander native CI failed before compilation: Git rejected long paths in historical test evidence.
Enable long paths before checkout in all CI jobs and the reusable build. Native qualification remains open.

The existing CI auto-format job changed 341 files, including archived evidence. Its exact automated commit is
reverted in the follow-up; CI now checks only changed owned C++ files with a pinned formatter and read-only access.
The local Full/Fresh run remains isolated at 83b703b7 while CI infrastructure changes are validated separately.

Package blocker diagnosed: TestSupport tried to find a checkout before reading the explicitly marked root.
A clean extraction has no checkout, so the existing local-writer package probe failed before testing the provider.
The helper now validates the explicit root first and discovers a checkout only for its default; package runtime
verification remains open. DxUi 3dae073 passes all 18 local suites in Debug, Release and ASan Debug.

The relocated native sandbox regression passes x64 Debug, Release and ASan Debug; the same ASAN fixture exits 12 with the original header, proving the checkout-discovery defect. Full package verification remains open.

RedSalamander CI reaches dependency restore after the long-path/formatter fixes. Its old vcpkg registry cannot resolve the already-requested BLAKE3 1.8.7. Both registry/tool pins advance to 11159b9 (the first BLAKE3 1.8.7 update); all twelve requested versions and the exact AWS override exist there. Dependency floors and the AWS override are unchanged. Native rebuilds qualify the resulting dependency graph.
