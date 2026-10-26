# RedMonitor extraction in RedSalamander

Status: ACTIVE consumer qualification. The canonical implementation plan is
[RedMonitor extraction and all four adoptions](https://github.com/RedSalamanders/RedMonitor/blob/ej/redmonitor-adoption-20261023/Specs/Plans/WIP/RedMonitorExtraction_2026-10-21.md).
The current consumer contract is [RedMonitor](../../Core/Core_RedSalamanderMonitor.md).

- [x] Extract viewer/model/resources/tests with source origin and MIT notices into Z:/src/RedMonitor.
- [x] Merge current upstream master into the extraction branch and preserve generated changelog history.
- [x] Remove the six compatibility viewer/test/satellite projects and all solution/installer references.
- [x] Add one pinned vendored header/license and local-only verification plus explicit update tooling.
- [x] Normalize existing module-owned providers without changing plugin ABI or provider lifetime.
- [x] Add installed Release discovery, explicit --redmonitor and optional bare Debug asynchronous startup.
- [x] Keep Preferences on the Release 7.0 monitor settings path in every consumer configuration.
- [x] Remove viewer ownership from ZIP/MSI/MSIX, symbol packages and consumer test plans; retain SDK notice.
- [x] Pass tooling contracts, local client controls and native format/validators on the revised boundary.
- [x] Build six native profiles, pass the accepted focused x64 Debug/Release/ASan scope and actual app ETW probes; Full runtime reruns explicitly omitted by the user.
- [ ] Qualify clean app-only ZIP/MSI/MSIX and Preferences Apply/restart/live filter behavior.
- [x] Retain reviewed source-bound evidence, fold changes, commit and push this integration delivery; qualification remains open below.
- [ ] Close timed focus, native ARM64, physical IME/AT, paired performance and release handoff gates.

The earlier managed-viewer pin13 checkpoint passed an integrated Release compile and targeted ARM64
builds, but its package run was blocked by an independent nested Release build. Those receipts remain
historical; they do not qualify the revised boundary. Earlier source-contract failures and cancelled
runs remain preserved. No independently launched process has been killed.

The helper is installed separately from its own production Release ZIP. Ordinary consumer builds
verify only local header bytes and never rebuild it. Its settings identity remains compatible pending
the separately tracked S4 narrowing; that later work must not silently break Preferences.

Qualification checkpoint (2026-10-23): all three ARM64 builds passed after the compiler-scoped aws-c-common workaround. The first integrated Debug run passed its native suites and 1000 ordinary tooling controls, with one stale inventory assertion still requiring the extracted viewer latency executable. That assertion now requires viewer surfaces to be absent. Archive validation also exposed retained absolute fixture links: the RedMonitor negative control now retires its known junction in finally while retaining file bytes/logs/link metadata; Compare fixtures create relocatable relative symlinks and retire their test-only junction. Failed receipts remain failed. Revised final-source builds and unfiltered Full runs, including the owner's authorized desktop suites, remain required before closeout.

Owner decision (2026-10-23): keep the independently launched x64 Release instance open and defer
Release qualification and packaging. Do not replace that executable or terminate the instance.
Debug/ASan work and the authorized desktop suites may continue. The completed ARM64 Release
cross-build is distinct from the deferred x64 Release/runtime/package gates.

The unfiltered Debug attempt `20261023T112244Z-4640-da4038fa92944c668559fcfea1c5e779`
passed 17 build-toolchain controls, 1001 ordinary tooling controls, the standalone noninteractive
entries, Compare, and both monitor Preferences cases (live filter/mask and settings-file link).
Commands recorded 597 passes and five failures; FileOps recorded five passes, two failures and
315 unexecuted/skipped cases. The native viewer parent stopped at its exact desktop-restoration
guard. The saved window identity was no longer valid; the available evidence does not establish
why. Preserve that guard and the failed run. Several failures overlap the existing
[repository qualification investigation](Repository_QualificationFollowups_2026-10-22.md);
this overlap does not prove an unchanged baseline or qualify the broader repository.

Archive validation then identified additional absolute Commands fixture junctions. Those fixtures
now retire only their own links through scope guards, after their original behavioral assertions,
and retain target bytes. This repairs evidence portability, without changing product behavior or
accepting a failed test. Fresh Debug/ASan validation of that repair remains pending.

The repaired Debug batch-rename and shell-commands source families passed and archived cleanly
(`20261023T131132Z-73260-fb4ccab5983a430f91d5753175887591` and the following Shell run).
Two malformed focused filter invocations matched no cases and are failed invocation evidence,
not product tests. Exact-name monitor checks remain scheduled. The first ASan attempt compiled
but failed final output attestation because a May 2026 RedLauncherConsole.exe/PDB remained in
the profile; those exact obsolete files were moved and hash-verified under .build/cache-preservation.
No artifact guard was weakened. Current onboarding, build skills, installer/localization/testing
contracts and diagnostics docs now describe independent ownership; a new validation snapshot is required.

Final Debug component checkpoint: both exact monitor Preferences cases passed without skips in
`20261023T132916Z-13460-4170db32f62248f993515fd21133b07a`. The actual Debug executable passed real
ETW Unicode/severity/PID/TID/time and explicit absent-helper controls in
`20261023T133234Z-72824-cdf25e08e9464aa58edc8f96c6051498`. These are focused/component passes.
Revised ARM64 Debug/ASan cross-builds also passed. The ASan attempt
`20261023T133039Z-13460-35b5a85870b5475db55835b133ddd0a1` stopped before tests with MSVC C1090,
PDB API error 23 / RPC 0x000006BA in three translation units. It remains failed; its cause is not
established. Retry serially with lower MSBuild concurrency, without promoting that retry to a causal claim.

Final build review found that Common-only header verification missed a direct plugin with no Common
reference. Verification now covers each first-party ClCompile project, excluding resources and external
dependencies. Four actual MSBuild fixture controls exercise that boundary. The icon tool inventory now
describes its existing one-application output rather than an extracted monitor asset. Revised tooling
and build qualification of this final boundary remains scheduled; x64 Release/packages stay deferred.

All four direct-MSBuild controls passed with no skips in
`20261023T134230Z-98288-3fd0a5b3422447e3a6d5beab51a92537`, including rejection of modified header
bytes before the direct plugin's build marker is admitted. The resource-only/external exclusions
passed with the same deliberately damaged fixture. ASan retry uses serial profile execution and
two MSBuild workers, then the ordinary Fresh Full runner against its verified solution receipt.

## Delivery and failed qualification record — 2026-10-23

The Fresh ASan Full run `20261023T135152Z-75348-852232e6ff774e55bec17597e760a3f7`
executed all 18 entries with no reused or invalidated entries: 12 promoted, six failed. The aggregate
records 2222 case passes, 21 failures and 349 skipped/unreached cases. Commands recorded 938 passes,
15 failures and six skips; both monitor Preferences cases passed. Compare recorded 249 passes and
26 prerequisite skips. FileOps recorded three passes, two failures and 317 skips/unreached cases.
ViewerPE failed both groups, PluginContract exited nonzero without an established cause, and the
artifact-writing deployment control failed during a compiler-PDB RPC error. Ordinary tooling passed
1002 controls; SQLite's two groups, FileSystemCurl, the other standalone entries and all 15
PerformanceTests2 cases passed. These observations do not establish baseline equivalence or a
performance non-regression. The native input, restoration, filesystem and archive guards remain intact.

The nested deployment writer had reconstructed only `RSBuildEnableTests`, omitting Full's existing
`Embedded` compiler-debug policy. It now reuses `Get-RSBuildEnvironmentOverrides -Suite Full` and
asserts the real targeted receipt's embedded mode and test surface. Its actual x64 ASan Debug rebuild
passed with no skips in `20261023T160028Z-500-791bafe185fb412bbccedf6456a96eb0` (one control,
618.19 seconds). The earlier RPC failures remain failed; this passing repair does not prove their
complete cause or turn the original Full run green. Full-solution re-attestation follows the writer.

The affected-test map now routes the vendored client and immutable lock through the complete native
consumer closure, replacing the two retired viewer rules. All 121 affected planning, client,
inventory and tooling controls passed without skips in
`20261023T155854Z-72156-178499b68c4b4150bcd0594607818157`. Four earlier actual-MSBuild direct-client
controls also passed. The 23 changed native files passed the discovered formatter check. Viewer-specific
coverage obligations transferred to RedMonitor's validation contract; this consumer retains client,
launch and Preferences coverage. The old bundled-viewer documentation image is explicitly historical.

Reviewed component receipts, complete failed selftest case records with reasons, aggregate/plan
identities and bounded failure logs are retained in
[the delivery evidence](../../TestRuns/Local-x64/Validation/2026-10-23_RedMonitorExtraction/README.md).
That record owns exact re-attestation/probe/canary outcomes and identities. It is diagnostic/component
evidence, not a successful Full or portable release attestation. The complete immutable raw archives
remain under the marked `Z:/RedSalamander.Perf`; no cache or failed record is promoted by a focused pass.

Remaining work keeps this plan active:

1. Reproduce and classify every failed broad Commands/FileOps/ViewerPE/PluginContract family under
   its existing domain owner, including relevant predecessors. Preserve all failed, skipped and
   unreached cases; repair confirmed defects and rerun the owning broad suites.
2. After the owner ends the x64 Release deferral, qualify the production/test-enabled Release modes
   and clean app-only ZIP/MSI/MSIX packages, symbols and license closure. Never terminate or replace
   the independently launched executable to bypass that decision.
3. Complete installed-helper handoff, Preferences Apply/restart/live behavior, native platform gaps,
   physical IME/AT and paired same-machine Release performance under the canonical shared plan.
4. Keep independent settings and private-support narrowing as RedMonitor's explicit S4 compatibility
   migration. Move the plans to Done only when their remaining implementation and qualification gates close.

### Effective compiler-policy correction

The post-writer r10 solution rebuild failed with C1090 in Terminal/Factory.cpp. Real MSBuild evaluation
then showed `ProgramDatabase` on that existing ASan compile item despite the requested Embedded receipt
policy. The owned r11 retry was interrupted before changing policy. Updating defaults alone was insufficient;
Embedded now explicitly updates existing compile items, and the SearchTextHelpers `/MP` exception selects
only the canonical shared source path. This fixes the verified policy defect; it does not establish the
complete cause of the RPC disconnects.

Read-only controls now inspect every compile item after all imports in Common, Terminal, the main app
and PluginContractTests across six profiles and two policies. The corrected pre-repair matrix
`20261023T163647Z-66052-61a551eff7de418aaa3abdcb3ffa00c6` passed ten controls and failed precisely the
two ASan Embedded cases. After repair, all 12 passed without skips in
`20261023T163947Z-76884-713382343cd24d4fb9282bc171ab0ee1`. Normal Debug Edit and Continue remains
unchanged; Release evaluation never rebuilt or replaced the open executable. Earlier control-authoring
and metadata-condition failures remain archived separately and are not product qualification.
The final clean ASan rebuild, actual compiler-command checks and component results are recorded in the
reviewed delivery evidence; old requested-mode receipts are not proof of effective compiler flags.

## Master integration and repository migration — 2026-10-24

- [x] Owner authorized conflict resolution preserving the MSVC 14.52 and final RedMonitor changes.
- [x] Reconcile Boolean COM-query admission, no-follow owned-link cleanup and retained target-byte checks.
- [x] Consolidate the ARM64 AWS overlay at Build/vcpkg-ports; retain /Ob1 with exact C compiler 19.52.37025.0 and ARM64 guards, negative patch/architecture controls and dependency evidence binding.
- [ ] Run fresh merged-source tooling, native/runtime and platform qualification; retain every failure and distinguish focus-taking, native ARM64, performance and release gates.
- [ ] Push the integrated master after validation.
- [ ] Publish the validated source snapshot on RedSalamanders/RedSalamander main without importing DualTail commit history, preserving existing destination history and tags.
- [ ] Repoint Z:/src/RedSalamander to the destination and verify remote/local source-tree equality.

The owner authorized focus-taking harness suites for this validation today. Existing failed Full receipts remain failed; no focused or old-branch pass qualifies the merged source. The independent RedMonitor consumer plan stays WIP until its remaining acceptance gates pass.

### Merged-source qualification repairs

The first Fresh x64 Debug Full at source 3946f4c5799de06022f3c819fc1a034cf0539dbe
failed six cases across four entries. Commands recorded 952 passes, two failures and six skips;
Compare recorded 252 passes, two failures and 21 prerequisite skips. FileOps passed all 378 admitted
F7 local/cross-volume/SMB samples, then stopped at a stale repeated-Move expectation: 13 passes,
one failure and 308 unreached cases. Both interactive viewer groups and all 15 PerformanceTests2
cases passed. The immutable archive remains under the authorized E:/RedSalamander.Perf evidence root.
These are failed Full results, not a repository qualification.

- [x] Publish the exact pinned DxUi 0def3d03 dependency on a separate public branch; verify a clean public restore without changing library main or its checkout.
- [x] Retain the original warmup failure and same-binary traversal control: 43,493 ms failed versus 2,698 ms passed. Preserve the 15-second status deadline and persisted-currentness checks.
- [x] Record isolated Terminal and gallery passes separately from their original broad failures. A leased gallery replay and its immediate predecessor also passed; the complete broad failure cause remains unconfirmed.
- [ ] Qualify the repaired retired-project inventory, whitespace-independent checked-cursor ordering, exact Missing source/summary axes, continuously pumped placement deadline and guarded static utility presentation.
- [x] Measure the new progress-counter fixture against the original producer, then batch only the two progress count metrics without changing callbacks, HRESULTs, throttling or sink guarantees.
- [ ] Repeat the owning broad suites and final native/platform gates on the repaired candidate; keep manual/native ARM64, installed-helper, packaging and release gates separate.

The original Compare metrics file reached 42,285,878,561 bytes, with repeated
search.progress.emit_calls and search.progress.emit_suppressed records on suppressed work.
Preserve that failed archive. Counter regression coverage must retain exact value0 totals,
success/rejection/cancellation/null-callback behavior and bounded record counts. Future focused
runs use the already marked C:/RedSalamander.Perf; no new drive-root output is authorized.

The original-producer regression fixture at b1264efba9bf31998d0384f8e8b5fc6c4c8e8216
failed as intended in 20261024T162715Z-98188-c4ad6348d6384111a7557e3d018d3230.
Its successful query recorded 5,880,051 bytes, 8,197 call rows, 8,188 suppression rows
and nine callback records. Rejection returned E_ABORT with one callback; early cancellation
returned ERROR_CANCELLED with three callbacks and 62 suppressed calls; null callbacks emitted
no progress counters. The fixture is unchanged for the candidate. The original failed Full,
failed regression fixture and smaller gallery passes remain distinct evidence.

The candidate at 20c564af6ee26bcd226758f6cbc183b7d0b04c5d passed the identical fixture
in 20261024T163728Z-24504-051f94a8ab7946608e5eb1d7f417f755. Successful-query bytes
fell from 5,880,051 to 10,750 while retaining all 8,197 calls, 8,188 suppressed calls,
five forced callbacks and nine callback records. Cancellation retained all 65 calls,
62 suppressed calls and three callbacks, including its terminal counter flush; rejection
and null-callback behavior passed. These are Debug component results and diagnostic-volume
evidence, not Full or production/Release performance qualification.
The [paired component archive](../../TestRuns/4cb089111a23/Validation/2026-10-24_163728_RedMonitorMergeMigration/README.md)
retains complete case verdicts, byte-preserved traces and raw-stream digests.

A serial Full at 8073 completed all 18 entries: Compare 255/0/21 prerequisite skips,
FileOps 297/0/25 prerequisite skips with all 56 families and 378 admitted F7 samples,
and Commands 953/1/6 (gallery only). Tooling's 1,005 assertions passed, but its one
mandatory version-database readiness skip failed qualification. The isolated build root
now has the exact independent pinned vcpkg checkout; a fresh run must qualify it.

The bounded real-input predecessor baseline at unchanged 8073 binaries passed the
speed-limit and decision-mirror cases, then reproduced the gallery's owned NULL
foreground event on Create Directory closure (run
20261024T211154Z-97124-8697e79b0f084a91b4af226e53d85275). A tentative two-prompt
static-input mode passed those prompts but failed later at Shortcuts (run
20261024T212222Z-78960-6449a4ae9fa94176813f8b104d42b068); that incomplete mode has
been removed. The next candidate centralizes production foreground requests so the
existing no-activation scope declines them before User32. Logical control focus and
native text-input setup remain intact. Capture and modal-return samples also drain the
existing event barrier to localize deferred notifications; NULL events remain failures.
The precise OS/native-input mechanism is unproven. Focused, broad and final Fresh Full
qualification remain required; neither failed run is promoted.

The 88570 Fresh Debug Full completed all 18 entries: 17 promoted and Commands
failed only at provider acquisition (953/1/6). Tooling passed 1,007/0/0, closing
the mandatory pinned version-database readiness gate. Compare and all 56 FileOps
families passed, including all 378 admitted F7 samples. The foreground correction
passed broad capture through Terminal Find; the provider failure is distinct.

The unchanged-binary reload/gallery replay reproduced 1/1/0 in
20261025T003434Z-63016-84858d3c06f84821b92eee17b4f183ae. A new native regression at
ba205371841d1876c74cbe307fa84cc8402292a2 failed against the unchanged producer in
20261025T004437Z-6572-9c8ee5fe311a43d2b7e795af454c2266: an ordinary quiescent
refresh left FTP loadable=false, deferred=true and pathDeferred=true. The gallery
reported the same state with HRESULT 0x80070490 and the expected exported fixture.
The correction retries the canonical unload sweep once after all manager-owned
instances retire. The native fixture also retains a real external interface to
require genuine deferral, callable ownership, identity preservation, no duplicate
aliases and recovery after release over three cycles. The fixture is unchanged
between the failed producer and candidate. Candidate focused and final Fresh Full
qualification remain open; all failed receipts stay failed. Overnight focus-taking
tests are explicitly authorized with the existing warning, lease and restoration.

The shared-provider correction at 47d22a521 passed the three-cycle ownership
regression, external reload and complete gallery in broad Commands. Its Fresh
Debug Full completed all 18 entries: 17 promoted, with Commands 954/1/6 failing
only `folderView_perf_scroll_render_stress` (one full-client repaint during the
first unchanged Brief horizontal request). Compare 255/0/21, FileOps 297/0/25
with all 56 families and 378 admitted F7 samples, both interactive viewer groups,
all 15 PerformanceTests2 cases, and tooling 29/0/0 plus 1,007/0/0 passed. That Full
remains failed. Unchanged-binary overlay/scroll and five-repeat scroll runs passed
2/0/0 and 5/0/0; they establish neither an original cause nor a repair.

A separate controlled native regression at 986e623325a66135553990d5f415f140bf23fb7f
proved that the existing warm-quiescence helper could return while a paint timer
was queued: four trials dispatched the paint, but trial five returned quiet with
pending=true, paintsAtReturn=0 and renderDelta=0 (run
20261025T040111Z-75708-0c710bf54e184d3385b9aa85b7f7c004). The helper now drains
after its delay and checks queue state before sampling counters. The original
scroll failure's precise cause remains unproven; its zero-paint assertion,
measurement boundaries, 100 ms preparation interval and 5-second deadline remain
unchanged. The [controlled baseline archive](../../TestRuns/4cb089111a23/Validation/2026-10-25_040111_WarmQuiescence/README.md)
retains the failed native result and all five observations. Candidate focused,
six-profile build and final Fresh Full qualification remain required.

At source 56232e438b32953ddf9829529b87e7f3d59f1e5e all six builds passed and Fresh
Debug Full passed all 18 groups (20261025T044405Z-46092-fb22a8af299748e39a5496ee1581aaa0).
Release Full passed 17 groups but failed `FileOps_PerItemCancelUndispatchedRoots`:
16 FileOps cases passed, one failed and 305 downstream cases were blocked, not qualified
prerequisite skips. The failure was a published Copy root marked Canceled before terminal
completion; every never-dispatched root remained correctly untouched. The unchanged Release
binary replay (20261025T084433Z-98192-f47832264975477d8a63100aa214343d) passed 17 repeats
before reproducing the same assumption on Move, with Published/Removed/Canceled truth.
Twelve later repeats were blocked. Both failed receipts remain failed; ASan Full did not start.

The fixture inferred completion from visible bytes plus 300 ms, although publication and
completion are independent and an open parallel prompt can hold a sibling callback. The
revised fixture uses explicit serial completed-prefix and two-worker blocked-root controls
for Copy and Move. It observes terminal/dispatch facts under the existing result mutex and
retains exact storage, untouched-root and summary assertions. No production result is changed
to Completed merely because it published. The [baseline archive](../../TestRuns/4cb089111a23/Validation/2026-10-25_084433_PerItemCancel/README.md)
retains complete broad/focused verdicts and digest-bound ordered failure observations.
Revised-source builds, focused native checks and final Fresh Full qualification remain open.

At source d195fe31b9b1948865fbfa06315d626a25f1a961 all six builds and 90 selected
cancellation repetitions passed; Fresh Debug and Release Full passed all 18 groups.
ASan Full completed 15 passing and three failed groups, with eight failed cases:
three service startup/query cases, four Commands cases and selection-shape concurrency.
The [failed broad/replay archive](../../TestRuns/4cb089111a23/Validation/2026-10-25_140118_ASanFailures/README.md)
retains every affected native verdict. Unchanged default seed reproduced all three
search failures; traversal-only controls passed. The UI replay reproduced cold memory
and breadcrumb failures, and exposed a reused synthetic clipboard sequence on later
repetitions. Focus and close passed isolated replays, which do not clear their broad
failures. Selection shape failed five of ten target repetitions without content loss.

The next diagnostic source retains the original 1,000-command/8 MiB memory guard,
100 ms close limit, actual concurrency equality, visible actions, focus identity and
storage assertions. It records memory/allocator values outside the measured interval,
native transfer phases, post-failure action readiness and teardown component times.
A controlled initializing-progress gate tests the optional status snapshot while its
volume mutex is held. Code inspection found a blocking optional overlay and a stale
pre-inspection focus authority; controlled regression, repairs and final qualification
remain required. No failed receipt is promoted and no traversal override is selected
for the ordinary startup/query qualification cases. Public source migration and the
primary checkout's remote cutover remain pending validation.

The [instrumented ASan baseline](../../TestRuns/4cb089111a23/Validation/2026-10-25_193147_ASanDiagnostics/README.md)
at b60d16c5a1cea3efc39cf04f85c5ac9a95be4fbc built with zero warnings/errors.
Its owned gate reproduced a 5,000,855 us optional status wait. The native case failed;
its wrapper separately rejected an omitted explicit-case-list registration, now
corrected in source. That wrapper failure remains failed and is not qualification.
The predecessor UI cluster reproduced empty default/cancel bindings on A's correct
focus and a card that compacted after its first snapshot. Cold private bytes grew
about 465 MiB while live ASan allocation bytes fell about 0.55 MiB. Selection phase
observations proved four real transfers where partial-byte sampling reported two
and zero. The unchanged close test passed at 84,336 us with its held-worker controls;
the earlier one-off 112,377 us cause remains unproven.

The candidate makes optional status enrichment a nonblocking, all-or-nothing
snapshot and chooses keyboard authority from the validated inspection-return task.
It gives each synthetic clipboard repetition a new sequence without clearing replay
protection, observes/expands terminal compaction within the existing three seconds,
and counts explicit admitted native transfer phases while retaining exact concurrency,
content and counting requirements. Ordinary Debug/Release retention setup is unchanged.
ASan warms the actual allocation path once, restores the original starting theme,
keeps the original 1,000-command/private-8-MiB measurement, and adds live-8-MiB bounds
covering cold, measured and combined bursts plus cold HWND/GDI/USER checks. Cold
private cost remains recorded; no sanitizer option or quarantine purge is used.
This is a measurement-protocol correction, not a product memory/performance gain.
The close/focus test predicates and all existing deadlines remain intact. Revised
focused tests, six native builds, three Fresh Full profiles and ETW/canary checks
remain open; publication and checkout cutover still await validation.
At source 4eba261d9122d16aa86762f48307b905716f2e38 the ASan solution build
passed without warnings/errors. The status gate and both original default-store
cases passed, but the original single-query case still failed. The UI predecessor
cluster passed retention and breadcrumb/close checks but reproduced empty decision
bindings on inspection return. Three breadcrumb repetitions and all five native
selection-shape repetitions passed with the original storage/concurrency guards.
These focused passes do not qualify the repository and do not clear failed Fulls.

A second production wait remains in EnumerateNoWait: even SQLite-authoritative
routing locks the optional in-memory volume held by background seeding. The next
controlled gate covers authoritative and opportunistic no-wait queries with the
same owned warmup gate, exact live result and 500 ms budget. Its unchanged-lock
baseline must be retained before applying the nonblocking production correction.
Inspection restoration now precedes binding and resolves the actual current
hosted tree's task identity; recycle safe-focus initialization cannot override a
successfully restored invoking control. All original native focus/default/cancel,
100 ms close, visible-action and retention requirements remain intact. These
changes remain unqualified until their focused and final Fresh Full runs pass.
The [first repair packet](../../TestRuns/4cb089111a23/Validation/2026-10-25_202056_ASanRepairChecks/README.md)
retains the failed 4eba search/UI checks and successful focused repetitions.
At e7f3b7188d9279e74fc1f9ac8d7f337d332445a3 the ASan build passed without
warnings/errors and the same thirteen-case UI predecessor cluster passed all cases.
The [owned no-wait regression baseline](../../TestRuns/4cb089111a23/Validation/2026-10-25_205025_NoWaitRegression/README.md)
failed both authoritative and opportunistic queries at 5,011,577 and 5,014,767 us
against the unchanged 500 ms gate, with exact live results only after gate timeout.
Its status control passed. Complete native/wrapper verdicts remain failed.

The production correction bypasses optional volume lookup/locking entirely for
SQLite-authoritative fallback. Legacy no-wait memory lookup uses try_lock and
selects live scan on contention, retaining warm-cache queries and reason metrics.
The controlled tests are unchanged from their failing baseline. The authoritative
search/UI contracts now state the lock and restored-focus ordering explicitly.
Positive focused validation, six builds, all three Fresh Full runs, ETW/ASan canary,
public snapshot/CI and primary checkout cutover still remain required.
At cfe43d8ddc92973fd47955876cb80ea0b061509e all six builds and 37 focused
checks passed without focused skips. ARM64 retained 37 baseline C4324 warnings.
The first Fresh ASan Full failed ViewerPE's private HTML origin and both PDF
navigation journeys. Ten entries promoted, six remained pending and Compare was
interrupted before further browser tests reused the live Release application's
WebView2 profile. This failed/incomplete Full is not qualification. The
[profile-isolation packet](../../TestRuns/4cb089111a23/Validation/2026-10-25_214951_ViewerWebProfileIsolation/README.md)
retains its exact native/evidence/terminal records and unchanged running-state file.

The ordinary browser profile was in use by the separately launched Z:/src Release
instance, which remains open. Three probes of the same executable/plugin digests
passed with isolated authorized sandbox profiles and all original assertions and
12-second deadlines unchanged. They support correcting test isolation, without
promoting the failed Full or claiming every possible browser failure cause.

Each ViewerPE process now sets an owned sandbox profile before creating WebView2,
overriding inherited profile state and failing closed on setup errors. The existing
HTML/PDF security fixture additionally queries the actual environment profile and
requires exact authorized ownership. Production browser configuration is unchanged;
profile readback is an ENABLE_TESTS observation. Revised candidate builds, focused
profile-isolation controls, all three Fresh Full runs, ETW/ASan canary and migration
remain required. The original eight assertions and budgets remain intact.

At 7b3ae28ce1f6cc99bbeb6453354c123fd7b43721 all six builds and three actual
WebView2 profile-ownership controls passed. The next Fresh ASan Full passed the
browser origin/PDF/profile checks but failed ViewerSpace's export lifetime gate.
Its ambient-desktop control passed. The already-failed run was stopped during
Compare to investigate before spending several more hours on the remaining
suites. Ten entries promoted, one failed and the remaining entries are incomplete;
the raw running-state files remain unchanged and are not qualification.
The [export-gate baseline](../../TestRuns/4cb089111a23/Validation/2026-10-25_233252_ViewerSpaceExportGate/README.md)
retains the exact failed native output/receipt and seven unchanged-binary probes:
six passed and one reproduced the identical gate failure. No cause is established.
The diagnostic candidate records admission, event/stage snapshots and existing
preparation metrics without changing the original five-second gate or any close,
retirement, output, focus, browser or original eight regression predicates.
Repair, all three final Fresh Full profiles, ETW/ASan canary and migration remain
required. The independent Release app and primary checkout audit edits are intact.

The first diagnostic build at 3ff97a52c8e171aba121b0598e07fdb1f9268154
failed linking because ViewerPETests does not link the direct Common JSONL sink
configuration API. No native probe ran. The
[failed build record](../../TestRuns/4cb089111a23/Validation/2026-10-26_005055_ExportDiagnostics/README.md)
is retained. The correction uses the existing governed capture environment in
the owned launcher before startup, preserving standalone boundaries and all
native predicates. Product export behavior and final qualification remain open.

At bf2ad18715abc6a406ae8b0ac302703d06b52617, corrected x64 ASan Debug
and Release builds passed. The
[preparation baseline](../../TestRuns/4cb089111a23/Validation/2026-10-26_010343_ExportPreparationBaseline/README.md)
retains 16 traced ASan, 16 untraced ASan and four traced Release passes of the
original analysis/export cases. ASan approaches the unchanged five-second wait;
one passing final message drain completes at 5.013981 seconds. The original
unobserved failure's precise cause remains unproven. Trace data identifies many
small preparation turns: the first held ASan export uses 146 turns and 195570 us
of active preparation across a 4713 ms metric wall span.

The candidate removes per-row display-string/attribute construction when checking
numeric timestamps, shares the exact existing Windows conversion/range policy,
borrows real names only synchronously and moves strict UTF-8 results into owned
frozen rows. Native tests additionally require exact exported valid FILETIME ticks
and explicit missing/invalid nulls; independent Windows conversion and boundary
tests protect the shared formatter. All original row, byte, timing, focus, output
and lifetime predicates remain intact. Candidate paired performance, six builds,
the final three Fresh Full profiles, ETW/ASan canary and migration remain open.

At b2cbf8f3849da29e1f6a35f53a1dc6ec0caa59e0 all six builds and the paired
export checks passed: all 24 independently parsed CSV/post-Close JSON byte pairs
matched the baseline. ASan median preparation CPU and native-gate ratios were
0.43619 and 0.43729; Release preparation CPU ratio was 0.66376. These focused
results do not qualify the repository. The complete
[ASan Full residual-failure packet](../../TestRuns/4cb089111a23/Validation/2026-10-26_022926_ASanResidualFailures/README.md)
records 16 promoted groups and two failed groups with three failing native cases.
Seven original failures passed, including selection-shape coverage. The original
100 ms Close case failed at 133173 us; stale-root service warmup and dense-files
draw observation also failed. An unchanged-binary six-case Commands cluster
passed all 18 repetitions, without clearing those Full failures. Extra SQLite
replays were stopped in preflight before native execution; their partial state
and interruption remain explicit.

The next candidate adds fine renderer/nonclient and NTFS seed diagnostics and
requires the dense-files test to observe the actual current presented generation
within its unchanged ten-second deadline. Every original draw/LOD assertion,
100 ms Close limit and 60-second warmup watchdog remains intact. Controlled
diagnosis, remaining repairs, all final Fresh Full profiles, ETW/ASan canary,
public snapshot/CI and primary checkout cutover remain required.

At fac05d5b70b9a5933b12093c1b2337f76aee8cdc, the full ASan rebuild and
Release build passed without warnings/errors. The strengthened six-case viewer
cluster passed all 18 focused checks. Fine teardown spans locate roughly
31–36 ms in final brush release; no Close repair is claimed. The normal native
MFT local stale-root case passed, but the service counterpart still timed out.
One complete 4,169,294-entry seed spent 89.327289 seconds enumerating and
18.169793 seconds retiring its temporary graph. The
[retained stage/baseline packet](../../TestRuns/4cb089111a23/Validation/2026-10-26_073646_ResidualStageDiagnostics/README.md)
binds three same-scenario fresh-process ASan and Release baselines each. ASan
totals were 95–98 seconds with roughly 3.2 GB sampled peak private memory;
Release totals were about 8.8 seconds with 832–848 MB sampled peaks.

The next candidate pools only temporary seed graph storage in a scoped arena,
moves unique candidate names into it, and copies retained tracked-root strings
out before release. Native NTFS routing and all original service watchdogs,
freshness, result, persistence, hardlink and corruption assertions remain intact.
The existing parser is shared across ordinary replay strings and pooled seed
strings, with additional Unicode V2/V3 and bounds checks. Paired timing/memory,
ordinary native service replays, the remaining Close repair, all final Fresh
Full profiles, ETW/ASan canary and public migration remain required.
The selected resource is a thread-confined unsynchronized pool, so graph growth
can reuse released buckets/child buffers. The initial monotonic prototype was
not built or measured and is not performance evidence. Retained paired Release
memory remains an admission condition for the pool candidate.

At 7ce7fcee0da65c4084c6f1fe1deb25f9d9b387d3 both x64 builds and 17 focused
native checks passed, including the stale-root service watchdog and the original
three service regressions. The
[default pool was rejected](../../TestRuns/4cb089111a23/Validation/2026-10-26_082420_NativeSeedPoolRejected/README.md):
ASan median seed time ratio was 0.623264 but sampled private-memory ratio was
1.369510; Release ratios were 0.908935 and 3.375519. This is a failed resource
admission, not qualification. No Full was started and no receipt is promoted.
The next candidate explicitly limits pooled requests to 256 bytes and chunks to
1024 blocks, with larger allocations retaining upstream deallocation. It must
pass the same original baseline comparison, native guards and final qualification;
no rebaseline, timeout increase or sanitizer change is authorized or applied.

At 487a99bc2b37f3f67e16924fffb0ba096447dab5 both x64 builds and all 17
focused native checks passed. The
[bounded pool was rejected](../../TestRuns/4cb089111a23/Validation/2026-10-26_090022_NativeSeedBoundedRejected/README.md)
because Release sampled private-memory median was 1.325656 times the original
baseline. ASan timing/memory ratios were 0.642555/0.728908 and Release timing
was 1.044102. No Full was started and no baseline was replaced.
The next candidate packs small seed nodes/names into scoped 64 KiB chunks,
keeps child buffers ordinary and borrows USN input names only until the graph
copy. Explicit parser reset avoids per-record replacement string construction.
New checks cover failure reset, alignment, chunk rollover, large allocation
release and retained-name independence. Paired admission, Close repair and
all final qualification/migration gates remain open.

At 676391cb3edb65c5b38ee4e3db8c0677678e9558 both x64 builds and all 19
focused native checks passed. The
[scratch candidate passed resource admission](../../TestRuns/4cb089111a23/Validation/2026-10-26_094332_NativeSeedScratchFocused/README.md):
ASan time/private-memory median ratios were 0.278035/0.615274; Release ratios
were 1.029977/1.014563, against the same original baseline and limits.
This does not clear the earlier Full failures. The next candidate explicitly
preserves per-object ASan boundaries and ended-lifetime poisoning inside chunks.
It also retains only the free-threaded final D3D device in the existing closing
worker, with a three-device module cap and synchronous capacity fallback. All
HWND/context/D2D cleanup stays on the UI thread. Tests retain the original
100 ms Close limit, strengthen synchronous destruction, and prove held ownership,
capacity fallback and drain while public viewer references remain owned.
Fresh resource comparison, six final builds, all Fresh Full profiles, ETW/ASan
canary and public migration remain required.

At bdf657a9559ed63579ad48e30f6fb306eb0e621a both x64 builds and all 54
focused checks passed, including five original 100 ms Close repetitions and
35 ViewerSpace checks. Actual trace evidence shows all 30 admitted ASan device
releases off the UI thread. Scratch boundaries and ended allocation lifetimes
remained visible to ASan. The
[resource comparison nevertheless failed](../../TestRuns/4cb089111a23/Validation/2026-10-26_100317_GuardedRendererResourceRejected/README.md):
Release memory ratio was 1.113858 against the unchanged 1.10 limit, while ASan
time/memory ratios were 0.282515/0.675939 and Release time ratio was 1.055365.
No Full was started. The live MFT grew past 4194304 entries; inspection suggests
power-of-two bucket growth contributed, but prior bucket counts were not measured.
The next candidate records actual bucket counts and uses 1.25 temporary-map load
factor. It must pass the same retained baseline comparison. Final builds, all
Fresh Full profiles, diagnostics, public snapshot/CI and cutover remain pending.

At c4c7399b0ccf700d5f761be5245f6a1e3c513fea all six builds, 54 focused
checks, seed resource admission, ProductUi and all 24 export/byte-pair checks
passed. ASan seed time/private-memory ratios were 0.279182/0.681069; Release
ratios were 0.896612/1.021731 against the original retained limits. Actual entry
buckets were 4194304 for 4225414 records. The user then permitted skipping Full
runtime reruns based on this green earlier qualification. No Full was run or
reclassified: the coordinator stopped at read-only planning, with all six profiles
verified uncontaminated after correcting the handoff timestamp conversion.
The [focused scope recheck](../../TestRuns/4cb089111a23/Validation/2026-10-26_113110_FocusViewportRecheck/README.md)
retains one failing original focus case after the theme/breadcrumb prefix, its
passing isolated diagnostic control, and the passing selection-shape case with
setup/cleanup. The disclosure lookup assumed B was visible; the candidate uses
the existing exact-task bounded-scroll helper and records pre-reveal visibility
and HRESULTs. No original focus, task identity, decision or timing predicate is
weakened. Corrected-source builds and the failing-order replay remain required;
Full runtime reruns remain omitted under the user's updated scope.

At a4209ee8cd7aa8f548f18706ca4a08c3bd272479 all six builds passed, but the
[focused replay failed one of twelve checks](../../TestRuns/4cb089111a23/Validation/2026-10-26_131427_FocusBindingsFailure/README.md).
All disclosure routes resolved correctly. Returning to A after B completed had
A's exact focused inspection control and current context, with empty default and
cancel bindings. The unchanged second iteration passed and does not clear that
failure. Focus callbacks updated visibility while binding updates waited for a
later frame. The correction reuses the last published task/generation projection
for synchronous focus rebinding, suppressing intermediate reconciliation callbacks.
An additional immediate A/B UIA focus assertion pumps no frames. All original
return, identity and timing predicates remain. Current-source focused qualification
and diagnostics must pass before migration; Full reruns remain omitted by the user.

At 7bb6990285ee0027eddde7bb8e68cbe2732c61b4 all six build profiles and
all eight original ASan regressions passed (12 Commands, three search and three
FileOps setup/selection/cleanup checks, no failures/skips). The following
[Debug pointer route failed](../../TestRuns/4cb089111a23/Validation/2026-10-26_140706_DisclosurePointerFailure/README.md):
B's retained disclosure provider existed and had focus while its clipped UIA
rectangle was zero-sized. Provider existence had been mistaken for visibility;
SetFocus on the already-focused control need not scroll again. The fixture's
existing bounded scrolling lookup now requires a measured disclosure rectangle.
Other action lookup semantics, all original pointer/focus/identity predicates,
the eight-page bound and timing limits remain intact. Production code is
unchanged after the green original ASan closure. Fixed-source qualification is
still required before publication; Full runtime reruns remain omitted by the user.


Final focused source handoff: all six builds, all eight original ASan regressions, 51 native checks, three real ETW roundtrips and the real ASan negative control passed. The [qualified source snapshot packet](../../TestRuns/4cb089111a23/Validation/2026-10-26_145454_MigrationFocusedClosure/README.md) preserves exact source/artifact provenance, no-frame A/B focus and original return assertions, the explicit Full runtime waiver, and fresh public-only source/history/dependency verification. Earlier failed receipts remain failed. This public-history child changes documentation/evidence only; public PR/main CI and primary-checkout routing are subsequent delivery operations. Installed-helper, packages, release and manual platform/accessibility gates remain open; I47 stays WIP.
