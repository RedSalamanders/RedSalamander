# File Operations: Beeline residuals and provider transfer routes

> **NON-NORMATIVE COMPLETED EXECUTION RECORD.** Closed on 2026-09-08. Current behavior is defined by
> `Specs/FileSystem/FileSystem_FileOperations.md`,
> `Specs/Plugins/Plugins_VirtualFileSystem.md`, and
> `Specs/Testing/Testing_PerformanceValidation.md`. This file routes the work that
> Operation Beeline deliberately placed outside its boundary together with the
> residuals its closing Fresh Full gate recorded. All eight outcomes below are
> closed; durable behavior is merged into the authoritative specs. The six future
> features remain separately deferred under H3.

## Status

- **State:** DONE (2026-09-08; Fresh Full passed all 35 entries).
- **Priority:** P1 cancellation/P2 follow-up, opened at the Beeline closeout.
  BR-1/BR-2 provider-description corrections are complete; their future routes
  remain deferred by the owner to H3. BR-3 through BR-8 are complete, including
  the actual NTFS placeholder copy repair, host FTP directory-Move coverage,
  popup teardown repair and Preferences Keyboard search-input correction.
- **Planned at:** `59d6b1038e20e455326b2e302231d4e79068568c`
- **Ownership boundary:** the Microsoft Drive and MTP provider-description
  corrections below, and the named validation debt left by Operation Beeline. Not
  owned: same-endpoint Move behavior, the rename merge, the Move/Copy link
  policy, and the removed `links.nativeMoveSemanticTransform` capability, all of
  which are delivered and now owned by the authoritative specs
  (`Operation_Beeline_SameEndpointMoveIsRename_2026-09-07.md` is frozen
  history). Also not owned: the six deferred features held under `H3`, including
  the future Microsoft Drive Copy and MTP native Move implementations, the
  broad test-contract migration owned by `I4`, and the File Operations stress
  evidence rows owned by `I3`.
- **Drift check:**
  `git diff 59d6b1038e20e455326b2e302231d4e79068568c..HEAD -- Plugins/FileSystemMicrosoftDrive Plugins/FileSystemMtp Plugins/FileSystemCurl RedSalamander/SelfTest/FileOperations RedSalamander/FolderWindow.FileOperations.State.cpp Plugins/FileSystem/FileSystem.FileOps.cpp Specs/FileSystem/FileSystem_FileOperations.md Specs/Plugins/Plugins_VirtualFileSystem.md Specs/Testing/Testing_PerformanceValidation.md`
- **Authoritative specs:** `Specs/FileSystem/FileSystem_FileOperations.md`,
  `Specs/Plugins/Plugins_VirtualFileSystem.md`,
  `Specs/Testing/Testing_PerformanceValidation.md`.

## Why

Operation Beeline made a same-endpoint Move one native rename and proved it with
a Fresh Full gate. It routed two provider transfer questions out of its boundary
because they are Copy-side or device-side work, and its gate recorded residuals
that are not Move defects but must not be forgotten. This plan holds exactly
those items so that no one reopens the frozen Beeline record to find them.

## Outcomes

| ID | Outcome | Kind |
|---|---|---|
| `BR-1` | Complete: normative provider row states same-endpoint Copy unsupported; future route deferred to H3 | Product truth |
| `BR-2` | Complete: `nativeMove=false` retained and host Copy-only/source kept documented; future route deferred to H3 | Product truth |
| `BR-3` | Complete: cancellation and Phase 7 cases pass in Fresh Full, followed by tasks-quiet and popup-host shutdown | Reliability |
| `BR-4` | Complete: real NTFS Cloud placeholder copies through a worker under `Skip links`; junction is skipped | Reliability / coverage |
| `BR-5` | Complete: host-driven FTP directory Move and refusal prove one rename pair, exact tree state and no fallback | Coverage |
| `BR-6` | Complete: configured Z:/D: test roots authorized and independent-volume control passed | Environment |
| `BR-7` | Complete: post-destruction focus restoration passes controlled teardown and canonical Commands | Reliability |
| `BR-8` | Complete: exact native input-host wait passes controlled, predecessor-context and Fresh Full validation | Reliability |

## BR-1 — Microsoft Drive same-drive Copy

Audit correction 2026-09-08: `CopyItem` returns `ERROR_NOT_SUPPORTED`, the typed
descriptor does not advertise same-endpoint Copy, and host admission rejects it.
Only cross-endpoint bridge import/export Copy is available. The authoritative
provider row now states that limitation; it no longer claims server-side Copy.
This closes BR-1's specification-truth alternative. A future same-drive Graph
Copy implementation remains a separate product enhancement, not an optimization
of an existing provider Copy route.

The owner's 2026-09-08 decision places the implementation and its monitor/fault/
performance tests in [H3](../WIP/Operation_FileOperations_LaterFeatures_2026-09-05.md#microsoft-drive-same-endpoint-copy).
This plan has no remaining provider-Copy implementation queue.

## BR-2 — MTP `nativeMove`

Audit correction 2026-09-08: MTP advertises `nativeMove=false` and has no
bound/conditional source-deletion proof. Host strategy selection therefore
produces Copy-only/source kept, including on one device; it does not copy then
delete. The authoritative provider row now states that limit. BR-2's guarded
capability alternative retains `nativeMove=false`: the legacy provider Move
entry point may select WPD move or stream-copy/delete and is not isolated native
authority. A future native WPD route must be proved independently, including
success, cancel, non-commit, and uncertain outcomes, before advertising it.
The owner's 2026-09-08 decision places that route and its fake/physical-device
qualification in [H3](../WIP/Operation_FileOperations_LaterFeatures_2026-09-05.md#qualified-mtp-native-move).

## BR-3 — Parallel-Copy reliability under the Full gate

Execution started 2026-09-08 at `c36d3497` for the owner's P1/P2 audit fixes
and BR-3. The bounded scope includes cancellation while queue-paused and
race-free test observations; provider-description corrections remain BR-1/BR-2.

- [x] Prove `BR3_CancelQueuePausedTransfer` fails before the cancellation fix,
  with both real Local transfers parked and the predecessor held paused.
  Debug run `20260908T064049Z-59444-dbd85e683bd044da8b0a2f23f3933019`
  on the instrumented baseline: expected failure at the three-second bound;
  both tasks drain after failure-only queue release. Build: zero warnings/errors.
- [x] Make task cancellation/stop leave pause/conflict waits and validate bounded
  completion without changing Queue mode or releasing the predecessor.
- [x] Read in-flight observations under their owning mutexes; retain the
  existing concurrency assertions.
- [x] Prove and fix conflict-wait cancellation's predicate/notification race;
  test the peer held immediately before condition-variable sleep.
- [x] Qualify the Phase 7 knob fixture with its limit present at admission and
  retained until concurrency/per-call sharing are observed; reject fast completion
  without those observations and keep retries bounded.
- [x] Identify the cancellation and fixture defects and validate the original
  Phase 7 sequence in Fresh Full. No unexpected timeout recurred; the historical
  timeout's exact worker stack remains unavailable and is not inferred.
- [x] Archive focused Debug/test-enabled Release evidence and a Fresh Full gate
  covering Phase 7, tasks-quiet, and popup-host shutdown; update domain specs.

The cancellation witness records `FileOps.SelfTest.QueuePausedCancelLatency`
in microseconds, the pass bit in `value0`, and the 3,000,000-us bound in `value1`.
Its RED cleanup explicitly releases Queue mode before failing, so the diagnostic
run does not retain a spinning worker. A single green isolated family does not
close BR-3.

Focused Debug candidate `20260908T064606Z-66052-62ad35df8dd940b5baf986310f2339ed`
passed all three selected/setup/cleanup cases, zero build warnings/errors.
Cancellation completed in **62,000 us**, below the 3,000,000-us bound, while the
predecessor remained paused. Raw results, trace and metrics are preserved in
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-08_064606_br3_queue_cancel_debug/`.

The initial whole-process diagnostic
`20260908T065123Z-75136-609108bc3bfe424789f7e00f0e1a6cd3` observed both original
Phase 7 cases and shutdown passing (177 passed, one alternate-volume authorization
failure, 20 skips). Its raw run was removed by the runner's stale-run cleanup before
promotion, and source changed during that diagnostic. This observation is not
archived qualification of the subsequent conflict-wake and knob-test changes.

The forced conflict-wake RED control failed all five test-enabled Release repetitions
in `20260908T071617Z-66932-d222086f33354384814a9bc45acf7f62`: cancellation returned
before the held peer could sleep, and all five `FileOps.SelfTest.ConflictCancelWake`
rows recorded `value0=0`, `value1=1`. Results, trace and metrics are preserved in
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-08_071617_br3_conflict_wake_red_release/`.

The same five-repeat Release control passed after synchronizing stop/cancel
notifications with both wait mutexes: 15 setup/case/cleanup passes, zero failures,
zero build warnings/errors, and five `ConflictCancelWake` rows with `value0=1`,
`value1=0`. Archive:
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-08_072857_br3_conflict_wake_green_release/`.

The production defects are (1) re-entering an already-awake pause/conflict wait
after cancellation while queue/peer state remains active, and (2) publishing an
atomic stop/cancel flag without synchronizing the notification with the conflict
or external-stop pause predicate. The forced controls distinguish the spin and
lost-wake paths. The historical Full timeout has no retained worker stack, so
these proofs do not claim its exact instruction pointer. The knob fixture also
had a post-admission limit race, an unconditional one-second release, and a retry
counter clamp that could prevent exhaustion. Its fix keeps the limit through
actual concurrency/sharing observation and rejects unobserved completion.

**BR-3 closeout:** Fresh Full invocation
`20260908T073251Z-67752-50d9f956c380407d9fb1ef91d0b66c49` passed FileOperations
with **178 passed, zero failed, 20 prerequisite/ownership skips** in one complete
process sequence. Queue-paused cancellation took **78,000 us** with the predecessor
still paused. `Phase7_ParallelCopyMoveKnobs` passed in 38,594 ms and observed both
parallel entries and per-call sharing at concurrency 4 and 8;
`Phase7_SharedPerItemScheduler` passed in 1,047 ms. The former cascade cases
`C1_ExitCloseDeferredUntilTasksQuiet` and `Phase14_PopupHostLifetimeGuard` passed
in 875 ms and 391 ms. The forced conflict-wake control also passed in this Debug
process. Builds and re-attestation reported zero warnings/errors.

Archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-09-08_073251_br3_fresh_full/`.
It preserves aggregate v2, all case statuses/reasons, build/plan/decision and
promotion identities, ordered bounded metrics and the raw-stream digest/counts.
The Full repository verdict is **FAILED**, not green: 2,108 passes, one Preferences
Keyboard search-rebuild failure and 52 explicit skips across 35 executed entries.
That unchanged Preferences case passed all three isolated repetitions using the
same build receipt, preserved in
`Specs/TestRuns/4cb089111a23/Commands/2026-09-08_084447_br3_full_preferences_followup/`.
No broad shuffle triage was performed, so the original failure is not cleared or
reclassified as flaky. This separate repository-gate follow-up is listed in H3;
it does not reopen the now-passing BR-3 File Operations sequence. BR-8's
final closeout below supersedes that failed gate with a green Fresh Full; the
original three isolated passes remain diagnostic history only.

Two consecutive Fresh Full gates on 2026-09-07 each failed exactly one Phase 7
parallel-Copy case while the same family passes in isolation.

- Gate `20260907T191824Z-55452`: `Phase7_SharedPerItemScheduler` timed out after
  240 s. Copy task B held `inFlight=1` for the whole budget after cancellation,
  and the stuck task then cascaded into `C1_ExitCloseDeferredUntilTasksQuiet` and
  `Phase14_PopupHostLifetimeGuard`, both of which wait for tasks-quiet or
  shutdown.
- Gate `20260907T205843Z-46028`: `Phase7_ParallelCopyMoveKnobs` never observed
  more than one in-flight entry with `copyMoveMaxConcurrency=8`, although the
  scheduler reported `maxConcurrency=8 workerCount=16` and the interlock showed
  no contention.

Both historical cases used the Local provider's parallel `CopyItem` path with a
bandwidth limit, as family 20 of 30 inside one long-lived process. The closeout
above preserves the whole-process regression cases and the claim-bearing
`Phase7_*`, scheduler, throttle and interlock evidence. A stuck in-flight item
that blocks shutdown was the reliability failure; the timeout was its symptom.

## BR-4 — Placeholder proof for Copy `Skip links`

The original `Beeline_CopySkipLinksKeepsPlaceholders` fixture could not prove
the placeholder half of Copy `Skip links` on the ReFS sandbox. Selecting D: NTFS
also exposed that WOF compression did not present the required reparse attribute.
The authorized scope allowed an equivalent non-name-surrogate placeholder.

The 2026-09-08 closeout replaced the hidden WOF fixture with a fully hydrated
Cloud Files placeholder on D:. Windows gave this app process placeholder-disguise
mode 1; an explicit fixture-lifetime exposure is needed to prove actual reparse
classification. Both process/thread modes and the temporary sync root are WIL-owned.
The resulting RED control uncovered a product defect: the parallel producer releases
its primary buffers, then `CopyLink` calls synchronous `CopyFile` for a regular-file
placeholder. No bytes reach the destination. Archive:
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-08_102141_placeholder_parallel_red/`.
The fix sends that classification into the normal bounded worker queue; the test
requires one worker admission, no hydration prompt for this fully hydrated file,
and complete payload bytes.

## BR-5 — End-to-end Curl directory Move

Complete: `BR5_CurlHostDirectoryMove` and `BR5_CurlHostDirectoryMoveRefused`
drive the real host engine against the owned fake FTP server. Success relocates
nested files and an empty directory with one RNFR/RNTO pair, preserving siblings.
The refused RNTO also occurs once, keeps the complete source tree and performs
no RETR/STOR/DELE/RMD/MKD fallback. Without an authoritative non-commit receipt,
the host conservatively reports Indeterminate and offers no automatic retry;
the fixture's source observation does not grant stronger host authority.
Both cases pass in focused coverage and canonical FileOperations in Fresh Full.

## BR-6 — Alternate-volume discovery control

Complete 2026-09-08: the machine resource manifest already enables exact alternate
root `D:\RedSalamander.Perf`. Fresh Full above used the governed runner's explicit
`-AllowAlternateVolumeTestRoot` flag with primary `Z:\RedSalamander.Perf` and passed
`R4A19_DiscoveryIndependentVolumes` in 3,750 ms. Its `IndependentVolumes` metric
records overlapping completion in 1,234,000 us and `value1=1`. The earlier
`alternate-volume-test-root-not-authorized` result was an invocation prerequisite,
not a product regression. Retain the flag for this configured machine's Full runs;
it does not authorize arbitrary roots on other machines.

## BR-7 — Navigation-view popup Commands case

The historical Full failure left native focus null after Escape. A controlled
reentrant-destruction witness proves that a restore posted inside WM_NCDESTROY
can run before native focus clearing finishes. The owning popup now restores
through its existing foreground-safe route only after destruction returns,
then queues the same guarded restore after the input callback unwinds.
The original assertion and three-second bound are retained. The controlled case
includes the historical focus-follows-pointer predecessor. Original popup,
owned-window activation and the controlled case all pass in canonical Commands.
This closes BR-7; I4 still owns its separate broad test-contract migration.

## Completed closeout pass — BR-4 / BR-5 / BR-7 / BR-8

Owner authorized all four items and exact local roots on 2026-09-08 at `892f9235`.
Local configuration is `Z:\RedSalamander.Perf\config\machine-resources.json`, with
volume facts in adjacent `local-volumes.md`: primary Z: is ReFS, alternate D: is
NTFS, both fixed and marker-owned. Retain `-AllowAlternateVolumeTestRoot`.

- [x] Recover original BR-7 Commands result from immutable entry evidence:
  popup closed, but native focus was null after Escape. Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-09-07_205843_br7_original_failure/`.
- [x] Run unchanged Preferences predecessor family: 174 passes, zero failures,
  one explicit foreground-ownership skip. This does not clear the failed Full.
  Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-09-08_093136_preferences_baseline/`.
- [x] Prove both controlled failures before fixes: the Keyboard wait accepts a
  different live HWND; a popup restore delivered during teardown is consumed
  before native focus clearing finishes. Two expected failures, zero skips:
  `Specs/TestRuns/4cb089111a23/Commands/2026-09-08_094051_ui_focus_red/`.
- [x] Implement exact native input-target waits, post-destruction popup focus
  restoration, NTFS alternate placeholder fixture, and host-driven Curl rename/refusal cases.
- [x] Validate repaired focus witnesses repeatedly and in their predecessor contexts.
- [x] Prove placeholder Copy preserves all 65,536 bytes and skips only the junction on D:.
- [x] Prove one RNFR/RNTO pair through host admission for success and refusal;
  preserve sibling/empty/nested shapes and reject any relay/delete fallback.
Focused repaired controls passed five repetitions each (10/0/0):
`Specs/TestRuns/4cb089111a23/Commands/2026-09-08_095421_ui_focus_green/`.
Keyboard input reacquisition took 94–141 ms; popup focus return took 16–31 ms.
Both host FTP cases passed (4 setup/case/cleanup passes):
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-08_095645_br5_host_ftp_green/`.
The initial D: WOF fixture correctly skipped because the reparse attribute was absent:
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-08_095536_wof_fixture_diagnostic/`.
A local probe confirmed WOF compression without that attribute; the replacement
uses an actual fully hydrated Cloud Files placeholder, whose observed tag is
`0x9000601A` and whose complete payload is available without any network provider.

- [x] Run canonical Commands and final Fresh Full; archive exact results and metrics.
- [x] Merge lasting contracts and move this completed plan to Done.

Focused BR-4 proof: all three repetitions passed, 9/0/0 including setup/cleanup,
with one worker admission and the complete 65,536-byte payload per repetition.
Archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-09-08_103402_placeholder_parallel_green/`.
Keyboard predecessor-context proof: 24 cases, three repetitions, shuffle seed 908;
72/0/0, including the original failing case in every repetition.
Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-09-08_103757_keyboard_context_green/`.
The BR-7 controlled witness already includes its historical focus-follows-pointer
predecessor in each of five repaired repetitions. The original popup variants and
both new controls also passed in canonical Commands inside the final Fresh Full.

## BR-8 — Preferences Keyboard search input identity

The former wait accepted any live `GetFocus()` HWND while DxUi reported logical
search focus. A controlled displacement onto the category host proves it returned
the wrong native input target. Require the active page's native text-input HWND
and repeated stable samples, reasserting search focus when it differs. Preserve
the real edit-message/search-rebuild assertions and report final focus/search/row
state on failure. This bounded repair coordinates with I4's shared-boundary policy;
it does not activate the general test-contract migration.

## Final Fresh Full closeout

Run `20260908T104319Z-97448-ff25c4b1b0d545fe9b3654fc1edfb85f` returned
repository/change **PASSED**: 2,113 passes, zero failures, 52 declared skips.
All 35 entries executed and were promoted, with no reused, provisional,
invalidated or incomplete entry. Commands has a fresh broad seal.

- Commands: 896/0/2. Original Keyboard search round trip passed in 686 ms;
  displaced-input control reacquired the exact HWND in 125 ms. Original popup,
  owned-window activation and the reentrant-close control passed; the latter
  restored focus in 15 ms.
- FileOperations: 180/0/20. BR-4 copied all 65,536 bytes through one worker on
  D: NTFS in 281 ms, with the actual Cloud tag and no hydration prompt. BR-5's
  success/refusal each issued one rename pair with no relay/delete fallback.
  BR-3 queue cancellation, both Phase 7 cases, tasks-quiet, popup-host shutdown
  and the independent Z:/D: volume control also passed.
- CompareDirectories: 242/0/30. All skip reasons remain explicit in the archive.
- Both builds reported zero warnings/errors. The accepted receipt is
  `4cefc3636a4084de4af21b81abab35941efa94ebfcf7ab9daf1d19f11ca32d45`;
  workspace snapshot is
  `586e01641d3b534ad321b5b3a4be717601c6398c4e08fde66d3fd85223a24e10`.

[Fresh Full evidence](../../TestRuns/4cb089111a23/FileOps/2026-09-08_104319_br4_br5_br7_keyboard_fresh_full/README.md)
retains complete case results, aggregate/provenance records, and digest-bound
ordered correctness/latency metrics. Source stayed fixed throughout the gate.
The required post-gate Done-path admission and archive metadata passed 75 focused
specification/tooling checks; both inventories report zero blocking findings.
The local volume configuration remains in `Z:\RedSalamander.Perf\config\`.
H3's six deferred features, I3's five stress rows, I12's bounded lifecycle work,
and the separately routed human/live-environment qualification remain open.

## Implementation and tests

- Provider-description corrections only for BR-1/BR-2; future implementation and
  fake-backend qualification are owned by H3 after promotion.
- `Plugins/FileSystem/FileSystem.FileOps.cpp` and the per-item scheduler in
  `RedSalamander/FolderWindow.FileOperations.State.cpp` (`BR-3`).
- `RedSalamander/SelfTest/FileOperations/*` for `BR-3` through `BR-5`; every new
  case joins `kFileOpsFamilyDefinitions` in
  `FolderWindow.FileOperations.SelfTest.cpp` or the Full suite skips it silently.
- `Specs/FileSystem/FileSystem_FileOperations.md` provider baseline row for
  `BR-1`; `Specs/Plugins/Plugins_VirtualFileSystem.md` capability rows for
  `BR-1` and `BR-2`; `Specs/Testing/Testing_PerformanceValidation.md` evidence
  rows for any new claim.

## Verification

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -TestRoot Z:\RedSalamander.Perf
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -AllowAlternateVolumeTestRoot -TestRoot Z:\RedSalamander.Perf
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

The Full suite serializes File Operations runs across worktrees through the
session-global mutex; a concurrent run exits with code 3 and must be re-run, not
reported as a failure.

## Done criteria

- `BR-1`: a same-drive Microsoft Drive Copy either performs the Graph action
  with no host payload bytes, or the provider baseline row no longer claims a
  server-side copy. The mutation-result contract is unchanged either way.
- `BR-2`: MTP `nativeMove` is either true with an isolated WPD move route proved
  against the fake backend on the success, cancel, and non-commit paths, or it
  remains false with the reason recorded next to the capability.
- `BR-3`: the root cause is named, and a Fresh Full gate runs the Phase 7
  parallel-Copy cases green without the cascade into tasks-quiet and shutdown.
- `BR-4`: the placeholder claim is evidenced on the gate machine, or the case
  states in its skip reason why the machine cannot carry it.
- `BR-5`: a host-driven Curl directory Move case exists and covers both the
  server rename and the refused `RNTO`.
- `BR-6`: the alternate-volume control is authorized and green, or waived in
  writing in `Specs/TestRuns/README.md`.
- `BR-7`: the case is classified; repair the owning popup teardown for a product
  focus-return defect or the shared wait for a test defect, preserving the assertion.
- `BR-8`: the controlled wrong-input-target case and original Keyboard failure
  pass with the corrected shared wait, including canonical Commands in Fresh Full.
- Performance evidence for any accepted claim is archived under
  `Specs/TestRuns/<MachineHash>/FileOps/` and validated with
  `Tools\Test-TestRunArchive.ps1`.

## STOP conditions

- `BR-1` turns out to require a Graph permission the product does not already
  request. Stop and route the permission question to the owner before writing
  code.
- `BR-2` cannot prove a WPD move that leaves the object intact on cancel. Stop
  and keep `nativeMove` false; a partially moved device object is worse than a
  slow Move.
- `BR-3` reproduces only under the gate and never in isolation after a bounded
  effort. Stop, record the exact evidence, and route it as a scheduler design
  question rather than continuing to retry the case.
- Any item here starts to change same-endpoint Move behavior. Stop: that
  behavior is delivered and owned by the authoritative specs, and a change needs
  its own plan and owner decision.

## Out of scope

- The Shell Copy/Move experiment stays on HOLD under `H3`
  (`../WIP/Operation_FileOperations_LaterFeatures_2026-09-05.md`). It covers single
  files only and is not a dependency of any item here.
- The accepted consequences of the rename route are settled, not tracked: an
  absolute link inside a renamed tree keeps naming its old target, a renamed
  tree keeps its source security descriptor without recomputing inherited ACEs,
  and a tree containing an open file can fail with a sharing violation instead
  of falling back to a copy. Explorer behaves the same way in each case.

## History

- 2026-09-08: completed BR-4/BR-5/BR-7 and the added BR-8 Keyboard failure;
  final Fresh Full passed all 35 entries. Merged durable contracts, archived
  controlled RED/repeated GREEN and full-process evidence, and moved I16 to Done.

- 2026-09-08: completed the P1/P2 cancellation and test-race fixes and BR-3's
  Full-process Phase 7/shutdown proof; BR-6's configured alternate-volume control
  also passed. Preserved the separate Preferences Keyboard Full failure and
  isolated follow-up in H3. BR-4/BR-5/BR-7 keep this plan ACTIVE.
- 2026-09-08: opened at the Operation Beeline closeout to hold that plan's
  routed-out provider work (`BR-1`, `BR-2`) and the residuals its two Fresh Full
  gates recorded (`BR-3` through `BR-7`), so the frozen Beeline record is never
  reopened to find them.
