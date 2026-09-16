# DxUi deferred ARM64 and ASan qualification

Status: **HOLD**. Owner: the user, who deferred ARM64 and ASan on 2026-09-13
and will validate them later. This is the remaining platform qualification from
[I19](../Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md), not new implementation scope.

## Checklist

- [ ] Qualify the final RedSalamander revision in x64 ASan Debug, including real
  seeded sanitizer detection, affected regressions and Fresh Full.
- [ ] Build the final RedSalamander revision in ARM64 Debug, Release and ASan Debug;
  verify the actual consumer module hashes, exact DxUi pin and compiler host.
- [ ] Execute native Windows ARM64 tests for DxUi, RedXe and RedSalamander in all
  three configurations, including real sanitizer detection in ASan Debug.
- [ ] Review each product's final tested source, skips and hardware restrictions
  before claiming platform runtime qualification or releasing that configuration.

## Evidence available for resumption

RedSalamander `d404fc90` x64 ASan Debug builds with zero warnings/errors in 4m21s.
The seeded heap overflow is detected. `content_pending_elided` and
`cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs` each pass
ten repeats; all 14 consumer module identities were checked after each run.
The earlier borrowed-name heap-use-after-free remains recorded with its original
dump and matching executable/PDB identities.

The next focused bandwidth run has 29 passed, one failed and zero skipped
(including setup/cleanup cases). Repeat 10 reports cancellation latency 765,000 us
against 250,000 us. It does not report the previous shortened-duration failure.
The driver stops before queue-after, don't-start, capability, Curl and Fresh Full.
Those unexecuted steps are not passes. No further ASan work is authorized by this
closeout pass; retain the failure for the user's later investigation.

Original records are under
`Z:/RedSalamander.Perf/evidence/I19-20260909/d404fc90-isolated-x64-ASan-Debug-*`.
[Reviewed original receipts](../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-ASanDeferredD404fc90/README.md)
retain the build, probe, case results, module checks and stopped driver.
Earlier successful ARM64 cross-builds and failed ASan Full runs retain their exact
source revisions in I19. Cross-builds do not prove native runtime behavior.

The supported build matrix remains Debug, Release and ASan Debug on x64/ARM64.
This explicit user deferral changes the current closeout scope, not the build
contract or the meaning of a passed test. No hosted CI is used.
