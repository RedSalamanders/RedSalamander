# Local adoption qualification, 12-13 September 2026

These receipts preserve failures as well as their focused follow-up. No hosted CI
was launched. `retained-files.json` hashes each copied original. Full diagnostic
archives remain under `Z:\RedSalamander.Perf\evidence\I19-20260909`.

## Final local qualification

Native revision `1232d181` passes x64 [Debug Fresh Full](../2026-09-13-Debug1232d181Isolated/README.md)
(2,122 passed, 0 failed, 52 skipped) and [Release Fresh Full](../2026-09-13-Release1232d181Isolated/README.md)
(2,118 passed, 0 failed, 56 skipped). Each promotes 20/20 entries and verifies all 14 final module identities.
The corrected overlap fixtures pass ten Debug repetitions at `f06f779f` and
ten final Release repetitions at `1232d181`; each final Full also executes them. The final
[portable A/B/A rehearsal](../2026-09-13-Packaging1232d181/README.md) passes, including
all 11 packaged identities and identical baseline payloads before/after.

The user accepted the measured resource trade-off and deferred further ARM64/ASan
qualification to [H4](../../../../Plans/WIP/DxUi_DeferredPlatformQualification_2026-09-13.md) on 2026-09-13.
Original failures and skip reasons remain retained. Publishing the tested library pin,
verifying public restore and consumer master merges remain separate release actions.
The [completed I19 record](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md) owns the final scope.
The [final documentation checks](../2026-09-13-FinalDocumentation/README.md) retain
the focused tests, inventories, archive and whitespace validation used for closeout.

Retained navigation follow-up:

- Failed Full runs at [f06f779f](../2026-09-13-Debugf06f779fIsolated/README.md),
  [a8350f4b](../2026-09-13-Debuga8350f4bIsolated/README.md) and
  [2f4fb03f](../2026-09-13-Debug2f4fb03fIsolated/README.md).
- [Refresh context](../2026-09-13-FocusedA8350f4bDebug/README.md),
  [unchanged history context](../2026-09-13-HistoryBeforeA8350f4b/README.md), and
  [history diagnostics context](../2026-09-13-HistoryDiagnostics2f4fb03f/README.md).
- [Source-guard correction](../2026-09-13-HistorySourceContract/README.md) and
  [diagnostic syntax correction](../2026-09-13-AncestorSyntax1232d181/README.md).
- [Final ancestor/history context](../2026-09-13-AncestorContext1232d181/README.md)
  and [final Release overlap/navigation repeats](../2026-09-13-Focused1232d181Release/README.md).

## Earlier qualification at a20f6680

Frozen `a20f6680` includes the corrected test-enabled Release ViewerText diagnostics.
The [three ARM64 builds](../2026-09-13-ARM64a20f6680/README.md) pass, with all 42 module
identities verified. The [portable package and A/B/A rollback](../2026-09-13-Packaginga20f6680/README.md)
also pass. The [first a20f6680 Release Full](../2026-09-13-Releasea20f6680/README.md)
passes every product suite and 19/20 entries; eight tooling checks fail because the
combined local driver loads modules from both checkouts. The
[isolated probe](../2026-09-13-LocalDriverIsolation/README.md) reproduces the cause.
The [isolated Release Full](../2026-09-13-Releasea20f6680Isolated/README.md) passes
20/20 entries: **2,113 passed, zero failed, 56 skipped**, with all 14 post-Full module
identities verified. ASan Debug and Debug qualification follow in separate clean
processes; native ARM64 runtime,
resource acceptance and publication remain open in the
[I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).

The [failed 3a43186a Release Full](../2026-09-13-Release3a43186a/README.md) remains
retained, alongside the [unchanged five-case before/after verification](../2026-09-13-ViewerTextReleaseDiagnostics/README.md).
The dated results below preserve earlier qualification and do not qualify later corrections.

## Functional checks

The [later Full Debug follow-up](f12acdba-Full-followup/README.md) at `f12acdba`
passes 19/20 entries: 2,115 cases pass, one fails and 52 skip. Menu hover and
ViewerImgRaw pass in Full. The remaining ViewerWeb close-time failure is retained;
an exact-duration diagnostic and subsequent isolated/group passes do not explain
its cause or replace final Full acceptance.

Fresh Full x64 Debug at `751437d2` completed 18/20 entries: 2,114 cases passed,
two failed and 52 skipped. The retained aggregate is the authoritative v2 result;
the focused single-suite aggregates use the earlier v1 format. FileOps completed
180/0/20, Compare 242/0/30 and Commands 897/1/2. Both interactive viewers and the
performance entries passed. This failed Full result is not final acceptance.

Persistent menu-bar hover subsequently passed ten isolated repetitions at
`50d2c4e5`. The failed Full attempt overlapped a separately launched viewer test;
the isolated result alone does not prove the original failure's cause.

The ViewerImgRaw scheduler fixture at `50d2c4e5` failed eight of ten quiet Debug
attempts. Its 1,024-message pump delayed predicate inspection until blocked reads
reached their existing 10-second emergency timeout; one failed attempt also closed
in 676 ms. Commit `f12acdba` gives that fixture one-message predicate checks and
explicit progress-message delivery checks. All existing scheduler assertions,
provider timeouts and the 500 ms close deadline remain unchanged. Ten Debug and
ten ASan Debug repetitions pass. The before/after Debug plugin DLL hash is equal:
this is a fixture correction, not a product-speed improvement. Raw logs and the
ASan close-duration summary accompany the comparison.

All six product profiles built at `751437d2`. The ARM64 Release x86-host compiler
ran out of address space; sequential x64-host Release and ASan Debug builds passed.
Warnings remain in the original logs and are not reported as zero. `50d2c4e5`
subsequently passed all three full x64 builds and detected the seeded ASan heap
overflow. The ViewerPETests target at `f12acdba` compiled in all three x64 profiles.
Final Fresh Full profile runs and final ARM64 cross-builds are separate gates.

## Preferences resource comparison

The same test-enabled x64 Release fixture ran ten Plugins-page round trips per
process, in A1/B1/B2/A2 order: master `0dd0bd6d` versus candidate `50d2c4e5`.
All 40 repetitions passed. The fixture hash and four native receipts are retained.
No build or other windowed test ran during sampling.

Median observed peak private bytes increased 0.384% (1,046,528 bytes, about 1 MiB);
working set decreased 1.071%. Last-sampled process CPU changed in opposite
directions across the pairs, so no CPU gain or regression is established. The
100 ms nominal sampler records actual timestamps; its last CPU sample is not an
exact exit-time total. This measures the whole test-enabled process during UI
Automation and repeated dialog interaction. It is not normal idle memory, a
DLL-versus-static-library isolation, or GPU/frame/idle acceptance. Later `f12acdba`
changes only the standalone viewer fixture and its documentation.

The current opening I19 checklist owns remaining acceptance. Native ARM64 execution
requires an ARM64 Windows host; these cross-builds do not establish it.

## Earlier ARM64 builds at ad7f7735

At `ad7f7735`, Debug, Release and ASan Debug root builds all pass. Their 42 consumer
module records match the actual module bytes, exact DxUi pin and x64 compiler host.
The [ARM64 record](ad7f7735-ARM64/README.md) retains the three original logs,
including 44/60/46 warnings, and the module-identity checks. These results supersede
older cross-builds for compilation; native ARM64 runtime remains unqualified.

## Earlier Debug follow-up at b2be60e4

The [b2be60e4 Full record](b2be60e4-Full-followup/README.md) retains the next failed
attempt: 19/20 entries pass; its only failure is the compiler-host test fixture's
long Git path. The corrected fixture passes all eight dependency tests. Product
Commands, Compare, File Operations, both interactive viewers and performance pass.
Final three-profile admission remains open; this result does not waive the failed entry.

The [earlier portable package and A/B/A record](../2026-09-13-FinalPackaging/README.md)
qualifies clean-extraction smoke and the 11 packaged DxUi identities at `3a43186a`.
All baseline payload hashes match after rollback; final Full and resource acceptance remain separate.
