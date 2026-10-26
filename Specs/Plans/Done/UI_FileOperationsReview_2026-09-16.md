# File Operations UI implementation plan

Status: **DONE for the qualified x64 I26 scope**. Priority: **P1**. Product decisions accepted 2026-09-19.

Implementation planning baseline: `1425d56799bbc5fce158e92be93da92ef9982292`.
Original captures: September 16, source/build identities in their manifest, starting at `9c68026dd`.
Before coding, check drift from the planning baseline in the source map below. Capture success and
the old 76 support checks are not implementation acceptance or a current performance baseline.

## Execution checkpoint

**Completed October 1, 2026.** The six accepted UI decisions are implemented and qualified on x64,
with explicit adoption of DxUi `19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1`.
Release Full/Fresh `20261001T115346Z-104056-8e2cfb6a0b8049218c622dd915a2697e` reports **2211 passed / 0 failed / 46 skipped**, a clean sandbox and
the exact reviewed source snapshot `f2e0a18a04fbc8a95368adc0c946bafb5ebccdf68d81b4a476d0fefbd323593c`. All required mapped cases pass;
the [final evidence packet](../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/README.md) records individual skips, current receipts, historical
failures, exact case bindings, the paired resource interpretation and original capture provenance.
Both V65 x64 builds and both two-case hosted-key runs pass. Broad Debug FileOps remains its
original V55 197/0/20 receipt, with V56's 21/0/0 cluster and reviewed V64 protected companion.

The full-run R0e routing check passes at 1,621,521 us against the unchanged 2,000,000-us ceiling;
its 4,096 parsing/admission decisions make zero provider calls. The independent-volume check
passes at 1,171,000 us against 1,582,200 us. Earlier timing failures remain recorded; this current
pass does not establish their cause. The final eight-process Embedded A78/B19 pair matches the
final Full build profile; the earlier eight-process ProgramDatabase comparison remains separately retained.
The V55 gallery retains 562 static and ten protected harness captures, with 512 reviewed image hashes.

Native 96/144-DPI and the same-HWND monitor round trip are qualified. Native 192-DPI and external
screen-reader execution remain unqualified; consumer ARM64/ASan/release handoff, I18 and H4 keep
their separate owners. HWND-message key tests do not substitute for physical injected input.
No additional implementation slice remains in this I26 owner. The ordered checklist below records
completion scope; the chronological checkpoints that follow preserve their original status and intent.

## Retained execution checkpoints
**October 1 final Embedded comparison passes; only archive/doc closeout remains.**
The baseline-only Embedded rebuild passes with zero diagnostics and unchanged source snapshot;
receipt `9e7ae9c5b52d45afff7bafe691e8427caae342e3aa6d9a8286854b572fb9ceb9`.
All eight runs in `integrated-a78-b19-embedded-final` pass 1/0/0 with clean sandboxes, using
the unchanged final Full B19 executable/receipt. Root independently reviewed native results,
raw metrics, identical source-delta CSV and feature/geometry witnesses. The actual review
`v65-integrated-pair-embedded-root-review.json` is SHA256
`af720ba97df6a576bc357f799457321f47fc3ade427d87b889225ec574f48b6b`.
All twelve median-of-process-median redraw differences are below 0.07%; outliers, the final
private-memory median difference of about 0.414 MiB, bounded cache difference and measurement
limits remain explicit. No causal or peak-memory claim is added.

The successor case binding is `v65-acceptance-results-embedded-root-reviewed.json`, SHA256
`d7535cb24e0c3e21088c000ba101e37a4721537a98009a18495b4972216ed857`; independent root review
SHA256 `46b8977a90f0ad23095d81f6d19bf68447f874f1495ba913d13631f49e2da4de`.
All 62 witnesses remain valid; the presentation witness now points to final Embedded B19-1.
No further build, native test, screenshot or foreground session is pending.

Finish the evidence packet under the unchanged 2-MiB/file and 5-MiB/run limits. Curation v8
caught two malformed README link expressions; v9 fixes them but the actual payload is
7,024,501 bytes. `v65-curation-size-audit.json` identifies the bytes for lossless packaging;
these packaging failures do not invalidate native results, and no packet has been published.
After archive validation, apply the reviewed closeout prose/ordered checklist, validate
spec links, commit/push and delete the continuation automation. Preserve every original run.

**October 1 V65 Full/Fresh passes; final build-profile resource pairing remains.**
Full/Fresh `20261001T115346Z-104056-8e2cfb6a0b8049218c622dd915a2697e` is complete:
**2,211 passed / 0 failed / 46 skipped, 2,257 total**, with all twenty scopes promoted and a
clean sandbox. Root verified the native results, skip reasons and exact acceptance-case binding.
R0e passes at 1,621,521 us against the unchanged 2,000,000-us limit; R4-A19 passes at
1,171,000 us against 1,582,200 us. The original failed runs and their limits remain retained.

Final archive review found that Full's targeted deployment rebuilt B19 using **Embedded**
compiler debug information; the original eight-process A78/B19 pair used **ProgramDatabase**
on both sides. Source, pin and toolchain identities are unchanged, but executable/library
hashes differ. The accepted original comparison applies only to those original binaries.
It does not yet qualify the final Full binary. Preserve B19's post-Full receipt
`7442adedcaf17d954a58b726f49756312beabb2c2426589e4fd71af42584fdac` and executable.

The sole native lane is rebuilding only A78 with Embedded debug information, under
`a78-embedded-v65-build`, preserving its 46-row source manifest and original receipts/binaries.
Then run one new eight-process matched comparison against unchanged final B19, review its
actual resource evidence, and bind that pair into the final packet. No Full or gallery rerun
is required by this profile mismatch. Curation attempts v3-v6 and their diagnostic failures
are retained; none published an acceptance packet. After the final pair, finish archive and
documentation validation, move only this plan to Done, commit/push, and stop the automation.
All remaining work is independent of additional user input or desktop availability.

**October 1 V65 paired resource qualification completes; instrumented Full is running.**
All eight serial A78/B19 processes pass with equal fixture bytes, asserted workload, 144-DPI
client/target geometry and clean sandboxes. Root reviewed the source scope and raw resource
rows; the final interpretation is `v65-integrated-pair-root-review-v2.json` (SHA256
`1f03b7d800efe556394e61e54471ba021b9d022b4e50c6ee1677c0914bbcf102`) under the driver
evidence folder below. The twelve redraw-series median differences are within 0.2%; every
hidden interval preserves render/resource counts, and host/attachment counts return.
Individual timing outliers, positive CPU median differences, process scatter and the one
additional bounded Issues text-format entry remain explicit. This is narrow same-fixture
consumer qualification, not whole-product equality, FPS or driver-memory/peak evidence.

Full/Fresh `20261001T115346Z-104056-8e2cfb6a0b8049218c622dd915a2697e` is active under
`release-v65-full-fresh`, using the reviewed 86/85 driver and frozen source. Preserve this
single native lane through its terminal result. On failure, read the actual new R0e phase
metrics; do not retry unchanged. On success, root must verify its exact cases/receipts, run
the reconciled binder and curator with actual root reviews, validate the archive, then
update canonical contracts and move only this plan to Done. The draft binder/curator now
preserve historical V64 identities separately from current V65. No Full pass or final
archive, commit, push or Done transition is claimed while the run is active.

**October 1 V65 keyboard qualification passes; paired resources and instrumented Full remain.**
V65 freezes 86 logical / 85 physical inputs (manifest
`3222ff1f2e90977a0bae368adb7a582fd54a938ca65b430fad5caaa0b9f121ce`,
source snapshot `f2e0a18a04fbc8a95368adc0c946bafb5ebccdf68d81b4a476d0fefbd323593c`).
Both x64 Debug and Release builds pass with zero warnings/errors. The exact ordinary
Close/Escape and new hosted decision-key cases pass **2/0/0 in each configuration**:
Release `20261001T112207Z-80192-1f5f873df736462f91655cb76638d1a9`, Debug
`20261001T113214Z-63820-2c8d88dff6c54bda82569d94542cb2ec`. Root read the native results
and all four trace outcomes: ordinary safe Enter, same-host Escape preserving the pending
Copy, and Recycle fallback safe Enter/Escape. This proves direct HWND routing; it is not
physical injected-keyboard evidence. The corrected Pester file also passes 24/0/0 from
an external working directory with Pester 3.4.0.

Five predeclared, independent, instrumented Release R0e processes pass **3/0/0 each**, with
totals 1,124,552 / 1,109,171 / 1,070,419 / 1,111,275 / 1,148,281 us against the unchanged
2,000,000-us ceiling. In the first, admission takes 1,115,010 us, parsing 9,251 us and the
rejection tail 131 us. The production Local volume-resolution and typed-query source files
are byte-identical to A78. The bounded API microprobe finds no consistent buffer-allocation
improvement. Fresh handle-based lookup is faster on the existing local directory, but has
unresolved SUBST/reparse/path-identity compatibility differences, so no routing optimization
is adopted. These isolated passes do not resolve the earlier Full-context timing failure.

Next, execute the reviewed A78/B19 eight-process resource pair as an independent gate;
preserve the original A78 product and identical resource fixture. Then execute **one** V65
Full/Fresh with the new inclusive phase metrics, corrected Pester assertion and hosted-key
fixtures. The targeted review found no concrete same-process predecessor cause: Full and
standalone FileOps use fresh child processes. This instrumented Full is needed to attribute
the failure in its original context; it is not an unchanged retry or a relaxed threshold.
The 86/85 Full driver supersedes its earlier 85/84 draft. Existing binding/curation drafts
still need reconciliation with this scope before use. Detailed raw receipts and root reviews
remain under `C:/RedSalamander.Perf/evidence/i26-ui/final-gates-drivers-20260930/`.
No Full pass, resource comparison result, final archive, commit, push or Done transition is
claimed. This primary-only checkpoint does not mutate the frozen candidate or A78.

**October 1 V64 Full finishes red; V65 retains the workload and adds timing attribution.**
Full/Fresh `20261001T093542Z-86260-1127b91800b845a8b9f36b5617113063` is terminal
**FAILED: 2208 passed / 2 failed / 46 skipped, 2256 total**, complete scope and clean sandbox.
Commands passes 923/0/2; FileOps passes 195 cases with one failure and 21 skips. Its R0e
route-containment total is **2,180,864 us**, above the unchanged **2,000,000-us** Release limit;
all 4096 parse/admission counts, the single rejected uncontained witness and zero provider calls
are correct. Independent-volume R4-A19 passes (1,172,000 us against 1,624,050 us).
The second failure is a Pester 3 assertion using file-content `Should Not Contain` for an array
of paths. The primary now uses an exact scalar membership count, preserving the same exclusion.
Original Full V2 SHA256 is `c1bbc57996ffd51b69e8c2d82a44c12d403cc799e7b6e7ecb20b196895ccaf5d`;
post-Full receipt SHA256 is `744ff8d68b3b4756e12edd4541a2b249bc441e414a9b96530f62348ddee6df2b`.

V65 will include the source-reviewed hosted keyboard test below, the Pester correction, and
diagnostic-only R0e phase durations with their normative documentation. The 4096 iterations,
their interleaving, rejection assertions and total-time threshold remain intact; diagnostic
clock overhead stays inside the total. No product optimization or stale path cache is introduced.
First qualify the focused keyboard cases in Debug/Release, the corrected Pester file from an
external working directory, and the provider-capability case with its timing breakdown.
Use that evidence to resolve R0e before another Full; do not retry Full unchanged. The old
V65 scripts expecting only two changed inputs are superseded and must not be executed.
The paired resource comparison and final documentation/Done transition remain outstanding.

**October 1 V64 qualification progresses; final coverage audit adds a bounded P2 key-route check.**
The qualified DxUi pin remains `19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1`; the user
accepted the separately measured Debug caret timing differences. V64's pending-scroll correction
passes the unchanged V63 RED fixture in both x64 Debug and Release, the five-case Release
predecessor cluster three times (15/0/0), and the protected 144→96→144 DPI route (1/0/0).
The latter's ten harness PNG hashes match the reviewed V55 images; its footer focus witness
does not substitute for the focused task-scroll regression test.

Unfiltered Release FileOps run `20261001T090450Z-99932-e4fc6b4b30da433da94d58284796c674`
passes **196/0/21, 217 total**, with a clean sandbox. R0e is **1,266,645 us** against
2,000,000 us; independent-volume completion is **1,172,000 us** against its unchanged
1,582,200-us ceiling, with overlap observed. The 21 explicit skips comprise eighteen missing
remote-profile cases, two separate 7z/Compare-owned cases, and one Debug-only plugin export
whose V55 Debug result is retained. The native runner succeeded. An external wrapper assumed
a native `case_filter` field that does not exist; the original error remains preserved, and
bounded offline verification recovered the qualification without a native rerun. Root checked
the actual ordered cases, reasons, scope, receipt, raw metrics and hashes in
`v64-fileops-full-root-review.json` (SHA256
`bfdf504b4f03b4d0dc9891a958f995656b0bb58258c2d9f874d40506f1fc1e0a`).

Frozen V64 Full/Fresh run `20261001T093542Z-86260-1127b91800b845a8b9f36b5617113063`
is **in progress**, not accepted. Its candidate/frozen manifest remains
`55b1d53c4e12b631dd66609d9d4390bc2dcc0eed50949bd636739793c4063730` (85 logical / 84 physical
inputs); this primary-only checkpoint update does not mutate that candidate.
The final acceptance-map audit found engine-backed same-host/Recycle consequences but no
direct hosted Enter/Escape proof for those prompts. Add focused real-task keyboard cases with
exact task/default ownership and unchanged-file assertions, and map the existing ordinary
Close/Escape case. The primary-only `cmd_pane_fileops_conflict_hosted_decision_keys` case now
covers ordinary default Enter, same-host Escape while the original Copy remains pending,
and real forced-Recycle-failure Enter/Escape. It is **implemented but not compiled or run**;
root rejected a design-only delegate sketch and wrote the concrete fixture using existing
task admission, scoped environment and hosted-key helpers. Root and the cheaper-model source
review found no blocking issue at test SHA256
`ce155d10601d6db24ebfb5da5b2a4ffbbf1bc76d22f06db6349ab08271701129`.
The bounded acceptance audit found no further missing journeys; the successor map references
26 registered Commands cases and 13 FileOps cases/selectors, including separate gallery and
resource witnesses. This source review is not a compile or runtime pass.
Do not substitute direct
`SubmitConflictDecision` calls or synthetic screenshots for these key-route assertions.
Complete the audit before freezing the next test-only delta; latest-source Full acceptance,
the paired resource comparison and documentation closeout remain required.

A78 retains its reviewed old product and pin78. The final comparison requires byte-identical
resource fixture/helper and workload, not identical application implementations. Root verified
this boundary in `v64-baseline-scope-root-review.json` (SHA256
`17a639f06efb00a3cd6da4396e6fc1e1bae601b2a234e0ac46f0f661804b7e3c`).
Never copy the new focus correction into A78 or rerun its Prepare stage. Keep one native lane,
the existing warning/desktop lease/restoration, all original failures and unchanged limits.

The checkpoints below retain the preceding execution history; their forecasts are superseded
by this checkpoint.

**October 1 V63 proves pending-scroll defect; V64 correction qualification begins.**
The test-only candidate preserves old product behavior. Its zero-warning x64 Release build
has receipt `54e237632507f8a0cab50c62cc8340331deb5cd6a511bf93db7d3d6640f06756`, source
snapshot `1f311234bb99a4cc11c6f35dfe5551e8988839a425281b93fccbcc091840784f`, and manifest
`e3dfb1c06a8d422cd1764c8174a5f5c7c13dafaeb0f80e1a545ba5978383db07` (85/84, pin19).
The exact one-case RED run `20261001T084255Z-86664-fca0c8b17eaf4181bf49894b7aaa7c97`
is **0/1/0, clean sandbox**, failing at the intended two-focus assertion: painted scroll 0,
first pending scroll 48, second pending scroll still 48, group top -22 / height 27 against
viewport [0,348]. The disclosure is above the viewport. Root verified runner/native/trace hashes
and the complete copied-input delta. This demonstrates the defect; prior V60/V61 setup failures
remain retained. RED summary SHA256 is
`bf75c45aa73fe2ac1bf4e13a429a6fd573213ea174fbe2072fe993aad8103ee8`.

V64 copies only the reviewed Popup correction (SHA256
`0ba4bef8d41b37678756e80b853cb02799848a8442501d99735c0c936580ab1d`), normative UI contract
and this checkpoint; the RED test remains byte-identical. Qualify the same exact case in x64
Release and Debug, then the five-case Release predecessor cluster three times and the existing
protected mixed-DPI interaction route. That latter route retains FooterOptions focus across
144->96->144; it is not direct evidence of task-action auto-scroll. Use the existing warning,
desktop lease and restoration. Run unfiltered FileOps against history load before another
Full/Fresh gate. No limits or old failures are waived. The later root-reviewed baseline scope
retains A78's old product with the unchanged resource fixture for the required paired comparison.
The plan remains ACTIVE; final Full, resource pairing and documentation closeout are still open.

**October 1 V61 setup failure retained; V63 test-only RED preparation.** I26 remains
incomplete. V58 Full/Fresh `20261001T054944Z-62272-841456e17a5b45f49747fff3203243e2`
is terminal **FAILED: 2207 passed / 3 failed / 46 skipped, 2256 total**, complete coverage,
clean sandbox. The failures are completed-group first-frame badge tooltips (expected 2,
observed 0), R0e route containment **2,191,679 us > 2,000,000 us**, and R4A19 independent-volume
completion **2,688,000 us > 1,876,500 us**. Counts, zero-provider calls, exact bytes and latched
volume overlap pass. No failed result or timing limit is waived. Original raw data is retained
under `V:\RedSalamander.Perf\evidence\i26-ui\raw-runs`; all drivers, receipts, root reviews and
resumable state are under `C:\RedSalamander.Perf\evidence\i26-ui\final-gates-drivers-20260930`.

V59's unchanged-product five-case Commands cluster passes **15/0/0** across three cycles,
run `20261001T072345Z-60020-c5084b3f352a4406a988ea5cbb29b22b`, clean sandbox. Three first
collapsed frames expose both badges. This bounded diagnostic does not replace failed Full.
Source manifest is `7b33dc98d8f951f73984cb8499bbf6341eb6bcd5890f7ba0b33486727f937029`.

V60 adds the raw-provider retirement witness and a two-focus-before-paint regression test;
product behavior remains unchanged. Release builds with zero warnings/errors: receipt
`9530db127bacabc3fea4c0a8131572f53fc27a4ff03d97eb46cb7871022c557a`, manifest
`ec81814f10867cf9533f73514568a8db8f610f9510653e5f12b350bedbe30f9f`.
Its one RED run `20261001T075625Z-19236-b056664a942a45ef997225a8e7996528` failed **before**
the intended assertion: task 2 had two raw providers but no visible Dismiss action. Compact
cards legitimately omit Dismiss. No focus-race reproduction or acceptance is claimed.
The subsequent FileOps diagnostic `20261001T075831Z-83440-782fcc0b4d424e36a7e735f002a7c5dc`
matched no native cases because this suite does not support comma-separated filters.
Both attempts are archived with clean sandboxes; neither was retried unchanged.

Root corrected the test to witness all three task IDs across at most eight viewport pages,
fail closed on incomplete raw traversal, and retain the unchanged first-frame absence check.
Footer focus prevents task-focus restoration from undoing the page scan. The later More
provider now uses its actual action data value 2. A cheaper read-only review confirmed both
fixture defects against production rendering. The candidate for the next RED build must retain
the old product: primary-only `EnsureHostedFocusVisible` coordinate translation and its UI
contract are deliberately deferred from the test-only freeze. The correction translates
painted bounds by `paintedScroll - pendingScroll` before checking visibility; it is not qualified.

V61's zero-warning Release build and one RED attempt are retained. Run
`20261001T082033Z-5620-84fe280455c04ad58e46c4004df002ee` failed before the intended
assertion: lower-action/group provider capture from the same top-scrolled frame failed.
The all-three-task raw-provider presence witness passed. This is a fixture setup failure,
not the expected regression failure, and does not qualify the product correction.
Manifest `977f4f6b0d9edcf744011551b925e97e2aefe6371193e28f4fe00925bb9c28a5` and receipt
`1dda9c594ccac02a9cf99f7873b5337bbb7dc272031832a1deca4f9a8141be19` identify that build.
V62 conditional GREEN scripts must not execute because their expected-RED prerequisite failed.

For V63, root expands all three case-owned completed cards, verifies the short 320-DIP client
height, then chooses an existing More provider among those IDs from one top-scrolled frame.
Live-to-history ordering is not assumed. Granular provider/QI/bounds/viewport diagnostics
precede setup assertions; the non-vacuous pending-scroll assertion and original visibility
and raw-provider retirement assertions remain unchanged. Test SHA256 is
`89bf23a6cb7d3b062bf8a15895dead2a76210715eeb267627a4dddf521d81ec3`.
Cheaper read-only review was checked against the actual rendering and history-order paths.
Old product and old UI contract remain in the RED candidate; no setup failure is waived.

The corrected performance diagnostic completed six serial processes, three per exact phase,
all **3/0/0** with mandatory Setup/Cleanup, original limits and clean sandboxes. R0e values
are 1,363,597 / 1,092,681 / 1,084,408 us against 2,000,000 us. Independent-volume completion
is 1,218,000 / 1,140,000 / 1,140,000 us against each run's unchanged derived ceiling, with
overlap observed in each run. Root verified native results, runner results and raw metric hashes
in `v60b-perf-root-review.json`. This establishes variance, not Full acceptance.
Source review found repeated whole-tree UIA publication during
attached-control configuration and retained 407-ms rebuild/170-ms update observations; the
hotspot predates pin19 and causality is unproven. No performance optimization is yet applied.

The A78 formatting-only baseline builds Release without warnings/errors, receipt
`b9781f0dffb29fff29fb8cc86f738f0b040879638ed462b0793b8e19d3148d77`, unchanged pin78.
Root verified all 46 source rows and identical resource fixture/helper bytes. Never repeat its
Prepare stage. The later baseline-scope review requires equal fixture/workload scope and preserves
the intended before/after product difference; it does not copy later product fixes into A78.

Next: freeze V63's test/checkpoint delta, build once and run its corrected RED fixture against
old product, then qualify the narrow focus correction if the expected failure is demonstrated. Resolve
remaining performance gates before another Full; resource pairing and final curation stay held.
Only one native/build agent runs; root owns repository edits. Preserve V55 screenshot provenance
and historical failed receipts. Conditional V58 closeout scripts remain inapplicable to failed Full.
**October 1 V58 source-guard repair prepared; V57 Full retained as FAILED.** The unchanged V57 Full
run `20261001T035746Z-80936-af79250aa05f43769b80d16fc565ba46` completed with **2206 passed,
4 failed, 46 skipped / 2256 total**, clean sandbox. File Operations passes **196/0/21** and all
22 mapped I26 Commands cases pass. Two failures are Commands source guards. Root traced both to formatter
alignment/line wrapping, then checked all 523 `.find` literals in that source-guard section against
all sixteen formatted files. Five expressions changed presence; three were hidden behind the first
reported Phase 0 failure. Production tokens remain unchanged.

The primary-only correction in `Commands.SelfTest.PluginConfig.cpp` preserves the exact required
and forbidden expressions and Scheduled Rename body uniqueness, ignoring only space/tab/CR/LF
for affected checks. It reuses the adjacent existing `std::erase_if` pattern and changes no product
code, policy, timeout, case registration or performance threshold. Candidate SHA256 is
`a2eef35462baeca89d85eef6f33067687e729209e3ef6f6ef295a899d068c19f`.
Pinned formatter 22.1.3 passes. The read-only structural proof
`final-gates-drivers-20260930/v58-all-source-guard-predicate-evidence.json`
(SHA256 `88d3ccf34aef41522fa0b6d02bacdda336621b616bcbee80c4eb5d58bb00f62b`)
checks ten old/new source rows, accepts formatting changes and rejects all 86 required-expression
removal / forbidden-expression insertion mutations. It is not native runtime evidence.
`Testing_SelfTests.md` now states the formatting-tolerance rule without upgrading source checks
to behavioral coverage. Root verified the cheaper-model review against the complete five-expression
correction and recorded `final-gates-drivers-20260930/v58-source-guard-root-review.json`.
The third failure is the Pester durable-state source guard: formatting the publication call
exceeded its arbitrary 220-character span. The primary correction matches the complete call,
argument order and fail-if-exists/replace policy with whitespace tolerance. All **191 source-contract
Pester cases pass**, and ten invalid-call mutations are rejected; this is structural evidence only.
Tool inventory, usage documentation and `Testing_TestCoverage.md` carry the same scope.

The fourth failure is a fresh ViewerPE child running the VLC focus/HUD case. It exited before
the deadline with a nonzero status, but the old harness did not preserve its numeric exit code
or separate child diagnostics. Earlier child executions passed. Its cause is unproven; investigate
with bounded retained diagnostics rather than waiving or reclassifying the failure.
One unchanged-binary exact-case diagnostic, run
`20261001T053040Z-99884-f514c443f7524f4389c25b4c73bf12ee`, then passed with exit zero,
all sixteen assertions, clean sandbox, unchanged cursor and original foreground restored by
normal window closure. Its protected wrapper used the existing artifact/desktop leases and native
warning; no focus repair or timeout extension was applied. Captured stdout SHA256 is
`2fdfacea3ffa93756a1a762fa4e6935122caca5042630fc073f8823922a5ea65`.
This diagnostic does not explain the earlier failure or replace final Full/Fresh acceptance.
The complete V57 raw run and post-Full receipt remain retained. Receipt
`7faa22aa852663b37bbd4a6c42f22d4186169f20e73ab2a6900f46622da12f02` retains source snapshot
`dcd29a8210761877b254e93acd35f9224a7738c04f078067841da7a8daf445b8`.

Next: freeze the reviewed eight-path V58 test/tooling/documentation delta,
build both x64 configurations, execute the two exact source-guard cases before another Full/Fresh,
then qualify the final equal-feature resource pair and close documentation. V55 captures and V56
interaction results retain their own provenance. The reviewed V57 binder/curator below need an
explicit V58 delta/receipt update; do not relabel old results or use their old pinned snapshot.

**October 1 V57 final qualification in progress.** Root verified all 85 logical / 84 physical
inputs across primary, isolated and frozen roots. Manifest
`86d11043dbae7aef033fb57507feb1fd68d45c2f98758708ae4d5302dfd561f3` retains the same qualified
pin19 and isolated base. The V56 delta is exactly seventeen physical paths / eighteen logical
rows because the Resource.h casing alias names the same file: sixteen formatted native files
plus this checkpoint. The immutable root review is
`C:\RedSalamander.Perf\evidence\i26-ui\interaction-v57-20261001\root-source-review.json`.

Both full test-enabled x64 builds pass with zero warnings/errors. Debug receipt
`016bd1ed868d02aacf966671079a6971396006532e05339210ff269f4dfa39d9`;
Release receipt `6f96996eec373c04c199300d7a4bb42d6304580055cc3a70b6c9da1d57e49171`.
Release Full/Fresh is now running on marked V: with explicit C: alternate. Hold the resource
pair until its terminal result and post-Full receipt/dependencies are reviewed. The baseline
resource helper then needs the same whitespace-only formatting and a new retained A78 build,
so both sides execute byte-identical final fixture/helper sources. The original baseline remains.
Primary specification/tooling preflight reports zero blocking findings. Final resource acceptance,
evidence curation and Done/index/link/automation closeout remain open.

Final evidence tools are ready under the external `final-gates-drivers-20260930` root.
Root reviewed the cheaper-model V57 curator, corrected two archived-schema mismatches,
and ran its bounded read-only source/build/carried-evidence preflight successfully. The final
script is `curate-i26-final-v57-draft.ps1` (SHA256
`b2203ff6625f80979d6d4c65ed61f64b6675ee0325f19af8760b44de64d52649`);
`v57-curation-root-review.json` records its exact scope. The case binder is
`bind-v57-acceptance-results-draft.ps1` (SHA256
`4c26310aaf56996faa9e16a011c7e5d961c3d5d81dab50d2664a777e30755bbc`), with
`v57-binder-root-review.json`. Both require actual final Full/pair inputs; neither has been
executed for final acceptance. Original V55/V56 run identities and failed diagnostics remain.

**V57 formatting review.** The repository CI formatter is pinned to
clang-format 22.1.3. Its non-mutating check found remaining formatting in sixteen changed native
files; the locally resolved 22.1.8 reports the same files. Root retained the original bytes and
applied the pinned formatter. A cheaper-agent review caught an important ordering dependency:
the Preferences dispatcher includes helper fragments by definition order, so its include remains
last in a separate documented group. No dispatcher case order or behavior was changed. All 42
changed native files now pass the pinned formatting check. Root verified identical raw tokens in
thirteen files; the other three contain only reviewed, same-membership ordinary-header reordering.
No executable statement, constant, resource text or visual-fixture logic changed. The external
`v57-format-root-review.json` records all before/after hashes and the include-order correction.
The original V55 captures and V56 interaction cluster keep their own
provenance; final Debug/Release builds, Full/Fresh and the resource pair must use formatted V57.
No V56 Full run was started, so this preparation does not discard or repeat a Full result.

**October 1 V56 retained fixture qualification.** Root verified all 85 logical / 84 physical
inputs across the primary, isolated and frozen roots. V56 manifest
`2d48670000336a7afe9351113d1c9e263e9b25b2767b29bbf9bccba1369eadeb` retains qualified DxUi
`19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1`. Its only compiled change from V55 is the reviewed
Commands readiness fixture; five documentation files also changed. All production and FileOps
gallery fixture bytes are unchanged. The root equivalence proof is
`C:\RedSalamander.Perf\evidence\i26-ui\interaction-v56-20261001\root-source-review.json`.
The broad Debug FileOps and 572 reviewed V55 captures below retain their original run/build
provenance; they are carried evidence, not relabelled V56 executions.

The V56 full x64 Debug build passes with zero warnings/errors, receipt
`16f2bf76ed1dc24d1cc6863ee783e1a31e9b4b12f0d506b40a65cd5b93f1fd17`
(file SHA256 `fd05ab174e4f7ecb2fcc5878062e8f069e078d563cb54ac7db11e26d6a25b7c8`).
The ordered seven-case predecessor Commands cluster passes **21/0/0** across three repeats in
one process, run `20261001T032939Z-69528-0f2e124ea9dd4721a752ccd97d5b9d93`. Root independently
checked each exact sequence, native/result hashes and clean sandbox. Release also builds with
zero warnings/errors, receipt `51c8e065e36b5989569fc0942b96f84280f9cddcc4cd8d2148337338d786df45`
(file SHA256 `6b9f39c458d21f29da224b8d684e66087a8a57fdb59d8b4e271f3b5c319620e9`).
The sole Sol lane holds before Full/Fresh for the V57 formatting freeze described above.
Original V55 failures and all five fixed-binary R0e diagnostic samples remain.
No final Full/resource pass, Done movement or automation closeout is inferred from these builds.

**October 1 V55 retained qualification and diagnostic history.** The frozen candidate uses DxUi
`19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1`; manifest
`d0fccec5b902e58352c457d6f820ce4c7efd98ed6a423067e4bb9831ec4cdba1`
contains 85 logical / 84 physical inputs. Root verified every input across the primary,
isolated and frozen roots before building. Subsequent primary documentation updates do not
change that immutable candidate or its native-test inputs.

- Both complete x64 Debug and Release builds pass with **zero warnings/errors**. Debug receipt
  `66a1bce053e7c1b78ea1236756ae350ffaeb5d42c0e6082d25545c5bf215ea29`;
  initial Release receipt `6c6155ca8f47676c435936477b5fa8131795fbce871bbc920b357abb8bd4fad9`.
- Broad Debug FileOps run `20261001T010532Z-89108-35279440c388486683c7646e58d45f24`
  passes **197 / 0 failed / 20 expected skips**, with a clean sandbox. Two skipped 7z cases
  belong to Compare coverage; eighteen require absent remote profiles/secrets. Independent-volume
  coverage passes, but its fallback-selected alternate drive was not logged.
- Current static run `20261001T012949Z-76256-65bb89cb9d96432b893602e7fe21a367` and protected run
  `20261001T013754Z-81300-8feb927cdd0b48518e3da429111876e5` each pass **1/0/0**, clean.
  The original static external checker rejected the normal `[tick]` trace prefix; root repaired
  only that parser and revalidated the unchanged archive, preserving exact ordered 96→144-DPI
  assertions and the original failure. No native test was rerun. The protected run proves same-HWND
  144→96→144-DPI transitions with native focus and action identity retained.
- The [V55 gallery](../WIP/FileOperationsUiReview_2026-09-16/implementation/2026-10-01-v55-final/README.md)
  publishes **562 static + 10 protected captures / 512 reviewed unique image hashes**. Root viewed
  all sixteen new hashes; the other 496 exactly match retained reviewed images. Eight new full-path
  root/scrolled captures verify the clipped caret and complete actions on both physical monitors.
  Ordinary rows may abbreviate a long path; the read-only detail exposes its exact full value.
- V55 **Release Full/Fresh FAILED** run `20261001T014547Z-61860-a2e7aafe26cb4786b9daef90990cc403`
  finishes **2,208 pass / 2 fail / 46 skip**, complete coverage and a clean sandbox (2,256 total).
  uses `V:\RedSalamander.Perf` and alternate `C:\RedSalamander.Perf`. Compare passes 249/0/23;
  Commands exposed `cmd_pane_fileops_speedLimit_prompt_live_dx_interaction`: the exact Resume
  UIA invocation failed after its fixture observed `HasStarted` and paused the worker. Root and
  Astra confirmed that `HasStarted` precedes discovery completion and that the test's snapshot
  reads the live model, not the rendered action set. The original trace does not identify which
  action branch was visible, so this is an insufficient fixture precondition, not proof of a
  production UIA defect. Commands finishes **922 pass / 1 fail / 2 skip**. Root reviewed
  Astra's fixture-only repair: wait for the exact visible/enabled Pause before setup-pausing;
  then wait for each eligible Resume/Pause/Cancel provider and invoke once without a message
  pump between acquisition and invocation. The existing five-second bound, source/neighbor
  assertions and ownership guard remain. clang-format 22.1.8 and the whitespace check pass;
  primary Commands source SHA256 is
  `3bd4d9bb627e0cdaf518453b994f00ba369bf2a40b83e262d8e8b8be71a35b8a`.
  This repair is not yet natively qualified. Do not retry unchanged or mark V55 passing;
  the sole Sol lane retains the original run to completion, then prepares a new V56 freeze
  for root verification and an ordered seven-case predecessor cluster repeated three times.
- V55 FileOperations subsequently finishes **195 pass / 1 fail / 21 skip**. The additional
  failure is `FileOps_ProviderCapabilityMatrix`: R0e duration is **4,130,567 us**, above its
  unchanged **2,000,000-us Release ceiling**. All count/safety predicates pass (4,096 parses,
  4,096 admissions, one rejected uncontained witness, zero provider operation calls).
  V47 Release measured 1,520,542 us and V55 Debug 1,471,323 us on the same V: test volume.
  Root/Astra source review confirms unchanged admission/fixture/Local route implementation;
  the DLL/executable artifact hashes differ, so no identical-binary claim is made. Neither a
  volume explanation nor a product regression is established. Five independent fixed-V55 Release
  diagnostic processes on V: will retain every result before the V56 freeze. This is variance
  diagnosis, not final qualification. Preserve the failed Full and ceiling; no retry-until-pass
  or timing waiver is authorized. The failed run's post-Full build receipt is
  `a96d1cc3c250a450b0072d4402c167b9a7483f3322e460777f5068f822e7aa29`.
- The fixed five-process diagnostic set is complete: **fail/pass/pass/pass/pass**, with R0e
  durations **2,667,544 / 1,091,386 / 1,092,836 / 1,095,473 / 1,637,394 us**. Root independently
  checked all raw metric/result hashes, exact Setup/target/Cleanup membership, counts and clean
  sandboxes. The same binary exhibits timing variability; a deterministic source regression or
  scheduling cause is not established. No threshold or provider source changes are justified by
  these samples. Preserve both failing records. Next: freeze V56 with only the reviewed Commands
  readiness repair and documentation changes, verify that delta, build both x64 configurations,
  run the ordered seven-case Commands cluster three times, and run final Release Full/Fresh plus
  the eight-process resource comparison. If the verified delta remains test-only, the unchanged
  V55 production/visual-fixture qualification and 572 captures remain valid with their original
  provenance; do not relabel them V56 or repeat unaffected captures merely for a new binary hash.
  The new Full must still pass the original R0e ceiling; these diagnostics do not satisfy that gate.
- The eight-run A78/B19 pair is held for a passing final Full and its post-Full build receipt.
  Full's governed rebuild changed the B19 library artifact hash; re-capture/review exact dependencies
  after it finishes. The old expectations and initial receipt remain evidence, not permission to
  bypass artifact checks. Frozen V55 sources remain immutable during the primary-only fixture repair.
- Final resource acceptance, current Full results, evidence curation, documentation checks and
  Done/index/link/automation closeout remain open. Native 192 DPI, external screen-reader execution
  and the separately owned consumer ARM64/ASan/release handoff are not claimed.

External continuation authority: `C:\RedSalamander.Perf\evidence\i26-ui\final-gates-drivers-20260930\v55-final-gates-routing.json`.
The dated entries below are retained diagnostic history; their pending statements describe that
earlier candidate, not the V55 results above.

**September 26–27 resume.**
- **October 1 V54 diagnostics pass / V55 final preparation:** final V54 manifest
  `01604b0ada1c0ce248a0deb70c9ddc3046e0957ad56061ddb0e51a8c44874808` was independently
  verified across primary, isolated and frozen sources (85 logical / 84 physical). Its earlier
  unbuilt freeze remains in `interaction-v54-20261001-pre-checkpoint` after a checkpoint-only
  race was detected before building. V54 x64 Debug builds with zero errors and one C5246 test
  aggregate-brace warning; receipt ID `fda750ac907acd98f7a9a7a742e668e34b01d5583822841152141ef577a9e8b0`
  and file SHA256 `3fbd66a6ad3c3f04ceb6cf980b38d70db2dbcfda88e80390639196b03294af3d` remain
  diagnostic-only, not final zero-warning qualification. Application run
  `20261001T005321Z-68628-51560f35c60448fbba926bd76daeb70b` passes **3/0/0**, clean sandbox:
  path disclosure, external settings reload, and remote case-sensitive nested inspection.
  Root checks original trace: generations 2/4 now retain both `s3` and `s3table`; inventory
  stays 14 entries, with S3 deferred after reload and loadable for the remote fixture.
  Separate diagnostic VSTest run `20261001T005600Z-39196-e4f4c7b3784442b287b20c98346a1baf`
  passes all **three exact lifecycle methods**, including the new mixed-alias regression,
  with zero failures/skips and a clean disk audit. Neither run retried a failed case.
  Root corrected the test's nested array braces (SHA256
  `7cc9a1c51359de610e4470de48d83bdd3bef6d2b3931ae0a642502f889d4b2e3`) and removed temporary
  discovery tracing plus incidental formatting, restoring the manager's exact original SHA256
  `24ea9c85c4dbf98a7712c262050d4a7c59e61716c161b60c032b95426acc0927`. No diagnostic buffer
  remains in the final source. Next: freeze V55 with the same 85/84 scope and pin19; root
  verifies it before the sole Sol build/test lane performs zero-warning Debug/Release builds,
  broad Debug FileOps, current static/protected French/native-DPI gallery, Fresh Full Release,
  and reviewed A78/B19 resource comparisons. Final gallery review, documentation and Done
  movement remain required. No additional mirrored source-string tests are needed: behavioral
  lifecycle and application regressions exercise the contract directly.
- **October 1 V54 lifecycle correction, reviewed and awaiting execution:** Astra implemented the
  shared-helper fix in `PluginModuleLifecycle.h`: successful per-reference unloads retain owned
  alias metadata while any owner still defers the same DLL path; refreshes deduplicate aliases,
  partial sweeps preserve them, and the final release removes all placeholders. Root verified
  that no COM reference or module pin is copied and that scratch pointers remain within unchanged
  source vectors. The standalone regression covers both traversal orders, repeated refreshes,
  partial sweeps, unrelated paths, exact pin release counts and final recovery. The production
  external-reload case additionally checks every initially known provider ID after both preference
  changes and restoration, allowing deferred/nonloadable entries without adding a fixture refresh.
  Root reviewed and formatted the three sources with clang-format 22.1.3; diff checks pass.
  Final SHA256 values: lifecycle header `deea3dbfcc8b5e0e873d22f72a8b76bd3283114f057ed5112dbc8644acee0eee`,
  standalone tests `4ad5574c163234f0ea0d7da7ca1c951e562a5fda3348119104931343a382cfc3`,
  Commands FileOps `174eb489a52df7415c56c581044f73cc15e6200a5532a10ea14c22d61d70cfea`.
  File-system/viewer lifetime contracts and the shared-helper catalog now specify alias retention.
  Sol owns the sole native/build lane: prepare V54 **85 logical / 84 physical** inputs on pin19,
  stop for root manifest verification, then run the focused lifecycle and three-case diagnostic
  qualification. The old-helper unit failure is inferred from its lost-alias assertion, not an
  executed red test. No runtime pass or final I26 gate is claimed yet.
- **October 1 V53 pointer/caret pass / deferred-alias diagnosis:** manifest
  `6fbb36d5ec12908be9934d48217600d8acc7da4e1a09f27294ec2e2bceb3562c` passes independent
  three-root source verification; full x64 Debug build has **zero warnings/errors**, receipt ID
  `40a903f98109e6a71f9c2684379f7616169944046fee09475a014d18e4cf5152` (file SHA256
  `c52b05bc9c6160a88db4d799bc3690e04a0537e65ebd28a1f3c3452054f19502`). Run
  `20261001T003246Z-27000-d611492be5d94007955900dcfe50bf9e` finishes **2 pass / 1 fail / 0 skip**,
  clean sandbox. Throughput disclosure and external reload pass; Space/UIA/mouse each opens the
  exact path with editor focus and unchanged eight hosts/attachments. Root visually reviews the
  French 480-DIP/144-DPI scrolled path capture: no caret spill and complete Copy/Close captions.
  This is scoped proof, not final gallery/resource acceptance.
  The remote S3 fixture still fails because its logical `s3` entry disappears. Crucially, trace
  generations 2/4 show deferred S3.dll with **only `s3table` retained**, and deferred Curl with
  sftp/scp/imap but no ftp. The earlier inference that both S3 aliases were absent was incorrect:
  the inventory diagnostic only reported `s3`. Per-alias mixed unload outcomes omit siblings when
  discovery skips a shared deferred DLL. General trace truncates before final rows, but these
  candidate/placeholder records are retained and sufficient to select the lifecycle defect.
  Astra High now owns the sole implementation/test slice for complete deferred-alias preservation;
  Sol has retained the original failure and is idle as native/build owner. No arbitrary fixture
  refresh, waived gate, or passing retry. Final qualification follows root review and meaningful
  multi-alias lifecycle regression tests; remove temporary discovery tracing before final resources.
- **October 1 V52 diagnostic build failure / V53 correction:** V52 manifest
  `11ca675d9ef56ffad92e7ecd7db093fb189cc876c1998175cdd8f161cc5cf3de` was independently
  verified across primary, isolated and frozen sources. Its full Debug build failed before tests:
  `PerformanceTests2` imports the manager source but does not link the application trace owner
  (`SelfTestCommon.h` unavailable); two diagnostic formatter counts also produced C4365 warnings.
  The failed build and the earlier pre-format freeze remain retained. Root restricts temporary
  tracing to test-enabled builds with the application harness header, and explicitly converts the
  bounded counts to `ptrdiff_t`. No project, discovery policy or release behavior changes.
  Corrected Manager SHA256 `670a10faa3a636d2a21da511364d5ab7d8c86dc7ccf1adc117298595f8c6c2dc`.
  Root reviewed Luna's pointer coverage observation and now requires an in-client mouse target,
  exact editor text and focused editor after every opening route; Commands SHA256
  `a593a3303074b7179906c012b66d2d454bff11809fcb31be49c413c22f4e3657`.
  Both corrected sources pass clang-format 22.1.3; prepare a new immutable V53 diagnostic freeze
  with the same 83/82 scope and pin19, then build/run the three cases once. Final app gates stay open.
- **October 1 V51 terminal / pointer root cause:** the unfiltered 925-case Commands run
  `20260930T232904Z-102716-4e078ce6e3eb4721b4b0eb89f8405b8b` is archived with **921 pass,
  2 fail, 2 documented opt-in skips**, clean sandbox. The two failures are the path-detail mouse
  route and S3 discovery described below. Root verified that both `DxUi::Label::HitTest` overloads
  return null: the consumer's interactive path subclass inherited this intentional passive-label
  behavior. It now overrides both overloads using `Control::HitTest`, preserving visibility,
  enabled-state and bounds checks. The existing direct mouse route is its regression witness;
  keyboard and UIA passed in the failing run. Formatted Popup source SHA256
  `cebf940372483f6bd6bc4b8c9e58fa6faaf9856d7629f5058469db536e07bba8`.
  V52's first unbuilt freeze captured mixed line endings in that insertion and is retained as
  pre-format evidence; regenerate its manifest before build. The isolated lock now explicitly
  selects qualified pin19, with the previous lock retained. Next: build reviewed V52, run its
  three diagnostic cases, establish the S3 omission cause, then remove opt-in diagnostics before
  final qualification. No passing app qualification is claimed for this source correction yet.
- **October 1 V51 demonstrated failures / V52 diagnostic preparation:** canonical Commands has
  reproduced two failures while the original run continues unchanged: throughput disclosure's
  new pointer detail route fails its combined field/host assertion after Space and UIA Invoke pass;
  remote case-sensitive inspection fails because S3 disappeared after the throughput external-reload
  case (inventory 14 to 12). This is not yet evidence of a host leak or the cause of S3 omission.
  Canonical hide/reopen passes exact bounds with no active motion. Root adds separate field/host
  assertions and pointer geometry/motion diagnostics, plus reviewed `ENABLE_TESTS`-only, opt-in
  discovery tracing using the existing sandboxed Commands log. The latter records bounded candidate,
  enumeration/deferred-alias and final-inventory facts; it changes no discovery policy or refresh order.
  Root and Astra review lifetime/buffer/root safety; clang-format 22.1.3 and diff checks pass, compilation
  remains pending. Primary source hashes: Commands FileOps
  `d1b2ea75d2769b70181062b3e5a1dbc4d2a65dd510cf8e1b18fae4fa54da61c3`, FileSystemPluginManager
  `539d09be3e14ca010bbbad678ed063ae90543cc1476794e793c2c48d4272e407`.
  After V51 terminates and all evidence is retained, explicitly select qualified pin19 in the isolated
  consumer and freeze **V52 diagnostic 83 logical / 82 physical inputs**, including the manager trace.
  Reviewed prepare/build/three-case driver hashes are recorded in the external driver folder; run
  throughput, external reload and remote inspection in one process with tracing enabled, not a blind
  full-suite retry. Remove diagnostic tracing after the cause is established before final performance
  acceptance. Primary UI contract/proposal now distinguish published V47 captures from newer path detail;
  the specification inventory passes with zero blocking findings. Earlier preflight CLI/schema errors
  remain in their logs. No parent gate is closed by these diagnostic changes.
- **October 1 approved caret handoff:** the developer explicitly accepted the recorded Debug
  clean-median/dirty-P95 timing differences and directed adoption. Root completed production,
  test, packet and cheaper-model review; all three documentation/dependency validators and
  format check pass after closeout. DxUi's separate caret plan is **Done** and commit
  **`19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1`** is pushed on
  `codex/multiline-caret-viewport`, with unchanged qualified production/fixture hashes and a clean
  worktree. The acceptance is bounded to those recorded measurements, not a causal explanation
  or waiver of other gates. Primary `Dependencies/DxUi.lock.json` now selects this exact commit;
  isolated `C:/RsI26Adopt` deliberately remains on pin49 until its V51 canonical Commands run
  terminates and is retained. Do not mutate a running candidate. Next: resolve demonstrated
  Commands failures, explicitly update the isolated lock, restore exact source, freeze V52,
  build and qualify new caret/path interactions plus final application gates. This library
  closeout does not close I26, I19 platform deferrals or other owners' work.
- **October 1 V51 exact geometry result / canonical Commands running:** old-pin diagnostic run
  `20260930T232432Z-101892-977dad67705b4c3792e8ba4240cab6eb` passes **2/0/0**, clean sandbox.
  Both repetitions preserve `(1297,625,2039,1371)` at 144 DPI, with no auto-resize/layout motion
  before or after. This does **not** reproduce or resolve the V50 intermittent geometry failure.
  Root verifies V51's 82/81 frozen/current bytes; manifest
  `de9ee16a5dcc1e10e6f8136b1ecc573253453fb2d7107943085d7c0ce8aa59b2`.
  Sol owns the next unfiltered **925-case canonical Commands** run, original order/repeat1, no
  automatic retries, retaining S3 first-loss inventory and geometry diagnostics. The first driver
  preflight read a nonexistent native `case_filter` field and stopped before native execution;
  root corrected that schema mistake while retaining exact hashed membership/order checks and
  the actual runner's unfiltered assertion. Original driver/error remain retained. Reviewed
  corrected driver SHA256 `189e3d9da8e7c24cfe056a90b2fe2d001be743de19c9cd8762d38fdd1c58bfbf`.
  Luna independently checked the library packet; root corrected its overly broad range-overlap
  sentence to refer only to flagged metrics and reverified all 182 members. Debug timing advice
  remains pending; no library pin/commit/Done movement is implied by these diagnostic app passes.
- **October 1 library qualification / V51 geometry diagnostic:** the multiline caret candidate
  completes all **60 x64 functional suite runs** (20 each in Debug/Release/ASan); Menu retains
  nine nonactivating capability skips per profile. Final strengthened Embedded witnesses pass
  **3,030 checks**, including 30 states with zero escaping pixels, in each profile. All three
  ARM64 cross-builds pass; native ARM execution remains separate. Final relinked Embedded binaries
  and earlier full-suite binaries have distinct recorded identities. Skills/specs/dependencies/
  format and five-theme gallery review pass. Root independently verifies all **182 packet members**
  in DxUi `Measurements/MultilineCaretViewport/2026-10-01`; the retained stale-timestamp build failure
  and original failing witness remain recorded. No claim that an unpaired benchmark is acceptance.
  Release's bounded interleaved comparison has no aggregate flags. Debug clean median rises
  **0.031 ms (8.5%)**, dirty P95 **0.221 ms (9.4%)**; the fixture never paints a TextField, causality
  remains unproven, and the requested developer disposition is pending. New library plan remains
  WIP and the consumer pin remains qualified `49c963f5751a6754e6b08e0362b96cfc7fc64414`.
  Independent work proceeds on **V51 as an old-pin diagnostic**, not caret adoption: 82 logical/
  81 physical frozen inputs, reviewed direct path routes/early guards and geometry motion logging.
  Debug build passes zero warnings/errors, receipt
  `a2bdcfe3cf25d2e021a5204e4872d6636583074875c59267a347bca61b035538`.
  The reviewed exact hide/reopen case runs twice in one protected session, with no automatic retry.
  Primary harness source contracts pass **191/0/0**, receipt `post-v50-harness-source-contracts.xml`
  under `final-gates-drivers-20260930`. Next: inspect exact geometry/motion evidence; fix only a
  demonstrated cause. Adopt the caret commit only after its performance disposition and final
  review, using a new frozen version; then execute affected and remaining app gates.
- **October 1 caret correction / reviewed consumer test guards:** the isolated library candidate
  clips multiline caret painting to the existing text viewport; root reviewed the balanced nine-line
  production diff. The retained unchanged-production WARP witness paints 20 caret pixels outside
  the viewport; the candidate passes all 30 editable/read-only, partial/offscreen/visible states at
  96/144/192 DPI and 3,018 Embedded checks. Full library qualification is still running. Initial common
  benchmark flags are retained and being investigated with paired measurements; that fixture contains
  no TextField/focus and does not execute the new clip. No performance acceptance or pin adoption yet.
  Root reviewed Sol's external patch and applied early-return guards before further detail input after
  a failed prerequisite, plus existing auto-resize/layout-motion diagnostics around hide/reopen. No
  geometry policy, settling wait or assertion was changed. Consumer source remains unbuilt, SHA256
  `0d5d943c93bdb5ef77f97b7178829e869e3515e6308450ebe38eb4d5b58ff004`; clang-format 22.1.3 and
  targeted diff checks pass. Luna's canonical Commands driver was corrected after root caught an
  undefined post-run variable; eight synthetic validator checks pass without executing native tests.
  Reviewed external artifacts: `v50-readonly-geometry-path-review.md`,
  `v50-path-guards-motion-diagnostic.patch`, and `run-v50-canonical-commands-review.md` in the driver folder.
  Astra retains exclusive native ownership; consumer freeze/build waits for qualified library adoption.
- **October 1 V50 predecessor result / library ownership handoff:** run
  `20260930T224753Z-69412-e4b288d64db04d36ac0842bab7ced4c3` finishes **47/1/0**, clean sandbox,
  across 24 cases repeated twice. All six previously failing V47 cases pass both repetitions;
  this subset does not close the S3 full-order gate. Repeat 2 of
  `cmd_pane_fileops_conflict_close_hides_escape_cancels_snapshot` fails exact hide/reopen geometry.
  The primary fixture now logs both rectangles, DPI and maximized state at that assertion; this
  diagnostic and the direct path-detail routes below remain unbuilt. No product geometry change
  or passing retry is claimed. Full receipts remain in `final-gates-drivers-20260930/debug-v50-predecessor-clusters`.
  Sol returned native ownership. Astra High now owns the isolated DxUi caret fix and qualification;
  no app native workloads or competing implementation edits run during that ownership. Root reviews
  the library change before explicit adoption, then freezes the new consumer inputs and resumes
  focused routes/geometry, canonical Commands and final required gates.
- **October 1 V50 detail pass / caret painting defect:** Debug build passes **zero warnings/errors**,
  receipt `a4bc8fc315852ca2721445c00bbe6bafb02cbcc4e19a422021bded100d3667ba`; root independently
  verifies all 82/81 frozen inputs (manifest `55741c098e7d84c8d23b8649c7b29364f385c13711890f8875bf8b4456970b11`).
  Focused run `20260930T224519Z-47436-f33ac5a0d95946d8b17a62a94ebb6b4c` passes **1/0/0**, clean
  sandbox: 6,155-unit Unicode/UNC path, exact TextPattern/copy, real viewport movement, read-only
  edits, Escape, stale-provider rejection, row removal/hide and 40 host-stable open/close cycles.
  Root inspected the actual scrolled screenshot and found the caret paints below the field into
  the button gap. Pinned DxUi TextField paints its multiline caret after the text clip has ended.
  This is a real library rendering defect, not an accepted visual limit. A bounded Astra High
  library fix/qualification is queued after Sol's already-running 24-case predecessor repetitions;
  no competing native workloads. Primary also adds direct Space/UIA/pointer opening, Copy-full
  and Close-button activation and both Tab-wrap edges within the existing path fixture; these
  added assertions are not built yet. The app remains on qualified pin49 until explicit adoption
  of a reviewed, newly qualified library commit. Do not close or publish V50 as final acceptance.
- **October 1 V49 focused failure / V50 test corrections:** V49 Debug builds with **zero
  warnings/errors**, receipt `6536368daddc917e82db559de0833c14f93992fc4008d4a98112b69468ebfb9b`.
  Focused run `20260930T223402Z-83304-df84806ea8854ebc9d1a185e4d841818` fails **0/1/0**, clean
  sandbox, before the new editor checks. Root and Sol confirmed the fixture cleared the path
  tooltip with Escape, then expected a same-control SetFocus no-op to reveal it again. Primary
  now traverses to the preceding path before refocusing; Escape still dismisses without losing
  focus. Per-path ID/text-length diagnostics and immediate failed-setup return remain.
  Separately, the V47 Full trace proves two completed Hide Selected cache refreshes before the
  reason assertion. The existing enumeration callback now records the first two-row/z.txt
  transition and requires GenericSuccessor/no selection there, preserving final focus, selection
  and non-resurrection checks. No product focus policy was weakened. S3's first loss boundary
  remains unresolved. No predecessor clusters ran after the V49 failure. Next: freeze/build V50,
  verify path detail and repeated predecessor clusters, then canonical Commands and remaining gates.
- **October 1 V48 build correction / V49 next:** V47 Full/Fresh is terminal **FAILED**:
  **2256 total / 2204 passed / 6 failed / 46 skipped**, clean sandbox. Commands accounts for
  all six failures; FileOperations finishes **196/0/21**. Original reports and the post-Full
  receipt `861d9a9c02f9d6766917d75b2d09caf882accfdc9a483cc3888509f8d3b810e6`
  remain under `final-gates-drivers-20260930/release-v47-full`; raw run is retained on V:.
  V48's Debug build then failed before tests: two new attachment-count assertions omitted the
  required thread ID. Primary now passes `GetCurrentThreadId()`, fixes three aggregate-brace
  warnings and explicitly retires default/cancel button aliases when closing path detail.
  Failed build logs remain under `debug-v48-build`; no V48 native acceptance is claimed.
  Root verified all 82 logical / 81 physical V48 frozen inputs and reviewed the V49 driver delta.
  Next: freeze/build V49, run the focused path-detail journey, then the 24-case predecessor
  clusters twice and canonical Commands. Resolve remaining failures before final broad gates.
- **October 1 V47 broad Commands failures / V48 repairs:** the running Release Full/Fresh
  Commands suite finishes **917/6/2**, archived under run
  `20260930T210622Z-74096-a0f5e4dd120c45db8a4ba22dbae0c6a5`, Commands `2026-09-30_235732`.
  Root reviewed two cheaper agents' diagnoses against source. Primary now uses the existing bounded
  exact-action scroll helper for the Completed-group disclosure and retained-Move More button;
  the latter also waits for the popup's completion-collapse transition. History navigation explicitly
  uses Home and waits for index zero before its adjacent-arrow check; root corrected initial Down
  after cheaper-model review caught its hover-relative starting point. ViewerSpace observes the delivered
  mouse message plus synchronous paint before pumping unrelated native input; its failure diagnostic
  no longer reads a snapshot in a separate argument before the operation that populates it.
  The S3 discovery loss and generic hidden-current repair remain unexplained: new case-boundary
  plugin inventory and separate reason/selection diagnostics identify their first failing transition.
  No speculative product repair, timeout increase, skip or passing retry is claimed. These changes,
  plus the long-path detail below, are not natively qualified. V47 Full continues serially through
  FileOperations; let it terminate and retain its result before freezing V48. Next: build V48,
  predecessor-cluster repetitions and new path-detail journey, then canonical Commands once;
  resolve remaining failures before another Full/Fresh and final paired-resource gates.
  Final source-contract pass is **191/0/0**, `v48-harness-source-contracts-r3.xml`; the intermediate
  **190/1/0** was the existing lexical guard requiring the old initial-selection expression and is
  retained. Its updated guard requires Home plus index zero, preserving bounded observation. Root
  reviewed the V48 build driver and its exact **82 logical / 81 physical** frozen-input checks;
  ViewCommands is newly explicit. Path-detail API/lifetime review found no actionable source issue.
- **September 30 path-detail gap found during closeout:** source review confirmed M15 still required
  a selectable, scrollable visual detail for values longer than the bounded tooltip. V47 remains the
  frozen baseline; its Full/Fresh runner finishes independently, then releases native ownership.
  Do not run the obsolete broad V47 Debug gate. Primary now adds a consumer-owned, read-only DxUi
  multiline editor inside the existing popup (no new HWND or DxUi pin), with exact full/selected copy,
  Escape return, hidden-task shortcut suppression and live-row retirement. Focus and tooltip access
  remain available on the path row. New tests cover a >4Ki Unicode/UNC fixture, actual viewport movement,
  UIA TextPattern, read-only input and 40 open/close cycles with host/resource receipts. These edits
  are not yet built or natively validated. EN/FR/cs-CZ/ja-JP/sk-SK resource contracts pass **9/0/0**
  after completing satellite parity; original 8/1/0 evidence remains retained. Root reviewed the
  cheaper agents' source/test/gallery work, corrected unsupported UIA-attribute checks and the
  exact source/folder route, and formatted with clang-format 22.1.3. The V47 gallery remains a prior capture; the next
  candidate must add detail captures and repeat affected native gates before final acceptance.
  Next: finish source/test review, freeze the new candidate after V47 Full ends, then build and run
  its focused Commands path journey, gallery, protected DPI/input, Debug FileOps, Release Full and
  paired resource gates. No required spec behavior was waived to close I26.
- **September 30 V47 gallery published / Release Full running:** Release solution build passes
  with **zero warnings/errors**, receipt
  `a4f651898f8ac0eb2ec601bc4ab10e8163d87c65beb0ed88641d55725d6248f8`.
  The Full/Fresh runner is active under `final-gates-drivers-20260930/release-v47-full`, using V:.
  Root independently verified the current 560 PNG paths / 500 unique hashes against positive
  review ledgers (442 exact V24 hashes and 58 V43 hashes) and inspected three current originals.
  The [published current gallery](../WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-30-v47-final/README.md)
  retains static/interaction originals, exact receipts, manifests and visual-review evidence.
  This closes capture/review for these current scenarios, not the pending Full/Debug/resource gates.
  Two external publication-script type/array errors were corrected without changing image or test
  acceptance criteria; failed utility attempts remain in external logs. Root also checked all
  139 portable scan-cache evidence ZIP members against original hashes, sizes and manifest, then
  bound the future curation script to those exact reviewed ZIP/manifest hashes. No native test or
  benchmark runs concurrently with the delegated Full sequence. Next: review Full, broad Debug
  FileOps, post-Full receipt and exact dependency identities, then execute the paired resource gate.
- **September 30 V47 native interaction qualification / final gates running:** Debug solution build
  succeeds, receipt `1606402fa029ceab2d5ea775ab75869217b1583d54e908fb1f7106e698c611cc`.
  Three protected interaction runs pass **1/0/0 each**, clean V:, including actual Tab/reverse Tab,
  hover/press and the same-HWND 144→96→144 DPI round trip:
  `20260930T205252Z-73948-5c3a6bf63ef1405cb66eb5d36a40fa99`,
  `20260930T205423Z-97208-a81110aad2fc4681b13dad5dc6469452`, and
  `20260930T205521Z-97208-d8cbee34679443b7baca9d3717b53d16`.
  Native Tab and retained-focus counters prove delivered input and changed exact focus targets.
  The V46 failure remains retained and unreproduced; this is diagnostic coverage, not a product fix.
  Current static gallery passes **1/0/0**, 550 PNGs, clean V:, run
  `20260930T205754Z-65632-890748afb5ff42869426f729d7b421bc`; final hash review is pending.
  Sol has exclusive native execution ownership for V47 Release build, Release Full/Fresh and broad
  Debug FileOps; root reviews receipts, source and evidence, with no competing builds or tests.
  A78's refreshed exact scan fixture Release receipt is
  `905bccdbb1698cef960bcc1064bb0ec1f9da73568d8007711624a78d0081ea1b`; original baseline inputs
  remain archived. Root reviewed the corrected scan-digest scope and CPU-delta checks in the external
  final comparator, and independently verified all 1,053 selected scan-cache metric lines against
  their nine original raw files. This evidence does not itself accept performance. Next: complete
  current gates, review every new gallery hash, run the serialized A78/B49 pair using post-Full
  receipt and reviewed dependency identities, publish evidence and close only the completed I26 plan.
- **September 30 V46 broad/static gates; V47 input diagnosis:** Debug FileOps passes **197/0/20**,
  clean V:, run `20260930T201340Z-98916-d48f41e9c21b4c3e973e2348c832bc98`; the 20 named skips
  remain coverage/environment limitations. Static gallery passes **1/0/0**, clean V:, run
  `20260930T203703Z-104344-ebe1584ff01f435ab32ab3200a71e2c4`. Protected interaction run
  `20260930T204018Z-42076-60d1f6174ea54fde821e6b90d09319a6` fails **0/1/0** at forward Tab:
  retained focus stays `FileOperations.Action.9.1.0`, with native and foreground focus both owned.
  This is not an ownership refusal. V41 passed the identical input route; source review cannot
  distinguish missing native delivery from navigation/reset with the old final-state-only trace.
  V47 adds test-only native Tab/focus-change counters, last two focus identities and pump diagnostics;
  no product focus repair, timeout relaxation or pass is claimed. A78's exact scan fixture refresh
  and full Release rebuild succeeded; evidence is `final-gates-drivers-20260930/a78-scan-v46-refresh`.
  Review found the old integrated comparator omitted scan phases and only required equal fixture
  hashes. Separate V46 draft pair/comparator now require approved `90f7301c...ce740` fixture bytes
  and cover all scan phases; root review/execution remains pending. The scan-cache comparison packet
  is external and descriptive, not acceptance. Next: V47 diagnostic build/run, resolve Tab,
  final protected/gallery/Release Full gates, reviewed resource pair, publication and closeout.
- **September 30 V46 interaction verification:** full-solution Debug builds with zero warnings/errors,
  receipt `5bdbc45634f49e329480a918f09bc992c40a4b89f3eab6c38f26f91a2a5e3410`. Only route diagnostics
  and the checkpoint differ from V45; product code is unchanged. Throughput alone passes **1/0/0**
  (`20260930T201012Z-74620-3513ad806650430a944a4b80cb079984`), and the original four-case sequence
  passes **4/0/0** (`20260930T201157Z-60300-50c1d2f2170746f9b7edddf9ce009de1`), both clean V:.
  Trace proves Invoke/Space/pointer each change the exact task state with owned native/hosted focus.
  V45's earlier failure remains retained as unreproduced; no product fix or timeout relaxation is
  claimed. The broad Debug FileOps gate is running under `final-gates-drivers-20260930/debug-v46-gates`.
  V45's 550-PNG/490-hash static set is fully reviewed: 489 hashes match V43's reconciled review,
  and root inspected the one new legacy technical Issues capture; ledger is `v45-static-hash-review`.
  V46 still needs its own final static/protected run receipts. The A78 baseline must receive only
  the identical updated integrated fixture, a new supplemental manifest and a new Release receipt
  before the final pair; original baseline manifests/receipts remain untouched. External curation
  review found and repaired a missing post-Full manifest-hash field in the future Full wrapper;
  this is evidence wiring only, not a test pass. Next: complete Debug gates, A78 fixture refresh,
  Release/Full, paired resources, publication and indexed closeout.
- **September 30 V45 protected journeys / scan cache qualification:** Debug builds with zero
  warnings/errors, receipt `c410254e39ce21c532d6612bf5e77edefdce1d6e23b3b10ba80b18a972aefc8b`.
  Clipboard preservation succeeds without clearing or replacing unreadable data. Protected run
  `20260930T195720Z-62984-c4d48bfc9e7d4614a5968bbb584692f0` is **3/1/0**, clean V: sandbox:
  qualified/different Issues subjects and retained source/mixed journeys pass; throughput disclosure
  fails its generic UIA/Space/pointer state-transition assertion. This is an actionable test/product
  diagnosis, not another clipboard or ownership blocker. Primary adds route-specific state/focus
  trace before the next V46 run; no guessed product fix or relaxed timeout was applied.
  Three V45 scan measurements pass **1/0/0 each**, stable AC/power/displays, matching fixture and
  dimensions/digests. Root-reviewed comparisons against V42d and V43 live under
  `final-gates-drivers-20260930/scan-perf-v42d-v45-comparison` and `scan-perf-v43-v45-comparison`.
  Median UI-thread CPU across 80 redraws drops from V43's 203125/218750/218750/265625 us to
  125000/171875/187500/171875 us (Compare/Case/Attributes/File List). Original V42d medians are
  140625/140625/203125/203125 us; the Case sample remains higher within overlapping coarse Debug
  ranges. Presentation-paced wall time is near 8.33 ms, not an FPS/CPU non-regression proof.
  Surface/cache identities and bounded release assertions remain required; full Release/Full,
  final native gallery/visual review and integrated library pair are still open. Sol verified V45's
  failed-layout fallback and bounded cache ownership; root checked those findings against source.
  V45 static gallery passes **1/0/0**, clean V:, run
  `20260930T200207Z-79996-1a130c29007340e2a0eea5b97d4e0b5b`; new cache reuse/absence assertions pass.
  Next: freeze/build V46's route diagnostics, resolve the disclosure
  failure, then final Debug FileOps/protected capture, Release/Full and final resource/curation gates.

- **September 30 V44 review correction / V45 next:** V44 cache code builds with zero warnings/errors,
  receipt `84f84e5db20c9d61ed61025b3bcdff2c2ecb9b0e3f81335db2a2aa8cd0b29375`; no V44 native gate is
  credited. The cheaper-model cache review found a layout-allocation failure could return before
  advancing the scan row. Primary now keeps the already measured DrawText fallback, accessibility
  descriptor and row height, leaves the failed cache uncommitted, and retries on the next paint.
  The gallery also requires zero retained scan layouts for its absent-row scenarios. V45 will freeze
  these two corrections; baseline fixture remains unchanged. All 58 V43 new images have now been
  reviewed by Sol; root independently checked six originals across scan/placeholder/confirmation
  states. Final acceptance still requires exact final-image hashes and native current-code gates.
- **September 30 V43 scan correctness / CPU optimization:** Debug builds with zero warnings/errors,
  receipt `4c44f7e0e3b39ca9644af9e1d600b3a3a9d1a48f76c98039123a1e85b069d6e3`.
  All three candidate scan measurements pass **1/0/0**, clean V:, with matching fixture/dimensions/digest.
  Root-reviewed comparison is `final-gates-drivers-20260930/scan-perf-v42d-v43-comparison`:
  redraw medians stay near presentation-paced 8.33 ms, but median UI-thread CPU across 80 redraws
  rises from 140625 to 203125 us (Compare), 140625 to 218750 us (Case), and 203125 to 265625 us
  (File List). These are quantized Debug CPU samples, not FPS or a performance pass. Logical surfaces,
  brush/format counts and host attachments stay equal; process memory is noisy and not a live-heap
  attribution. V43 static native gallery passes **1/0/0**, clean V:,
  `20260930T194116Z-95964-42465926dcc347f0a84b802f7c14f6ec`; all 15-present/8-absent scan witnesses
  per physical monitor pass. Its 550 PNG/490-hash set has 58 new hashes awaiting full visual review.
  Primary now retains one measured scan layout per current snapshot slot, reuses unchanged paints,
  and releases unpainted/removed/theme/DPI entries. The gallery asserts reuse and bounded retention.
  This V44 optimization remains uncompiled/unaccepted; its performance fixture stays identical to V43.
  The older failed geometry runs and V43 CPU cost remain retained. Root also corrected the external
  comparator's per-repetition resource cardinality, frozen-fixture binding and culture-dependent
  date parse, and the curation draft's Debug Suite result filename; none supplies a product pass.
- **September 30 V42d baseline qualified; V43 compiling:** unchanged V41 product plus the final
  scan fixture builds with zero warnings/errors, receipt
  `0a14baf91c7b033de40d5e05c7d7ec630c98f41d460738a2983e38edae54e042`.
  Three fresh-process measurements pass **1/0/0 each**, clean V: sandbox, stable AC/power/display
  conditions: `20260930T192903Z-11768-5afec6aed9574da09c9aef989939b796`,
  `20260930T193001Z-11768-4442ccb5f45d4a83855ce2918b798273`,
  `20260930T193102Z-11768-efc0a68f515d4cac96b6ac43eabcaeeb`.
  Source fixtures are retained beside conditions; V43 normalizes twelve LF-only lines to CRLF,
  and root verified exact text equality after removing CR characters, with no other byte changes.
  The comparator checks both physical hashes and this exact line-ending-only relation.
  V43 freezes the scan product repair, diagnostics, gallery witnesses and normative docs (81/80),
  with Debug output under `final-gates-drivers-20260930/debug-v43-build`. Two Sol reviews found
  no actionable scan correctness/focus/Unicode issue; additional text-measurement cost remains
  pending the candidate pair. The external closeout ledger maps 21 unchecked parents to current
  evidence and remaining acceptance; it does not close any gate. Next: V43 build, paired scan
  measurements, final Debug FileOps/static/native checks and remaining Release/Full/resource gates.
- **September 30 V42c diagnostic / V42d settled baseline:** V42c builds with zero warnings/errors,
  receipt `ad7f7303e7248f92535321c11ec8dc3ce5c04ac6d9449a416d6cf3b7d87d0455`.
  Run `20260930T191825Z-41936-751e158b22fc4a1f955237a362b69ddb` fails **0/1/0**, clean V:,
  with the same 720x1080 -> 720x480 surface change. The one-time model reset did not settle
  geometry: the creation-time throughput-size preservation flag suppresses the first paint's
  natural resize. V42d refreshes the identical model on each bounded warm paint, exercising
  the same update route as changing samples before recording the initial surface. It retains
  the strict surface/cache equality gate. V42d is frozen 81/80 at `interaction-v42d-20260930`,
  with only the common measurement fixture changed from V41 product. Baseline build/evidence
  are `debug-v42d-scan-baseline-build` / `scan-perf-v42d-baseline` under the final driver root.
  V42b/V42c failed runs remain diagnostic, excluded from all accepted comparisons.
- **September 30 V42b baseline diagnostic / V42c warm-up correction:** the test-only
  baseline builds with zero warnings/errors, receipt
  `e1874517b8720e3f7ba48dbbffa8d12985df8e5b96a28552538cd21c4aa57939`.
  First measurement `20260930T191123Z-97360-7397e0298bfb4853a7354bc149c2819f` fails **0/1/0**
  with clean V: sandbox; before/after provenance remains valid and all raw metrics are retained.
  It records a real fixture geometry mismatch: Compare's warm surface is 720x1080, but first
  changing refresh settles the natural height to 720x480 (brush/format counts stay 3/2).
  No paired performance pass is inferred. Root now refreshes the identical synthetic model
  after the explicit initial resize and before all warm/timed samples, preserving the surface
  equality requirement. V42c freezes only that fixture correction on unchanged V41 product;
  baseline build is under `debug-v42c-scan-baseline-build`. Successful baseline measurements
  must precede copying the V43 product repair. The original V42/V42b records remain separate.
- **September 30 scan-status repair / paired fixture in progress:** primary now shares a
  measured localized scan row across Compare/Case/Attributes/File List, preserving full tooltip,
  UIA text and nonempty exact copy path. Gallery checks cover 15 present/eight absent scenarios
  per native monitor, long emoji/combining paths, terminal scan counts and a nonfocusable empty-path
  placeholder. Root review found the old public layout snapshot excludes informational tasks;
  primary's test-only projection now admits those exact task IDs. The common performance fixture
  instead reads existing host-wide observations through the owned diagnostic message, so the same
  fixture can measure the unchanged product. It adds four bounded informational workloads and
  separate timing/resource/dimension/digest metrics; existing Copy/Move/Issues metrics stay intact.
  V42 is V41 product plus the fixture only: it builds with five C5246 aggregate-brace warnings,
  receipt `f3bb4a0147d1ad0a54123ef29c25d8a250ee6683a21fd3e4352f00032f25dca3`; no native run.
  Root fixed those warnings and the diagnostic-wrapper limitation before measuring; V42b is the
  corrected test-only baseline, frozen 81/80 at `interaction-v42b-20260930`, currently building
  under `final-gates-drivers-20260930/debug-v42b-scan-baseline-build`. Candidate product remains
  V41 until baseline capture finishes. Next: three baseline measurements, freeze primary as V43,
  then V43 Debug/static/native measurements and remaining gates. No other owner's code is staged.
- **September 30 V41 verified keyboard / new scan-line visual defect:** Debug builds with
  zero warnings/errors, receipt `ce2412c49a9cb62968d7373bb21edf54516bd1a9ea5ac859e0c67ea84014adf2`.
  M02/M12 pass **2/0/0**, clean V: sandbox, run
  `20260930T183405Z-20248-7bdc8e9ef5c247b8af0945977b379d70`. UIA ExpandCollapse, Invoke,
  Space and pointer routes preserve both task pause states and conflict identities. Source
  contracts pass **191/0/0**. Static gallery native assertions pass **1/0/0**, run
  `20260930T183608Z-60264-6b9d57e8748d4060bf867aa8aa8bec3b`; visual acceptance remains FAILED.
  Of 548 PNGs / 488 unique hashes, 444 are identical to previously reviewed captures.
  Root and Sol reviewed the 44 new hashes, with independent root checks of delegate indices
  2/3/10. Nine captures show a hard-clipped `Analyse` path (four informational task kinds,
  including retained scan counts on terminal cards). Ledgers and reconciliation are under
  `final-gates-drivers-20260930/v41-static-hash-review`. Primary repair measures the localized
  scaffold and path, preserves full tooltip/UIA text, and adds native Unicode/fit/presence checks;
  it is not yet compiled or accepted. A paired informational redraw fixture is being prepared
  because the existing integrated Copy/Move/Issues benchmark does not exercise these rows.
  The first protected gallery attempt correctly refused foreground acquisition before input,
  run `20260930T183949Z-98916-547b17ad32984907836a51d3cf54f39f`. After fresh user availability,
  retry `20260930T184533Z-79304-9b95e520eb8142f6a0a4490ccee5a1dd` passes **1/0/0**, clean V:,
  with physical **144 -> 96 -> 144 DPI**, exact retained action ID, and native/hosted focus
  on each hop. The latest availability reply is consumed by that successful session.
  Next: retain scan baseline measurement, qualify the scan repair, then final Debug/FileOps,
  repaired clipboard-dependent Commands, Release/Full, galleries and A78/B49 resource pair.

- **September 30 V40 Debug diagnostic / V41 keyboard repair:** Debug builds with zero
  warnings/errors, receipt `ee77cc6e18857022c4f39c2a7b62dd7db84ed395eb7af92c547ec73a38cd2739`.
  Protected run `20260930T182019Z-6120-af6ca43e5e244504b2873fa77b08fb81` passes M02's
  tightened live Copy/Move options assertions; M12 fails exactly at Space after UIA
  ExpandCollapse and Invoke both successfully collapse and expand B. Foreground/native
  focus, retained control identity and both pending decisions remain valid. The popup's
  task-pause shortcut intercepts Space before the disclosure receives it. Primary now
  exempts TaskToggleCollapse, with per-route assertions that both pause states are unchanged
  and the normative keyboard contract clarified. This source repair awaits V41 native tests.
  The disk blocker is resolved: governed V:/RedSalamander.Perf is newly initialized with
  about 224 GiB available and a clean audit; use V: for remaining gates and both resource
  pair candidates. No legacy C:/D: data was changed. V36 raw migration completed after
  verifying all 49,757 files / 48,382,179,292 bytes: complete payloads are now under
  `C:/RedSalamander.Perf/evidence/i26-ui/raw-runs/20260930T155713Z-62344-a242b127316e45238c27cd6d362de331`.
  Intentional reparse fixtures remain at the original E: root with targets in the migration
  receipt; the former retention receipt is preserved and the current receipt links both.
  Full evidence remains failed, not reclassified. Four clipboard-dependent Commands still
  require readable clipboard contents; the pending user question has not been answered.
  Next: V41 build and protected M12, independent Debug FileOps/static gallery, then repaired
  Commands, final Release/Full, native mixed-DPI gallery and paired resources.
- **September 30 V39 protected diagnostic: 8 passed / 5 failed / 0 skipped**, clean E:
  sandbox, run `20260930T180722Z-71376-000cbf70557046e5a55f1ca15dc6580b` under
  `focused-commands-v39-diagnostic-20260930`. Release has zero warnings/errors, receipt
  `98aa17ac1b98b060cd3be139f7ebfff1a883ef1651cd18ce6cbe30ffc52933bc`; before/after provenance
  checks pass. The actual accepted Copy route records Supported, Verify On, enabled/focusable;
  native Move records Not applicable, Off, disabled/nonfocusable, Invoke and SetFocus
  `0x80040200`. Both operations validate immutable options and final bytes/disposition.
  Issues primary grid, remote case-sensitive inspection, both renderer-failure fallbacks,
  completed-group navigation, close/Escape behavior and the alert overlay also pass.
  Four cases stop safely at clipboard preservation (format 50383, `0x80030005`): qualified
  Issues, different-subject Issues, throughput and retained-source journeys. The user has
  been asked to copy short text; no clipboard mutation bypass is allowed. M12's new disclosure
  assertion fails without enough subcondition detail; primary now records route/state/focus/
  UIA HRESULT and both conflict identities without relaxing the assertion. A cheaper reviewed
  M02 strengthening also now requires a successful Invoke-pattern lookup before asserting
  Invoke rejection; V39 predates that final assertion strengthening and remains diagnostic.
  The 191 source-contract tests pass (`v39-source-contracts/pester.log`); its wrapper's
  all-stream redirection lost the returned object, so counts are explicitly reconstructed
  from the retained log in `results-from-log.json`, with no unchanged test rerun.
  V36 raw migration recovery was stopped before any receipt/deletion because provider-based
  ancestor checks made metadata enumeration excessively slow. The reviewed equivalent
  `File.GetAttributes` ancestor checks retain all guards; recovery2 is verifying existing
  C:/E: copies before removal. Next: V40 Debug diagnostic, clipboard-dependent cases after
  the user's reply, then remaining Debug/FileOps/gallery and final Release/Full/resource gates.

- **September 30 V37b/V38 diagnostic correction (supersedes the prior live-red claim):**
  Both diagnostic runs fail before reading Verify properties: local Copy is a managed
  transfer with Supported (or deferred-check) verification, while this native Move is
  Not applicable. The test incorrectly searched Copy for the Not applicable caption.
  `enabled=1`, `focusable=1` and `0x8000000A` were untouched test defaults, **not observed
  UIA values**. Neither run establishes the intended disabled-option red witness.
  V38's product correction is independently supported by source review; native acceptance
  remains open. The next test uses the captured prepared availability, stable option ID,
  property-read HRESULTs, enabled Copy cycling and disabled native-Move rejection, then
  checks the accepted Verify value in the immutable plans. V38 Release builds with zero
  warnings/errors, receipt `76cb17494f1ed489f52bcdafaffd946db1922c2ca23e3933d42b981e52da57c7`.
  Its C: diagnostic retains the separate 260 pre-existing noncanonical-directory audit
  failures. The full E: raw-file migration copied/hashed the payloads but stopped at its
  receipt byte-sum expression, before deleting any source file; safe verify-existing recovery
  is next. E: has only about 0.77 GiB free; the user has been asked for space or another fixed
  drive for Full and stable paired measurements. Independent UI/test repairs continue.

- **September 30 V37b test-only build / V38 product correction:** test-only Release passes with
  zero warnings/errors, receipt `1886efdc27f170134cb7abf885900163ae62b98bdc098537102a459ee3904980`.
  The actual Copy confirmation test fails **0/1/0** on the incorrect Verify-caption lookup;
  Links/Start/Bandwidth each change successfully. See the correction above for the
  distinction between initial diagnostic defaults and actual observed UIA properties.
  Run `20260930T174053Z-100700-968766526bf44b54bf5068df4a58952c` is under
  `disabled-confirmation-red-v37b-20260930`; before/after source/artifact checks pass.
  E: fell below the external launcher's conservative 64-GiB reserve, so this diagnostic
  used marked C:. Its audit separately reports 260 pre-existing noncanonical run directories;
  no unrelated directory was changed and this run is not a clean qualification pass.
  Root now applied the reviewed two-file AlertOverlay correction: fixed/absent choices
  retain accessible names/bounds but reject activation/focus, are skipped by Tab, and use
  read-only chrome. V38 will use E: for the 13 focused Commands routes and bounded synthetic
  graph case, with a 1-GiB reserve appropriate to those fixtures. This changes only that
  external lane's capacity preflight; no product/sandbox/performance gate is waived. Full's
  larger capacity reserve and final paired resource evidence remain outstanding.

- **September 30 V37 compile correction / raw retention:** V37's test-only Release build
  fails on two chrono literals inside the new nested generic UIA lambda and reports four
  implicitly-deleted-member warnings. Root changed those literals to explicit milliseconds
  and declared the shared state noncopyable/nonmovable; V37b is the corrected test-only
  candidate, with original V37 logs retained. No runtime red witness has been obtained yet.
  The supplemental archive helper now moves a completed run on its own volume outside the
  automatic stale-run sweep, preserving all raw logs, metrics and intentional link fixtures.
  It copies only supplemental reports/captures, recording omitted raw categories and
  SHA-256 checking every copied file. V36's complete raw record is now
  `E:/RedSalamander.Perf/evidence/i26-ui/raw-runs/20260930T155713Z-62344-a242b127316e45238c27cd6d362de331`.
  The corrected supplemental copy contains 107 files / 16,354,200 bytes; both final result
  files match the independently retained reports. The old partial copy remains separately
  named `runner-run-partial`, and the previous helper is preserved for audit.

- **September 30 V36 Full terminal evidence:** run
  `20260930T155713Z-62344-a242b127316e45238c27cd6d362de331` finishes
  **2201 passed / 9 failed / 46 skipped**, with a clean sandbox audit. Its authoritative
  Full/Fresh repository and change verdicts remain **FAILED**; the two unsuccessful entries
  are the three source-contract failures and six Commands failures described below.
  FileOperations itself passes **196/0/21** (`2026-09-30_191652`): the extra Release-only
  skip is the Debug-export scheduler-shutdown proof, alongside 18 absent remote-profile
  checks and two explicitly delegated 7z checks. No other native failure appeared.
  Root-reviewed direct disabled-option tests are now in the primary gallery layout fixture
  (zero/one-choice discovery, no hover/input/focus mutation, enabled cycling and Tab order).
  The product fix remains unapplied pending the live M02 red witness. V37 test-only freeze,
  Release build and 12-case fixture diagnostics are next; V38 drivers are prepared for the
  subsequent product fix and 13-case focused qualification. All application binaries and
  screenshots still run only through the protected harness; independent primary app PID
  99420 is preserved. Full's original reports remain at the exact E: run root. The supplemental
  recursive copy stopped at the intentionally broken Compare work link; its partial copy is
  retained and is not claimed complete. Root separately archived the final reports and
  post-run receipt `887f4ff8b8208d44f625c935633b322b99d3fda36e0e7ef9fb647e1acc0345eb`,
  and `Assert-I26Candidate` passed its 81-row source/artifact check. These receipts are under
  `final-gates-drivers-20260930/release-full-fresh-v36`. Supplemental evidence copying must
  omit seeded work directories/reparse points while preserving original reports and captures.

- **September 30 V36 Full Commands / next regression:** Commands finishes **917/6/2**;
  its two explicit ViewerSpace opt-in skips and all six failures remain in the original
  `2026-09-30_190026` archive. Full now continues with FileOperations. Root confirmed the
  minimum-Issues failure measured model row zero after selecting a different retained task;
  its actual 144-DPI capture has a complete 132-pixel selected row. The primary assertion now
  checks the exact selected task and its existing visible-row bounds, retaining the 88-DIP
  requirement. The S3 fixture now enables/restores its provider before the export loader,
  which correctly refuses disabled entries; load failures also record their HRESULT.
  The two injected-renderer fixtures now explicitly expect the native fallback during the
  metadata pause; normal fixtures retain all loading/UIA checks and cannot bypass them on
  unexpected host failure. The result-badge fixture exposes its header before checking the
  first collapsed frame. Completed Open source now reports invocation, exact provider/folder,
  item focus/visibility and filter state, then stops immediately on failure; its cause remains
  unknown. These corrections pass 191/0/0 source-contract checks in
  `final-gates-drivers-20260930/v37-source-contracts-fixture-corrections`; native validation is pending.
  The reviewed M02 test now drives actual local Copy/Move confirmation controls, accepted
  immutable options and file bytes, with bounded UIA/modal workers and owned-task cleanup.
  That review exposed a real confirmation defect: single-choice Verify is advertised as
  enabled/focusable, contrary to the disabled Not applicable contract. V37 will first retain
  a failing UIA enabled/focus/Invoke witness against unchanged production code. A bounded
  Astra review has prepared an external input/UIA/rendering correction, not yet applied.
  M12/M23 and these primary test repairs are not yet natively qualified. Final candidate
  version, Full/Debug/gallery/resource receipts and documentation closeout remain pending.

- **September 30 V36 Full findings / V37 preparation:** Fresh Full remains running against
  unchanged V36 candidate files. The toolchain entry passes 12/0/0; the main tooling entry
  reports **812/3/0**. Root is checking the three source-contract failures against actual
  queue cancellation, safe default/focus policy and activation guards. The completed-card
  Issues test had a real missing immediate return after activation refusal; the primary
  source now stops before its pane-wait loop. The source-contract repair retains the actual
  queued-card UIA Cancel and stale-provider checks, no-start/zero-byte cancellation outcome,
  safe nonactivating RecycleFailed default focus, and immediate stops for every audited
  activation assertion. Root's final focused Pester run passes **191/0/0** under
  `final-gates-drivers-20260930/v37-source-contracts-final`; original failed attempts remain.
  CompareDirectories finishes **249/0/23**;
  the explicit skipped cases remain in its original archive. Commands is running next.
  The reviewed V35 focused packet is now retained under
  `Specs/TestRuns/4cb089111a23/Validation/2026-09-30_i26-v35-focused` (38 files, 2,967,008 bytes);
  exact source-part membership/hash verification and canonical archive validation pass.
  Root checked the cheaper M01–M32 audit and is filling three bounded route gaps together:
  live confirmation option changes reaching typed admission, real pending-card disclosure
  input parity, and actual Issues Copy Detail invocation/readback. The last two primary
  test extensions are applied and reviewed but not yet natively qualified. Clipboard
  testing retains every advertised format, checks generation immediately before invocation
  and restores only its owned publication. Card checks use live A/B prompt identities before
  the separate synthetic long-text reflow. V33's unexplained native-focus failure remains
  retained; later diagnostic-only passes do not establish a cause or a production fix.
  Do not copy primary changes into the running V36 candidate or claim the new tests passed.

- **September 30 V36 final-gate run:** the exact Done-path admission, sibling rejection and
  tooling documentation are prepared without moving this plan. Root corrected the delegate's
  incoming-link depth calculation; direct retained asset links will need `../../Done/`, and
  nested implementation galleries `../../../../Done/`. The four focused tooling suites pass
  **82/0/0** and both inventories have zero blocking findings. V36 freezes **81 logical/80
  physical inputs**; the full Release build passes with zero warnings/errors, receipt
  `b1586d7b5c61703cf5ac3402a2eb555295bdf180752f3a74952fbeb11f787e27`.
  Fresh Full is running as `20260930T155713Z-62344-a242b127316e45238c27cd6d362de331` under
  `final-gates-drivers-20260930/release-full-fresh-v36`; its build-toolchain entry passes 12/0/0
  before the governed restoration rebuild. This is not yet a Full pass.
  Exact-docs-head DxUi checks pass skills/specs/dependencies/format at clean
  `c5dccff08bb00c2668a72f337c0a7eaf3446f53f`, and its native-provider packet verifies all 475
  logical files/348 unique payloads. No library code/tests changed after qualified pin49.
  The remaining interaction audit found a concrete M23 gap: complete conflict Copy Detail
  payloads are asserted, but the actual button-to-clipboard route lacks readback coverage.
  A bounded external test draft is being reviewed while V36 remains frozen. Do not claim all
  M01–M32 journeys passed from the inventory table. P0 baseline audit confirms programmatic
  checkbox/French-fit red witnesses and the speed defect's 144-DPI numeric red plus 96-DPI
  visual baseline; current native-96 geometry is green, not a claimed native-96 numeric red.

- **September 30 V35 focused gates passed:** full test-enabled x64 Release builds with zero
  warnings/errors, receipt `31076b5d632bbc4d0bbf2ab196938ea84a3c0b9c24c6dfdfe2d56fc48ea8eb80`.
  The corrected queue packet passes **3/0/0**, clean sandbox, run
  `20260930T154408Z-87960-c93e1973d1314f52809041562e859f93`, with the conservative provider
  interlock preserved and the completed-neighbor replacement bounded to one. The protected
  Commands packet passes **3/0/0**, clean sandbox, run
  `20260930T154548Z-60668-d55b3078a2cd4dbdae26c48eb509f077`: actual throughput ExpandCollapse,
  persistence/default omission/idempotence, all sixteen inspection/return routes and the
  strengthened two-pending-task case. Exact frozen inputs and unchanged build receipts are
  retained in `phase5-cancel-slot-v35-20260930` and `focused-commands-v35-20260930`.
  Remaining: final Fresh Full Release, current Debug FileOps/static/protected gallery,
  image review/publication, paired integrated resources, normative reconciliation and Done/index
  closeout. Prepare the exact Done-path tooling admission before the final source freeze; admission
  does not mark the plan complete. Original failed packets remain available.

- **September 30 V34 focused qualification:** full test-enabled x64 Release builds with zero
  warnings/errors, receipt `ee6b9a0a1a8f3c3c3daee44bdda3da624bf5ef04a304298aa43c770efb76a0b6`.
  The brief protected inspection packet passes **2/0/0**, clean sandbox, run
  `20260930T152012Z-86960-2b0fbf62890f481c8101c1d2ee230378`. All sixteen UIA/Space/Enter/pointer
  source/destination routes in both panes pass with unchanged input timestamps and correct native
  focus immediately after View and after pumping/redraw. The two-pending-task case also passes,
  including exact narrow/wide geometry, safe fallback, cross-task Escape and returning to A after
  resolving B. Evidence: `brief-inspection-v34-diagnostic-20260930`. This diagnostic-only change
  does not establish the cause of V33's return failure; original failures remain retained. The
  exact queue packet is **2/1/0**, clean sandbox, run
  `20260930T152154Z-88592-f465e5c0a0dd415bb7f01c4d9fef414c`: the attempted V34 fixture correction
  cannot admit B while A holds the dummy provider's conservative mutation interlock. This rejects
  the earlier read/read-only interpretation. Start now must not bypass that interlock. The next
  fixture restores explicit cancellation of A and, if reordered C legitimately completes before B
  enters, creates one replacement queued neighbor under throttled B for the unchanged stale-control
  safety assertions. Root's review of the cheaper audit also covers the interval between C releasing
  its interlock and publishing completion: the fixture waits for that publication and the real
  queued replacement row, permits at most one replacement, and explicitly rejects concurrent
  conservative mutation admission. The spec audit found an application-level ExpandCollapse
  pattern witness gap. Root reviewed and applied the bounded Sol draft to the existing throughput
  case, with bounded message pumping, exact-provider reacquisition, persistence/default omission
  and idempotence checks. V35 is frozen at 78 logical/77 physical inputs and its full Release build
  is running; no production behavior changed after V34. Current normative text now distinguishes
  implemented destination admission, focused speed/native-DPI evidence and the V34 two-task pass
  from remaining final gates. Fresh Full, final Debug/gallery, paired resources and normative
  closeout remain required.

- **September 30 V33 qualification:** full x64 Release passes with zero warnings/errors,
  receipt `d0feb2ba7c33ae64a27bd84a967a9eadacec7dee5937a658ae414b7f16e4404a`.
  The repeated inspection attempt stops at **0/1/5** after an input timestamp changes during a
  handoff; the product correctly declines the protected focus transfer. Original failure is
  retained under `repeated-inspection-v33-20260930`. Following the user's explicit idle-session
  confirmation, the shorter one-repeat attempt is **0/1/1**, clean sandbox,
  `20260930T144341Z-31548-72b54c43fba0470288be529d338af7c4`. Both backend owner-navigation
  witnesses and all sixteen explicit source/destination handoffs pass. The final pointer/right/
  source return preserves the same popup, hosted focus, scope, scroll and engine actions but does
  not regain native foreground/focus. Return-side input witnesses are absent, so its cause is
  not yet established; neither failure is recategorized as passed. Evidence is retained under
  `brief-inspection-v33-idle-20260930`. The two-pending-task case remains fail-fast skipped.
  The frozen V33 full Debug build passes with zero warnings/errors, receipt
  `6b26d3c6e31f597edaebf42df50ee7f8271a95baa21c17e33721e82589be212e`. Root and a read-only Sol
  review agree the return trace needs immediate-command versus later-pump/native-input witnesses;
  these diagnostics are in the primary test source for V34, preserving guards and assertions.
  Root reviewed the cheaper delegate's grouped Debug driver and verified the actual planner marks
  broad FileOps/static capture nonactivating and the exact interaction gallery protected. V33
  broad Debug FileOps finishes **190/1/26**, clean sandbox, run
  `20260930T144932Z-24088-981b6fbb0eeb4f9f8ef6520dbe69908c`. Its sole failure,
  `Phase5_DiscoveryCancelReleasesSlot`, occurs before either
  stale-provider invocation at the later started-task check. Source review identifies a fixture race:
  after releasing B's scheduling gate, the test cancels A without waiting for B's actual operation
  entry. Reordered C can legitimately enter while the active count is zero. The primary V34 fixture
  waits for B to enter while A remains admitted, preserves the unchanged engine guards, and logs
  before/after neighbor entry/order plus each invocation. Root checked the read-only Sol review;
  it confirms the scheduling race and bounded fixture correction. V33's 26 skips comprise six
  cases not reached after failure, eighteen missing-profile remote cases and two intentional 7z
  delegations; none are counted as passes. V34 freezes the corrected queue fixture, return
  diagnostics and 24-MiB fixture comment at 78 logical/77 physical inputs; its Release build is
  running. The independent Debug/static/protected driver draft prevents one failed gate from
  suppressing unrelated evidence. Next: qualify Debug
  background checks, resolve the return failure with evidence, run the two-task correction, then
  Fresh Full, final captures/hash review, paired resources and normative closeout.

- **September 30 V32 focused results / V33 diagnosis:** full x64 Release V32 passes with zero
  warnings/errors, receipt `0ad8fa21c9c89d6f445de27d6e6cef487ab8b2c763bc77aa0a58abb875c97ebd`.
  The nine-case run is **4/1/4**, clean sandbox,
  `20260930T142353Z-64272-315971f31f0e4f52a5c23c34098c87ed`: progress, throughput, loading/scope
  and comparison pass; the previously passing owner-navigation aggregate intermittently fails
  before explicit inspection handoff. Its path/selection/OS-focus subfacts were not logged, so no
  cause is claimed. The primary now records those facts, foreground PIDs, input ticks and the
  completion focus witness; the assertion is unchanged. Read-only Sol review confirms selection
  finishes synchronously, but the fixture pump can drain more messages after the callback.
  A separate unexecuted-case run is **0/1/3**, clean sandbox,
  `20260930T142615Z-66592-c4ee4de7e2414e75ab8e6e074ffbce7e`: exact 480/640-DIP frame-adjusted sizes,
  long-French two-card scrolling/reflow, nonactivating safe fallback, cross-task Escape isolation
  and A Source return finish; final A Destination return after resolving B fails a second stale
  pre-invocation default/cancel comparison. Primary now checks A's exact published engine actions
  and native return focus, as already done for A Source, rather than B's old bindings.
  The remaining cases pass **3/0/0**, clean sandbox,
  `20260930T142744Z-79336-a1821ec9ac994216b503f34332493349`: real nested case-sensitive loopback
  S3 inspection in both panes, live French short-body/locale speed reflow, and actual protected
  Tab/Shift+Tab, invalid Escape, changed-value Enter. Its separate graph-history group also passes.
  Original failures are retained; passing later cases do not clear them. V33 freezes the corrected
  assertion and diagnostics at 78 logical/77 physical inputs; Release build is running. Next:
  repeat both inspection cases three times, diagnose any failure, then Fresh Full and final Debug/
  gallery/resource gates. The external gallery-hash draft has been root-reviewed and selfchecked:
  all 558 V24 PNG paths resolve to 498 reviewed hashes, with zero new hashes; final images remain
  uncaptured/unreviewed and the script never marks a new hash reviewed.

- **September 30 V31 qualification / V32 fixture correction:** full x64 Release V31 builds with
  **zero warnings/errors**, receipt `7647b6619e55593d23ce1fed1f22a245a302929b3b0fb97dee28b03c18b9b2ae`.
  The changed Commands run is **5/1/1**, clean sandbox,
  `20260930T135939Z-82900-39e1174d1d4c4c8682c0a4e263b6a557`. Progress, collapsed rate/qualified ETA,
  live loading/repeat-scope UIA, comparison rows and every explicit inspection-return route pass.
  In particular, UIA, Space, Enter and pointer handoffs restore native foreground/focus in both panes.
  The two-task case stops at its newly exact geometry check: it requested an outer width of 480 DIP,
  below the actual minimum of 480 DIP client plus frame. The primary fixture now converts client
  geometry with AdjustWindowRectExForDpi and reports measured dimensions separately from activation.
  Remote inspection is fail-fast skipped, not passed; its primary fixture also waits for metadata
  completion and published actions before invoking UIA. The independent storage rerun passes
  **18/0/0**, clean sandbox, `20260930T140522Z-13444-d442a4ab2d7b47b08e1d7bb9884a238a`.
  The 187-bucket graph-history case plus native setup/cleanup passes **3/0/0**, clean sandbox,
  `20260930T140932Z-97800-1172cd89d62b4fb6a8705fb79fe526f2`. Its external validator originally
  expected one case; reviewed-result-validation.json checks all three exact names against retained
  results and current frozen inputs, preserving the original validator error without a native retry.
  Spec inventory reports zero blocking findings. These primary corrections are not in frozen V31.
  Root reviewed and applied the Sol draft in the existing live speed-dialog case: UI-thread popup
  activation, modal foreground/PID/native-focus verification, Tab/Shift+Tab with UIA witnesses,
  invalid Escape preserving 128 KiB/s, and valid Enter applying 256 KiB/s before restoring the
  fixture's 128 KiB/s premise. Review corrected thread ownership, lambda captures, null-window
  checks and the weak same-value Enter assertion. V32 is frozen at 78 logical/77 physical inputs;
  its Release build is running. Next: exact focused cases, then Fresh Full and final Debug/gallery/
  resource evidence. No new production API or desktop-control path was introduced by the fixture.

- **September 30 V29 diagnostic qualification / V30 corrections:** V29 links in x64 Release,
  receipt `37c8bfdc18c3ab7cede0d590b9d5696d565078a1544f6a8be01563a96f3fca2a`, but has two C5219
  test-coordinate conversion warnings. Its native results are diagnostic only; V30 fixes the
  conversions and requires a warning-free rebuild. The five-case run is **3/1/1**, clean sandbox,
  `20260930T132852Z-68644-53c5c63fd87144ba83e8e399ba1d51e4`: progress contracts, actual blocked
  loading/scope UIA hierarchy and comparison rows pass; inspection return fails its native-focus
  assertion, so the two-task case is skipped. The separate two-task diagnostic is **0/1/0**,
  `20260930T133225Z-20364-d5ec72a4ba5243d2a95606a50dfe616b`: long-French reflow and safe background
  fallback assertions finish, then A's original Source-inspection return fails. Both original
  reports remain under `diagnostic-focus-loading-v29-20260930` and
  `diagnostic-two-pending-v29-20260930`; neither is recategorized as passed.
  Storage's eighteen actual cases all pass in `20260930T133829Z-73440-802576091bc8412da4b156e4e1c2bf29`,
  but the runner rejects a stale inventory name and adds `selftest_result_coverage`, yielding **18/1/0**.
  V30 updates the renamed case in the native inventory; final runner acceptance remains required.
  V30 removes native focus transfer from the two remaining presentation-only default/inspection
  restore calls. Explicit View activation still owns the unchanged foreground/input witness.
  The reflow test now checks both axes, visible fallback bounds, exact resize and nonzero scroll.
  Its subsequent inspection must bind A's published Enter/Escape actions: the deliberately focused
  global Options baseline has no implicit decision while two tasks are pending, so comparing against
  those pre-invocation empty IDs was invalid. V30 checks exact current engine action IDs and native focus.
  Root reviewed and integrated bounded cheaper-model test drafts: actual loopback S3 nested
  case-sensitive inspection in both panes, exact collapsed rate/qualified ETA UIA content, and the
  existing four-stream graph sampler driven beyond its 180-bucket capacity. Review corrected raw
  versus qualified paths, provider/configuration cleanup ordering, inaccessible formatter usage,
  and stream totals so those fixtures preserve valid semantics. They remain unqualified until run.
  Agents do not edit repositories or run native workloads. V30's first build rejects the new fixture's
  lvalue passed directly to WIL scope_exit and warns about pointer-to-bool conversion. V30b supplies
  an explicit lambda and null comparison; the failed V30 log is retained. V30b is frozen but not
  built: a targeted review then confirmed a cross-task Escape defect. V31 keeps default/cancel
  unbound when focus belongs to a present task whose decision is not published, rather than
  assigning another task's sole actionable decision. The two-task witness checks those empty
  bindings while preserving both real prompt contexts. Next: freeze/build V31, qualify exact
  changed cases, then Fresh Full and final Debug/gallery/resource gates.

- **September 30 V28 qualification / V29 focus closeout:** full x64 Release V28 builds with
  zero warnings/errors, receipt `900c8bba6031917126791fa15f481628757edfa9865fcbb9b2db3178f3f19167`.
  The strengthened repeat-scope test passes **1/0/0**, clean sandbox, run
  `20260930T131400Z-86980-99e9239744334baa94fc540eed778ec3`; a model-originated toggle survives
  two redraws with matching model/paint/UIA state and no reciprocal toggle. Protected interaction
  passes **1/0/0**, run `20260930T131622Z-90640-eede4caa71894bd98d9f309555422625`, preserving
  native and hosted focus on the same action through physical **144 -> 96 -> 144 DPI**.
  Storage improves to **17/1/0**, run `20260930T131057Z-98120-8653cdb797c247ac9ef2f3ca11c8b652`;
  all final progress/fallback assertions pass. The sole residual checks an internal Debug string
  in foreground stdout; V29 reuses the structured completion predicate already passing in the
  companion query-failure case. No test run is silently recategorized as passed.
  Root checked cheaper-model P2/P3 ledgers and found a concrete focus-rebuild omission: a removed
  action had no explicit fallback, and restoring retained focus could transfer native focus.
  V29 preserves stable identity without native transfer, falls back to the same task disclosure
  or global Options, and adds a two-real-pending-task/long-French/reflow/background-focus witness.
  This new behavior is not yet qualified. Next: finish bounded loading/coverage review, freeze
  and build V29, exact storage/focus tests, then Fresh Full and final Debug/gallery/resource gates.

- **September 30 V27b search qualification / V28 final fixes:** Release V27b builds with
  zero warnings/errors, receipt `474fe5f9eb56ffbb21db5aadd892dd2fe4acedfc7aeb4f49602e3ea4b69adaa7`.
  Storage qualification improves to **16/2/0**, run
  `20260930T125917Z-79772-9297e287502842699219ce6c8247f0bf`; regex valid/case/rejection/recovery
  and the corrected foreground completion check pass. The two remaining failures identify
  unready legacy SQLite rows reaching the candidate callback before fallback, and throttled
  progress ending one unique item behind final stats. Root-reviewed V28 rejects unready roots
  before SQLite enumeration, flushes final qualified live progress, and removes a duplicate
  progress increment. It also executes a model-originated repeat-scope change through two
  redraws, requiring model/paint/UIA agreement without reciprocal toggles. These are unqualified
  until the next build and focused run. Gallery copy findings were synthetic-fixture mistakes:
  V27b uses the production count phrase and populates omitted current paths; ingress wording is
  unchanged. All 498 V24 image hashes and the 479-new/19-prior review union are root-verified in
  `final-gates-drivers-20260930/v24-root-hash-reconciliation.json`; V27b/V28 captures remain pending.
  Next: freeze/build V28, exact storage and scope/interaction qualification, Fresh Full, final
  Debug/captures and paired resource evidence. Preserve all failed receipts; no silent rebaseline.

- **September 30 V26 native focus qualification and V27 repairs:** frozen V26 builds x64
  Release with zero warnings/errors, receipt
  `9efbc265c3558f986aa1c1c1098e47882dba703fb9eb259623bbf7dc3ab6c5c6`.
  Protected interaction passes **1/0/0**, clean sandbox, run
  `20260930T124241Z-94068-0f902a216221465d98c945a55d30825d` in
  `release-interaction-v26-20260930/qualification.json`. Tab and Shift+Tab settle on distinct
  controls; the same popup/action retains both native and hosted focus through physical
  **144 -> 96 -> 144 DPI**. No availability reconfirmation is needed for already-authorized
  protected sessions. The eighteen storage regression cases finish **13/5/0**, clean sandbox,
  run `20260930T123508Z-92360-0f6d6b785c2d46889c7413b3db1dd28f` in
  `storage-qualification-v26-e-20260930/invocation.json`. Exact-current local fixtures,
  real stale/freshness guards, replay and warmup/multiclient checks now pass. Remaining failures
  expose missing compiled name regex in the broker/test caller, a synthetic broker progress
  frame that clears real fallback facts, and a stale stdout expectation for an internal Debug log.
  V27 repairs these with request-local validated regex lifetime, rejection/recovery/case tests,
  preservation of qualified progress, and a structured foreground-completion assertion.
  It also strengthens protected-input ownership and key-release cleanup. These edits are not
  yet built or qualified. Cheaper-model visual inspection covers all **498 unique V24 PNGs**;
  root spot-checks and SHA reconciliation remain part of acceptance. Two minor punctuation
  observations are under source review. Next: finish review, freeze/build V27, qualify changed
  storage/interaction routes, then Fresh Full and the equal-feature A78/B49 resource pair.
  No terminal Full or overall completion is claimed.

- **September 30 V25b placement and diagnostic qualification:** the first V25 build exposed
  a local-name collision and timeout conversion warning; both are fixed and the failed log retained.
  Frozen V25b builds x64 Release with zero warnings/errors, receipt
  `14226b662102e3d82ce251eafff5ad1626e231a1ccef11e14f9b6cf70b95376c`.
  The four placement/hide/recreate Commands pass **4/0/0** with a clean sandbox,
  run `20260930T121536Z-78960-55ce419961574474ad596091b15e1bab`, in
  `placement-native-v25b-20260930/qualification.json`. This covers actual Issues rectangle
  persistence and restoration of fixture settings, plus pure workspace/screen offset cases.
  The exact eleven storage rechecks still fail in `storage-recheck-v25b-e-20260930`;
  diagnostics prove the stored/live journal IDs match but live cursors advance after fixture writes.
  Production correctly rejects those stale cursors. Root-reviewed V26 test corrections supply
  controlled exact cursors for local positive-path tests and isolate synthetic service SQL-failure
  and prefilter fixtures; real stale-root and freshness-policy negatives remain unchanged.
  Release interaction `20260930T122101Z-65444-f6d9b5e7468242ac9287c58447eda62d`
  fails before Shift+Tab capture, **before any DPI hop**. Its initial wait could accept retained
  focus before the first Tab was processed. V26 requires an actual focus transition and records
  forward/reverse native and retained focus; this correction is not yet built or qualified.
  A cheaper-model review inspected another 80 distinct V24 original PNGs; root checked narrow
  and scrolled confirmation samples. Complete gallery inspection remains in progress.
  Next: finish the narrow storage-fixture corrections, freeze/build V26, run focused storage and
  interaction checks, then the required Fresh Full and equal-feature resource gates. No Full or
  native-DPI pass is inferred from the placement result.

- **September 30 V24 Debug and V25 corrections:** full Debug FileOps passes **197/0/20**,
  run `20260930T113256Z-89140-2b2e314f131e43bf98c72570d6e884b3`; skips are the two delegated
  7z cases and eighteen unavailable remote-credential cases. Static and protected galleries each
  pass **1/0/0**, runs `20260930T115425Z-89140-81b8da1a89cc494e9a02d90095632fad` and
  `20260930T115733Z-89140-8380f31e795a4125ac5df932b4f6f203`. All sandbox audits are clean;
  receipts are in `debug-fileops-gallery-v24-20260930`. The actual same-HWND focus TSV records
  144-to-96-to-144 and an unchanged action identity. A bounded cheaper-model visual review covers
  fifteen exact PNG hashes; root independently samples the filename-first conflict and target-conflict
  captures. This is not a claim that all 498 unique renders have been visually inspected.
  Root reviews and integrates V25 improvements: shared screen-coordinate placement capture for
  File Operations/Issues; real hide/reopen and destroy/recreate placement assertions; serialized
  restoration of settings after the fixture; native/retained focus witnesses and a separate observed
  DPI qualification receipt. V25 also records actual direct-SQLite freshness facts without changing
  the production decision, and replaces a fixed warmup delay with a bounded status wait.
  Frozen V25 contains 74 logical / 73 physical inputs; its Release build is in progress.
  These new changes have no native pass yet. Maximized cross-monitor placement with different
  work-area offsets remains a targeted qualification gap; no speculative legacy-record migration
  is added. Next: V25 build, exact storage diagnostics and placement/focus checks, then required
  Fresh Full and equal-feature resource gates. Existing product decisions and approvals stand.

- **September 30 V24 focused qualification:** frozen candidate V24 builds in x64 Release
  and Debug with **zero warnings and errors**. Receipts are respectively
  `d253b9af61b99c79e06c83b1df1caec2c586dd86c329e156dd03819225328608` and
  `3d5a430a6c9b07be4879f397974f2795cf5aec348d28acbe91a891e9ee2fefe4`.
  Release protected Commands passes **17/0/0** and queued-cancel FileOps passes **3/0/0**,
  with exact source/artifact checks and a clean sandbox, retained externally in
  `focused-native-v24-e-20260930`. This executes the new mixed Cancel-all, Conflict More,
  queued Cancel, disclosure, completed Issues, hide/reopen geometry and Issues selection/scroll
  assertions. The shortcut and Compare-window cases that failed during the full-disk run also pass.
  No separate running/paused-worker Cancel-all or mid-confirmation admission proof is claimed.
  E: now has adequate free space and its governed runner cleanup succeeded; subsequent native
  qualification uses E: with the explicitly marked C: alternate. C: as the primary test root
  is unsuitable for a clean audit because it contains 260 older noncanonical run directories
  owned by other work; those are preserved. The earlier C-root storage attempt is not acceptance.
  With ample E: capacity, the exact SQLite/service recheck still fails **0/11/0**, retained in
  `storage-recheck-v24-e-20260930/runner-run`; disk exhaustion does not explain this gate.
  Source review identifies an invalid currentness premise in the tests: journal availability
  alone does not prove exact stored/live journal-cursor equality after fixture and log writes.
  Production freshness remains unchanged while a deterministic correction is investigated.
  Full Debug FileOps and both deterministic galleries are running against V24; the native
  same-popup focus/DPI proof remains pending. Next: finish that sequence, close the explicit
  Issues-placement assertion gap, resolve the Full SQLite/service failures, then Fresh Full and
  equal-feature A78/B49 resource qualification. No terminal Full or overall completion is claimed.

- **September 30 V23 recovery and final interaction routes:** the protected Full V22b run
  was interrupted after its Commands process stopped advancing. The owned-process dump,
  traces, process identities and free-space receipt are retained externally under
  `final-gates-drivers-20260930/full-v22b-commands-stall`. E: had **zero free bytes**;
  the concurrent shortcut case recorded `STG_E_MEDIUMFULL`. This invalidates storage-dependent
  qualification; it does not prove the SQLite failures or Compare stall have the same cause.
  Only the verified runner and its owned Commands process were stopped. The independent
  application was preserved. All 18,784 artifact files (14.44 GB) were copied to C: and
  individually SHA-256 verified before any cleanup. No terminal Full pass exists.
  The initial recovery proposed C: with D: as alternate; V24 above supersedes that choice after
  inspecting those roots. Paired resource runs must use the same qualified root for both candidates.
  The long-path restore correction now passes focused Pester **39/39, zero skips** in
  `longpath-focused-v23b-20260930`. Its fixture pads the actual runner prefix so Git's
  discovery/config path stays below MAX_PATH while cloned template metadata and tracked
  receipt content exceed it; the earlier overlong-fixture failure remains retained.
  Root reviewed and integrated the cheaper-model mixed Cancel-all draft: decline preserves
  conflict/queued/completed/informational records, acceptance cancels the two live copies
  without changing completed/informational output or publishing queued bytes. Cleanup now
  drains the owned cohort before restoring Queue/Parallel even on early failure.
  M19 exercises the real More menu, Escape and Skip All with exact typed accepted/cached
  decisions. These native routes still require compilation and execution in frozen V23.
  No simultaneous separate running/paused worker or mid-confirmation admission pass is claimed.
  Governed exact-run cleanup partially removed scratch before failing on nonempty directories;
  direct deletion of the archived metric duplicates was rejected by automatic approval review
  (`blocked by policy`) and was not retried. C: and D: have adequate free space; no other E: data
  is removed. The archive also retains all 491 scratch files with individual hash verification.

- **September 30 final-gate failure investigation (V23 preparation):** Fresh Full V22b
  `20260930T100200Z-66244-4ff44f5ca2cd4a89bdf20972fc2e720d` passes the 12 build-toolchain
  cases and its full-solution re-attestation, but Tools Pester reports 813 passed / 1 failed.
  Git fails while creating clone metadata in the deeply nested cold-restore fixture, before
  repository-local `core.longpaths` can apply. Primary-source restore now supplies process-level
  `-c core.longpaths=true` as well as clone-local configuration. Its strengthened fixture uses a
  controlled Git metadata template and verifies actual restored metadata and tracked content
  beyond 260 characters, plus refusal to repair dirty sources. See the later focused pass above.
  Compare Directories reports 238 passed / 11 failed / 23 skipped; ten failures concern SQLite
  mirror/warmup/query routing and one concerns a busy service pipe. Preserve the original results;
  investigate before any Full pass claim. Commands subsequently stalled and the run was stopped
  with the disk-full evidence above. External evidence is `release-full-fresh-v22b-20260930`.
  Reviewed route-gap work is limited to actual queued Cancel, disclosure, Conflict More, completed
  Issues and exact focused-action continuity on the same popup across native monitor moves.
  No already-covered inspection/export path is treated as an absent assertion. Root owns all
  source edits and native execution; cheaper agents provide read-only audits or external drafts.
  Root integration now extends ordinary hide/reopen with exact geometry and explicit View-command
  focus (without a test-side focus repair), actual disclosure UIA/Space/pointer activation,
  completed-card Issues selection, and exact queued-card Cancel plus retired-provider replay.
  The queued Cancel draft was reviewed and hardened to reacquire task identities after UIA dispatch.
  The same-popup monitor gallery now records the focused action before/after each physical hop.
  These new routes are source-complete but not yet executed. Two additional Commands failures
  are stale source guards after helper extraction: the revised guards trace the destination-name
  validator and locked interlock-removal helper instead of requiring their former inline text.
  Current tooling/spec inventory checks report zero blocking findings; they are not native passes.
  Next: retain the terminal Full failure, qualify the restore correction, finish/review the narrow
  route drafts, freeze V23, then run the remaining Debug/gallery, Fresh Full and paired resource
  gates. V22 focused passes and library receipts remain valid only for their recorded scope.

- **September 30 V22 combined interaction qualification:** full x64 Release builds with zero
  warnings/errors, receipt `73a44b3810781ac65633b38729204a68ef7a022826d23d017cf578f2ae4f7fbe`.
  The protected ten-case sequence passes 10/0/0, run
  `20260930T094450Z-33212-62b0bcadf1c64622b40444b7686f17ed`, with a clean sandbox audit.
  This includes the corrected provider switch, two actual pending tasks, Export, Show log,
  Options denied-save recovery and running-worker speed/Pause/Resume journeys. Root visually
  checks the reopened and natively resized French error captures: both retain complete text
  and close glyphs without diagnostic replay. Phase5 destination-history keyboard selection,
  stale-provider rejection and discovery Skip also pass, runs
  `20260930T094631Z-42852-748f728347be4a649e4b41b139b0961a` and
  `20260930T094708Z-42852-88aaf02695374d6a9ec5dfdb50a99a92`.
  Receipts are retained in `protected-actions-v22-20260930` and `queue-controls-v22-20260930`,
  now copied into the [V22 repository packet](../../TestRuns/4cb089111a23/Validation/2026-09-30_interaction-v22/README.md).
  All 238 original ZIP entries are byte-verified. Pre-stage curation retains 144 entries and records
  94 omitted raw/duplicate identities; the complete original packet stays external. The curated
  packet passes the canonical 2 MiB/file and 5 MiB/run budgets. All 51 original PNGs have copied hashes and a visual
  review covering 36 distinct images. Root samples the delegated review, including missing-item
  feedback and exact selected-pane captures. UI/Settings normative status and Options failed-save
  semantics are reconciled; specification governance reports zero blocking findings including this
  new packet's links.
  The first Fresh Full Release attempt was aborted during its owned build-tooling deployment test:
  the external driver omitted the pinned vcpkg environment, causing a bundled-VS restore. Retain
  `release-full-fresh-v22-20260930` and its verified descendant-process stop receipt; unrelated apps
  were preserved. A matching full-solution `-Rebuild` is repairing the interrupted artifact epoch
  in `final-gates-drivers-20260930/release-v22-repair`. That clean rebuild passes with zero warnings/
  errors, receipt `4e59b27448c5fb77a5509de97a94e72b3d9ef75605440b8b06fb2b7e461e00f7`.
  Corrected Fresh Full is running in `release-full-fresh-v22b-20260930`. Its driver retains the
  runner's legitimate post-mutator re-attestation receipt, not require the initial receipt unchanged.
  Next: corrected Fresh Full, current Debug full FileOps and both galleries, reviewed
  A78/B49 equal-feature resource comparison, normative docs and complete after-capture review.
  Original failed V19–V21 evidence remains; focused passes do not substitute for these final gates.

- **September 30 V21 native proof and V22 fixture corrections:** Release builds without warnings,
  receipt `4524e59555639b1fc633ee2d2668862108ffb6f104e1d81615c8bc5aad8b0ff6`;
  207 focused Pester checks pass. Actual Export run
  `20260930T093203Z-58820-36c432ebe6b44776af86d095c3132cae` passes 1/0/0, clean disk audit.
  Root inspects its original `write-failure__fr-FR__144dpi.png`: the French error text and close
  glyph render fully without diagnostic replay. Combined actions retain 3 passes, 1 new-fixture
  failure and 6 unexecuted cases. Root/agent source review proves the dummy provider's actual
  short ID is `fk`, not `dummy`; V22 derives ID/root from plugin metadata instead of invented
  strings. The separate Phase5 run passes the original stale-provider reorder and Skip behavior,
  but its new keyboard-history witness searches a leaf instead of the intentionally qualified
  nonlocal label. V22 derives that label from the captured admission. Both original failures stay
  archived (`protected-actions-v21-20260930`, `queue-controls-v21-20260930`). No ownership failure
  occurred. V22 also captures reopened and natively resized export error overlays, covering first
  full-frame presentation after resource reset. Its 207 Pester checks pass; build/native pending.
  DxUi documentation/evidence closeout is committed and pushed (`c5dccff08bb00c2668a72f337c0a7eaf3446f53f`);
  both packets pass byte verification from a fresh Git archive, with explicit attributes preventing
  newline conversion. Consumer pin49 remains the qualified code revision; the later commit is docs only.

- **September 30 V21 and library handoff:** V20 full Release passes, receipt
  `f54bffcabe13bf68dda4afd8cd86026fe3e7940985059dd5e8ce62c4c8a29a5e`.
  Its Export run `20260930T091014Z-44796-97e26030a30e452db069058cb8a90fc9`
  correctly fails the original black WGC capture. A failure-only full replay yields 100,352
  nonblack pixels and the subsequent independent WGC capture passes its French glyph witness
  (74 matches). Root visually confirms readable French error text; the original failure remains.
  The first nonmodal paint can present only a region before initializing the entire swap chain.
  Root applies the existing `RenderInitialFrameForShow` API after first successful preparation
  and region application, retaining the post-region frame guard and ordinary partial paints.
  V21 freezes 67 logical / 66 physical inputs and is building; no fixed-runtime pass yet.
  Root reviews the Sol provider-switch draft, corrects missing provider cleanup, and adds exact
  plugin/context/path restoration before disabling dummy. The Luna destination opener/keyboard
  draft now uses actual retained UIA and bounded Home/Down/Enter. Root rejects its scenario58
  checkbox-absence assumption: replacement withholding and repeat-scope eligibility are separate.
  Scenario55 permanent deletion supplies the true policy-ineligible absence witness. The gallery
  now moves the same invalid French speed prompt across real monitor work areas and back,
  checking edit identity, selection, error, focus and footer bounds; execution remains pending.
  Primary `Dependencies/DxUi.lock.json` explicitly adopts qualified published
  `49c963f5751a6754e6b08e0362b96cfc7fc64414`; prior pin73 remains the rollback identity.
  DxUi's NativeProviderLifetime plan moves to Done with its exact CI, original failing/passing
  consumer journeys and retained paired-resource review. Its final validators and format pass
  after correcting one moved-plan docs link (`native-uia-closeout-final-20260930`). This closes
  only that library correction. Consumer I26 full suites, current gallery, integrated resources
  and final documentation remain required; H4/other owners are not closed.

- **September 30 V19 and alert diagnosis:** full Release passes without warnings/errors,
  receipt `d56aac031eded43bc46927c696484986b6f2efdfaff183624df8efdb87159f7b`;
  207 focused Pester checks pass without skips. Protected combined run
  `20260930T085745Z-60456-d933044f0fae41c18554558958efba5d` passes six cases including the
  new two-actual-pending-task inspection/other-task completion identity journey. Export fails
  the new semantic pixel witness: the presented PNG is still entirely black. Its downstream
  three cases are unexecuted by fail-fast. The separate tail run
  `20260930T090150Z-14192-69cfaffe9d8c4c1cbb74b459506cb056` passes Show log, Options denied-save
  recovery and actual worker Pause/Resume together (3/0/0). Both runs have zero disk findings
  and no desktop-ownership failure. Retain all original captures/failures.
  V20 freezes a failure-only full-frame replay diagnostic to distinguish alert content from
  HWND presentation; it preserves the original failed screenshot/result. Release build is
  running in `interaction-v20-20260930`; no product rendering change has been made yet.
  Root rejected the delegated provider-switch draft's reliance on previously enabled dummy
  state and its weak current-path check; it remains external and unapplied. The long DxUi
  crossover/restoration completed and its supplemental packet is copied/hash-verified in the
  qualified library checkout. Root's reconciliation excludes an agent's unrelated September 23
  menu experiment, retains all current flags, and does not infer a uniform performance pass.

- **September 30 V18 executing-worker coverage:** full Release passes with zero warnings/errors,
  receipt `a20456cd1dc137335c374f58c52d3ab9421155d97660213722b49f1a0caf1683`.
  Isolated protected speed run `20260930T082645Z-17880-15991585978a4ca7950c0d804fac844f`
  passes 1/0/0 after correcting the four Copy fixtures to the supported PerItem executor.
  Isolated Options run `20260930T083035Z-73300-3d8b2f6eb2ff490391039963616cec88`
  also passes 1/0/0, including denied save and distinct durable recovery. This does not resolve
  the earlier sequence-dependent failure: rerun the combined sequence on the next frozen build.
  Both archived receipts are under `protected-speed-v18-20260930` / `protected-options-v18-20260930`.
  Root visual inspection finds V17's Export `write-failure__fr-FR__144dpi.png` entirely black despite
  its behavior pass; retain it as failed visual evidence. The next test revision waits for the
  actual alert's painted frame and uses the existing screenshot color-probe facility to reject
  blank captures. V17's inspected 480-DIP French conflict capture shows the common filename once,
  stacked parents and complete wrapped actions; this is not whole-gallery acceptance.
  The reviewed external matched 1,000-frame DxUi crossover driver is running serially, preserving
  original short reports and restoring original benchmark source/builds; no performance result
  or final pin adoption is claimed yet. Remaining next actions: review/apply the delegated actual
  two-pending-task journey, freeze V19, build and rerun the combined actions and export capture.

- **September 30 V17 qualified routes and V18 correction:** V17 Release builds without warnings,
  receipt `e24f017009c3d01a782bae3384f6593fc75cb8fbfb2aceb0b37da5234abd7dec`.
  Native `protected-actions-v17-20260930`, run `20260930T081629Z-30208-7043667edb724f98b0f1f4e91aad1fae`,
  passes seven cases: the four existing conflict cases (including new filtered-item recovery),
  pending-inspection teardown/stale replay, completed Export (including French write-failure feedback
  and overlay teardown), and Show log. Options fails its combined denied-save compact guard; do not
  promote the isolated V16 pass to whole acceptance. Speed was unexecuted by fail-fast. Its separate
  run `20260930T081819Z-2876-3c16c2c716224aa88228b9ae29c733a1` fails waiting for execution: retained
  JSONL proves `strategy.bulkTransferRejected`, because four old speed Copy fixtures still request
  BulkItems while the current engine requires PerItem. Root verifies all four are Copy fixtures and
  updates their admission mode, retaining every failure. Options gains distinct invocation, applied
  setting and footer snapshot diagnostics. Both native runs have zero disk findings and no ownership
  failure. V17's 198 focused source/inventory checks pass; stdout also captured the Pester result
  object, so `pester-focused-recovered.json` explicitly records offline recovery from the retained log.
  V18 freezes the four fixture corrections and diagnostics in `interaction-v18-20260930`; build/native
  results are pending. No new product behavior or pin adoption is inferred from these test changes.

- **September 30 V16 native results and V17:** full V16 Release passes with zero warnings/errors,
  receipt `70569bba6d70608d32cd69eaec97a25961f9f3c882f89a6b118e7a027067ad5b`.
  Protected actions pass four existing conflict/inspection cases, then fail the new Export Escape
  guard before export execution. The independent Show log route fails opening completed More.
  Tail passes actual footer Options including denied-save/recovery, then fails the Resume guard.
  Preserve `protected-actions-v16-20260930`, `protected-unexecuted-v16-20260930` and
  `protected-tail-v16-20260930`; no ownership failure or disk finding occurred. Nonactivating
  Phase13 passes 3/0/0 in `phase13-v16-20260930`, run
  `20260930T080706Z-65692-0df1adbb935a4b22ad2eb960e5df9e5d`, proving future-only auto-dismiss.
  Root identifies the new completed fixture can expand before the one-time completion collapse;
  V17 waits for that transition and bounded visible-provider readiness, and splits Resume failure
  diagnostics without weakening identity/state requirements. A focused lifetime review catches
  two callback-storage cleanup defects in the external inspection draft and the same ordering
  issue in the existing navigation fixture. Root corrects destruction order, verifies synchronous
  cancellation, and applies actual filtered-target and pending-state-teardown/stale-callback tests
  with exact successful folder settlement witnesses. `interaction-v17-20260930` freezes 66/65
  inputs on isolated pin49; full build/native execution remain pending. Primary still pins73.

- **September 30 resumed after usage interruption:** V15's frozen full build failed with three
  C2597/C2641 errors: the Show log test directory seam referenced instance members from a static
  accessor. Root changes the two private path helpers to const instance methods and makes only
  the test mutex mutable; all callers already belong to the same state. Keep V15's failed build.
  Reviewed M29 coverage now drives actual Options actions while atomic settings replacement is
  denied, proves exact durable bytes remain unchanged while live settings/layout survive, then
  releases the lock and persists on/compact values distinct from the original off/expanded values.
  Root rejected the draft's initial recovery assertion because unchanged disk bytes could pass it.
  The same task, popup and retained completion are checked, and cleanup flushes restored settings.
  V16 freezes both corrections in `C:/RedSalamander.Perf/evidence/i26-ui/interaction-v16-20260930`;
  full build/native tests remain pending. Exact49 CI now succeeds in all six native configurations:
  root verifies 122 successful suite exits, eight consumer receipts and two expected ASan detections.
  Each ARM64 Menu suite records nine desktop-dependent skips; all other suite skip lists are empty.
  This does not claim those ARM desktop paths ran. The offscreen crossover report preserves all
  flags and does not establish a stable causal regression or a blanket performance pass; matched
  longer-duration confirmation is next, without replacing the original baseline or reports.

- **September 28 V15 qualification in progress:** the same 207 localization, harness source-contract
  and inventory checks now pass, including the two immediate activation-failure returns; a second
  207/207 run also covers the reviewed Show log addition. V15 freezes 66 logical/65 physical
  inputs on isolated candidate `49c963f`, retaining V14 unchanged. The completed More/Show log
  journey invokes the actual menu action, checks Escape leaves it uncalled, validates the selected
  task's owned UTF-16 JSONL path/content, and suppresses only the external application launch.
  Root verified the delegated draft against the production filename/BOM/writer and cleanup.
  V15 also includes the export `.txt` observer, worker-start fixture correction and UINT warning fix;
  its full Release build and these native journeys are pending. The fresh unchanged73/current49
  common benchmark A-B-B-A series preserves all eight reports: unchanged73 also triggers thresholds.
  This is being compared with the original baseline; no failed comparison or resource gate is waived.
  CI49 x64 artifact audit verifies 62 suite receipts and four consumer receipts with no suite skips;
  ARM Debug/Release jobs pass and ARM ASan completion is still pending. Raw current work is under
  `C:/RedSalamander.Perf/evidence/i26-ui/interaction-v15-20260928`,
  `protected-v15-driver-20260928`, `native-uia-final-perf-confirmation-20260928` and
  `native-uia-ci49-20260928`. Primary pin adoption and I26 closeout remain gated.

- **September 28 V14 actual route results:** full Release receipt
  `e53f1fa569709a6bbd733500b19b0d9f81bf8d23ca10acfe3f97e0915d96421f` passes with four test-only
  signed-conversion warnings (fixed in primary by explicit UINT lambda return). Both Phase5
  queue/Skip runs pass 3/0/0, including the formerly failing stale-provider reorder, with zero
  disk findings. Protected run `20260927T233038Z-91768-f3063a2326e4413399eab3ce9d423c91`
  passes completed navigation/Clear then exposes an export-test assumption: the existing TSV
  content is saved as `.txt`, not `.tsv`. Primary corrects only the filename observer.
  Independent tail `20260927T233225Z-72660-a6e34653529d4fa48bcc09cb2b281377` passes Cancel all,
  Queue/Parallel/Options, and retained Move/mixed informational Dismiss (3 passes); new worker
  actions stop at fixture setup because speed-dialog setup paused before ExecuteOperation sets
  HasStarted. Primary releases that setup pause, waits for actual execution, then prepares the
  paused state before testing real Resume/Pause/Cancel. No ownership failure or disk finding.
  A focused Pester run has 203/207 passes: three old assertions name fields replaced by immutable
  selected admission, and one exposes two missing immediate activation-failure returns in test
  fixtures. Root verifies the replacements against current key/scope/receipt ownership, applies
  the two fail-fast guards and updated assertions together, and reruns the same 207 checks.
  V14 remains frozen; these corrections require V15. Exact49 CI x64 Debug/Release/ASan passes;
  native ARM jobs and noisy offscreen resource confirmation remain pending.

- **September 28 library publication and V14 preparation:** reviewed lifetime/focus fix is
  pushed as `49c963f5751a6754e6b08e0362b96cfc7fc64414`; local 54 requested suite receipts and
  all three ARM64 cross-builds pass. Exact-source CI `36358394952` is running. Primary still
  pins 73ba until qualified adoption; the isolated V14 consumer will test 49c963f explicitly
  as a candidate. Root reviews/applies completed More export's actual UIA/TSV/refusal fixture
  and localized feedback together, with safe host-entry release on clear/popup teardown.
  Source establishes the original silent failure; no pre-fix native export failure is claimed.
  The fixture checks French feedback, retained task/report bytes, suppressed external launch,
  explicit clear and owner teardown. Independent lifetime review finds no new actionable UAF;
  the new consumer code and journeys still await the V14 full build and native execution.

- **September 28 final local library matrix:** all 18 requested nonactivating suites pass in
  each of x64 Debug, Release and ASan (54 passes). Missing ARM64 dependencies stop the first
  cross-build before compilation; the supported restore succeeds and three builds continue.
  Exact source hashes and original failure logs remain retained. Native foreground/CI and
  paired resources remain separate gates. Export feedback review identifies a stale-window
  overlay ownership concern; resolve it before applying the proposed feedback route. Current
  consumer pin remains 73ba and the new connected journeys are not yet compiled or passed.

- **September 28 native lifetime hardening and connected routes:** root retains a failing
  DxUi stale-provider regression, fixes native token publication/guards/runtime identity and
  canonical-root acquisition, then verifies the full Accessibility suite. A second regression
  reproduces focus-callback retirement; the reviewed fix preserves the chosen focus successor
  and prevents an old/replacement click. Final library Debug's 18 nonactivating suites and all
  static validators pass; Release/ASan/ARM cross-builds and paired-resource reconciliation remain
  in progress. The consumer still pins 73ba until explicit qualified adoption of the new commit.
  Root reviews/applies cheaper-model actual UIA journeys for Pause/Resume/Cancel, Cancel all,
  Clear completed, Queue/Parallel and Options. They await the next frozen consumer build.
  Source review also confirms auto-dismiss violated the approved future-retention contract:
  enabling it erased already retained records. Primary now removes only that retrospective
  purge; Phase13 and actual Options tests preserve prior records while requiring future success
  and cancellation auto-dismiss. Completed More export currently hides write/open failures;
  its connected fixture and feedback correction are being drafted separately.

- **September 28 V13 speed and independent evidence review:** protected speed run
  `20260927T225111Z-84340-b6cf08e963af4aae8d2e007d3f5d69c7` passes all four cases (4/0/0),
  zero disk findings, including live locale/edit preservation and repeated dialog lifetime.
  The earlier tail run `20260927T224934Z-90012-fc9251233fc84a189fb7acab54bd44fe` passes More
  placement, then fails M30's test observer: that observer deliberately selects file-operation
  cards, not informational cards. Root reviews/applies actual exact-ID informational disclosure
  and Dismiss discovery instead; the new journey remains unverified. No ownership failure occurs.
  Root independently checks the cheaper agent's four-run V12 resource comparison, recalculates
  all 16 median/P95 series and verifies eight retained source hashes. French-culture CSV export
  initially serialized the fractional median with a decimal comma; root fixes numeric CSV
  serialization to invariant culture and regenerates offline as
  `integrated-final-comparison-4run-v2-20260928`. Original outputs remain retained. All four app
  results pass; the B1 collector's false classification remains alongside its separate recheck.
  Timing samples remain descriptive; one extra configured text-format cache entry appears in
  Issues, with logical targets/render/resize counts unchanged. No blanket resource pass is claimed.

- **September 28 V13 runtime attribution:** full Release builds, receipt
  `bef9d9349dab216ef55c7f45356e368a25338679e528c344a9911818322f92cf`.
  Two test-only LONG-to-FLOAT warnings are corrected in primary with explicit double conversion;
  frozen V13 evidence keeps those warnings and its exact inputs. Queue run
  `20260927T224744Z-94348-ef87c5cf22024e3fa035042660351a0a` confirms the native lifetime defect:
  the same retained C-Up provider changes AutomationId from task 3 to task 2 after rebuild,
  returns S_OK and reorders the wrong task. Focus/foreground remain unchanged. Do not weaken
  the stale-reference assertion. DxUi's separately indexed `NativeProviderLifetime_2026-09-28`
  plan owns the generic guard and native regression; existing Grid/Menu closeouts remain scoped.
  Protected V13 More/M26/M30/speed verification is running with existing warning/lease/restoration.
  A bounded interaction audit also identifies missing actual action-route witnesses for per-task
  Pause/Cancel, Cancel all, Clear completed, Queue/Parallel and Options menus, and completed More
  export. Existing engine/internal-helper checks do not replace those user-route checks.

- **September 28 V13 correction/build and library closeout:** DxUi Grid/Menu plans move to
  Done in pushed documentation/evidence commit `3022541`; exact qualified code remains 73ba.
  Root checks Sol's closeout audit and corrects local-versus-paired resource wording. All library
  skills/spec/dependency/format checks and both portable packet verifiers pass. This completes
  those library slices and pin handoff, not consumer I26/H4, main merge, or native ARM foreground.
  V12's live-locale failure is a real product bug: `SetWindowTextW` reaches the hosted edit's
  `WM_SETTEXT` compatibility route and replaces the invalid speed with the window title.
  Root reviews/applies the narrow default-window caption update plus independent chrome-read
  assertions and per-condition diagnostics. V13 freezes the change, reviewed M26/M30 journeys,
  corrected More alignment witness and queue diagnostics; full Release build is in progress.
  New V12 `Phase5_DiscoverySkipContinues` passes 3/0/0, zero disk findings. Queue reorder passes
  its first real UIA mutation but fails later identity/order checks; V13 records each stage.
  Source review finds native UIA control paths lack the existing embedded lifetime-token guard;
  confirm whether the retained provider reaches another control after rebuild before fixing it.
  V12 resource B1, old-pin A2 and V12 B2 each pass 1/0/0 with zero disk findings; the offline
  comparison remains pending. B1's collector incorrectly indexed a single expected name as a
  string; the original false verification stays retained, with independent result recovery.
  Do not rerun successful application measurements to conceal a collector error.

- **September 28 current head of checkpoint — V12:** full x64 Release passes with zero
  warnings/errors, receipt `db1424aa978a5589369df2bc44316f1dff8cce9addcc57918ef1080f195e24cd`.
  Static gallery passes 1/0/0, zero disk findings, run
  `20260927T222051Z-104844-c758af929aaa425d9452accab98351f4`; 576 runner files are archived
  under `queued-unknown-v12-20260928`, including both distinct 124a/124b subjects on real 96/144 DPI.
  Protected tail run `20260927T222457Z-98000-65bd1e717ec34015ac9e0f9bd82274d9` passes five
  cases, including the corrected connected accepted-decision witness; completed More placement
  fails and four speed cases are unexecuted by FailFast. No ownership failure occurs.
  Root confirms that measured actions may occupy the left half: the existing anchor chooses
  that button's left edge there. The old test always demanded its right edge. Primary corrects
  the assertion to match the invoking side and adds actual coordinate diagnostics; this awaits
  execution. The independent protected speed run stops on its live-locale state assertion
  (0 passed / 1 failed / 3 unexecuted); all evidence remains in `protected-speed-v12-20260928`.
  Diagnose the exact differing state, never remove the preservation requirement.
  Root reviews/applies the external M26/M30 draft: actual pane Cut/Paste/Move/Skip and terminal
  source-recovery menus, plus host informational update/Dismiss isolation beside a real pending
  Copy. The latter does not claim producer-engine cancellation. Both new journeys await build/run.
  Next: execute V12's new nonactivating Phase5 UIA controls and matched final resource fixture;
  finish locale diagnosis, freeze the next candidate and rerun changed/previously unexecuted cases.
  Older checkpoint entries below are chronological evidence, not current pin/grant/gate status.

- **September 28, V11 protected results and V12 correction:** the protected session finishes
  with interaction gallery PASS and Commands 43 passed / 1 failed / 9 unexecuted by FailFast,
  zero disk findings. Full evidence is `protected-session-v11-20260928`; Commands run
  `20260927T220900Z-98472-c58bf15c443742a1aa774136871fb1c8`. No ownership failure occurred.
  The new unchecked-scope test prematurely required the first parallel mutation to finish
  while a second prompt is active. `Task::WaitWhilePaused` intentionally blocks sibling workers
  during that prompt. Only this connected round now uses the existing one-worker test override,
  asserts its admitted budget, and restores it after fixture cleanup. The original concurrent
  cached-scope tests remain unchanged. The speed and pointer-inspection cases were not reached.
  Root also reviews/applies actual hosted Start/reorder/Skip Phase5 paths with exact task IDs,
  stale-provider/neighbor isolation and an enforced nonactivating guard. Aggregate gallery scenes
  now use distinct spring/winter photo subjects: their prior duplicate-looking rows were separate
  task occurrences; production already coalesces each exact task/generation and remains unchanged.
  Freeze V12, compile and execute these changes before claiming their coverage. Remaining native
  work uses notified protected continuation with strict ownership checks, never unchanged retries.

- **September 28, explicit qualified pin adoption:** primary `Dependencies/DxUi.lock.json` now
  selects `73ba9365726cce30299290ab9ae8afa065c4fe3d`. Root verifies the six native successful jobs
  in run `36352406442`, source/architecture identities and all 361 curated CI originals. Each ARM64
  profile retains nine Menu interactive-desktop skips; x64 profiles have none. This is explicit
  feature-commit adoption, not the main-only `Update-DxUi.ps1` selector and not consumer H4 closure.
  V11 full x64 Release passes with zero diagnostics, receipt
  `3709184ac07d4424e867e8fbe9ecc1cbebdb1f7df3dba9d0107aede420b85cb0`; static gallery passes
  1/0/0, zero disk findings, run `20260927T220505Z-95296-ab0a0caf2d3648a881db1ad5bdcb3223`.
  New 124a/124b scenes/export assertions pass on both real displays. Protected interaction gallery
  passes run `20260927T220822Z-98472-5afe00bcfc054fb4a939e5c7bc0a1ed3`; the same authorized
  protected session continues with all 53 `cmd_pane_fileops_` cases, with FailFast. The desktop
  availability grant is now in use, not a fresh grant for a later retry. Keep V10's failed compile.

- **September 27, V10 compile repair before V11:** the full V10 Release build fails with two
  C2664 errors converting `wstring_view` to the speed field's owned accessible-name string.
  Primary fixes both calls with explicit `std::wstring`; the frozen V10 source and failed build
  log remain unchanged. No V10 gallery ran. V11 is frozen (64 logical / 63 physical inputs) under
  `queued-unknown-v11-20260927` and is building in the same isolated checkout with exact pin73ba.
  It includes the repair and the reviewed connected conflict checkbox/action/worker witness:
  real checked consent reuses the accepted choice; unchecked consent leaves the next generation
  unresolved and its destination unchanged, then cancels through the actual hosted action.
  Source/destination inspection also gains owned pointer-message invocation in both active panes.
  These are source changes awaiting execution, not physical-pointer or external-AT passes.
  The next protected-session availability grant remains unused. Its reviewed driver uses FailFast
  and the existing warning/lease/restoration; no second attempt follows an ownership failure.

- **September 27, final candidate V10 frozen:** DxUi optimization `65b0257` and metadata-only
  correction `73ba9365726cce30299290ab9ae8afa065c4fe3d` are pushed. All 54 local nonactivating
  suites pass; exact-head native run `36352406442` is still in progress. Its x64 ASan Menu bound
  now passes; no full six-profile pass is claimed yet. The isolated consumer `C:/RsI26Adopt` freezes
  64 logical / 63 physical inputs with candidate pin 73ba under `queued-unknown-v10-20260927`;
  primary pin remains 36b until explicit qualified adoption. In addition to 124a/124b, it contains
  the reviewed speed-dialog scrolling body/fixed actions and owned live-locale refresh. Root fixes
  the draft's throwing COM-pointer query before applying it. The extended existing Commands case
  checks real validation at 132 DIP, help End/Home, UIA focus and retained English/French edit state.
  This is unverified until build/native execution. The user confirms desktop availability for the
  next protected harness session, with the existing warning/lease and restoration; that grant has
  not yet been consumed. Fresh old-pin resource baseline A1 passes 1/0/0, zero disk findings, run
  `20260927T214004Z-107560-263f0c4b63604bc2a8cc379b57f95e95`, under `integrated-final-A1-20260927`.

- **September 27, large-menu blocker corrected in final qualification:** the forced-rebuild
  ASan comparison preserves the failing original result (End-to-visible 1,022,545 us) and measures
  344,345 us with exact-capacity, in-place accessibility snapshot construction. No semantic rows
  or timing bounds are removed. Independent review passes; temporary production probes are removed.
  Final library x64 profiles are running under `menu-plain-uia-20260927/profiles-navigation-final`,
  followed by a new native six-profile CI run and explicit consumer adoption. The earlier d45 run
  is complete: five native profiles pass and x64 ASan fails that latency bound; it is not qualified.
  The invalid incremental A comparison is explicitly excluded because its old source timestamp
  did not force recompilation; both corrected variants rebuild and retain source/binary receipts.
  Primary also adds reviewed Copy/Move aggregate-Unknown gallery/export witnesses (124a/124b),
  with exact accepted-decision provenance and an empty per-item TSV result. These are unbuilt until
  the next frozen consumer candidate; the aggregate operation success must not invent a conflict
  result association. Current ordinary-menu resource cost and previous grid/menu timing approvals
  remain accepted. Next: finish library profiles/CI, freeze the next consumer with its qualified pin,
  run the new scenes and current broad/native/performance acceptance; I26 remains open.

- **September 27, ordinary-menu native focus correction qualified locally:** DxUi `d45c3610`
  passes all 18 nonactivating suites in x64 Release, Debug and ASan Debug, with fresh timestamp/
  executable-hash receipts in `menu-plain-uia-20260927/profiles-d45`. Its native CI run
  `36349836006` is still running. Earlier `15be545` CI exposed first-Down skipping a command after
  native activation; the test fails before the narrow menu-owned focus correction and passes after.
  Plain/described initial selection, focus reentry, clipped-row UIA geometry and scroll-to-focus
  are covered. These local passes do not replace native CI or consumer qualification.
  The final ABBA ordinary-menu fixture measures about 23 KiB additional live heap for 12 rows and
  226 KiB for 128 rows, without corresponding retained growth in four closed-menu cycles.
  The user accepts that specific accessibility cost on September 27; previously accepted grid
  memory and grid/menu timing tradeoffs remain accepted. Original common-benchmark flags and noisy same-
  variant controls are retained, not silently rebaselined. Next: isolated V9 consumer candidate
  build/output checks while native CI is pending; primary dependency pin stays unchanged.
  Native x64 Debug/Release pass; ASan fails the existing 4,096-row menu End-to-visible one-second
  bound. A test-only timing breakdown and nonactivating reproduction retain the original limit.
  V9 full Release builds with zero diagnostics, receipt
  `b0d673186caa9bbd225ede48d1e348b19bff20a149ad4542acf1bc0d0bdce4f0`, and component passes
  3/0/0, run `20260927T210753Z-51644-dda0050ee8ae41c9993f236d8aa67d78`, with zero disk findings.
  The actual menu/output journey also passes 3/0/0 in 1,593 ms for the selected case, run
  `20260927T210904Z-84856-1bf9a29f87a1469fb277c99b4c19b168`, with zero disk findings. It invokes
  real history-menu UIA commands, checks cancel/stale replies, writes both Copy files only to the
  new folder, and verifies the logical clipboard Move's bytes, source removal, durable destination
  record and once-only admission barrier. This is isolated V9 consumer evidence, not final pin adoption.
  All nine R4 overlap/live-output cases plus setup/cleanup pass 11/0/0 on that same build, run
  `20260927T211358Z-76148-e8da9514160f4ba9bc6cc8eca147d1b4`, zero disk findings. Root verifies the
  [V9 packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_queued-retarget-v9/README.md)
  against all 254 payloads, 64 frozen inputs, original bytes and exact archive membership.
  The large-menu failure now reproduces locally under the established no-activation blocker:
  open 3,971,199 us; End dispatch 1,097,268 us, pump 8,673 us, query 138,909 us (1,244,850 us total).
  It reaches the correct last row and paints only eight rows. Next library action is attribution
  within synchronous dispatch, retaining the one-second limit and all semantic rows.

- **September 27, V8 source-continuity witnesses pass:** frozen 64-input full Release build
  `cc9e08145d63e188d74189c08b390063af7cffb94e7112cfdd3727cffbf54f38` succeeds. Component passes
  3/0/0, selected case 687 ms, run `20260927T201900Z-95836-00afb59d8d0440938c4578bbc8c4958a`,
  with zero disk findings. It exercises real cancellation after durable Move replacement,
  original-source replacement before/after retarget, provider-relative nested/renamed mappings,
  root DirectoryShell mapping and four late source replacements at actual transfer attempts.
  One C4456 shadow warning is fixed in primary after this freeze; the next full build must verify it.
  V8 does not yet pass the actual menu/output journey or current-source broad acceptance.
  Root verifies the [V7 packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_queued-retarget-v7/README.md)
  and [V8 packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_queued-retarget-v8/README.md)
  against original bytes, archive membership and hashes; both contain portable verification.
- **September 27, library ordinary-menu correction in qualification:** the two incorrect
  description-only UIA guards are removed. Fresh Release `MenuAccessibility` runs B4/B5 pass,
  with invocation timestamps and executable hashes verified, followed by 18 nonactivating Release
  suites. Their A1/A2 old-code controls fail the expected missing-child assertion. Evidence is under
  `C:/RedSalamander.Perf/evidence/i26-ui/menu-plain-uia-20260927` pending repository curation.
  B1/B2 wrappers stopped at the common performance comparison before running the UIA suite;
  their collector copied stale A2 receipts. Those copies are invalid candidate correctness evidence
  and remain retained with a correction. Diagnostic instrumentation is removed. Independent review
  additionally finds that clipped menu rows incorrectly report `IsOffscreen=false`; the newer
  correction above preserves those rows in navigation and refreshes geometry after scrolling.
  Library resource advice, native profiles and explicit new qualified consumer adoption remain open.

- **September 27, V7 menu failure diagnosed:** pinned DxUi `36b2f4b` returns before creating or
  synchronizing UIA rows whenever `descriptionLayouts` is empty. The real Dummy destination menu
  contains ordinary single-line entries, so its visible popup exposes no item Invoke providers.
  Root confirms the cheaper-model diagnosis in library source. Fix ordinary-menu accessibility in
  the canonical library, qualify it independently and adopt an exact new pin; do not add artificial
  consumer descriptions merely to activate UIA. The unchanged Release library performance baseline
  is retained at `menu-plain-uia-baseline-20260927`. The V7 failure remains evidence, not a fixture pass.
- **September 27, transaction V7 build/component receipt:** frozen 64-input full Release build
  `3cf1761d3c671f0c260500df1f03bdc4c935cb803e4eb06b782d3d46763ab6da` succeeds; two new shadowing
  warnings are fixed in primary for the next candidate. Component passes 3/0/0, run
  `20260927T195637Z-86712-912aede2066b44f9bd12f0d324c6ab4d`, disk findings 0. The real menu/output
  case fails at its UIA menu lookup in 234 ms, run `20260927T195742Z-84640-3c510dfe1f264b0e9e7324e82b3b8386`,
  preserved under `queued-retarget-output-v7-20260927`. Diagnose the actual history/accessibility
  mismatch before rerunning; do not bypass the menu action. Resource localization passes 9/0/0
  with transcript/XML under `retarget-static-v7-20260927`.
  Root and independent review catch a missing transfer reader: retaining old source scopes alone
  did not stop pathname rebinding to a replacement file. Primary now checks exact retained object,
  revision and kind before candidate publication and again against the transfer guard's source
  binding. Real Move cancellation after durable replacement, source replacement before/after
  retarget and logical clipboard-Move once-only witnesses are applied but not yet built/run.
  Empty-relative-path DirectoryShell mapping now retains its legal mapping to the new root.
  Next: finish explicit-mapping/menu diagnosis, freeze V8, run full build and these exact witnesses.
- **September 27, V6 complete and transaction candidate in progress:** root verifies the
  [eight-outcome V5/V6 packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_queued-ownership-v5-v6/README.md),
  including 392 original payloads and both frozen 58-input source sets. Full Release V6 has zero
  warnings/errors, receipt `af1cc0cd8342702c1a907f294770befe6981cd313c498375b901bd7781e2f77a`.
  Component, readiness, provider matrix, permanent Delete and inline-F2 interlock each pass 3/0/0;
  all nine overlap cases plus setup/cleanup pass 11/0/0. Every run has zero disk findings.
  V5's component failure, V6's focused-build attestation refusal and the PowerShell 5.1 launcher
  refusal remain preserved. The same-machine V4/V6 descriptive comparison flags four timing
  groups above its 10% review threshold (two have one observation per side); no regression waiver
  or repeatability claim follows from those single runs. Repeat paired qualification remains open.
  Primary source now contains private retarget preparation, immutable source-authority retention,
  provider-relative mapping, captured menu context and the Move-record helper. Queue transaction
  and real-output fixtures are under review; this newer source has not compiled or passed M06.
  B remains the frozen V6 source until the complete candidate is reviewed. Next: finish/apply
  the transaction kernel and asynchronous output witnesses, full-solution build, component/race
  and real Copy/Move output tests, then current-source broad/native/resource acceptance.
- **September 27, coherent destination migration under review:** root applies the reviewed
  11-file selected-view draft, checks every file against the supplied candidate (normalizing Git
  line endings), removes one redundant initial view allocation and formats changed ranges.
  Destination metadata, typed plans, receipt, facts and interlock now share one retained view;
  aggregate execution/UI/completion readers are migrated. This does not implement M06: its old
  display-only setter is an explicit no-op pending the next contiguous request transaction.
  No release/acceptance or M06 pass is claimed for this intermediate state. Root also applies
  the reviewed Issues append/no-op/sort/remove scroll-and-selection fixture, with prior pane
  existence, disclosure, settings and visibility restoration. Its protected Commands lane remains
  unexecuted. Source is in the primary checkout; the running B clone remains the prior V6 input set.
- **Ownership V5/V6 qualification correction:** V5 compiles with zero diagnostics, but its
  component fails because Dummy does not implement bound objects. The strong independence
  assertion remains; V6 uses canonical Local provider construction and distinct real destination
  parents. All V5 source bytes are frozen. V6's focused application build succeeds, receipt
  `d247c42cc5d1c914312569e5e047f86ea9a8f023c0a43c2a7f87c2e7c8375d0a`, but the test runner
  refuses its non-full-solution attestation before any application case runs. This setup failure
  remains at `queued-admission-contracts-v6-20260927`; a required full-solution retry runs at
  `queued-admission-contracts-v6-full-20260927` against the same 58 frozen input hashes.
  Next: component and same-build R4/R1d/provider/permanent-Delete regressions, retained paired
  ownership metrics, then selected-view qualification and the bounded retarget transaction.

- **September 27, immutable registration ownership in qualification:** root reviewed and applied
  the external draft, with named retained scope owners, immutable decision revisions and exact-key
  withdrawal. The component now binds a real second destination to prove authority independence,
  and rejects a late reply from request 2 while request 3 is pending. Logical owner-object counts
  and bytes are instrumented separately from heap/control-block bytes; their timers exclude metric
  emission. V5 freezes 58 inputs; its fixture failure and V6 correction are recorded above.
  Coherent destination inputs and the transaction remain pending; M06 is still deliberately open.
- Root verifies and copies the [V3/V4 seven-run packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_queued-preparation-v3-v4/README.md):
  373 mapped original payloads and seven passing result files. V4 component, preparation, permanent
  Delete and provider-capability cases pass 3/0/0 each; the nine overlap cases plus setup/cleanup
  pass 11/0/0. All have zero disk findings. The original 256-by-256 scope and 4,096-root readiness
  fixtures remain available for the ownership comparison. These are functional receipts, not a
  blanket timing non-regression claim. The reviewed packet fixes an omitted V3 run and truncated ID.

- **September 27, exact-admission permissions:** Release V3 builds without diagnostics,
  receipt `2a610f5846a439f139d65a9f0fd8ad1045898ce90c3eede9e3c6a3488ee07536`.
  Component case passes 3/0/0 (`20260927T184118Z-35836-5fd99df866d74ec9b9af8f5886881ed3`);
  all nine overlap cases plus setup/cleanup pass 11/0/0
  (`20260927T184223Z-92996-c94b8cec548b4dd995a433a5618f8247`), both with zero disk findings.
  All 56 inputs are frozen under `queued-admission-contracts-frozen-v3-20260927`.
  V4 also qualifies the reciprocal-grant owner-key check. Exact key predicates do not yet make
  task-owned authority/relations safe for simultaneous original/candidate registrations.
- The next slice generalizes the existing Batch Rename artifact-prompt transport for private
  candidate preparation. Request IDs correlate replies; a successful receipt remains private
  until the caller installs it. Stale nonzero payload tokens are ignored; only the existing
  zero-lParam fallback can wake an uncorrelated failure. The cancellation/output-preservation
  and stale/duplicate-reply component tests plus existing nested-shutdown fixture are updated.
  V4 Release builds with zero diagnostics, receipt `6d60675eba886f47be0411aa3aebc412a7d2762e9a90ff17479d637ae675dc97`.
  The component passes 3/0/0 (125 ms selected case), run
  `20260927T185332Z-86932-61453f9ff9ad42c2a169dcabf62b0dc4`, zero disk findings.
  All 57 inputs are frozen under `queued-admission-contracts-frozen-v4-20260927`.
  The actual Commands nested-modal shutdown test still needs the protected interactive lane;
  the private component does not establish that gate. No queued destination transaction is enabled.
- Next: registry-owned immutable scopes/relations, coherent selected destination inputs, then the
  queue reservation/durable Move-record commit and actual output/race witnesses. Root remains the
  sole repository writer and native runner; external agent drafts are reviewed before application.
- Root checks the [six-run V2 packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_queued-preparation-v2/README.md):
  272 mapped raw payloads, every copied runner hash, both 55-file source audits per run and
  all 55 frozen source inputs match. Tool/spec inventory validators report zero blocking
  findings and TestInventory Pester passes 7/0/0 with transcript and NUnit receipts in
  `queued-admission-static-v3-20260927`. The first Pester attempt requested unavailable v5;
  the version-compatible invocation passes. Its first console capture was empty, so the
  final repeated invocation retains an actual transcript and per-case XML.

- **September 27, private preparation qualification:** x64 Release V2 builds with zero warnings
  and errors, receipt `379bafb0cde9ce6106ee840a94e2edf2611e79d552056282063f2bcdfbe08bea`.
  `FileOps_QueuedAdmissionContracts` passes 3/0/0 including setup/cleanup, selected case 109 ms,
  run `20260927T182800Z-93180-0e6473488c044f4bab262801132a842a`, zero disk findings. Five subsequent
  exact cases on those frozen 55 inputs also pass 3/0/0 each: ProviderCapabilityMatrix,
  RC3_9_TransferDestinationNameRefused, R1d_PreparingLifecycle, C1_PermanentDeleteConfirmsOnCard,
  and R4A02_RenamePublishKeepsIndexPrecise. Raw evidence lives under
  `C:/RedSalamander.Perf/evidence/i26-ui/queued-preparation-<case>-v2-20260927`;
  source bytes are frozen at `queued-admission-contracts-frozen-v2-20260927`.
  These passes qualify the extracted preparation/breadcrumb helpers, not the M06 transaction.
- The next independent slice binds overlap Run/Queue relations, registrations and live-output
  indices to exact admission keys while retaining stable task IDs for presentation. Private tests
  exercise the real scheduler/live-output predicates against stale and renewed grants. Name tests
  now include rejected-leaf facts and later success clearing; cancellation is rechecked after the
  provider call. These newer edits are under review and require their own build/test receipt.
- Root independently verifies both cheaper-model evidence packets against all raw/frozen bytes:
  [logical host resources](../../TestRuns/4cb089111a23/Validation/2026-09-27_host-resource-pair/README.md)
  (180 mapped payloads, 13 phases), and
  [destination history gallery](../../TestRuns/4cb089111a23/Validation/2026-09-27_destination-history-gallery/README.md)
  (79 mapped payloads, 50 frozen inputs, four trace witnesses, 572-entry runner inventory).
  Both packets are copied into repository evidence. The resource comparison still reports Issues
  text-format entries 8 vs 7 and unavailable driver/peak allocation; it is not an FPS comparison.

- **September 27, queued destination follow-up:** the corrected real-menu witness reproduces
  M06's display/plan disagreement in 281 ms, run `20260927T180744Z-93896-22046a94b2334680a0e4e78cec94bc4f`
  (0/1/0, zero disk findings; Release receipt `59e588c957c1465c6b243fc2546f1b6f89921a258cf764754d96ed1fed7f27dd`).
  The first helper mistakenly treated the normal activation handler's zero return as menu failure;
  root adopts the reviewed fix and verifies the owned menu HWND instead. No test result is waived.
  The new engine groundwork splits captured admission construction from publication and adds atomic
  Admitted-breadcrumb replacement with rollback. `FileOps_QueuedAdmissionContracts` is a new normal,
  nonactivating component case; compilation/execution is pending. The existing Commands lifecycle
  test requires foreground isolation, so its attempted noninteractive preflight correctly stopped
  before launching; no desktop attempt was made and no pass is claimed.
- Described destination gallery now passes 1/0/0 with 572 hash-verified artifacts, zero disk findings,
  run `20260927T180002Z-93948-f8786715040c479db2fc4556f331194c`, on receipt
  `9d356df19ee57097f81c986915e91cac11702f0016f75ed15085b6b99d22e957`.
  Root and a cheaper-model audit inspect the originals: all seven rows fit at 480/760 DIP on native
  96/144 DPI, including repeated leaves, ampersand, drive and UNC roots. Full UIA names are unique.
  This synthetic display case does not establish the real queued callback, 192-DPI or AT acceptance.
  The same candidate also passes the resource case, run
  `20260927T175445Z-81116-f1a60bf507014577a99b6130dc358a04`. A/B target extents/bytes, brushes,
  renders and resizes agree in all 13 phases; Issues retains one additional text-format entry (8 vs 7).
  These are one-run sampled counters, with different destination-menu source inputs; GPU allocation
  and peaks remain explicitly unavailable. Original50 inputs are frozen outside the mutable clone.
- **September 27, 17:33 UTC checkpoint:** combined DxUi native CI/resource publication is
  committed/pushed as `4919471` on `codex/fileops-ui-qualified`; qualified code remains `36b2f4b`.
  Consumer `Dependencies/DxUi.lock.json` explicitly adopts that code revision in the working tree.
  The short candidate checkout `C:/RsI26Adopt` builds x64 Release and passes the integrated smoke
  (`20260927T171727Z-22032-af3b24d361c948508b2a3dc5fad1cb64`, receipt `8602c2633e9aabfbeb5705da5fa8a16aaa4e1fc27e526365600e0e0f6eeb4700`).
  Eight fresh application processes pass in predeclared A1/B1/B2/A2/B3/A3/A4/B4 order with old
  `78b3` versus new `36b2`, identical fixture/workload/geometry, zero skips and disk findings.
  Root checks all 258 raw payloads and the reversible map, actual results and sample medians in
  the [paired consumer packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_integrated-pin-pair/README.md).
  Redraw-batch differences are small and mixed-sign; sampled private/working-set memory likewise
  shows no consistent new increase. Retain the isolated B3 +9 process-handle endpoint difference;
  other repeats do not reproduce it. These samples do not measure surface/cache/driver peaks.
- Cold restore exposed historical DxUi receipt paths beyond MAX_PATH. The initial candidate
  restore failure remains under `consumer-pin-36b2-20260927`; clean-source rejection worked.
  New managed dependency clones now set repository-local `core.longpaths=true` before checkout.
  An owned failed clone was repaired explicitly; tooling never repairs an existing dirty source.
  The real local-Git Pester witness passes 10/0/0 under `dxui-restore-longpaths-tests-v2-20260927`.
  Its V1 test-fixture parsing failure remains separately retained.
- Working-tree destination history now uses local leaf/parent descriptions, complete UIA names,
  literal ampersands, and exact fallback for drive/UNC roots or unprofiled remote paths. The
  nonactivating gallery checks seven rows, repeated leaves, long French text, roots, geometry and
  raw UIA identity. New test-only host snapshots observe current-target logical bytes/cache entries
  without inferring actual driver allocation. These newer edits still require compilation/native
  execution. A cheaper-model audit found missing actual queued-task selection/cancel/start-race
  coverage; its UIA-based fixture is being drafted, with root as the sole repository writer/runner.
- The old-pin logical-resource extension now builds without diagnostics and passes 1/0/0, zero
  disk findings, run `20260927T174751Z-96192-042acaa40777481ba30848057222f7ff`, receipt
  `5993bff200a20958578297ceb9a52b3cfd7fe4e3c74806104710976a4fa0fe14`.
  Popup target is 9,054,592 logical bytes (1114x2032 BGRA8), Issues 3,283,200 (1140x720);
  forty steady redraws, 256 churn updates and both hidden intervals preserve extent/cache counts.
  Actual GPU allocation/peaks remain unavailable. V1's WIL `com_ptr::query_to` return-policy
  compile failure is retained; V2 uses `com_ptr_nothrow` for explicitly checked HRESULTs.
- **New M06 correctness finding:** source review of the real journey shows `SetDestinationFolder`
  changes only the display/backing field, while execution resolves the previously admitted
  `TransferPlan` destination. There is also a pre-lock started-state check. The new phased test
  must check the admitted plan, not merely a changed caption. Correct re-admission must requalify
  endpoint authority, mutation interlocks and Move breadcrumbs; do not patch an immutable plan's
  destination string or weaken its identity guard. Root/cheaper-model architecture review is active.
  Candidate build `described-destination-v1-20260927` includes the honest failing witness plus
  menu display and resource changes; do not count it as M06 completion until actual execution agrees.

- **Current next action, September 27:** the user answered **"Accept the recorded timing
  tradeoffs"** for the stated grid/menu costs. This supplements the prior grid memory acceptance;
  it does not waive correctness, new resource regressions, combined-build, DPI or accessibility
  gates. Historical unresolved-advice notes below are superseded for those recorded costs only.
  Menu ablation and the durable contract are committed/pushed as `0f35bab`; all six ordinary
  A1/H1/B1/B2/H2/A2 runs succeed, all nine original comparisons remain `advice-required`.
  Root verifies 86 staged payload hashes and all four library static gates. The failed V1 H
  dependency preflight is retained; V2 restores dependencies and completes the experiment.
  Different executable images and C:/Z: paths prevent a causal claim about the header fields.
- Root merges qualified grid `047414e` and menu `0f35bab` in the existing managed checkout,
  on `codex/fileops-ui-qualified`, committed/pushed as
  `36b2f4b49d42248790fa1a4ad6b6409ac56207dc`. Both benchmark command families, JSON diagnostics,
  contracts and tests are preserved. Root checks a cheaper-model review and preserves the public
  debug record's original aggregate prefix. All four static validators pass. Native CI
  `36332676309` passes all six native profiles; the local nonactivating Release lane passes
  all 17 selected suites with zero skips and republishes all five gallery themes. Root and a
  cheaper-model reviewer inspect the originals. Combined qualification and explicit consumer
  adoption are recorded above. Inherited component receipts cannot qualify the merge.
- V14 visibility diagnostics are committed/pushed as `b5b0996`. V15 adds a bounded raw UIA
  traversal for the real hosted scope Checkbox: unchecked/checked controls each occur once;
  metadata loading exposes none, including offscreen descendants. Full Release builds without
  warnings/errors, receipt `5077f8a8d1de0a89b3a4adef42e32c5d4d6409c5084819c4ffaf3a00e3e8ff9c`.
  Run `20260927T160901Z-16232-22c410209b8245a698010fcddf238554` records all six scope witnesses
  valid at 96/144 DPI, all sixty terminal transitions and both migration witnesses. It nevertheless
  fails 0/1/0: existing waiting-animation checks observe only one frame in several 144-DPI phases
  after long delays, with fitted geometry and a true settled predicate. Root verifies all 571
  copied files and zero disk findings. Do not promote this aggregate as a pass.
  Evidence: `C:/RedSalamander.Perf/evidence/i26-ui/scope-uia-v15-release-20260927`.
- The next fixture revision bounds waiting-transition/collapse message batches to 16 instead
  of the shared default 1024 and requires two observations inside the settled predicate as well
  as the unchanged final assertion. It logs elapsed time and message counts, retaining the
  1500-ms budget and all fit checks. This addresses observation starvation as a harness concern;
  V15 did not record enough pump detail to prove the long-delay cause. Production is unchanged.
  V16 now builds without warnings/errors and passes 1/0/0 in run
  `20260927T163450Z-4896-0058439a37c24305a46cd78aadd5accc`, receipt
  `f23f969006941dc7b07f1e9887d10574faf3d762813c04584880cd4be39de2b1`.
  Root checks all 572 copied artifacts, 40 frozen inputs and 48 fitted/settled waiting observations
  with at least two frames. Six scope, sixty terminal and two legacy-migration witnesses pass;
  disk audit is zero. [V15 failure](../../TestRuns/4cb089111a23/Validation/2026-09-27_scope-uia-v15/README.md)
  and [V16 result](../../TestRuns/4cb089111a23/Validation/2026-09-27_scope-uia-v16/README.md)
  are retained separately. The original V9 footer failure remains unexplained.
- Root reviews and corrects a Sol draft of `FileOps_IntegratedPresentationPerf`: eight running
  cards, first-card graph disclosure, all-card churn, 127 warnings plus one typed detail record,
  hidden/show/teardown samples and fixed dates. Cancel must remain a live hit target. Copy retains
  its source and does not claim verification. The unchanged common P0 fixture remains separate.
  Full x64 Release builds with zero warnings/errors and the exact case passes 1/0/0, run
  `20260927T165846Z-87056-f1e2ed16ce844359bb2e0ad319f8e1f6`, receipt
  `43590230e40359ab6f85486c8ac207907c2da9532f8698ca5af351704ecc3295`.
  Root verifies all 44 frozen inputs and 28 copied artifacts. The baseline records 17 resource
  phases, five repetitions per redraw state, 32 churn batches, actual 144-DPI geometry and
  workload digest `11213680259031909754`; host/attachment counts return to baseline and the disk
  audit is zero. It is a smoke/baseline observation, not yet a paired performance pass.
  Evidence: `C:/RedSalamander.Perf/evidence/i26-ui/integrated-perf-preadoption-v1-20260927`.
  The combined-library A1/B1/B2/A2 Default and multiline-retention runs also complete: all eight
  executions pass and all twelve unchanged comparisons retain `advice-required`, including
  same-variant controls. Exact inputs, source/binary hashes and original flags remain retained.
  Next: preserve this runnable consumer baseline, review the combined receipts, adopt the exact
  qualified pin and compare the identical integrated fixture. Missing WindowHost surface
  counters remain explicit; no zero-allocation claim follows from the fixture.

- User resumed implementation and authorized cheaper-model delegation with root review. The existing
  continuation automation is ACTIVE again. Root remains sole writer/native runner; three Sol High
  agents prepared external Skip coverage, performance fixture and evidence-curator drafts.
  Root reviewed those drafts. Follow-up agent work hit a usage limit; root continues the native lane.
  Subsequent bounded evidence/dispatch audits use an actual `gpt-6-luna` / `medium` subagent override;
  root checks every conclusion against source/receipts and remains the only writer/native runner.
- Primary feature checkouts were clean at resume; no prior candidate source drift was found. All 35
  frozen isolated input hashes matched the paused candidate. Interrupted v1 is preserved; the
  driver copies and verifies runner artifacts after each case.
- Candidate v2 correctly refused the interrupted artifact set before compiling; retain its setup
  failure. Candidate v3 uses the required full-solution Release|x64 `-Rebuild` recovery, with
  unchanged35 inputs, at `exact-bound-results-candidate-v3-20260926`. Never bypass the guard.
- V3 completed with zero warnings/errors: full Release|x64 receipt
  `69898b9bf90205dc930d66b24284b8ebb5975222542b8ae411f77a4a3d7f6173`.
  Batch run `20260926T102340Z-37620-c3820504fa5e4ee1a9f0579b5e81656e` and Recycle run
  `20260926T102415Z-37620-cdb691b468294ef6a6a98613e8954b7b` each pass 3/0/0 including setup/cleanup,
  with zero disk-audit findings. Root verified all 33/30 copied runner files and reviewed the
  exact choice/result axes. Seven Batch rounds and three malformed Recycle receipts now have
  matched red-to-green evidence. The reviewed eight-file producer/fixture patch and canonical
  contract changes are applied in primary; September 23 statements below describe the old pause.
  [Curated V3 evidence](../../TestRuns/4cb089111a23/Validation/2026-09-26_exact-bound-results/README.md)
  retains the failed setup attempts and true-red baseline without overwriting them.
- V4 freezes 36 inputs at
  `C:/RedSalamander.Perf/evidence/i26-ui/exact-bound-results-candidate-v4-20260926`.
  Only four inputs change from V3: the new common-P0 performance fixture and its registration,
  one real Recycle Skip round, and the two already committed synthetic gallery path corrections.
  Full Release build passes with zero warnings/errors, receipt
  `2b8aef878e162811721a0962ed90309b3325dc79406b3bfc8e51089edefbdb8b`. Batch passes 3/0/0;
  gallery `20260926T103536Z-7556-2f3544f9f9da4fd8983eef3b508ca6f5` passes 1/0/0.
  Recycle fails because the new test incorrectly expects `ERROR_PARTIAL_COPY` for intentional Skip;
  the canonical reducer and recorded runtime metric correctly return `S_FALSE`. Preserve the failed run.
  Root reviewed all Sol drafts, corrected the evidence curator and formatted the new fixtures.
  Common performance source SHA256 is
  `9c817b6cfe7570a3fea9c2d0b5a95da9379a5368df7612debd80628f095224fd`;
  use those identical bytes on original P0 and current source. Compilation is not a performance result.
- V5 changes only that test oracle, with separate status/byte-preservation diagnostics. Full Release
  passes with zero warnings/errors, receipt
  `e231b251d2a6a43e2e209b45aae9b2959b8b969727caeda6e693040f4d34f2da`.
  Recycle `20260927T115532Z-63916-4b2ddc2bcac8414592da8ae2957a8168` passes 3/0/0, including
  retained bytes, one exact accepted/result pair, summary/popup projection and export.
  Common fixture smoke `20260927T115612Z-63916-28b477e021bb45daa34f32b583e6bd22` passes 1/0/0.
  Both disk audits are zero; all 31/27 runner files were copied and rehashed. This is runtime fixture
  qualification, not a paired performance pass. All 36 frozen inputs match primary after LF normalization.
  [V4/V5 packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_recycle-skip-gallery/README.md)
  preserves every outcome. Its archive contract passes; V3's 115 staged payload hashes pass too.
- Root compared all 538 V4 PNGs against semantic-detail V6: 534 are RGBA-identical, four 129c/129d
  captures change, no dimensions change. All four originals were visually reviewed at 96/144 DPI.
  [Corrected gallery paths](../WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-26-corrected-paths/README.md)
  now show Delete without a destination and Rename with a new leaf in the same folder.
- Original P0 checkout `C:/RsI26P0` is at `e61e349c3e8d133e9d52a4c63ba4c49266bc4db5`, pin
  `f5c7fc7403e352425b8e82b7a4da1c042d942ede`, with only identical common fixture, opt-in registration
  and four-line test-only render-count accessor overlays. Driver/source evidence:
  `C:/RedSalamander.Perf/evidence/i26-ui/common-p0-baseline-v1-20260927`.
  V1 failed compilation because the older self-test translation unit lacked the include for its
  already implemented memory helper. V2 adds only `State.Private.h` to test registration; the
  common fixture stays byte-identical. V2 full Release build passes with zero warnings/errors,
  receipt `f510718fd9d42c265bcc6317b12012088f266f7d4776a0e2930b4827bdeca7ed`.
  Smoke `20260927T121059Z-66920-e1995a32a97a433188a3af89acfaa57f` passes 1/0/0; root rehashes
  all 27 runner files and confirms the current fixture's 144-DPI/1140x720 geometry and workload digest.
  V1's two compiler errors are retained. The matched comparison below starts after both smokes pass.
- Original-P0/current common comparison is complete: five fresh processes each in predeclared
  ABBAABBAAB order, all ten pass 1/0/0 with zero disk-audit findings. Root verifies all 270 copied
  files, identical fixture/digest/dimensions/DPI and host/attachment cleanup. Pooled median eight-draw
  wall times remain effectively level (popup 66.711/66.652 ms, Issues 66.649/66.684 ms);
  process-private memory and CPU are lower in this workload. These are synchronous wall times,
  not FPS; negative stage deltas establish only a net process-memory decrease, not its cause.
  Lightweight file checks, validators and Git ran during the sequence, so the timing observations
  do not establish an idle-only hard-budget result. No new budget is invented.
  [Matched common-workload packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_common-p0-current/README.md)
  records all samples, five-process scatter, hardware/fonts/power, both pins and baseline overlays.
  Newer detail/graph feature parity, surface allocations and final combined-pin resource gates remain open.
- Grid full-value UIA witness is qualified on exact code `40384eae` through existing CI run
  `35918640900`: Grid/Rendering/Embedded/Accessibility all pass without skips in all six native
  x64/ARM64 configurations, including ASan probes. Root reviewed the receipts and a Luna audit;
  packet/docs/plan commit `d2f2e3a` is pushed after validators, format and all 84 staged hashes pass.
  Menu run `35918649268` at `36a1c43` also passes; root confirms its src/include/Tests are unchanged
  from the already retained `f72941b` qualification. Nine ARM64 Menu desktop skips per profile
  and the unresolved menu timing decision remain separate. No duplicate native builds or pin adoption.
- A checked Luna follow-up distinguishes grid functional receipts from remaining resource gates.
  Root verifies the French pixel/copy/mutation witness, fixed 32-slot/4096-unit retention path and
  six-profile full-value UIA evidence. No dedicated over-capacity grid-cache assertion was found;
  the embedded brush/format-cache test does not establish this separate cache's lifecycle.
  The retained V11 multiline report still flags dirty CPU-composition P95 0.0152 -> 0.0182 ms
  and clean frame P50 0.3936 -> 0.4182 ms, despite roughly halved dirty frame time. Preserve those
  timing flags and the separate default-scene comparison; the accepted V11 memory tradeoff does
  not automatically accept timing degradation. Do not repeat rejected cache experiments or
  create another memory-approval question for the already accepted recorded grid cost.
- Broad current V5 Release FileOps qualification completed through the governed nonactivating
  runner with the required `-AllowAlternateVolumeTestRoot`, under
  `C:/RedSalamander.Perf/evidence/i26-ui/fileops-current-v5-full-20260927`. Run
  `20260927T122219Z-64752-80f406468c454126a4b99e52c3e9187d` reports 194/0/22, but the runner
  correctly exits 1: the new Batch Rename phase was absent from the default child-family table
  and reported `not executed`. The other 21 skips have declared capability reasons; this missing
  dispatch is not waived. Disk audit is zero and root verifies all 34 copied runner files.
- V6 adds the missing Batch Rename family and a compile-time guard that every default phase
  belongs to a dispatched family. The only candidate input change is this 26-line test registration
  repair; production is unchanged. Root checked a cheaper-model review against the dispatch path.
  A single sequential lane freezes 36 inputs, builds full Release and reruns broad FileOps, then
  builds Debug and runs broad FileOps plus the opt-in common presentation fixture. This strengthens
  the original three focused Debug cases before that build starts; both exact-result cases belong
  to the broad run and are not redundantly repeated. The prelaunch decision is retained with the driver.
  Evidence roots are `family-dispatch-v6-release-20260927` and `family-dispatch-v6-debug-20260927`
  under the external I26 evidence folder. Release now passes 195/0/21 with zero disk-audit findings:
  build `4dfbd85b1f60fc54c297a396f057f90abf9b694719c33928c95971ba39f92080`, run
  `20260927T124545Z-53864-e906e559483a4bbdb13d171bf130b1fc`; root rehashes all 41 copied files.
  Debug builds with zero warnings/errors, receipt
  `660d77a72da0359700d1de4f1f93357c8d5975e4087ef82c8990960c2470fc78`. Broad Debug run
  `20260927T130427Z-53864-274b0ec34ed6445a940a55530aad264e` passes 196/0/20; the common smoke
  `20260927T132731Z-53864-c4ed998b3267435e8d19ffa8003e3594` passes 1/0/0. Both wrapper exits
  and disk audits are zero. Root rehashes all 41/27 copied files. The sequential lane is complete;
  [V5 failure and V6 qualification](../../TestRuns/4cb089111a23/Validation/2026-09-27_default-family-dispatch/README.md)
  retain all four outcomes, source identities and 73 reviewed payloads. This is FileOps qualification,
  not the Full runner or final combined-pin gate.
- The separate grid checkout now has a reviewed candidate for its uncovered cache-bound witness:
  observed retention diagnostics, over-capacity French scroll/reattach pixels, full copy, oversized
  hidden-tail non-retention and detach cleanup. Its existing pixel fixture's borrowed models now
  outlive the host, including assertion unwinding; no crash was observed. Static skills/specs/
  dependencies/format checks pass after a retained formatting-only correction. Source and drivers
  are frozen under `grid-cache-bound-witness-20260927`; x64 Release and Debug Grid/Rendering/
  Embedded/Accessibility pass without skips. Release's archive driver failed on culture-dependent
  date re-parsing after all native suites passed; fresh source/binary-bound receipts were recovered
  without rerunning native tests, and the driver was corrected. ASan and all three ARM64
  cross-builds now pass too; all six builds have zero warnings/errors. Root verifies 88 staged
  payload hashes and that compiled/staged source differs only in CRLF normalization. A checked Luna
  review and final validators pass. Library commit `826fa56335f41c99c1f007e8879c7d124d319b74`
  is pushed; exact-commit native CI `36323665013` fails its relocated-consumer checkout on
  Windows path length in the retained UIA evidence filenames. Root is shortening those archive
  names with a reversible hash map; native results are not promoted as a green aggregate. The repair
  is pushed as `15f98822b934b5014f0d69b6eb26b1d320d20eaf`; all 90 staged payloads and static
  gates pass. Replacement CI `36324654165` passes all six native profiles and relocated consumers.
  Root verifies all 24 focused zero-skip receipts, both ASan probes and all 85 staged packet hashes;
  a checked Luna audit concurs. The source-bound `cache-bound-six-profile-ci` packet and docs
  are committed/pushed as `047414e`; this closes the cache witness's ARM64 execution gap.
  The local packet is
  `Measurements/GridTextOverflow/2026-09-27/cache-bound-witness` in the grid checkout.
  It changes no paint policy, user-facing pixels or consumer dependency; no native ARM64 pass
  is borrowed from preceding source. Paired timing and adoption gates remain open.
- Fresh current-consumer Release PR V6 (`20260927T134337Z-71884-fc01b2aecd2b4e1896d5e231c43c932c`)
  runs all 13 nonactivating entries: 822 pass, two Pester guards fail, zero skips. All 366 runner
  files are retained. The guards still expected inline decision qualification and counted the
  explicit presentation benchmark as a broad phase. Root repaired their ownership/exclusion checks,
  updated tooling inventory/coverage, and checked Luna's review: focused Pester 226/0/0 and both
  inventory validators pass. `SOURCE_INPUT_CHANGED` is the planner's changed-input reason, not
  run-time drift; all 36 before/after snapshot entries and frozen inputs are identical.
  V7 stops before native execution because the changed tooling closure invalidates its reused
  build attestation (exit 4); its setup/manifest corrections and rejection are retained.
  V8 Release rebuilds with zero warnings/errors and receipt
  `2c242338c720034372ea0d98da0d3ed28c78805aac55443846d903599469f7e2`; PR run
  `20260927T140758Z-66980-c64f831c560c47088f0a2f6ae195157c` passes 824/0/0 with all thirteen
  entries executed/promoted and no reuse. Root verifies exact build/result/snapshot binding and all
  forty frozen file hashes (canonical `Resource.h` spelling is matched case-insensitively on Windows).
  Debug also builds without warnings/errors (receipt `6080db5a7893b18c76940f32c7652c40a9d0357e074975bd62626b03cdbf4cc8`);
  run `20260927T142115Z-66980-6a26045089ba4f24878fe602e6873892` passes 824/0/0, all thirteen
  entries executed/promoted. Root rehashes 366/382 copied V8 runner files plus all 366 V6 files;
  the [64-payload guard/PR packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_tool-guard-pr/README.md)
  binds both new builds to their runs and source snapshots. The archival checker needed Windows
  extended paths for one 261-character scratch filename; its original script is retained externally,
  the corrected verifier passes, and no native run was repeated for that packaging issue.
  The tooling guards and reviewed PR packet are committed/pushed as `91f19f901`.
  V8's serial session is complete. The next source-only fixture slice is the seeded legacy-layout
  witness (133a/133b), using the private gallery owner/settings and no foreground activation.
  Root checked Luna's review against Grid's actual model-index API, then used explicit ID lookup.
  Formatting/spec inventory pass. V9 Release builds without warnings/errors, receipt
  `cd642eb461c37c78cfc5f62d7871103ab2e3cb41b826a94bb3459d5869b8eb5c`; gallery
  `20260927T143631Z-73516-0cd21f6d07d04f678233b8a91e378cdc` fails an existing terminal-facts
  footer-restoration assertion at 96 DPI. Both 133a/133b migration witnesses pass at 96/144 DPI,
  but the gallery aggregate is not qualified. All 571 runner files and the zero disk audit are
  retained under `legacy-migration-v9-release-20260927`; the serial driver correctly stops before
  Debug. Root adds failure-only rectangle/content/motion diagnostics without changing production.
  V10's diagnostic-only Release build passes without warnings/errors (receipt
  `c7d48c0ce93e1fbc317a24770c336ae231eb782893f31de82c316190e19b3db5`); gallery
  `20260927T144638Z-22308-87a6009773f5415ba4691e1ef75181ef` passes 1/0/0 with 571 copied
  files and zero disk findings. This does not fix or explain V9's intermittent failure. Root checks
  and rejects an unsupported resize-debounce hypothesis: the owner requests reduced motion.
  V11 adds eight same-HWND footer cycles, the actual disclosure handler, and both successful and
  failing restoration geometry diagnostics. It preserves every original assertion and raises the
  exact transition count from 12 to 30 per display. V11 full Debug builds with zero warnings/errors,
  receipt `9d59ea7d600b1520f7e8266ba0908c548f6714cab1a18aec86ab370aa4b2dade`; run
  `20260927T145732Z-8716-4973ac91ddc04cab80d617e75ee47dbc` passes all sixty transition observations
  and both migration witnesses, but reports 0/1/0 because one aggregate count oracle still expects
  twelve per monitor. Root repairs both checks to use one shared thirty-per-monitor constant.
  V12's scoped Debug build passes, receipt
  `316fa7dbd4f1161801d2ab1b47718e903d1e61519d791275f6aa6d71d0c17be6`, but the runner refuses
  its project-only receipt before application execution. Preserve both failures. V13 freezes the
  same forty inputs. Its full-solution Debug build passes without warnings/errors, receipt
  `503bcd43929331fd4e3bc93ed7474821282fec0ddef3c066e99be5370785b1b6`; gallery
  `20260927T151143Z-84768-53972d29b05f4ea990a0d55ef3ad9e62` passes 1/0/0, all sixty
  transition observations and both migration witnesses. All 571 copied files and zero disk findings
  are retained under `transition-guards-v13-debug-20260927`. The prepared V9/V10 Debug drivers were never launched.
  None of these later successful observations explains or waives the original V9 intermittence.
  Full additionally contains foreground entries and remains separate; PR does not replace its verdict.
- Remote dependency-submission workflow `36317958678` on `a86dd111` did not start its job:
  GitHub's check annotation reports failed account payments or a spending-limit restriction.
  This is not a product compile/test failure. Raw jobs/annotations are retained beside the
  common comparison's external driver. Local required validation continues; account billing is
  an external owner action if that remote check is required for publication.
- A managed combined-qualification checkout is prepared at
  `C:/Users/eric/.codex/worktrees/fileops-ui-qualified/DxUi`, initially detached at grid `40384eae`.
  No menu changes are merged into it and no consumer pin changes. The menu timing decision remains
  separate; use the existing grid checkout for its remaining standalone qualification first.
- Final evidence checks before the September 27 checkpoint commit: both packets' 199 staged payload
  hashes plus four original PNGs pass; full archive inventory passes for 7,918 files; spec inventory
  has zero blocking findings; staged whitespace check passes. These checks do not close the native gates.
- Subsequent family-dispatch closeout: a Luna audit confirms all 73 new packet hashes, both V6
  profiles and the retained V5 forbidden skip. Root additionally verifies all 242 staged payload
  hashes across the new and common-comparison packets. Archive policy and spec inventory pass;
  historical checkpoint text is moved to its linked history file with link rebasing and trailing
  blank-line cleanup. The current checkpoint and remaining gates stay in this WIP plan.
- Standing user authorization permitted one notified, leased Commands inspection attempt after V10.
  Run `20260927T144925Z-65508-cbfaa35ad0ab48839db54af54c056088` fails 0/1/0 at the first
  activation: foreground PID7840 is Windows `ApplicationFrameHost`, test PID70728 cannot obtain
  ownership. All 26 files and zero disk-audit findings are retained under
  `conflict-inspection-v10-release-20260927`. No immediate foreground retry is made. The trace also
  exposes the fixture continuing into owned UIA invocation after recording activation failure;
  five missing immediate exits are repaired, preserving the two existing guards. A new source
  contract fails on old code (190/1/0) and passes after repair (191/0/0). Its initial empty Pester
  selection is retained and is not a test pass. V11 Debug compiles these two additional test files.
  After the user confirms desktop availability, one protected V13 Debug attempt passes 1/0/0:
  `20260927T151422Z-87984-01e82b162f8d41ada5bd932ec31fbe5b`, same full Debug receipt as V13.
  Root checks all twelve pane/role/UIA-Space-Enter handoffs and verifies 54 copied files and zero
  disk findings. Twenty-eight unmodified harness screenshots preserve selection and returned-card
  states. The existing warning, lease and focus/cursor restoration scope guard are retained; this
  focused Debug pass does not replace Full, Release, assistive-technology or mixed-DPI acceptance.
  The historical timing question is not a new permission requirement. Menu performance advice
  remains unresolved independently of the accepted grid memory cost. Preserve the independent
  primary application (historically PID31268) and all other independently launched applications.

- The [migration/transition packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_migration-transition/README.md)
  retains all six native outcomes and the V12 setup refusal. Root and a checked Luna audit verify
  103 payload hashes and 32 unmodified published PNGs; archive policy passes for 104 files,
  spec inventory has zero blocking findings, formatting/tool inventory pass. All forty V13 frozen
  source hashes match, and primary source matches after line-ending normalization. Root corrected
  the external curator's source aliases, duplicate-storage size and README creation order; failed
  curator logs remain external, and no native run was repeated to repair packaging.
  [Inspection captures](../WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-27-conflict-inspection/README.md)
  and [legacy-layout captures](../WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-27-legacy-layout/README.md)
  state their limited scope. Known technical-row clipping still requires the qualified grid pin.
- Migration/interaction closeout is committed and pushed as `d8892fc35a15c6ee336ae7e4ac4f5bce7335eef3`;
  root verifies all 137 staged packet/image/index hashes before committing. V14 adds only the
  existing snapshot's group/viewport/painted-card/action/scroll diagnostics to the same transition
  fixture. Forty inputs are frozen at `transition-visibility-v14-release-20260927`. Full Release
  passes with zero warnings/errors, receipt
  `164674b2c8088a615fe4caddb9d941bad635dc8f3eb7e9eba73726d5665f0bf2`.
  Nonactivating gallery `20260927T153240Z-90320-cbb07f251f7e4a2c846d41bc6fb6bb2d` passes 1/0/0,
  all sixty transitions and two migration witnesses; all 571 copied files and zero disk findings
  are verified. At footer restoration, both displays report no hidden group or residual scroll,
  with painted cards inside their viewports. No further foreground attempt was made.
  [V14 diagnostic packet](../../TestRuns/4cb089111a23/Validation/2026-09-27_transition-visibility/README.md)
  retains nineteen payloads. These passes do not explain or repair the retained V9 failure.
- Next dependency-ready investigation: use the new group/viewport/card/scroll geometry if
  the unexplained V9 footer restoration recurs; do not repeat unchanged galleries for a green result.
  Root rejects an unsupported group-auto-collapse explanation: only explicit group toggle/scenario
  initialization changes group expansion. Off-viewport row omission remains a hypothesis, not a cause.
  Menu timing stays separate: the reviewed source-ablation driver splits the two-float debug-header
  delta from the menu implementation. A is clean `fb857d4`; B is clean `36a1c43`, whose qualified
  source is unchanged from the earlier candidate. Managed checkout
  `C:/Users/eric/.codex/worktrees/i26-menu-header-probe/DxUi` is at fb857d4 with exactly those
  two header fields overlaid (H). The baseline and menu checkouts remain unchanged. Root corrected
  the cheaper-model draft's stale B identity, scalar Git-output indexing, overlay indentation and
  causal overclaim before its strict read-only preflight passed. Fresh sequential x64 Release builds
  and declared A1/H1/B1/B2/H2/A2 ordinary runs are active in the sole native lane under
  `C:/RedSalamander.Perf/evidence/i26-ui/menu-header-ablation-20260927`.
  Preserve every comparison and unchanged threshold. H's different checkout drive/path and executable
  layout remain confounds; the experiment may localize differences but cannot establish causality,
  waive timing flags or qualify adoption by itself.
## Retained implementation history

[September 23 receipts and working notes](../WIP/FileOperationsUiReview_2026-09-16/execution-history-2026-09-23.md)
and [earlier execution history](../WIP/FileOperationsUiReview_2026-09-16/execution-history-through-2026-09-21.md)
retain the superseded checkpoints and all failed attempts. The checkpoint above is the current
source for status and next actions; older memory-advice and paused-state statements are historical.

## Ownership

The initial harness blockers were closed and the full redesign direction is approved. V55's
capture review passes; final Full and resource acceptance remain open above. This is the single consumer
implementation owner. Start with [approved intended behavior](../../UI/UI_FileOperationsPopup.md),
[proposal](../WIP/FileOperationsUiReview_2026-09-16/proposal.md), [gallery](../WIP/FileOperationsUiReview_2026-09-16/gallery.html)
and [historical validation](../WIP/FileOperationsUiReview_2026-09-16/validation.md). No product question remains
before P0. Completed slices and remaining qualification are tracked below.

I18 retains when confirmations appear and richer metadata/consent outcomes; I26 changes their
presentation using current typed facts. Do not wait for all of I18, invent unavailable outcomes or
change its defaults here. I3 retains stress qualification; I25 retains provider/harness residuals.
H4 retains explicitly deferred consumer ARM64/ASan adoption, which is not silently closed by I26.
DxUi owns generic controls/layout in its completed `Specs/Plans/Done/LocalizedAdaptiveLayout_2026-09-19.md`.
The consumer explicitly adopts qualified `19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1` in its
dependency lock, including the separately qualified grid/menu and multiline caret changes.
Library qualification and handoff are complete; consumer Full and paired resource gates remain
owned here. The original P0 rollback pin is retained as historical provenance:
`f5c7fc7403e352425b8e82b7a4da1c042d942ede`.

## Model routing and economical execution

Accepted 2026-09-19: execute this plan to implementation and verified closeout, using lower reasoning
effort where appropriate without reducing any quality gate. Start P0 and the first shared-layout
slice with **GPT-6 Astra / High**. This recommendation is not a guarantee of a particular token saving.

| Work slice | Model / reasoning effort |
|---|---|
| Baseline design, shared DxUi measurement/rendering, DPI/focus/UIA, conflict authority and C++ lifetime/concurrency | `gpt-6-astra` / `high` |
| Established fixtures, resource/localization updates, ordinary settings persistence and UI wiring using qualified APIs | `gpt-5.6-sol` / `high` |
| Narrow review of dangerous-action semantics, difficult unexplained regression or final high-risk diff | `gpt-6-astra` / `xhigh`, then return to the implementation setting |

Classify by actual risk rather than phase name: a settings lifetime bug or new test-harness architecture
belongs on Astra even inside a routine slice. Review the changed code and its contracts/evidence;
do not restart an unrestricted repository audit. Measure available usage and rework per slice before
claiming savings. A cheaper model can use more tokens through retries.

Use one implementation agent and one coherent code/test/spec slice at a time. Read the checkpoint,
relevant contracts and changed source sections; retain full logs as files and surface only failures
and receipt summaries. Update the checkpoint after meaningful progress, not after every tool call.
Keep required native suites, French/DPI/UIA capture checks, regression tests and paired measurements.
The baseline must precede changed rendering. Model choice never justifies skipped or weakened gates.

Recurring execution is attached to the existing task `01a0a940-8323-7cd0-bf53-1b687224c4cf`, with an
hourly continuation while unfinished. Each run resumes the next unchecked dependency-ready slice;
it performs implementation, tests and spec updates, rather than only reporting status. Continue useful
work during a run; the cadence is a recovery opportunity, not a requirement to stop every hour.
Do not start a second writer or competing build/capture lane when work is already active.

At a slice boundary, use the supported task continuation model override for the next slice when its
model differs; persist the checkpoint first and queue at most one continuation. The heartbeat itself
is not a multi-model worker pool. Do not claim a model switch based on prompt wording alone. If an
override cannot be applied, record that limitation and retain the last verified configuration.

On a genuine external blocker, preserve its exact evidence and continue independent items. Avoid
repeated expensive retries of an unchanged failure. Required unresolved gates keep the plan ACTIVE.
After all implementation and required validation pass, update domain specs in place, move completed
WIP plans to Done, repair links/indexes and disable this continuation. Domain specifications themselves
remain in their canonical folders. Do not archive this user task or close unrelated HOLD owners.

## Completed capture work

- [x] Persist harness-only documentation screenshots and the explicitly authorized notified focus exception in both repositories.
- [x] Implement reusable owned-window PNG capture and named synthetic presentation scenarios.
- [x] Capture all 24 conflict classes, related task phases, results, representative themes, widths and long text.
- [x] Close destination history, Issues empty/populated/selected/scrolled, recovery menu and live confirmation driver gaps.
- [x] Load genuine French UI resources and long French sentences; capture both physical displays at 144 and 96 DPI.
- [x] Capture real keyboard focus, hover and pressed states, with owned input and a same-HWND monitor round trip.
- [x] Inspect 437 native PNGs / 124 named scenarios and three revised French mockups; preserve per-batch provenance.
- [x] Replace the rejected horizontal comparison with mandatory Incoming-above-Existing at every width.
- [x] Update [UI contract](../../UI/UI_FileOperationsPopup.md), [selftest workflow](../../Testing/Testing_SelfTests.md) and [helper catalog](../../Core/Core_SharedHelpers.md).
- [x] Validate focused native Debug interaction / Debug static capture, 76 runner/inventory/governance checks and spec/tool inventories.

## Accepted decisions and implementation interpretation

1. Ordinary eligible name conflicts expose Keep both, Replace, Skip and Cancel task directly, with
   whole-button row wrapping. Cancel remains initial/default/Escape; the engine publishes placement.
2. Ordinary replacement uses French **Remplacer**. Specific read-only/link replacements retain
   descriptive variants; internal action IDs and operation authority do not change.
3. Throughput starts **expanded**, is collapsible, and remembers the last explicit choice using
   `fileOperations.popupThroughputExpanded`, alongside existing window placement persistence.
   The popup contract defines multi-card behavior: each existing card keeps its live choice; the
   last explicit choice is the default for new/recreated cards. No per-task settings history.
4. Incoming stays above Existing at every width. A short window scrolls the body and keeps the
   focused decision's controls reachable, with no clipped critical text.
5. Issues prioritizes the file, problem, operation, decision and result, with complete selected
   detail. Distinguish conflict, accepted decision and observed result using labels AND distinct
   icons AND semantic color; a decision is never drawn as successful execution.
6. Deliver the entire review in phases, including progress, queue/preparation, conflict, confirmation,
   outcomes, Issues, menus/history, footer and accessibility. DxUi limitations are not product limits.

`06-progress-french.png` now illustrates an explicitly collapsed graph only, not the default.
All mockups remain concepts; the accepted contract overrides their illustrative copy/metrics.

## Source and ownership map

| Work | Production owner / supporting evidence |
|---|---|
| Card, footer, graph, conflict geometry and hosted controls | `RedSalamander/FolderWindow.FileOperations.Popup.{h,cpp}`; stable action descriptors and shared bounds |
| Allowed actions, placement/defaults, immutable prompt and typed diagnostics | `FolderWindow.FileOperations.State.cpp`, `.State.Runtime.cpp`, `.State.Diagnostics.cpp`, `.State.Private.h`, `FolderWindow.FileOperationsInternal.h` in `RedSalamander/` |
| Issues grid, selection/detail, export | `RedSalamander/FolderWindow.FileOperations.IssuesPane.{h,cpp}` |
| Confirmations and speed validation | `RedSalamander/HostServices.cpp`, `RedSalamander/Ui/AlertOverlayWindow.cpp`, popup speed-dialog implementation; reuse existing host helpers |
| Settings and localization | `Common/SettingsStore.h`, `Common/Common/SettingsStore.cpp`, `Specs/SettingsStore.schema.json`, settings save/hot-reload adapters, canonical resources and all satellite resources |
| Native fixtures / layout assertions | `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp`, `.SelfTest.VisualGallery.cpp`, `.SelfTest.VisualGallerySurfaces.cpp`; existing Commands settings selftests and `Tests/SettingsSchemaTests` |
| Generic behavior | Canonical DxUi public headers, `src/Controls`, `Tests/Controls`, embedded suites; no copied library implementation in the consumer |

Source observation at planning: `ConflictActionPolicy::AddPrimaryIfAvailable` capped three controls,
and name collisions placed Overwrite before KeepBoth. The focused September 19 policy slice removes
that cap and orders eligible collision choices as approved. `TaskDiagnosticEntry` still lacks typed
accepted-decision/result correlation. These are explicit implementation changes, not layout-only
assumptions. Preserve existing bounded queues and identity validation while adding the needed facts.

## Ordered delivery and exit criteria

This section records the completed I26 delivery only after the V65 Full/Fresh, case binding, independent
root reviews and final packet named below exist and pass their exact source, run, receipt and artifact
checks. The chronological execution checkpoint above retains the original RED runs, narrow PASS runs,
retries, source revisions and their scope; this closeout does not relabel them. The acceptance matrix
below remains the coverage contract. Six accepted product decisions above remain unchanged: four
ordinary name-collision actions with Cancel default; French Remplacer with authority unchanged;
default-expanded, persistent throughput disclosure; Incoming above Existing with bounded scrolling;
typed and visually distinct Issues problem/decision/result; and the entire review across progress,
queue, conflict, confirmation, outcomes, Issues, menus/history, footer and accessibility.

### P0 — Characterization and matched baseline

- [x] Characterized the existing popup, DxUi capability boundary and ownership map in
  [P0 characterization](../WIP/FileOperationsUiReview_2026-09-16/p0-characterization.md). The retained
  ordinary-action and four-primary policy RED/PASS witnesses show the original failure and the focused
  fix, including all 24 conflict classes and authority masks. A passing candidate is not a rewrite of
  the RED run.
- [x] Bound the remaining repeat-scope checked state, French long-content fit and speed-dialog
  locale/error/DPI/focus assertions to the reviewed
  [V65 case map](../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/reviews/v65-acceptance-case-map.json)
  and actual case results in
  [binding](../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/reviews/v65-acceptance-results-root-reviewed.json). Its 16 plan rows, 62 witnesses and 39 unique native case IDs are exact evidence mappings, not proof
  that every surrounding product path ran.
- [x] Retained the unchanged-product P0 common workload and the final equal-feature A78/B19
  eight-process Embedded-profile comparison with source/build/fixture, power, DPI and sample provenance in
  [pair evidence](../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/integrated-pair/root-review.json).
  The original eight-process ProgramDatabase comparison remains separately retained under its original
  identity as historical evidence; it is not substituted or relabelled as the final Embedded pair. The
  reviewed final comparison covers this presentation and logical-resource workload, including detail/graph,
  churn and hidden/idle observations; it does not extrapolate to every product path or replace
  Full/lifetime gates.

Exit: the failure baseline, capability owner and paired workload are retained; the V65 map, actual
results and independent interpretation bind the claimed P0 facets. See the acceptance matrix for
configuration and interaction breadth.

### P1 — Measured layout and hosted controls

- [x] Adopted the separately qualified DxUi pin and kept consumer layout, localization and settings in
  their owning sources. Measured complete French action/summary text and reserved footer height before
  task content; ordinary and conflict actions, discovery Skip, destination affordance and long Unicode
  paths retain complete values, usable targets and accessible names.
- [x] Qualified repeat-scope checked/unchecked presentation and omission while loading or ineligible,
  Space/pointer/UIA parity and model-sync behavior with the actual hosted checkbox. The focused native
  tests also bind ordinary queue/options and history-menu callbacks; window-message key dispatch is
  reported as HWND input, not OS-injected physical keyboard input.
- [x] Qualified speed/validation reflow after creation, locale/DPI/error changes and constrained work
  areas, keeping input, help/error and actions distinct. French 480/640/760-DIP captures and native
  96/144-DPI geometry remain linked in the original harness galleries and final packet.

Exit: measured text and controls are reachable through real pointer, keyboard and UIA paths at the
qualified widths and physical 96/144 DPI. Native 192 DPI and external screen-reader execution retain the
explicit limits below.

### P2 — Conflict workspace, inspection and typed decisions

- [x] Qualified Keep both, Replace/Remplacer, Skip and Cancel task placement, safe default/Escape, stale
  prompt rejection, hosted Enter/Escape routing, same-host Escape and Recycle Enter/Escape behavior against the final Debug/Release
  and Full evidence. The engine remains the authority for eligibility and granted scope.
- [x] Presented one common conflict filename and state, Incoming above Existing parent/metadata rows,
  differing leaf names where meaningful, and exact-path detail/copy. A long comparison scrolls within
  the task viewport while subject/actions remain fixed; reflow preserves action identity or takes the
  safe fallback without stealing OS focus or another task's choice.
- [x] Qualified source/destination inspection through the last-active pane: exact item selection and
  focus, then return to the same unresolved card, scope and control. Both roles/panes, missing or
  filtered items, refusal, stale generations, intervening navigation/focus, multitask and teardown are
  represented in the final map and actual interaction results. Inspection neither resolves nor retargets
  the operation.
- [x] Kept immutable problem/accepted-choice facts and exact typed mutation references distinct. Inline
  Rename, Batch Rename and per-item Recycle have qualified exact carriers; cached grants do not become
  second consent. Ordinary aggregate Copy/Move completion can include descendant prompts, so without an
  exact mutation-attempt carrier its prompt result stays explicitly unknown/unrelated. No path,
  root-index or task-completion join fabricates success, and missing/evicted legacy facts remain
  unknown.
- [x] Separated problem, decision and observed result in Issues by labels, icons and semantic color,
  with no Issues-side mutation. The original Release/Debug typed-detail and policy receipts remain in
  the historical checkpoint; final current-source case bindings are in the packet.

Exit: the actual placement, subject/inspection and decision-to-result cases pass with unchanged
authority. A focused or synthetic case does not replace the whole P2 transition and real-control matrix.

### P3 — Progress, confirmation, outcome and persistence

- [x] Throughput starts expanded, remembers the last explicit choice for new/recreated cards, and
  preserves each live card's choice. Collapse keeps rate/qualified ETA/actions; hidden plots perform no
  visual work. Bounded history, settings round trip, failed save, external reload and process restart
  retain the original qualification evidence and current binding.
- [x] Distinguished discovery, transfer, verification and honest unknown totals, with complete French
  progress text and qualified waiting/running geometry. Queue, pause, stop and stream transitions retain
  their actual action targets and lifecycle semantics.
- [x] Qualified confirmation layout, locale/theme/input and resources while retaining the existing
  eligibility, consent ingress, safe default and Escape. I18 continues to own when confirmations appear
  and richer metadata/consent outcomes; I26 qualifies their presentation only.
- [x] Qualified truthful terminal summaries and independent publication, source-removal and verification
  facts, with an Issues route for partial or failed outcomes. Destination-history entries distinguish
  same-name folders by parent; actual task-button, menu, cancel and stale-result routes retain their own
  native receipts.

Exit: the acceptance matrix's task phases and terminal fixtures pass through their real controls;
persistence and monitor movement preserve task/decision state, reachability and safe recovery.

### P4 — Issues list and readable detail

- [x] Presented file/problem/operation/decision/result priority at default width, with Source above
  Destination and complete selected detail. Selected subject, exact path copy, partial-pair fallback and
  typed unknown states match the underlying exported facts.
- [x] Preserved keyboard/UIA selection, scroll and focus through append, refresh, sort, removal,
  hide/reopen and recreate. Legacy nondefault column order/width/sort survives disclosure and save;
  missing summary columns are added and invalid IDs ignored without discarding valid settings.
- [x] Qualified long French detail, legacy missing facts, cached scope, partial Move, failed replacement
  and bounded-history eviction/reintroduction in original native 96/144-DPI harness captures and final
  current results. The Problem field never borrows decision/result prose as its original fact.

Exit: selected details and exports agree, values remain accessible without default-width horizontal
scrolling, and large-fixture work remains bounded.

### P5 — Integrated qualification and closeout

- [x] Qualified distinct V65 x64 Debug/Release builds and 2/0/0 hosted-key tests in each configuration, then ran the Release Full/Fresh runner.
  Original carried V55 broad Debug FileOps 197/0/20 and V56 seven-case cluster 21/0/0 retain their
  separate run/source identities.
  [Full result](../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/full/run-all-tests-results.json)
  run `20261001T115346Z-104056-8e2cfb6a0b8049218c622dd915a2697e` reports `2211 passed / 0 failed / 46 skipped`, a clean sandbox and source snapshot
  `f2e0a18a04fbc8a95368adc0c946bafb5ebccdf68d81b4a476d0fefbd323593c`; its post-Full receipt and native Commands/FileOps results are bound by
  [acceptance binding](../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/reviews/v65-acceptance-results-root-reviewed.json)
  and
  [independent root review](../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/reviews/v65-acceptance-binding-root-review.json). Historical failed Full runs remain failed in the checkpoint and packet.
- [x] Reviewed the final eight-process A78/B19 Embedded-profile pair separately from the 39-ID behavioral
  binding. Its source/build/fixture/profile provenance and root review are in
  [pair evidence](../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/integrated-pair/root-review.json).
  The original eight-process ProgramDatabase comparison remains separately retained under its original
  identity as historical evidence; it is not substituted or relabelled as the final Embedded pair. The
  final comparison covers the named presentation/resource fixture and does not assert equal-feature
  performance for unmeasured paths. Retained physical 96/144-DPI captures include same-HWND 144→96→144
  and protected focus; the 562-static/ten-protected V55 gallery and 512 exact reviewed image hashes retain
  original bytes and provenance, not V65 recaptures.
- [x] Reconciled UI/FileOperations/Settings contracts, capability wording, schema, tests, gallery and
  dependency provenance against the accepted behavior.
  [Final V65 packet](../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/curation-manifest.json)
  records the final manifest, source review/carry, builds, case map/binding, Full, galleries, pair
  interpretation, legacy failures and archive hashes.
- [x] Updated the plan/index/link state only after the final packet and independent root review passed.
  I18 confirmation authority, H4 consumer ARM64/ASan adoption, required native 192-DPI execution when
  available, external AT/screen-reader execution and other separately owned platform gates remain
  explicitly outside this I26 completion claim. Physical keyboard/pointer injection is not inferred from
  HWND message tests.

Exit: `4a3a69952f7dacd732c2ddfc29bf37808802ac2741334d59bc01394a0bdf303e` identifies the reviewed packet, with `d7535cb24e0c3e21088c000ba101e37a4721537a98009a18495b4972216ed857`
and `46b8977a90f0ad23095d81f6d19bf68447f874f1495ba913d13631f49e2da4de` binding its actual V65 cases and Full run. This closeout is
effective only for the qualified scope recorded above; separately owned or unavailable gates remain
named limitations rather than silent passes.

## Acceptance matrix

| Dimension | Required cases and pass conditions |
|---|---|
| Scenario coverage | Reuse all 124 named scenarios / 437 historical captures as the review map; all 24 conflict classes, all task phases/results and every adjacent surface. Add scenarios for the new disclosure, four-action policy, typed decisions and Issues detail. Inventory accounts for unchanged, replaced and new cases. |
| Complete interaction coverage | Reconcile all current popup action kinds and menu/dialog/window routes against M01–M32 in the September 20 review. For each, execute initial state → trigger → visible and underlying outcome → cancel/failure/return with ownership restoration. Static capture or a helper return is not a journey pass. |
| Conflict subject and inspection | One common name and state, parent-only role rows, explicit differing names/fallback, no hidden duplicate UIA providers. Pending source/destination inspection selects the exact item in the last-active pane without consent; returning restores the same unresolved context. |
| Locale / width | English and genuine French resources; French long sentences, filename, parent paths and disabled reasons at 480/640/760 DIP, plus a wide layout. No clipped critical labels, overlapping regions, lost root/extension or zero hit targets. |
| DPI / short height | Native 96/144 DPI on the two physical monitors, actual work-area bounds, same HWND 144->96->144. Required 192-DPI native qualification when an environment is available; synthetic 192-DPI layout tests do not prove physical presentation. Report that gate pending until run. |
| State / theme | Light, dark, rainbow light/dark, high contrast; normal/reduced motion. Rest/hover/pressed/focused/disabled/checked. Correct distinct status icons/words even without color; readable normal text and disabled explanations. |
| Input / UIA | Tab/Shift+Tab, Space checkbox/disclosure, Enter/Escape defaults, pointer release outside, scroll/resize/DPI during focus, vanished action fallback, stale prompt rejection. Names/full values, Toggle/ExpandCollapse, selection, reading order and screen bounds agree with display. Screen-reader review supplements automation; record environment and limitations. |
| Persistence / lifecycle | Default-expanded, explicit collapse/expand, canonical round trip/restart, external reload and failed save; hide/reopen, footer-only and expanded placement, minimized/hidden idle, device loss/fallback, multiple tasks and prompt arrival without focus theft. |
| Outcomes / diagnostics | Pending choice vs accepted decision vs verified outcome; partial/failed/unknown, cached scope and missing legacy facts. Never derive success/authorization from text, color, percentage or an accepted choice. |
| Scale / resources | Existing bounded task/history/diagnostic/grid limits, 1 and 4 transfer streams, large Issues data, update/scroll/reflow, repeated open/close and hidden intervals. No unbounded text/layout caches, timers, HWNDs, surfaces or row work. |

Capture input follows the existing opt-in warning/lease, process/foreground ownership checks and
cursor/focus restoration. Static captures remain non-activating. No desktop screenshot substitute.

## Paired performance and command routing

Measure before implementation and after each affected rendering phase, serially on the same machine,
configuration, pin, fixture, DPI and power profile. If the harness changes, run its final version on
both implementations. Record p50/p95/p99 latency, presented frame pacing, CPU/preparation/render cost,
allocations, private/working-set and surface/cache peaks, handle/thread counts and idle wakeups.
Retain all repetitions, not just a favorable run; the common DxUi comparison uses five-round medians.
Apply existing library and [consumer performance requirements](../../Testing/Testing_PerformanceValidation.md)
without loosening thresholds. Confirmed degradation needs measured optimization/scope/deferral advice.
No native runtime evidence is inferred from a cross-build.

Use supported commands, selected from `Tools/README.md`. On this qualification host,
use marked `V:\RedSalamander.Perf` as primary and marked C: as the explicit alternate;
the required independent-volume case must not be launched without its opt-in switch:

```powershell
$env:RSBuildEnableTests = 'true'
$env:REDSALAMANDER_TEST_ALTERNATE_ROOT = 'C:\RedSalamander.Perf'
.\Tools\Run-AllTests.ps1 -Suite FileOps -Platform x64 -Configuration Debug -TestRoot V:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_VisualGallery -Platform x64 -Configuration Debug -TestRoot V:\RedSalamander.Perf -TimeoutMultiplier 3
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_VisualGalleryInteraction -Platform x64 -Configuration Debug -TestRoot V:\RedSalamander.Perf -TimeoutMultiplier 3
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Platform x64 -Configuration Release -TestRoot V:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

Add focused Commands settings/schema cases through the runner after P0 discovers their exact names;
do not invent passing case names or replace runtime assertions with source-shape tests. Capture
lanes are opt-in and not automatically covered by Full. Do not stop unrelated running applications
to acquire a build output; report an output lock and use supported isolated outputs where available.

## Safety and closeout

Default capture does not activate windows. The exact interaction case uses the existing warning and runner lease, owned-window/foreground checks and cursor/focus restoration. Abort on wrong process, unsafe paths, lost ownership or missing renderer content. Never substitute desktop control or manually edited application screenshots.

Product decisions are resolved; this I26 owner is complete within the qualified scope above.
The capture record and reproducible commands remain in the review folder. A documentation-only plan
update changes no compiled inputs and requires no new native performance claim or application capture.
