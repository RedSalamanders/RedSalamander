# I27 File Operations safety review remediation

Status: DONE — delivered implementation, archived on 2026-10-05 following the user's explicit approval to separate delivery from remaining qualification. Historical implementation owner: Codex. This is not a green Full, current-pin performance, release-readiness or universal filesystem-safety claim.

## Delivery closeout — 2026-10-05

The implementation and reviewed integration are published on `master` at
`cfe94556c94e971a753d692c55793eddda172bd7`. The delivered task commit is
`02841dd4b62022a5c3d04f72efa18a3dea84a2bb`; integration commit
`6e54d8679fa110bb2b8364ac306404788e641638` preserves incoming master changes and
its DxUi `49a99881cecd37014dc1a92075fd1a610fdfded2` / API revision 3 pin.
The integrated test-enabled Embedded x64 Debug solution build passed with zero
warnings/errors. The two affected Commands fixtures passed 2/0/0 on that build;
focused merge tooling passed 255/0/0. Their exact source/build/run identities and
scope limitations remain in the historical checkpoint below.

The user asked to converge and merge before broader testing, then approved moving
the delivered implementation to Done with unfinished qualification in a separate
WIP plan. [I29: File Operations post-merge qualification](../WIP/FileOperations_PostMergeQualification_2026-10-05.md)
is now the sole continuation owner for this plan's residual obligations:

| Historical unfinished gate | Active owner |
|---|---|
| Integrated x64 Debug/Release builds, regressions and Full verdicts | I29 Q1 |
| Source-bound final Release comparison and reviewed evidence publication | I29 Q2 |
| Final normative/coverage/consistency reconciliation and governance closeout | I29 Q3 |

No obligation is waived or counted as passed by this split. The failed Full runs,
adverse measurements, skipped mounted routes and historical API-2 receipts remain
unaltered. The independent DxUi I27, I28, I18, I3, I25 and H4 plans retain their
ownership and acceptance boundaries. The exact new Done path and its existing
positive/negative admission test authorize this file only; existing Done history
and the already-published harness gallery remain unchanged.

Delivery-split validation passed 37 SpecInformationArchitecture and 18
DocumentationDriftContracts checks, with no failures or skips. Root independently
verified all 77 local Markdown links across the three changed documents; spec/tool
inventories have zero blocking findings and protected Done history is valid.
The 65 informational spec findings remain. These documentation checks do not
satisfy the Full/performance/evidence gates transferred to I29.

All checkpoints, instructions and unchecked qualification rows below are
**historical execution records**. They do not authorize resuming or rewriting
this completed file. Continue from I29, using canonical domain specs as current
authority. The implementation checkout and source branch remain available; no
additional Full/native/build session is started by this delivery closeout.

## Historical execution record

Continuation authority: the user's 2026-10-02 request extends this same indexed plan to the
verified follow-up below, authorizes implementation and cheaper-agent delegation with independent
review. Current follow-up checkout: `C:\RsI27Followup`, branch `codex/fileops-safety-followup`,
originally starting at `0e139fb889346827ba73b776a3a558b65e7287f1` (then-current `master`), then rebased onto
`fbe113a0f747f89af32d6fb337e4e67b9f5bdd6b` for the completed dependency compatibility fix. Unrelated dependency-recovery
changes in `Z:\src\RedSalamander` are owned elsewhere and must remain untouched.

### Historical execution checkpoint — 2026-10-05 16:01 UTC — integrated, ready to publish

The user's delivery order is now applied: no additional pre-merge Full cycle.
Task commit `02841dd4b62022a5c3d04f72efa18a3dea84a2bb` is preserved, and integration
commit `6e54d8679fa110bb2b8364ac306404788e641638` merges current master
`272aaf38e145cd8cd3e02235198b1bf7290fbdc7` while retaining DxUi
`49a99881cecd37014dc1a92075fd1a610fdfded2` / API revision 3. Root and cheaper
independent review checked the four conflict resolutions and repaired the
automatic-merge debug-message collision; both named snapshot handlers survive.

The integrated whole-solution x64 Debug build passed with **zero warnings/errors**,
test-enabled Embedded, receipt
`0bfec155d5d67fae48227cd791b1edbca1e7317c5e9e2887e3f63a4059f01f17`, source
`565bac37210db3e3c0df6b66b5867d7dc831527d28a4302ad6e1743cf2ed7a01`.
The two affected Commands fixtures passed **2/0/0**, repeat/timeout scale one, in
`20261005T155340Z-46636-1f9c0561c42245cbb9d571b1ad07655d`: throughput/path-detail
clipboard interaction and forced D2D-target fallback. Root and independent review
verified all six original byte bindings, matching archived workspace source and
the complete validated no-follow archive; manifest SHA-256
`5ce1b427218e3a70a66b9d5e06fece801c145f98c7d544b4ebea19cab6f2350e`.
The packet and immutable build receipt are retained under the existing external
`final-run-review` root. This establishes exactly these two integrated Debug cases;
it does not establish Full, repeated reliability, current-pin performance/resource
equivalence, a clean historical disk audit, or the cause of old failures.

Focused merge tooling passed **255/0/0**: SanitizedEnvironment 10, ToolInventory 17,
SpecInformationArchitecture 37, TestHarnessSourceContracts 191. Tool/spec inventory
validation has no blocking findings; 65 informational spec findings and the
historical 265-entry disk audit remain distinct. Formatting/whitespace and the
integrated tree are clean. The only subsequent changes are this checkpoint and its
WIP index description; they change no compiled inputs. Full Debug/Release and
source-bound final performance/evidence/spec closeout are still open for proper
post-merge qualification. Both safety and DxUi owner plans remain WIP.

Next: ordinary push/fast-forward publication to master, preserving other owners'
checkout state and plans. Do not start another Full before publication. After
integration is available for the user's other tests, resume only the explicitly
remaining qualification/closeout items; do not rewrite failed originals or claim
Done from this bounded gate. No build/native/archive process is active.

### Historical execution checkpoint — 2026-10-05 15:47 UTC — converge for merge

The user explicitly redirected this work: "why another round of full validation ?
we need to converge and merge for other proper tests". No Full/native/build process
is currently running. Do not start another pre-integration Full cycle. Freeze the
reviewed repairs, preserve the current work, integrate with current master, then
run the focused affected cases on the integrated build. Outstanding Full Debug/
Release and final qualification remain recorded for proper post-merge testing;
this changes delivery order, not the historical verdicts or their claimed scope.
Keep this plan WIP while its required qualification remains open.

Current master is `272aaf38e145cd8cd3e02235198b1bf7290fbdc7` and has meanwhile
adopted DxUi `49a99881cecd37014dc1a92075fd1a610fdfded2`, API revision 3.
The safety branch and all performance/native receipts below used API revision 2
and the recorded `19ca44d...` pin. Do not rebuild that old pin again as a substitute
for checking the actual integration. Preserve master's chosen pin and its separate
active DxUi follow-up; the existing API-2 receipts do not qualify API-3 adoption.
Inspect and resolve integration compatibility through bounded review/checks.

The last two fixture repairs have root and independent Astra full-diff review;
source contracts pass 191/0/0, spec information architecture 37/0/0, and pinned
format/whitespace checks pass. The three-entry consistency ledger update has root
and cheaper independent review, retaining all 62 entries' historical fields.
The new C++ repairs still need the integrated build and focused native checks.
All prior source-bound originals, failed broad verdicts and accepted performance
scope remain retained. The merge preparation must preserve other owners' plans,
independently launched applications and unrelated checkout state. No claim of Done,
green Full, current-binary performance equivalence or public merge is made yet.

The task work is now preserved in commit
`02841dd4b62022a5c3d04f72efa18a3dea84a2bb` (1,519 changed files, including retained
original evidence and the deterministic gallery). A normal merge of current
`origin/master` is prepared in this checkout. Four overlapping document/test
conflicts are resolved by preserving both the task's strict child-environment
semantics and master's first-entry failure-report routing, contained-job crash
behavior and default-toolset overlay. Master's API-3 renames and registered-message
integration are retained; the active source has no legacy API-name residue.
Independent merge review found one concrete automatic-merge defect: the task's
menu-attachment snapshot and master's new ViewerSpace draw-item snapshot both
used `WM_APP + 0x636`. The task-owned named message now uses the previously free
`0x637`; all callers use that shared symbol, and both handlers remain intact.
The two existing I27 index rows name different plans: this safety plan and the
separate DxUi follow-up. Neither plan's qualification is implicitly closed by
integration. Next: review and commit the integration, build the integrated x64
Debug product, check only the throughput-disclosure and D2D-fallback fixtures,
then publish the merge with their actual outcomes and Full qualification open.

### Historical execution checkpoint — 2026-10-05 15:25 UTC

Whole-solution x64 Debug passed 0 warnings/errors: source
`074caab6080370f7136df87c92abc38a30d325f62a9e8260ba782bac39bcfadf`, receipt
`dcc160009a884abf58ed60cdf86b2425ac164ee099d33d12e74b892a6ce37756`.
The exact throughput case passed 1/0/0 at
`20261005T150647Z-86936-41082d2fe82c48c2a533e3c29ee7fbf5`: all three actual
system clipboard routes, complete Unicode/UIA text, retained-format restoration,
protected focus and forty later open/close cycles. Root and a cheaper independent
review verified its six originals and complete validated no-follow archive,
manifest SHA-256 `a9dd3e0620310f7146bffc63adb695a3ccedcd2e170a41296f474b7e54b4911d`.
This isolated success does not establish repeated reliability or explain old failures.

The broad family then closed FAILED 63/2/0 (repeat/timeout scale both one) at
`20261005T151229Z-93344-397b0800eb484f7584012682438c9b69`. The throughput
case passed its first two actual reads but refused the THIRD copy guard; its
original quiet/generation result was not logged. Later cleanup owner zero and
sequence 17682/17389 cannot identify that boundary or its publisher. Observed
reader processes were RadeonSoftware and Explorer; reading does not prove writing.
The D2D-target-failure fixture also failed its blocked-metadata fallback assertion.
Root verified all six original bindings and the complete validated archive;
manifest SHA-256 `b3adabf29941ae87d0155ea474f5a457654713c3b869798500dba79879248291`.
The original native verdict remains FAILED, unclassified, with no retry/quarantine.

Root and independent Astra source review found a reachable fixture ordering gap:
popup creation invalidates for later painting, and worker metadata pause can precede
the first render that consumes the D2D injection. The fixture now drives the same
ordinary owned-window synchronous redraw before both normal/failure assertions,
retaining the fallback assertion, metadata pause and unconditional cleanup.
New failure diagnostics report snapshot, readiness, host, child and pause state.
No production fallback defect or historical run cause is proven by that source finding.
The third clipboard guard now logs non-pumping snapshots after rejected read-only
keys and before/after its quiet wait, separates quiet/generation observations, and
rechecks the admitted generation after logging. It adds no wait, copy retry,
generation adoption or intervening-content overwrite. Source review found no
unexpected Return/Tab clipboard route. All existing deadlines/ownership gates stay.

These further fixture changes are UNBUILT/UNQUALIFIED. A three-entry consistency
ledger update adds current owning anchors and append-only scoped notes, preserving
all historical fields. Retained-before/review files are under external
`resume-20261005T060429Z/broad-fixture-paint-boundary-*`; the staged tree,
production/library/pin scope and earlier performance originals remain unchanged.
Next: independent complete-diff review, format/focused contracts, fresh whole Debug
receipt, the broad family, then genuine Fresh Full Debug/Release and evidence/spec
closeout. Do not move this plan to Done before those required gates pass.

### Historical execution checkpoint — 2026-10-05 14:50 UTC

The corrected whole-solution x64 Debug build passed (0 warnings/errors): receipt
`0522280ddfcf1baae901e27c752a5c6a04e6428cd15b8547a860c15a64f33fc5`, source
`5199b3e14e8e7fea444b6842d33197ea56d763e0ec6fef9f76000c7522f6a04d`.
The new queue regression closed 1/0/0 at
`20261005T143246Z-71452-5ab193c62e2d4094b582505aa3ef42ef`. Its finite producer
yielded after three work messages with work still pending (40,110 us), then the
fresh post-cleanup fixture processed one message with none pending (13,857 us).
These cooperative observations do not promise a hard ten-ms dispatch ceiling.
Six original packet byte bindings and the complete no-follow archive validated.

The real throughput fixture closed FAILED, 0/1/0, at
`20261005T143726Z-96172-26f31a7e9b0d44cbb07f48fefc867c0c`, with the same receipt
and source. The first complete 6,155-character copy and actual system read passed
by 575 ms from path-detail opening; the failure is before the SECOND suffix-copy
route, at the combined quiet/generation guard. Cleanup later observed owner zero
and sequence 16767/16474, after another quiet wait. That later snapshot cannot
identify the guard's original failure or prove an external publisher. The exact
failed verdict, six original byte bindings and validated archive remain retained;
manifest SHA-256 `eab347982a83c00f699af2752db3b368ac34596c6a22bd51342bafe9729922b4`.
An external dated correction fixes the earlier root-review shorthand that no
product copy had occurred; no original native verdict is rewritten.

Root and cheaper read-only review found no concrete clipboard accounting/lifetime
defect. The next diagnostic separates quiet and generation results, logs metadata
immediately after that wait (including failures), and adds owner/sequence/open
window information around navigation, screenshot and selection. It neither
adopts a newer generation nor retries a product copy or overwrites intervening
contents. Root also found that the existing full-path-button read still allowed
the debug fallback; this same named fixture now requires actual system readback
on that final route as well, retaining the single button activation and generation
guards. Legacy and Shortcuts clipboard policies stay unchanged.
Independent review identified that diagnostic writes widen the interval after the
sampled ownership check. The final pre-copy guard now rechecks the same admitted
generation after logging; this adds no wait, pump, retry or generation adoption
and does not make clipboard admission atomic.
These further fixture changes are UNBUILT/UNQUALIFIED. Retained-before/candidates
are under the external `resume-20261005T060429Z/path-detail-clipboard-boundary-*`
packet. Next: full-diff independent review, format/whitespace check, fresh Debug
receipt, this exact case, then affected broad families and Fresh Full Debug/Release.
The staged tree and frozen production/library/pin scope are unchanged.

### Historical execution checkpoint — 2026-10-05 14:25 UTC

The first fresh build of this queue regression closed FAILED at
`debug-whole-build-20261005T141957Z-77616` (0 warnings, 1 error): the pinned WIL
rejects a named lvalue lambda passed directly to `scope_exit` (C2338). Its original
log and wrapper record are retained. Root corrected only the guard construction
to a direct noexcept lambda invoking the same stop/drain callback; cleanup order
and assertions are unchanged. No native run used the failed build. The complete
five-file review, pinned format checks and focused tooling had already passed
(inventory 7/0/0, source contracts 191/0/0); those do not establish compilation.
The corrected source and this checkpoint require a fresh whole-solution receipt
before the reviewed exact queue regression and actual clipboard/focus checks.
Retained-before/candidate files and compiler failure bindings are kept under the
external `resume-20261005T060429Z/commands-pump-wil-fix-*` packet.

### Historical execution checkpoint — 2026-10-05 14:07 UTC

The further stage diagnostic compiled cleanly in whole-solution x64 Debug
Embedded/test-enabled form: source
`94af857674b65c762250024998c7927741c64bf82e9c07d3b30c1c84b5214ba0`, receipt
`2e3023b31f7ddf7b756e95daa42c8c0b9ebe811d86c06732b699f308c7ce8e1a`.
Its exact case closed **FAILED**, 0/1/0, at
`20261005T135000Z-87392-2bf45ca6f9e84b5fa0f7f09adf4d7598`: actual clipboard
read error 5 with exact debug fallback. This time foreground class/process,
native focus, active HWND and last-input ticks stayed on the same owned popup
throughout the recorded stages. No Ghost or external-input cause is established.
All six original packet byte bindings were root-verified; exact C archive closed
and validated, manifest SHA-256
`b0fe61ed00a53a73fe006f2ff3b496721b25125acb4626d3f4357cb5020e2ea8`.

The trace measured about 15 seconds in `redraw`, which includes the Commands
512-message pump, and 9,101 ms inside the unchanged three-second clipboard-read
wait. Original per-paint rows are roughly 33–37 ms. They do not attribute fifteen
seconds to three paints or establish a library rendering cause. Root and cheaper
review identified that sustained dispatch can consume a wait's deadline before
another read attempt. Exact read-attempt count was not measured. The pinned
multiline selection renderer also warrants separate investigation; no library
change, pin adoption or rendering-cause qualification follows from these timings.

Root has written a narrow Commands pump adapter: reuse canonical TestSupport
dispatch one message at a time, retaining the 512 count limit and adding a ten-ms
cooperative budget between calls. At least one pending message is serviced;
already-started or internally dispatched sent callbacks cannot be preempted.
Clipboard retry order/deadlines, product copy count and every ownership assertion
remain unchanged. A native Settings-family regression uses finite replenished
work and real cancellation messages, verifies caller control before exhaustion,
drains only its owned private messages on every exit and tests a fresh fixture.
The helper catalog and normative self-test contract are updated. Retained-before
and reviewed candidates are under
`resume-20261005T060429Z/commands-pump-budget-20261005T135917Z-e906d669/`.

These latest changes are **UNBUILT/UNQUALIFIED**. Next: independent complete-diff
review, format/inventory checks, fresh Debug whole-solution receipt, the new native
regression and the real throughput case, then the affected broad families and
Fresh Full Debug/Release. No build/native/archive is active at this checkpoint.
The staged tree is unchanged. All earlier failed runs and performance/build
source scopes remain; no pass repairs their verdicts. Capped publication,
normative/governance closeout and the same-plan Done/index move remain open.
Desktop reservation and existing warning/lease/restoration safeguards persist;
H4/I18/I28 remain with their separate owners.

### Historical execution checkpoint — 2026-10-05 13:45 UTC

The stage diagnostic was compiled in a whole-solution x64 Debug Embedded/test-enabled
build with zero warnings/errors: source
`d683fda0e71c040c8a2f3d0b90d068c66252a2776aa4f30336420d2de9277edf`, receipt
`0a07ec5bb78ffe4456040d273ad0f0c5f26c8d84da912a93bb358ce4596621a5`.
The exact throughput-disclosure diagnostic closed **FAILED**, 0/1/0, at
`20261005T132322Z-67024-cd3cd77a9139482d95f4b4eea5dd0f95`. This time the failure
occurred before close: the actual clipboard read returned `OpenClipboard` error 5;
the debug fallback contained the expected 6,155 characters but was correctly rejected
as a system-copy witness. Clipboard sequence/owner still matched. Native focus was null
at failure; the initial opening trace identified the owned popup and unchanged input ticks.
No later foreground identity was captured, so no Ghost, external-input, clipboard-lock
or product-root-cause classification is established. Original C archive closed and
validated, all six packet bindings were checked, and the archive manifest SHA-256 is
`720eb749317e1671ded54ab9f5a8aa8a6045be181427c9629911de8021a5ee0c`.

Root has added elapsed-time and read-only stage traces around full-document UIA,
the existing clipboard quiet point, the single product copy and actual read.
Retained-before and formatted candidate bytes are in
`resume-20261005T060429Z/path-detail-copy-stages-20261005T134047Z-bc856066/`.
This further diagnostic is **UNBUILT/UNQUALIFIED**. Format/whitespace checks pass;
cheaper independent review, a fresh Debug solution receipt and a changed diagnostic
run are next. No extra copy, focus acquisition, message pump, timeout change, assertion
removal or production change has been added. No build/native/archive is active at
this checkpoint; the staged tree and other owners' work remain unchanged.

The actual Commands-family result remains 64/1/0; the Viewers and ViewerSpace prefix
passes below retain their original source scope. Full Debug/Release, capped evidence,
normative/governance reconciliation and the three closeout boxes remain open. The
user's desktop reservation persists until tests finish, subject to the existing
warning, lease and restoration guards. H4/I18/I28 remain separate.

### Historical execution checkpoint — 2026-10-05 13:10 UTC

Whole-solution x64 Debug Embedded/test-enabled build passed with zero warnings/errors,
source `208c5cd5498336088fa4b5835c421fdd00e38c4df715edeaea52d1bc122d503f`,
receipt `6d441747791701acff4a11cab71e2de0607bab41bc9d34bccb0c6ab3f38ab9b0`.
Its immutable original receipt is retained beside the reviewed four-file fixture bundle.
The same-process Viewers prefix passed 29/0/0 at
`20261005T121945Z-53364-9fbc4a53667d490aaa47cd6e7ff1485b`; ViewerSpace passed
20/0/1 at `20261005T122639Z-58472-13f8c746c6c54972be3b56d95469cac5`, including
hit-grid settlement. Its single declared skip is the optional 20,000-item fixture.
Both original no-follow archives closed and passed their exact validators; root checked
all six packet byte bindings per run. These focused passes do not repair the old Full verdict.

The File Operations Commands family closed **FAILED**, 64/1/0, at
`20261005T125426Z-88396-6298ba5832824414819143ef1c663481`. The real clipboard
read succeeded; the subsequent first path-detail close failed foreground preservation.
Original trace records expected popup `0x24A1D68`, actual foreground `0x505CA`, and
native focus still on the popup. It does not identify that other window or the exact
transfer stage. Root and cheaper static review established no native cause or environment
classification. Its complete validated C archive, original failure and six verified byte
bindings remain; archive manifest SHA-256 is
`ce2a38bc85352e2c29d3d363cdc79623a4437900277df0faae61ca3ecda20857`.

Root has added a scoped fixture diagnostic in `Commands.SelfTest.FileOps.cpp`: existing
trace stages now capture foreground process/thread/class, active window and known last-input
ticks before/after clipboard restoration, Tab, Escape and redraw. Added ownership
preconditions stop further keys after a lost foreground; no reacquisition, extra product
copy, timing increase or original assertion removal is introduced. Retained-before and
formatted v2 bytes are under
`resume-20261005T060429Z/path-detail-focus-diagnostic-20261005T130519Z-5cd7c66e/`.
This latest diagnostic is **UNBUILT/UNQUALIFIED** and requires independent review, a fresh
Debug solution receipt, then bounded diagnosis and the affected broad family. No runner,
build or archive remains active at this checkpoint. The staged tree remains unchanged.
Full Debug/Release, final evidence/spec/governance and the three closeout boxes remain open.
The agreed desktop reservation and all ownership guards remain in force; H4/I18/I28 are separate.

### Historical execution checkpoint — 2026-10-05 12:05 UTC

Fresh Full Debug `20261005T093437Z-90912-a440bb9f89f046bda245891a1473a39f`
closed **FAILED**: all 20 entries executed, 17 promoted, native 2038 passed / 5 failed /
246 skipped. The three Commands failures are Viewers reordered-column clipboard copy,
File Operations throughput path-detail clipboard copy, and ViewerSpace hit-grid settlement.
F7 separately refused an observed foreground transfer from Windows Terminal to Codex;
ViewerPE's first native child passed before the parent's changed-input restoration guard
refused further children. No native root cause or environmental classification is inferred
for the three Commands failures, and neither desktop safeguard is waived.

All C, D and C-validation no-follow archives closed and passed their exact validators.
The failed-run receipt keeper completed 37 metadata checks with zero issues and retained
both immutable solution receipt epochs `e1ff74e1…` and `d54d99f8…`; the repository verdict
remains FAILED. Its report SHA-256 is
`54c74b1ab5a49748fe0355b6e074fa10b3681b2689b99cda9bade2966d818d23`.
No build or source adoption preceded that keeper's closure.

The guarded Local route candidate at source `3a1e1d3…` was compiled in whole-solution
Debug and Release, with zero warnings/errors. Its unchanged R0e ceiling passed in all five
Release processes; focused F7 passed its complete ordinary fixture; five fresh paired v2
processes per variant retain all original frame/resource rows and adverse observations.
These bounded measurements and the 22 externally prepared evidence packages do not
replace the failed Full gate, qualify unavailable mounted behavior, or attest later source.

Root adopted four independently reviewed Commands fixture changes and two scoped normative
updates after verifying every retained-before and proposed byte hash. Viewers copy now
requires successful clear/quiet setup and rechecks its target after message pumping. The
two affected copy witnesses require actual system reads and retain transport/fallback and
ownership diagnostics. Clipboard setup errors are captured before pumping. ViewerSpace
failure text uses the post-wait snapshot, and standalone probes reuse process-scoped window
enumeration rather than class-only lookup. Product copy attempts, native deadlines and
desktop guards remain unchanged. The cheaper read-only review's error-capture and overly
broad documentation findings were corrected and checked by root. This latest fixture bundle
is **UNBUILT/UNQUALIFIED**; all original failures and earlier reviewed proposals are retained
under `resume-20261005T060429Z/debug-command-diagnostics-draft-20261005T112748Z-7a4083eb/`.

Next: fresh whole-solution Debug build, then the affected cases with their actual same-process
predecessor clusters (Viewers prefix, File Operations family and ViewerSpace prefix). Resolve
demonstrated failures without raising limits, then complete genuine Fresh Full Debug and
Release and exact evidence/normative/governance closeout before moving this same plan to Done.
The desktop reservation persists; no repeated availability question is needed. The three
overall closeout gates and separate H4/I18/I28 ownership remain unchanged.

Baseline: `e860022a0` (I26 completion), branch `codex/fileops-review-remediation`. The historical [I26 plan](../Done/UI_FileOperationsReview_2026-09-16.md) stays frozen. Review evidence: `C:\RedSalamander.Perf\evidence\i26-code-review-20261001-e860022a0\Review.md`.

Current authority: [File Operations](../../FileSystem/FileSystem_FileOperations.md), [popup](../../UI/UI_FileOperationsPopup.md), [FolderView](../../UI/UI_FolderView.md), [shared grid/alerts](../../UI/UI_DxUiSharedGrid.md), [shared helpers](../../Core/Core_SharedHelpers.md), [selftests](../../Testing/Testing_SelfTests.md), and [performance validation](../../Testing/Testing_PerformanceValidation.md). This is a bounded safety follow-up; I18's proposed confirmation preferences and H4's deferred platform qualification remain with their existing owners.

## Historical implementation and qualification

### Historical execution checkpoint — resumed on 2026-10-05

Root verified all 103 compact source bindings from the paused recovery packet against the current working files: zero unexpected changes, HEAD `fbe113a0f747f89af32d6fb337e4e67b9f5bdd6b`, unchanged staged tree `2f2aab8ac16426b0af8da6e8745f0b89a2812d8b`. The resumed evidence is `final-run-review/resume-20261005T060429Z/`. The completed read-only Sol reviewers are resumed for bounded ownership/test and admission-path reviews; root remains the only writer/build/native owner and independently checks their findings.

The owner-local shutdown slice is now compiled and focused-qualified. Whole-solution x64 Release Embedded/test-enabled build passed with zero warnings/errors, receipt `5dee74305930d29dc5b8960810be16e93135e2ac9aa7488685dff08a48525375`, source `25c6c6f9be8f30da06b4b5929d2e99f121ebb466a838eec2b60d11e7e5e14cb9`. The linked settings case `20261005T072143Z-56400-7f023bf1d56f4542b4a5e7acfb6d5b36` passed 1/0/0 with complete original C archive closure: positive production-setter/real-shutdown child persisted the admitted density generation in 519198 us; negative idle-owner shutdown returned in 16 us, leaving the unrelated generation in progress. The linked theme case `20261005T072937Z-54628-45405134296142298ed031f225e4c64f` also passed 1/0/0; its original Full failure remains unclassified history. No Full or performance gate is closed by these focused passes.

The fresh whole-solution Release and Debug Embedded/test-enabled builds both stopped at `LNK1104` for `Common.dll`, before the main target compiled. Independently launched SearchService test probes hold those outputs: Release PID 46208 and Debug PIDs 46496, 46724 and 46800, all started at 06:02:22 UTC under Windows Terminal PID 43828. Root retained exact executable/argument/start-time diagnostics and requested permission to close only those four probes, leaving Terminal open; no independent process was killed. Build logs and outcomes are under `resume-20261005T060429Z/`. No old diagnostic receipt authorizes the changed source.

The user subsequently authorized closing those four probes only. Root validated pinned process handles, exact start ticks, executable, arguments and parent for all four before stopping them at 06:55 UTC; Windows Terminal PID 43828 stayed unchanged. The first validation attempt stopped before any termination because reparsing PowerShell's already-typed JSON date lost its precision and timezone; direct typed-date comparison fixed the diagnostic without relaxing identity checks. The exact closure receipt is `resume-20261005T060429Z/approved-exact-probe-stop-20261005T065504Z.json`. This resolves the recorded linker blocker, not the native qualification gates.

The separately reconstructed baseline `C:/RsI27FinalBaseline` passed both pre/post-build source bridges (`ae5ce86941a307830316ccebcfeb668570292f963bd4543d854249237abbc221`, all 12739 raw paths) and a fresh whole-solution x64 Release Embedded/test-enabled build, zero warnings/errors, receipt `e72275b9ab9669b9d2c73b4bd4db0721c6f9cd18ef300f5f2e93f9b007c0fec7`. Five-process measurements remain open. Current specification inventory reports 438 files, 62 normative files, 24 WIP files/index entries and zero blocking findings; retained reports are under `resume-20261005T060429Z/source-spec-inventory-20261005/`. This is a focused source/link consistency check, not a runtime or final closeout pass.

The second cheaper reviewer found no concrete bug in the current shutdown flag, fresh-child tests or documentation mappings; root independently reviewed them and corrected the new optional access to the repository's `.value()` convention. Root removed only temporary ViewerPE child desktop-phase tracing before the next build. Failure-only parent restoration diagnostics and the original instrumented evidence remain retained; no focus, input or restoration predicate was weakened.

The corrected volume-path API shadow v2 closed successfully in its own original C archive, packet `volume-path-api-diagnostic-v2-20261005T063252Z-7648-ed7aa8af5f4b4cbda886b37dc09bffa2/`: 4096 successful calls took 236757 us at the short root and 1468280 us at the exact R0e namespace-shaped deep directory. This is a different synthetic managed caller and run id, not a provider/admission pair or qualification pass. Preserve the invalid-depth v1 probe unchanged. The cheaper reviewer found no safe existing drop-in alternative; root verified the path/operation contract and API differences. A separate finite API strategy shadow `existing-path-handle-shadow-v1-20261005T073335Z-20284-13b4823b676647eca557c33449d4d423` retained 4096 deep calls at 1732357 us for the old volume resolver, 203521 us for following normalized handles, and 186525 us for following opened-name handles including cached-mount GUID revalidation; all 4096 handles per method closed. This covers only ordinary existing paths in a different managed caller, not production, alias/mount/SMB semantics or a qualification gate.

Current native candidate: `Plugins/FileSystem/FileSystem.cpp` now implements guarded following-handle observation for a fresh existing Local path beneath an unexpired cached mount. It may reuse the existing route only when the system-reported endpoint GUID and a fresh cached-mount GUID query both match the retained GUID. Missing, inaccessible, unsupported, oversized, remote, expired and mismatched observations retain the original resolver; known GUID mismatch must not authorize a stale same-root cache hit. No path handle is retained across queries, no lexical volume identity is invented, and typed per-path/operation querying and mutation bindings remain mandatory. `resume-20261005T060429Z/before-existing-path-guard/` retains the actual preimplementation source and compiled originals before any change. Synthetic Debug controls cover success, failed/oversized observations, endpoint/mount changes, failed re-resolution, TTL, UNC/no-GUID fallback and repeated real-handle cleanup. Root independently reviewed the code and the targeted Astra High read-only review found no actionable static finding; neither establishes native correctness or timing. Focused owning contracts and test obligations are reconciled. These edits are UNBUILT/UNQUALIFIED and invalidate `-SkipBuild` against receipt `5dee743…`. Next: format/build, then qualify unchanged R0e and actual admission/performance gates; do not shorten the fixture, raise the timing limit or close unavailable mounted coverage.

Root verified and retired exactly four already-archived diagnostic C run directories through canonical `Clean-TestSandbox.ps1`; all original archives remain retained. The settings run's 265 historical foreign/manual audit entries match the earlier original audit exactly (zero additions/removals), remain globally not clean and are untouched. Only selected-run closure and unchanged external entry sets are claimed. The three overall closeout gates remain open; any new native/spec edit invalidates the prior linked-build source snapshot until rebuilt.

### Historical pause checkpoint — 2026-10-05

All owned native/build sessions and both cheaper read-only reviewers have completed. No process was killed by name and independently launched applications remain untouched. The desktop reservation is released for this pause. There is no active goal. No closeout checkbox is newly checked, no Git changes are committed/pushed, and this same plan remains in WIP.

The durable resume and recovery packet is
`C:/RedSalamander.Perf/evidence/i27-followup-20261002-0e139fb/resume-20261003T121349Z/final-run-review/pause-20261005T053345Z/`.
Start with its `RESUME.md` and `checkpoint.json`; they bind the pending source, staged tree, source recovery patches, controller paths, complete failed/pass diagnostics and next dependency-ready actions. Existing large screenshots and whole-run archives are retained at their recorded locations rather than copied into another gallery. Original receipts and adverse evidence remain unchanged.

The last diagnostic whole-solution x64 Release Embedded build passed with zero warnings/errors, receipt `ba432eddab6ad1b5a1125a6abc90ee78fa394c4ab44c450d3dafb377bd09f927`, source snapshot `fc37636565d199134b4dfe5a35414a02ae3f061e14f0d76417d6f91d1ab3051c`. It does **not** bind the latest pending shutdown patch. Fresh Full Debug and Release still fail. The unchanged finite R0e series exceeded the existing 2,000,000 us ceiling in three of five processes; no timing gate is waived.

Instrumented settings diagnosis retained ordering-child success and teardown-child body/result success, followed by a parent timeout at real `FolderWindow::Destroy`. Root traced the production `FileOperationState::Shutdown` to an unconditional five-second flush of an unrelated deliberately stalled global settings save. Root then wrote a conservative sticky own-save-admission flag: a state that has admitted an `S_OK` settings snapshot retains the existing flush; an unused state does not flush another owner's save. The new child test checks bounded idle shutdown while that unrelated save stays in progress. This code, test and normative-spec patch is **UNFORMATTED, UNCOMPILED and UNVALIDATED** at the pause.

Pending patch paths are `RedSalamander/FolderWindow.FileOperationsInternal.h`, `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp`, `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`, `Specs/Core/Core_SettingsStore.md` and `Specs/UI/UI_FileOperationsPopup.md`. Before compiling, review the new teardown artifact destination validation and the debug snapshot's `completedGeneration` field. Review and range-format the actual diff, then produce a fresh whole-solution build and retain its receipt before linked tests. Do not use the earlier receipt for `SkipBuild`.

The protected instrumented ViewerPE group passed all 50 child runs, with 49 quiet checks and 44 paired desktop capture/restoration boundaries. It is auxiliary evidence; it does not repair the original Full restoration failure. Temporary ViewerPE child-phase tracing still needs removal before the final uninstrumented Fulls; retain the reviewed failure-only parent diagnostics and bounded settings child trace retention if their final review supports them.

The Win32 volume-path shadow probe is also auxiliary: its deep path accidentally duplicated `selftest/last_run`, so it has no exact R0e path parity and proves neither optimization nor causality. Preserve that packet. If a corrected probe is useful after resume, give it a new version and fresh packet; never replace the original. No handle/GUID strategy or lexical volume cache was adopted.

Next after an explicit user resume: qualify the pending owner-local shutdown fix and its existing owned-save persistence coverage, investigate R0e under unchanged provider identity semantics and budgets, remove temporary diagnostic instrumentation, then run genuine fresh Full Debug/Release and five paired Release/resource processes. Finish the normative, evidence/governance and gallery provenance updates, move this same plan to Done and retire the continuation only after every required gate passes. DxUi pin `19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1`, I26 and separate H4/I18/I28 ownership remain unchanged. The older checkpoints below are history, not authority for the current unbuilt source.

Latest execution checkpoint, 2026-10-04 22:33 UTC: the exact progressive
diagnostic `20261004T215536Z-29724-520f14e0274649b8b19b3f0f4dfde3bf`
closed PASSED, native 1/0/0: first-paint wait 254 us, settlement 174478 us,
121 directories/10000 files and no pending rows/bytes. Its original C no-follow
archive and exact reader closed. Scope correction: the canonical Commands exact-case
runner builds RedSalamander and dependencies, not the full solution. Both retained
targeted receipt records describe 133 outputs; the launcher wording does not expand
their scope. This diagnostic does not replace Fresh Full qualification.

Root fixed an actual desktop-coordination gap: every FileOperations process entry
now acquires the existing interactive desktop lease independently of activation.
Only the exact interaction-gallery case permits activation; all other FileOps filters
remain non-activating. Family/prefix filtering is native, without a second case
allowlist. F7 rejects a missing initial foreground and checks its captured HWND/PID/TID
after every sample and final cleanup, with saved/current identity diagnostics. The
cheaper read-only reviewer checked the fixes; root independently reviewed the scoped
diff. No ownership or restoration guard is weakened and no independent app is controlled.
Focused plan/inventory/governance Pester closed 87/0/0, run
`20261004T222621Z-110200-e2bdc43587364ff3b1e83ff2ab721f93` (receipt SHA-256
`66ec452259dfd2e9d9f300391c4aac9fd412151121af576e185c9b18815313c6`).
The earlier 86/1 invocation used a forward-slash spelling where the sandbox test
requires the canonical Windows path; its failed receipt is preserved. The new F7
native source and lease policy remain UNCOMPILED/UNQUALIFIED until the next native gate.

The new source-only baseline finalization in
`final-run-review/final-v2-normalized-source-finalized-20261005T000100Z`
closed PASSED: 12739 actual paths/raw bindings, all 24 normalized-content checks,
687 retained pre-overlay native bindings and the exact three already-applied v2
overlay hashes. It did not reapply the overlay or copy build products. Its bounded
experiment reproduced PowerShell 7.6.5 restoring null as an EMPTY `GIT_INDEX_FILE`:
Git saw 3926 paths; explicit environment-provider removal restored all 12739.
The failed v3 closure and original 24 raw newline mismatches remain unchanged.
The newly recorded digests bind the currently retained v3 evidence; they do not
infer an earlier independent digest receipt. This qualifies only the explicitly
labelled fresh Git-normalized source reconstruction, not the historical raw
checkout/executable or any performance result.

Next: freeze the reviewed lease/F7 bundle and run genuine Fresh Full Debug and
Release serially. Retain every receipt epoch before a later build, then run fresh
serial baseline/candidate Release builds and five exact v2 process pairs. Finish
normative/governance/curation/Done only after all required gates pass. The user's
desktop reservation persists; no further availability question is needed. All
three overall closeout gates remain open and I26/DxUi/H4/I18/I28 are unchanged.

Earlier execution checkpoint, 2026-10-04 21:54 UTC: genuine Fresh Full Debug
`20261004T193009Z-22068-332644cb008d4c2d8f298709c08c888a` closed FAILED after
all 20 entries executed: 17 promoted, native 2040 passed/3 failed/246 skipped.
The C, D and C-validation whole-run no-follow archives closed and passed the
controller's exact readers. Root independently checked packet hash/length bindings.
Its immutable receipt epochs `501f7ac20d59d147350e05dfe1815f0006ecc4bc74c199a7d1a626f752c92ea6`
and `725b26a0b016abe570d686630a2859f5a3a30d7c2f8dbfabd5ca57bf8c554378`
are retained; the receipt keeper reports 37 checks and zero issues before any next
build. Packet: `final-run-review/final-full-debug-20261004T193009Z-22068-332644cb008d4c2d8f298709c08c888a`.
The runner and launcher exited; archival custody does not change the failed verdict.

The progressive ViewerSpace scan missed its unchanged 20-second deadline. Its
artifact actually recorded Running, 121 directories and 6500 of 10000 files;
the zero-valued failure sentence was formatted before the wait updated its
snapshot. Root separated those expressions, added measured wait/deadline outcomes
and requires the completed default model to retain all 121 directories/10000 files.
This diagnostic/test correction and the owning ViewerSpace contract are written,
formatted with pinned clang-format 22.1.3, but UNCOMPILED/UNQUALIFIED. No production
scan deadline or asynchronous worker-retirement contract is changed.

F7 stopped at prime sample47 when exact foreground ownership changed from
ChatGPT to an independently launched WindowsTerminal. ViewerPE's first interactive
native case passed, then its restoration guard observed a changed LASTINPUTINFO
and stopped subsequent children. Both refusals preserve unrelated desktop input;
their causes remain unproven. The user's reservation until testing finishes still
authorizes the next protected checks; do not repeat availability questions or weaken
warning, lease, ownership, containment or restoration safeguards.

Historical baseline checkouts/builds were absent. The source-only reconstruction in
`C:/RsI27FinalBaseline` preserves the failed original 12739-file raw-copy gate:
24 physical-newline mismatches, with all 687 normalized native Git blobs matching
the retained I26 source. Root reviewed the cheaper reader's methodology assessment:
use an explicitly labelled fresh Git-normalized I26 reconstruction, retaining every
raw mismatch and new actual source/build identities. This is not reproduction of the
historical raw checkout or executable. Complete exact v2 overlay/source proof before
fresh serial baseline/candidate Release builds and five process pairs per variant.

Next: freeze this diagnostic bundle, compile and run the exact progressive case,
then genuine Fresh Full Debug and Release. Review original failed/skipped witnesses,
preserve receipt epochs before every later build, qualify final v2/resource evidence,
and finish governance/curation/Done only after the required gates actually pass.
All three overall closeout gates stay open; I26, DxUi/H4/I18/I28 ownership is unchanged.

Earlier execution checkpoint, 2026-10-04 19:27 UTC: the canonical Debug PR gate
closed PASSED, with 13/13 promoted entries and native 844/0/0, after the recovered
solution build (zero warnings/errors, 192 outputs). Its original no-follow C archive
is complete and independently validated (manifest SHA-256
`70c79b1a13f5a67c123ad0f3b0918600318f049d0bbda0052bdd83f70d64bbbe`).
The immutable build receipt `9e984997f4ad9bcfae93b683db151c5a33f740a72c58c64428778c2ea72cd7e7`
is byte-identically retained in `final-run-review/menu-ownership-recovery-receipt-retention-20261004T175900Z`
(raw SHA-256 `096b3baeecc391473c6606bd5e36e7e75b44c9507106c344162e44159566e210`).
Packet: `final-viewerpr-debug-20261004T180510Z-20128-ca107e93a87f45138403fcd86c5a0241`.
PR runs ViewerPE Noninteractive only; it does not qualify the prior F8/picker or
other interactive failures. Earlier failed Full evidence remains unchanged.

The saved `C:\RsI27Followup` checkout and build directory were unexpectedly absent
at continuation. Git retained its index. Root preserved raw metadata and the staged
binary diff, compared all 19,221 path/mode/blob entries, and recovered the identical
indexed source without changing the main checkout or adopting its newer DxUi pin.
Recovery record: `final-run-review/missing-checkout-recovery-20261004T191225Z/recovery.json`;
recovered index tree `6c225393ef6b326e1f44996bb492fb4c5ccef842`. Unindexed draft
changes and build products were not recoverable from that index; no new pass is claimed.

Root reapplied the reviewed, unindexed ViewerPE harness correction. Every owned child
is now joined; remaining job descendants are stopped and observed quiet before releasing
the job and attempting foreground restoration. Changed input or unrelated foreground,
including a different HWND in the saved app's PID, refuses restoration and stops outer
and nested churn. No foreign queue attachment or foreign-control focus is used. Failed
encoding-picker prerequisites stop dependent input and perform only owned-window cleanup.
A private noninteractive ready-before-timeout child tests root/job containment without
requiring global focus stability. It does not prove activation restoration. Its readiness
argument rejects empty, repeated and normal-case use. Root checked the cheaper read-only
reviewer's findings and corrected the overbroad saved-PID and ungated cancellation checks.
Shared-helper/selftest contracts and their focused consistency anchors are updated.
The restored draft is UNCOMPILED/UNQUALIFIED at this checkpoint. The pinned formatter
22.1.3 was restored from the CI hash-locked requirements; the installed 22.1.8 is not
used as qualification evidence.

Next: freeze and stage this bundle, run canonical genuine Fresh Full Debug then Release
(no SkipBuild or reduced suite), review all original outcomes/archives, and preserve each
immutable receipt epoch before any following build. Existing human desktop reservation,
warning and lease remain in force; do not repeat availability questions. All three
closeout gates remain open. I26, DxUi's pin and independent H4/I18/I28 ownership are unchanged.


Earlier execution checkpoint, 2026-10-04 17:56 UTC: the first focused Viewer PR
controller stopped during build preparation and has no original completion or
native verdict. Packet `final-viewerpr-debug-20261004T174341Z-45128-e3965a12cb4c472482ab35dd06b913c0`
remains INCOMPLETE; its separately labelled later empty-run capture passed its
reader without repairing that missing original custody. The next controller
`20261004T175134Z-11784-c6ba3fa242914b79992f14989d139b24` refused before creating
a run: the canonical artifact guard detected the abandoned operation and required
`build.ps1 -Rebuild`. Its original refusal and missing-run archive report are retained.
No native result or pass is claimed for either attempt, and no marker is removed
manually. The independently launched RedPrism build has disjoint outputs and is
untouched. Next: canonical test-enabled entire-solution Debug recovery rebuild,
then the protected Viewer PR run using its verified receipt; final Full Debug and
Release remain required. A hidden root-owned launcher preserves the one serial
harness across chat turns; the existing warning, lease and restoration still govern
all native input. The source stays frozen during build/test execution.

Root reviewed the cheaper reviewer's final EXIF regression check: the existing
handler supports the programmatic toggle/restore assertions, without a claim of
physical menu selection. Two focused consistency notes now accurately distinguish
the injected failure from the actual preceding DxUi detachment. All three overall
closeout gates and the independent I26/H4/I18/I28 ownership boundaries remain unchanged.

Earlier execution checkpoint, 2026-10-04 17:40 UTC: the menu-ownership implementation
compiled in the canonical test-enabled x64 Debug entire-solution build with zero
warnings/errors. Receipt `c93fa08fa0eed82d8a286305fbd5b1345c5bca4fea58beee4776e3742c961e17`
and all 192 output bindings passed canonical semantics; its raw immutable record is
byte-identically retained under `final-run-review/menu-ownership-solution-receipt-retention-20261004T172413Z-41936`
(SHA-256 `76ff79168221693e94f68b1c673df0a82e827d724064ae74be4587eba24f25f0`).
The linked exact throughput-disclosure/clipboard case closed 1/0/0 in packet
`final-namedcase-debug-20261004T172413Z-41936-3e85de3e1d1e4e9994126eece984e3da`.
Its original C archive is complete and independently validated, with no invocation
or archive error. Clipboard restoration and all path-detail cycles passed. The
earlier Full clipboard failure remains unexplained and is preserved unchanged.
This focused result does not establish Full qualification.

A cheaper read-only reviewer checked the ownership changes again and proposed
focused consistency anchors. Root reviewed and applied only the three owning
records, without restamping their whole-document reviews. Root also added a
meaningful fallback command regression: both early and post-detach native menus
must execute the EXIF toggle, update their check state, then restore it while
keeping the original menu identity. That added assertion and the new metadata
still require the next build/native qualification. No library/pin change is made.
Next: canonical focused Viewer PR diagnostics, then genuine Fresh Full Debug and
Release; preserve each immutable receipt epoch before any following build. The
desktop remains reserved by the human under the existing protected harness.
All three overall closeout gates remain open; I26/H4/I18/I28 are unchanged.

Earlier execution checkpoint, 2026-10-04 17:11 UTC: Fresh Full Debug
`20261004T124815Z-21664-e1d90d57fff64fa0b666adf73276e30f` closed FAILED:
20 entries executed, 17 promoted, native 2040/3/246. The original C archive
(56,808 files), D archive (one file), and C-validation archive (78 files) are
complete and passed their independent readers. Immutable solution receipts
`d9567b3d315ff20373976c198a8ea7ac71697906ebc9c25c787af352576db59e` and
`0453b2f99ac5bbb2e14594f87ed292b845b3e7187dea537c81fefc3f1c414f0a` are
byte-identically retained in that packet's `receipt-epoch-retention`; its canonical
37 checks passed with zero issues. Custody does not change the failed verdict.
All external packets remain under the existing `final-run-review` evidence root.

Commands failed the rich-format clipboard flush-contention fixture after its first
guarded injected attempt; its UIA, Space and pointer disclosure routes passed.
FileOps stopped during actual F7 sample 213 after the foreground changed from
WindowsTerminal to claude; the input guard did not transfer focus. ViewerPETests
Interactive failed encoding/picker and later menu-host checks, plus a loader-submit
deadline whose numeric measurement was missing. Runtime causes remain unproven.

Root applied the independently reviewed fail-only clipboard/input/loader diagnostics.
A separate source-confirmed consumer ownership hole is now drafted: the pinned menu
host can fail after removing the native menu. Five viewers attach only after the
creation-time guard hands off its handle; `Common/ViewerMenuHost.h` restores fallback,
reconciles reentrant close, preserves a replacement, and retains a detached handle
only when restoration fails. Four content layouts run after attachment. Meaningful
hidden Win32 ownership cases and real ImgRaw post-detach probes are added. A cheaper
read-only reviewer identified layout and rollback gaps; root corrected and reviewed
them. These edits are UNCOMPILED/UNQUALIFIED. The library and its pin remain unchanged.

The human reserved the desktop until testing finishes. Continue inside the existing
warning/lease/restoration safeguards without repeating availability questions.
Next: format, compile, focused ownership/encoding/clipboard diagnostics, then genuine
Fresh Full Debug and Release on the frozen final source; retain each receipt epoch
before another build, curate within existing caps, run final governance, and move this
same plan only after required gates pass. I26 and the independently owned H4/I18/I28
work remain unchanged; no failed evidence or deadline is waived.

Earlier execution checkpoint, 2026-10-04 12:44 UTC: the reviewed native fixture repairs
are compiled and their linked focused Debug chain has closed successfully. The exact
Shortcuts case passed 1/0/0, its family 40/0/0, ViewerSpace 20/0/1, LocalBandwidth
3/0/0 and ParallelBandwidth 3/0/0. The one skip is the explicitly opt-in 20,000-item
ViewerSpace layout benchmark. Every original focused archive passed its independent
reader; the ViewerSpace archive retained 24,256 ordinary files. Its timer trace records
initial idle false, synchronous rearm true and final idle false on the same HWND and UI
thread. LocalBandwidth records ordered handler/request/call-return/completion ticks and
16 ms delivery against the unchanged 250 ms gate. ParallelBandwidth exercised four
active workers in both modes. These are focused witnesses, not Full qualification.

The actual test-enabled x64 Debug entire-solution rebuild passed with zero warnings and
errors, immutable receipt `7ee3adcd413266e851a5c85bbb4d9e17845f315a8d98f10c5bd8809f638e2df5`,
source snapshot `88521f6d0232ab3c26e4a9b7ccd2dda43e7f5b548b6cc5741b3a27cf9044376f`.
Root and the independent reviewer checked its canonical schema/semantics, 192 output
bindings and byte-identical retained receipt. Raw receipt SHA-256:
`fec0d7b0910c8dfcc6dfd34be93743fba556565666a3843d0e07417c47c567b1`.
The earlier `2448fb...` rebuild selected only the application, so the linked skip-build
attempt was correctly refused before native launch. Its focused timer pass and the
refused attempt remain retained separately; neither is described as a solution build.

Canonical containment/governance passed 353/0/0, inventories had zero blocking findings,
and all 48 changed native files passed formatting. The five-run receipt is
`post-fixture-focused-chain-20261004T121606Z.json` under the evidence root below; it binds
each original summary and run. Before another build, the actual solution receipt was
retained under `focused-solution-receipt-retention-20261004T121607Z-21608`.
The earlier failed Full runs, incomplete original captures and separately labelled later
recovery remain failed history. The 265 historical sandbox issues are still visible;
there is no claim of a globally clean sandbox. The independently launched Release
application was observed and left untouched.

Next: freeze the reviewed code/test/tool/spec and asset inputs after this checkpoint,
then genuine serialized Fresh Full Debug and Release with authorized D-volume controls.
Require each original archive/absence closure, canonical immutable-epoch retention before
the next build, capped curation and final governance before the metadata-only Done/index
move. Root remains the sole writer/build/native/input owner; no commit/push, I26/H4/I18/I28
closure or unrelated-application changes are included.

Earlier execution checkpoint, 2026-10-04 11:40 UTC: genuine Fresh Full Debug
`20261004T054036Z-80612-60e6bb68d8934d4aa370d59535713f24` finished FAILED:
20 executed entries, 19 promoted and one failed Commands entry. File Operations passed
203/0/20, including the actual destructive storage witnesses and the unchanged 250 ms
UI cancellation callback gate (110 ms). Commands passed 930 cases, failed the Shortcuts
reordered/resized clipboard copy and ViewerSpace idle-timer rearm observations, and
skipped two cases. These failures remain failed; no fresh Release Full pass is claimed.

After native containment closed, root stopped only the positively identified owned
post-native archiver after more than two hours of incomplete archive work. The original
controller/partial archive remain incomplete. The owning archive contract permits a
separate later capture; it does not require keeping a failed archiver running indefinitely.
No independent application or native test was stopped. Original interruption SHA-256:
`4d45b5ef74f2b44141671c654b71454fae5b9c8f258004e3b294a3bf04332b79`.
The corrected separate later capture closed at 11:23 UTC with C's 52,751 ordinary
files and seven no-follow reparse leaves, plus 80 C-validation files. Both independent
readers passed; D was absent at this later observation only. Recovery receipt SHA-256:
`68c62fcba33dff2fa082ae26839ccb1aa66029a0bee72e98045ad83e353e6c48`.
Original boundary proof and qualification remain false. A first later-capture draft was
interrupted before admission to correct provenance/write guards; its partial files and
both wrapper boundary-rejection controls remain retained.

Before any new source edit or build, failed-only immutable receipt retention passed
all 39 canonical record/epoch/output-map checks, zero issues, and retained both original
immutable receipts. Retention report SHA-256:
`f7ab2fa4f732e1ab3d798034a11d00b9ae3e6d43fc93d1fdbeb75d018f55b602`;
wrapper SHA-256 `c57dce67e511f9d1c8e4567c732c7e5daf491087b6a648cf730061a1be14f258`.
It restored the exact private helper; canonical source and both root markers were unchanged
through that closure. This is historical failed-run retention, not original archive repair.

Root has now applied the reviewed fixture changes: synchronous owner-thread ViewerSpace
rearm observation before pumping, successful clipboard clear/quiet and selected-row focus
before one Shortcuts copy attempt, and ordered cancellation handler/request/return/UI
callback ticks. Existing TSV/content, both idle-stop limits and the 250 ms callback gate
remain unchanged. The canonical private sandbox helper now reads fresh attributes for
every ancestor without caching; only not-found permits missing descendants. The corrected
ten-case containment test, inventory/README and normative contracts are applied together.
External parity and counterbalanced small-fixture diagnostics passed. Canonical focused
tooling/governance passed 353/0/0, including all ten new containment cases; tool/spec
inventories had zero blocking findings and all 48 changed native files passed formatting.
The [bounded diagnostic packet](../../TestRuns/4cb089111a23/Validation/2026-10-04_114600-containment/README.md)
passed explicit pre-stage archive validation and preserves the failed diagnostic, all
eight original records and their exact byte/hash map. Native fixture changes remain
UNCOMPILED/UNQUALIFIED at this checkpoint. No atomic concurrent-mutation, large-archive
performance, source/output rebuild or native Full pass is inferred.

Next: canonical containment/governance/format checks, fresh test-enabled Debug solution
build and exact/family fixture plus bandwidth regressions. Freeze all code/test/tool/spec
and asset inputs before genuine serialized Fresh Full Debug and Release. Require original
complete archive/absence closure, immutable-epoch retention before the next build, capped
curation and final governance before the metadata-only Done/index move. Root remains the
sole writer/build/native/input owner; no commit/push, I26/H4/I18/I28 or unrelated-app changes.
All failed originals, later-capture records and proposed/corrected controllers remain under
the `final-run-review` evidence root below.

Earlier execution checkpoint, 2026-10-04 05:30 UTC: the fixture repairs compiled in the
test-enabled x64 Debug full-solution build with zero warnings/errors, receipt
`5ed48edca0e422a04485ca5a0a667c3fece5428f2a68fa2195666d4e95c87aaf`, source snapshot
`449fb0c81495c29f98b45c3947f284921b0f998f3183363ca11c941755e016ac`. Three separate
ViewerVLC native processes passed (19.743/26.251/25.547 ms forced-parent return, unchanged
250 ms budget, cleanup release withheld at return). Complete ProductUi passed 27/0,
TerminalCluster 10/0, Safety 3/0, environment/inventory tooling 26/0 and focused governance
317/0. Formatting passed for all 45 changed C++ files. These are bounded focused Debug
witnesses; the final repository Full and Release gates remain open.

Root-reviewed no-follow archive v6 passed thirteen runtime checks, empty-source and real
file/directory symbolic-link controls, including rejection of mutated ordinary bytes,
unsafe paths, missing/extra members and invalid raw blobs. The original failed Full and
its incomplete seven-link archive remain failed history; later link capture cannot repair
their original provenance. The new owning archive contract is in
`Testing_ValidationEvidence.md`; external keeper/curator checks also bind the original
closed-controller manifest receipt. No literal clone, link-target payload, global clean
sandbox, ARM64/ASan, mounted-route or unexecuted Full qualification is asserted.

FileOpsAll Debug `20261004T040227Z-78724-6789727bd6f84325bc0e09598bcab641` finished
198 passed/2 failed/23 skipped. Its unchanged cancellation callback gate observed 265 ms
against 250 ms, while the retained worker boundary reported 15 ms and the UI thread
recorded a 255.922 ms paint. This is a failed delivery-latency witness, not a demonstrated
limiter defect. The other failure was the independent-volume control without alternate-root
authorization; two cross-volume storage cases consequently remained capability skips.
The complete original v3 archive passed independent validation, manifest SHA-256
`35c885040dcab571dd957c1fe6a16e6f6b3b0a92833ad9e93ceedac14b60dec4`.
One unchanged isolated diagnostic, `20261004T052208Z-60772-2b7593f894624c1bb039c6f64dec4147`,
passed Setup/LocalBandwidth/Cleanup (3/0), with 15 ms worker cancellation and 218 ms UI
delivery. It does not classify or repair the broad timing failure. The 250 ms gate,
original failures and skipped ParallelBandwidth case remain intact for genuine Full.

Next: freeze the reviewed code/spec/tooling/assets after focused documentation governance,
then run genuine serialized Fresh Full Debug and Release with explicitly authorized D
controls. Each Full must pass the unchanged delivery gate and remaining destructive cases.
After each runner and exact C/C-validation plus D archive-or-original-absence outcome close,
retain every actual used immutable receipt epoch before another build. Final capped
curation, staged archive inventory, separately reviewed metadata-only Done/index changes
and continuation retirement remain required. Root is the sole writer/build/native/input
owner; independently launched apps and H4/I18/I28 remain with their owners. No commit/push.
All focused packets and original logs are under the `final-run-review` evidence root below.

Earlier execution checkpoint, 2026-10-04: the user's "take the computer for testing now"
authorizes the continuing protected harness session; the warning, desktop lease, exact
ownership guards and restoration remain required. Root remains the sole checkout writer,
builder and native/input owner. No commit/push is requested for this continuation.

Genuine Fresh Full Debug `20261004T010926Z-24184-5f43d912bdca4d87802311e5ec40c166`
finished FAILED: 20 executed entries, 17 promoted, 3 failed; native totals 2025 passed,
3 failed, 247 skipped. Source snapshot `c9b9cad92243fda972287f0f5a089bafe30fa14fd7809f986b6b4f7c8fbe9edc`
and original initial/post-writer receipts remain bound to that failed run. File Operations
stopped before its remaining broad cases; its earlier destructive byte witnesses passed.
The three failures are synthetic ViewerVLC forced-parent timing, asynchronous floating
Terminal restore snapshot, and F6 optional environment reading. Current repairs retain
the original timing/identity/error assertions and add exact diagnostics and regression
coverage. They are written, not yet compiled or qualified at this checkpoint.

The original ordinary-file native archive copied 52,736 files/3,008,176,448 bytes but
missed exactly seven fixture reparse leaves, so its original archive remains incomplete.
A separately reviewed failed-metadata keeper retained both immutable receipt epochs and
canonical entry records with qualification explicitly false. Future archives need
no-follow raw link metadata capture; no original raw link payload is retroactively proven.
Evidence and proposed controllers are under
`C:/RedSalamander.Perf/evidence/i27-followup-20261002-0e139fb/resume-20261003T121349Z/final-run-review/`.
Next: format/review the repairs, focused tooling and fresh full Debug build, exact VLC,
Terminal and File Operations checks, then final Fresh Full Debug/Release with reviewed
no-follow closure. Keep all original failures, H4/I18/I28 and mounted-route limitations.

- [x] Retain a same-fixture pre-change Release performance baseline and source/module provenance.
- [x] P1: Carry prompt generation through all consent bindings, including primary, overflow, Escape, UIA, Apply-to-all and safe initial focus.
- [x] P1: Reject held activation/cancel key repeats before hosted dispatch; a refreshed B binding must require a fresh press. Preserve repeated navigation.
- [x] P1: Extract coherent conflict acceptance/consumption: generation validation, one-shot decisions, and event/state transitions under one lock; stale signals never imply Cancel.
- [x] P1: Extract a pure typed-item reducer; missing cancellation receipts after possible mutation retain unknown storage truth. Preserve stronger host-managed evidence.
- [x] P2: Preserve path-detail text/selection across reconstruction while restoring native text focus only when the popup owns it.
- [x] P2: Extract Issues projection and use consistent model revisions; hidden/minimized panes perform no timer rebuilds and refresh on showing/restoring.
- [x] P2: Fence alert providers and queued actions by presentation identity, with owner-thread dispatch and HWND lifetime protection.
- [x] P2: Make native failure-surface creation transactional, clean partial children, and retry failures safely.
- [x] Add deterministic regression tests for stale/duplicate consent, concurrent wait ownership, receipt matrices and provider integration, focus policy, hidden Issues, stale alert providers, and every native-child failure.
- [x] Review delegated implementations against production callers and contracts; remove redundant policy paths and document extracted ownership boundaries.
- [ ] **Transferred to I29 Q1/Q2, not passed here:** Run x64 Debug and Release builds, targeted regressions and Full gates. Compare retained Release performance on the unchanged integrated fixture. Preserve skips, failures and any deferred platform boundary truthfully.
- [ ] **Transferred to I29 Q3, not passed here:** Update affected normative contracts, helper catalog, coverage/evidence and consistency anchors; pass spec/tool inventories, focused governance tests and format validation.
- [ ] **Final publication/qualification transferred to I29 Q2/Q3:** Admit the exact completed-plan path with a focused governance test, archive evidence under Specs/TestRuns, then move this plan to Done and retire I27 from the active index. The exact-path admission, implementation-plan move and safety-index retirement are performed by the user-approved delivery split above; remaining publication/qualification is still open.

## October 2 verified follow-up

The October popup qualification is not proof of all filesystem outcomes. Source review confirmed
the following remaining defects/opportunities without establishing a current data-loss bug.
Historical failed Full runs remain evidence; later fixes do not turn those runs into passes.
The original I27 qualification checklist above remains required after these changes.

- [x] F1: Carry key repeat metadata through the native failure-surface button hooks and parent
  dispatch. Ignore repeated Enter/Escape activation while preserving fresh input, navigation,
  pointer/UIA invocation and standard Space key-up semantics. Test the actual child subclass,
  consent boundary, Close's hide-only behavior and fixture teardown without desktop activation.
- [x] F2: Project a durable cancellation request as explicit `Stopping` for unfinished tasks,
  including conflict/queue/pause/verification states. Normal finishing also publishes the
  Stopping lifecycle phase, so phase alone is not cancellation authority. Terminal truth wins.
  Cover English/French text, UIA/status/taskbar counting and all relevant action availability.
- [x] F3: Format Copy/Move and permanent-Delete confirmations with the captured provider path
  identity contract. Preserve provider separators, case, roots, Unicode and exact source/destination
  identities; group parents only under that contract. Reuse canonical path helpers. Unknown
  contracts must preserve raw text rather than silently apply Windows rules. Add local/FTP/S3
  and case-distinct-parent controls, checking display independently from operation endpoints.
- [x] F4: Guarantee owner-thread consumption of stored completions when all message posts fail,
  including the post after the orphan worker joins and threadpool-submission failure. Preserve
  one result per task, worker ownership, shutdown joining, payload registry fencing and host
  lifetime. Terminal status is already published before posting: the defect is pending UI
  completion/refresh/callback/cleanup, not necessarily a card displaying Running. Inject every
  posting failure and prove bounded recovery, exactly-once callbacks and quiet teardown.
- [x] F5: Reject a claimed successful Delete receipt that commits but retains the original
  item as successful completion. Partial/canceled directory deletion may legitimately retain
  the root; preserve those independent storage/cancellation axes. Cover producer, callback,
  pure reducer, finalizer and Issues projection with a hostile provider receipt.
- [x] F6: Retain a baseline for no-work/hidden/minimized/running/completing popup states, then
  suspend telemetry when no operation or UI animation needs it. Publish final taskbar state,
  rearm on admission/show/relevant state transitions, and preserve hidden/minimized active-work
  taskbar updates. Measure snapshot/timer/taskbar counters and idle CPU/resources on the same fixture.
- [x] F7: Measure actual UI admission and provider-call/allocation counts for available Local/UNC/alternate-volume
  routes and provider partitions at 1/128/4096 selected items; attempt real mounted routes and retain
  capability-bound skips as unqualified route coverage when Windows denies registration. The retained baseline Local implementation allocated a 32,768
  wchar buffer and resolved the volume before consulting its two-second cache; a fallback Move
  queries Copy capability again. Optimize only after paired evidence, without conflating mounted
  roots or bypassing endpoint/capability/consent authority. Old synthetic-loop timings are not
  production admission measurements.

### Reviewed bounded F6, integrated v2 and F7 disposition

The retained six original F6 Release ON/OFF process receipts cover the same eighteen states;
all twenty-seven ON idle intervals have zero popup timer/snapshot/rate/taskbar work, while
hidden/minimized active taskbar updates continue. Root accepts that scheduling policy with its
measured transient resource differences and adverse active-work directions. It does not claim
zero whole-process CPU, CPU/memory neutrality, a leak cause or post-destroy resource release.
The ten original strict v2 Release product runs form five baseline/candidate pairs on identical
fixture bytes and geometry. Root accepts the bounded descriptive comparison with its adverse
changing-Attributes and individual timing/resource samples, unequal vcpkg identities and
unresolved lifetime limits; the earlier v1 graph regression remains unaccepted history.
[F6/v2 assessment and retained representative packet](../../TestRuns/4cb089111a23/FileOps/20261003T211512Z-80796-0cb73e59c3c741f9bcae95a38bb43929/review/ROOT-F6-V2-BOUNDED-ASSESSMENT.md).

The actual F7 before/after Release matrix qualifies 252 available-route semantic pairs and
126 separate prime-ingress pairs with canceled/no-write byte and BLAKE3 evidence. Candidate
root/GUID output-buffer heap allocations are zero for those qualified samples and eight
capability/root-resolution call families match. All 53 adverse-direction observations and
multi-second UI-admission limits remain in the [F7 assessment and retained candidate packet](../../TestRuns/4cb089111a23/FileOps/20261003T191317Z-82816-199e9f25af4d49739d1410c6881e3e0d/review/ROOT-F7-ASSESSMENT.md).
All 24 actual mounted samples are `S_FALSE` capability skips after `SetVolumeMountPointW`
returned access denied; native mounted behavior and mounted-route non-regression remain
unqualified. Debug synthetic controls do not substitute. This is the allowed unavailable-route
boundary in `Testing_SelfTests.md` and `Testing_TestCoverage.md`, not a mounted pass or a
new privilege change. Eighteen original F6/v2/F7 packets are curated under
`Specs/TestRuns/4cb089111a23/FileOps/`; the pre-stage review reported 724 files passing.
Final staged archive inventory and all final governance remain pending.

This bounded measurement disposition does not close the pending final Debug/Release Fresh Full,
protected/gallery publication, native/schema/governance or archive gates. Final receipt slots:
Debug Full `<PENDING-DEBUG-FULL-RUN-ID-AND-BUILD-RECEIPT>`; Release Full
`<PENDING-RELEASE-FULL-RUN-ID-AND-BUILD-RECEIPT>`; governance
`<PENDING-SPEC-TOOL-ARCHIVE-FORMAT-RECEIPTS>`. These are placeholders, not passes;
the three overall closeout checkboxes above and the active WIP index remain open.
Source-review corrections: native BUTTON Space does not emit repeated clicks in the isolated
hidden-window message probe (ten repeat key-downs: zero clicks; key-up: one). Native fallback Enter
explicitly sends `BM_CLICK` per key-down and needs the metadata fix. The junction self-copy test
already admits a real task and checks rejection/no recursive child, but not the entire source
tree. A missing storage assertion is a proof gap, not proof of a destructive product defect.

### Destructive outcome proof

Every fixture must use governed, owned test roots; retain distinct before/after content hashes,
tree inventories and item-result axes. An HRESULT, guard enum, task-id result or file size alone
cannot establish byte preservation. Deterministic mutation barriers should select the cut point,
not timing sleeps. OS ACL/lock fixtures must restore only case-owned resources.

- [x] D1 (priority): Cancel Local recursive Delete after actual child removal; prove removed
  children, retained subtree bytes, cancellation status and no removal after the cancellation gate.
- [x] D2 (priority): Fail a real cross-volume/cross-provider Move after some files publish;
  prove both trees' contents and the source-removal/verification boundary. Rename the existing
  `FileOps_CrossVolumeMovePartialFailureStatus` copy-only Dummy success case truthfully.
- [x] D3 (priority): Execute same-file and hard-link Copy/Move via production admission/execution;
  prove both aliases retain their original bytes and no existing alias is truncated or replaced.
  Direct same-path Copy retains its specified Keep Both behavior; prove only its expected new
  sibling is published, with identical bytes and a distinct provider object identity.
- [x] D4 (priority): Fail overwrite with an exclusively opened destination and after a controlled
  write failure; prove the destination's original content, source content and owned-stage cleanup.
- [x] D5: Strengthen direct and aliased copy-into-self tests with complete source tree/content equality.
- [x] D6: Deny access with an actual Windows child-directory DACL; inspect surviving/published trees
  and typed results separately from injected ACCESS_DENIED or junction-walk protection.
- [x] D7: Skip one conflict while copying others; assert completed-byte accounting counts copied
  data only and preserve every skipped occupant's content.
- [x] D8: Exercise file case-only rename through the actual host conditional-rename route;
  separately fail the public provider's existing temporary-name fallback after its first hop.
  Prove exact casing, original object identity/bytes and no abandoned temporary sibling. The host
  uses handle-bound `RenameExactTo`, not the provider's `.rs_case_tmp_` fallback; do not invent a
  host fallback or claim a provider-only witness establishes a host failure path.
- [x] D9: Submit zero selected items; prove no task, provider mutation or filesystem change.
- [x] D10: Destroy the real owned main window while a write is paused at a deterministic mutation
  boundary; inspect source/destination bytes after the actual shutdown path and worker quiet point.
- [x] D11: Start the same Copy/Move twice through a non-clipboard route. Specify the expected
  conflict/queue behavior first; a second intentional user command is not automatically invalid.
  Prove safe bytes and result/focus disposition, separately from clipboard-sequence deduplication.

### Execution checkpoint

2026-10-04 01:09 UTC — pre-Full source freeze. The original protected/static
packets and all capture companions are published. The generated gallery resolves
all 1,165 link occurrences to 590 distinct local files, with 562 static and 13
protected original images. The existing protected-history guard admits exactly
the 596 reviewed companion paths; a new fixture rejects an unreviewed image in
that same directory. No subtree wildcard or historical rewrite was introduced.

The first archive staging exposed 441 Git normalization mismatches; exact owned
`-text` attributes and explicit index renormalization corrected all 1,404 staged
publication files to their raw working-file object IDs. Original console whitespace
is preserved with reviewed path-specific attributes. Failed audits remain retained.
The canonical pinned CI formatter 22.1.3 now passes all 41 changed/new native files;
the earlier 22.1.8 formatting was not treated as the pinned gate.

Focused governance passed 311/0/0 and the additional current tooling tests passed
29/0/0. Spec/test/changed-archive inventory passed. Tool inventory initially found
the missing `RunAllTestsPlan.Tests.ps1` consumer of `ValidationFingerprint.psm1`;
the exact manifest correction now passes with zero blocking findings. Prior failed
Pester/spec/consumer reports remain original evidence. Whole archive inventory
passed 9,602 files before the additional governance-only edits; final archive
inventory remains required after actual Full curation and the metadata closeout.
Receipts: external `final-run-review/pre-full-governance-fixed-20261004T010131Z/`,
`pinned-format-final.json`, `root-staged-original-byte-review-v2.json` and the
original root publication review. No runtime Full pass is claimed by these checks.

Next execute genuine serialized Fresh Full Debug then Release with this frozen
implementation/test/tooling/normative/asset bundle. Do not edit tracked inputs
between or during those runs. Immediately after each contained runner and its
whole C/C-validation archives close, retain all actual used immutable receipt
epochs before another build. Only after both final runtime gates pass, update
final Markdown/index/status/receipt links, move this same plan to Done, validate
that narrower metadata delta and retire the continuation. Original Full source
identities must remain distinct from the later metadata snapshot. The user's
continuous protected desktop interval is active; no commit/push.

2026-10-04 00:45 UTC — corrected full x64 Release Embedded test-enabled
receipt `da364927b9432c7e329e7e1305c173981170c4455b1dc0593ca635db66544e24`
binds source snapshot `eaad52c9e8e75de253c82268a7d1c9e98e49418d7eefc1c40cd451bd4eb43b9c`
and all 192 outputs, with zero warnings/errors. The earlier failed compiler
attempt remains retained; it caught an invalid test-only snapshot field and
was corrected by capturing the live arbiter item under its lock. The immutable
receipt original is retained beside a root correction of the separate retention
note's erroneous artifact-count field; the receipt bytes were never rewritten.

Focused live Issues capture `20261004T002150Z-82732-b047c1091224434cba98d32db548ea77`
passed 1/0/0, including four actual English/French primary/technical PNGs with
render, geometry, locale, selection, column and disclosure witnesses. Root
visually reviewed the two French captures. Repaired Commands family
`20261004T002346Z-43352-6d615f0acafc45218b7e91321f3f85c6` passed **65/0/0**,
with no quarantine, repair or skipped cases. The original trace proves one
owned exact clipboard restoration before the forty reconstruction cycles and
stale generation rejection against the two actual concurrent item indices.
Its whole C archive contains 155 regular files / 249,465,009 bytes and is closed;
the pre-existing 265 foreign audit findings remain unchanged. These are focused
results, not repository Fresh Full or global-clean evidence.

Next: finish byte-exact publication of the reviewed original protected/static
packets and generated navigation; validate and stage only those owned archives,
finish normative/gallery/governance checks, then freeze all implementation,
test, tooling, normative and asset inputs for genuine Fresh Full Debug and
Release. Retain every actual used immutable receipt epoch before the next build.
After both pass, allow only the documented final Markdown/index/status move,
run its structural/focused/archive gates, and explicitly preserve the original
Full source identities rather than claiming Full on the later metadata snapshot.
The user's protected desktop interval remains authorized. No commit/push.

2026-10-04 00:10 UTC — full x64 Release Embedded test-enabled receipt
`aa261e2bf802576713e3a5a739b84f9b8b23e352c62ad73049e6f9edee9d2fc4`
built with zero warnings/errors. Protected Interaction
`20261003T231906Z-20588-9a03d56cc9484b5db5b1358771aee583` passed 1/0/0:
real pointer/keyboard, readonly Unicode text, selection/caret/focus and native
144-to-96-to-144 DPI checks. Static Gallery
`20261003T232201Z-75488-0a7264eaad564dcd8aa1014d5093c463` passed 1/0/0,
562 captures on the two native displays. Both complete C archives and the
unchanged foreign 265 audit findings remain retained under final-run-review.
These are focused witnesses, not Fresh Full or external AT qualification.

Commands `20261003T233032Z-85672-02f2c2e7e9e54bf1a7a5adef3e20641f`
closed FAILED 62/3/0. The taskbar initialization/backoff repair passed. Failures
were the live Issues language capture, exact clipboard restoration (ERROR_BUSY),
and the same-task stale prompt fixture. Root has written UNCOMPILED test-only
repairs: compare replacement item against the actual first concurrent item;
restore all original clipboard formats immediately after the final copy, with
no second publication after a successful Set; observe a real localized render
and stable geometry, retain the precise capture HRESULT, and recheck every
asserted UI identity before and after capture. Read-only cheaper-agent findings
were verified and incorporated. No product policy, capture fallback, deadline,
clipboard ownership or focus refusal was weakened.

Next: format/review and build this frozen bundle, qualify the three failed
cases and Commands family, then publish the reviewed evidence/gallery, run
fresh Full Debug/Release and final governance. Full epoch receipts must be
retained before another build. F6/F7 bounded assessments now close only their
recorded measured scope; all adverse timings/resources, unqualified mounts,
historical failures and separately owned H4/I18/I28 work remain explicit.
The user's continuous protected testing interval is active. No commit/push.

2026-10-03 23:11 UTC — full Release build
`15366809dc8a1eb13fbf450cdedfcd6cdb220c7bf99c31d14a72a56d34b2a7a0`
completed with zero warnings/errors. Independent source review caught a readiness
guard gap: resetting the stability counter on focus loss could allow later recovery.
Root now latches any observed native/foreground ownership loss, ends that attempt
and rejects input regardless of subsequent recovery; an unreadable window rectangle
also resets consecutive stability. No activation rescue, synthetic hover or production
layout change was added. These last test-source changes require a new full Release
build before protected/Gallery/Commands and Fresh Full Debug/Release.
The final focused documentation/harness run at
`final-focused-doc-harness-20261003T230536Z` now retains typed results and NUnit XML:
209 passed, zero failed/skipped. Its prior redirected-result failure remains historical.
All eighteen performance packets retain their exact original bytes and canonical
pre-stage pass; staging, final native/visual/Full/governance and Done/index remain open.

2026-10-03 23:04 UTC — Release receipt
`ce9d17efa72afff623d1cc567047239fdf546accdfc1b56b769d77590a6eb491`
built with zero warnings/errors. Protected run
`20261003T230003Z-49900-a6354bf3537b4b08ab1c11c6553db680` closed FAILED 0/1/0,
complete C archive: the new diagnostics prove real WM_MOUSEMOVE delivery, but queued
auto-sizing moved FooterOptions from y=520 to y=429.333 DIPs during the pump. The
guard refused the final pointer position and sent no mouse-down. Root's fixture now
waits for no pending/active layout motion and three stable native-window/button
observations with continuous owned focus before recording pointer targets; no product
sizing policy or focus-preservation assertion is bypassed. A fresh full Release build
and protected rerun are next. The documentation/harness Pester log records 209/0/0;
its shell redirected the typed result before capture, so the wrapper failed rather
than generating a typed green receipt. Preserve that capture limitation; canonical
Full and final focused receipts remain required.

The canonical explicit pre-stage archive validator passed all 724 files across the
eighteen copied original F6/v2/F7 packets (`canonical-perf-prestage-20261003T225433Z.log`).
They remain unstaged. Independent static-gallery review counted 562 actual PNGs,
281 each at native 144/96 DPI, all French, with selected long-name/confirmation/
Issues/path-detail/light frames reviewed. This static evidence does not qualify
routed native input or external AT. The latest user-authorized protected interval
remains active; final native/Full/governance and Done closeout remain open.

2026-10-03 22:53 UTC — the latest full x64 Release test-enabled build completed
on receipt `1d112f434c2748a24cc80bb272fbf71df8d318e4158fa272decf076de6090c9a`.
Static Gallery `20261003T223746Z-24624-72ab9544c11043e7a46d2f7b2dbfdbd4`
closed PASSED 1/0/0, with a complete C archive and the unchanged foreign 265 audit issues.
Protected Interaction `20261003T223529Z-50116-7067b736952242a999a7dbebef87f389`
closed FAILED 0/1/0: routed PageDown, exact text/selection/caret/native focus at
144-to-96-to-144 and Tab/Shift+Tab passed; the pointer hit the exact owned popup,
but real hover was not reached. No mouse-down was sent. Root retained both originals.
The next fixture repair records actual retained hit/control bounds, capture and
WM_MOUSEMOVE delivery, and requires two distinct owned interior pointer movements
before real hover. The four new LONG-to-FLOAT warnings use explicit conversions.
The reviewed narrow helper/selftest prose and four consistency-anchor updates are
applied; these new changes require a fresh build and native qualification.

All eighteen F6/v2/F7 performance originals have now been copied byte-exact to their
original `Specs/TestRuns/4cb089111a23/FileOps/<run-id>` directories. Root independently
verified every source/copy/after digest, all 15,386 metric rows, original receipt/component
bindings and file/run caps. `final-run-review/root-publication-20261003T224041Z-d105cf73f3e94e19948e688c2da9ced7/root-review.json`
records that bounded publication; no staging or canonical archive admission is claimed.
The bounded performance assessments and all adverse/unknown observations remain intact.
Next: freeze this source, build full Release, serialize protected Interaction, Gallery
and Commands, then fresh Full Debug/Release with every used receipt epoch retained.
Canonical pre-stage archive validation, visual review, final governance and Done/index
closeout remain required. The user-authorized continuous protected testing interval is active.

2026-10-03 22:24 UTC — current protected-test repair: Release Interaction
`20261003T220516Z-110824-fc06eddbd27d45dea5ec448adcd283ac` closed FAILED 0/1/0
on receipt `c337bd5bdea9b6bc9d7dea1ae94152027881efa31618fc93a327798279d6ff83`.
Routed PageDown, exact Unicode text, selection/caret and actual text/native focus passed
both 144-to-96 and 96-to-144 transitions. Tab and Shift+Tab passed; the subsequent
pointer ownership guard correctly refused input. Root and the read-only cheaper reviewer
found that static placement explicitly lowers the popup after initial activation.
The protected fixture now restores owned z-order once after all text witnesses and
Escape restoration, with pre-existing native/foreground ownership required. Pointer
geometry/hit diagnostics and a second pre-mouse-down ownership check retain strict refusal.
No static capture or production focus policy changed. These edits require a fresh build.

Static Gallery invocation `20261003T221259Z-113388-14b766c3fc5547d3ad404054ba2b3ba4`
was BLOCKED before native execution: the focused 133-output build was not a compatible
full-solution receipt for `-SkipBuild`. Its original packet remains retained. Next action:
build the test-enabled full x64 Release solution, then serialize protected Interaction,
static Gallery and Commands with that receipt; run fresh Full Debug/Release and retain
each used receipt epoch before any later build. Complete publication and normative/
governance closeout remain required. The current user-authorized testing interval remains
active through the existing warning, desktop lease and restoration guards.

2026-10-03 22:05 UTC — final testing continuation: fresh Release receipt
`669816b069eed0f111e675d8b9a5d16d9c7c9af457a17398bb86451956a08a04`
built with zero warnings/errors. Gallery run
`20261003T215545Z-69436-254c207b64614fd39b11a78a6e320324` remains FAILED 0/1/0:
the repaired static path workspace passed real control-handler PageDown (viewport
0 to 1, exact read-only text and unchanged native focus) and both physical DPI hops
with exact selection/logical focus. The next scope-provider witness rejected the
synthetic conflict's zero generation. Root retained the complete failed C archive
and original foreign 265 audit issues. The fixture now supplies deterministic nonzero
prompt generations; production generation fences and UIA assertions remain unchanged.
The protected text-focus extension uses the real guarded keyboard route, separate
text/action DPI receipts, nonempty interior selection/caret preservation, exact row
restoration and original action context. Independent cheaper-agent review found a
buffered evidence-writing gap; all final interaction streams are now flushed and
checked before success. These last fixture fixes are uncompiled and require a new
Release build, protected Interaction, Gallery and Commands before fresh Full gates.

Closed F6/v2 originals were byte-preservingly curated externally (all 16 originals).
Root reviewed all 108 F6 intervals, all 290 v2 absolute resource phases, comparison
arithmetic and lifecycle source. `final-run-review/ROOT-F6-V2-BOUNDED-ASSESSMENT.md`
records bounded acceptance with every adverse/transient/resource/dependency limit
retained; no CPU/memory neutrality, process/GPU leak proof or retroactive v1 pass is
asserted. F6/F7 publication, final native/visual/governance and fresh Full Debug/Release
remain required, so their checkboxes stay open. The current user authorizes computer
testing now through the existing warning, desktop lease and ownership/restoration
guards. Root remains sole checkout writer/build/test owner.

2026-10-03 21:40 UTC — current continuation: the six-process F6 counterfactual
and ten-process integrated v2 comparison completed on frozen Release receipt
`256d18be002b07877b2e00cd18181d35fc30a3f3a702e6c436b8f8fc1ff35de7`.
All originals, adverse observations and failed earlier attempts remain separate.
Root and the cheaper read-only reviewer checked comparison arithmetic; resource
and timing interpretation is still a gate, not an automatic pass. V2 dependencies
differ, so it is a whole-product fixed-fixture comparison, not isolated native causality.
Protected interaction passed 1/0/0 on both physical displays, with actual 144→96
and 96→144 DPI changes and the same focused action preserved. Commands family
run `20261003T212008Z-13968-6e8bc8ad3ed446a7a59b40a88564d9d5` failed 64/1/0:
its taskbar fault hook was armed after the real synchronous cached-model replay.
Gallery `20261003T213304Z-92580-83d34d6d5ae8477d8af6abb6bd48a641` failed 0/1/0:
it sent PageDown to a deliberately unfocused, non-activating editor. Both complete
C archives have unchanged foreign audit findings; both original runs remain failed.
The test repairs arm the taskbar hook before its message and verify backoff, and
make static scrolling use the real TextField handler with explicit logical-focus
provenance. Protected Commands retains actual routed PageDown/focus assertions.
Production taskbar replay and native text-focus policy are unchanged.

The Full runner also refreshed its cached build-receipt binding too late after an
artifact writer. A retained metadata-only production-AST regression failed before
the repair. The digest now refreshes from the byte-verified restored receipt before
downstream fingerprints and decision publication. Tests, owning validation spec,
tool inventory and usage documentation change together. These latest test/tooling
changes are UNCOMPILED/UNQUALIFIED until the new builds and runs below finish.
Next: focused Pester, fresh Release gallery/Commands, fresh Full Debug and Release;
retain all used immutable receipt epochs before another build, then final archive,
visual/governance review and same-plan Done/index/automation closeout. Root remains
the sole checkout writer and build/test owner. The user has authorized computer
testing now; every input lane still uses its warning, desktop lease and restoration.
2026-10-03 — D10 review corrected the nested child's native run-ID parent-PID field,
whole-run retention through an identity-checked same-volume rename into parent artifacts after
natural exit, and reacquisition of parent HWND/FolderWindow/provider after message pumping.
The storage/lifecycle catalogs now describe actual fixture cuts and independent result axes;
they do not claim runtime qualification. The evidence-only eight-unit /WX check is running
before the next full-solution build. Original failed receipts remain immutable. No foreground
session, library adoption, commit or unrelated checkout mutation has occurred in this continuation.

2026-10-03 — latest qualification/resume: the complete test-enabled Debug build passed
with zero warnings/errors, receipt `fd2ff784e5e4104fa50ada16bf1042cb7236cd4020080be7238eb081c212ca86`.
The next linked Safety first repetition passed actual D1 post-child-delete cancellation and
D8 host case-only/provider temporary rollback, F2/F4/F5/D3/D9 and all eighteen F6 states,
but failed D4 and D6 fixture expectations. Preserve the full original failed packet at
`C:\RedSalamander.Perf\evidence\i27-followup-20261002-0e139fb\resume-20261003T121349Z\failed-safety-retained-first-whole-run`.
D4 wrongly rejected the real Exists prompt whose Replace action is withheld by an actual
exclusive destination lock; D6 wrongly expected ERROR_PARTIAL_COPY instead of the specified
S_FALSE aggregate for accepted child Skip. Reviewed corrections preserve all storage and
security proofs, add observed prompts/receipts, and retain failed diagnostics/trees. The outer
repeat controller tried to restart retained scratch and stalled; only its exact owned
launcher was stopped, with before/after process evidence retained. New owner-thread
HasRetainedFailureEvidence stops the aggregate rather than starting a refused repeat.
This is a failed run, never a repeat-three pass.

The parent independently reviewed/applied D10's private child-only actual production close,
normal settings-save/destroy receipt, real positive unpublished write gate and post-exit
closed byte/hash tree comparison; the contained runner gains child-only Unicode environment
overrides and actual/pure inheritance tests. F7's true directory volume-mount matrix and
aggregate alternate-root preservation are applied; failed F7 observations retain both trees
after safe unmount. These revisions are under fresh compiler/build/runtime qualification;
no DxUi, I26 baseline or unrelated repository change occurred. Final brief warned input-session
availability is explicitly confirmed by the user; static work remains non-activating.
Next: finish compiler check, freeze source/spec and rebuild Debug, then linked Safety/D2/D11/D10
and strict v2 fixture with whole-run and nested child retention before every runner GC.
Retain real Release F7 admission baseline before optimization; identical v2 baseline/candidate
pairs and final whole/native/performance/docs gates still prevent Done.

2026-10-03 — reviewed D1/D2/D4/D6/D8/D11/F7 bundle is written and its affected
translation units pass an evidence-only MSVC Debug /WX compiler check. The retained
compiler command now redirects object/PDB output only into the owned evidence folder,
uses the existing tlog source exactly once, and retains unique log names. The prior
wrapper failure accidentally used original intermediate output paths; those artifacts
are UNATTESTED and require the following full-solution build before any linked test.
Two actual optional content-proof calls were corrected to QueryInterface the owning
IFileSystemBoundContentProof interface; narrowing numeric environment text now uses
canonical UTF conversion. Compiler failures remain recorded, not runtime passes.

D6 uses the actual ordinary-child DACL and serial native Local route, with original
security restoration. D2 uses actual C/D Managed Move, first publish/verify/delete then
second real positive stage-write failure; the old Dummy case is now truthfully named
FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof. D11 registers two independent
non-clipboard commands with real Queue-after, then Copy Exists/Skip or Move SourceMissing.
Parent reviewed actual callers, all typed axes, callback lifetime and storage evidence;
failed D11 witnesses preserve primary/alternate scratch after quiet. F7 instruments actual
UI admission/provider calls/allocations without changing cache/capability policy. All new
included chunks are listed as None in both projects. None of these written witnesses is
runtime-qualified yet. D10 and a true directory-volume-mount F7 route remain reviewed
agent proposals in progress. No DxUi pin or unrelated checkout changed.

Next: freeze this bundle, full test-enabled x64 Debug build, linked repeated actual-storage
qualification and strict v2 rendering fixture. Archive every whole run before runner GC.
Then finish D10/mount coverage, retain real Release admission baseline before optimization,
repeat same-fixture paired performance and final Debug/Release Full/native/docs gates.
The final protected desktop availability question is pending; non-activating work continues.

2026-10-03 — RESUMED qualification receipts and next implementation bundle. The test-enabled
x64 Debug full build passed with zero warnings/errors, receipt
`33615353462243d3e7bd5f6ab41c9ed448e74dc04d3193317dc400366aaf5b6a`.
Whole linked runs were copied before every subsequent runner cleanup beneath
`C:\RedSalamander.Perf\evidence\i27-followup-20261002-0e139fb\resume-20261003T121349Z`:

- Safety repeat-three `20261003T123855Z-107928-1dacee73320245ffb6561bc2d7f95b14`
  failed only the first informational-completed-idle-visible F6 interval. F4's five posting-failure
  cuts and secondary-owner destruction, F5's actual `task.result`/`E_UNEXPECTED` selection and
  retained bytes/hash, and D3's four actual storage/identity cases passed individually in all three
  repetitions. These bounded items are checked above; fresh whole Debug/Release Full qualification
  remains required. F6 observed bounded queued UI motion during the fixed warm interval: three
  ticks, five snapshots, three rate updates and five taskbar attempts, then interval zero. The
  failure remains retained. Its repair waits for an actually stopped timer within a bounded
  settlement period before the unchanged 650 ms idle observation; it does not stop/reset counters.
- D5 direct repeat-three `20261003T124200Z-94868-df0c43076fb04fb9bcd5c56960dd4a5d`
  and alias repeat-three `20261003T124329Z-42960-796739dd7ae44507b64267bc50b41097`
  passed complete owned tree/byte/hash preservation.
- D7 Local repeat-three `20261003T124455Z-82672-5de7b4df892d4be1839ce59b41c03131`
  and bridge repeat-three `20261003T124656Z-55376-1e11cdf64d484170bb08c20b520c472d`
  passed exact copied-byte progress and retained source/occupant contents. The cuts are Local
  sharing-violation Skip and bridge directory/file-conflict Skip; their roles stay explicit.
- Strict-render performance repeat-three `20261003T124801Z-77972-7f086da96de14213a7920a1b25f8a0d6`
  failed legitimately: queued auto-sizing changed the popup's measured viewport and the scan's
  changing path range added a brush-cache entry after warming. Those results cannot establish a
  paired throughput comparison. The proposed v2 benchmark holds only its synthetic viewport,
  prewarms the complete measured path range and still requires exactly eight rendered frames and
  unchanged geometry/resize/cache/surface counters. Ordinary popup auto-sizing stays enabled.

The parent has reviewed and applied D1's real serial child-removal cancellation barrier, D4's
exclusive-destination and actual owned-stage write-failure witnesses, and D8's distinct host/provider
case-only storage witnesses. F6 settlement/metric-unit corrections and the v2 benchmark repairs
are also written. All changes after the above build are UNCOMPILED/UNQUALIFIED. D2, D6 and F7
remain evidence-only agent proposals until parent review/application. The parent is the only
checkout writer and build/test owner. Next: complete/review that bundle and remaining D10/D11
storage witnesses, compiler-check, freeze exact source/spec, rebuild and repeat linked qualification;
then identical v2 baseline/candidate performance pairs and the required whole/native/docs gates.

2026-10-03 12:13 UTC — HUMAN RESUME: continue all required work to verified Done. Source/patch/hash
backups at `pause-20261003T114728Z` remain intact. No source changes occurred during the pause.
First action is the precise F5 Issues/HRESULT/foreground failure; the parent owns all implementation
and builds/tests. Cheaper agents are preparing read-only diagnosis and case-owned proposal patches
for parent review, with no competing checkout writers or builds. The prior successful Debug receipt
and failed aggregate remain unchanged evidence. No gate is waived or plan completion inferred.

The parent confirmed F5's selected-row mismatch against the real diagnostic JSONL and the
production projection: successful consent precedes `task.result`, and generic task selection
chooses the first matching row. The fixture now selects the actual category/HRESULT diagnostic
through an ENABLE_TESTS hook that reuses the same retained grid selection path; production task
selection semantics are unchanged. Separate category, exact HRESULT and foreground assertions,
plus wrong task/category/HRESULT selection-refusal controls, retain stronger evidence.
The parent reviewed/applied Sol's D5/D7 proposals (actual direct/alias self-copy tree hashes and
closed inventories, Local/bridge Skip bytes and retained contents) and Luna's exact-eight-render
performance-fixture guards. These edits are unqualified until compiler/build/runtime checks pass.
The historical 29% regression predates F6; its new per-paint policy must not be presented as its
cause. Repeat matched exact-render-count baseline/candidate samples before selecting an optimization.

2026-10-03 11:47 UTC — USER-REQUESTED PAUSE. All implementation/review agents have completed;
no task-owned build/test/application process remains under `C:\RsI27Followup`. Preserve the
30 changed source/spec files, including the untracked SameObject fixture, without resetting,
committing, pushing or changing other owners' checkouts. The existing autostash
`ed7d9a5b7c6bb2707e5c3c30361999f1208eab7a` remains retained.

The latest full test-enabled x64 Debug build passed with zero warnings/errors, receipt
`87a0c42d2a7918e2f750c1a7e73abe2403681afbcdc0ba1c6f68d15d82bef83a`.
The repeated linked Safety run `20261003T113435Z-84672-73a53c3d0ee64b6789b450377529603f`
FAILED in all three repetitions. Its only reported individual witness failure is F5:
`selected actual Issues row lost E_UNEXPECTED or changed foreground`. The failure does not
yet distinguish the selected-row HRESULT assertion from foreground preservation. Do not
weaken either assertion or assume the implementation versus fixture cause. D3's four actual
same-object/Keep Both/hard-link byte/hash/identity/inventory witnesses, all five F4 all-posts-fail
recovery cases, real secondary-owner callback teardown, D9 and all eighteen F6 states passed
individually in each repetition. Aggregate failure remains failure. Candidate idle intervals
recorded zero timer/snapshot/rate/taskbar work; active hidden/minimized telemetry retained
100 ms scheduling. Untimed Explorer replay passed with one successful COM submission and no
snapshot/timer rearm. Resource/CPU comparison and broad regression qualification remain open.

The complete failed run was copied before runner cleanup to
`C:\RedSalamander.Perf\evidence\i27-followup-20261002-0e139fb\safety-f4-f5-f6-d3-debug-packet`.
The same evidence root retains the complete previous baseline packet, logs, build/source
receipts and two unapplied, uncompiled agent proposals: `d5-d7-proposed.patch` and
`integrated-perf-count-proposed.patch`. Parent must review their entire diffs and production
callers before applying them. No new library pin, code qualification checkmark or Done move
is implied. The explicit pause snapshot and detailed next actions are in
`pause-20261003T114728Z/RESUME.md` beneath that evidence root, with all changed file copies,
hashes, Git status and a binary diff against the exact HEAD.

Next action AFTER explicit human resume: inspect F5's actual selected Issues row/cell values
and foreground evidence at `RunFileOpsCommittedRetainedDeleteChecks` (SelfTest.cpp, around
line 7338); add precise diagnostics if needed and repair the actual cause. Then independently
review/apply the bounded D5/D7 and exact-render-count proposals, compiler-check affected TUs,
freeze source/spec, rebuild and repeat/archive linked qualification. Retain F6 paired baseline
and candidate evidence. F7 needs real admission instrumentation/baselines BEFORE optimization;
D1/D2/D4 and D6/D8/D10/D11 still need actual filesystem outcome witnesses. F1/F3 native Commands,
fresh Debug/Release Full, French/mixed-DPI/focus/UIA, the unresolved graph-redraw performance
regression, governance, final evidence/domain/index/archive work all remain required.

The desktop availability question is still unanswered; no focus/input session is authorized
by elapsed time. The recurring continuation `implement-file-operations-ui-to-done` pause
could not be confirmed: its view rendered an app card without saved editable fields, and the
two partial pause updates were rejected by argument validation. Do not manufacture its missing
name/prompt/schedule or claim it disabled. Regardless of scheduler state, this explicit human
pause takes precedence; any wake must remain idle until the human resumes.

2026-10-03 11:20 UTC continuation: the parent reviewed the four actual Local D3 fixtures written by
Sol High, including normal admission, canonical binding/revalidation, exact endpoint spelling,
before/after bytes/BLAKE3, closed inventories, summary axes and quiet cleanup. Same-path Copy is
specified to Keep Both; it does not need a new conflict policy. Hard-link refusal revealed an actual
reporting defect: the qualified-result adapter hard-coded possible mutation before the identity
guard allowed execution. The parent now carries Preparing/MutationPossible at actual mutating
provider, shell or bridge entry after live-output/breaker/source-revalidation gates, retaining
conservative truth across retries and the resolved destination on early refusal.
Runtime qualification remains open. The parent also corrected the delegated fixture's ignored
cleanup/settings-save results; normal success now requires both quiet points.

F4 secondary teardown and F5 Issues fixture repairs are written: the former now observes state
without recreating it; the latter expands actual technical columns and checks the HRESULT column.
The complete failed baseline packet remains retained. F6's unchanged eighteen timed intervals have
retained baseline counters and resources; its idle policy is written but unqualified. A read-only
Luna review found that progress interpolation could raise active telemetry to the motion cadence.
The parent fixed precedence so active work keeps 100 ms except for the existing explicitly armed
motion deadline. A separate, untimed Explorer taskbar-cache replay witness follows all measurements.
No foreground session has been run while the desktop availability question remains unanswered.
Next: compile the frozen F4/F5/F6/D3 bundle, run/retain repeated linked safety qualification, compare
the timer measurements with their retained baseline and continue F7 and the remaining destructive
proofs. F1/F3 native Commands, exact-render-count performance repair, fresh Debug/Release Full and
final normative/evidence/archive/index gates remain required. No unchecked gate is waived.

2026-10-02: verified current source and retained I27's open Full/foreground/performance gates.
Created the isolated follow-up checkout above. One `gpt-6.1-sol` High agent owns only F1's
popup implementation/header, existing Commands fixture and failure-surface normative paragraph;
the parent owns this plan/index, independent code review and all builds/test execution. No
competing code writers/builds are authorized. Review must inspect production callers, keyboard
metadata, native Space semantics, actual consent effects and fixture isolation, then execute
focused native tests; agent self-reports are not qualification. Existing unrelated builds may
constrain runtime/performance timing and must not be killed or used as uncontended evidence.
2026-10-03 checkpoint (continuation began 2026-10-02):

- F1 implementation is written and independently reviewed by the parent, including the real
  main-loop dispatch and native child/consent callers. Enter opens Cancel-all consent on release;
  repeated Enter/Escape cannot invoke it, native Space retains its standard release behavior,
  and focus loss/hide/reopen/destruction clear the arm. Existing host-attach and D2D-failure
  Commands cases assert declined consent preserves the task and exact source/destination bytes.
  MSVC compile-only checks of the actual Popup.cpp and Commands.SelfTest.cpp with `ENABLE_TESTS`,
  `/std:c++latest /W4 /WX /permissive- /sdl` passed. Linked/native runtime remains unchecked.
- F5 is implemented by the parent and independently reviewed by a `gpt-6-luna` High read-only
  agent. Successful known committed Delete with a retained selected root now fails with
  `E_UNEXPECTED`, preserving `NotAttempted`/`Retained`; explicit Skip/cancel, provider partial
  failure, late task cancel, cleanup debt and Copy/Move semantics keep their precedence.
  The actual pure repository reducer matrix, compiled against the actual production header,
  fails against the initial master header on this regression and passes against the candidate.
  The callback/finalizer matrix also covers the committed-retained Delete receipt. The linked
  `FileOps_SafetyRemediationContracts` case now passes three repetitions, including real Task
  callbacks/finalization and the existing independent Issues/lifecycle controls. The FileSystem
  domain contract is updated in place; a dedicated hostile-receipt-to-Issues witness and broader
  filesystem/platform gates remain unchecked.
- F3 implementation is written by the same `gpt-6.1-sol` High agent and independently reviewed
  by the parent against all confirmation callers and endpoint-contract/lifetime helpers.
  Captured qualified endpoint path rules govern formatting; unknown/mixed domains use a bounded
  multiple-locations summary, and displayed raw spelling remains unchanged. The parent added an
  exact English/French resource assertion so a missing French satellite cannot silently pass the
  new prompt matrix using English twice. Remote Delete coverage is still formatter-only; the
  actual native permanent-Delete card control also checks Unicode spelling and Cancel defaults.
  `C1_PermanentDeleteConfirmsOnCard` passes two nonactivating repetitions in the linked application.
  The provider prompt matrix remains a Commands test behind the existing foreground deferral.
  No additional provider query, mutation-route change or archive cleanup change was introduced.
- Full x64 Debug build on the initial base failed before application compilation in
  `aws-sdk-cpp` 1.11.880: the restored CRT lacked `Base64EncodedLength`/`Base64DecodedLength`.
  The separate recovery owner has since committed master `fbe113a0f747f89af32d6fb337e4e67b9f5bdd6b`,
  including the matching AWS CRT 0.43.4 override. The isolated checkout now fast-forwards to that
  committed base, retaining all task changes. The sole index conflict was resolved by preserving
  both our revised I27 row and upstream I28. All changes are unstaged. The exact task-owned
  recovery autostash `ed7d9a5b7c6bb2707e5c3c30361999f1208eab7a` remains available until validation;
  unrelated shared stashes must not be dropped. The ordinary full-solution x64 Debug build on
  this base passed with zero warnings/errors, receipt
  `e6a47369062214078d29faff55efd9f2676f15f442f9995f2da3587787bebea8`.
  Do not copy uncommitted recovery edits or infer qualification from an older owner's binary.
- Evidence is retained at `C:/RedSalamander.Perf/evidence/i27-followup-20261002-0e139fb/`:
  full build/spec logs, parent review receipts, reproduced baseline/candidate reducer source,
  MSVC response files and build/run logs, source hashes and the compile-only popup/Commands logs.
  Parent formatting with discovered clang-format 22.1.3 and `git diff --check` passed on F1/F3/F5.
  Retained `safety-debug-packet` contains the successful three-repeat linked safety result;
  `permanent-delete-debug-packet` contains the successful two-repeat consent result. Each packet
  includes the runner receipt, actual case JSON and traces. Runner stale-run cleanup removed the
  first successful safety packet before archival, so that same fixture was rerun and archived;
  do not cite its removed run directory as retained evidence. The initial nonnumeric shuffle-seed
  invocation was rejected before cases ran; its failed log remains. Both successful runs use
  numeric seed `20261003`. The sandbox audit separately reports 265 historical, unattributed run
  directories; focused runner exit zero does not qualify that outstanding governance finding.
  No application screenshot or desktop input was substituted.
- Spec inventory still has seven existing missing-link findings in I26's support material:
  the interaction-v22 archive README and six historical ignored evidence logs in validation.md.
  No new blocking I27/index finding was reported. Retain the original findings; do not invent logs.
- D1 preparation: a cheaper read-only audit, checked against the production walker by the parent,
  confirms no existing provider barrier after child removal. Local Delete uses handle disposition
  (`DeleteOpenedPathNoFollow`/`SetFileInformationByHandle`), and parallel child batches can mutate
  siblings concurrently. The next byte-preservation witness needs a case-owned post-removal
  gate, explicitly serial provider concurrency, before/after sibling content checks, and cleanup
  that releases the gate on every exit. Existing host preflight gates and successful deep-tree
  deletion do not establish this cancellation outcome. D1 remains unchecked.

- F2 source implementation is written by Sol High and independently reviewed by the parent and
  a read-only Luna High agent. Durable cancel is published before the phase; terminal results
  win; unfinished canceled tasks have a neutral localized Stopping status, distinct footer count,
  retained active taskbar/counter/history facts, no ETA, and withdrawn task/conflict decisions.
  Stale pause is guarded under the existing pause mutex, and stale speed/destination/queue
  callbacks are rejected. Global Cancel-all availability includes hidden engine work; the parent
  replaced the new per-layout borrowed-task vector with an allocation-free ownership-locked
  atomic predicate. Native failure-surface Cancel all is unavailable when all work is stopping.
  Source review repaired host-lifetime borrow ordering and optional access before qualification.
- The F2 regression uses the existing metadata pause to hold a real Local Copy unfinished after
  RequestCancel, checks all eleven stale task decision kinds, and proves exact source/destination
  bytes remain unchanged. Negative UIA action checks require a successful bounded tree traversal
  plus a positive Stopping status-name witness. A second real Copy is held in hidden Preparing
  under a shared-owned finite readiness gate: it stays absent from popup snapshots yet enables
  exactly one native Cancel-all consent and receives cancellation. Both gates are released before
  bounded draining on every exit; both terminal results and four exact byte references are checked.
  Literal resource labels cover all five cultures, including French. These new checks are written,
  not yet run. Resource localization checks and the new full Debug build are in progress.
- Parent formatting now uses discovered clang-format 22.1.8 (the installed version changed),
  with dry-run/Werror and git diff --check passing. The version and complete review are retained
  in format-f2-version.log and F2-parent-review.md in the same evidence folder. Source changes are
  frozen for the next full-solution build and linked no-activation regression run; no older binary
  or source-only agent report can establish the new F2 behavior.
- F2's first full Debug build failed before linked qualification: new IsTaskDecisionHit produced
  nineteen C4061 warnings for excluded enum cases, and the test compared typed taskStatusKind with
  a uint32 cast (C2676). The parent replaced the default-only exclusions with explicit enum cases
  and compared the typed enum directly. The original failed build-debug-f2.log is retained.
  Fresh formatting passes; resource-localization-f2-retained.json reports 9/9 passed. The first
  resource run also passed, but its shell redirected the typed result before capture; the corrected
  invocation preserves the actual Pester object and JSON. Source is frozen again for a full build;
  these repairs do not count as an executed application pass.
- F2's repaired full-solution x64 Debug ENABLE_TESTS build passed with zero warnings/errors,
  attestation `d98b8a499b42c447f78b36c68031670b03de33cc83fa603fa9f85c14b14ed258`.
  `FileOps_SafetyRemediationContracts` passed all three repeats (6719/7078/6875 ms), including
  the actual unfinished RequestCancel snapshot, hidden-work native Cancel-all consent, positive
  stopping UIA witness, absence of all eleven stale decisions, exact bytes and terminal results.
  `C1_PermanentDeleteConfirmsOnCard` passed both repeats (438/375 ms) on that same binary.
  Both complete runner packets were archived immediately as `safety-f2-debug-packet` and
  `permanent-delete-f2-debug-packet`; runner exit codes are zero, with respectively 9/9 and 6/6
  cases including setup/cleanup. The source hashes still matched the repaired freeze at execution.
  Only the checkpoint is updated afterward; fresh attestation may be required before later gates.
  F2 is checked for this bounded implementation/runtime coverage, while the original broad
  Full/foreground/performance/platform requirements remain unchecked. Native physical-key/modal
  handoff, external accessibility clients and the provider Commands matrix are not claimed.
- The parent verified Luna's F4 preparation against production posting, transfer, shutdown and
  presentation-timer callers. Recovery must seed an owner-thread active-work heartbeat at admission,
  survive absent/hidden/minimized popups, and avoid dependence on any successful message post.
  A transferred/nonjoinable Task thread is not proof of exit: publish fallback eligibility only
  after an off-thread join, and preserve pending entries on failed reaper submission for a safe
  owner-tick retry. Early wakes must not consume pending entries; successful initial fallback wakes
  must also reach the retry/join protocol. Retain shutdown's outstanding-reaper lifetime fence.
  The next real regression must force every wake to fail and assert one owner-thread callback,
  one summary, no retained task/fallback, unchanged bytes and no owner-thread join while parked.
  Preparation is retained in `F4-parent-checked-preparation.md`; F4 remains unimplemented/unchecked.

Next: implement F4's owner-thread fallback recovery and deterministic all-posts-fail witness, with
explicit post-join readiness and failed-submission retry ownership. Sol and Luna have stopped
writing/reviewing; the parent has independently checked their work and owns subsequent review,
builds and runtime receipts. Reuse the existing pause/failure hooks and add only missing bounded
test controls. Do not mistake taskFinished or a transferred thread for completed worker teardown.
Native fallback and provider prompt Commands tests retain their foreground deferral.
Independently launched external builds remain untouched and make performance timing unqualified
while they run. F4/F6/F7 and D1-D11 remain
unchecked; prior Full/foreground/29% graph-performance gates remain open. Neither compile-only
evidence nor a pure reducer pass qualifies the whole application or moves I27 to Done.

## Evidence

2026-10-03 retained baseline: full test-enabled Debug build passed with zero warnings/errors,
receipt `3f155779d46786308a709fbfebbbbe5969eb0965de87391c42a44aed65b2c02b`.
Run `20261003T104811Z-75708-2d45bfda559b4f38aea6f6b52bac66dc` remains FAILED in all three
aggregate Safety repetitions. Individually, D9 and all five F4 posting/refusal/quiet-point
scenarios passed each repetition. The complete F6 unchanged-policy measurement passed all
18 display/work states per repetition: 54 counter/CPU/resource samples are preserved in
`popup-telemetry-baseline-results.json` and the complete archived run packet. Idle states
still produced roughly 5-6 ticks/snapshot/taskbar attempts per 650 ms. Observer reads were
scalar-only; CPU granularity and the real bounded worker-barrier contribution are retained.
F5's real receipt/summary/byte checks passed before its Issues witness incorrectly requested
hidden technical columns and confused localized status with the numeric hresult column.
The secondary-owner witness recreated released state by calling an initializing debug getter.
Parent and cheaper read-only review independently traced the real post-join dispatch/shutdown
quiet point; its state-release assertion stays intact and now uses a noninitializing getter.
The Issues witness explicitly expands technical presentation and asserts the numeric hresult.
The timer fixture now verifies successful cleanup and restored-settings persistence, in addition
to its early-exit restoration scope; timed measurement phases are unchanged.
F6 idle scheduling is now written in the consumer, with independent owner recovery, bounded
visible motion, active hidden/minimized updates, transition rearming and cached Explorer replay.
It is not yet compiled/run or marked complete. No DxUi change/pin update was made. Next: review,
targeted precompile, fresh full Debug build and repeated linked qualification of the repaired
witnesses and paired F6 candidate; native/full/performance and D1-D8/D10-D11 remain open.

2026-10-03 next bundle: the linked F4 Debug run
`20261003T082353Z-48304-0d2350d7d9554db98f0615727dd937d9` failed all three Safety repetitions.
Recovery scenarios 0-2 passed, but scheduled Rename used a different Local instance from its
owner pane; the secondary-owner fixture omitted its Local pane routes; retained idle-measurement
windows polluted the next repetition. All failures and full packets are preserved under
`C:\RedSalamander.Perf\evidence\i27-followup-20261002-0e139fb\safety-f4-debug-failed-packet`.
Sol repaired those fixture endpoints and cleanup without waiving any assertion; the parent
checked the actual admission guards, callback cleanup and owner lifetime. Fresh runtime is required.
A second Sol slice added F5's real PermanentDelete hostile receipt witness: normal consent,
Failed/Retained completed summary, retained task.result diagnostic, authentic Issues status,
exact bytes and canonical BLAKE3. Parent review prevented private-pane teardown from saving
test layout into application settings. This witness is unbuilt/unrun so F5 remains unchecked.
Before F6 policy optimization, the parent replaced observer-induced full snapshots with scalar
telemetry counters and extended the unchanged-policy fixture to naturally pumped initial-idle,
real running/terminal-owned Copy, completed-idle and informational-running/completed states,
each visible/hidden/minimized. It records task/rate/timer/taskbar counters plus CPU/memory/handles
and host observations; taskbar attempts and successful COM submissions are distinct. Its real
worker barriers contribute process CPU and must remain identical for paired runs. No policy
optimization or new pass is claimed. Next: format/freeze this bundle, full test-enabled Debug build,
three repeated linked Safety checks and retain complete F6 baseline before changing timer policy.
The first F5/F6 bundle build failed on two delegated fixture API errors (WIL query_to returns void;
the UTF conversion needs the existing canonical wrapper). Parent repairs reuse try_query_to and
Utf16FromUtf8ForFileOpsSelfTest; the failed build log remains. Sol added D9's production empty
Copy/Move/Delete admission fixture, checked independently against the real early rejection branch:
no task or task-history change, exact distinct sentinel bytes/BLAKE3, closed directory inventory,
no prompt/payload/alert/host change and preserved foreground. D9 is not qualified until runtime.
The timer fixture also fixes/then restores auto-dismiss to keep its workload deterministic.
The repaired full build then rejected D9's missing local Task alias and aggregate array braces;
the parent fixed both and compiled the exact SelfTest.cpp with the discovered project's MSVC
arguments, redirected outputs, and warnings as errors. That compile-only check passed. The
two failed full-build logs and source freezes remain; neither is counted as qualification.
Current next action: freeze this repaired bundle, run a fresh full test-enabled x64 Debug build,
then run and archive three linked Safety repetitions. This qualifies the new F4/F5/D9 witnesses
and retains the complete unchanged-policy F6 baseline before any idle-policy implementation.
A cheaper read-only agent is reviewing the baseline fixture; the parent retains all write/build
ownership and reviews its findings. An independently launched RedPrism ARM64 build is untouched;
any overlapping performance timing must be labeled contended and repeated without that load.

2026-10-03 continuation: F4 is implemented but not yet qualified. A synchronous
owner-window heartbeat is admitted before both ordinary and scheduled-rename workers;
fallback results remain pending until a background reaper completes the exact worker join.
The owner heartbeat retries refused reaper admission and consumes ready results once, even
when every completion and presentation-refresh post is suppressed. Stale nonzero payload
tokens cannot consume fallback results. A targeted Astra High review found and the parent
repaired scheduled Rename's separate admission and raw Task use after owner-destroying
callbacks. Completion dispatch now re-finds the task after each external callback.
New linked nonactivating controls cover five real Copy/BatchRename cases, parked terminal
workers, both transfer/refusal paths, failed presentation posts with no preexisting popup,
early and stale wakes, owner heartbeat responsiveness, one callback/summary and exact bytes.
A persistent threadpool-admission refusal control additionally permits owner join only after
a zero-time wait proves the exact native worker handle already signaled; a parked worker
cannot grant this exception. A separate real secondary FolderWindow is destroyed by its completion callback; the primary
owner must survive. New test-only popup tick/snapshot counters and unchanged-policy idle
measurements precede F6 optimization. The first full Debug build passed with zero warnings/errors
and attestation `2866e9d876071862b818da5304929f6157c1b5b22739211aaf0a9c55465f66a5`.
Before running its new fixture, targeted review caught failure cleanup retaining a callback after
its captures expired; the parent repaired that scope and strengthened production shutdown's
recursive callback fence. A second targeted review found no additional concrete defect in the
repairs or native-handle quiet-point route. Formatting passes; a fresh build/runtime of the
repaired source is next. Popup snapshot counter reads include the two observer snapshots;
they are not claimed as pure timer rebuild counts or CPU/resource non-regression proof.
No new runtime pass or F4/F6 closeout is claimed yet. Full/native/performance and D1-D11
remain required. The user's latest request authorizes continuing all required qualification;
desktop availability has been asked separately while nonactivating work proceeds.

The user authorized committing and merging the implemented changes on 2026-10-02 while
qualification remains open. Final R21 x64 Debug and Release builds have zero warnings/errors;
governance passed 239/239, and the held-key safety case passed three repetitions per build.
Six isolated ViewerPE close checks passed without changing the 500 ms limit; the original
R19 Debug PR timing failure remains failed. The reviewed evidence packet preserves all
6,325 integrated claim rows from 23 captures and passes the archive limits. Performance
non-regression remains unqualified: current interleaved graph-redraw medians worsened
about 29%, with attribution unresolved. Foreground checks and fresh Full runs remain
deferred by the user. Publication does not move this plan to Done or clear these gates.

The final source review found a further held-key defect: after B replaces A's published cancel
binding, repeated Escape could cancel B with no fresh press. The consumer now rejects repeated
Enter/Space/Escape/task-Delete before hosted dispatch. The existing non-activating popup fixture
adds native repeat/fresh-press/navigation controls; the real A-to-B Commands assertion remains
pending the user's foreground-test deferral. The R19 broad FileOps and Compare suites passed
in Debug and Release. Release PR passed; Debug PR recorded one ViewerPE close-time miss whose
functional blocked-read and module-lifetime assertions passed. Its unchanged 500 ms assertion now
prints the measured duration for isolated diagnosis. Those R19 receipts predate this final repair.

The second fresh Debug Full run (`20261002T005651Z-55492-bc32144485214f708ccd42efaa3e8a9c`)
failed: 2,110 passed, 55 failed and 96 skipped, with 18/20 promoted entries. Its failures exposed
parallel Local Copy receipt aggregation and Issues selection integration defects, a leaking
clipboard-ready fixture, cold-index UI fixture assumptions, and stale action lookup/unsupported
Retry fixtures. Repairs and original outcomes are recorded in the
[qualification report](../../TestRuns/4cb089111a23/Validation/2026-10-01_fileops-safety-remediation/ReviewRemediation.md).
Focused follow-up found incorrect DPI units in header-input fixtures, a stale generation in the
acceptance fixture's next-prompt Cancel lookup, and a real Local Native Move receipt that falsely
claimed retention after failed source validation. The fixes preserve actual sort transitions,
exact prompt identity and unknown source truth, with producer/callback/finalizer and indexed
single/bulk controls. The [qualification report](../../TestRuns/4cb089111a23/Validation/2026-10-01_fileops-safety-remediation/ReviewRemediation.md)
records the boundary between the executor's earlier source guard and these receipt controls.
A final adversarial check found a successful empty-directory Copy merge being classified as
`E_UNEXPECTED` because the reducer equated completion with mutation commitment. A clean, exact
`S_OK` Copy no-op now preserves `Completed`/`NotPublished`/`Retained`, while destructive non-commit,
uncertain receipts and cleanup debt retain their stricter rules. Empty/existing/sentinel/nested/new
directory controls and the canonical completed-task summary cover this result boundary.
The first executor witness also exposed successful host-bridge no-op merges whose exact
`NotPublished` fact was overwritten as `Published`. The result store now preserves that proof
on success; executor/bridge and direct provider/callback/finalizer controls remain distinct.
Fresh gates and candidate performance must be repeated after these repairs.

The third fresh Debug Full run (`20261002T070305Z-42792-e7d0008e473245c6976c74610ce8edec`)
recorded 2,204 passed, six failed and 51 skipped; 18/20 entries were promoted. Four Commands
failures included a whitespace-sensitive source guard and three interaction assertions needing
more precise diagnostics. The two File Operations failures exposed a direct live-output Skip
being normalized as partial work, and a recursive-copy fixture that requested four workers without
establishing the provider's effective configuration. Direct `S_FALSE` Skip remains no-attempt
Skip; only exact `S_OK` with a consumed descendant Skip is normalized to partial work. The
recursive-copy receipt controls now establish and restore configured serial/parallel limits and
check storage independently of worker execution. The live-output journey checks the immutable
per-item completion, publication and source axes as well as task status and disk state.
Adjacent inspection review also found that clearing FolderView's path did not cancel enumeration:
an already queued successful result could repopulate the cleared view. Clearing now uses the
existing generation/cancellation path; the reentrant-selection control retains an actual row arena
and submits a pre-clear payload to prove rejection. Diagnostic and failed runs remain preserved;
fresh final gates are still required.

The revised recursive-copy and live-output journeys pass focused Release validation (three
passes each). The source guard, final Preferences selection/clipboard journey, and its native-focus
reliability control each pass three Debug repetitions. Reentry traces also confirm that a queued
pre-clear enumeration result is rejected in all three repetitions. Physical Space delivery and
inspection handoff remain under investigation: accepted `SendInput` is not delivery evidence,
and inspection traces record intervening input and one foreground change to Mouse Without Borders.
The assertions remain intact; these diagnostic failures are not counted as qualified passes.

The next repeated Debug group passed 12 cases and failed all three throughput-disclosure
repetitions while opening path detail late in its 40-cycle loop. Exact text, read-only/multiline
semantics, retained focus and host counts remained correct; foreground ownership changed to DWM.
A separate failing run (`20261002T105144Z-85132-36399f15733d4b88a293d3c2e360c322`)
matched its failing foreground HWND to a read-only observer's exact `Ghost` window class.
The fixture's long direct-paint sequence had not serviced its queue. It now uses the canonical
bounded pump between paint batches, outside timed scopes. Ghosting remains enabled and all
foreground, native/retained focus, storage and host-lifetime assertions remain intact.
Fresh focused and Full qualification must establish the repaired fixture's outcome.

The next Debug group ran while Release was compiling and recorded nine passes and six
failures before the repaired path-detail cycle could qualify. Its third inspection witness
showed a new cache refresh overtaking the earlier settlement (`enumerationGeneration=80`,
`completedGeneration=79`): selection correctly returned `Pending`, so no reentrant callback
ran. The fixture now forces a new refresh and waits on actual current-visit `Selected` status
immediately before installing that callback, while retaining the clear/late-payload assertions.
An isolated throughput run (`20261002T112526Z-22008-a4744fdbe7624255b0dcd9ab1a21d412`)
passed physical Space and initial Unicode detail checks, then failed at cycle ten when
Zoom took foreground ownership; exact text and host counts remained correct. Both failures
remain diagnostic history. Fresh qualification must run without build load, and real foreground
input remains subject to the governed desktop lease and visible warning.

Pending: qualification records are added here as checks complete. Build success alone does not establish runtime correctness. No ARM64 native pass or final ASan adoption qualification is inferred from x64 results.

2026-10-03 linked qualification checkpoint (receipt 59360b1ea20cfc0bb73d6baeef1cec442234f2ff4314eeb972318f2cde521014):
Debug full Rebuild passed with zero warnings/errors. Every exact linked run and each D10 owned
child whole run was archived before the next runner GC under
`C:/RedSalamander.Perf/evidence/i27-followup-20261002-0e139fb/resume-20261003T121349Z/`.
Safety's D1 actual post-child-delete cancel, D4 locked/positive-write overwrite and negative
controls, D3/D8/D9, F2/F4/F5 and all eighteen F6 telemetry states passed in the retained run.
D6 failed only its stage expectation: Native Local direct-final publication correctly records
Published. Its retained failure polluted the following idle path-detail witness; that cascade
is not an independently confirmed product failure. D2 and D11 Copy failed only verification
expectations (requested but not reached is NotApplicable); D2 ERROR_PARTIAL_COPY plus actual
WRITE_FAULT diagnostic remains required. D2 lacked sticky failure retention; its erased payloads
cannot be reconstructed as proof. Fresh repeats remain required for all corrected witnesses.
D10's three natural children completed Native Local Copy before the host-managed pause hook;
all failure receipts remain. Strict v2 rendering failed before measurement because synthetic
host client sizing assumed a standard nonclient frame. No timing pass or usable pair is claimed.

Next correction bundle (unbuilt): fix only those receipt expectations; stop Safety sub-witness
execution after sticky retention and retain D2/D10 failed runs; use a private test-only atomic
Native stage observer, with exact expected/desired registration, pinned actual provider module,
synchronous borrowed proof, and unregister only at the production worker quiet point; diagnose
terminal-before-cut immediately. Exact v2 synthetic viewport derives current outer/client delta
and retains exact client/render/cache/resource assertions. Root reviewed Sol's proposal and fixed
blind observer replacement and optional dereference before adoption. No DxUi pin change.

Frozen baseline: C:/RsI26Adopt was freshly verified at 687 native blobs, binding e860022a088f8b5fc3c604a61e091452cf83337e.
A separate ordinary checkout C:/RsI26V2 reproduces all 12,739 frozen source inputs with no raw
byte differences, independent Git objects and no copied build products. Keep original fixtures,
29% adverse v1 pair and failed source comparisons. Install identical final v2 fixture bytes in
both current and isolated baseline; retain incremental overlay provenance before Release pairing.
F7 production admission baseline must run before the proposed root/GUID scratch allocation
optimization. Final Full Debug/Release, paired Release, actual French/keyboard/UIA/mixed-DPI,
resource/gallery and governance gates remain open; no completed-plan move is yet authorized by evidence.

2026-10-03 fresh Debug linked outcomes (receipt be7c23af881fdb8843d5d3d0ed2f26b42982022f071f9904c8c9b085f3aaef6d):
Safety 20261003T152620Z-87740-15507313ab9a4d6fa97bca8c163d930d passed all three iterations,
9 passed/0 failed/0 skipped, including D1/D4/D6/D8 and idle/no-activation reconstruction.
Real partial Managed Move 20261003T152841Z-87740-5f1ed9b3278e4bb8ad38b0adf7c49ad3 passed
three iterations,9/0/0; repeated non-clipboard Copy/Move 20261003T152925Z-87740-2bf883578560485895c2e45d2eebe2b5 passed three iterations,12/0/0.
Their whole packets retain actual byte/BLAKE3/identity/tree/result-axis observations. These close
the dedicated Debug destructive witnesses; fresh final Release/Full remains the aggregate gate.
Strict v2 integrated fixture 20261003T153744Z-74812-e9563585bd1348599297d0fb28b29299 passed
1/0/0 with no shuffle; final viewport fixture SHA B0848BE7AB47D0D4FD4804B858E8A7729242A3AD8BA03468A151CAC0878377D4.
All raw packets retained under the current resume evidence folder before runner GC.

D10 20261003T153012Z-87740-6397864f0d244323b66f6cebcae9e589 failed once and retained/stopped
correctly. Source tracing supersedes the earlier Native inference: requested verification forces
the host bridge while the admitted Copy strategy remains Copy (enum0). The Native observer was
unreached. It and its new header/module/project include are removed; the actual bridge positive
write cut is restored. Fixture paths now join src/dst and filename as separate native components;
the former dst/payload.bin literal had mixed-slash spelling unlike the constructed admitted destination.
Require actual Copy strategy, one-shot Overwrite, exact persisted owned-stage prefix, normal WM_CLOSE,
responsive Stopping, natural quiet settings-save/destroy and canceled/unpublished/retained/removed axes.
New runtime is required; the failed proposal and natural child receipt remain immutable.

The independent C:/RsI26V2 Release Embedded test-enabled baseline build is currently running.
Candidate edits during it do not touch any baseline input. Pair capture must remain serial with
no competing builds/tests and fresh compatible receipts. Root repaired seven stale historical I26
support links only, preserving the original frozen Done plan and all historical failure/verdict text.
Canonical sandbox audit is non-blocking;265 historical foreign/manual run entries stay untouched,
fully reported, and cannot be relabeled globally clean by a comparator.

2026-10-03 corrected-route qualification (receipt 167eaae1890e93ee1c05785a83caa127c821499107bd9f616aa2d7ba63fdb1a7):
Full x64 Debug build passed, zero warnings/errors. D10 run 20261003T155627Z-73460-c72900633e264962a3e6ae8a4ba65f17
passed 9/0/0 across three iterations. Each real main-window child paused after 2 MiB of actual
owned-stage bridge Copy writes, accepted one Overwrite decision, serviced ten heartbeats while
Stopping, then naturally completed WM_CLOSE, worker quiet, settings save and WM_DESTROY.
Parent independently verified exact source/destination/sentinel bytes and BLAKE3 after process
exit; stage removed, final unpublished and source retained. Failed Native-proposal receipts stay failed.
F3 provider confirmation run 20261003T155837Z-4592-eee72734be8b4b74bcd38b5f79190160 passed 3/0/0.
F1 actual failure-surface attach/target/child-retry cases each passed 3/0/0 in runs
20261003T160129Z-105012-581bc564a8fd4e87b616b52b5d6053c5,
20261003T160222Z-105012-32979e23ffa34f978c73d297ec2f5318 and
20261003T160310Z-105012-eadcb4118fb24b5abf68e165286963ff. Each whole packet was retained before runner GC.
A literal wildcard case-filter trial matched no cases and remains a failed coverage invocation,
not product qualification. The evidence controller now sanitizes wildcard characters in filenames;
actual named-case reruns above are the witnesses. Static/native-message tests did not use the
user's pending final protected real-input session.
Independent I26 v2 Release Embedded full build completed, zero warnings/errors, receipt
2fb4d7aabad65f136080922bd94c2999d10e6d13f85350b39b2806632c9405c9. No pair has yet run.
Next: build current Release Embedded with the actual F7 instrumentation and retain its real admission
baseline before any root/GUID allocation optimization. Then candidate pairing, final Full gates,
protected French/keyboard/UIA/mixed-DPI/gallery and governance. F6 comparison is being reviewed
from the retained same-fixture baseline and candidate. No Done move or new performance pass claimed.

2026-10-03 F7 diagnosis and next bounded slice: the full current Release Embedded build passed
with zero warnings/errors, receipt 51a8a731e00e0760860e4802b3c119dea3b5c18c94765ee1694b991ef53ffadf.
Actual admission run 20261003T161102Z-80976-3974460c379e48348ee7e5a6ef0d98a2 failed at n=4096.
The complete C run and exact D run are retained under the resume evidence folder; no mount was
registered. Eighteen n=1/n=128 samples retained exact cancellation/storage proof, but n=4096
did not reach the preconsumption gate, so it has no passing preservation assertion.
Measured actual plan construction was 2,868,966 us and worker interlock preparation 55,278,144 us
for 8,192 scopes. The private builder searches prior scopes and authority nodes linearly on
each addition; review its scaling before retrying. Preserve the 60-second gate, selection sizes,
exact authority/binding semantics and failed raw baseline. Root/GUID output allocation policy is
still unchanged; a complete baseline with that original policy is required before its optimization.
The mount fixture's primary predicate incorrectly expected scratch, while the owning harness
supplies artifacts/selftest/last_run/fileops/work. The written correction requires the exact
acquired C case path and exact acquired D scratch case path, with unchanged marker, GUID, NTFS,
no-reparse, actual-registration and identity-checked unmount gates. Accepted UI ingress counters
are now logged before worker preparation, explicitly marked pending preparation/storage proof.

F6's offline parent-reviewed report retains all 108 observations: the candidate has zero popup
periodic work in all 27 settled idle intervals; active timer/taskbar work remains live. Coarse
Debug CPU and resource differences, differing preceding fixtures and source closures prevent
isolated resource attribution. State-level adverse signals and the failed baseline stay visible;
fresh controlled Release/resource qualification is still required. No performance gate is waived.
Spec inventory preview reports zero blocking findings, and all 193 discovered Tools files have
one classification with zero inventory findings. These previews do not replace final governance.
The v2 pair collector's native-commit assertion needs a truthful provenance correction: the
candidate's native helper ignores Git worktree commondir and emits a blank commit, while the
canonical build receipt and actual common ref bind fbe113a0. The ordinary detached baseline
emits its actual ccebd5af commit. Never inject metadata or weaken receipt/source/output binding;
record the explicit missing native field and its limitation. Final Full/protected/gallery and
Done/index/archive/automation retirement gates remain open. No competing writers/builds.

2026-10-03 indexed preparation qualification: full x64 Release Embedded solution build
passed with zero warnings/errors, receipt 6898adbc4dfae7605d22af877d32b2b0cbc6b46d39dda1d66494fc9b76c47bb5.
The Unicode/first-match/qualified-contract/reallocation/storage-cap differential passed in
Safety run 20261003T170652Z-49696-8a4f86c3b9af4efbb72cb4f9916928a1, as did all its destructive,
completion and status witnesses. F6 failed without a specific diagnostic after its first fifteen
intervals; the next idle motion settlement reported success and zero cadence. This is not yet
proof of a scheduler defect. The complete failed packet is retained; split failure diagnostics
are now written and uncompiled, with all original predicates preserved.

F7 run 20261003T171032Z-96760-77679b86c0c94a9e83220d09d2ce6700 reached and qualified its first
actual n=4096 canceled/no-write sample. Preparation measured 735,000 us versus the original
55,278,144 us failure (same selected/scopes/node sizes; sequential measurements, no variance claim).
It recorded 8,191 authority predicate evaluations and 4,393,600 bytes of temporary key capacity
for 12,301 entries. The full run subsequently failed its combined foreground/provider predicate
at prime sample20; it does not qualify the entire matrix. Both complete C and D run packets are
retained. Split exact foreground HWND/PID/thread and pane-provider diagnostics now distinguish
that check from the independent filesystem inventory. Original root/GUID scratch allocation
policy is still unchanged; no buffer optimization or complete F7 baseline pass is claimed.
The index now additionally caps distinct-contract buckets at256 and switches every lookup to
the original scan after overflow, preventing unindexed same-root contracts from disappearing.
New cap differentials and these diagnostic changes need a fresh build and linked qualification.

2026-10-03 next frozen Release bundle: F6 has exact-case, owner/thread-checked current-source
idle-suspension on/off controls and failure diagnostics, preserving the eighteen 650 ms windows,
foreground assertion, natural motion/cadence settlement and production default. A failed bounded
fixture drain now hard-retires that primary owner's workers before restoring settings, retains
scratch and stops later family/repeat reuse; it cannot retry cleanup into a passing result.
The private preparation index additionally caps contract buckets at 256, entries at 32,768 and
key capacity at 16 MiB, with full original predicate scans after any contract overflow. These
changes are written but not yet compiled. Original root/GUID output buffers are unchanged.
Next: fresh Release build, controlled F6 and complete actual F7 baseline, then bounded provider
buffer optimization, candidate comparisons and the final Full/protected/gallery/governance gates.
The user's final protected desktop agreement remains unused. No failed receipt is reclassified.

The frozen control bundle compiled successfully in Release Embedded, zero warnings/errors,
receipt 57164358a28ca86a590600c0747e227e1b3c9a13f4555c106d2af88850bf2c3a.
Before linked runs, the independent review found that the mount fixture built paths from an
overlong unvalidated environment buffer. Root verified the finding and applied bounded
construction only after run-ID validation, retaining all exact sandbox/GUID/NTFS checks.
Current source-only normative anchors and position-independent history references are reconciled;
runtime/performance/Full acceptance remains open. A fresh receipt must include these final edits.

2026-10-03 complete original-buffer admission baseline: test-enabled full-solution
x64 Release Embedded build e0d3bf3d2d2a135165aa29f4541b26e1ebbd3be98b52bd6081223cca2eb5f072
passed with zero warnings/errors. Linked Safety passed 3/0/0. Actual F7 run
20261003T182209Z-51656-126e7ed59f724fefae9d270d755c8153 passed native 3/0/0 with
252 prepared, canceled, byte/BLAKE3-preserving samples and 126 separate prime ingress
samples. All available Local, distinct C/D and loopback UNC routes are represented.
Windows denied the first actual mount-manager registration with ACCESS_DENIED;
24 original mounted route skips remain S_FALSE and do not qualify those routes.
The complete independent C/D run archives and canonical receipt/source/output
binding are retained under resume-20261003T121349Z. The earlier PowerPoint foreground
failure and its independently recovered C archive remain failed evidence.

After retaining that baseline, root applied and range-formatted the reviewed Local
root/GUID buffer change: MAX_PATH root stack output with one same-input fallback at
the previous 32,768-character bound; documented 50-character GUID stack output;
NUL/trailing-slash validation and reused Debug hook controls. Cache TTL, endpoint
identity, strategy and readiness deadline are unchanged. The change and its updated
normative contracts are UNCOMPILED/UNQUALIFIED at this checkpoint. Next: fresh full
Release build, complete candidate F7 on the unchanged matrix, retained build binding
and parent-reviewed comparison; controlled F6 on/off must use that final binary.
Then run the unchanged v2 five-pair Release/resource comparison, fresh Full Debug and
Release, protected French/keyboard/UIA/mixed-DPI/gallery checks and final governance.
No timing/resource gate is waived. The final protected desktop agreement is unused.

## October 5 bounded diagnosis checkpoint

The fresh Full Debug and Release runs remain FAILED; their closed no-follow archives and immutable receipt epochs are retained under the external I27 evidence root. Release had settings-child timeout, theme timing, R0e timing and ViewerPE restoration failures. Full Release F7 completed all 378 matrix rows, but this does not qualify the failed Full run.

Root rechecked 30 original packet bindings for a finite, unchanged Release R0e series: 1,798,508 / 1,945,047 / 2,261,573 / 2,107,117 / 2,040,769 us in five distinct processes. Three fail the unchanged 2,000,000 us ceiling; all eight metric rows and admission counts remain intact. This is variability diagnosis, not a before/after optimization qualification. The isolated settings case also failed; the isolated theme case passed without erasing its Full failure.

After that series and all owned archives closed, root applied bounded settings-child trace retention and temporary read-only ViewerPE phase diagnostics. A fresh whole-solution Release build and focused diagnosis are next. These instrumented results will not qualify the original uninstrumented runtime. Keep existing timeouts, native jobs, exact input/foreground predicates, warnings and desktop lease. The user's desktop reservation remains active. No speculative handle/GUID admission replacement was applied: reviewed link-object and volume-alias semantics require proof first. Original source backups and all external draft/controller hashes remain in `C:/RedSalamander.Perf/evidence/i27-followup-20261002-0e139fb/resume-20261003T121349Z/final-run-review/`.

No I27 closeout checkbox, deferred H4/I18/I28 gate or DxUi pin is changed. Remove temporary child phase instrumentation before final Full runs; archive final paired performance and update normative contracts before moving this same plan to Done.
