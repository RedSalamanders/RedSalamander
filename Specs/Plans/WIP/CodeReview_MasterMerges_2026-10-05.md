# I36 review remediation — remaining work

## Execution checklist (reopened 2026-10-31)

Owner: current implementation session. User authorized challenge and implementation of remaining
work. Starting identity: `6b6f5646b`, clean checkout. The staged-file directions below describe an
older checkout; no staged candidate exists at this starting identity.
Checked implementation items mean source, regressions and owning specifications are present;
qualification remains a separate requirement.

Owner delivery decision (2026-10-31): remove MSI and MSIX packaging from every
project in the RedSalamander solution; retain portable ZIP and winget delivery.
MSI/MSIX-specific repairs and acceptance gates below are historical and superseded.
Other repositories and immutable historical evidence are preserved.

- [x] Retire MSI/MSIX projects, commands, workflow paths and active contracts;
  retain portable ZIP and winget delivery and their receipt/integrity guards.

Retirement acceptance requires affected tooling checks and an actual Release ZIP
build; their exact-source results are recorded in the shared execution packet.

- [x] Repair the SearchService busy-pipe fixture's listener-recreation race within
  its existing five-second acquisition budget; add a bounded missing-listener control.

The subsequent Release PrePush passed 1,420 checks but failed that fixture once.
Its original result remains preserved. The repair changes test acquisition only;
production broker retry behavior and query/status/shutdown/cancel budgets are unchanged.
Fresh focused/Compare qualification and final delivery results belong to the execution packet.

- [x] Preserve the parallel search coordinator's final-batch callback failure after
  worker quiescence; add deterministic success/abort/error/cancel coverage.

The later affected repeat exposed a separate product race in `local_search_callback_contract`:
it recorded `E_ABORT` after observing completed workers, then returned stale success.
The candidate reconciles the first terminal status under the existing post-join lock.
The root-scoped test-only join gate and perf row prove the last-batch ordering;
fresh qualification remains separate from the preceding failed run.

The current build captures native artifacts and module provenance after native
compilation, then validates uniform test-surface metadata and final native bytes
before publishing a receipt. ZIP staging must preserve that captured closure.
The MSI/MSIX source was preserved before removal, including diagnostic failures
that exposed excluded evidence during MSI staging and seven stale native hashes
after MSIX project-reference compilation. Those historical installer repairs no
longer define shipping requirements. The later Debug PrePush run was interrupted
when the owner changed delivery scope; it is not a completed pass. Qualification
results and remaining gates belong to the shared execution packet. Focus-taking
work remains pending by request.

- [x] Reconcile every residual wave and coverage gap with current code, delivered history,
  the follow-up record and overlapping I50 findings; remove stale execution directions.
- [x] Independently verify consequential reviewer conclusions and record rejected hypotheses.
- [x] Implement confirmed remaining defects with meaningful regressions and owning specs.
- [x] Review simplification proposals for current need and repair risk before implementation.
- [x] Run noninteractive affected explanation, appropriate builds/tests and required
  performance scenarios; retain exact-source results and original failures.
- [x] Complete fresh noninteractive PrePush coverage accounting and record deferred gates.
- [x] Validate both tests-disabled Release ZIPs, their receipt/integrity closure,
  and stock Windows checkout compatibility.
- [ ] Run foreground Commands, FileOperations, ViewerPE and ViewerSqlite acceptance
  under the desktop lease at an agreed time; deferred by the owner.
- [ ] Execute native ARM64 qualification and classify the ARM64-only WIC decode
  discrepancy; explicitly pending by owner direction, not closed by a cross-build.
- [ ] Complete approved physical-device and provider/environment checks, including
  MTP, SMB and elevated/NTFS/live-journal prerequisites; explicitly pending.
- [ ] Complete mixed-DPI, real assistive-technology and real shutdown/logoff acceptance.
- [ ] Move to Done and update index/exact-file admission only after all required gates close.

## Current acceptance checkpoint (2026-11-01)

Owner authorized commit and PR publication with native ARM64 and physical-device
checks left pending. The earlier decision to defer foreground tests also remains
in force. Both plans stay ACTIVE in WIP until the unchecked acceptance gates close.
Earlier checkpoints below are historical; this checkpoint and the shared execution
packet own the current result.

At qualified implementation snapshot
`eadb3ec0caa3f840950eae4eb9c60f6b993f2b5af28624f6a8f67b102292c56c`:

- Fresh affected and PrePush execution each passed **1,421 / 0 failed / 31 accepted
  capability skips**. All fourteen selected entries were executed and promoted;
  none were incomplete, invalidated, provisional or reused. Source hashes before
  and after each run match. The change verdict is PASSED; the repository verdict
  remains NOT_EVALUATED because four foreground entries were deferred.
- Complete Compare passed **252/0/31**, including both repaired search cases.
  The forced final-batch regression proves four dispatches with zero remaining
  workers and one queued result. PerformanceTests2 passed **15/15**, the compiler
  lane **30/30**, and tooling **1,128/1,128** in both runs.
- Native builds reported **zero warnings and zero errors**. Tests-disabled x64
  and ARM64 production ZIPs each have 174 byte-verified receipt outputs, 175 ZIP
  entries, no excluded build/test sidecars and no hash mismatches. The x64
  fresh-extraction native package smoke passed; no ARM64 executable was run.
- Stock Windows checkout proof `sf-daf065ef` passed with explicit long-path and
  symlink support disabled: 28,032 regular files, maximum relative path 200 and
  absolute path 251, no round-trip mismatches. All 3,002 historical relocation
  blobs and 6,004 text/eol attribute comparisons match. Source index, status,
  HEAD, shared indexes and object-store content remained unchanged during proof.
- Final source, shipping-byte, archive, changelog, tool-inventory and whitespace
  checks passed. The final archive guard checked 1,298 paths with zero violations.
  Original failures, interrupted attempts and earlier snapshots remain preserved;
  a passing repeat does not change their verdicts.

The [execution packet](../../TestRuns/Local-x64/MergeReview/2026-10-31_Remediation/README.md)
contains the exact run IDs, receipts, source identities and remaining gates. This
plan/status closeout follows the qualified snapshot. Integration with current main
also preserves its newer test and CI policy changes and reuses its canonical pipe
fixture connector with deterministic attempt observations. Merged-source validation
is recorded separately; old receipts are not retagged as that integration or as
native ARM64, physical-device, foreground, mixed-DPI, assistive-technology or real
shutdown acceptance. I50 retains the incoming revision-2 review/routing scope.

Qualification checkpoint before the final fixture edits: noninteractive PrePush
`20261031T164304Z-21260-b7c1915171574a9abd900578f6a31e5b` failed with 1,433 passed,
two failed result groups and 29 skips; twelve of fourteen entries were promoted and none
invalidated. The compiled PluginContract, Compare (256/0/27 accepted capability skips), Curl
and PerformanceTests2 (15/15) entries passed. The build-toolchain lane passed 30/30 and
re-attested x64 Debug without diagnostics. Test-enabled x64 Release and all three ARM64
cross-build profiles also have zero diagnostics at the same source; ARM64 execution is unverified.

The candidate corrects the two ViewerSpace cache fixtures' build-versus-restore wait, using
real provider-read and post-Close empty-state witnesses. It also checks immutable pinned
vcpkg baseline data without altering the dirty managed checkout. The preceding tooling
assertions were 1,137/0/2, but the native runner correctly rejected those two prerequisite skips.
Pinned registry-data checks do not qualify the live tool/bootstrap state. Fresh results for
these later test edits and all remaining gates belong to the
[shared execution packet](../../TestRuns/Local-x64/MergeReview/2026-10-31_Remediation/README.md).
Focus-taking tests remain pending by request, so this plan remains WIP.

Earlier integrated follow-up: a tagged relocated Debug fixture warning and an ordinal path-key
collision for supplementary names were confirmed while qualifying I50/I36 together. The warning
and key algorithm are repaired locally, with required-key availability and classifier controls;
fresh native execution remains required. The interrupted PrePush attempt and original failing
NLS oracle remain preserved in the shared execution packet. Focus-taking tests remain pending.
Later build-only receipts `3c106a54` (x64 Debug) and `72d38756` (ARM64 Release cross-build)
have zero diagnostics. The deployment-test capture-path correction is later source and still
requires fresh qualification; neither an interrupted runner nor a cross-build is native acceptance.

Status: ACTIVE. Renumbered from I30 to I36 on 2026-10-11: the archived delivered record claims I30 in its title and plan ranks are never reused (`Specs/README.md`, "Plan lifecycle"); the content below is unchanged and its ledger references still name the I30 delivery. Reconciled 2026-10-10 at `ab5f40b57`. Completed review and committed implementation are archived in [Done](../Done/CodeReview_MasterMerges_Delivered_2026-10-10.md). Delivery does not imply completed runtime, performance or release qualification.

## ViewerSpace nonactivating desktop witness follow-up (2026-10-16)

Owner: `ej/viewerspace-desktop-preservation`, based on merged `5141305d9`.
The user authorized investigation, fixes, regressions, specs/docs and strictly
nonactivating verification. This slice does not close the unrelated I36 work.

- [x] Trace the failed foreground/cursor comparison and audit the UIA path for
  physical input writes; retain the failed PR run and the two passing diagnostics.
- [x] Reproduce ambient cursor interference deterministically through read-only
  observation injection, without moving the real pointer or activating a window.
- [x] Replace desktop-global equality assumptions with guarded process-scoped
  foreground evidence and explicit ambient-state diagnostics. Observation or guard
  failures and owned foreground transitions must still fail closed.
- [x] Cover ambient input, transient owned activation, unavailable observations,
  missing guards and observation lifecycle; forbid physical input in this UIA path.
- [x] Update shared-helper, ViewerSpace, testing and tooling contracts/docs.
- [x] Repair the lifecycle capture readiness race exposed during broad validation:
  hold the provider gate, wait for `Scanning` and no queued progress, and retain
  unchanged before/after stability assertions plus unsettled-state negative controls.
- [x] Build and run focused/noninteractive validation, archive exact inputs and
  failed/candidate evidence, and record remaining qualification separately.

Protected scenario: ViewerSpace UIA grid/filter/resize/collapse/lifetime coverage
while the person continues using the desktop. Retain scenario duration and observer
counts; no focus-taking, cursor restoration, input injection or timeout relaxation.

Bounded result: x64 Debug ViewerPETests builds with zero warnings/errors; all 29
fresh-process noninteractive cases and all 204 harness source contracts pass. The
original failed PR, expected fail-first comparison and first candidate's refresh
capture failure remain alongside the corrected run in the
[DesktopObservation archive](../../TestRuns/Local-x64/Viewers/2026-10-16_125146_DesktopObservation/README.md).
Production ViewerSpace bytes are unchanged. The original desktop actor remains
unknown; this fixes attribution and capture readiness, without claiming a new
interactive, other-platform or whole-repository Full verdict.

## Authorized supplied-review remediation (2026-10-16)

Owner: this branch, `ej/review-safety-fixes`, based on `cca51ab09`. This bounded
slice does not close the other I36 work below. The user authorized fixes, regression
tests, specification/documentation updates and verification for review items 1, 2,
6, 7, 8 and 10; item 4 is a separate shutdown investigation. Items 3, 5 and 9 remain
unsubstantiated as originally claimed and are outside the patch.

Implementation checklist (checked means code/fixtures delivered; runtime qualification is
recorded separately in the evidence ledger below):

- [x] Add the shared checked directory-record decoder, migrate compatible Local walks,
  and add malformed/valid layout regressions. The search adapter also checks
  a name crossing into a valid following record; its source guard enforces shared decoding.
- [x] Retain no-follow parent authority through child joins and exact final disposition;
  add serial/parallel, bound/unbound, 32-level descendant and cancellation fixtures.
- [x] Bind PR delegation to workflow/action/tooling execution closure and test committed
  reusable-workflow changes, imported dependencies and failed discovery.
- [x] Validate actual benchmark destination names/content on both routes outside timing;
  add missing/corrupt-output negative controls.
- [x] Correct Dummy directory-record strides exposed by the real benchmark verification;
  cover mixed filename lengths and canonical streamed/indexed traversal parity.
- [x] Return typed XP-09 skipped/failed prerequisite results and exercise both branches.
- [x] Use random destination/message/data-type-bound payload tokens, migrate receivers,
  and add spoofing, ICF-equivalent type, entropy-failure and teardown regressions.
- [x] Investigate the scheduler separately with source/lifetime tracing and a bounded
  nested-wait probe; retain the unresolved valid app-exit lifecycle explicitly.
- [x] Update owning specs, helper guidance, tooling inventory and native test inventory.

Verification must complete before declaring this bounded slice finished. Its exact commands,
failed attempts, build/source identities, case outcomes and performance observations are
maintained in `Specs/TestRuns/Local-x64/ReviewSafety/2026-10-16_Remediation/README.md`.
A focused pass does not close Full, interactive or other-platform qualification. The user
requires nonactivating validation; interactive gates remain pending. The larger I36 plan
below remains active for its unrelated residual work.

Protected scenarios: bounded Local recursive enumeration/delete, no-follow containment,
posted-payload adoption/coalescing under callback load, actual small-file Copy on both
routes, and honest PR/case coverage. Retain failed receipts and exact source/profile
identities; a build or scoped pass does not certify Full or unavailable SMB/platform gates.

### Separate scheduler-shutdown investigation

Source-verified at `cca51ab09` plus this patch; production scheduler behavior is unchanged.
`SharedFileOpsJobScheduler::WaitJob` in `FileSystem.FileOps.cpp` cooperatively executes
work on a worker thread and otherwise waits without a stop predicate. `ShutdownAndJoin`
requests worker stop, joins, then marks remaining jobs finished. A nested waiter for an
unfinished parked dynamic job with no remaining producer has a wait/join cycle. This is
a scheduler-level failure condition, not yet a reproduced app-exit lifecycle.

The production shutdown entry is the final `FileSystem` destructor in `FileSystem.Watch.cpp`;
its documented precondition stops producers first. Host tasks retain provider COM references
while executing. `LocalBoundObject` does not retain a `FileSystem` instance, so that boundary
needs special care; its current recursive delete/manifest path uses finite jobs, which a
nested helper can still drain. The dynamic Copy producer closes or fails its queue, supplies
a final ready unit on early stop, and notifies before `WaitJob`. No valid host path was found
that releases the last provider with that dynamic producer abandoned.

The existing `Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker` covers an external
waiter and callback quiescence, not a worker nested on a parked child. Merely returning on
`request_stop` is unsafe: job callbacks capture operation contexts and buffers on caller
stacks. A scheduler change requires an admission/closing state, cancellation/abandonment of
pending jobs and a proof that every in-flight callback finishes before any waiter returns.
A bounded native probe using the exact scheduler source prefix held shutdown for 500 ms
with a synthetic nested child at ready-count zero, then released/notified that child and
observed parent and shutdown quiescence. This demonstrates the scheduler condition without
stranding a worker; it intentionally does not establish a valid app-exit lifecycle.
Provider/bound-object lifetime stress remains an investigation gate. This patch claims no
shutdown fix or app-exit proof.

## Immediate execution sequence

1. Reconciled 2026-10-31: the previously staged popup/path-detail/localization candidate is historical; this checkout started clean. Re-read current code and delivered evidence instead of reconstructing that old staged state.
2. Waves 0-3 now have implementation commits, including Escape routing, retarget feedback, admission failure reporting, popup recovery, stable controls and idle cadence. Run their outstanding regression/performance checks and update each finding with exact commit and evidence identity. Do not repeat delivered implementation.
3. Wave 4's own status note states all code work landed. Treat its repeated bullets as historical implementation scope and reconcile missing evidence; do not replace the deliberately adopted per-window hook with the old proposed SetWindowSubclass design.
4. Wave 5 refresh/cancel fixes landed in 6442a5b91. Check production inspection cancellation callers, name-and-kind matching and test-only listener removal.
5. Wave 6 localized projection, Apply-to-all aggregation and Issues ordering landed in bc7728fd5, ff888db98, 45779b0df and ab5f40b57. Finish residual UIA/localization/selection details below and acknowledge Q10's two-row presentation.
6. Q14, Q18, Q2 and Q17 have implementation commits. Keep their remaining owner acknowledgments and verification obligations open in remediation-followups.md.
7. Implement the confirmed residuals in Waves 7-10 with surrounding-callers review, meaningful coverage and authoritative spec updates. Review the simplification track and twelve coverage gaps, retaining optional architecture leads and missing-prerequisite gates explicitly rather than treating every suggestion as a required rewrite.
8. Run the authorized noninteractive checks and available build/Release/package acceptance, then complete PrePush accounting. The user explicitly kept focus-taking tests pending; elevated VHDX, native-platform and device gates retain their own prerequisites. No affected-only pass establishes production readiness.

### Current revalidation ledger (2026-10-31)

- `runtime-queue#3` overlaps I50 M2. One candidate synchronization repair and its boundary tests own
  both entries; they must not become independent competing implementations.
- `state-c#3` is already repaired at starting HEAD: Batch Rename uses
  `kMaxCachedModifierAttemptsPerBucket = 1`, per-bucket attempt accounting and a forced fresh prompt
  after its cached attempt. The earlier still-open verdict was rejected after re-reading HEAD.
- `tests-fileops#2` is confirmed and repaired in the candidate: pure timing/telemetry rejection
  fails the case without retaining scratch or suppressing later families. Failed worker quiet
  points, byte integrity or settings restoration still retain/prohibit reuse. A real fixture
  negative control injects a telemetry-only failure after exact bytes/cleanup and then runs the
  following owned fixture; execution remains pending.
- `fs-core#1` cold/warm followed-root identity differs in the current route cache. `fs-core#6`
  retains per-instance characteristics indefinitely. Candidate resolves the lexical mount root and
  fresh direct volume identity for every route/characteristics query and reuses only matching identity.
  Unstable/unknown identities bypass persistent storage reuse. Cold/warm/nested/remount/SUBST and
  failure controls are added; execution and actual query-cost evidence remain pending.
- Wave 4 source work is delivered; the Wave 5-10 residuals are reconciled in the dispositions
  below. The remaining inline-F2 queue-slot lead is not a confirmed defect without a contract
  requiring Wait for others to exclude inline rename. G1-G12 are coverage obligations, not twelve
  proven defects. Current dispositions retain G5 latency, G6 logoff/shutdown, G9 physical input,
  G10 tests-off Release, G11 measurement, G3/G4 consumer acceptance and G8 provider-result execution
  as explicit qualification work; source presence alone closes none of those gates.
- The two-row Apply-to-all presentation keeps the explicit user choice separate from its cached
  aggregate to preserve provenance. Do not merge them without evidence that the model stays unambiguous.
- Follow-up claims are snapshot-specific, including the old Pester assertion investigation and
  Recycle Bin incident. Never enumerate or empty the user's Recycle Bin to reconcile that prose.
- User directed focus-taking tests to remain pending in this session. Elevated/device/platform and
  release qualification retain their owning prerequisites; no affected pass closes them.
- Wave 7's write-only 64-entry consent receipt table is removed, without weakening typed consent
  scope/identity checks; the 65th-answer regression is added. Inline F2 now marks possible mutation
  immediately before its actual provider rename. Execution pending.
- Wave 9 cache policy is superseded by the owner: Close forgets all private ViewerSpace scan/cache/
  workspace state. Only the same open window may reuse its cache. Embedded locale refresh and
  tooltip progress repaint now use the existing locale path and 30-Hz timer; runtime evidence pending.
- Wave 10's active positive `Should Not Throw` assertions are replaced with captured success/error
  checks and a shared deliberate-error negative control. An AST guard covers positional and
  parameter spellings without treating comments or explanatory strings as executable assertions;
  focused execution remains required. Historical packets stay immutable.
- G1's regex authorization/unready-store behavior and G2's alias-retention code are delivered at
  HEAD; the relevant current native cases still need execution. G5's followed-path observation is
  removed by the route repair, but lexical mount-manager latency still needs measurement. G7's
  escaping/export fault coverage is extended with the L22 repair. G8's reducer remains conservative
  when destructive proof is missing; never infer source deletion authority from HRESULT success.
- G3/G4 dependency and DxUi acceptance attach to the current immutable pins and consumer tests,
  not a replay of the obsolete upgrade range. G6 real end-session durability, G9 physical-input
  coverage and G10 tests-off Release/native acceptance remain explicit gates. G11 is a performance
  hypothesis until mutex/stage measurements show material cost. G12's current owner is
  `RedSalamander/FileOperationsFooterLayout.h::Arrange`, which delegates to
  `DxUi::ArrangeMeasuredActions`; the old out-of-range line number is rejected. Footer fit and
  localized-width behavior remain owned by the existing popup layout/scenario tests.
- The simplification track is advisory architecture work. Shared-helper reuse and dead-state removal
  coupled to confirmed repairs proceed here; broad file decomposition, a second ConflictBroker or
  speculative model rewrites are not prerequisites merely because the old review suggested them.

The following additional residuals were checked against the current source and repaired in the
local candidate. These entries describe implementation and test intent, not executed acceptance:

| Residual | Current disposition |
|---|---|
| `common#1` | Preferences merges only fields edited from its baseline into the latest settings, preserving popup/Issues and JSON-only changes. A concurrent-field regression extends the existing settings case. |
| `host#1`, `popup-c#1` | Explicit Show no longer rejects its own command key-up as unrelated input; foreground/focus ownership checks remain. Saved popup size is normalized against the created popup's actual DPI. Native focus and mixed-monitor acceptance remain pending. |
| `popup-c#6`, `popup-c#7`, `x-ux#5` | Path detail names the source/destination subject, retains logical Copy/Close focus across synchronization, and raises localized success/failure feedback through the existing UIA notification helper. The existing path-detail fixture uses real provider Invoke with a clipboard-result fault seam. Physical focus/assistive-technology acceptance remains separate. |
| `l10n#3`, `l10n#5`, `l10n#7`-`l10n#10`, `l10n#12`, `host#3`, `popup-b#6` | Counts use shared per-culture cardinal forms; result summaries retain typed formatting arguments and refresh on locale revision. Issues terminology, separators, unavailable speed/remaining text and progress nouns agree across code/resources/docs. A localization matrix and existing presentation/language-switch cases cover these contracts. Broader raw-diagnostic claims require typed-category evidence, rather than a blanket rewrite. |
| Issues export / G7 | Log scanning occurs outside the operation-state mutex; exact writes publish through `LocalFileTransaction` with FailIfExists. Collision/write/flush/publish controls retain the existing target and avoid announcing success before Commit. The shared transaction now retains exact handle authority for both Ex and legacy rename and abort, with no close-then-path fallback. |
| `tests-other#3`, `tests-other#4`, `tests-other#6` | Restart children receive child-only environment and run off the UI thread while owned messages pump; child completion precedes reading isolated settings. Clipboard format preflight emits explicit unsupported-prerequisite skips before mutation, with malformed/read failures still failures. UIA worker/deadline budgets use the existing timeout scale. Commands child execution retains its normal desktop admission. |
| `tests-other#5` | Stopping a suite after an unsafe real-input desktop boundary remains deliberate containment. Nonactivation evidence cannot authorize another activation case. |
| `tests-fileops#3`, `tests-fileops#5` | A redundant digest after an exact byte comparison is evidence overhead, not a correctness defect. Hand-built Task fixtures are model coverage and do not claim production admission coverage. Global Space injection in the throughput helper is replaced by owned popup message routing; its activation classification is retained. |
| `fs-core#1`, `fs-core#6` debug fixtures | Synthetic storage probes and route caches are scoped to the executing test thread and a private cache. A second thread verifies that hooks/cached identities do not escape into another thread or mutate production cache state. |
| `x-perf#6`, G11 | Full row reformatting on projectable revisions and localization-lock cost are performance leads. Nonprojectable diagnostics already avoid projection revisions; no incremental model/cache redesign is justified without measured cost. |
| `build-tooling#4`, `build-tooling#5`, `x-spec#7` | Checkout-path repair uses exact immutable relocation identities. Old dependency/cache folders need their own lease/recovery inventory before cleanup; this work does not delete unrelated output trees. Historical evidence bytes and provenance remain preserved. |
| `host#5` | Source/destination text selects shared display path rules across provider roots without weakening root-qualified strategy/admission equality. The caller and its common-parent/mismatched-rule regression use the same selection helper. |
| Completion cache context / I50 L18 | A pane switch can unregister a task's retained provider pointer, causing metadata fallback to lose its nondefault instance context. Completion now supplies the admitted source/destination logical contexts through canonical cache notification overloads. A real cache A-task/A-keeper/B regression includes the legacy-route stale-entry control; this API fixture is not an end-to-end completion-callback claim. |
| `x-ux#8` | Routine/informational reveals preserve an existing popup's minimized state and z-order. Actionable, interrupted-notice and explicit Show retain deliberate reveal behavior. The fake-clock deadline fixture uses the canonical guarded desktop observer and an owned nonactivating minimize/restore. |
| `l10n#6` | The old Failed Items ID is already removed at HEAD. The remaining unused Discovering-as-needed ID/string is removed from the header and main/four satellite resources; no new orphan-resource framework is needed. |
| `viewerspace#4` | Obsolete: the current header has no Up affordance; double-click reaches only Cancel/Problems. Navigation is hosted separately. The old two-level Up counterexample has no current caller. |
| `viewerspace#5` | Capped/model-less read errors propagate to the visible owning Other bucket, including cached snapshots, tooltip/analysis Error state and high-contrast outline. The root/unrelated retained sibling keep their own truthful state. Synthetic provider coverage checks exact totals, one unreadable folder and the visible marker; high-contrast pixel acceptance remains separate. |
| `viewers-dxui#1`, `viewers-dxui#2`, `viewers-dxui#6` | The cited ColorTextView source belongs to the independent RedMonitor owner repository and has no active source/build target here. Preserve the upstream lead; do not recreate a bundled viewer or change its pin without consumer qualification. |
| `viewers-dxui#3`, `viewers-dxui#5` | ViewerText treats pre-attach/embedded menu absence as normal, preserves a native menu during initialization, and warns on a real post-attach fallback. The impossible embedded attach branch is removed. Existing noactivation open fixtures check embedded/standalone attachment state; shared menu-host failure coverage remains authoritative. |
| `viewers-dxui#4` | ImgRaw's custom attach branch deliberately supports early/post-detach fault injection. The shared adapter already owns normal production callers; removing the test seam is optional, not a demonstrated repair. |
| I50 M23 | A shared classifier oracle now runs through the existing exact noactivation remediation case with 256 real Local source files and production retained-anchor preparation. Fixture preparation is outside classifier timing. Disjoint and last-pair overlap outcomes are asserted; comparisons are measured, not required to remain Cartesian. The broader ProviderCapabilityMatrix desktop admission is unchanged. |

Required focus/device/platform/performance gates above stay open even when a source repair is
complete. Current commands, preserved failures, exact receipts and measurements are recorded in
the [shared execution archive](../../TestRuns/Local-x64/MergeReview/2026-10-31_Remediation/README.md).

Current implementation and verification evidence is shared with the
[I50 execution record](CodeReview_MainMerges_2026-10-28.md). Neither plan is Done while its remaining
review, fixes or required qualification are unfinished.

The latest shared checkpoint built x64 Debug with zero warnings/errors (receipt `561b2f04`),
passed focused FileOps 3/0/0 and the exact cache API case 1/0/0, and passed focused tooling
52/52 plus 315/315. Ten harness FileOps images were visually reviewed. These are focused results
for the preceding snapshot; final breadcrumb compatibility/counter repairs and the Release paired
classifier fixture require a new receipt and execution. The separate
[delivery review](DeliveryReview_Remediation_2026-10-31.md) owns its W1/W2/W3 outcomes; there is
one integrated implementation, with exact results retained in the shared archive above.

## Detailed residual checklist

The original detailed wave definitions below retain full scope. For items already delivered above, only missing validation and follow-ups remain. Archive an item only with its fixing commit and evidence; retain unknown coverage explicitly.
### Wave 4: alert overlay lifetime and accessibility (M)

Status note (2026-10-12, from the I35 S5 research brief, spot-checked in code): every Wave 4 item below is
implemented at `3044781f0` and only its evidence archival remains. `alert#2` (a) is the `WM_NCDESTROY` reset
(`Ui/AlertOverlayWindow.cpp`, commit `69caa9dbb`); `alert#2` (b) landed as one owned `GWLP_WNDPROC` link per tracked
window with per-overlay listeners rather than `SetWindowSubclass` (deliberate, `Specs/UI/UI_DxUiSharedGrid.md`);
`alert#3` (`5e581d8c0`), `alert#8` (localized UIA names), `alert#9`/`alert#10` (`SetTrackedWindowDestroyedCallback`, no
`onDismissed`), `alert#1`/`alert#4`/`alert#6`/`alert#7`/`x-ux#12`, and `alert#5` (registered
`RedSalamander.AlertOverlay.UiaCall` message, `WM_APP + 0x76` retired, `ff2773d2b`) are **closed** pending their
evidence lines. I35 S5 (the decision mirror) builds on these and is not blocked by Wave 4.

- `alert#2`:
  - Set `_visible=false`, `_blocksInput=false` and clear `_restoreFocus` in the overlay's `WM_NCDESTROY`, so
    nested prompt loops always exit.
  - Replace the shared per-HWND property hook with `SetWindowSubclass(..., (UINT_PTR)this, this)`.
- `alert#1`: no content jump on option click or arrow cycling in scrollable alerts.
- `alert#3`: raise UIA focus and structure events on in-alert focus moves and model replacement.
- Low items:
  - `alert#4`, `alert#6`, `alert#7` (mnemonics: Q in ledger).
  - `alert#8` (hard-coded English UIA names).
  - `alert#9` (window-scope overlays leak).
  - `alert#5` **closed** (I35 S5 closeout, 2026-10-12): the `WM_APP+0x76` handler trusted `LPARAM` from any
    sender; fixed in `ff2773d2b` with the registered `RedSalamander.AlertOverlay.UiaCall` message and an opaque
    payload token (`WM_APP + 0x76` retired, `Specs/UI/UI_DxUiSharedGrid.md`). Evidence:
    `FileOps_AlertOverlayUxAccessibility` sends a raw call pointer on both ids and requires that nothing runs;
    passed in `FileOpsFamily_AlertOverlayLifetime` on branch `ej/i35-s5-decision-mirror` (run
    `20261012T015250Z-17460-4e15d4b3a4394701820349d6945a6fb2`, 4/4).
  - `alert#10`: dead `onDismissed` path with a latent use-after-free.
  - `x-ux#12`.

### Wave 5: conflict inspection (S-M)

- `folderview#2`: a cancelled enumeration must settle, so the inspection stops showing "Opening the folder…"
  and stops blocking later inspections.
- `folderview#1`: a refresh during the wait must not turn the inspection Stale.
- `x-lifetime#1`: give `CancelPendingConflictInspection` production callers (prompt resolved, task ended, popup
  destroyed).
- Low `folderview#4`: match by name and kind.
- Info `folderview#3`: move test-only listener slots behind ENABLE_TESTS.

### Wave 6: Issues pane, status text and localization (M)

- `x-arch#1`: localized typed problem text in the primary column; English engine diagnostics only in Technical
  details (Q9).
- `model#2`: Info decision/result observations must not evict real Warning/Error issues from the bounded
  reservoirs. Aggregate cached apply-to-all per prompt generation (Q10). I35 S1 (2026-10-12) extends the rule:
  first-attempt questions and expected outcomes are Information notes (original rows), evicted before any
  Warning/Error problem; with `fileOperations.showInformationalNotes` off the Issues pane shows only Warning/Error rows,
  their joined decision/result rows and the cached apply-to-all row of a visible originating prompt; original
  Information rows use the `Information` caption with the information glyph.
- `issues#1`: detail follows the grid's current row.
- `x-ux#1`: UIA notifications for Canceled, Indeterminate and fail-before-first-card.
- Low items:
  - Issues pane: `model#5`, `issues#4`-`issues#9`, `x-perf#6`, `x-ux#6`, `x-ux#7`.
  - Localization: `l10n#3`-`l10n#10`, `l10n#12`, `host#3`, `popup-b#6`.
  - Path detail and accessibility: `popup-a#2`, `popup-b#4`, `popup-c#5`, `popup-c#6`, `popup-c#7`,
    `x-ux#5`, `x-ux#8`.
  - Confirmation text: `host#5`.

### Wave 7: queue and runtime edge cases (S)

- `runtime-queue#3`: notify `_queueCv` under its predicate mutex.
- `state-a#2`: recycle-escalation receipts must not refuse the 65th answer.
- `state-c#3`: cap cached modifier retries in Batch Rename, as Copy/Move/Delete do.
- Low `runtime-queue#4` (inline F2 occupies the Queue slot) and `x-lifetime#4` (deferred exit; Q14).

### Wave 8: FileSystem route identity (S-M)

- `fs-core#1`: the route `rootId` must not depend on what was queried in the last 2 s.
- `fs-core#2`: route-cache hits now open the path and query the mount manager twice.
- Low `fs-core#6`: per-instance characteristics cache never expires.
- Coverage gap G5: the route observation follows reparse points with `CreateFileW` on UI-thread callers. Bound
  its latency for offline UNC and removable targets.

### Wave 9: ViewerSpace, viewers, Monitor (S)

- `viewerspace#1`: add an absolute maximum age to the sliding TTL, and decide whether to cache scans that
  contain failed folders.
- `viewerspace#2`: the locale refresh must reach the embedded preview pane.
- `viewerspace#3`: the tooltip must respect the 30 Hz throttle.
- Low `viewerspace#4`-`viewerspace#6`.
- `viewers-dxui#1`, `viewers-dxui#2`, `viewers-dxui#6`: ColorTextView swap-chain reuse, null factory, use
  `DxUi::IsDeviceLossHResult`.
- Low `viewers-dxui#3`, `popup-c#8`.

### Wave 10: tests, tooling and spec hygiene (M)

- `tests-fileops#2`: timing failures must not abort every remaining Full family.
- `tests-other#3`, `tests-other#4`: environment mutation and UI-thread blocking; clipboard formats.
- `popup-c#1`: saved size rescaled with the target monitor's DPI.
- `host#1`: the explicit Show command focuses the popup.
- Spec hygiene:
  - `x-spec#1`: duplicate I27. The working-tree index now also has two I28 rows. **Done (Q18, 2026-10-09):**
    the DxUi plan is I31 and Scoped testing is I32; the Local route-observation obligation is I29 Q2-L; the
    spec-index check now rejects shared ranks and ranks claimed by archived plan titles.
  - `x-spec#2`, `x-spec#3`: merged with open qualification or perf evidence.
  - Low `x-spec#4`-`x-spec#7`: user and dev docs still describe removed pre-calculation; 225 MiB of captures
    under Specs/Plans.
- **New, found during Wave 0 (`impl-w0-1`, agent-reported, not yet re-verified by hand):** under PowerShell 7 the
  repository's Pester 3.4 `{ ... } | Should Not Throw` passes even when the block throws (a minimal
  `{ throw 'boom' }` repro passed). Existing positive-path Pester cases that rely on it assert nothing; for example
  `VcpkgInstallSafety` "accepts manifest version floors present in the selected registry". Audit every
  `Tools/Tests/*.Tests.ps1` use, replace with captured-exception assertions or move to a Pester version whose
  assertion works, and add a contract that rejects the pattern. Medium (silent false passes in governance gates).
- Remaining Low: `tests-fileops#3`-`tests-fileops#5`, `tests-other#5`-`tests-other#7`, `build-tooling#4`,
  `build-tooling#5`, `fs-core#4`, `common#1` (Preferences Apply clobbers popup-owned settings).

### Simplification track (Info, schedule after Waves 0-3)

1. **Test-only code out of shipping translation units** (`x-arch#5`, `state-c#7`, `state-b#7`, `fs-core#5`,
   `viewers-dxui#4`, `tests-other#9`). About 6,200 ENABLE_TESTS lines are interleaved in the FileOps TUs. Move
   them to `*.SelfTestHooks.cpp` and unify the fault-injection plumbing.
2. **One conflict-prompt representation** (`x-arch#7`, `x-arch#8`, `state-a#3`, `state-b#6`). Finish the
   ConflictBroker extraction: one state, one wake mechanism, one consent-fact copy, shared overlap-consent code.
3. **One result-classification seam** (`x-arch#9`, `model#7`, `state-c#6`, `x-arch#4`, `fs-fileops#6`,
   `x-arch#6`, `host#7`):
   - Replace the 17-argument positional lambda with a struct input to the pure reducer.
   - Share receipt validation and the source-index-to-plan-item mapping.
   - Stop re-deriving terminal status from collapsed HRESULTs.
4. **Popup decomposition** (`popup-a#5`, `popup-a#4`, `popup-c#10`, `popup-c#11`, `x-arch#12`, `popup-c#12`).
   Split prompt, pure status model and debug structs out of the 15.4k-line TU. Remove dead footer hit kinds,
   merge the five anchored-menu routines, use one DIP measurement system, and add RTL to the footer and
   completed-group layout.
5. **Dead code** (`fs-fileops#5`, `viewers-dxui#5`, `viewerspace#7`, `common#3`, `issues#11`, `l10n#6`,
   `alert#11`, `x-arch#10`, `fs-core#3`, `viewerspace#8`).
6. **Queue duplication and perf** (`runtime-queue#7`, `runtime-queue#8`, `state-c#5`, `x-perf#8`).
7. **Test and tooling hygiene** (`tests-fileops#8`-`tests-fileops#10`, `tests-other#8`, `tests-other#10`,
   `build-tooling#6`, `host#8` BOMs, `common#2` and `x-lifetime#6` window-message IDs).

## Questions for the owner

Each question has a recommended default; implementation of the affected item waits for the answer. The ledger
keeps all 148 reviewer questions, with the reviewer that raised each one.

### Owner decisions (2026-10-06)

| Question | Decision | Effect |
|---|---|---|
| 1-7, 9-11, 13, 15, 17, 19, 20 | Recommended default accepted | The waves implement the default stated under each question. |
| 16 | **A test-enabled (ENABLE_TESTS) Release build must never be distributed.** | Release packaging and the release workflow must build with `RSBuildEnableTests=false` and fail if a shipped binary exports a test entry point (for example `RedSalamanderFileSystemTest*`). Simplification 1 (moving test hooks out of the shipping translation units) is cleared. |
| 8 | Recommended default accepted after explanation | "Clear completed" never acknowledges interrupted-Move notices (they keep their own Dismiss) and clears only the task IDs that were visible when the user clicked. |
| 12 | **Mounting a real user volume in a test is forbidden.** Use a test-owned virtual disk (option A). Mark `D:\RedSalamander.Perf` as a sandbox root, and reinforce the rule that no test file or test-created link/mount may exist outside a marked sandbox. | Remove the real-volume mount from F7 first. Then add the D: marker through the governed initializer. Then implement the VHDX route (skips when not elevated). Then strengthen the AGENTS.md / `Testing_SelfTests.md` rule and add a source contract forbidding real-volume mounts and outside-pointing links in test code. Order matters: marking D: before the mount code is gone would make the forbidden mount eligible. |
| 14 | **Refuse new work** while a "Cancel all and exit" is pending | New admissions are refused with a localized message. A further close attempt does not re-prompt; it brings the popup forward. The app-close consent counts only cancelable tasks. |
| Pane Delete and continue-on-error (2026-10-07) | **The code must follow the human decision:** the BulkItems executor (pane Delete/F8) honours continue-on-error exactly like the per-item executor. | When continue-on-error is on, a failed Recycle is kept, reported Failed and skipped without any prompt; the escalation prompt (and the task-scoped "Delete all remaining permanently" choice) appears only when it is off. **Done 2026-10-09** with the Q2 task-scoped escalation (`Phase10_RecycleEscalationBulkItemsContinueOnError`, `Phase10_RecycleEscalationTaskScopedGrant`). |
| 18 | Proposal approved | Plan IDs are never reused. DxUi plan I27 becomes I31; Scoped testing I28 becomes I32; Release dependency recovery keeps I28. The open Local route-observation obligation moves to I29. A spec-index check rejects duplicate IDs and IDs claimed by archived plan titles. **Done 2026-10-09** (renumbering, I29 Q2-L routing and the `WIP_DUPLICATE_RANK` / `WIP_RANK_ARCHIVED_ID` checks). |

1. **Cancel semantics.** Should roots never dispatched before a Cancel (or before the first failure cancels the
   task) be NotAttempted/Retained, as before `b883d5f8e`? *Recommended: yes; "Outcome unknown" only for roots
   that reached a provider.*
2. **Recycle on volumes without a Recycle Bin.** Should Local fail the item with a known non-commit so the
   RecycleFailed escalation asks for permanent-delete consent? Should SMB/removable routes still advertise
   `recycle:true`? Should a Delete of N items on such a volume get one "Delete all permanently" consent instead
   of N prompts? *Recommended: yes, keep advertising (escalation handles it), and add one task-scoped
   escalation consent.*
3. **Recycle escalation on folders.** Recursive permanent delete with counted consent, or not offered for
   directories? *Recommended: recursive, showing discovered item and byte counts.*
4. **Retarget overlap "Don't start".** Keep the current destination (relabel the button), or cancel the task?
   *Recommended: relabel "Keep current destination".*
5. **Name-collision Escape.** Escape and Enter both cancel the whole task. Now that Skip is a primary button,
   should Escape stay "Cancel task"? *Recommended: keep it, since the spec chose it, but confirm.*
6. **Popup failure surface.** Should transient graphics and layout failures retry and re-attach instead of
   entering the permanent native surface? *Recommended: yes; the failure surface only when Attach itself
   fails.*
7. **In-flight stream rows.** Should each stream be a separate focusable region (Tab stop and UIA element), or
   should there be a stable Source/Destination pair per task with the other streams in help text?
   *Recommended: stable pair.*
8. **"Clear completed".** Should it acknowledge interrupted-Move breadcrumbs in bulk, and clear results the user
   never saw? *Recommended: no and no.*
9. **Issues "Problem" column.** Localized typed text, with the English engine diagnostic only in Technical
   details? *Recommended: yes.*
10. **Cached apply-to-all logging.** One accepted-decision record and Issues row per item (a 300k-conflict copy
    writes about 300k JSONL lines and evicts real errors), or one per originating prompt generation?
    *Recommended: per generation, with a count.* *Implementation note (w6-2, 2026-10-10), awaiting owner
    acknowledgment:* the cached applications of one originating prompt are one record and one Issues row with
    `appliedCount`, but the user's own answer to that prompt keeps its separate record and row (joined to the
    prompt's Warning), so each Apply to all prompt shows two rows: the user's choice and the automatic aggregate.
    Folding the aggregate into the prompt's row would make the projection join a User and a CachedScope decision
    (and the prompted item's result with the aggregate's) without marking the row ambiguous.
11. **Dependency process.** Ignore Dependabot's vcpkg-baseline and AWS bumps (audit-gated per `FileSystem_S3.md`)?
    Require green CI before merging, even during a billing outage? *Recommended: yes to both.*
12. **F7 mount.** Is mounting a whole user volume acceptable on developer machines, or should F7 use a
    test-owned VHDX? *Recommended: VHDX.*
13. **Release gate.** Is `d93dea1c6` releasable while I29 Fresh Full Debug/Release on the API-3 pin is
    unpassed? *Recommended: no; block release until I29 Q1 passes and Waves 0-2 land.*
14. **Deferred exit.** After the user accepts "Cancel all and exit", should new operations be refused, and
    should a second close re-prompt? *Recommended: refuse new work and do not re-prompt.*
15. **Perf evidence.** AGENTS.md requires before/after perf pairs. Do you waive them for the merged ViewerSpace
    Slice 0 (X2/X14/X16) and the popup/Issues cadence changes, or must they be captured before the next change
    in those areas? *Recommended: capture before the next change.*
16. **Test-enabled Release.** Are ENABLE_TESTS Release builds ever distributed? *If not, Simplification 1 can
    move about 6k hook lines out of shipping TUs safely.*
17. **Search service.** `e860022a0` carried unreviewed search-service changes: server-side regex,
    unready-store cutover, progress flush. Should over-complex regexes keep failing as silent non-matches, or
    raise a warning flag? Should that commit get its own review (coverage gap G1)? *Recommended: warn, and
    review it.*
18. **Plan IDs.** I27 is used by two plans, and the working-tree index has two I28 rows. Which keeps each ID?
    This plan uses I30. *Resolved by owner decision Q18 and renumbered on 2026-10-09: DxUi I31, Scoped testing
    I32, Release dependency recovery I28, the archived safety plan keeps I27. This plan moved from I30 to I36 on 2026-10-11 because its own archived delivery record claims I30.*
19. **Native Move with a vanished source.** Indeterminate (blocks Retry, shows "Outcome unknown") or a known
    failure with the real error? *Recommended: known failure.*
20. **Preferences Apply.** Should Apply and "Reset all" preserve popup-owned view state (throughput disclosure,
    density, Issues layout)? *Recommended: preserve it.*

## Coverage gaps (second review round)

The completeness critic found areas no reviewer covered in depth. Schedule them as one focused round after
Wave 0.

- **G1, search service:** server-side regex, unready-store cutover and progress changes in
  `SearchServiceBroker.cpp` and `LocalSearchIndexCore.cpp`. The service runs as LocalSystem, so review it for
  security and DoS.
- **G2, plugin unload/reload:** `PluginModuleLifecycle.h` deferred-alias changes, for multi-plugin DLLs.
- **G3, upgraded dependencies:** behavior of yyjson 0.12 to 0.13 (settings and breadcrumb round-trip), curl,
  sqlite3, 7zip and libjpeg-turbo for their consumers.
- **G4, DxUi upstream:** semantic diff of DxUi `f5c7fc7..49a9988` (modal loop, grid selection/focus,
  accessibility over registered messages).
- **G5, route observation latency:** `TryObserveCachedLocalRoute` on UI-thread callers with offline UNC or
  removable targets.
- **G6, logoff and shutdown:** `WM_QUERYENDSESSION`/`WM_ENDSESSION` with a running Move (breadcrumb
  durability).
- **G7, diagnostics export:** JSON/TSV projection escaping in `State.Diagnostics.cpp`.
- **G8, providers without receipts:** typed results for MTP, 7z and Dummy, which never set mutation receipts.
  Every success may show "Outcome unknown".
- **G9, unreviewed Commands self-tests:** about 13.6k lines in `Commands.Tests.FileOps.cpp`; check for real
  desktop input and false passes.
- **G10, tests-off Release:** never built or exercised in the range; confirm the release workflow uses
  `RSBuildEnableTests=false` and compiles.
- **G11, LocalizationManager:** the process-wide mutex is now taken on every popup render.
- **G12, layout-helper anchors:** re-anchor `popup-c#12` (the cited `FileOperationsFooterLayout.h:252` exceeds
  the file length).

## Closeout

- Each wave closes when:
  - its regression tests pass in the owning family;
  - the owning spec is updated;
  - its ledger IDs are marked fixed here with the commit.
- Waves 0-2 plus I29 Q1 gate any release.
- When every wave and the second round are done, move this plan and its ledger to `Specs/Plans/Done/`.
