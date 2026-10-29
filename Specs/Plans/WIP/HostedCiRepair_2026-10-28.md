# I51 Hosted CI repair — 2026-10-28

Status: ACTIVE. Priority: high. Planned at `fbcc9fc0ddd3b6890741f6bc350f2016c755614b`;
integrated main baseline `8ea956a7e2e271368082a953f28efd2a58e1013f`.
Drift check: `git log --oneline fbcc9fc0ddd3b6890741f6bc350f2016c755614b..origin/main`
and inspect the affected workflow, source, tests, tooling and owning-contract diff before integration.
Ownership: RedSalamander hosted/runtime repair; RedMonitor formatter closeout remains a separate merged record.
Done criteria: causal fixes, unchanged acceptance gates, applicable local/hosted qualification and green main.
Stop on an unowned process/path, unsupported required environment, missing source/artifact identity,
or focus work outside the agreed time; retain the failed/incomplete receipt rather than bypass admission.

- [x] Isolate this repair from unrelated primary-checkout work.
- [x] Retain original nightly/ASan logs and all six validation artifact archives.
- [ ] Classify each failed case using original broad-run and retry evidence.
- [ ] Reproduce focused failures and implement causal fixes with unchanged assertions and budgets.
- [x] Repair current scheduled checkout/staging policy without changing sealed historical workflows.
- [ ] Validate affected tooling and native behavior locally and on hosted x64/ARM64 profiles.
- [ ] Merge reviewed fixes and confirm push, nightly, ASan, flake and Ghostty main workflows.
- [ ] Reconcile durable contracts and record completed qualification before moving to Done.

Baseline: public main `fbcc9fc0ddd3b6890741f6bc350f2016c755614b`.
Original nightly run `40473021621` fails native x64/ARM64 Debug/Release. ASan run
`40484804050` fails ARM64 runtime cases and cancels x64 after six hours. Flake run
`40477285880` and Ghostty status run `40477455425` fail Windows checkout on long paths.
Push PR-suite run `40433814634` passed; that narrower run does not close these failures.

Preserved baseline artifacts and the diagnostic failure ledger live under the existing
approved `E:\RedSalamander.Perf\evidence\hosted-ci-repair-20261028` root. A later passing
retry does not erase an original failure or prove its cause. Keep configuration, source,
dependency, environment and scope identities with every new receipt.

No failure may be removed by reducing coverage, loosening an assertion/performance budget,
turning a failed prerequisite into a pass/skip, or treating pending hosted work as accepted.
Local focus-taking runs require agreement to their timing and the existing warning/lease
and restoration. Native ARM64 runtime acceptance comes from native execution.

Owning contracts to reconcile as fixes are established: `Specs/Testing/Testing_SelfTests.md`,
`Specs/Testing/Testing_ValidationEvidence.md`, `Specs/Testing/Testing_ToolingGovernance.md`,
`Specs/Build/Build_Toolchain.md` and any affected product domain contract.

## Established causes and candidate repairs

- The formatter repair is owned and completed in RedMonitor PR #2; main run 40671228756
  passed all seven jobs at `1e3e9d6059728e90542e013876c1c9783234f505`. Its plan closeout is
  separate from RedSalamander native acceptance.
- Both failed scheduled consumers omitted Git long-path configuration before checkout.
  The shuffle consumer also lacked explicit sandbox initialization and bounded upload selection.
  All 38 release/workflow policy checks passed locally, including exact-byte deep checkout.
- A real occupied search-service pipe deterministically fails the original 750 ms connection
  limit. The strengthened existing multi-client case failed with that premature return in local
  run `20261028T102900Z-52128-19c0a6bd4b3942ab90c7fbcedcfd6d15` on build receipt
  `7f29e8ab0712a07d65532e40ed28c56db0ef6434e40a056daa7f3443508d7acf`.
  The candidate uses the existing frame/operation budgets while preserving short missing-pipe fallback.
- Hosted four-worker selection-shape receipts report exactly three active leaf transfers versus
  four folder transfers: the shared scheduler counted its starvation reserve inside that pool.
  The candidate adds that reserved worker to the bounded I/O pool; exact concurrency assertions remain.
- At hosted DPI, the wide 260-line note does not exceed the required 4096-pixel height. The fixture
  now uses 520 lines and a stronger line-count assertion while retaining the height/scroll checks.
- Animation-phase fixtures inherited the OS reduced-motion default. They now select the intended
  animation policy through supported settings, restore it, and retain the dedicated On case.
- The original ARM64 runner had one fixed volume but an already validated harness-owned VHDX.
  The independent-volume case now admits that exact mount as its final fallback, preserving all gates.

The local focused Commands subset passed before these runtime changes. The same campaign's FileOps
entry rejected a comma-separated union filter with no matching cases: that entry is invalid coverage,
not a runtime result. Subsequent FileOps reproductions use its supported exact case/family selector.
Further runtime and hosted validation remains open; none of these candidate changes is a green-main claim.

## Follow-up validation and review

- Candidate `53fe837c8b9978895cb60c0a647869948390cfb8` built x64 Debug successfully,
  receipt `be8e8b4d2aee9d2606dc347966a87486da6175eca0907e3a722101e1e7ae8f84`.
- Focused independent-volume run `20261028T110609Z-58744-6475e570cc05408eadfdc344112cd77d`
  passed Setup, the unchanged independent-volume gates and Cleanup, with no skips or retries.
  This is local focused coverage, not hosted or broad-family acceptance.
- Ghostty Online observer run `40677907341` passed on that candidate. Its unchanged observer
  report is retained in `candidate-ghostty-online` under the repair evidence root.
- Review identified an overbroad connection-wait change, operation deadlines beginning after
  control connection, a missing alternate-volume opt-in in the shuffle consumer, and a fixture
  restoring settings without refreshing live theme state. Correct those before acceptance:
  only queries receive the longer contention budget; status remains short; all control budgets
  include connection; the occupied-pipe regression also checks prompt status/shutdown failure;
  the shuffle lane explicitly opts into alternate coverage; both fixture scopes reapply restored UI.

The no-journal index path also contradicts the owning Search contract: its fallback calls generic
`BuildIndex`, which can enumerate the complete NTFS MFT, although the contract and original failing
case require tracked-root traversal when the journal is unavailable. A deterministic unavailable-journal
probe and an MFT-attempt witness now exercise that distinction while native MFT seeding remains eligible.
Regression run `20261028T114657Z-20084-58c602e9b16e47dcb241772da654700e`, receipt
`216643ba312fe899c658854785286193f188ae5c9af66b460fb50c2a49ded8bc`, failed deterministically:
`journal=false traversal=false nativeMft=true attempts=1`, taking 48.9 seconds for the tiny fixture.
The failed run also exposed a missing expected-case inventory entry; that coverage error is preserved
and the new case is now registered. The production unavailable-journal branch now calls tracked-root
traversal directly. Normal journaled MFT seeding and the original replay/candidate assertions remain intact.

Plugin contract failures now mirror their exact assertion into the existing captured perf JSONL stream;
the GUI host did not preserve the FTP export's stderr, leaving only its opaque 141/142 aggregate before.

The shuffle artifact consumer now provisions the same pinned MSVC family before portable attestation
rebind; matching runner labels alone do not ensure the producer-installed compiler is present.

## Integrated qualification

Candidate `6dd56d99c55db8f141570d165f305242f0b20cf2` built x64 Debug successfully with receipt
`efe0fcd9c7540e2364791ac3de92ebfb8c097fa0e12ddae0e7dc828973b530e3`.
Its six focused Search cases passed without skips or retries in
`20261028T121410Z-76456-24af46693e4d43a2bd202be44ca51afd`; the no-journal regression took 40 ms.
The exact history-menu and delete-focus cases passed in
`20261028T122212Z-78288-979e5f220c8b4c9c8df9333f5227f3f3`. A mixed prefix/union filter selected
only those two exact cases, so this is not animation coverage. The remediation-contract case and
its Setup/Cleanup passed in `20261028T122824Z-67024-cde935e3e3584d538f127b589b9f3a51`.
All 304 focused tooling controls passed before integration; they do not qualify the newer source.

Main advanced concurrently through PR #145 to `8ea956a7e2e271368082a953f28efd2a58e1013f`,
including the canonical PowerShell formatter installer and guidance/tooling corrections. Preserve both
changes. The newly started local Full run was interrupted before native execution and remains incomplete.
Earlier hosted runs remain tied to their original source; none is promoted to integrated acceptance.
Resolve the changelog merge by reconstructing the six existing repair fragments and folding them over
main's generated changelog. Refresh the reviewed PR execution digest after inspecting the complete
combined workflow/module closure; this records reviewed inputs, not a CI pass. Integrated local and
hosted qualification, plan closeout and main confirmation remain open.

The integrated `b79613627d35ec1a84bc4722297d298797b4f20f` x64 Debug rebuild passed in 5:23,
and 428 relevant tooling controls passed without skips. Inventory validation caught a missing
machine-readable WIP table row; the repair was indexed as rank I50; integrating main on 29 October allocated unused rank I51 because the separately merged main review owns I50. No frozen rank is reused.
The retained hosted desktop log proves the real x64 display changed from 1024x768 (work area
1024x720) to 1920x1080 (work area 1920x1032); its foreground owner was Windows Terminal.
Commands isolation now reuses the existing warning-gated queue handoff and explicitly requires
owned foreground, instead of ignoring raw activation return values.

The pinned DxUi popup seeds its pointer sample only when it has no initial keyboard row. A
keyboard-opened popup can therefore treat the first stationary event as movement and clear its
keyboard row. Add a deterministic create/End/stationary-event regression before changing the pin,
with a real-movement control and existing warning/lease/restoration. The suspected library cause
is not yet a runtime-confirmed fix, and existing focused history success does not close it.

Regression source `1512f7aaa9d9d1cdd0d95dc7ec0bb40c433c09f8`, build receipt
`2318d0f2122671bac3a4c550a2a0f3f1d7d1dd276afe71a96ac1fdbc2a3a5aea`, reproduced the pointer defect
in `20261028T131806Z-78428-c12b0ffda10b4833a9f289d0371f8180`: keyboard row 2 became unset and hover
row 0 replaced it after a stationary event. The exact new case failed in 70 ms, with no skip or retry.
Canonical DxUi already contains the capture-time history repair. The owner explicitly selected the
broader consumer adoption of `881e5d5ff252d5988c1357a5c64d14f480b46793`; main push run `40622836997`
passed all six native profiles, validation, formatting, Windows tooling and the CI gate. Its paired
benchmark was skipped, and existing performance/platform holds remain separate. Keep API revision 4
and the single target, and qualify the new consumer source without editing the sibling checkout.

Integrated shuffle run `40691121033` passed portable import but failed receipt use before native tests.
Its fresh consumer had no restored DxUi source; the first use-time identity guard correctly rejected it.
Restore the exact pin before import, preserving that guard, all artifact hashes and repeat/shuffle controls.
Check the child's exit inside each protected action, where its status is available, and propagate failure
through display restoration. A real nonzero-child regression complements the restoration controls.

The broader pin's first consumer build (`0e12ce08f`) failed with four C2589 diagnostics: the Analysis
and Navigation host headers included Windows without `NOMINMAX` before the new public RichDocument
header. Apply the same Windows-header boundary already used by the other ViewerSpace host/header;
do not edit managed dependency products. Repeat a full-solution rebuild and the original regression.
Both specification and tooling inventories passed without blocking findings. The workflow/receipt
controls, including a real child exit 7 and display restoration, passed 95/95 with no skips.

The next broader-pin rebuild (`e53451e09`) reached the application and found two obsolete local
`void CompactChildren` adapters conflicting with the new public counted compaction API. Remove
those duplicate erasers and the otherwise empty scroll-panel subclass, use the public Panel and
ScrollPanel API, and retain its tree/accessibility invalidation. Keep all reconciliation, pending
decision, focus, lifecycle and resource assertions; validate those paths before acceptance.

Candidate `1f3cd13f7ab342dcc5812677ea020492efdf695a` passed its full-solution x64 Debug rebuild
in 13:03, receipt `650cf08b4103fd5b85b3777025e907527786c83fcc33434b3525188d42e07368`.
All 14 exact Commands cases passed in `20261028T143733Z-26820-ea86d5ea88ef4574957e6530f1f271d5`,
with no skips or retries: the original history case, the deterministic stationary-pointer regression
and all twelve theme-overlay cases. The pointer regression took 59 ms; its real-movement control
remains required. Popup stable-tree Setup/case/Cleanup passed without skips in
`20261028T144733Z-73716-425c1baf430a479b841c34cec4b23f24`.

The next remediation run, `20261028T145445Z-73716-17432a0e510a4343a891c176f90b36ad`, failed
its unchanged final foreground-equality assertion after all four auto-collapse scenarios passed.
Its original trace lacks the before/after process identities, so do not attribute that change to
ambient input or the candidate without evidence. Retain the failure and add identities while
keeping the same assertion. The sequential campaign stopped there; later selections did not run.

The superseded hosted `b79613627d35ec1a84bc4722297d298797b4f20f` x64 Debug terminal Commands
record retains 956 passes, two failures and six declared skips; its last retry exposes a remaining
speed-dialog layout capture failure. That callback runs on a worker and bypasses the screenshot
helper's owner-thread capture admission. Route just that capture through the prompt's existing
synchronous test-message handler, using the canonical helper on the owner UI; require successful
capture and exact extended-style restoration. Do not change capture, geometry or foreground gates.

Fresh candidate workflows `40709199266` (PR), `40709196329` (nightly), `40709202121` (ASan)
and `40709208031` (shuffle) were dispatched at `1f3cd13f7`; observer `40709212481` passed.
Those results remain bound to that source and do not qualify subsequent capture repairs.

Capture-bridge source `e04bb4ffed003a5b04c6e7608aa7e26c32ebf91f` passed all 282 focused
source/specification tooling checks without skips in governed run
`20261028T151428Z-58492-137ea5fbcf314a119cd808528e8aa52c`. Its native build failed C1083:
the application root needs the explicit repository-relative TestSupport header path rather
than a bare include name. Correct that boundary and rebuild before native acceptance.

Source `a9598b9c0262a572670976271110a337be24d1fa` rebuilt x64 Debug in 9:10, receipt
`37725ddf63722bb6a5d46fc719a5ef1bf1331f29461410ffc9b4f260c89b5d19`. Its focused speed-dialog
run `20261028T153154Z-60116-607957bed6824dd0b6af32ec1c6c2265` failed in 2.5 seconds, without
skips or retries: the UI-thread capture bridge returned `0x8001010D` (the original signed-format
diagnostic prints `0x-7FFEFEF3`). WGC cannot call out from that input-synchronous SendMessage
handler. Keep COM capture on the interaction worker and route only canonical admission and
restoration to the owner UI while it pumps. Extend the same admission policy to ordinary owned
modal windows, which also lack a taskbar entry, and require the APPWINDOW bit while retaining
all other bits. Add a worker-refusal control, exact restoration and unsigned HRESULT diagnostics.
Use the existing `TestSupport/` include root to remove the new C4464 warning rather than suppress it.
None of these candidate repairs is accepted until the new runtime and hosted receipts pass.

Source `4b094967d1f2175105b9c754aa81123d86d4e8e3` built x64 Debug in 8:01, receipt
`6ab41572cfdce08c453086bd3aa7668dcd986db19b8bc2669637590c02d52678`, with no new warnings.
All 282 focused source/specification controls passed without skips in governed tooling run
`20261028T155818Z-72880-4c6c50d91ce3427f9888a23025ad83d5`. The exact speed case passed in
3.983 seconds without skips/retries in `20261028T160422Z-86640-8f259136c2fa4dc892393ee46af198da`;
initial and short French harness captures were inspected. These are focused results only.

Review comments `4403208253` and `4403208365` identified valid request/teardown defects:
the new WM_APP data was still a raw pointer, and DxUi handles WM_NCDESTROY before the
late switch cleanup. Stop the local PrePush campaign (its BuildToolchain Pester entry was
in progress), retain its incomplete record, and fix both before acceptance. Convert all
data-bearing commands in this prompt channel to canonical typed tokens with an embedded
operation and shared atomic response. Initialize/drain the payload window; retire admission
before host dispatch. Add malformed-request and active-admission nested-shutdown controls.
Source inspection also found the fifth short-locale capture bypassing setup; route all five
through the same scenario-local capture helper, retaining every original assertion.

The PrePush explanation verified the unchanged reviewed PR digest and six local obligations:
BuildToolchain tooling, Compare/Search, Commands, FileOps and the two interactive viewer groups.
The interrupted campaign is not a local remainder pass or a Full result. Current-source hosted
qualification must follow the corrected code; earlier runs remain diagnostic source-bound records.

### Residual qualification after the capture repair

Source `173ef7434aa4f3e05093487c831ffad3d81ae7fc` built x64 Debug in 5:41, receipt
`37261316683ddd32cb9fbd8b8c402d67f5ddfecf566edbc1c1082939004e39c1`. All 282 focused
tooling controls passed without skips in `20261028T165535Z-86156-a530ce5cfbdd45c19fc31b0983eca9c6`.
Its strengthened speed-dialog case passed in 3.429 seconds without skips/retries in
`20261028T165503Z-78724-d613e8ca64004c8583c79cb4f1ee09d9`, including malformed tokens,
all five captures and held-admission destruction. Normal and short French captures were inspected.

The same source reproduces the settings ordering-child timeout locally at 12.002 seconds
against 12 seconds in `20261028T170424Z-80468-41cd243fa0694c50b944d5f18f4e57ea`. The child
stops before UI isolation/body entry while the parent blocks its foreground queue.
Run `20261028T171441Z-86368-4447dedf6a294411a0f30d1ca2a13a1b` reproduces the premature
Manual-buffer seed and empty Shortcuts handled-copy mismatch; durable mechanics and theme
performance pass locally. Its 200 largest-graph resolver samples are 7,210–13,364 us,
mean 7,898.79 us. These are retained baseline observations, not candidate acceptance.
The two original journal cases pass locally without skips in
`20261028T172456Z-79636-8025e69bf5cc41919d0cfb5eea31849c`; they do not reproduce ARM64's
driver rejection and do not close that hosted gate.

Additional original `1f3cd13f7` artifacts retain eight distinct residual failures across profiles:
settings child, speed capture, empty-grid copy, Manual seed, durable private sibling cleanup,
theme resolution, and two journal seed queries. Its x64 ASan 200 largest-graph samples are
80,818–120,999 us, mean 98,026 us, against the unchanged 100,000-us guard. ARM64 ASan
retains an incomplete native MFT scan after 9,134 nodes and the two `E_INVALIDARG` seed failures.
Keep that failed source, per-case ledger and raw streams.

Repair the UI wait with the canonical bounded pump, the premature mode transition with busy
admission plus the existing blocked-provider control, and the new pin's empty-copy contract
with explicit unhandled/exact-clipboard assertions. Reduce resolver duplicate ownership and
cache allocation, retaining every graph and timing oracle. Roll back through the owned file
handle and add a still-open-reader control. A driver-rejected native enumeration falls back
to complete root traversal, while a malformed-data negative control still fails.
These latest changes are unbuilt and unqualified until fresh receipts pass.

Source `5221736ddf217d8f7806771acd33c3e2d82bfee8` built x64 Debug with receipt
`cec809bd7dee0926c4c50d931150403491bc01d32c18e4755eeef7e523af7f6a`.
Focused run `20261028T175950Z-17620-8aed146ddc0745b2b4c8651eef16edf4` confirms all
four settings child modes exit naturally within the unchanged deadline (4.936–6.594 seconds)
and the 200 largest theme resolver samples are 4,585–10,402 us, mean 5,239.59 us.
The process stops during blocked-provider Batch Rename before five remaining case results;
its count-mismatch verdict is incomplete coverage, never a campaign pass. Review found
timer cancellation touching freed controls after DxUi host teardown. Remove that late
UI update and retain the existing blocked-provider close/recovery assertions.

Review `4404286278` correctly identifies parent environment leakage during the pumping child
wait. Use the canonical child-only Unicode block and assert exact parent environment values
throughout the wait. Superseded hosted `4b094967d1f2` retains a stale cp437.zip pane snapshot
(two displayed entries versus three disk entries) and ARM64 ASan foreground refusal against
WWAHost. Retain both failures. An in-flight directory invalidation must survive publication;
add a gated post-snapshot mutation control and retain the original archive test. Use a warned,
owned registered-hotkey foreground handshake with refusal/cleanup controls, retaining actual
foreground assertions and the existing focus deadline. These repairs require fresh native,
hosted and local-remainder qualification; source-522 workflows are diagnostic until then.

Source `c28ecab60db2586846908e2ef0593b97df670d36` passed its full-solution x64 Debug
build in 5:19, receipt `0bda355e7897c6ca0ff748a14f7be52fd93f6b21cad848d244fe96f9ea413f83`,
and all 282 focused tooling controls without skips in
`20261028T174151Z-68132-e5362f038e2347579b99f8eb2bedd025`. Review caught the new
enumeration control reusing a synthetic unchanged journal position after file mutations;
advance that fixture position before native execution. The product currentness proof is
unchanged. Rebuild and qualify the corrected source before acceptance.

### Focused qualification and remaining MTP timing

Source `58d1e03fdef9ec93e0cfe63139eed80461c9ca18` passed its full-solution x64 Debug
build without new warnings in 6:55, receipt
`b4f599f045639901f9f8b063eedee11ee20befecc8dab53305ef631f9cd5ab03`.
All ten focused Commands cases pass without skips or retries in
`20261028T185620Z-83432-9a6524d619a24c3b9552cee37f760974` (32.802 seconds).
This includes all four natural settings-child exits, theme guard, both Batch Rename cases,
the open-reader rollback, unchanged clipboard, five captures, hotkey admission/refusal/cleanup,
and the gated old-snapshot invalidation race. The latter requires two provider calls and
observes the current snapshot in 384 us. Initial and short French captures were inspected.
See the [complete focused packet](../../TestRuns/4cb089111a23/Commands/2026-10-28_185620_HostedCiRepair-Commands10-Focused/evidence.json).
Its entire original metric stream is retained in two ordered, byte-preserving shards.

The same source passes all eleven journal/search/MTP baseline cases without skips or retries
in `20261028T190623Z-68876-624d286bf52c40879e81f6c05f64103f` (283.392 seconds).
All five journal/index controls and all four service transport controls pass. The existing
4,096-identity MTP recovery fixture measures 232,237 us in local Debug before optimization.
The [baseline packet](../../TestRuns/4cb089111a23/Compare/2026-10-28_190623_HostedCiRepair-SearchAndMtpBaseline/evidence.json)
retains complete case results, ten claim-bearing metric rows and a digest-bound summary of
all 15,595 original rows; the raw 5.3 MB stream remains in the governed run.

The exact FTP control passes all 142 assertions and all eight 4,096-file matrix variants,
without skips/retries, in `20261028T191809Z-68876-560949bb942545478dd2e078ae8800be`.
Its native case takes 232.375 seconds; Setup and cleanup also pass. The
[packet](../../TestRuns/4cb089111a23/FileOps/2026-10-28_191809_HostedCiRepair-CurlCopyTruth/evidence.json)
retains the complete original metric stream. The earlier 141-of-142 failure did not reproduce;
do not attribute that passing repeat to a causal Curl repair. Broader/order qualification
and the retained original failure remain necessary.

The [theme observations](../../TestRuns/4cb089111a23/Commands/2026-10-28_185620_HostedCiRepair-ThemeResolverPair/evidence.json)
retain every original 200-sample resolver and 250-sample dynamic row for baseline `173ef7434`
and candidate `58d1e03fd`. Resolver p50/p95/max are 7,745/8,761/13,364 to
4,779/5,107/6,507 us. Dynamic p50/p95/max are 141/242/603 to 135/146/1,161 us;
the candidate maximum outlier is retained. These same-machine Debug observations have
different campaign order and do not establish independent statistical non-regression or
an ASan speedup. No threshold or baseline was changed.

Superseded source `5221736dd` x64 ASan exposes one additional failure, the existing
`mtp_r0c_exact_recovery_identity` two-second ceiling: 2,235,725 us originally and
2,096,910 us on isolated retry. Retain [both failures and original case statuses](../../TestRuns/136659e3341a/Compare/2026-10-28_183027_HostedCiRepair-MtpAsanFailure522/evidence.json).
Its first directory enumeration alone takes 1,047,256 us. Repeated canonical MTP paths
currently rebuild through separator and scheme normalization in enumeration/recovery loops.
Return their exact owned UTF-16 copy directly; preserve the complete normalization route
for noncanonical inputs, all identity/recovery checks, the 4,096 objects and two-second cap.
Expand the existing normalization case to seven root and five exact-content read aliases.

Review `4404905671` validly identifies a new focus deadline including unrelated preview/terminal
retirement. Deduct only measured foreground-handshake time from the existing 1.5-second
focus wait; retain the separate teardown deadline. These latest source changes are unbuilt
and need fresh native, hosted and local-remainder qualification. All five curated packets
pass explicit pre-stage archive validation; inventory and final closeout remain open.

### Archive boundary lost by the history-free import

The required repository inventory is red at source `58d1e03fd`: the original contract
`8b9e36191835cc71b5ee0dfeea4dddd5d942dd5a` is not an ancestor of the public snapshot.
Its fallback validates all 19,500 tracked artifacts, including grandfathered June/early-July
logs over the current size ceiling and old mismatched machine-profile folders. Retain this
failed outcome; do not change legacy bytes, raise ceilings or describe the inventory as passed.

Construct parentless baseline `4a9489e9cb2d1190528b8ec35971fba885fd7792`, tree
`3fb1f79ac855236a8cc2e0bf9b348601ec8cd5f3`, from only the 2,075 exact path/mode/blob
matches between original pre-contract tree `eaf640798eb0fa21f2b5b6e3883891ac2442a255`
and already-published source `58d1e03fd`. Every entry is under TestRuns and already public;
the new commit has no parents and imports no DualTail history. Its
[construction proof and failed inventory record](../../TestRuns/4cb089111a23/Validation/2026-10-28_200100_ArchiveContractBaseline/proof.json)
are structural provenance, never runtime acceptance.

Use this exact baseline only when it is an ancestor, preserving the original full-history
contract when available and the fail-closed all-files fallback otherwise. Existing limits,
schema/identity checks and explicit legacy RunPath validation remain. New committed archives
and committed/staged/worktree legacy edits must all be rejected when invalid; an unrelated
baseline object cannot grant grandfathering. Compare branch additions against public main
before the legacy master fallback. Fetch complete public history in native CI so the real
inventory test executes. Add regression controls for every boundary. Preserve the parentless
baseline through a normal merge of PR146; squashing it out would lose the ancestry proof.
The original failed inventory remains recorded and the corrected inventory is pending.

Candidate `0906482f4ab41db6a9194db298febff421e5f582` builds the full x64 Debug solution
without new warnings in 6:18, receipt
`21a119e97980f2f65452c53b9c618feb85a93fbe6e065365e3f72fcb5e69c647`.
Governed run `20261028T201034Z-69960-b26347c43c0645808ac2557407294993` passes all 300
archive/source/specification checks with zero skips. The real complete archive inventory now
executes and passes (299.84 seconds); the four new migration/rejection controls also pass.
The explicit corrected Inventory entrypoint and focused MTP/native handoff/gallery checks remain.

Review the enabled PR execution surface after complete-history checkout and archive helper/test
changes. `ci.yml` still runs the existing noninteractive Suite PR in native x64 and ARM64 Debug;
in-product suites, interactive viewers and build-toolchain work remain local obligations.
Bind reviewed coverage to digest `9a4fe488ec02ae7a41a0549823b119c22e0a569206ba02a1ab8ce552b1a2cae2`.
The changed binding is not hosted acceptance. Rebuild the final candidate, verify its scoped
coverage explanation, and retain CI_PENDING until the matching PR checks actually pass.

Candidate `032547ca7` builds the complete x64 Debug solution in 5:46 without new warnings,
receipt `b46267cb90ac7f5e70e00913b60e05b15cbcd27498cb2c22f74a19dda4b1eedc`.
Governed run `20261028T202241Z-90024-cde963f1b9cc4f28a62ec40f168fe915` passes all eight
scoped-testing contracts and the separate corrected archive Inventory entrypoint.
The original failing inventory remains preserved.

### ARM64 asynchronous image-test readiness

PR run `40749715601` at source `58d1e03fd` passes x64 Debug but fails the ARM64 Debug
ViewerImgRaw orientation case: TIFF and WIC-fallback snapshots are inspected with no image.
The test's `DecodeCompleted` incorrectly treats any `ClearAlert` as terminal, although
`BeginLoadingUi` clears alerts before the worker reads the file. Replace all three callers
with the existing UI-applied final-result counters plus request identity and ended loading.
Keep the eight-second completion deadline and all original warning, image, orientation,
menu-ownership and resource checks. Gate the actual TIFF provider read using the existing
blocking-reader fixture: startup must have cleared alerts while the request is still loading
and has no final result; only explicit release may allow completion. Require no fixture timeout.
The PNG probe additionally requires a retained decoded image. This correction is test-only;
fresh native and ARM64 hosted validation remains required before claiming the failure fixed.

Candidate `624f9552b` passes the complete x64 Debug build in 5:46 with no new warnings,
receipt `d26e5e3f00cb0f106efc68c66d05e4975211ac044f5be48c86364f7d9000a30a`.
Governed run `20261028T204108Z-83756-15cf203bc6994790a55373563c49ba3a` passes all 282
source/specification controls without skips. These are not native runtime acceptance.

Retained partial source58 ASan artifacts expose two remaining failures: x64
`cmd_fileops_durable_store_safe_mechanics` still fails the immediate private-sibling cleanup
assertion; isolated ARM64 Find retries fail before entering their cases because foreground
admission returns false. Preserve those failures. The hotkey helper currently removes its
message and activates outside the window procedure. Dispatch it through a scoped subclass
on the owned warning and activate inside the actual handler, retaining the 500 ms bound.
Strengthen the existing native control with an independent owned-thread window-procedure
hook witness; actual foreground, key release, registration cleanup and refusals remain.
For the unreproduced rollback failure, record each disposition/delete Win32 outcome plus
the exact enumeration status, entry count and entry attributes without retries or relaxed
cleanup assertions. No further causal rollback repair or foreground acceptance is claimed
until new native/hosted evidence establishes it.

Review comments `4405923108` and `4405923200` are valid. A queued close can delete
BatchRenameWindow during its synchronous debug pump, before the next predicate or caller
UI access. Reuse its existing dispatch-depth/deferred-delete scope for the pump and outer
settling methods; reject post-close results and compare exact instance identity at global
adapters. Extend the real blocked-provider case with queued closure during snapshot settling
and a proven pending-preview flush, retaining the original close bounds and explicit release.
Archive grandfathering is path-specific: disable rename detection at all six committed,
worktree and index diff sites. A regression fixture enables rename detection, moves an old
oversized blob, and requires both staged and committed new paths to fail the same limits.

Source `986156585` fails its x64 Debug solution build on ViewerPE's missing common-control
link dependency; its source/spec run `20261028T205942Z-81940-4d1e9260892b4503996d13f3cc900cc4`
is 281 passed, one failed, zero skipped because inserting diagnostics separated the original
fault-injection declaration from its unchanged source guard. Retain both failed receipts.
Supply the test header's Comctl32 dependency and place diagnostics after the existing fault
helpers; no guard is relaxed. A full Rebuild is required after that failed build.

The reviewed execution binding after the archive rename controls is 84e767c50fbf91c1177ee8069c77581f56c86d347d14e4dd4eb34f087e5f4733. The enabled Suite PR and two native Debug profiles are unchanged; pending hosted work remains unaccepted.

Source `887a9b079` builds the full x64 Debug solution in 4:50, receipt
`1fe9f0a7dd320aac4d25c9e5824fbfbc8513e61a7e3db8c84359ce13b7771a85`, with no new warnings.
All 11 review threads are resolved; its latest review has no open findings, while native
acceptance remains pending. The PrePush explanation still selects the six required local
groups (about two hours), with only equivalent PR entries delegated and CI_PENDING.

The hosted migration guard fails on the new native witness's retired comctl subclass APIs.
Keep the guard unchanged and migrate both the permission handler and independent witness
to the existing Win32Callback hook helpers. The witness must restore its outer hook before
forwarding to the handler and prove the exact original procedure restored afterward.
All three migration/surface guards pass locally after this correction; a new build and
native qualification are required. Focused run
`20261028T213339Z-93684-b8a598fcd023434584aeec2d1bf3add9` was interrupted while waiting
for another worktree's desktop lease, before launching its native cases. Retain it as
incomplete coverage. The person approved the remaining focus-taking tests overnight into
29 October with the existing warning, desktop lease and focus/cursor restoration.

### Empty deferred scheduling partition

Source `16d07ae35` passes its full x64 ASan Debug build in 6:17, receipt
`e26403f8a837a3562f89093206654a7096867b2fe58933bc20869497987f36b3`.
Its local PrePush remainder `20261028T221818Z-96528-20173e430f4748dda03713946a170b62`
passes all 29 build-toolchain controls without skips, then successfully re-attests x64 Debug
in 5:44, receipt `6395312c94b2f98787ccbc19612b972cbbb02e2a7d1df61bd31525ed0af0b91c`.
Before launching native cases, the runner throws at its empty cheap-entry selection:
strict-mode PowerShell rejects `$earlyStandaloneConfigs.Id` when only the two interactive
viewer entries remain. The interrupted run is retained; it is not a runtime or Full pass.

Execute the actual scheduling block in four focused controls, including the real deferred
Full plan. The original code fails exactly the interactive-only and real-remainder cases;
empty/all-cheap and mixed cost/order/identity controls pass. Materialize cheap IDs through
pipeline enumeration before excluding them from later work. Preserve all entries, ordering,
desktop admission, assertions, deadlines and incomplete-coverage semantics.
The pending ASan rollback run `20261028T222658Z-77500-c602ca6c8ec14322a92496808aaf8cc7`
was stopped before native launch to permit this source correction and archived as interrupted.
No rollback or MTP native acceptance is claimed; restore the ASan profile through a canonical
full Rebuild after that interruption. Recompute and review PR execution coverage for the changed
runner before a new PrePush attempt. Hosted x64 Debug at source `16d07ae35` is green; ARM64
and the six-profile runtime qualification remain pending.

The corrected actual-runner controls pass all four cases with zero skips. The reviewed
execution digest is now `b34cc09f8a684a2cebb3978a75af9174907b8cc66f213a28bf944f10cb8ab6b6`.
The PR still runs the same two native Debug profiles and unchanged Suite PR contracts;
the runner fix changes their dependency binding and requires fresh matching CI evidence.
Tool inventory validation has zero blocking findings and the changelog is folded.

### 29 October focused qualification and fixture isolation

Candidate `29615cf6f` builds x64 ASan Debug in 7:15 (receipt
`068959a40c7c07b019febd23c10e05fd4195d22dadeaa930b00a9ebc829f95c3`)
and x64 Debug in 7:33 (receipt
`19e836edd66ab817c6c2e9efd5eb995bbfc2c383cc4d5cd72f80ced6bf8a121a`),
without new warnings. Governed run `20261028T225721Z-94716-3f6fcf547c804311bb3027cd94d1f624`
passes 364 source/specification/scoped/run-plan controls with zero skips.

ASan run `20261028T230612Z-72492-fdfe8ff6df644a73a1ca2c0d95b0274e` passes the two MTP
cases ten times: 20 passed, zero failed/skipped/retries. All ten 4,096-object recovery samples
remain below the unchanged 2,000,000 us cap: min 930,844; median 1,464,865.5; max 1,504,989 us.
Keep the [complete raw metric stream and source/build-bound receipt](../../TestRuns/4cb089111a23/Compare/2026-10-29_001519_HostedCiRepair-MtpAsanFocused/evidence.json).
The older hosted failure used another machine; these observations are threshold acceptance,
not a paired speedup or Full acceptance claim.

The exact static gallery run `20261028T231645Z-92276-3eeed7f92fad4fef83a0e2515e7858ec`
passes and produces all 45 nonactivating captures. Retain all raw captures externally, the
complete scenario table and byte identities, and [six inspected original samples](../../TestRuns/4cb089111a23/Commands/2026-10-29_001645_HostedCiRepair-StaticGallery/evidence.json).
The samples show readable layouts and synthetic fixture data; the other 39 have capture
evidence without an individual visual-inspection claim.

ASan rollback run `20261028T231645Z-44800-54177ba5f2f94fe98f939193d75fa2e4` passes its
first invocation, fails its second, and leaves eight unexecuted through fail-fast. Both
invocations report owned-handle POSIX deletion success. The failed listing contains
`oversized.json` from the preceding invocation, plus `record.json`; no private sibling is
listed. Preserve the [failed repeat and exact diagnostics](../../TestRuns/4cb089111a23/Commands/2026-10-29_001645_HostedCiRepair-RollbackRepeatFailure/evidence.json).
Give every invocation an unused child of the existing owned sandbox using the existing
GUID helper, preserving prior fixture records and all cleanup/reader/target/limit/reparse/
concurrency assertions. The older hosted first-iteration failure lacks entry diagnostics;
this independent repeat defect does not establish its exact cause. Fresh hosted ASan
qualification remains mandatory.

Focused Debug run `20261028T232522Z-84764-ccd1b965be1647f2876418f1fe3bacfb` completes all
ten native cases successfully in 32,908 ms, including the 19 ms real-hotkey witness, queued
close during snapshot and pending-preview settling, settings children, five prompt captures,
clipboard and the cache race. Its wrapper is interrupted after native completion while
waiting for a second desktop lease for inventory; retain that incomplete wrapper separately.
The CLI inventory path exits before UI startup, so inventory queries must not acquire the
desktop lease. The actual-function regression requires normal native admission true and
inventory admission false, plus exact filter/list arguments and both contained cleanups;
the original scalar admission control fails and the corrected five scheduler/inventory
controls all pass. Actual tests retain their existing warnings, lease and restoration.

The reviewed execution digest after the inventory admission correction is
`ee8d355646a112b6de1fba8e5d06b67a2e6dccdfc418c5d4e488d1c69a8995e8`.
Native case execution, suite/filter contracts and the two enabled PR profiles are unchanged;
the modified runner and its explicit admission controls require fresh matching CI evidence.

### Rollback repetition passes; hotkey key-up ordering

Source `5cd2fa704` passes its full x64 Debug Embedded build in 7:05, receipt
`3346a3f6913a55352c6983a980b6fde2ba43b586c29c3f341ceb67189eaf81b7`, and
its full x64 ASan Debug Embedded build in 7:00, receipt
`17c476946019b095a4167c22881574aa522ebc9684f1c350231113291171a91d`.
Supported Embedded compiler debug information avoids a shared compiler-PDB RPC failure;
final linker PDBs and ASan instrumentation remain enabled. No PDB service is terminated.

The corrected ASan rollback run `20261029T002919Z-75984-99949c75645748b09c0025e4a1391519`
passes all ten invocations with zero failed/skipped/retries and canonical wrapper exit zero.
Native duration is 15,897 ms. All ten unused owned GUID children report successful POSIX
rollback and immediate enumeration of exactly `record.json`. Keep the complete
[source/build-bound receipt and trace](../../TestRuns/4cb089111a23/Commands/2026-10-29_004026_HostedCiRepair-RollbackAsanTen/evidence.json).
This is focused local acceptance; current hosted qualification remains required.

Retained hosted source `16d07ae35` diagnostics also enumerate `oversized.json` beside the
original record, with successful POSIX deletion. Its ARM64 isolated Commands retries still
refuse foreground activation after receiving the real hotkey (`sent=2`, no held key).
The helper queues key-down and key-up together before dispatch. Windows routes ordinary
keyboard messages to the foreground thread and restricts foreground activation using last
input ownership ([keyboard input](https://learn.microsoft.com/en-us/windows/win32/inputdev/about-keyboard-input),
[foreground activation](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow)).
An early key-up can therefore return input to the previous foreground process before the
hotkey handler attempts activation; this is the candidate explanation for those refusals,
not a claim of completed ARM64 acceptance.

Add a key-state oracle to the existing independent hotkey dispatch witness before changing
the helper. Its original-helper control
`20261029T005019Z-52116-167656286d44465996a9b126bd26471d` fails exactly the new assertion:
both real dispatches see `keyDownAtDispatch=false`, while foreground acquisition, exact hook
retirement, key release, repeated registration cleanup and refusal checks remain successful.
The full Debug build receipt is `90ac8bed6ebfcb26a301bd6d97b53eccbfc5b0b3536e2e5fa4263455b1a4141e`;
the retained test-only patch identifies its uncommitted test input precisely. Keep the
[failed control](../../TestRuns/4cb089111a23/Commands/2026-10-29_005019_HostedCiRepair-HotkeyReleaseOrdering/evidence.json).

Send key-down alone, dispatch through the existing owned handler, then release after the
activation attempt. Keep scoped release on every failure, require successful key-up injection,
and verify actual foreground after release. Do not attach input queues, alter foreground-lock
policy, relax the 500 ms deadline or change the original ownership/refusal/restoration gates.
The unchanged stronger witness and fresh native/hosted qualification must pass before merge.

### Captured console isolation and duplicate writers

The local remainder at `7a3a70576` passes all 29 build-toolchain checks, then stops
Compare at 226 passed, one duplicate-instance failure and 56 fail-fast omissions.
The duplicate child exits zero; retain that failure. The original helper subsequently
passes the same case in isolation, with second exit 1, the primary still alive and
the expected already-running output. That retry is diagnostic, not a causal repair.

The canonical captured-child helper omits the existing required `CREATE_NO_WINDOW`
flag. Its independent console-state control at the same public HEAD plus the retained
test-only patch fails: child exit zero, console present, zero passed/one failed/no skips.
Keep the [original helper control](../../TestRuns/4cb089111a23/Compare/2026-10-29_065644_HostedCiRepair-ConsoleNegative/evidence.json).
Add the flag without changing restricted handles, suspended Job assignment, independent
bounded drains, exit-code handling, descendant containment, cancellation or deadlines.

The corrected uncommitted native candidate passes ten repetitions of each existing
runner, duplicate-instance and resource-scoped writer case in both profiles:

- [x64 Debug](../../TestRuns/4cb089111a23/Compare/2026-10-29_071025_HostedCiRepair-ConsoleDebugTen/evidence.json):
  30 passed, zero failed/skipped/retries, wrapper exit zero, native 26,319 ms.
- [x64 ASan Debug](../../TestRuns/4cb089111a23/Compare/2026-10-29_071701_HostedCiRepair-ConsoleAsanTen/evidence.json):
  30 passed, zero failed/skipped/retries, wrapper exit zero, native 28,518 ms.

Both full builds pass with Embedded compiler debug information and final PDBs;
ASan stays enabled. All twenty console observations have no console, and all twenty
duplicate probes return exit 1 while the primary remains alive. Raw build receipts,
native input patches, full native/wrapper results and traces retain the exact dirty
source identities. These focused results prove the console-isolation correction;
they do not alone prove the precise cause of the earlier broad exit-zero failure.
Fresh broad and hosted/native ARM64 acceptance remains required.

The existing tooling source guard now requires all three launch flags and the actual
console probe. Review of that execution-closure change yields PR digest
`c12ab6b16cac10b9493ac6d1c70b4f9d8b49a40ef66f81e6ee6daab3bdf12e20`;
native suite membership, limits and the two enabled PR profiles are unchanged.

The passing ten-repeat Debug and ASan held-hotkey packets at source `7a3a70576` are
also retained beside the earlier red ordering control. Hosted PR checks at that
source pass x64/ARM64 Debug, while longer runtime qualification remains red:
Viewer harness/blocked-close, Release FileOps disclosure/quiescence/independent-volume
and ARM64 foreground/mounted-volume cases, plus shuffle viewer/Terminal deferred
creation failures and its retained timeout. Keep the draft open and these gates
explicit. The focused ViewerText diagnostic was interrupted before launching any
child while waiting for the shared desktop lease; it supplies no runtime acceptance.

## Follow-up runtime controls on 29 October

The published search-console fix is `ed8f6683b`; PR run `40836790471` passes x64
and native ARM64 Debug, formatting and migration guards. Main remains `8ea956a7e`.
This is not green-main acceptance.

The existing Terminal roundtrip, strengthened with retained-owner, placeholder-kind
and ordinary-create recovery controls, fails on the original manager in run
`20261029T075127Z-92084-509cec17ae40438e98b923e31d013e28`: kind is lost and
ordinary creation remains busy after 5,010 ms. The manager candidate passes ten
repetitions with zero failures/skips in `20261029T082549Z-64564-753e0db2692341aeaffabccdf77946cd`
(native 5,616 ms; receipt `11a2bec9122b3cebcd87bf6781ce9cc0aa9ce6334ecb7e4b12423c814086d275`).

A diagnostic-only four-worker provider build reproduces independent-volume starvation
in `20261029T083234Z-72968-860900ab169e4aedb613d98f265324aa`: both operations
copy all 16,777,216 bytes, but overlap is zero; isolated times are 1,188,000 and
1,125,000 us, concurrent time 2,234,000 us, gate 1,603,800 us. The candidate adds
one bounded reserve, caps each job below the actual created pool size and expands
stream storage for the additional distinct ID. Restore hardware discovery before
publication; the literal four-worker control is never a product setting.

An ad hoc Curl launcher incorrectly passed viewer arguments and launched a contained
ordinary app instance, PID 68260. It executed no native test. The owning operation
was interrupted, all its descendants exited, and user instance PID 56680 remained
untouched. Keep its invalid/interrupted evidence, never count it as coverage.
A full Debug Rebuild then passed (receipt
`c483bb03ef1771837740348048b09cf9a6a79c6890a5e6b3dad2f0a9a1ec0072`).
The canonical original-socket Curl control in
`20261029T090345Z-37156-66191dd46f3449c9a6fbefef55a774e6` passes all 142 checks,
with no case failure/skip and native 186,500 ms. Its 4,096-file CopyItem controls take
106,752,348 us at concurrency one and 57,208,650 us at four. The ARM64 baseline
instead reaches the unchanged 180-second deadline; reserve-name and promote probes
dominate. The TCP_NODELAY candidate in run
`20261029T095411Z-64280-a4da63f7a8884706a9cacc61930a4ba1` also passes all 142
checks, but takes 252,484 ms; the two wide CopyItem controls take 143,169,433 and
73,864,257 us. Concurrent tooling CPU load prevents a causal regression claim.
Both traces instead show 8,195 complete payload builds: the original spends
44,049,024 us constructing their rows and the candidate 56,449,874 us. Remove the
unsupported socket change. Serialize the fixed grammar directly with bounded
decimal sizes, preserving every row, operation and deadline. Keep the existing
literal-byte/4,097-file controls and add a maximum-size literal oracle to Entry
lookup (298 to 299 checks); Copy traversal stays at 142. Fresh local and native
ARM64 measurements remain required.

Prepared viewer repairs preserve complete child output through restricted inherited
handles and canonical bounded concurrent drains. The contained timeout case now also
requires exact Unicode stdout/stderr markers and an actual no-console child. It
separately checks read-only exact foreground/focus observation across thread/process
boundaries; restoration writes retain the original stricter ownership predicates.
Hosted ARM64 refusals show the same foreground and input timestamp with focus thread
10668 under foreground thread 1560. The Web combo fixture must invoke production
shutdown and reach CanUnloadNow before dropping its module; the original hosted child
reported all assertions passed and then exited 0xC0000005. Optional-backend VLC keyboard
coverage must settle loading and present a real laid-out HUD without assigning focus.

VHDX operation arguments derive only strict child aliases of the freshly revalidated
owned volume. Seeding/inspection/cleanup retain the sandbox mount spelling, and all
foreign-volume, mount, alias and escape refusals plus product ancestry guards remain.
These repairs remain candidates until focused, broader and hosted qualification passes.

The four-transfer-plus-reserve diagnostic passes ten unchanged R4A19 repetitions
in `20261029T093953Z-74488-dbab28ce4d6e4d628a0acef4314b75c0`: 30 passed,
zero failed/skipped, native 40,250 ms, receipt
`b60f62efe5c28581bd6cd91dc906ad3f4ec9b4e6541a5bc9a2116e447f3308ec`.
Restore hardware discovery for the production candidate. A full ASan Debug Rebuild
on that diagnostic source also passes in 5:59, receipt
`7c0276dca43f6ea6ce2f5a97d8e4c177f69dc8cc2e96949a60a9b311724f7aa8`.
Neither build substitutes for runtime acceptance of the final source.

Viewer diagnostic run `20261029T103028Z-58920-9f320e0ecea34028a7fe8373a0c2fbd7`
passes containment, Web shutdown/quiet unload and the real VLC HUD focus journey.
The Text child also passes every assertion, including blocked close in 45,359 us
under its unchanged 500,000 us gate. Its parent correctly refuses restoration:
foreground changes from Slack PID 41372 to unrelated Claude PID 58440 and the input
timestamp changes during the run. Retain the aggregate failure and rerun without
weakening input ownership. This passing child is not a causal repair of the older
hosted close failure.

The governed source/spec/scoped/archive controls pass all 384 checks without skips
in `20261029T101740Z-95300-91fa02cf404e497e8506f38d3eb4eed8` (453,929 ms).
Those controls precede the serializer and Commands follow-ups and must not be
reported as qualification of newer source. Documentation must pin an ordinary
unopened Terminal before export lookup; a deferred catalog entry has no HMODULE.
The warm-quiescence helper must drain and resample at its final quiet decision,
with a real late-queued paint control retaining all five original timer trials.
Find-sort diagnostics must format the actual post-wait snapshot; argument
evaluation order previously formatted its zero-initialized value before waiting.

### Source-bound follow-up and independent repeat controls

Candidate v4 at public parent `ed8f6683b` has source snapshot
`8133b2cc775785150341c13d41f01b017e708a612a8edd7684c4ca8d143e25b0`.
Its Debug, ASan Debug and test-enabled Release receipts are respectively
`c94752f45d45f7e17fb792fb7d7a663b3bd7e1dc97c25726234780774031c6cd`,
`17d871317ef76260d49114b7d70fc27e8a7dc86786f4e7839a64be008a13e61a`
and `792bf7adfe7e39d74635195ba6f8ef523fa3a4fadb408cead288f6f7687a8768`.
The initial Release build omitted test-enabled Full overrides; all four attempts
correctly refused attestation before execution. Preserve them as blocked attempts,
not runtime acceptance. The corrected Release build uses the canonical Full
overrides. Embedded compiler debug information and final PDBs remain enabled.

- Debug run `20261029T105813Z-72596-9f3de0032d574f2c9537bb974136cc75`
  passes nine exact Commands cases three times: 27 passed, no failures/skips,
  native 37,598 ms. All three late-paint controls actually queue, drain and render
  once before the quiet return. Every original disclosure route passes.
- Static gallery runs `20261029T111249Z-72596-c0b31c15ec0c4f4ab1c0de4ff9a28ba0`
  (Debug) and `20261029T112023Z-9060-989b7e4a90eb4b02b6f25c4162661341`
  (Release) pass with no skips and no activation. Inspect the actual synthetic
  Terminal Find and File Operations preferences captures; both Unicode separators
  render. Static frames do not qualify keyboard focus.
- Fresh ASan search/runner run
  `20261029T111243Z-58980-ff027f7f6b6346efbaff6224d82bd79a` passes all
  30 repeated console, duplicate-writer and resource-scoped ownership cases.
- Entry lookup runs `20261029T111547Z-72868-77eec78043c940a3bf3cc3a35ed8eb49`
  (Debug) and `20261029T112343Z-9060-08af2b90f7a848b1a4ada6b2bc04d12b`
  (Release) pass all 299 internal checks plus Setup/Cleanup, with no skips.
  Debug's 64 complete 4,097-file renders preserve all 14,159,232 bytes; this
  unpaired direct measurement is not a performance non-regression claim.
- Release live-progress run
  `20261029T112207Z-9060-c5c846dceeaa43b3bfe93deab03bedae` and ASan run
  `20261029T121138Z-94316-3f437f72b79d4c51a70353bb7fc27bf8` each pass
  all three repetitions, including Setup/Cleanup (nine cases each). The original
  both-roots/content/progress/close/row assertions remain.
- ASan remediation run
  `20261029T120930Z-94316-dbdfa9026a3e4a13a95b73f8f3bde4c7` passes
  its original checks. It did not invoke the F7 owned-volume operation-path
  helper; do not claim that native acceptance from this v4 run. Add the same
  pure controls to the normal remediation entry and qualify them afresh.
- The newer governed source/spec/scoped/archive run
  `20261029T105320Z-55260-856680000623443db3ed22dd29fa4391` passes all
  384 checks, zero failures/skips, in 416,759 ms. It predates the repeat-fixture
  repairs and does not qualify those newer inputs.

Original repeat control
`20261029T111446Z-58980-9bf5189606b64503a834b8f9594484a7` reproduces
six ASan failures (six passed, six failed, no skips; native 52,617 ms): both
settings fixtures and the thumbnail fixture pass once, then fail on repetitions
two and three. An independent fixture deletes/reseeds its target while its old
coordinator lineage remains; use the existing test-only quiet/reset API only
before setup. Within-case external/stale guards must remain untouched. The
thumbnail command is a legitimate no-op if the prior invocation left 96 DIP;
establish and assert 48 DIP, retaining the original cancellation gate.

The isolated theme case passes those three original invocations. Its old loop
nevertheless has a net 120 ring steps, so its System assertion depends on catalog
length. Strengthen it with a real 16-theme catalog before changing the loop.
That control at snapshot
`ca9b88a0f173b382588f7792d98a3dbc9235195b1a97950a872020f2d7c867a3`,
receipt `6469c207e5569054b1bc45abb3f0c573a3b5ab32f3e337f89fa4a9f0a4414103`,
fails exactly the three System assertions in
`20261029T122930Z-72524-0a01c8e4738540149027a3566418ec0c`.
The repaired settings and thumbnail fixtures pass all nine corresponding
invocations; native duration is 78,863 ms overall. Retain this deliberate
negative control. The balanced candidate keeps every original 240-command,
resource, animation and System assertion and adds exact 120/120 direction
counts. Fresh Debug/Release/ASan and hosted/native ARM64 acceptance is pending.

Only this repair's queued C0-copy controllers, PIDs 72868 and 9060, were stopped
before native launch so the reproduced repairs could proceed. Both owners were
verified dead; interrupted runs `20261029T115103Z-72868-7687f358223e473ab61c673bca3ef96f`
and `20261029T115117Z-9060-da04384f12634920a24278c5df9b1e97` are preserved.
No Copy traversal acceptance comes from either attempt. Full-solution Rebuilds
precede renewed qualification. Independently launched user PID 56680 and the
other chats' current entries remain untouched. Queue coordination with both
Audit MSVC and Integrate libghostty is explicitly authorized; focus-taking work
retains its agreed timing, warning, shared lease and verified restoration.

Fifteen bounded follow-up packets retain the failed controls and successful
focused results under `Specs/TestRuns/4cb089111a23`. Raw metrics exceeding the
archive cap remain in the completed-run archive with their exact byte digest;
the R4A19 packet retains every claim-bearing row in original order and its full
input digest/count/line provenance. Native scope, skips and failures remain
explicit. No limit is raised or baseline silently replaced. The reviewed PR
execution digest remains `c12ab6b16cac10b9493ac6d1c70b4f9d8b49a40ef66f81e6ee6daab3bdf12e20`.

### Integrated candidate focused qualification, 29 October

Committed runtime candidate `a319e2f398788f35c58968b6dad18fa7632a933a` integrates main
`b4ff85bead78c1caf2a1e9e7fb80545ca00831de`. Full-solution test-enabled x64 Rebuilds pass
in Debug, ASan Debug and Release, with the same one pre-existing C4464 warning and no
errors. Their exact receipts are respectively
`853ff02833077cdc5dfa273f730dbfb96a80ec327036ba78c109603530886609`,
`03233ac0698015cc18e4c2277b341656ece695139c877795139753ee9da0f51c` and
`c061b0578727d0d02aba14816e36d93dd316da6ea6a0abe85ddc7ba1ca726b2f`.
All focused counts below pass with zero skips and no automatic retries. FileOps counts
include Setup/Cleanup; the internal Curl assertion counts remain separate.

| Focused scope | Debug | ASan Debug | Release |
|---|---:|---:|---:|
| Owned-volume remediation | [3](../../TestRuns/4cb089111a23/FileOps/2026-10-29_131153_HostedCiRepair-FinalDebugOwnedPaths/README.md) | [3](../../TestRuns/4cb089111a23/FileOps/2026-10-29_131736_HostedCiRepair-FinalAsanOwnedPaths/README.md) | [3](../../TestRuns/4cb089111a23/FileOps/2026-10-29_132543_HostedCiRepair-FinalReleaseOwnedPaths/README.md) |
| Captured service and writer ownership (10 repeats) | [30](../../TestRuns/4cb089111a23/Compare/2026-10-29_131333_HostedCiRepair-FinalDebugSearchTen/README.md) | [30](../../TestRuns/4cb089111a23/Compare/2026-10-29_132000_HostedCiRepair-FinalAsanSearchTen/README.md) | [30](../../TestRuns/4cb089111a23/Compare/2026-10-29_132717_HostedCiRepair-FinalReleaseSearchTen/README.md) |
| Settings, thumbnails and theme catalog (3 repeats) | [12](../../TestRuns/4cb089111a23/Commands/2026-10-29_131449_HostedCiRepair-FinalDebugRepeatThree/README.md) | [12](../../TestRuns/4cb089111a23/Commands/2026-10-29_132152_HostedCiRepair-FinalAsanRepeatThree/README.md) | [12](../../TestRuns/4cb089111a23/Commands/2026-10-29_132836_HostedCiRepair-FinalReleaseRepeatThree/README.md) |
| Terminal, Find, disclosure, ViewerSpace and quiet-boundary cluster (3 repeats) | [27](../../TestRuns/4cb089111a23/Commands/2026-10-29_131717_HostedCiRepair-FinalDebugClusterThree/README.md) | [27](../../TestRuns/4cb089111a23/Commands/2026-10-29_133035_HostedCiRepair-FinalAsanClusterThree/README.md) | [27](../../TestRuns/4cb089111a23/Commands/2026-10-29_133212_HostedCiRepair-FinalReleaseClusterThree/README.md) |
| Curl entry lookup (299 internal checks plus Setup/Cleanup) | [3](../../TestRuns/4cb089111a23/FileOps/2026-10-29_132507_HostedCiRepair-FinalDebugCurlEntry/README.md) | [3](../../TestRuns/4cb089111a23/FileOps/2026-10-29_133311_HostedCiRepair-FinalAsanCurlEntry/README.md) | [3](../../TestRuns/4cb089111a23/FileOps/2026-10-29_133617_HostedCiRepair-FinalReleaseCurlEntry/README.md) |
| Curl traversal (142 internal checks plus Setup/Cleanup) | [3](../../TestRuns/4cb089111a23/FileOps/2026-10-29_133050_HostedCiRepair-FinalDebugCurlCopy/README.md) | [3](../../TestRuns/4cb089111a23/FileOps/2026-10-29_133644_HostedCiRepair-FinalAsanCurlCopy/README.md) | [3](../../TestRuns/4cb089111a23/FileOps/2026-10-29_133723_HostedCiRepair-FinalReleaseCurlCopy/README.md) |
| Multiple-folder live progress (3 repeats including Setup/Cleanup) | [9](../../TestRuns/4cb089111a23/FileOps/2026-10-29_133534_HostedCiRepair-FinalDebugMultipleFoldersThree/README.md) | [9](../../TestRuns/4cb089111a23/FileOps/2026-10-29_134303_HostedCiRepair-FinalAsanMultipleFoldersThree/README.md) | [9](../../TestRuns/4cb089111a23/FileOps/2026-10-29_134350_HostedCiRepair-FinalReleaseMultipleFoldersThree/README.md) |
The original repeat-control failures and the independent non-divisor theme catalog
failure remain intact in separate packets. Every original 240-command, animation,
resource and System gate remains. Ordinary remediation now really invokes the owned-volume
operation-path controls, as its retained native trace confirms. C-root functional Curl
runs preserve all eight traversal variants; they are not a matched E-root performance
pair, and establish no causal throughput or non-regression claim.

Focused viewer batches pass five supported cases in Debug and ASan Debug and four in
Release. These include containment/read-only identity controls, final-applied image
orientation, Web quiet teardown and actual VLC HUD keyboard routing. The _DEBUG-only
Text async-open/UTF-8/hex case passes with its original 500 ms close gate in Debug and
ASan. A driver mistakenly selected that case for Release; the native registry correctly
rejected it before execution. Preserve that original exit-1 attempt separately from the
fresh corrected four-case Release pass. No compiled or registered test was removed.

The governed five-file run `20261029T130632Z-99464-e9f865a30733457986517103d1b64622`
passes 383 checks and fails one unchanged source guard at its 900-character loop window.
The added explanatory comment enlarged that window. Move only its three lines before
the loop; the original guard matches the 793-character body afterward. No executable
C++ token, assertion, test count, deadline, or test-policy limit changes. Retain the
failed run and qualify the adjusted source through the final tooling gates.

Twenty-five bounded packets retain exact receipts, source patch, raw results and assertion
traces: twenty-one successful in-product batches, three supported viewer passes, and the
invalid Release selection. Native input patch SHA-256 is
`2722da4cceb58f66d81d1388601a91521a408da26aec6364e6de40c84981b91b`.
The handoff after this runtime candidate changes only explanatory comment placement and
evidence/documentation; it does not turn these focused results into Full qualification.
PrePush selected obligations and fresh hosted x64/ARM64 PR/main workflow acceptance
remain explicit gates. Keep this plan ACTIVE until the main workflows actually pass.
