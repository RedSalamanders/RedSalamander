# I50 Hosted CI repair — 2026-10-28

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
machine-readable WIP table row; the repair is now indexed as unused rank I50, without an exemption.
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
