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
