# Hosted CI repair — 2026-10-28

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
