# I50 Ghostty 1.4.0 terminal runtime integration

> Non-normative execution record. The durable runtime contract belongs to
> `Specs/Terminal/Terminal_EmbeddedPlugin.md`; testing and tooling contracts retain
> their existing owners.

## Scope and source identity

Integrate libghostty-vt from the official Ghostty `v1.4.0` tag, commit
`83fb95624b0439ee660c2679b43b38c8b63769cd`, tree
`4dbd485a3b72c6ced65ad7ea447644a5967e382e`, into RedSalamander.
Ghostty and libghostty have separate release cycles; this is not a claim that
upstream has released a separately versioned libghostty 1.4 package.

The starting product lock selects `9c3ec931d64561a8407dde7ac984ce156ae91539`
(`1.3.2-dev`), Zig 0.16.0, and three dependencies. The explicit Candidate
observation confirms 933 commits ahead and no divergence. Work starts at
RedSalamander `8ea956a7e2e271368082a953f28efd2a58e1013f` on
`ej/libghostty-1.4-integration`.

I15 continues to own the older remediation qualification. This plan does not
reinterpret its historical runs as current evidence or close that plan.

## Progress

- [x] Create and attach an isolated clean worktree from refreshed `origin/main`.
- [x] Resolve the official annotated tag and collect an exact Candidate status.
- [x] Freeze authenticated source/archive/ABI/security inventory and descriptor.
      Manual source dispositions remain part of the overlay/API review.
- [x] Rederive the Windows/private overlay against the exact candidate; review
      every retained concern, enum allocation, and changed upstream contract.
- [x] Lock Zig 0.17.0 and the complete transitive runtime dependency/notices
      closure; build and inspect isolated x64 and ARM64 candidate outputs.
- [x] Adapt host integration and meaningful regression coverage to public API
      changes while preserving clipboard, placement, teardown, and pairing rules.
- [ ] Execute deterministic upstream/runtime/plugin coverage and collect a
      same-machine paired 200-sample Release VT corpus before admission.
- [ ] Complete required native ARM64 candidate and admitted-product gates;
      cross-compilation alone does not satisfy native runtime qualification.
- [ ] Admit the reviewed lock/overlay, update authoritative specs/documentation,
      and complete applicable affected and PrePush obligations.
- [ ] Fold/check the changelog, reconcile remaining gates, and move this record
      to Done only when the required outcome is qualified.

## Protected scenarios and acceptance

Reuse `Terminal.Tests.VtUpgrade.cpp` and the existing VT corpus and comparison
tooling for VT ingest, render snapshots, resize/reflow, Unicode, input encoding,
clipboard rejection, and bounded Kitty image/placement retention. Retain raw
samples and exact source, artifact, profile, corpus, control, and machine
identities. Candidate p95 must remain within baseline x1.10 under the owning
performance contract. Keep upstream test results distinct from production
plugin execution and portable packaging.

Use the existing marked `E:\RedSalamander.Perf` for test scratch and evidence;
its ownership marker was verified before use. Build inputs and candidate sources
remain under this worktree's `.build/TerminalEngineUpgrade/<commit>`.
No focus-taking run is authorized by this plan; such gates remain pending until
the user agrees to the execution time.

## Evidence and decisions

- Official release notes: <https://ghostty.org/docs/install/release-notes/1-4-0>.
- Candidate status: `.build/TerminalEngineStatus/ghostty-1-4/status.json`,
  `update-observed`, exact source/tree and ancestor relationship confirmed.
- Initial source scope: 145 files, 34,521 insertions and 4,057 deletions across
  the governed lib-vt review paths. The candidate requires Zig 0.17.0 and changes
  the public C API; changing only a version string is insufficient.
- The canonical x64 baseline runtime rebuilt successfully: 2,141,184 bytes,
  raw SHA-256 `dc048bb106385058fd145f5f6fe74e0cb1811c418308fe0a13e1237869b387eb`.
- The authenticated candidate freeze passed: both archives have SHA-256
  `7759929eadaddf602d6fd5ad6b3403d854b91b7083975cd984bb525bc8275187`;
  261 lib-vt commits are inventoried and all current loader exports are declared.
- The first freeze failed because the old descriptor concatenated the incumbent
  lock's three dependencies into scalar hashes. Descriptor schema 2 and overlay
  rebinding now retain every dependency separately. Historical finding templates
  are restricted to ranges containing all their supporting commits.
- Focused Pester regression run `ghostty140-tooling-20261028b` passed 11/11.
  Its predecessor remains under the marked sandbox; it exposed an incomplete
  synthetic enum fixture, which was corrected before the passing repeat.
- The rederived overlay uses options 48/49 after upstream 0-47 and Kitty data
  values 3-6. Its seven files retain all five reviewed concerns (372 insertions,
  27 deletions), SHA-256
  `995cc2b9d39907dfb4b6dac179e844b3f7a7c4279a81df502a693ff7834c8a1d`.
  Placement-map reservation precedes garbage reaping, preserving state/pins
  even when allocation fails.
- The Zig 0.17 closure includes Aro (translate-c transitive parser), Pixels
  (Wuffs build fixtures), translate-c, uucode, and Wuffs. Eight notices are
  staged centrally for both platforms. Verified packages are passed through
  `--system` so missing dependencies cannot fall back to network fetching.
- Governed candidate builds passed for x64 (1,973,248 bytes, raw SHA-256
  `8299405198f67e19015c0d9ebc806517a277904b5dc834b1211ca27001baa10f`)
  and ARM64 (1,732,608 bytes,
  `fb1ba734aa742b59b72451c725f6c43edd3d02a0b414969e66dd330884a08a21`).
  Both have 207 exports and seven locked import modules. Cross-builds establish
  PE shape and closure, not native ARM64 execution.
- Native x64 upstream C-ABI tests passed 3,720 with 36 skipped and zero failures
  in `ghostty140-upstream-x64-20261028b`. The first run's 26 missing-fixture
  failures remain in `ghostty140-upstream-x64-20261028a`; the repeat used the
  identical binary with upstream snapshot fixtures under sandbox scratch.
- Baseline test-enabled x64 Release Terminal builds passed without warnings or
  errors. The 200-sample VT corpus passed in
  `20261028_174540_ghostty140-product`; baseline p95 was 2,198 us VT write,
  11,333 us snapshot and 29,944 us hosted render. Exact pre-upgrade lock, patch,
  runtime and receipt bytes are preserved alongside that run.
- The proposed canonical pin/overlay and centrally staged notices now select
  1.4.0 on this branch. Product execution, paired performance, affected checks,
  and both native ARM64 gates remain open; this is not release qualification.
- Fresh bootstrap rejected initially flattened remote-cache representations.
  Local Zig 0.17 `fetch` retains the source transport wrapper in its canonical
  archive. The final lock records those exact archive hashes and nested roots
  for all five dependencies. Independent comparison verified identical file
  bytes, modes and sizes for 98 Aro, 10 Pixels, 16 translate-c, 42 uucode and
  14 Wuffs files. Original archives, corrected archives and the comparison
  record remain in `ghostty140-pixels-bootstrap-20261028a`. The corrected
  offline Product runtime build passed and reproduced the x64 candidate DLL
  byte-for-byte. No hash check was relaxed.
- The missing-package experiment `ghostty140-missing-package-20261028a`
  rejected absent uucode/translate-c immediately in `--system` mode.
  `ghostty140-missing-package-20261028b` repeats that check with an explicit
  integer exit-code record (exit 1); the first diagnostic record had omitted
  that value and remains preserved.
- The native workflow now supports hash-bound checked-out inputs, Zig 0.17,
  the shared closed-package helper, and sandboxed fixtures/logs/dumps. Its x64
  harness parsed and the 187-type ABI schema and 3,720 C-ABI tests passed. The
  combined local harness hit its 20-minute timeout during the separate Zig
  suite; its output is retained in `ghostty140-workflow-harness-20261028a`.
  A direct run of the already compiled Zig binary continues that unresolved
  coverage without rebuilding or rerunning the passed C-ABI suite. This does
  not replace native ARM64 execution.

- The standalone Zig binary printed 3,233 passed, 36 skipped and zero failed.
  Its ad-hoc local exit-code record is null, so that output is diagnostic and
  does not establish a clean process exit. Native workflow acceptance remains
  required; the C-ABI standalone run separately returned exit 0.
- Final focused tooling run `ghostty140-tooling-final-20261028a` passed 172
  tests, zero failed, with one historical archive-inventory check skipped
  because its contract commit is absent from this squashed checkout.
- The corrected x64 Release Terminal build passed with zero warnings/errors:
  receipt `45251b11ac7ce1443b3c0c28a4d3f5011dfdfdfafb895ed5abdcfad752002349`.
  The actual loaded plugin passed all 162 deterministic selftests, localization,
  sized-record and unload contracts (`ghostty140-terminal-release-20261028a`).
  The final ARM64 Product offline cross-build also passed with the same runtime
  bytes as its candidate build.
- The 200-sample candidate corpus `20261028_190224_ghostty140-candidate`
  passed behavior/unload checks and the pair validator with the explicit
  reviewed Zig 0.16.0 to 0.17.0 transition. p95 (baseline to candidate): VT write
  2,198 to 1,108 us; snapshot 11,333 to 3,776 us; hosted render 29,944 to
  30,135 us. Every measurement stays below baseline x1.10. This measures the
  complete compiler/runtime upgrade and does not isolate either one's effect.
  The two curated runs under `Specs/TestRuns/4cb089111a23/Terminal/` retain
  original samples, traces, exact runtime/lock/overlay and build-receipt bytes.
  These are targeted dirty-tree receipts at starting HEAD `8ea956a7e2e2`, not
  final immutable-commit Full acceptance.

## Source review dispositions

- Required loader prototypes remain present; no host-required API needs a
  signature adapter. New search, render-image and animation APIs are not wired
  to new host features by this upgrade.
- The new Windows CRT sentinel exports and runtime import are explained by
  upstream `src/lib/windows_dll.zig`: DLL startup runs the C++ constructors.
  Actual DLL-based Unicode/plugin tests remain necessary because static unit
  tests cannot establish DLL startup correctness.
- Inspected the upstream cursor-page replacement fix (`65ff54c67`), orphaned
  wide-character resize fix (`e6db5b633`), snapshot exact-length copy fix
  (`8f78fe4c1`), and reference-count accounting change (`5f5495826`). Their
  deterministic upstream coverage supplements the existing host corpus. This
  is a bounded review, not an exhaustive security certification of 34,521
  inserted lines.
- The private placement path validates capacity/count/image-counter bounds
  before mutation, reserves before garbage reaping, preserves replacement pin
  ownership, and keeps screen-limit changes atomic. Clipboard read remains
  null, decoded writes remain bounded at 256 KiB, and Kitty OSC 5522 writes
  remain rejected before buffering. File/shared/temp image media remain
  disabled by the host.
