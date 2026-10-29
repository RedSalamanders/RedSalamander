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
- [x] Execute deterministic x64 upstream/runtime/plugin coverage and collect a
      same-machine paired 200-sample Release VT corpus before admission.
- [x] Complete the native ARM64 candidate gate, including original canonical
      report publication and independent artifact/identity validation.
- [x] Complete the admitted-product native ARM64 gate, including Debug Terminal
      contracts, representative Commands, production Release, portable smoke
      and original receipt/artifact publication.
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
The user authorized the four pending interactive qualification suites on
2026-10-28: Commands, FileOperations, ViewerPETests.Interactive and
ViewerSqliteTests.Interactive. Execute them through the canonical desktop lease
and warning; acceptance remains pending their actual results.

## 2026-10-29 14:14 UTC qualification checkpoint

- Application and build sources remain unchanged between `3f2b0f1ce` and
  `f72618d62`; the current x64 Debug test-enabled Embedded build uses receipt
  `224e9a7766a213ef9d91f34dcea8a5567862f12f96a6e4bdb648cce53e723c67`.
  Its rebuild passed with one unchanged ViewerWeb C4464 warning and no errors.
- Current PR CI `40847813175` passed on x64 and native ARM64 at `f72618d62`.
  Each platform promoted all twelve entries: 1,076 passed, zero failed and
  two historical-evidence skips, including all 162 loaded Terminal selftests.
  Both original validation artifacts, complete logs and metadata were downloaded
  and independently checked in
  `20261029T103355Z-64404-57469813dc7e4797abc86f97ad52db4c` (x64) and
  `20261029T103418Z-79536-67e027f8b42547509ad68f777ee908f9` (ARM64).
- Fresh local run `20261029T100906Z-94716-97da1f410e0041d688bc0b29d42913d7`
  is archived as FAILED: 1,542 passed, three failed and 52 declared skips.
  Build-toolchain passed 29 checks; Compare passed 260/0/21, including the
  previously timed-out cold-start case in 38,488 ms; File Operations passed
  297/0/25, including `R4A19_DiscoveryIndependentVolumes` in 3,734 ms; interactive
  SQLite passed. The alternate-volume control reused the existing marked
  C: sandbox with the explicit opt-in. The original failed aggregate is retained.
- Commands recorded 955 passed, two failed and six declared skips. The failures
  are `cmd_pane_embedded_terminal_plugin_lifecycle` and
  `cmd_pane_embedded_terminal_shell_exit_closes_only_its_tab`. At the first
  five-second readiness deadline, the listener reported stage 2, S_OK and no
  authenticated client. Exact failed logs, source, receipt and terminal binaries
  are preserved in `20261029T115415Z-85364-2b11848ba7f04ed98bd3d278c81b7ad4`.
  This diagnostic does not establish why the shell was late.
- The isolated four-case embedded-terminal repeat passed 4/0/0 at normal timeout
  scale on the unchanged build, native run
  `20261029T120142Z-72404-39188a0963824958b852d8aa752eceb4`. Its aggregate uses
  the v1 schema; the wrapper's null v2 count fields do not replace those native
  results. The ordered Navigation family then passed 168/0/0 across three
  repetitions in `20261029T134326Z-69680-3b409fc1b3994df288a7f88a28518103`.
  Neither repeat explains or relabels the failed broad run. No application
  code, readiness deadline or assertion was changed.
- Interactive PE's VLC child passed and its owned job became quiet, but the
  parent refused restoration after foreground moved from Codex to the independent
  PowerToys Mouse Without Borders helper while last-input state had changed.
  This does not prove physical user input. After the user explicitly paused
  that utility, the complete group passed all 57 child runs and all twenty
  native identity controls, with unchanged guards and successful post-run
  source/artifact attestation, in
  `20261029T132040Z-68476-09c3b0c26f634329a72cfe4c28788942`.
  That run observed no independent-owner restoration branch; the native
  counterfactual and source-wiring checks retain their separate scope.
  The user was told they could resume the utility after the desktop phase.
- Original application and ViewerPE images were byte-verified and preserved
  before any further replacement in
  `20261029T134223Z-41384-31d6afb93c4744feb82461458c74e83d`.
  Read-only observers recorded owned process metadata without command lines,
  window captions, personal profile contents or process-memory capture.
- Both raw paired VT runs were archived with all nine files preserved per run;
  seven original corpus/provenance files per run match their published copies.
  The candidate's three separately published review records remain distinct.
  Verification is in `20261029T104411Z-5356-e8d1426c6dfc4cee87c474215fb6c67d`.
- A complete Commands rerun with failure classification is now executing on
  the unchanged source and artifacts, native run
  `20261029T135851Z-95996-75f5e460bf094a7e83b6cd6ab91b9b3a`.
  Its passive observer adds owned process counters and terminal-window geometry
  to diagnose a recurrence. The original readiness failure's cause remains open.
- The shared sandbox audit is explicitly non-blocking under
  `Specs/Testing/Testing_SelfTests.md`. It records 47 remaining layout issues
  after preserving this task's two older paired runs. Earlier wording treating
  those diagnostics alone as a Full acceptance blocker was incorrect.
  This checkpoint still makes no complete repository Full or release claim.

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

- The original approved local run completed at application source 3f2b0f1ce
  and is canonically archived as
  20261029T012426Z-93224-2b01e80f5a6341bcb1d0691d944ae606:
  1,539 passed, three failed and 55 skipped. Commands and interactive SQLite
  passed. The three failures are the initial search-service warmup timeout,
  the mandatory independent-volume control without its opt-in, and the
  interactive PE parent harness's saved-desktop identity check. Its 49 disk-audit
  issues remain recorded; this failed broad run is never relabeled as passed.
- On the same unchanged application source and Debug receipt, the isolated
  warmup repeat passed in 58,067 ms (one passed, no failures/skips), run
  20261029T081219Z-63824-c15eece088984e45922715cf0039d481, wrapper evidence
  20261029T081219Z-63824-5983646682544d67af818c5c64440404. This passing repeat
  neither establishes the original timeout's cause nor repairs its verdict.
  The independent-volume repeat with the explicit opt-in passed its requested
  case in 3,625 ms (three including setup, no failures/skips), run
  20261029T082424Z-4548-2a3de1c003064d54934d0b20098e0ac0, wrapper
  20261029T082424Z-4548-e284800eb77e4028a1f692d14daca287. It reused the
  already declared, marked C:\RedSalamander.Perf without initializing a new root.
  These focused repeats are not complete Compare or File Operations acceptance.
- The PE child passed its UI assertions, exited zero and left its owned job
  quiet; the parent then refused restoration because the saved focus window's
  process/thread differed from the foreground root. Exact logs, source snapshots,
  plan and build receipt were preserved with verified hashes in
  20261029T080657Z-58752-fc166cf4667d46aab6bdd86b68cf34a5. Read-only native
  observation in 20261029T081333Z-92508-9a1d8bf2d99c46e08e005c4da4f11eea
  confirmed that both saved handles retained those identities and the focus
  window was a legitimate cross-process child beneath that root. It performed
  no desktop writes and does not prove the GUI-focus state at the earlier failure.
- The PE parent now captures foreground and focus identities independently.
  Exact automatic foreground/focus restoration can pass through observation
  alone; the write path retains its focus-owner, unchanged-input and foreground
  restrictions. The new noninteractive native counterfactual failed exactly its
  two cross-process/thread positive controls with the old restriction in
  20261029T083456Z-90376-eb59dea6c6184a2ca8d789c8078bca1c, then passed all
  twenty identity/permission checks after repair in
  20261029T084444Z-91992-d6e5f1d8aa3248519d31d5f2c676e6a9. Both runs retain
  exact source, executable hashes, project-build receipts and output; they are
  targeted dirty-tree checks at bd0c27040, not final Full qualification.
  The harness, scoped-testing and tooling-governance checks passed 232/232 in
  20261029T085613Z-45924-147137bbafcf405fb9ab1210c007c7e9. The preceding
  invocation used an unsupported Pester option and failed before test execution;
  it remains interrupted evidence 20261029T085148Z-98700-76d9e35365764193a90688c33aed2bc9.
  Full interactive PE qualification and matching CI for this repair remain open.

These are source-bound execution checkpoints. Statements about a run still
being queued or pending describe that checkpoint; later successful records do
not erase earlier failures. The live status is maintained in draft PR #148.
At the 2026-10-29 07:34 UTC checkpoint, both native ARM64 gates and both PR CI
platforms have passed. The local application run remains frozen at `3f2b0f1ce`;
File Operations coverage validation and the two viewers still await the shared
desktop lease, followed by the two focused failure checks. No complete local
PrePush or repository verdict is inferred from the native gates.

- Admitted-product ARM64 run `40821309490` passed every step at `17058e31f`:
  test-enabled Debug, all 162 loaded Terminal selftests, all 15 representative
  Commands cases (10 terminal-shortcut, four embedded-pane and one shell-route),
  tests-disabled production Release, and native portable clean-extraction smoke.
  Both artifact uploads passed. Original artifacts `12878221818` (product) and
  `12877564861` (test root), complete logs and metadata are preserved in
  `20261029T063245Z-42988-f7459e2ab88942db90e34c7ddb84d8d4`.
  Independent verification in
  `20261029T063510Z-78348-fef3861180164219ae9b84619162593d` checked both
  canonical receipt pointers against their immutable records and projection
  digests, their source/profile/test-enable identities, all three Commands
  result sets, and the packaged runtime against the Release receipt and native
  candidate. Debug receipt is
  `3056a1c82efeb432843bcd9a78a870a77953eae904a73563561beae1142e724d`;
  Release receipt is
  `321f6ed68ccab52c5cfe7423313e82476b652cb4009e79838e0a2baf48181fda`.
  `RedSalamander-7.0.1-ARM64-Portable.zip` is 25,163,597 bytes, SHA-256
  `c1b85e06387c89ccfeb2e3e116b5d7e7365822db9b2a381b54216778709c1433`.

- PR CI `40822116591` passed on x64 and native ARM64 at `72d874428`.
  Each platform passed 1,064 PowerShell checks and eleven other gate entries,
  including 162 loaded Terminal selftests, with two historical-evidence skips
  and zero blocking classifications or quarantine entries. The ARM64 original
  aggregate confirms PR scope, 1,075 passed, zero failed, two skipped and all
  twelve entries promoted. ARM64 logs and original validation artifacts are
  preserved with the product collection above; x64 logs are in
  `20261029T055405Z-91012-3d89758a38904245ba9f0dbeab999e39`.
  These matching PR gates do not turn the separately failed local application
  run into complete repository qualification.

- Native ARM64 run `40819647447` passed at `9a46463e8`, including original
  evidence publication: 3,720 C-ABI tests and 3,233 Zig tests passed, each with
  36 skips and zero failures, and the ABI manifest validated 187 types. All
  targets passed on their first execution; no code-5 retry was needed. Original
  artifact `12877308139` was downloaded and independently checked for exact
  canonical/schema validity, the native ARM64 host, candidate/overlay and
  canonical lock identity, and the admitted runtime hash. Complete job metadata,
  logs, downloaded bytes and verification are preserved in
  `20261029T062010Z-18316-64d5746c843848e3937c5d5bf0c2f02e`. The original
  report SHA-256 is
  `e71a6323145a329fbd886e6e3db6ef752c988f9a3d54972ba0a719e8d40ec758`.
  This successful publication supersedes the earlier upload failure's open gate
  while preserving that failed run and its explicit recovered-report provenance.

- Local File Operations execution at frozen `3f2b0f1ce` finished with 293
  passed, one failed and 28 skipped. Its only failure,
  `R4A19_DiscoveryIndependentVolumes`, reports
  `alternate-volume-test-root-not-authorized`: this invocation omitted
  `-AllowAlternateVolumeTestRoot`. The existing machine resource already enables
  `C:\RedSalamander.Perf`; its marker and reparse-free authorization were
  independently validated without creating or initializing a new root. Preserve
  the original result and repeat that exact control with the explicit opt-in
  after the active profile releases. The initial failure snapshot and receipt
  are preserved in `20261029T054109Z-74128-5e844c6cbfce492aa8964048f1859e9d`.
  Coverage-list validation and both already-authorized interactive viewers are
  still queued behind the shared desktop lease.

- Product run `40819043691` was cancelled during dependency installation,
  before application build or product tests, after verifying that its artifact
  list referenced obsolete `.build/BuildReceipts/`. The canonical producer
  `Get-RSBuildReceiptStore` retains pointers and immutable records beneath each
  actual output profile instead. The pre-cancellation state and interruption
  reason are preserved in
  `20261029T050127Z-81836-042b4d74e0cc4cce8ef45c04d950f577`.
  Product artifacts now retain both ARM64 profiles' receipt pointers, immutable
  histories and toolchain identities. The regression derives expected paths
  from the real producer helper; focused workflow, governance and scoped checks
  passed 61/61 in `20261029T050250Z-52904-fcdd228b4f7c466db64215025f95000e`.
  No product execution was discarded or counted as accepted by that cancellation.
  The complete cancelled-job log is retained in
  `20261029T050439Z-57212-cc6ddb60798546d6a4cb16b1ca9f9c50`.
  Replacement product run `40821309490` uses commit `17058e31f`; the separate
  native publication verification `40819647447` uses `9a46463e8` and does not
  duplicate the product run.

- Native run `40810497954` at `7152e3823` completed all three upstream targets
  on the first execution, without a code-5 retry: the ARM64 C-ABI suite passed
  3,720 with 36 skips, the Zig suite passed 3,233 with 36 skips, and the ABI
  manifest validated 187 types. The canonical report step also passed and
  bound the expected `fb1ba734...` runtime. The job then failed during artifact
  upload: combining workspace and external-sandbox files yielded a drive-relative
  common ancestor, which the artifact client resolved to the workspace and
  rejected. Product continuation was skipped; the overall job remains failed.
  Complete logs are preserved in
  `20261029T043734Z-36808-395bdefc2d454998bafe862392c2d784`.
  The report printed by the successful canonical step was recovered and locally
  schema-validated, with source-log hashes and explicit missing-original-artifact
  provenance, in `20261029T044531Z-12496-b601263634bd4dcc86c94a6fa905e6ea`.
  This is a recovered structured record, not a downloaded original artifact.
  Publication now stages a byte-verified identity copy beside the report and
  native output, rejecting existing or mismatched destinations; all uploaded
  files share that owned root. The actual staging block and focused workflow,
  governance and scoped checks passed 60/60 in
  `20261029T044341Z-79576-5aa7efc2b4c54729a6afcd59f429a0e9`.
  Admitted-product run `40819043691` was dispatched separately at application
  commit `dd3f5a732`, preserving the successful upstream execution while the
  publication correction is verified independently.
- Combined-head PR CI `40814412271` passed on x64 and native ARM64 at
  `dd3f5a732`. Each platform passed 1,062 PowerShell checks and all eleven
  native entries, including the loaded Terminal's 162 selftests, with two
  historical-evidence skips and zero blocking classifications or quarantine
  entries. Complete logs are preserved in
  `20261029T043522Z-41352-3f2727b5a70544ac87bc06a5df8f2806` (x64) and
  `20261029T051210Z-84604-1b041457a32f4208a5335b53dcc08be8` (ARM64).
  At `3f2b0f1ce`, the local Commands entry
  has passed its coverage check and promoted its result. File Operations
  execution and its pending coverage validation are recorded above; both
  interactive viewer entries remain pending.

- Combined workflow, candidate-review, scheduler, governance and scoped checks
  passed 138/138 in
  `20261029T034752Z-96224-e7af37b690674067aa963c3769376e9d`, excluding the
  tagged native build-toolchain lane. The preceding direct invocation included
  that lane without an explicit built validation profile and failed its setup
  guard; all 138 other checks passed. That failed invocation remains preserved
  in `20261029T034442Z-93132-5d1004b9d6084f0c8ee26ee94fffb6ca`.
  The real primary profile's 29 toolchain checks are recorded below, separately.
  Integration is prepared through the diagnostic worktree while the ongoing
  application run keeps commit `3f2b0f1ce` frozen. Its independently recomputed
  workspace snapshot still equals its starting snapshot
  `ae31f655059018eedd3fa02a6ba98eb34022b8a6eb7fbfe703d06a4c34488230`.

- Qualification at application commit `3f2b0f1ce` is recorded separately from
  the later native-workflow corrections. PR CI `40798928378` passed on x64 and
  native ARM64: each job passed 1,060 PowerShell checks and eleven native
  entries, with two historical-evidence skips. Both actual loaded Terminal
  plugins passed 162 deterministic selftests; the ARM64 runtime reproduced the
  locked raw DLL hash. Complete logs are preserved in
  `20261029T022034Z-96736-640b462155354373b5948006c647eeda` (x64) and
  `20261029T030424Z-95880-86eec96a24ab4eb799c5bc15e1e63675` (ARM64).
  Local run `20261029T012426Z-93224-2b01e80f5a6341bcb1d0691d944ae606`
  passed all 29 build/toolchain checks and refreshed its Debug receipt to
  `1d997c49ed7c7c55114a0682a2343bd4028cea440a606fad4b4c9569259a6fd5`.
  Commands execution completed with 957 passed, zero failed and six declared
  opt-in skips, including all 17 Terminal cases. Its coverage-list check passed
  and its entry was promoted. File Operations execution is recorded above;
  its coverage validation and both viewers await the shared desktop lease.
  Compare Directories recorded 259 passed, one failed and 21 environment skips:
  `search_service_sqlite_cold_start_stale_root_refreshes_before_query` timed
  out waiting for initial service warmup. Failed results remain intact; the
  empty SQLite store and input were preserved with verified hashes in
  `20261029T023504Z-38628-8a8656333db44f17a38b3371ef24c861`.
  No cause or successful retry is inferred. This run is not repository acceptance.
- The tests-disabled full x64 Release rebuild and portable clean-extraction
  smoke passed at `3f2b0f1ce` with zero warnings/errors, receipt
  `abfb4e76bf7318bb33baeef86fb1d7379582425348475a4a62918506921da11c`.
  The packaged runtime loaded natively and all built-in plugin contracts passed.
  `RedSalamander-7.0.1-x64-Portable.zip` is 25,228,838 bytes, SHA-256
  `77b79f55bc84a268e37abfd3587cbfdb9efc3a258d8b059705d383c57714eb19`.
  The complete build/package record is preserved in
  `20261029T032623Z-94028-75eae22297f4467e9c6dc90660881e18`.
  This shipping-profile result retains its original application source identity.

- Native run `40805550221` at `5e721727a` completed the ARM64 candidate build
  and seed/WER setup, then failed while launching `uucode_generate.exe` from
  the external locked package directory. No upstream test had executed. The
  complete log and initial-output artifact are preserved in
  `20261029T024956Z-87616-8962f735210041da8f7a5d8ecb3b08fe`.
  A captured-output generator reproduction using the pinned x64 Zig compiler
  and matching source/package layout failed with the same `FileNotFound` when
  the local cache was inside the source tree (exit 1), then passed with an
  external absolute cache (exit 0). Both graphs and logs are preserved in
  `20261029T030404Z-92292-029bc5aef59349ff9b8ccccf2592c178`.
  Native qualification now uses the exact candidate's `native-local-cache`,
  including its diagnostic executable boundary and event filters. The upstream
  test bodies, seed, filters and clean-first-execution rule remain unchanged.
  Workflow YAML and all eight PowerShell blocks parse. Focused workflow,
  governance and scoped checks passed 59/59, including execution of the actual
  diagnostic path guard with absolute/relative images and sibling, traversal
  and missing-image rejection, in
  `20261029T030812Z-73288-2b0efe7f120e4241b0c50dbfe7aebf90`.
  Native acceptance still requires a fresh successful run.

- Native run `40801003949` at `a635ffdfd` completed the ARM64 build and WER
  setup, but Zig rejected the command before upstream execution. PowerShell
  unwrapped the modern one-element seed array returned by `if`; splatting
  the resulting scalar passed each character as an argument. A typed string
  array now preserves the complete seed option in both branches. The regression
  executes the workflow's actual assignment and passes it to a real child
  process, checking both modern and legacy argv vectors. The failed log and
  initial-output artifact are preserved in
  `20261029T020745Z-81776-950f4b2bb08f487b8e79f155fa808128`.
  The actual modern and legacy argv vectors passed the child-process witness;
  focused workflow, governance and scoped checks passed 58/58 in
  `20261029T021539Z-93516-6241a7a2d5f84d0ab8148555a21df4ba`.

- Native run `40796626342` at `a7f11e810` completed the exact ARM64 candidate
  build with the shared fetch cache, then failed before upstream execution.
  Its clean runner profile lacked WER's `LocalDumps` parent, so creating an
  executable key with `New-Item` failed. The log is preserved in
  `20261029T011759Z-18144-1bf839eac61d4468b9a42dd9a84258dc`.
  Setup now opens or creates that parent through the registry API and disposes
  its handle; it still rejects existing executable keys and never force-creates
  them. Native unit and product acceptance remain pending the fresh run.
  YAML/PowerShell parsing and focused workflow, governance and scoped-testing
  checks passed 57/57 in
  `20261029T012734Z-92760-ffe72c07a2174b788837901e68b7c82f`.
- At `8a0e89a02`, local PrePush run
  `20261028T235142Z-81748-2cdfeb7bf8454498b806bdbf8b5f7a95` passed its
  Debug build, all 29 toolchain checks, profile refresh (receipt
  `fc0510ee949ad8d19860301629c4c74247dd6e9565fba22775d57bbca84e1dac`)
  and 260 Compare Directories cases with 21 environment skips. Commands then
  waited for the shared desktop mutex, owned by a separate DxUi validation.
  A read-only native owner query is retained in
  `20261029T010300Z-91732-d528a87c9f7d444e83e395e7fcdd7fef`.
  No desktop suite from this run launched. Its exact owned runner was stopped
  before integrating the cache fix, and the complete run was preserved with
  an interruption record. These results remain bound to `8a0e89a02`; the
  integrated source still requires qualification.
- PR CI `40790194244` passed its x64 job at `8a0e89a02`: 1,057 tooling checks
  and all eleven native entries passed, with two historical-evidence skips.
  The complete x64 log is preserved in
  `20261029T003709Z-86740-f7e4744d19cf4bbb929760ebe4ae0d92`.
  Fresh native candidate/product qualification of the cache fix is running as
  `40796626342` at `a7f11e810`; no passing result is inferred yet.

- The approved interactive PrePush at `ec6ad6c11`, run
  `20261028T224447Z-93944-65b4e072bbd7461399d2ba61becfd44b`, passed its
  Embedded-symbol Debug build, all 29 toolchain checks and profile refresh
  (receipt `8a3a6ab5ca7e0d9d1d349beed5ae6129ba2a3adae044bf05b57e6b6105430c67`).
  The runner then exited before Compare or any interactive entry. Direct
  execution of its scheduler against the two canonical deferred viewer entries
  reproduced strict-mode `The property 'Id' cannot be found` on an empty early
  set. Scheduling now projects early IDs through the pipeline; regression
  scenarios cover empty, late-only and mixed inputs. Interactive acceptance
  remains pending the corrected run. The focused scheduler, scoped-testing and
  governance regression passed 91/91 in
  `20261028T233856Z-78424-b76fb160d3b74359990865796b2af6f7`.

- PR CI `40779610989` at `ec6ad6c11` passed the x64 solution build and all
  eleven runtime entries, but tooling reported 1,052 passed, one failed and two
  historical skips. Its long-path Git regression received a successful Git
  stderr warning about `.git/objects/info/alternates`; combining stderr with
  stdout made that warning look like an inventory row. Metadata now consumes
  stdout only, with successful diagnostics still visible and failed-command
  diagnostics retained. Real Git aliases emitting both streams and a nonzero
  exit cover both boundaries. The combined candidate-review, scheduler,
  scoped-testing and governance checks passed 99/99 in
  `20261028T234753Z-75512-55b70c3d9559450e9039da9720a2a66a`.

- Diagnostic run `40790104792` at `8149476b4` exhausted its 20-minute
  contained build limit during the `translate-c` fetch. Aro and Pixels fetches
  had completed in 191 and 187 seconds; the third Zig parent was still running
  after 660 seconds. No output-drain stall was observed. Its explicit
  `qualification=false` record and complete logs are preserved in
  `20261029T002335Z-80328-9cb3b1ae17a2483eb64a12ace14079ff`.
  Each fetch previously deleted its compiler cache, repeating Zig 0.17 maker
  compilation. One repository-serialized, disposable fetch-session cache now
  spans the dependency closure; all package/archive/license checks remain.
  Cold/warm/second-fetch-failure ownership checks passed with the lifecycle
  suite, 39/39, in
  `20261029T003210Z-83120-67aaadf3eeb44a19afc3230afde187f8`.
  A cold x64 build passed in 399,025 ms, reproduced the exact locked DLL and
  removed the cache, in
  `20261029T003622Z-89180-65cfc0e461644515b6b647bd29826f24`.
  Its first fetch took 106 seconds and each later fetch less than one second.
  The record binds lifecycle SHA-256
  `335880ddfbd03844ca36dac906ba816c53eea031f040717debdadddcbe6e8ae4`;
  subsequent indentation cleanup does not change the executed statements.
  This local result does not establish the native timeout's underlying cause
  or native ARM64 acceptance; a fresh native qualification remains required.
  The final cache implementation passed lifecycle, workflow, scoped-testing
  and tooling-governance checks, 96/96, in
  `20261029T004714Z-90328-752ebfc3157d4edeb715fcb4265ee46e`.

- Native ARM64 candidate run `40761642957` at `62d181cf4` reached the
  120-minute job limit during the exact-candidate build; no upstream test ran
  and admitted-product continuation was skipped. The terminal build in PR CI
  `40761500830` also remained unfinished before the obsolete run was cancelled.
  Neither establishes native qualification or a compiler crash. Build progress
  now distinguishes repository/dependency/drive phases, running owned commands,
  and output pipes remaining open after process exit. The optional build-only
  diagnostic dispatch has a separate check name, a 20-minute contained watchdog
  and an explicit `qualification=false` artifact; it cannot publish qualification
  evidence or start product continuation. The workflow YAML, eight PowerShell
  run blocks and generated worker parsed successfully. Focused lifecycle,
  workflow and governance checks passed 85/85 in
  `20261028T234817Z-90600-7717d6bb3cd34622b885d5f114db47d1`.

- The standalone x64 Zig repeat `ghostty140-zig-x64-clean-exit-20261028a`
  passed 3,233 tests, skipped 36 and returned exit 0. Its record retains the
  executable and log hashes and verifies unchanged executable bytes. It closes
  the earlier missing-exit-code evidence gap for that standalone suite; it does
  not replace native ARM64 or complete workflow qualification.
- The final local non-interactive PrePush attempt at `62d181cf4`, run
  `20261028T205029Z-58832-bcd0d2c07e1844db8a929d83588bed49`, passed all 29
  build/toolchain tests and 260 Compare Directories cases, with 21 environment
  skips. Its profile refresh passed with receipt
  `bf0077b96f5178761fccbb0c8a8f2c465397a9ef438c5dcfc69e4be1db5f4980`.
  The repository verdict remains NOT_EVALUATED: CI/interactive obligations were
  deferred and the shared sandbox audit reported 74 unexpected top-level/run
  entries. The preserved run is not a complete PrePush or Full pass.
- PR CI run `40761500830` passed its x64 build and all eleven native runtime
  entries, but its tooling profile had two failures (1,050 passed, two skipped):
  the staged Terminal output expectation omitted Aro/Pixels, and a nested Git
  metadata path exceeded the Windows short-path limit. Its complete job log and
  validation artifact are preserved in `ghostty140-prci-x64-40761500830`.
  Staging expectations now include both notices. Candidate metadata queries use
  invocation-local long-path support, with independent fixture repositories and
  a real clone/inventory/ABI regression beyond MAX_PATH. The seven candidate
  review cases passed with system/global Git configuration isolated in
  `ghostty140-long-git-path-20261028e`; the deeper-TEMP repeat also passed all
  seven in `ghostty140-long-git-path-20261028f`. Fixture owners use short,
  canonical sandbox scratch paths so Git's separate working-directory bound
  does not obscure file/metadata long-path coverage. The complete CI tooling
  profile then passed 1,053 tests, zero failed, with two historical-evidence
  skips in `20261028T220741Z-6780-0c0a868b09834aa1adfa3c7e42563316`.
  Its execution digest remained
  `efb23e1071f0ca6bf2773b078950c9b9655cf822772c4b4bb944f2ab771e7073`
  before/after execution. This is tooling-profile evidence, not repository Full.
- All 35 completed `ghostty140-*` runs were preserved through
  `Save-RSTestSandboxRunEvidence` after checking live command references. The
  helper verified complete inventories before and after each same-volume rename;
  raw passing and failed results now live under `evidence/completed-runs/`.
  Transaction receipts are retained in
  `20261028T221107Z-93048-dbbd98b0ba314902b85285c4c0a11558`. The shared
  sandbox's other layout issues remained non-blocking audit diagnostics; incomplete
  CI and interactive coverage kept repository acceptance open.
- Review thread `4405126908` originated at `77ba80042` and is resolved after
  verifying the `62d181cf4` executable/schema/event-filter corrections and the
  diagnostic-only retry rule. No runtime acceptance is inferred from resolving
  a review thread.

- Native run `40753290262` on `fbcece451` failed during dependency-fetch cache
  cleanup: Windows denied removal of the Zig 0.17 `maker.exe` in the Pixels
  fetch cache. That secondary error could mask a primary fetch failure; the
  underlying fetch outcome is not inferred from it. Failed logs are preserved
  under `ghostty140-native-arm64-40753290262` in the marked sandbox.
- Temporary cleanup now revalidates containment and retries I/O/access errors
  up to 20 attempts with 250 ms delays. A failed fetch keeps its primary error
  when cleanup also fails. The real open-file regression, permanent lock,
  primary-error preservation and path rejection passed; focused lifecycle,
  governance, build-evidence, plugin-source and workflow tests passed 136/136
  in `ghostty140-cleanup-retry-20261028b`. Updated workflow tests passed 36/36.
  A fresh forced x64 candidate rebuild reproduced the same DLL. Compile-only
  workflow harness validation returned exit 0 in
  `ghostty140-workflow-20261028200931`; no unit execution is claimed for that
  compile-only check.
- Native simple runners now receive `--seed=0x97d35bc7` explicitly. Failure
  artifacts require acquired scratch, and event diagnostics require both the
  source root and a known test executable name. A diagnostic retry still never
  qualifies an initially failing run.
- The first real PrePush attempt at `fbcece451` completed a full x64 Debug
  solution build (receipt
  `1c447fd8b0f550a097262ace08a778a2307c1cdb245c1b5802217ef4abe5b2c8`)
  with one C4464 warning in unchanged `Plugins/ViewerWeb/ViewerWeb.cpp`.
  Build/toolchain tests passed 29/29, including the real targeted deployment
  rebuild. Its owner PID 79472 was then cancelled before Compare Directories
  because the native failure requires changed validation inputs. Run
  `20261028T192005Z-79472-6a990b54e976475e8a82a3dd1c479bd1` retains its
  checkpoint and explicit cancellation record. It is not a PrePush or Full
  pass; a matching rebuild must recover any abandoned artifact scope before
  validation of the updated commit.

- Cache coexistence follow-up: Zig 0.17 canonical archives use distinct
  `-zig-0.17.0.tar.gz` download names. Reusing the pre-upgrade Wuffs filename
  would otherwise reject a valid warm 0.16 cache. The old archive is retained,
  and new namespace/governance/scoped tests passed 51/51 in
  `ghostty140-cache-namespace-20261028a`. The preceding Release/performance
  receipts remain bound to their original lock bytes; this follow-up changes
  archive filenames, with source/package contents unchanged.
- Draft PR: <https://github.com/RedSalamanders/RedSalamander/pull/148>.
  The initial native dispatch was cancelled while queued to qualify the final
  cache-namespaced lock instead. PrePush planning delegates matching PR entries
  as pending CI and retains the build-toolchain and Compare Directories local
  obligations. The four focus-taking suites are authorized; acceptance remains
  pending their actual results.

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
