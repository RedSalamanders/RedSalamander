# TestRuns

This folder contains local build/test run logs used to validate critical fixes. Generated run folders are ignored by default to keep the repository healthy; force-add only small, explicitly curated evidence when a spec or release review requires it. Do not maintain a per-run exception list in `.gitignore`: already-tracked archives remain tracked, and new files under this folder require `git add -f`.

## Conventions

- Create a new subfolder per run under `Specs/TestRuns/<Profile>/<Area>/yyyy-MM-dd_HHmmss/`.
- Prefer `.txt`, `.md`, `.json`, or `.jsonl` files (the repo ignores `*.log`).
- Keep logs small and avoid secrets.
  - `--fileops-selftest` includes optional Phase 16 “Remote*Secret” checks; these should normally be **skipped** in curated runs unless explicitly validating WinCred/Connection Manager behavior.
- Archived self-test results must preserve per-case status and reasons. Every declared case must be represented as `passed`, `failed`, or `skipped`; skipped cases must retain their reason text.
- A newly checked-in curated artifact must be at most 2 MiB, and the aggregate checked-in payload for one run must be at most 5 MiB. Raw dumps, broad traces, and high-frequency metric streams belong in local or CI artifact storage; retain only the compact rows and summaries needed to review the claim. When raw row ordering supports the claim, compaction must preserve a digest-bound deterministic summary with the original byte/row counts and the claim-bearing rows in their original order; otherwise retain the bounded raw input outside Git and do not claim closeout from a lossy summary.
- When `perf_metrics.jsonl` records a 12-hex `machineHash`, a hash-shaped profile directory must match that value. Commit hashes belong in metric metadata or the run name, never in the machine-profile slot.
- Digest-bound raw evidence must disable Git text normalization in its owning `.gitattributes` archive pattern so checkout preserves the measured bytes. Preserve original trace line endings and terminal blank lines as evidence.
- Before force-adding a generated run, execute `Tools\Test-TestRunArchive.ps1 -RunPath <run-folder>`. This explicit pre-stage mode includes ignored raw files and must pass before the run can be accepted. The validator parses every nonblank JSONL row, not just the first row, and dispatches area/schema-specific checks when a governed evidence format is present.
- `Tools\Test-TestRunArchive.ps1` remains the fast changed-set gate. Closeout must also execute `Tools\Test-TestRunArchive.ps1 -Inventory`, which validates every checked-in archive added or modified since the archive contract was introduced by commit `8b9e36191835cc71b5ee0dfeea4dddd5d942dd5a`; older untouched evidence is grandfathered, but changing it brings it under the current contract. Validation failure is a red closeout and must never be summarized as green. Full Pester coverage exercises pre-stage rejection, unchanged contract-era inventory, exact limits, and this failure behavior.

Rename detection is disabled in both changed-set and inventory diffs. A staged or committed
move of a grandfathered blob is an addition at its new path and must pass the current
contract; exact old path identity is part of grandfathering.

The history-free source migration preserves the same boundary through parentless public baseline
`4a9489e9cb2d1190528b8ec35971fba885fd7792` (tree `3fb1f79ac855236a8cc2e0bf9b348601ec8cd5f3`).
It contains only 2,075 exact path/mode/blob matches between the pre-contract tree
`eaf640798eb0fa21f2b5b6e3883891ac2442a255` and already-published source
`58d1e03fdef9ec93e0cfe63139eed80461c9ca18`, all under `Specs/TestRuns`.
It has no parents, imports no DualTail commit history and publishes no additional blob contents.
Keep this immutable baseline as an ancestor when merging; a squash that drops it cannot retain
its grandfathering proof. Inventory uses the original contract ancestor when available, otherwise
this exact imported baseline ancestor. Missing or unrelated ancestry validates every tracked file.
Any later committed, staged or worktree edit is governed; explicit `-RunPath` validation also checks
legacy files. Changed mode compares committed branch additions/edits with `origin/main`, falling back
to legacy `origin/master`, and includes index/worktree edits. Hosted validation fetches complete public
history so the real inventory check executes rather than skipping for missing ancestry.

### Composite validation evidence

Terminal VT pairs require identical environment metadata by default. A reviewed
compiler upgrade may declare `-ZigToolchainTransition <baseline>,<candidate>`
to `Tools/Test-TestRunArchive.ps1`; the two distinct exact versions must match
the recorded Zig values. Every other environment field, behavior digest, raw
sample binding, and baseline x1.10 gate remains enforced. Describe these results
as a comparison of the complete runtime/toolchain upgrade.

Only a fresh Full run provides final closeout evidence. An explicit exact Resume can
provide a complete repository verdict for local continuation, but it must be described as
composite evidence with direct origin provenance. A local ledger, provisional result,
explanation report, monolithic logical candidate, Commands-family result, or Affected run
is diagnostic and must not be described as Full repository evidence.

A curated composite run must retain aggregate v2, immutable plan and decision identities, promotion-certificate
digests, per-entry verdicts and coverage scope, accepted build-receipt/portable-attestation
identity, and proof of launches actually avoided. Do not archive the complete reusable
evidence cache. `Test-TestRunArchive.ps1` must support both the legacy and v2 aggregate
contracts. The legacy v1 sidecar may accompany v2 during migration but cannot replace
its provenance/verdict axes. Validation-planner performance evidence uses the
`Validation/` area under the normal machine-hash profile; a blocked Commands avoidance
scenario records zero avoided Full duration and the payload-boundary reason.

## Capture workflow (HARD REQUIREMENT: always archive runs)

Runner-owned selftests write their raw artifacts under:

- `REDSALAMANDER_TEST_ROOT\\runs\\<runId>\\artifacts\\selftest\\last_run\\`
  - `results.json` + `trace.txt` (overall selftest run)
  - `fileops\\results.json` + `fileops\\trace.txt` (FileOps suite)
  - `compare\\results.json` + `compare\\trace.txt` (CompareDirectories suite)
  - `commands\\results.json` + `commands\\trace.txt` (Commands suite)
- `REDSALAMANDER_TEST_ROOT\\runs\\<runId>\\artifacts\\selftest\\last_run\\run-all-tests-results.json`
  when `Tools\\Run-AllTests.ps1` created the invocation.

Promote only reviewed runs into this folder so durable performance and behavioral evidence can be compared over time. Ordinary runs remain under `RedSalamander.Perf` and do not dirty the checkout.

`REDSALAMANDER_TEST_ROOT` is exact `X:\\RedSalamander.Perf` on a fixed local drive, defaulting `X:` to the repository drive. All raw runtime data and validation evidence remain beneath that root. Do not use repository `.build\\TestSandbox`, historical `RSPerf`, `%LOCALAPPDATA%`, or process temp as test-data locations.

In Debug builds, the selftest harness creates any compact automatic archive inside the current `RedSalamander.Perf` run. Promotion into this checked-in tree is explicit:

- `Specs/TestRuns/<ComputerHashName>/<Area>/yyyy-MM-dd_HHmmss/`

The promotion workflow is:

1. Run the scenario (build + selftest).
2. Review the run-owned compact archive and raw artifacts beneath `REDSALAMANDER_TEST_ROOT`.
3. Explicitly copy the allow-listed evidence into `Specs/TestRuns/<ComputerHashName>/<Area>/yyyy-MM-dd_HHmmss/` and validate it with `Tools/Test-TestRunArchive.ps1`.

Curated runs may legitimately contain skipped cases when prerequisites are absent. Skips are expected for conditional coverage as long as the reason is preserved in `results.json`.

In-product selftests never write directly to `Specs/TestRuns`. This separation is mandatory even from a developer checkout.

## Compare and trend analysis

Compare two archived runs (suite auto-detection; use `-Suite` if needed):

```powershell
.\Tools\CompareTestRuns.ps1 `
    Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-27_121904 `
    Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-28_100446 `
    -ShowTraceDiff -MaxTraceDiffLines 50
```

Analyze timing evolution over *many* archived runs:

```powershell
.\Tools\AnalyzeTestRuns.ps1 Specs\TestRuns\<ComputerHashName>\FileOps
```

## Profiles

- `<ComputerHashName>/` — Default profile per machine (SHA-256 of `MachineGuid`, first 12 hex chars), so timings don’t get mixed across different hardware.
- Custom/legacy profiles are allowed (e.g. `Tiny/`), but prefer per-machine hashes for performance comparisons.

## Areas

- `FileOps/` — `RedSalamander.exe --fileops-selftest` (plus copied SelfTest `results.json` + `trace.txt`) and related validation runs.
- `CompareDirectories/` — `RedSalamander.exe --compare-selftest` (plus copied SelfTest `results.json` + `trace.txt`).
- `Commands/` — `RedSalamander.exe --commands-selftest` (plus copied SelfTest `results.json` + `trace.txt`); prefer deterministic local scenarios covering registry, shortcut, menu, and command-behavior contracts.
- `SelfTest/` — `RedSalamander.exe --selftest` (captures all suites; useful to archive a single run folder with everything).
- `Monitor/` — historical bundled-viewer receipts, including `RedSalamanderMonitor.exe --chrome-selftest`. New standalone viewer qualification belongs to RedMonitor. Consumer client evidence uses `Validation/`; Preferences interaction stays under `Commands/` with the owning source/build identities.
- `Viewers/` — focused viewer-plugin runtime, safety, unload, and responsiveness scenarios; retain the assertion trace, compact results summary, and relevant `viewer.*` JSONL rows together.
- `Terminal/` — focused embedded-Terminal runtime and responsiveness scenarios. Test-enabled Release Kitty evidence uses `PluginContractTests.exe --terminal-selftests` and retains `results.json`, `trace.txt`, and `perf/perf_metrics.jsonl` together; real host lifecycle evidence remains under `Commands/`. A VT-upgrade corpus archive is identified by `TerminalVtUpgradeEvidence.schema.json`/the VT corpus identity and must bind its schema-valid `results.json`, archive run id, machine profile, repository/configuration identity, `trace.txt`, and all 600 ordered metric rows. Compare a claim-bearing pair only with `Tools\Test-TestRunArchive.ps1 -BaselineRunPath <product-run> -CandidateRunPath <candidate-run>`; the command validates both archives, recomputes nearest-rank percentiles from all raw samples, verifies the control/delta digest contract, and applies the candidate ceiling from the baseline p95.
- `TerminalEngine/` — Gate-0 candidate collection/finalization evidence created only by
  `Tools/Evaluate-TerminalEngines.ps1`. Each immutable run directory is
  `Specs/TestRuns/<ComputerHashName>/TerminalEngine/<runId>/` and contains
  exactly `terminal-engine-evidence.v1.json`,
  `terminal-engine-evidence.v1.sha256`, and an optional human summary that adds
  no evidence. The canonical record binds the repository, schema, frozen
  requirements/anchors/corpus/protocol/candidate-set identities, provider
  harness, native lane, toolchain, aggregate measurements, candidate outcome,
  and cross-manifest references. Never archive source trees, patches, binaries,
  packages, extracted files, local paths, certificates or keys, secrets,
  commands, terminal/screen/clipboard/image payloads, user data, or user/machine
  names. Raw candidate artifacts and reproducers remain disposable under
  `.build`; the evidence record retains only their content-free identities,
  sizes, categories, and digests. A round with no executable
  `disposition="collect"` candidate creates no directory here: static source
  rejections and controls cannot truthfully attest a native lane or its derived
  toolchain identity. Verify an existing record with
  `Read-RSTerminalEvidenceManifest` from
  `Tools/TerminalEngineEvaluator.psm1`, verify the complete three-lane set by
  passing those exact paths to
  `Evaluate-TerminalEngines.ps1 -Mode Finalize -RequireAllEvidence`, and run
  `Tools/Test-TestRunArchive.ps1` before
  publication. Retain every manifest referenced by an engine decision for the
  selected lock's supported lifetime and through validation of its rollback
  lock; never replace it with a newer run having the same nominal lane.
- `TerminalEngineRecovery/` — small, tracked, machine-independent Gate-0
  closeout archives used to verify a published round from a fresh clone after
  ignored coordinator state has been removed. This is an explicit exception to
  the machine-profile layout above. Every Round 5+ Gate-0 round must publish
  exactly
  `Specs/TestRuns/TerminalEngineRecovery/round<ordinal>-<yyyyMMddTHHmmssZ>-<16-lowercase-hex>/TerminalEngineRound<ordinal>RecoveryArchive.v1.json`.
  The timestamp is ASCII and must round-trip through invariant
  `yyyyMMdd'T'HHmmss'Z'` UTC parsing; impossible timestamps, Unicode decimal
  digits, case aliases, backslashes, rooted paths, ADS, dot segments, and extra
  path segments are invalid. The file is one RS-JCS canonical JSON value,
  UTF-8 without BOM or trailing newline, at most 2 MiB, with
  `recordKind=terminal-engine-recovery-archive`, the exact ordinal and audited E
  decision-cutoff commit, embedded canonical boundary/receipt/quiet objects and
  raw digests, plus only the bounded content-free cleanup inventory. It must
  not contain source, checkout, build, binary, package, cache, repro, terminal,
  clipboard, user, machine, credential, or secret content. Verify a named
  archive without enumeration using the explicit invocation below. Clone
  verification reads only the tracked archive and embedded objects; it never
  reopens ignored `.build` paths:

  ```powershell
  $archivePath = '<path>'
  $roundOrdinal = 5
  $archiveSha256 = '<sha256>'
  $decisionCutoffCommit = '<E>'

  .\Tools\Test-TerminalEngineRecoveryArchive.ps1 `
      -ArchivePath $archivePath `
      -RoundOrdinal $roundOrdinal `
      -ExpectedRawSha256 $archiveSha256 `
      -ExpectedDecisionCutoffCommit $decisionCutoffCommit
  ```

  Retain every archive for as long as its Gate-0 Done round or any product/rollback
  lock derived from that lineage is supported, and never replace bytes at an
  existing path.
