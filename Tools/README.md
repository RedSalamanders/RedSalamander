# Repository tooling

`Restore-DxUi.ps1 -Platform x64 -Configuration Debug -CheckUpdates` restores the exact
`Dependencies/DxUi.lock.json` pin before an IDE build. Root `build.ps1` performs this automatically.
The source and compiler/SDK/CRT-isolated outputs stay under `.build/dependencies/DxUi`; a dirty
managed source rejects build/test reuse. The optional newer-version notice never edits the lock; a validated candidate
is yellow and prints the normal update command plus its `-UpdateOnly` alternative. A candidate without a successful
completed DxUi validation is red and must remain pinned.
`Update-DxUi.ps1` makes that branch update: it accepts only a current `main` commit with successful completed DxUi
CI, changes the lock, and runs the Full suite. `Update-DxUi.ps1 -UpdateOnly` skips the local suite only when matching
product validation already passed in another environment; neither form creates a commit.
Restore evaluates the selected configuration's first-party compiler host and preserves explicit
`PreferredToolArchitecture` overrides. The library project reference forwards that same host.
See `Specs/Build/Build_Toolchain.md` for the manual upgrade and per-module provenance contract.


`Tools/` contains the repository's supported developer commands, build and test
helpers, release policy, and preserved terminal-engine evaluation evidence. Start
here instead of guessing from a filename.

The machine-readable source of truth is [`tool-inventory.json`](tool-inventory.json).
The specification inventory admits the three exact reviewed I19 Done paths;
its focused test also rejects unreviewed siblings. ARM64 and ASan qualification
remains separately indexed under WIP and is not inferred from plan closeout.
It records each file's kind, public/internal visibility, lifecycle, purpose,
consumers, prerequisites, side effects, outputs, owning specification, tests,
exact replacement path, historical `manualEntryPoints`, and any required
retention `supportPolicy`.
Reviewed exceptions for legacy public filenames are explicit `namingPolicy`
fields; they are not precedents for new commands.
Validate it after any tooling change:

```powershell
.\Tools\Get-ToolInventory.ps1 -FailOnFindings
.\Tools\Get-ToolInventory.ps1 -Format Object |
    Where-Object Lifecycle -eq 'historical'
```

Every file directly beneath `Tools/` has an explicit manifest entry. A new
top-level file therefore fails validation until its ownership and use are
documented. The only directory rule is the conventional
`Tools/Tests/*.Tests.ps1` Pester lane; unusual tests and historical tests can
override that rule with explicit entries.

Current source-contract checks complement native behavior tests. The history-popup guard
requires a readable popup state and keyboard selection within the existing bounded wait;
separate diagnostic bookkeeping does not relax that success condition.

## Where to look

Foreground tests reuse the native `DirectedSelfTestInputWarning` in
`Tests/TestSupport/`: a large warning centered on the test app appears three
seconds before input and remains during the foreground suite or isolated case.
Pause keyboard and mouse input until it closes. Background lanes retain their
no-activation guards. The unified runner continues to own the desktop lease.

The runner gives each FileOperations child process a fresh `LOCALAPPDATA` below
the selected run's `scratch/fileops-profile` directory, including classification
retries. Fake MTP journals remain inside that child's lifetime and cannot collide
with the caller's journal history. The parent environment is unchanged. Use this
runner for FileOperations qualification; a direct executable diagnostic must also
provide an isolated local-data directory under the marked TestSandbox root.

| Location | Contents | How it is used |
|---|---|---|
| `Tools/` | Stable commands, machine contracts, two documented compatibility loaders, and sealed top-level history | Invoke commands by their documented path. New reusable implementation does not belong here. |
| `Tools/Modules/Auditing/` | Shared repository source discovery and audit primitives | Imported by source-audit commands. |
| `Tools/Modules/Build/` | Build locking, project selection, process/environment, vcpkg, version, receipt, and portable-evidence helpers | Imported by `build.ps1`, CI workflows, vcpkg tooling, test orchestration, and installer commands. |
| `Tools/Modules/Packaging/` | Release artifact, portable package, runtime dependency, and winget publication policy | Imported by installer and release workflows. |
| `Tools/Modules/Reporting/` | Shared test-run and source-line reporting primitives | Imported by stable reporting commands. |
| `Tools/Modules/Testing/` | Test inventory, archive validation, and canonical run-plan implementation | Imported by stable test commands and current contract tests. |
| `Tools/Modules/Tooling/` | Specification and Tools inventory governance | Imported by stable governance commands and focused tests. |
| `Tools/Tests/` | Pester policy and behavioral tests plus reviewed `TestSupport.psm1` | Executed through the current Tools Pester profile in `Run-AllTests.ps1`. |
| `Tools/TerminalEngine/` | Current selected-runtime helpers, dormant generic lifecycle support, and sealed Gate-0 history | Current runtime files are `active`, stable dormant generic-lock files are `compatibility`, and frozen candidate-selection machinery is `historical`. Do not infer currency from the directory alone. |

Lifecycle values have precise meanings:

- `active`: supported current behavior.
- `compatibility`: a stable legacy path retained for known consumers; new callers
  should use its `replacement` when one is named.
- `historical`: retained to reproduce or authenticate sealed evidence, not a
  current product workflow.
- `generated`: checked-in output owned by a named generator.
- `data`: declarative input consumed by tooling.

Historical retention is a closed manual graph. Every `historical` manifest row
must name one or more exact `manualEntryPoints`, limited to these reviewed
commands:

- `Tools/TerminalEngine/Invoke-HistoricalTerminalTests.ps1` runs the sealed
  historical Pester cohort.
- `Tools/Evaluate-TerminalEngines.ps1` reproduces the sealed evaluation.
- `Tools/Test-TerminalEngineRound4Closeout.ps1` reauthenticates Round-4
  closeout.

Each reviewed root names itself. Inventory validation rejects an orphaned
historical file, an unreviewed or missing path, and a path that does not resolve
to an explicit `historical` command entry.

Before running a command that downloads, deletes, modifies credentials, changes
firewall rules, accesses a device, or rewrites checked-in assets, read its
`sideEffects` entry in the manifest and its comment-based help.

## Common entry points

| Goal | Command |
|---|---|
| Build the repository | `./build.ps1` (repository root) |
| Run the canonical tests | `./Tools/Run-AllTests.ps1 -Suite CI` |
| Run the full closeout suite | `./Tools/Run-AllTests.ps1 -Suite Full` |
| Build the selected terminal runtime | `./Tools/Build-GhosttyTerminalRuntime.ps1` |
| Check the selected Ghostty runtime | `./Tools/Get-GhosttyTerminalRuntimeStatus.ps1 -Mode Locked` |
| Freeze and review an exact Ghostty candidate | `./Tools/New-GhosttyTerminalRuntimeCandidateReview.ps1` |
| Validate a rederived Ghostty candidate overlay | `./Tools/New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1` |
| Inspect test coverage | `./Tools/Get-TestInventory.ps1` |
| Validate specification governance | `./Tools/Get-SpecInventory.ps1 -FailOnFindings` |
| Validate this inventory | `./Tools/Get-ToolInventory.ps1 -FailOnFindings` |
| Run raw cloc for solution directories | `./Tools/InvokeSolutionCloc.ps1` |
| Review archived performance runs | `./Tools/Show-PerfRuns.ps1` |
| Validate archived or paired Terminal VT evidence | `./Tools/Test-TestRunArchive.ps1` |
| Remove NTFS alternate streams | `./Tools/remove-ads.ps1` |

## Top-level inventory

This table is the concise human map. The manifest contains the full operational
contract and is authoritative when the table is intentionally brief.

| File | Role | Used by / use it for |
|---|---|---|
| `AnalyzeTestRuns.ps1` | Public test-run trend command | Manual historical timing/result analysis. |
| `Audit-ComctlReportSurfaces.ps1` | Public UI source audit | CI and documentation drift checks for list-view/tooltip surfaces. |
| `Audit-RemainingWin32UiDependencies.ps1` | Public UI source audit | Inventory or reject remaining Win32/GDI UI dependencies. |
| `Audit-VisibleNativeSurfaces.ps1` | Public UI source audit | CI and documentation drift checks for visible native controls. |
| `Build-GhosttyTerminalRuntime.ps1` | Public active runtime builder | Thin stable entrypoint for canonical Product builds and explicit isolated candidate builds; delegates validation, restore, build, and attestation to `GhosttyRuntimeLifecycle.psm1`. |
| `Build-TerminalEngine.ps1` | Dormant generic lifecycle compatibility command | Stable facade for `External/TerminalEngine/TerminalEngine.proj` and reviewed future-upgrade evidence; not the product builder. |
| `Clean-TestSandbox.ps1` | Public maintenance command | Lists or explicitly removes one exact run directory beneath marked `X:\RedSalamander.Perf`; it never touches historical outside-root locations. |
| `CompareTestRuns.ps1` | Public report command | Compares two archived self-test runs. |
| `Evaluate-TerminalEngines.ps1` | Historical Gate-0 command | Reproduces/authenticates the sealed multi-candidate selection; not a product build. |
| `GenerateDebugIcons.ps1` | Internal asset generator | Regenerates checked-in debug-badged icon resources. |
| `Get-GhosttyTerminalRuntimeStatus.ps1` | Public read-only runtime status command | Validates the lock without network, observes official upstream/advisories, or freezes one explicit candidate; writes only below `.build`. |
| `Get-SpecInventory.ps1` | Public governance command | Validates specs/WIP/links and writes derived reports below `.build`. |
| `Get-TerminalDependencyStatus.ps1` | Dormant generic lifecycle command | Explicit-lock validation and opt-in advisory checks for a reviewed future upgrade. |
| `Get-TestInventory.ps1` | Public discovery command | Derives test registrations, projects, Pester cases, and run-plan coverage. |
| `Get-ToolInventory.ps1` | Public governance command | Validates and renders `tool-inventory.json`. |
| `Get-VisibleTypographyAudit.ps1` | Public UI source audit | Finds visible typography paths outside the intended DirectWrite model. |
| `Invoke-SanitizedMsbuild.ps1` | Public build diagnostic command | Runs contained MSBuild with normalized environment plus explicit profile/dependency coordination. |
| `InvokeSolutionCloc.ps1` | Supported manual compatibility command | Preserves hand-run raw cloc arguments while forwarding solution discovery to `Measure-SourceLines.ps1 -Scope Solution`. |
| `Manage-WindowsCredentials.ps1` | Public support command | Lists or, with confirmation, removes selected Windows credentials. |
| `Measure-SourceLines.ps1` | Public metrics command | Produces the categorized repository source-line report. |
| `New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1` | Public candidate-overlay command | Proves clean apply, exact diff reproduction, review bounds, complete concern disposition, every newly added private ABI value's separation/parity, and candidate-descriptor binding below `.build`. |
| `New-TerminalEngineNotices.ps1` | Public packaging command | Generates deterministic terminal third-party notices. |
| `New-GhosttyTerminalRuntimeCandidateReview.ps1` | Public candidate-review command | Authenticates and freezes one explicit Ghostty commit into canonical provenance, inventory, ABI, security-review, and input-bound descriptor evidence below `.build`. |
| `Prepare-TerminalEngineUpgrade.ps1` | Dormant generic lifecycle command | Collects/finalizes explicit-lock future-upgrade evidence on a disposable runner. |
| `README.md` | Human inventory | This navigation and safety guide. |
| `remove-ads.ps1` | Public filesystem command | Canonical explicit NTFS alternate-stream removal. |
| `ResolveVersionForMsbuild.ps1` | Internal build command | Supplies synchronized version properties to MSBuild. |
| `Restore-DxUi.ps1` | Public dependency command | Restores the exact public DxUi source and isolated per-platform build inputs; optional advisory never advances the pin. |
| `Update-DxUi.ps1` | Public dependency-update command | Advances the DxUi lock only to a successfully validated main commit, then runs or explicitly skips Full product validation. |
| `Restore-TerminalEngineInputs.ps1` | Dormant generic lifecycle command | Downloads or verifies inputs only for an explicit reviewed generic lock. |
| `Run-AllTests.ps1` | Public canonical test command | CI/Full execution with profile-scoped build inputs and all test-generated local data beneath exact `X:\RedSalamander.Perf` on a selected fixed local drive. First initialization and alternate-volume FileOps roots require explicit opt-in; `path.local.alternate` can pin cross-volume coverage to one exact second-drive root. Optional per-machine resources come from `config\machine-resources.json` beneath the primary root. |
| `Run-MtpLiveCloseout.ps1` | Public opt-in device command | Runs MTP/PTP live-device closeout smoke and archives all local evidence beneath the selected `X:\RedSalamander.Perf` run. |
| `SanitizedEnvironment.ps1` | Internal compatibility loader | Preserves the frozen Terminal tooling path; current callers import `Modules/Build/SanitizedEnvironment.psm1`. |
| `Show-PerfRuns.ps1` | Public performance report command | Lists, compares, and trends archived performance evidence. |
| `TerminalEngineEvaluator.psm1` | Historical Gate-0 module | Sealed candidate collection/finalization implementation. |
| `TerminalEvidence.psm1` | Historical evidence module | Authenticates sealed terminal evaluation records. |
| `Test-TerminalEngineRound4Closeout.ps1` | Historical closeout command | Reauthenticates the frozen Round-4 selection evidence. |
| `Test-TestRunArchive.ps1` | Public evidence command | Validates changed, staged, complete, or paired Terminal VT `Specs/TestRuns` archives, including area schemas and identity/performance bindings. |
| `TestRunPlan.ps1` | Internal compatibility loader | Preserves sealed historical test paths; current callers import `Modules/Testing/TestRunPlan.psm1`. |
| `tool-inventory.json` | Tooling-governance data | Machine-readable contract consumed by the validator and tests. |
| `validation-impact.json` | Trusted-validation policy data | Reviewed path-to-impact/build mapping, including executable `Specs` data; specific rules accumulate and unknown relevant JSON/JSON5 widens to Full. |
| `Verify-NoSubclassManager.ps1` | Internal policy command | Rejects production use of the retired SubclassManager pattern. |
| `Verify-TerminalEngine.ps1` | Dormant generic lifecycle command | Verifies explicit-lock generic outputs for historical or reviewed future-upgrade use. |

## Implementation module map

Top-level commands are intentionally thin. These modules own reusable behavior;
the machine inventory carries their complete consumers, effects, outputs, owners,
and focused tests.

| Module | Responsibility |
|---|---|
| `Modules/Auditing/RepositorySourceScanner.psm1` | Deterministic repository source discovery and match scanning. |
| `Modules/Build/ArtifactOperationLock.psm1` | Profile artifact, platform-scoped vcpkg-root, and packaging coordination with native immediate-child-authenticated delegation, scoped contamination, path-qualified residual-process policy, reparse-safe guarded output paths, and same-directory atomic file publication. |
| `Modules/Build/BuildProjectSelection.psm1` | Solution project selection and transitive dependency resolution. |
| `Modules/Build/BuildEvidence.psm1` | Manifest-derived artifact/runtime closure, legacy/current vcpkg app-local path normalization, complete Ghostty lock/build-input identity, measured test-surface identity, canonical compiler/linker/RC/SDK identity, verified immutable receipts, receipt-derived portable payload staging/import, and final output rendering. |
| `Modules/Build/DxUiDependency.psm1` | Exact-pin public restore, compiler/SDK/CRT identity, source cleanliness, bounded update notice with red/yellow severity, linked-module provenance and package sidecar verification. |
| `Modules/Build/DxUiUpdate.psm1` | Validated DxUi-main selection, atomic product-lock update, and caller-selected product validation. |
| `Modules/Build/MSBuildInvocation.psm1` | MSBuild launch planning and one diagnostic-header classifier for streaming colors and captured-log counts, including unnumbered warnings/errors. |
| `Modules/Build/ProcessStreaming.psm1` | Live stdout/stderr subprocess execution. |
| `Modules/Build/SanitizedEnvironment.psm1` | Canonical environment normalization, kill-on-close contained process launch, and local current-parent identity. |
| `Modules/Build/VcpkgInstallSafety.psm1` | Manifest floor/override validation, safe triplet validation, and confined vcpkg merge paths. |
| `Modules/Build/VcpkgToolIdentity.psm1` | Pinned vcpkg executable attestation. |
| `Modules/Build/Versioning.psm1` | Atomic locked version resolution/persistence and validated diagnostics plus pure explicit package-version construction. |
| `Modules/Packaging/PortablePackageSmoke.psm1` | Clean-extraction package smoke plus receipt-bound Ghostty identity, notice, and native load/free validation. |
| `Modules/Packaging/ReleaseArtifactPolicy.psm1` | Release artifact names, architectures, completeness, and checksums. |
| `Modules/Packaging/RuntimeDependencies.psm1` | Central runtime-dependency manifest resolution and validation. |
| `Modules/Packaging/WingetPrPublication.psm1` | Deterministic winget publication state and pull-request operations. |
| `Modules/Reporting/SourceLineMeasurement.psm1` | Shared solution source discovery for line-count commands. |
| `Modules/Reporting/TestRunReporting.psm1` | Shared archived-run discovery, summaries, and formatting. |
| `Modules/Testing/TestInventory.psm1` | Source-derived test registration plus stable-ID/policy/resource run-plan inventory. |
| `Modules/Testing/TestInvocation.psm1` | Self-test, Pester, and build invocation parameters plus result classification. |
| `Modules/Testing/TestQuarantine.psm1` | Quarantine policy parsing, status, and metadata-preserving repair planning. |
| `Modules/Testing/TestMachineResources.psm1` | Schema validation, non-secret capability digesting, and scoped environment projection for per-machine Connection Manager, SMB, MTP, and alternate local-root resources. |
| `Modules/Testing/TestRunArchive.psm1` | Archived evidence structure/content validation, area-schema dispatch, complete JSONL validation, and paired Terminal VT admission comparison. |
| `Modules/Testing/TestRunPlan.Common.psm1` | Stable plan entries, canonical validation-plan digests, path normalization, run identities, and process snapshots. |
| `Modules/Testing/TestRunPlan.psm1` | Canonical compatibility facade that preserves the established exported function surface. |
| `Modules/Testing/TestRunSummary.psm1` | Legacy v1 plus verified composite v2 summaries, verdict axes, the canonical 0..5 terminal-state mapping, case history, dashboard, JSONL, and step-summary rendering. |
| `Modules/Testing/TestSandbox.psm1` | Fail-closed exact `X:\RedSalamander.Perf` authorization, ownership markers, run contexts, cleanup planning, stale-run removal, and live-owner-consistent disk audits. |
| `Modules/Testing/TestSuitePlan.psm1` | Declarative stable-ID suite construction, governed Commands families, and selected-profile artifact-mutating Pester coverage. |
| `Modules/Testing/ValidationFingerprint.psm1` | Exact canonical JSON/digests, streaming read bounds, atomic replace or create-only record I/O, timed base/index/worktree snapshots, one-time receipt-closure binding, and versioned component-aware entry fingerprints. |
| `Modules/Testing/ValidationEvidence.psm1` | Schema-validated normalized terminal results, atomic hashed/quota-bound result artifacts, immutable content-addressed decisions/evidence/promotions, one exact direct-child run resolver and semantic promotion loader shared by Resume/summary/history/prune provenance, crash-safe artifact epochs/read barriers, interruption repair, and identity-bound pruning. |
| `Modules/Testing/ValidationImpact.psm1` | Governed impact-manifest/build-graph validation and deterministic shadow or enforced Affected decisions. |
| `Modules/Tooling/SpecInformationArchitecture.psm1` | Specification inventory, WIP indexing, consistency, link validation, and exact reviewed Done-path admission. |
| `Modules/Tooling/ToolInventory.psm1` | Tools manifest reconciliation, validation, and rendering. |

## Adding or changing a tool

Trusted validation commands:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -TestRoot D:\RedSalamander.Perf
.\Tools\Run-AllTests.ps1 -Suite Full -ExplainPlan
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Resume -ResumeFrom <run-id>
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Affected -ImpactBase <commit-ish>
.\Tools\Run-AllTests.ps1 -Suite Commands -CommandsFamily file-operations
```

Fresh is the default and final-closeout mode. Resume reuses exact promoted evidence only.
Affected and Commands-family runs are iteration surfaces whose repository verdict is
`NOT_EVALUATED`; Commands families remain non-cacheable while they share the monolithic
application payload.

Keep stable public paths thin. Put reusable implementation in a purpose-named
module directory, and preserve an existing path as a compatibility shim when a
known build, CI, documentation, or user-facing consumer depends on it.

In the same change:

1. Add or update the manifest entry, including an explicit `prerequisites` array.
   Do not rely on the Pester directory rule for an unusual lifecycle or owner,
   and never copy a legacy `namingPolicy` onto a new command.
2. Add complete comment-based help to public commands, including side effects and
   generated outputs.
3. Add focused Pester coverage and name it in `tests`; an active entry must not
   rely only on a sealed historical test. Use `[]` only when no useful automated
   contract exists.
4. Update this concise map when the top-level surface changes.
5. Give every `compatibility` record a concrete `supportPolicy`; `replacement`,
   when non-null, is one exact existing repository-relative `Tools/` path, never
   prose.
6. Update the owning normative specification when behavior, public parameters,
   lifecycle, or consumer contracts change.
7. Run `Get-ToolInventory.ps1 -FailOnFindings` and the focused Pester suite before
   the canonical repository test gate.

Do not delete a no-caller script based only on text search: manual support commands,
asset generators, MSBuild entry points, and frozen evidence can have intentional
non-code consumers. Mark the lifecycle and known audience first, then follow the
retirement contract in the tooling-governance specification.
