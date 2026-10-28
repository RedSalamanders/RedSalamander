# Repository tooling

Hosted CI/Full and nightly shuffle UI runs wrap native validation in
`Invoke-RSWithHostedTestDesktop` from `Modules/Testing/HostedTestDesktop.psm1`.
It selects a supported sufficient display mode, reports work-area/session/foreground facts,
and restores the original mode on failure or success. It refuses local invocation; native
focus warnings and leases remain required. See the [self-test contract](../Specs/Testing/Testing_SelfTests.md).

The nightly shuffle consumer restores the exact pinned DxUi source before importing its portable
attestation; receipt use still verifies that source identity and every attested output byte.
Protected native actions check child exit status at the call site and throw on failure. A workflow
reports success only after that check and display restoration complete.

RedMonitor shares one vendored client header/license. Verify-RedMonitorClient.ps1 checks local hashes
without Git/network/viewer work. Update-RedMonitorClient.ps1 -Commit FULL_SHA is the explicit dependency
update operation. Restore-RedMonitor.ps1 is a compatibility alias for local verification only.


The I42 Copy/Move selection-convergence completed record has exact-path admission in the specification
inventory; a date-shifted sibling remains rejected. Admission does not complete the WIP plan or qualify
its implementation. The reviewed storage cases below retain their separate no-activation contract.

The five 2026-10-19 File Operations UI/screenshot records and the
[five-stash review](../Specs/Plans/Done/StashReview_2026-10-19.md) also have individual
exact-path admission. Focused tests reject all six date-shifted siblings and an
unreviewed companion; admission does not establish runtime qualification.

The completed October 20 waiting-card record also has exact-file admission.
Its focused guard rejects a date-shifted sibling and an unreviewed companion.

The I49 agent/skill remediation record has exact-file admission under `Specs/Plans/Done/`.
Its guard rejects a date-shifted sibling and an unreviewed companion. This records guidance and
package-check closeout; native runtime qualification remains outside that review's scope.

`Restore-DxUi.ps1 -Platform x64 -Configuration Debug -CheckUpdates` restores the exact
`Dependencies/DxUi.lock.json` pin before an IDE build. Root `build.ps1` performs this automatically.
The source and compiler/SDK/CRT-isolated outputs stay under `.build/dependencies/DxUi`; a dirty
managed source rejects build/test reuse. The optional newer-version notice never edits the lock; a validated candidate
is yellow and prints the normal update command plus its `-UpdateOnly` alternative. A candidate without a successful
completed DxUi validation is red and must remain pinned.
New managed clones enable Git long paths for metadata creation and persist them locally before checkout; user/global Git settings are unchanged.
`Update-DxUi.ps1` makes that branch update: it accepts only a current `main` commit with successful completed DxUi
CI, changes the lock, and runs the Full suite. `Update-DxUi.ps1 -UpdateOnly` skips the local suite only when matching
product validation already passed in another environment; neither form creates a commit.
Restore evaluates the selected configuration's first-party compiler host and preserves explicit
`PreferredToolArchitecture` overrides. The library project reference forwards that same host.
See `Specs/Build/Build_Toolchain.md` for the manual upgrade and per-module provenance contract.


`Tools/` contains the repository's supported developer commands, build and test
helpers, release policy, and preserved terminal-engine evaluation evidence. Start
here instead of guessing from a filename.

The exact FileOps-only `FileOps_ReviewRemediationContracts`, `FileOps_CrossRootLocalCopy`,
`FileOps_SmallFileCopyCost`, `FileOps_SelectionShapeLocal` and `FileOps_SelectionShapeSmb` storage/route regressions run without activation or a desktop
lease, alongside `FileOps_RepeatDecisionContracts` and the two reviewed folder-progress
witnesses. Broad filters, other suites and proof injections retain normal desktop requirements.

PR delegation uses `Get-ScopedPrWorkflowDigest` over reusable workflows, local action
contents, tooling, scripts and build/test policy inputs. Changing a delegated test runner
invalidates coverage even when `ci.yml` is unchanged. Recompute the reviewed manifest digest
after execution-policy changes; missing or unresolved local dependencies retain local work.

The PR build-policy guard checks the formatter job's PowerShell installer wiring and rejects
Python/pip bootstrap only within that job. Formatter acquisition integrity is covered separately
by `Tests/ClangFormat.Tests.ps1`; other build dependencies retain their own prerequisites.

The machine-readable source of truth is [`tool-inventory.json`](tool-inventory.json).
Ghostty lock rejection retains field-specific schema diagnostics when PowerShell's `Test-Json`
returns false without nested error records. The fallback uses the shipped schema engine with
reference fetching disabled and cannot admit a rejected lock.
Artifact action upgrades also update `ReleaseWorkflowPolicy.Tests.ps1` pin
assertions and verify build, release, nightly, and qualification consumers together.
The exact-byte historical Gate0 workflow is excluded from Dependabot updates. Its older artifact
action pins are accepted only while the entire workflow matches its existing sealed digest;
current workflows continue to require their reviewed current pins.
The specification inventory admits the three exact reviewed I19 Done paths and
the exact I26 File Operations UI and [I27 safety-remediation](../Specs/Plans/Done/UI_FileOperationsSafetyRemediation_2026-10-01.md) plan paths; focused tests reject unreviewed
siblings. Admission does not establish completion of implementation or qualification
gates. ARM64 and ASan qualification remains separately indexed under WIP and is
not inferred from plan closeout.
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

Shared-runtime cleanup tests execute both common MSBuild cleanup and the installed native
C++ cleanup task, including production Utility transitions; shared DLLs survive while owned
targets and private intermediates remain cleanable.

Current source-contract checks complement native behavior tests. The history-popup guard
requires a readable popup state and keyboard selection within the existing bounded wait;
separate diagnostic bookkeeping does not relax that success condition.
Clipboard-admission guards ignore formatting whitespace while retaining the admission-sequence
and no-terminal-consumption assertions. WinGet minimum-Windows checks resolve Monitor source
through the verified clean external pin; compatibility project directories contain no native owner.
The durable-state guard matches the complete publication call and overwrite policy with
formatting-tolerant whitespace; native durable-store tests establish runtime behavior.
ViewerSpace guards check the workspace action/layout menu order, unique static access keys,
embedded-only unknown-share resource and all 13 published Space frames with origin receipts. Its monotonic ancestor-total source guard complements the gated native tile test.
The `ViewerSpace nonactivating desktop observations` Pester group bans physical input
and activation writes in the UIA observer/fixtures, enforces process-filtered event
coverage with a bounded private drain, and keeps ambient-interference/negative-control
cases in the noninteractive native inventory. Native cases exercise those contracts;
source checks alone are not desktop behavior qualification.
File Operations guards bind bridge cleanup permission to its constructor argument order and
keep selected-root and descendant Keep Both grants separate, alongside native isolation and
matching-descendant reuse witnesses. Loopback guard refusals must fail; the reviewed reason-only
throughput result is admitted only with non-passing, non-skipping defaults and runner failure routing.

## Where to look

Foreground tests reuse the native `DirectedSelfTestInputWarning` in
`Tests/TestSupport/`: a large warning centered on the test app appears three
seconds before input and remains during the foreground suite or isolated case.
Pause keyboard and mouse input until it closes. Background lanes retain their
no-activation guards. The unified runner continues to own the desktop lease.
FileOperations children hold that lease for their foreground and UIA checks
while retaining no activation; only the exact documentation interaction case
opts into focus. Two reviewed exceptions run without the desktop lease under
`-Suite FileOps -NonInteractive`: exact, case-sensitive
`-CaseFilter DiscoveryScope_FolderLiveProgress` and
`-CaseFilter DiscoveryScope_MultipleFoldersLiveProgress`. Their real-copy
witnesses use direct owned-window hooks and WGC captures with the existing
no-activation guard. Broad suites, families, prefixes, wildcards, multiple-case
filters, casing near-misses and proof injections retain the lease. The lease
serializes governed tests, not independent apps.
Noninteractive planning and execution report the pending entry count even when
it is zero; an empty pending list must not prevent admitted cases from running.

The runner gives each FileOperations child process a fresh `LOCALAPPDATA` below
the selected run's `scratch/fileops-profile` directory, including classification
retries. Fake MTP journals remain inside that child's lifetime and cannot collide
with the caller's journal history. The parent environment is unchanged. Use this
runner for FileOperations qualification; a direct executable diagnostic must also
provide an isolated local-data directory under the marked TestSandbox root.

After a Full artifact writer, the runner byte-verifies the rebuilt profile and
refreshes its cached receipt binding before it computes downstream fingerprints.
The receipt ID and build-artifact component therefore describe the same verified
receipt; exact Resume can reconstruct that terminal artifact epoch.

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

`Run-AllTests.ps1` keeps active execution beneath the marked `runs/<runId>` tree.
After children and owned-disk teardown finish, it verifies and renames the complete
tree into `evidence/completed-runs/<runId>` on the same volume. Summary/dashboard
links use that final location. Failed, skipped and interrupted witnesses are retained;
startup archives abandoned parseable runs instead of deleting their evidence.
Create-only records under `evidence/run-archive-records/<runId>` preserve the original
path mapping and verified inventory for recovery. Unsafe links, unresolved disks,
changed data, live owners and collisions retain the tree for review. Explicit cleanup
is still separate. Archive retention does not qualify a test or bypass exact Resume.

| Goal | Command |
|---|---|
| Build the repository | `./build.ps1` (repository root) |
| Run the canonical tests | `./Tools/Run-AllTests.ps1 -Suite CI` |
| Run the full closeout suite | `./Tools/Run-AllTests.ps1 -Suite Full` |
| Measure the opt-in I26 presentation workload | `./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_IntegratedPresentationPerf` (see `Specs/Testing/Testing_PerformanceValidation.md`; paired evidence required) |
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
| `Clean-TestSandbox.ps1` | Public maintenance command | Lists or explicitly removes one exact run directory beneath marked `X:\RedSalamander.Perf`; it never touches historical outside-root locations. Apply protects live owners using the runner's PID-reuse policy. Preview/WhatIf/success exit 0; a failed removal returns `Status: Failed`, warns and exits 1. Review manually named runs and verify evidence preservation first. Removal never follows reparse points: junctions and directory symlinks are unlinked, and a run containing a volume mount point is refused whole (an administrator must remove the mount). A run that records a test-owned VHDX receipt is torn down first when elevated; unelevated, a VHDX kept detached after a failed FileOps suite is deleted with the run and any other receipt refuses the whole run with the receipt path and the exact recovery commands; a dry run warns. |
| `CompareTestRuns.ps1` | Public report command | Compares two archived self-test runs. |
| `Evaluate-TerminalEngines.ps1` | Historical Gate-0 command | Reproduces/authenticates the sealed multi-candidate selection; not a product build. |
| `Fold-Changelog.ps1` | Public changelog command | Immediately folds pending fragments; supports read-only -Check and -WhatIf preview. |
| `GenerateDebugIcons.ps1` | Internal asset generator | Regenerates checked-in debug-badged icon resources. |
| `Get-GhosttyTerminalRuntimeStatus.ps1` | Public read-only runtime status command | Validates the lock without network, observes official upstream/advisories, or freezes one explicit candidate; writes only below `.build`. |
| `Get-SpecInventory.ps1` | Public governance command | Validates specs/WIP/links and writes derived reports below `.build`. |
| `Get-TerminalDependencyStatus.ps1` | Dormant generic lifecycle command | Explicit-lock validation and opt-in advisory checks for a reviewed future upgrade. |
| `Get-TestInventory.ps1` | Public discovery command | Derives test registrations, projects, Pester cases, and run-plan coverage. |
| `Get-ToolInventory.ps1` | Public governance command | Validates and renders `tool-inventory.json`. |
| `Get-VisibleTypographyAudit.ps1` | Public UI source audit | Finds visible typography paths outside the intended DirectWrite model. |
| `Install-ClangFormat.ps1` | Public formatter acquisition | Checksum-pinned native formatter download/offline extraction, cache verification and failure-preserving publication. |
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
| `Restore-RedMonitor.ps1` | Compatibility command | Read-only local client verification; no restore or build. |
| `Update-DxUi.ps1` | Public dependency-update command | Advances the DxUi lock only to a successfully validated main commit, then runs or explicitly skips Full product validation. |
| `Verify-RedMonitorClient.ps1` | Public client verifier | Read-only local header and MIT hash verification. |
| `Update-RedMonitorClient.ps1` | Explicit dependency update | Fetches only the header and MIT notice at one full commit; writes their lock. |
| `Restore-TerminalEngineInputs.ps1` | Dormant generic lifecycle command | Downloads or verifies inputs only for an explicit reviewed generic lock. |
| `Run-AllTests.ps1` | Public canonical test command | CI/Full execution with profile-scoped build inputs and all test-generated local data beneath exact `X:\RedSalamander.Perf` on a selected fixed local drive. First initialization and alternate-volume FileOps roots require explicit opt-in; `path.local.alternate` can pin cross-volume coverage to one exact second-drive root. Optional per-machine resources come from `config\machine-resources.json` beneath the primary root. |
| `Run-MtpLiveCloseout.ps1` | Public opt-in device command | Runs MTP/PTP live-device closeout smoke and archives all local evidence beneath the selected `X:\RedSalamander.Perf` run. |
| `SanitizedEnvironment.ps1` | Internal compatibility loader | Preserves the frozen Terminal tooling path; current callers import `Modules/Build/SanitizedEnvironment.psm1`. |
| `Show-PerfRuns.ps1` | Public performance report command | Lists, compares, and trends archived performance evidence; the FolderView preset separates `folder.sort_product_paint_us`, legacy `folder.sort_toggle_us` and inclusive writer attribution `folder.sort_jsonl_write_us`. |
| `TerminalEngineEvaluator.psm1` | Historical Gate-0 module | Sealed candidate collection/finalization implementation. |
| `TerminalEvidence.psm1` | Historical evidence module | Authenticates sealed terminal evaluation records. |
| `Test-TerminalEngineRound4Closeout.ps1` | Historical closeout command | Reauthenticates the frozen Round-4 selection evidence. |
| `Test-Changelog.ps1` | Public read-only validator | Checks fragment names/content and the changelog; -RequireFolded enforces handoff. |
| `Test-ChangelogWorkflow.ps1` | Internal tooling fixture entrypoint | Runs shared behavior fixtures in the existing marked sandbox through the Changelog Pester adapter. |
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
| `Modules/Documentation/Changelog.psm1` | Shared Changes validation, deterministic merge, staged publication and retry-safe cleanup. |
| `Modules/Build/ArtifactOperationLock.psm1` | Profile artifact, platform-scoped vcpkg-root, and packaging coordination with native immediate-child-authenticated delegation, scoped contamination, path-qualified residual-process policy, reparse-safe guarded output paths, and same-directory atomic file publication. |
| `Modules/Build/BuildProjectSelection.psm1` | Solution project selection and transitive dependency resolution. |
| `Modules/Build/BuildEvidence.psm1` | Manifest-derived artifact/runtime closure, legacy/current vcpkg app-local path normalization, complete Ghostty lock/build-input identity, measured test-surface identity, canonical compiler/linker/RC/SDK identity, verified immutable receipts, receipt-derived portable payload staging/import, and final output rendering. |
| `Modules/Build/ClangFormat.psm1` | Native formatter package/cache validation, exclusive cache lease and guarded atomic executable publication without Python. |
| `Modules/Build/DxUiDependency.psm1` | Exact-pin public restore (Git long paths, sparse checkout, failed-restore cleanup), API-revision gate, compiler/SDK/CRT identity, source cleanliness, bounded update notice with red/yellow severity, linked-module provenance and package sidecar verification. |
| `Modules/Build/RedMonitorDependency.psm1` | Resolves consumer-owned source-contract paths and rejects obsolete viewer ownership. |
| `Modules/Build/DxUiUpdate.psm1` | Validated DxUi-main selection, atomic product-lock update, and caller-selected product validation. |
| `Modules/Build/MSBuildInvocation.psm1` | MSBuild launch planning and one diagnostic-header classifier for streaming colors and captured-log counts, including unnumbered warnings/errors. |
| `Modules/Build/ProcessStreaming.psm1` | Live stdout/stderr subprocess execution. |
| `Modules/Build/SanitizedEnvironment.psm1` | Canonical environment normalization with distinct null-removal and present-empty child overrides, kill-on-close contained process launch (the job also applies `SEM_NOGPFAULTERRORBOX`, so a crash ends the process without a Windows Error Reporting dialog), and local current-parent identity. |
| `Modules/Build/VcpkgInstallSafety.psm1` | Manifest floor/override validation, safe triplet validation, confined vcpkg merge paths, and the default-MSVC-toolset overlay-triplet pin shared by `vcpkg-install.ps1` and `build.ps1`. |
| `Modules/Build/VcpkgToolIdentity.psm1` | Pinned vcpkg executable attestation. |
| `Modules/Build/Versioning.psm1` | Atomic locked version resolution/persistence and validated diagnostics plus pure explicit package-version construction. |
| `Modules/Packaging/PortablePackageSmoke.psm1` | Clean-extraction package smoke plus receipt-bound Ghostty identity, notice, and native load/free validation. |
| `Modules/Packaging/ReleaseArtifactPolicy.psm1` | Release artifact names, architectures, completeness, checksums, and the fail-closed distributable test-surface check used by every packager. |
| `Modules/Packaging/RuntimeDependencies.psm1` | Central runtime-dependency manifest resolution and validation. |
| `Modules/Packaging/WingetPrPublication.psm1` | Deterministic winget publication state and pull-request operations. |
| `Modules/Reporting/SourceLineMeasurement.psm1` | Shared solution source discovery for line-count commands. |
| `Modules/Reporting/TestRunReporting.psm1` | Shared archived-run discovery, summaries, and formatting. |
| `Modules/Testing/HostedTestDesktop.psm1` | Hosted-only supported display selection, work-area/session/foreground diagnostics, and restoration around native qualification. |
| `Modules/Testing/TestInventory.psm1` | Source-derived test registration plus stable-ID/policy/resource run-plan inventory. |
| `Modules/Testing/ScopedTesting.psm1` | Native test naming, reviewed candidate PR coverage, and latest-attempt scheduled source-monitor decisions. |
| `Modules/Testing/TestInvocation.psm1` | Self-test, Pester, and build invocation parameters plus result classification. |
| `Modules/Testing/TestQuarantine.psm1` | Quarantine policy parsing, status, and metadata-preserving repair planning. |
| `Modules/Testing/TestMachineResources.psm1` | Schema validation, non-secret capability digesting, and scoped environment projection for per-machine Connection Manager, SMB, MTP, and alternate local-root resources. |
| `Modules/Testing/TestRunArchive.psm1` | Archived evidence structure/content validation, area-schema dispatch, complete JSONL validation, and paired Terminal VT admission comparison. |
| `Modules/Testing/TestRunPlan.Common.psm1` | Stable plan entries, canonical validation-plan digests, path normalization, run identities, and process snapshots. |
| `Modules/Testing/TestRunPlan.psm1` | Canonical compatibility facade that preserves the established exported function surface. |
| `Modules/Testing/TestRunSummary.psm1` | Legacy v1 plus verified composite v2 summaries, verdict axes, the canonical 0..5 terminal-state mapping, case history, dashboard, JSONL, and step-summary rendering. |
| `Modules/Testing/TestSandbox.psm1` | Fail-closed exact `X:\RedSalamander.Perf` authorization, ownership markers, fresh uncached reparse-component checks, `New-RSTestSandboxLink` (the only governed creator of test junctions, symbolic links and hard links, refusing any link that would live or point outside one marked sandbox), run contexts, cleanup planning, stale-run removal that never follows reparse points (links are unlinked; a run with a volume mount point is retained whole), the per-run test-owned VHDX fixture (elevated provisioning that proves the attached disk is the blank file-backed VHDX before formatting and the volume is its own before mounting, a durable receipt before mounting, reverse-order teardown that can keep a detached VHDX as evidence, and receipt-aware cleanup that retains a run when not elevated unless its disk is already detached), and live-owner-consistent disk audits. |
| `Modules/Testing/TestSuitePlan.psm1` | Declarative stable-ID suite construction, governed Commands families, the two exact no-activation folder-progress witnesses, and selected-profile artifact-mutating Pester coverage. |
| `Modules/Testing/ValidationFingerprint.psm1` | Exact canonical JSON/digests, streaming read bounds, atomic replace or create-only record I/O, timed base/index/worktree snapshots, one-time receipt-closure binding, and versioned component-aware entry fingerprints. |
| `Modules/Testing/ValidationEvidence.psm1` | Schema-validated normalized terminal results, atomic hashed/quota-bound result artifacts, immutable content-addressed decisions/evidence/promotions, one exact direct-child run resolver and semantic promotion loader shared by Resume/summary/history/prune provenance, crash-safe artifact epochs/read barriers, interruption repair, and identity-bound pruning. |
| `Modules/Testing/ValidationImpact.psm1` | Governed impact-manifest/build-graph validation and deterministic shadow or enforced Affected decisions. |
| `Modules/Tooling/SpecInformationArchitecture.psm1` | Specification inventory, WIP indexing, consistency, canonical inline/fenced Markdown masking for links/headings, and exact reviewed Done-plan/capture-companion admission, including the 2026-10-16 File Operations card hierarchy record, with unreviewed sibling rejection. |
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

## Published documentation capture guard

Application foreground requests use `Common::RequestForegroundWindow`: no-activation
scopes decline them before entering User32, while ordinary interactive calls retain
their Win32 behavior. Source contracts ban direct bypasses in application production
code and recognize this wrapper in all physical-input/activation-write bans.

The exact Commands case `cmd_product_documentation_gallery` captures static
Search, Preferences, plugins, connections and Command Palette examples from
private synthetic settings. It uses the runner's no-activation guard and owned
foreground observer without global input or a desktop lease. Broader filters,
case variants and proof injections retain the ordinary Commands classification.
`ViewerPETests TestViewerDocumentationGallery --no-activate` captures generated
viewer fixtures; `TestViewerSpaceLifecycleCaptures --no-activate` covers Space.
These focused captures do not qualify keyboard interaction or a Full run.

`Tests/DocumentationDriftContracts.Tests.ps1` runs through the current tooling
profile. It binds the ten public File Operations PNGs to adjacent native capture
provenance, matching SHA-256, invented fixture paths, no activation and a recorded
complete-frame review. The guard rejects unreviewed byte replacement; it does not
perform visual privacy review for the publisher. The authoritative rules are in
[Testing_SelfTests](../Specs/Testing/Testing_SelfTests.md#documentation-screenshot-scenarios).

The exact static product gallery also covers eleven utility/shortcut forms, three
shell-free native Terminal layouts and five loopback/fake-provider browse scenes
(45 scenes in total). `ViewerPETests TestViewerVlcDocumentationPlayback --no-activate`
adds real playback/pause of generated silent media; publication requires a present
identified VLC runtime and passing decoder observations, never a skip.

## Changes workflow

See [Changes/README.md](../Changes/README.md). Test-Changelog.ps1 validates without writes; Fold-Changelog.ps1 folds before PrePush/final handoff and after every integration, even one fragment. Use -WhatIf to preview and -Check as the final read-only gate. Test-ChangelogWorkflow.ps1 runs shared fixtures through the current Changelog Pester adapter in the existing marked sandbox. The implementation lives under Modules/Documentation/Changelog.psm1.

Verify-RedMonitorClient.ps1 is also called by direct Common MSBuild; Update-RedMonitorClient.ps1 downloads exactly the public header and MIT notice at an explicit full commit.

Hosted MSVC provisioning uses the existing contained streaming-process module for
the signed GUI installer, preserving argument boundaries and actual exit status.
It waits at most 45 minutes, retains stdout/stderr under RUNNER_TEMP, accepts only
0 or reboot-required 3010, then verifies the required compiler/ASan/SDK/LLVM tools.
A real GUI fixture exercises delayed completion, Unicode/spaced arguments and
0/3010/failure exits without installing software; local -VerifyOnly remains read-only.
Workflow-selected VS18 MSBuild may live in a custom/temporary installation root;
build discovery checks its real binary version/identity and standard host layout,
then applies the required compiler/SDK policy independently of the parent name.


Tools/Install-ClangFormat.ps1 installs the Build/clang-format.json-pinned native formatter using PowerShell 7, without Python. It validates package and executable bytes, supports an offline -WheelPath and nonmutating -WhatIf, and preserves the existing executable on failure. The cache stays beneath .build/format.

Formatter output paths reuse the guarded path and atomic publication helpers in
`Modules/Build/ArtifactOperationLock.psm1`. Existing reparse components are rejected through
the full output chain, including `.build`, cache files and executable parents. Relative
`-WheelPath` inputs resolve once against the PowerShell location; hashing and extraction use
the same open stream. An exclusive `.build/format/.install.lock` lease covers cache mutation
and publication, fails immediately on contention and leaves an empty lock file after closing.
`-WhatIf` creates neither the cache nor its lease file.
