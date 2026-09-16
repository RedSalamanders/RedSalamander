# SelfTest Specification

## Overview

RedSalamander ships three debug self-test suites:
- `--compare-selftest`
- `--commands-selftest`
- `--fileops-selftest`

This document defines the normative result contract shared by those suites.

FileOperations qualification runs through `Tools/Run-AllTests.ps1`. Each contained
FileOperations process, including an isolated or shuffled classification retry,
gets a fresh `LOCALAPPDATA` directory under the current run's scratch root. Fake MTP
recovery journals must not reuse or mutate the caller's profile. Keep that directory
stable until the contained child exits, because a provider worker may outlive an
individual case. The runner never changes the parent's environment. A direct native
diagnostic must supply the same isolated local-data setup; it is not Full evidence.

Compare-only startup is headless only when its selected cases do not require host
services. `CompareDirectoriesSelfTest::RequiresHostServices` applies the same
declared-case filter semantics as `ListCases`: saved-profile remote cases (except
UNC-only SMB), remote FTP fallback search, credential storage/Hello cases and
cloud connection-configuration gates use the existing application/FolderWindow
host. They retain `--selftest-no-activate` and the ordinary owned host teardown;
this is not authorization to bypass credential, TLS, SSH, or Hello policy.
`host_services_startup_selection` tests exact, prefix, comma-set, case-insensitive,
unknown and pure/headless selections; `host_services_startup_context` requires a
unique missing-profile metadata lookup to reach `ERROR_NOT_FOUND`, with no secret
or network access. Missing host context in configured remote credential checks
and selected credential/cloud-gate cases is a test failure, not an
environment/capability skip; the selection case also tests this classification.

Related documents:
- `Specs/Testing/Testing_TestCoverage.md` — coverage policy and required risk map
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/TestRuns/README.md`
- `Tests/README.md` — central test infrastructure index
- `Tools/Run-AllTests.ps1` — unified test runner with summary reporting

`--fileops-selftest` must cover the verification setting-to-confirmation snapshot chain, Preserve/Skip
link handling and semantic tree-Move retargeting, destination containment, name/limit feasibility,
capability-v2 provider truth, identity replacement races, publication/delete uncertainty,
Recycle escalation, clipboard consumption on admission, discovery Skip/non-starvation, and
cancellation at bind/read/write/verify/delete boundaries. `FollowTargets` cases are retired and must
not be registered under a new name.

Capability-v2 contract tests must reject an empty document, a missing required operation member, or
a non-boolean operation member. Every complete provider document contains both `operations.move`
and `operations.nativeMove`: `move` admits the provider Move entry point for the queried path/profile,
while `nativeMove` independently proves the qualified native strategy. Provider-matrix tests must
exercise them independently so a legacy Move bit cannot silently become native-strategy authority.
The same matrix must exercise the pure Move selector: Native wins without querying Copy; a proven
Copy pair plus Move pair, exact bound/conditional Delete, exclusive/conditional publication, and
both binding interfaces selects Managed; a missing destructive fact selects Copy-only; and a
missing Copy pair rejects. Executable Local cases cover exact regular-file and nested-directory
Managed cleanup plus a pre-existing external writer that forces a successful Copy-only/source-kept
downgrade before source deletion can be armed.
Every complete provider document also contains Boolean `operations.createDirectory` and positive
`names.maxComponentUtf16`. File Operations provider-matrix coverage queries the concrete parent and
candidate with `FILESYSTEM_CREATE_DIRECTORY`, proves Local and Dummy admission, rejects read-only
7-Zip, empty `rootId`, Local device/`GLOBALROOT` namespaces, and invalid provider-native component
separators, and verifies that only a qualified builtin Local candidate may use the Win32 fallback.
S3 debug contract coverage proves conditional zero-byte trailing-`/` marker creation plus
empty-prefix Copy/Move publication and source-marker cleanup; S3 Table remains non-advertised and
unsupported for Create Directory.
S3 multipart coverage also creates four simultaneous known-size 4-KiB writers and requires an
aggregate reserve no greater than 16 KiB. MTP fake-backend coverage includes
`mtp_public_writer_and_reader_memory_is_bounded`: a 12-MiB public write/read must use staged-file
publication, zero private payload buffering, at most 4 MiB for upload/relay, at most 8 MiB for
verification, and one reusable reader request state of at most 8 MiB. A failed ordinary seek must
not poison the reader; only a watchdog-abandoned device command may make it terminal.
Managed-cleanup cases also pause after destination publication and prove the retained Local binding
blocks a pathname rename/replacement, inject one known Sharing Violation and require the central
Retry action to reissue only conditional Delete, and inject an unknown mutation outcome that leaves
source plus destination, returns indeterminate, and creates no whole-item Retry prompt. Pure outcome
tests distinguish Removed, Retained, Indeterminate, and provider-contract contradiction.

## Trusted incremental validation

[`Operation Startrail`](../Plans/Done/Operation_Startrail_TrustedIncrementalValidationAndResumableFullEvidence_2026-08-13.md)
introduced trusted affected-test selection and resumable Full evidence. Fresh remains
the default and the mandatory CI/nightly/release/final-closeout path; explicit local
Resume and Affected have the narrower authority defined below.

Build evidence Phase 1 is active. `Run-AllTests.ps1 -SkipBuild` MUST reject path-only
reuse and require a schema-valid, test-enabled full-solution receipt for the exact source
snapshot, platform, and configuration. Before any test process launches, it MUST verify
the receipt pointer against its immutable content-addressed record, rehash every attested
output, and require every file-backed executable or plugin in the selected CI/Full plan
to be present in the declared runtime closure. The same verification runs after an
artifact-writer re-attestation. CI/nightly consumers MUST verify and locally rebind the producer's
same-workflow portable attestation after all staging and before invoking `-SkipBuild`.

A family filter, case filter, repeat/shuffle variant, or capability-reduced run is not a
Full closeout even if a legacy artifact or display label contains `Full`. A raw passing
entry, interrupted entry, provisional entry, or logical result from a changed monolithic
`RedSalamander.exe` is not independently reusable Full evidence.

Resume reuse requires an immutable promoted entry, a
matching post-entry workspace snapshot, a compatible artifact epoch and receipt, an
unchanged outcome/skip policy, and aggregate schema-v2 support in every consumer.
Artifact-mutating tests must declare their artifact read/write sets and may not silently
invalidate earlier evidence. Commands-family reuse after a changed monolithic executable
requires a proved independent logical artifact; otherwise the affected Commands surface
must run comprehensively before Full can be green.

`-Suite Full -ValidationMode Resume -ResumeFrom <origin>` reuses only exact-artifact
promoted evidence from that direct origin and executes every non-cacheable or mismatched
entry. `-Suite Full -ValidationMode Affected` executes only the governed affected set,
does not reuse results, and reports repository `NOT_EVALUATED`. Invalid combinations fail
before artifact locking/building. CI/nightly/release workflows explicitly pass Fresh.

### Stable validation-plan entries (Phase 2 active)

Every `Get-RSTestRunPlan` entry MUST have a stable lowercase semantic `Id` independent
of its display `Name`. Duplicate or missing IDs fail plan construction. The canonical
`validation-plan.json` binds the requested suite/mode/coverage scope and, for every
entry, kind, repository-relative executable/script and working directory, ordered
arguments, result contract, impact/input sets, artifact read/write/produced sets,
receipt invalidation scopes, cache policy, environment inputs, coverage/repeat/order
identity, outcome/skip policy, resource classes, duration key, and desktop requirement.
Display names are serialized for operators but excluded from `plan_digest`; every
behavior, coverage, resource, or policy field is included.

Fresh runner execution order and commands come from these same entry objects.
Quarantine repair projections MUST copy all entry metadata while changing only the
reviewed focused-case arguments. `ToolsPesterTests` excludes
`RequiresBuildToolchain`; Full runs those cases through the separate
`ToolsPesterBuildToolchain` entry, which declares profile artifact mutation and receipt
invalidation. The runner passes its selected platform/configuration through isolated
process environment inputs, and the deployment test must delete, build, lock, and inspect
only that exact profile. No hidden build-toolchain Pester case may execute inside a
read-only entry, and no Release or ARM64 plan may silently mutate x64 Debug.

### Workspace barriers and entry fingerprints (Phase 3 active)

Before launching a test, the runner MUST publish a stable initial workspace snapshot
and one schema-valid fingerprint for every planned entry. It MUST acquire and publish a
post-entry snapshot after every normal entry and quarantine repair, plus a final
snapshot. A mismatch from the initial snapshot is a blocking source mutation even when
the test itself passed. Runner evidence under `.build`, archived evidence under
`Specs/TestRuns`, Git metadata, and logs are excluded to prevent self-invalidation;
ordinary untracked source and tooling files remain included.

Each fingerprint MUST bind the plan/entry semantics, workspace and input identities,
verified build receipt/runtime closure, current artifact epoch, reviewed runner/schema
dependencies, classification and outcome policy, quarantine state, and every declared
non-secret capability input. Unknown declared environment inputs fail planning. Fresh
still executes every entry; a fingerprint match authorizes reuse only through the
explicit Resume promotion/receipt verification path.

### Impact explanation and Affected mode

`-ExplainPlan` MUST construct the same stable suite plan, acquire the canonical workspace
snapshot, validate the governed build/import/runtime graph, and publish the deterministic
affected-set decision without acquiring the artifact-operation lock, building, or
launching a validation entry. An optional `-ImpactBase` resolves locally through
merge-base and never fetches. Normal Fresh runs publish a shadow decision after their
fingerprints and still execute all entries. Affected derives its filtered execution plan
from that complete Full decision and cannot report repository health. The console and JSON MUST distinguish
affected from unchanged entries and state `advisory_build_action` without suggesting
that unchanged means reusable. That field explains impact scope only; it does not dispatch,
suppress, or reorder the runner's independently governed build phase.

### Resumability checkpoints and exact reuse

CI and Full Fresh execution MUST persist the plan/decision and checkpoint each stable
entry before launch. A passing entry remains provisional until terminal artifacts are
preserved and its post-source/receipt/epoch barrier publishes an immutable promotion.
The aggregate run is passed only when all required entries are promoted and no source,
quarantine, or runner failure remains. Interruption leaves inspectable state. A later
explicit Resume may reuse only entries that already reached immutable promoted state and
still satisfy every exact identity.

The artifact-mutating Tools Pester lane MUST be first in the canonical Full plan and use
PowerShell 7. Before launch, `Run-AllTests.ps1` MUST verify the current epoch barrier,
suspend the selected receipt pointer, persist a `mutating` epoch, and invalidate every
intersecting promotion. Artifact readers MUST reject mutating, interrupted, epoch-mismatched,
or receipt-mismatched state. After the writer completes, the runner MUST rebuild/re-attest
the same full-solution profile, byte-verify that receipt, publish a create-only
content-addressed epoch decision, atomically switch run state back to `stable`, and
recompute the impact decision and every downstream entry fingerprint before any artifact
reader starts. A crash after suspension remains receipt-less and repairs to interrupted;
it cannot expose earlier evidence as current. Rebuilt MSVC binaries and PDBs are not
required to reproduce the pre-entry byte
identity. Re-attestation MUST apply the same suite-owned build environment as the initial
build, including `RSBuildEnableTests=true` for Full, even when the run entered through
`-SkipBuild`; a byte-valid receipt for a production-only Monitor binary is not compatible
with the Full executable selftest plan. Full's initial build and post-writer
full-profile re-attestation MUST use embedded compiler debug information while retaining
final linker PDB output. This avoids the process-global MSVC compiler-PDB service during
a promoted full Rebuild, so disjoint worktrees remain concurrent without C2471
contention. Ordinary `build.ps1` and CI builds retain ProgramDatabase by default. The
MSVC 14.51 `/Z7` + `/MP` internal-compiler failure for
`Common/SearchTextHelpers.cpp` is contained by compiling only that translation unit
without `/MP`; the Full profile keeps project and remaining compilation parallelism. The
runner MUST restore the caller's process
environment after the scoped re-attestation. The writer uses `ExactArtifact`: an exact Resume may reuse its promoted
post-entry receipt/epoch evidence, but if the writer executes again all earlier reuse
decisions are conservatively replaced with current-epoch execution decisions. General
non-isolated writers likewise increment the artifact epoch and invalidate intersecting
earlier promotions.

Nested `build.ps1` and in-process Pester boundaries may alter PowerShell's loaded-module
scope. `Run-AllTests.ps1` MUST re-import its Startrail fingerprint, impact, evidence, and
build-evidence command surface after each such boundary before snapshotting, re-attesting,
or publishing evidence. Shared evidence modules MUST NOT force-reload a dependency they
do not own, because removing that dependency can also remove commands imported by the
runner. Focused source-contract tests guard both sides of this boundary.

## Result Contract

### Case coverage

Every declared self-test case must produce exactly one case result in the suite `results.json` during
normal single-pass execution. When `--selftest-repeat=N` is supplied, each matched case must produce
one result per requested repeat attempt and each repeated result must include a one-based
`repeat_index`.

Allowed statuses are:
- `passed`
- `failed`
- `skipped`
- `crashed`

There must be no declared case with a missing status.

### Skip semantics

- `skipped` is the correct result when a declared case cannot run because a required precondition is absent.
- Every skipped case must carry a human-readable reason.
- Every skipped JSON row must carry `skip_class`. The shared emitter classifies ordinary
  reviewed precondition skips as `declared-capability`, list-only rows as
  `static-declaration`, and runner backfill/early-stop rows as `unexpected`.
- Skipping is part of normal suite behavior only when the entry's `outcome_contract`
  explicitly allows the case or its skip class. All-skipped, unclassified, forbidden,
  or otherwise unallowed skip sets fail the normalized terminal result.

### Setup failure behavior

If a suite encounters a fatal setup failure before all declared cases can run:
- the setup failure itself must be recorded explicitly,
- unreached declared cases must still be emitted as `skipped`,
- each backfilled skipped case must explain that it was not executed because of suite setup failure.

## Conditional Coverage Rules

Conditional coverage must remain part of the suite membership. Preconditions change execution status, not case existence.

Examples:
- remote smoke tests skip when required connection profiles, secrets, or sandbox roots are absent,
- plugin-dependent tests skip when the plugin is not available,
- machine-dependent filesystem coverage such as ReFS skips when the required volume is not present.
- self-tests that run before the normal application window exists may skip host-owned connection UI paths only when the host returns `HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE)`; other secret, profile, and plugin-configuration HRESULTs remain real failures unless they are an explicitly documented precondition.

A case must not disappear from the suite just because the current machine lacks its prerequisites.

Environment variables may select alternate test inputs such as profile names, but they must not be used to remove a declared case from suite membership.

## Suite and Aggregated Results

- Suite `results.json` files must preserve per-case status and reason.
- Aggregated self-test results must count `passed`, `failed`, and `skipped` consistently with the suite artifacts.
- The runner must select an aggregate suite by exact suite name; the mere presence of a
  `cases` property in another suite is never a match. It must reconcile declared counts
  against case rows and reject missing/malformed JSON, no cases, wrong-suite results,
  invalid statuses, aggregate mismatch, or incomplete/all-skipped outcomes.
- Pester success requires a PassThru result with nonzero discovery, at least one pass,
  zero failures/pending/inconclusive results, and only explicitly allowed skips.
- Every entry kind must pass through `ValidationTerminalResult.schema.json`; evidence
  promotion binds that terminal-result digest and cannot infer success from process exit
  code, failure count, and classification alone.
- In-product self-test suites must use the shared result-emission helper for suite-level case insertion, status counts, and first-failure propagation. Suites may keep distinct execution models, such as `SelfTest::RunCase` or FileOperations phase-state execution, but summary emission must stay centralized.
- In-product self-test suites must flush suite `results.json` through the shared case-result emission path after every recorded case. A crash or timeout must not erase the identities and statuses of earlier completed cases.
- The unified runner must compare each executed in-product suite's result names with runner-native `--selftest-list-cases` output for the same suite/filter and fail on duplicate expected names, duplicate actual names, missing declared results, or unexpected extra results. When `--selftest-repeat=N` is active, duplicate actual names are valid only when every expected repeated case appears exactly `N` times. CompareDirectories may emit an extra `setup` result for explicit setup failure reporting.
- `Tools/Modules/Testing/TestInventory.psm1` is the source-derived manifest for static Commands
  registrations, Compare registrations, FileOperations active phases, Tools
  Pester files, native `Tests/*.vcxproj` surfaces, and canonical CI/Full run-plan
  entries with execution kinds. Native project names must have set equality with
  project-backed run-plan surfaces; `PerformanceTests2` is `CppUnitTest`, other
  native test projects are `Executable`, and script/self-test surfaces retain
  their declared kinds. `Tests/README.md` and `Testing_TestCoverage.md` describe
  coverage but must not own mutable current totals. Use
  `Tools/Get-TestInventory.ps1 -Format Json` and runner-native
  `--selftest-list-cases` output for live inventory instead of copying counts
  into documentation or Pester expectations.
- `Tools/TerminalEngine/HistoricalTerminalProfiles.json` seals the exact Gate0/Round4
  test bytes and identities. Current Terminal runtime upgrades must add or update
  assertions only in its `currentRuntimeTests` cohort; they must never edit a
  `historicalTests` file or its expected hash to reflect the active runtime.
- When `Tools/Run-AllTests.ps1` is invoked with a non-empty `-CaseFilter`, runner-native case listing must return at least one expected case for each executed self-test suite. A zero-expected, zero-actual filtered run is invalid evidence and must fail as result coverage drift.
- Runner-injected coverage failures must contribute to the effective suite status, displayed failure counts, runner aggregate `exit_code`, and process exit code even when the native self-test process exits successfully.
- Runner exits are typed: 0 PASSED, 1 FAILED, 2 NOT_EVALUATED, 3 invalid CLI/plan,
  4 blocked/missing attestation, and 5 corrupt evidence. Explain and Affected return 2,
  which CI/release consumers must reject as a completed repository gate.
- Checked-in archived runs may contain skipped cases; that is valid when the skip reason documents the missing precondition.
- `Tools/Run-AllTests.ps1` must tolerate both direct suite JSON (`commands_results.json`, `compare_results.json`, `fileops_results.json`) and aggregated run JSON (`selftest_run_results.json`) when summarizing archived results.
- `Tools/Run-AllTests.ps1` must also write a runner-owned `run-all-tests-results.json` artifact for every invocation. Multi-suite runs must preserve each suite's status counts, case names, failure reasons, skipped reasons, wrapper exit code, canonical `test_root`, and per-run `run_id` in that aggregate file so later suites cannot overwrite earlier evidence. The aggregate must also preserve a non-blocking `test_sandbox_audit` object that records the selected root, allowed run ids, unexpected sibling run directories, and unexpected direct children. Normal execution audits only the selected root and must not enumerate or mutate legacy locations outside it.
- `Tools/Run-AllTests.ps1` must also derive `run-all-tests-case-history.jsonl` and `run-all-tests-dashboard.md` from the aggregate summary. History rows must include `harness`, `case`, `duration_ms`, `status`, `reason`, `classification`, `seed`, `attempt`, and whether the row came from native case output or retry/shuffle evidence.
- Runner aggregate summaries must preserve blocking classification metadata: per-suite `classification`, `classification_reason`, `retry_attempts`, and top-level `classifications` counts for `flaky`, `regression`, `isolation_suspect`, and `unclassified_failure`.
- `Tools/Run-AllTests.ps1` must use exact `X:\RedSalamander.Perf` as `REDSALAMANDER_TEST_ROOT`, where `X:` is any selected fixed local drive. The default `X:` is the repository drive; `-TestRoot` or `REDSALAMANDER_TEST_ROOT` may select another fixed local drive but may not change the directory name or select a descendant. First creation requires `-AllowExternalTestRoot`, an existing non-empty unmarked directory must never be adopted, and later use requires the exact `.red-salamander-test-root` marker. Arbitrary absolute paths, drive roots, UNC roots, repository-local `.build\TestSandbox`, reparse-point components, alternate-data-stream syntax, and historical `RSPerf` names are unauthorized.
- All local test-generated scratch, process temp, logs, compact archives, dashboards, and validation evidence must stay beneath `X:\RedSalamander.Perf`. The root may contain only its ownership marker plus `runs`, `config`, and `evidence`; one invocation lives under `runs\<runId>`, and Startrail evidence lives under `evidence`. Build products and source fixtures may be read from the repository, but test execution must not write there. Promotion into checked-in `Specs\TestRuns` is a separate explicit review action, never an in-product self-test side effect.
- The runner sets `REDSALAMANDER_TEST_RUN_ID`, redirects process `TEMP` and `TMP` into the current run scratch tree, clears inherited `REDSALAMANDER_SELFTEST_ROOT`, and restores every caller environment value on exit. Native helpers write their existing `last_run` tree under `runs\<runId>\artifacts\selftest\last_run`. Invalid root configuration fails closed and must never fall through to `%LOCALAPPDATA%`, process temp, or repository-local scratch. Direct harness launches derive the same drive-root location and require its marker.
- Native self-test scratch files must be acquired through `SelfTest::AcquireTestSandbox(...)`. The helper writes under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\<suite>\<case>\`, separate from `artifacts\selftest\last_run`; a marked direct launch uses `runs\adhoc-<pid>` beneath the same root. Commands plugin-config, viewer perf/close, BatchRename, ShellCommands, and Compare dummy fixtures use this helper instead of fixed drive-root, repository, LocalAppData, or process-temp paths. Nested test processes must remain beneath the current run scratch tree and must not create transient direct children of the root.
- Standalone harnesses must consume `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID` and use the same `runs\<runId>\scratch\<harness>\<case>` layout. Direct launches derive exact `X:\RedSalamander.Perf` from the checkout drive and fail closed when its marker is absent; `.build\TestSandbox` is not a fallback. DxUi generated artifacts remain beneath the current run `artifacts\dxui`, and caller-selected output parameters must also pass the selected-root containment check before a test writes them.
- Runner preflight cleanup must recognize both canonical runner IDs (`<UTC>-<pid>-<guid>`) and direct-harness fallback IDs (`<harness>-<pid>-<tick>`). It may remove only non-current/non-allowed run directories whose parsed owner PID is no longer live, or whose directory predates the current process that reused that PID; actual live-owner and unparseable manual directories remain protected. The post-run disk audit must reuse this owner/PID-reuse classification: live-owner siblings are protected, dead-owner or PID-reused parseable siblings are stale issues, and unparseable non-allowed directories are unexpected issues. Process-start lookup must use a Windows process-metadata fallback when ordinary process APIs deny access, so protected system processes cannot accidentally shield an unrelated stale run after PID reuse.
- Tooling Pester tests that create scratch files, fake toolchains, or synthetic perf-run trees must use `New-RSTestSandboxScratchDirectory(...)` beneath the selected run. Pester uses the same `X:\RedSalamander.Perf` root as native suites; no auxiliary repository sandbox is permitted. The runner's `TEMP`/`TMP` projection keeps Pester `TestDrive` state inside the current run.
- Native Full CI imports Pester 3.4.0 explicitly, matching the current suites' assertion and setup semantics. The artifact-writing lane and subsequent re-attestation inherit the initial build's vcpkg imports, configuration, include/link paths and runtime path. Governed test evidence outside the checkout and repository build diagnostics use separate upload artifacts so neither root is silently omitted.
- FileOperations real cross-volume and filesystem-capability coverage (including NTFS WOF placeholders) use the only sanctioned second-local-drive root. It is disabled by default and may enumerate or mutate another fixed local volume only with `-AllowAlternateVolumeTestRoot`. An enabled `path.local.alternate` entry in the primary root's `config\machine-resources.json` restricts selection to that exact `<AltDrive>:\RedSalamander.Perf`; without the entry, the explicitly authorized run retains fixed-volume capability discovery. It allocates through `SelfTest::AcquireTestSandboxOnVolume(...)` under exact `<AltDrive>:\RedSalamander.Perf\runs\<runId>\scratch\<suite>\<case>`, validates the same marker, and rejects non-empty unmarked roots. No other direct child may be created on the alternate drive.
- FileOperations queue-state selftests must prove the queue state they assert. Tests that cancel a queued task must start the active holder first, wait until that task has entered operation and owns the active slot, then start the queued task, assert `IsWaitingInQueue()` before queued-cancel assertions, and fail fast if the queued task enters operation or completes before cancellation. Timeout diagnostics must include task ids, live/completed state, queue flags, progress, and HRESULTs for every task participating in the queue proof.
- Compare selftests that launch `RedSalamanderSearchService.exe --run-foreground` must create a JobObject with `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE`, start the child process suspended with `CREATE_NO_WINDOW`, assign it with `AssignProcessToJobObject`, and resume the main thread only after assignment succeeds. The selftest-owned child must not share the broad runner's console-control lifetime, and it must not survive a selftest harness crash or timeout and hold SQLite/WAL or scratch-root handles into the next run.
- Compare selftests that gate foreground SearchService readiness must poll the explicit service status for the intended pipe and must diagnose early child-process exit with exit code plus captured stdout/stderr. Query-performance and parity tests must establish service readiness before the first warmup query; the first query is not an acceptable readiness probe.
- Foreground SearchService readiness polling uses a scaled, bounded 30-second default. Freshly linked images and cold SQLite stores can legitimately spend more than 10 seconds in image verification and initialization before creating their unique pipe; successful warm starts still return immediately, and cases with a tighter semantic budget may pass an explicit override.
- Foreground SearchService fault-injection and first-query gates must use the process-aware status helper with child output capture. A timeout, broken pipe, or early exit must retain the child exit code and captured stdout/stderr instead of reporting only the client-side pipe HRESULT.
- Foreground SearchService broad-order cases that depend on a long-lived primary process, including cold-root refresh, duplicate-instance rejection, and local root-rejection fallback, must enable captured stdout/stderr and diagnose an early primary exit at the point where the process is expected to remain live.
- Compare selftests own foreground SearchService lifetime through the unique pipe: readiness requires a successful status response whose reported pipe matches that intended pipe, normal teardown sends the foreground-only graceful-shutdown request, and the child must exit within the case's bounded shutdown budget. `--max-requests` and incidental later status probes are not valid teardown mechanisms.
- Before child tests launch, `Tools\Run-AllTests.ps1` must sweep parseable sibling run directories whose names match `runs\<timestamp>-<pid>-<guid>` and whose owner PID is no longer live. The current run id, explicitly allowed run ids, live-PID sibling runs, and unparseable/manual directories must not be removed by this dead-PID sweeper. Removal failures must warn and return failed cleanup rows; they remain visible evidence and must not silently turn into a green disk proof.
- Test tooling never performs legacy cleanup outside `X:\RedSalamander.Perf`; `-ApplyLegacySandboxCleanup` is rejected. `Tools\Clean-TestSandbox.ps1` is dry-run/`ShouldProcess` protected and may resolve and remove only one explicitly named direct-child `runs\<runId>` directory beneath the selected marked root. Historical outside-root artifacts are never enumerated or mutated by test tooling.
- GitHub Actions pull requests run `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` on native x64 and ARM64 runners in Debug, Release and ASan Debug. Full includes the build-toolchain Pester lane, ProductUiTests, the separate in-product suites, ViewerPE interactive cases, FileSystemCurlTests, RedConfigureTests, PluginContractTests, SettingsSchemaTests, and CrashHandlingTests, plus RedSalamanderMonitorEtwLatency. Each native ASAN profile must prove detection before ordinary contracts. The smaller `-Suite CI` remains an explicit iteration command; it does not replace the native Full gate. Runner summaries and governed artifacts are uploaded on failure as well as success. Cross-compilation alone does not qualify native ARM64 runtime behavior.
- Runner preflight, artifact-mutating deployment tests and native Commands inventory must accept the same six-profile matrix. Inventory launches only the selected platform/configuration executable; no x64 Debug sibling fallback. Invalid profile names still fail before artifact mutation.
- `MonitorTest.exe --document-model-selftest` owns the exact empty/BOM/unterminated/terminated/CRLF/blank/Unicode/line-budget read-export-read matrix and immutable snapshot lifetime across append, eviction, clear, and replacement. The test-enabled Monitor Chrome case owns injected open/save thread-construction failure through the production admission helpers, exact retained-state output bytes, zero copied text/active-tail bytes, and the three-size/five-repeat Release scaling drill.
- `PluginContractTests` must require each built-in filesystem plugin's `RunDebugSelfTests` export in Debug configurations. A missing expected export is a test failure, not a silent pass. Configurations that intentionally omit debug exports must report an explicit skip; if an optional export is present, the harness still executes it and requires its reported failures to remain zero.
- ViewerText snapshot handlers, their backing counters and the embedded-preview bridge use `ENABLE_TESTS` consistently across Debug, Release and ASan Debug. A test-enabled Release must exercise the same text/hex/diff and preview assertions; `_DEBUG` must not silently remove the observations while the cases remain registered. Ordinary builds with tests disabled omit these hooks.
- Compare `content_pending_elided` must hold real content workers at a bounded synchronization gate while asserting the initial folder-level pending state, then release them and require all 200 equal comparisons to finish with no surfaced items or pending state. Queue notifications on the calling thread must remain unblocked. A fast completed initial snapshot is valid product behavior, so file counts, sleeps or scheduler timing cannot establish the pending-state fixture. Gate timeouts fail explicitly, and teardown must release the gate and join workers before destroying callback captures, including after failure.
- Synthetic FolderView fixtures must retain the backing names for every borrowed `FolderItem::displayName` until the view and its retained items are destroyed. Initializer-list name vectors are valid only for synchronous calls that do not retain the resulting items. Removal-focus ownership/selection fixtures must exercise the same assertions under ASan without substituting owning product fields or changing production behavior.
- `Phase6_LocalBandwidthThrottle` must measure whole-copy duration from the worker's published operation start, not the first UI observation. Its fixture deliberately defers the first observation by at least 500 ms while continuing to pump messages, so delayed polling cannot shorten the measured runtime. Preserve the 250 ms duration-lead allowance, sampled throughput checks and 250 ms cancellation limit; report observation delay separately from duration.
- The R4A02 queue-after and don't-start fixtures hold the predecessor Copy at the existing test-only publication pause until the Delete is observed waiting in the overlap queue or completing explicit cancellation. Provider latency alone must not establish that state. Preserve the final nonzero wait-count, ordering, source-preservation and cancellation assertions; release and disable the pause on failure and after the decision witness.
- A test-enabled Release selftest that asserts an info-level File Operations diagnostic must explicitly enable that diagnostic level for the case and restore the captured settings afterward. It must not wait for an info record while relying on Release's intentionally disabled default.
- FileSystemCurl selftest compilation and its exported selftest entry point use `ENABLE_TESTS`, so the same provider proof runs in Debug and an explicitly test-enabled Release build while an ordinary production Release omits the export. The provider proof must use bounded loopback FTP endpoints and the public cross-connection Move path to cover both a `226` upload that retains a strict prefix and a complete-upload control. It must assert the staged `STOR`/`SIZE`/cleanup command shape, source/final/staging byte state, `ERROR_PARTIAL_COPY` on the short case, success on the control, bounded teardown, and mutation sensitivity when the independent staged-size comparison is bypassed.
- Portable ZIP creation must expand the completed archive into a new directory, validate the shared runtime-dependency manifest, authenticate the packaged Ghostty identity/DLL/notices against the receipt-bound shipping profile and canonical lock, start the packaged app through `--help`, natively load/free the packaged Ghostty DLL when the target architecture is executable, and run the matching plugin harness there with `--package-smoke`. Package-smoke mode validates load/enumeration/exports/schemas/capabilities for every built-in plugin but deliberately excludes Debug-only selftests and runtime-refresh probes. ARM64 package execution may be skipped on an x64 host only after full file/dependency and Ghostty identity validation; native admitted-product qualification must execute it.
- An admitted Ghostty upgrade must also dispatch `.github/workflows/ghostty-runtime-product-qualification.yml` on native ARM64 Windows. The lane must bootstrap the exact executable-tool commit in `vcpkg-tool.json`, materialize `vcpkg.json`'s independent `builtin-baseline` as the ports/version registry root, and use the repository `vcpkg-install.ps1 -Platform ARM64` entrypoint before compiling, so a fresh runner cannot depend on ambient headers, packages, or a stale tool checkout's version database. The executable checkout must remain on the tool pin while `-VcpkgRoot` explicitly names the baseline worktree; `VCPKG_ROOT` is insufficient because vcpkg ignores it when the executable already resides in a valid checkout, and treating the two pins as one checkout is invalid. Before compilation, the lane must import the pinned checkout's `vcpkg.props` and `vcpkg.targets` process-locally, disable manifest installation during MSBuild, bind `VcpkgInstalledDir` to the repository-owned ARM64 install root, and provide the same installed include root directly to the compiler. This setup is required because a native hosted image need not include Visual Studio's vcpkg integration; user-wide integration and ambient package paths remain forbidden. One positive build number binds its test-enabled Debug build, canonical runtime schema/load validation, complete Terminal plugin contracts, representative embedded-terminal Commands filters, production Release build, and natively executed clean-extraction ARM64 portable smoke. Selecting `REDSALAMANDER_TEST_ROOT` is not the same as authorizing it: because this lane launches native harnesses directly instead of only through `Tools/Run-AllTests.ps1`, it must explicitly initialize the selected exact `X:\RedSalamander.Perf` root and its ownership marker through `Initialize-RSTestSandboxRoot -AllowExternalInitialization` before the first direct launch. A direct harness that acquires sandbox scratch fails closed against an unmarked root, so an uninitialized root turns the Terminal contracts step into a sandbox-acquisition failure rather than a contract result. That root is outside the workspace, so its evidence must be uploaded as its own artifact; mixing it into a workspace-rooted artifact path has no valid common artifact root and fails the upload. Each representative Commands invocation must use its own runner-owned `REDSALAMANDER_TEST_ROOT`, so a preceding filter's valid run directory cannot become a later filter's disk-audit residue. A pointer-follow assertion inside a native lifecycle filter must prove the foreground-owned success path when the hosted desktop grants foreground ownership and the fail-closed, focus-preserving path when it does not; child focus alone is not proof that the app root owns foreground. Its runtime identity artifact path is `.build/TerminalEngine/Product/ARM64/terminal-runtime-identity.json` and required evidence uploads fail when any named artifact is absent. The release workflow must require this reusable gate before artifact builds. Cross-compiled ARM64 build and package validation remains required but does not replace this native product gate.
- Terminal runtime lifecycle selftests must inject every engine-construction stage after runtime load and prove complete reverse-order rollback: no Ghostty object, worker, callback registration, PNG system-callback lease, or loaded runtime may survive. A second initialization attempt must fail without replacing the live engine. The production OSC 52 callback must be exercised through the admitted DLL and prove exactly one `BUSY` reply per callback, no inner engine retry while the terminal mutex is held, and payload coalescing while UI work is pending. Output-ready and Kitty-ready wakes must use typed posted-payload registry tokens carrying the session generation; teardown tests drain the registry, reject stale tokens, and prove a failed/closed post clears its pending gate. Snapshot fixtures must call `renderStateEndUpdate` before any state/cell traversal and distinguish successive styled generations. Reader-thread COM/WIC initialization failure must be observable while leaving text rendering available.
- Both reusable official build paths must set `VcpkgXUseBuiltInApplocalDeps=false` before invoking `build.ps1`. Build receipt publication consumes each project's vcpkg tlog as output evidence; the pinned PowerShell copier records copied destination paths under the selected output directory, while the built-in copier records package source paths that must remain rejected by the receipt boundary.
- Both the reusable build and admitted-product bootstrap must compare the positive
  `schema-version` in `scripts/vcpkg-tools.json` at the executable pin and ports
  baseline before checkout/bootstrap. The executable schema must be greater than
  or equal to the baseline schema. This compatibility preflight fails before
  dependency installation and directs a baseline bump to update
  `vcpkg-tool.json`; cache hits or ambient tools cannot bypass it.
- `.github/workflows/asan.yml` owns the scheduled and high-risk-path x64 `ASan Debug` plugin-contract lane. Before accepting a green harness run, it must invoke the opt-in `--asan-seed-heap-overflow` probe, require a nonzero result with an AddressSanitizer diagnostic, and then run the ordinary contracts successfully. `_DISABLE_STL_ANNOTATION` remains documented technical debt for non-ASan vcpkg dependencies; the seeded proof covers ordinary heap detection, not STL container annotations.
- `.github/workflows/nightly-flake.yml` owns the expensive scheduled/manual shuffle-repeat lane. It must stay separate from PR CI, build the Debug self-test artifact, run `Tools/Run-AllTests.ps1 -Suite All -SkipBuild -SelfTestRepeat 5 -SelfTestShuffleSeed <seed> -ClassifyFailures`, and upload the runner-owned `REDSALAMANDER_TEST_ROOT` tree as `selftest-artifacts-nightly-shuffle`.
- Failure classification is blocking evidence, not a pass policy. `FLAKY` means an entry failed and passed on retry; for in-product suite failures it requires the failed cases to pass under `--selftest-case` and then pass shuffle triage. `REGRESSION` means retry evidence failed again, including any shuffle-triage seed that reproduces the failure. `ISOLATION_SUSPECT` means an in-product suite failed and failed cases passed under `--selftest-case`, but shuffle/order triage evidence is missing. Shuffle triage attempts must preserve their `shuffle_seed` in `retry_attempts`. Focused subset runs with `-CaseFilter` must still use failed-case retry plus shuffle triage, preserving the original subset filter for shuffle attempts instead of falling back to a whole-entry retry. When a native suite implements seeded/repeated order by dispatching one exact case at a time, that exact-case replay remains suite/shuffle context for classifier proof; only the runner's failed-case retry is isolated evidence. All non-`PASSED` classifications must keep the aggregate run exit code non-zero until the underlying test is fixed, replaced, or explicitly tracked through the reviewed quarantine process.
- Before every Commands case dispatch, including the normal single-pass order and seeded/repeated exact-case order, the harness must close same-process non-main top-level windows, exit inherited native menu mode, clear test-owned menu-hover overrides and the visible theme-cycle overlay, verify the main window is enabled, reset shared pane/transient state, and reactivate it. It must then prove stable FolderView focus, pump a bounded directed-input/message quiet point, clear navigation and incremental-search transients again, and re-prove FolderView focus before dispatch. This second clear/focus proof is required because input and teardown completion from the preceding case can arrive after the initial activation. A failed isolation boundary must be recorded as that case's failure instead of running the fixture against inherited UI state. Fail-fast bookkeeping may skip this cleanup after the first recorded failure because remaining cases do not execute.
- A same-process Commands case that delivers confirmed `WM_ENDSESSION` may validate persistence without performing normal teardown, but its test-only cleanup must reset both the runtime-settings save owner and every committed-shutdown gate before later cases run. In particular, the main-window command gate and host prompt-shutdown state must be restored; leaving the window alive while retaining production shutdown state is invalid suite isolation.
- A broad command-dispatch smoke case MUST NOT invoke commands whose contract performs a whole-runtime reload from durable or machine-local settings, such as `cmd/app/rereadAssociations`, or begins provider/OS lifecycle preparation before user confirmation, such as `cmd/pane/disconnect`; those commands require dedicated deterministic cases. After a smoke-dispatched command can alter navigation or presentation state, the smoke fixture MUST restore its in-memory settings snapshot, both pane providers and paths, and each pane's live display mode, sort key/direction, and visibility options, then prove both current-folder enumerations have settled before dispatching the next command. Restoring only the settings object does not reapply already-mutated `FolderView` state, and matching path text alone is not quiescence evidence.
- A Commands case that asserts navigation-shell stability after changing a pane provider/path or focus MUST pump pending messages and observe a bounded quiet baseline across repeated samples before capturing refresh/path/focus counters. A single matching snapshot is not sufficient because queued provider and focus notifications can invalidate it immediately after capture. The Refresh-shell fixture captures its protected counters from the final quiet sample and dispatches Refresh immediately; it must not insert another focus helper/message pump between baseline capture and the command. Timeout diagnostics for this fixture and drive-menu focus restoration are formatted after the wait so their state describes the observed result.
- A fixture that recreates files outside the provider at a previously visited path MUST invalidate the pane's cached enumeration before asserting the new contents. Reuse `ForceRefreshPaneForCommandSelfTest`, then wait for the expected items; increasing the item wait cannot make a stale snapshot current. Repeated Find fixtures share this setup through `OpenFindWindowFromLocalPaneRoot`.
- Bidirectional tab-traversal cases MUST re-establish and verify the boundary control before testing the opposite-direction wrap; a prior wrap's debug focus snapshot is not by itself proof that native/DxUi traversal state is ready for the next wrap. Focus failures must report the actual logical target, native focus identity, current page/category, and live child-host invariants.
- UI redraw-churn cases MUST first prove that render activity settles and that renders are attributable to invalidations. Their hard cap MUST bound the invalidation source rather than impose a lower fixed render cap that can reject a settled, fully attributable transition under broad-process load.
- A UI selftest that observes producer-side completion before asserting a paint-owned transition or popup placement MUST pump messages and poll the final observable UI state within a scaled bounded deadline. Task completion, window visibility, or the first readable debug snapshot is not proof that auto-collapse, animation, or anchored placement has reached its settled geometry. Before activating a moving control, the test MUST observe stable client-space hit-target geometry across repeated samples and map it to screen coordinates immediately before activation; keeping the resulting popup open cannot make stale click or anchor coordinates authoritative. The wait must remain condition-based and must not be replaced by a fixed delay or timeout inflation. History-dropdown keyboard coverage must obtain both a readable live popup state and a keyboard selection inside its existing scaled three-second poll. Diagnostics may retain the query result separately, but must not accept window visibility or a readable state without keyboard selection as success. Focus/capture/cursor observations remain read-only and do not alter the input sequence, assertions or deadlines. Full-path ancestor coverage checks the ellipsis client rectangle and navigation-window position together before resolving its click; post-click ancestor navigation and popup cleanup remain required.
- A Commands case that closes and reopens a UI Automation-backed window in the same process MUST wait for the old HWND to close, release thread-held UI Automation client/cache objects, pump pending teardown, and reacquire the current window, page, host, and provider identities before the next UIA read or mutation. Reusing a provider, element, page HWND, or child fragment obtained before close is invalid evidence even when Win32 recycles the same numeric HWND.
- Preferences category navigation under broad-process load MUST require repeated stable snapshots of the requested category and ready page surface. The helper must reassert the requested category when queued initialization temporarily restores an earlier selection. Tests that measure category-tree render/invalidation stability MUST settle again after moving focus from the tree to a page control before capturing the protected baseline.
- UIA state-setting helpers used after window recreation MUST be bounded and idempotent: reacquire the current host on every attempt, restore the intended logical/native focus, perform the Value/Toggle/Selection mutation, and report success only after the requested state is observed through the current model or UIA surface. A transport-level success without observed state is not a pass. Do not increase the suite timeout to conceal stale provider identity.
- When a UIA failure appears only after earlier cases in the same Commands process, isolated exact-case success is diagnostic only. Reproduce the failure with its predecessor cluster in one process, repeat that cluster after the repair, then require a green canonical Commands run. The focused cluster and canonical run complement one another; neither may be replaced by timeout inflation or a standalone retry.
- A Commands case that closes and reopens a UI Automation-backed window in the same process MUST wait for the old HWND to close, release thread-held UI Automation client/cache objects, pump pending teardown, and reacquire the current window, page, host, and provider identities before the next UIA read or mutation. Reusing a provider, element, page HWND, or child fragment obtained before close is invalid evidence even when Win32 recycles the same numeric HWND.
- Preferences category navigation under broad-process load MUST require repeated stable snapshots of the requested category and ready page surface. The helper must reassert the requested category when queued initialization temporarily restores an earlier selection. Tests that measure category-tree render/invalidation stability MUST settle again after moving focus from the tree to a page control before capturing the protected baseline.
- UIA state-setting helpers used after window recreation MUST be bounded and idempotent: reacquire the current host on every attempt, restore the intended logical/native focus, perform the Value/Toggle/Selection mutation, and report success only after the requested state is observed through the current model or UIA surface. A transport-level success without observed state is not a pass. Do not increase the suite timeout to conceal stale provider identity.
- When a UIA failure appears only after earlier cases in the same Commands process, isolated exact-case success is diagnostic only. Reproduce the failure with its predecessor cluster in one process, repeat that cluster after the repair, then require a green canonical Commands run. The focused cluster and canonical run complement one another; neither may be replaced by timeout inflation or a standalone retry.
- The absence of `Tools/test-quarantine.jsonl` is the canonical no-quarantine state; do not keep an empty placeholder ledger. A reviewed temporary quarantine creates that default JSONL file or supplies an explicit file through `Run-AllTests.ps1 -QuarantinePath`, with one JSON object per line. Empty/blank lines in a present ledger are ignored. Every nonblank entry must include `harness`, `name`, `owner`, `opened`, `expires`, `issue`, `root_cause_hypothesis`, and `fix_or_replace_plan`; `opened` and `expires` must use `YYYY-MM-DD`, `expires` must not be in the past, and an entry must not span more than 30 days. Invalid entries and valid active entries both keep `Run-AllTests.ps1` red; quarantine is a blocking repair ledger, not a nonblocking suppression list. Active entries must match a runnable harness adapter name from the current runner plan, and self-test entries must also match that adapter's `--selftest-list-cases` inventory. Matching self-test entries run in a separate repair lane with `--selftest-case=<name>` after the main gate, standalone entries rerun their matching adapter entry, and `run-all-tests-results.json` must preserve `quarantine.repair_attempts`, `repair_attempt_count`, and `repair_reproduced_count`.
- If a self-test process exits early, crashes, or writes only partial artifacts, `Tools/Run-AllTests.ps1` must report the runner failure from the available JSON/trace data instead of failing its own parser.
- If a self-test crashes while a case is in flight, the SEH path must write a partial aggregate `last_run\results.json` that marks that in-flight case with status `crashed`; the runner must treat `crashed` as failure evidence.
- Debug self-test builds expose `--selftest-crash-case=NAME` solely for crash-signal proof. When the exact matching case starts, the harness raises `EXCEPTION_ACCESS_VIOLATION` after recording the in-flight case, so reviewers can verify partial `results.json` preservation without waiting for a real crash.
- Debug self-test builds expose `--selftest-flaky-proof-case=NAME` and `--selftest-order-proof-case=NAME` solely for classifier proof. The flaky proof hook fails the named case in suite context and lets exact-case plus shuffle-triage retries pass, proving the runner labels it blocking `FLAKY`. The order proof hook fails the named case in suite/shuffle context and lets exact-case retry pass, proving shuffle triage labels it blocking `REGRESSION` instead of `FLAKY`. Commands, CompareDirectories, and FileOperations must carry explicit classifier proof suite/shuffle context when their seeded/repeated execution plans replay exact case names internally.
- Debug self-test builds expose `--selftest-repeat=N` for repeat-in-process flake reproduction. `N`
  is clamped to `[1, 100]`, suite and aggregate JSON record `repeat_count`, and repeated case rows
  record `repeat_index`. Commands and CompareDirectories repeat through the shared `RunCase` path;
  FileOperations repeats by expanding the selected phase-state run plan and preserving each attempt
  as a distinct `repeat_index` row in aggregate results.
- Debug self-test builds expose `--selftest-shuffle=SEED` for reproducible seeded case-order
  exploration. Commands and CompareDirectories support explicit seeded case ordering through the
  shared `RunCase` path. FileOperations supports explicit seeded phase ordering by expanding the
  selected phase-state run plan into individual phase filters while preserving `Setup` and
  `Cleanup_RestorePluginConfig` rows in aggregate results. Runner forwarding sends shuffle seeds to
  all three in-product suites.
- `--selftest-list-cases` must not overwrite prior run artifacts. List mode must set the shared
  self-test options to `writeJsonSummary=false` before building inventory JSON.
- Fatal-error UI must never block a self-test lane. When `IsRunningAnySelfTest()` is true, `ShowFatalErrorDialog(...)` must append a self-test trace row, write diagnostic output, and return before entering any modal loop; the fatal caller remains responsible for returning a non-zero process exit code.

## Command-Line Safety

- `--selftest-timeout-multiplier=N` accepts only finite numeric values.
- Invalid or non-finite multiplier values are command-line errors.
- Finite values outside `[0.1, 100.0]` are clamped to that range with a diagnostic.
- Scaled nonzero timeouts must remain finite and at least 1 ms.
- In-product self-test waits that gate correctness must express their base budget
  through `SelfTest::ScaleTimeout(...)` or a helper that scales internally before
  creating the deadline. Fixed literal waits are allowed only as short polling
  slices after a scaled deadline has been established.
- PowerShell harnesses that need the `RedSalamander.exe` self-test exit code
  must use `Start-Process -Wait -PassThru` or `System.Diagnostics.Process`
  rather than relying on `$LASTEXITCODE` from direct invocation of the
  GUI-subsystem executable.
- PowerShell/Pester tests that launch build-toolchain child processes must use
  a finite timeout and must capture stdout/stderr to named log files. A hung
  child build must fail the test with those log paths instead of blocking
  `Run-AllTests.ps1 -Suite Full` indefinitely.
- The complete test-enabled solution build owned by `Run-AllTests.ps1 -Suite Full`
  must cap MSBuild at four workers. That build can otherwise fan out enough
  compiler, linker, resource-compiler, and vcpkg applocal child processes for
  unrelated tool processes to fail initialization on an interactive host.
  The artifact-only CI build retains its build-host worker policy.
- Builds and self-tests that use the same normalized repository root,
  configuration, and platform must be serialized for the complete artifact
  lifecycle. Different configurations, platforms, and worktrees use independent
  artifact locks; target, project, and suite labels are diagnostics and must not
  split a shared output profile into separate locks. Production callers must
  always supply both configuration and platform. Same-platform MSBuild phases
  additionally serialize on the configuration-neutral vcpkg dependency lock;
  tests and runner bookkeeping do not hold that shared-dependency lock. See
  `Specs/Build/Build_Toolchain.md` for the authoritative lock and process-evidence
  contract.
- `Tools\Run-AllTests.ps1` executes only the exact normalized
  `.build\<platform>\<configuration>\RedSalamander.exe` covered by its selected
  artifact profile. `-ExePath` is retained as a compatibility spelling for that
  same canonical path; a foreign repository, different profile, or arbitrary
  executable MUST fail before the runner acquires any artifact lock.
- `build.ps1` must never close or kill an independently launched application.
  When a process executable path exactly equals an output the selected build can
  replace, the build aborts with PID, exact path, and available command-line
  diagnostics and asks the operator to close it. A same-name process at another
  path or in another profile/worktree is unrelated and must not block the build.
- `build.ps1`, standalone `Tools\Invoke-SanitizedMsbuild.ps1`, deployment build
  tests, and the complete `Run-AllTests.ps1` lifecycle, including `-SkipBuild`,
  must hold the same per-profile exclusive lock using `FileShare.None`. The test
  lease starts before TestSandbox cleanup and ends after result archival and disk
  audit. Active ownership records PID, process start time, executable path, exact
  lock path, random token, operation, and scope. Balanced nesting is limited to
  the owning managed thread. A direct child may reuse ownership only when the
  synchronous native launcher creates it inside the kill-on-close Job Object,
  restricts inherited handles, and completes the reviewed one-use delegation
  handshake. Delegation setup, temporary child-environment mutation, process
  creation, and restoration share one exception-safe cleanup boundary on both
  PowerShell 7 and Windows PowerShell 5.1.
- Only test-plan entries marked `RequiresInteractiveDesktop` hold the
  session-wide `Local\RedSalamander.InteractiveTestRun.v1` mutex, and only for
  the entry's execution boundary. Foreground, desktop-global UI Automation,
  pointer, clipboard, prompt, and ETW-monitor entries must be marked; Pester and deterministic
  noninteractive script/process entries must remain outside the mutex. A
  quarantine repair attempt inherits the matched plan entry's marker. This
  repair attempt also inherits the stable ID, coverage/outcome policy, artifact
  sets, cache policy, environment inputs, order contract, resource classes, and
  duration key; focusing one case does not create a new ungoverned adapter. This
  serializes desktop-global work across worktrees without blocking unrelated
  builds, sandbox preparation, noninteractive tests, result archival, or cleanup.
  An abandoned mutex may be recovered with a warning, and the three-hour
  acquisition timeout fails explicitly.
- Desktop serialization and permission to take foreground activation are
  separate contracts. Every GUI-capable plan entry must explicitly select one
  activation class. Focus-independent coverage runs with the shared
  `RedSalamander::TestSupport::ScopedWindowActivationBlocker` held on its UI
  thread; the guard adds `WS_EX_NOACTIVATE` to newly created top-level windows
  and rejects CBT activation/focus transfers. Focus-sensitive coverage runs
  without that guard and must be marked `RequiresInteractiveDesktop`. A new GUI
  suite/case must not inherit an implicit activation default.
- The governed no-activation set is CompareDirectories, FileOperations, DxUi
  suites other than Menu and NativeTextInput, the ViewerPE noninteractive
  group, the ViewerSqlite noninteractive group, and the Monitor chrome selftest.
  The governed activation-required set is the broad Commands suite, DxUi Menu
  and NativeTextInput, ViewerPE keyboard/combo/prompt cases, and ViewerSqlite
  tab traversal. DxUi Accessibility also retains desktop serialization for its
  process-global UI Automation work but runs with no activation. Monitor ETW
  latency retains serialization for desktop-global ETW measurement while its
  window remains `WS_EX_NOACTIVATE` and is shown with `SW_SHOWNOACTIVATE`.
- Activation-required suites MUST reuse `Tests/TestSupport/DirectedSelfTestInputWarning.h`,
  the shared version of the existing Commands warning, inside the runner's
  interactive desktop lease (including repair attempts). Show one large,
  DPI-scaled, non-activating, click-through warning three seconds before input
  begins; keep it visible until the foreground suite or isolated native case
  ends. Center it on the owned test app/dialog, clamped to that monitor's work
  area; before an app window exists, use the primary work area. The text names
  both keyboard and mouse and states when input can resume. Nested directed-input
  guards borrow the suite warning. Do not add a second runner or test-family
  warning implementation. Failure to show the required surface fails execution
  before its input assertions. No-activation lanes and case listing remain silent.
  The Commands process starts its warning before startup can activate the app;
  its suite and input probes borrow that process-lifetime surface.
  Commands isolation excludes only the exact live, helper-owned warning HWND
  from product-window cleanup; unrelated STATIC/tool windows are not exempt.
- Commands remains activation-required as one broad-process contract: the
  mandatory pre-case isolation boundary reactivates the main window and proves
  stable FolderView focus before every dispatched case. It may be split only by
  introducing reviewed per-case interaction metadata without weakening that
  isolation proof.
- Script-launched build and test process trees must use a kill-on-close Job
  Object. Automatic termination is limited to descendants launched and contained
  by the current operation. Residual `MSBuild.exe`, `cl.exe`, or `link.exe`
  processes block only when boundary-aware command-line evidence names both this
  repository and the same profile output/intermediate path or matching MSBuild
  properties. A process name alone, an ambiguous path substring, or another
  profile/worktree is insufficient. Failure to enumerate process metadata is a
  blocking diagnostic, not permission to kill or guess.
- A non-empty abandoned lock or stale authenticated owner record contaminates
  only its configuration/platform profile. Tests, direct MSBuild, clean-only
  builds, incremental builds, and targeted rebuilds remain blocked for that
  profile. Only a successful serialized full-solution `build.ps1 -Rebuild` for
  the same configuration and platform clears its marker; clearing or repairing
  one profile must not affect another. The platform-shared vcpkg marker is
  configuration-neutral and clears only after a successful full-solution rebuild
  for that platform. Source/build inputs must not be edited
  while an operation that consumes them is active.

## Self-Test Roots and UI Navigation

- `REDSALAMANDER_TEST_ROOT` is the runner-owned sandbox base for all harnesses and
  must equal exact `X:\RedSalamander.Perf` on a fixed local drive. Relative paths,
  UNC paths, descendants, alternate names, and repository-local fallbacks are
  invalid. The runner also supplies `REDSALAMANDER_TEST_RUN_ID`, and native self-tests resolve their
  artifact root as
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\selftest`. The legacy
  `REDSALAMANDER_SELFTEST_ROOT` variable is compatibility-only for deliberate
  override launches, including the reviewed quarantine repair lane; normal
  runner execution clears inherited values before child tests start.
- First initialization requires `-AllowExternalTestRoot`; subsequent use requires
  the exact ownership marker. Filesystem-specific tests inspect the selected
  volume and skip when its capability is absent; the root contract itself is not
  restricted to NTFS. New evidence must not use unique drive-root siblings,
  source/worktree roots, process temp, LocalAppData, or historical `RSPerf` paths.
- In-product self-test scratch roots are not artifact roots. Use
  `SelfTest::AcquireTestSandbox(suite, caseName)` for temporary files and
  directories that a case creates while it runs. The helper creates
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\<suite>\<case>\` for normal
  runner execution and logs the acquired root in the shared self-test trace.
- Shared self-test filesystem helpers must preserve extended-length Windows path
  semantics for directory creation, binary fixture writes, existence probes,
  trace-file creation, and recursive cleanup. A case that deliberately creates a
  path beyond `MAX_PATH` must not fail first in generic fixture setup or cleanup;
  the product API named by the case must be the component under test.
- Self-test fixture directory and file names should keep on-disk path segments
  concise. Worktree-local absolute roots can already be long; tests should use
  variables, comments, and assertion text for readability rather than embedding
  long descriptive names into temporary filesystem paths.
- Adversarial filename fixtures that claim near-maximum filename coverage must
  budget for the full absolute self-test temp root plus every generated
  subdirectory before choosing the file-name component length. A component that
  is individually legal can still break the test when the total path exceeds the
  active Windows path budget.
- Column, grid, or viewport fixtures that depend on row counts must derive those
  counts from the same visible UI mode they assert. Horizontal scrollbars,
  display mode, DPI, and density can change the pane's visible height; a
  no-scroll probe row count is not enough evidence for a scrolled fixture.
- Commands UI tests that call `FolderWindow::SetFolderPath(...)` must wait for
  the requested pane path to settle before dispatching commands that consume
  the current pane roots.
- Commands UI tests that clear pane visibility/filter state and then navigate
  to a fixture folder must not let the reset refresh an old pane path race the
  fixture enumeration. Clear state without refreshing, or perform the reset
  after the fixture path is active and wait for the resulting enumeration/debug
  state before asserting empty, filtered, or watermark UI.
- Empty-folder and filter-watermark self-tests must also clear pane operation
  alerts, explicit empty-state messages, and background watermarks before
  asserting the built-in empty-folder placeholder. These surfaces intentionally
  outlive some folder navigation paths, so visibility/filter reset alone is not
  enough isolation evidence.
- Commands GUI self-tests that validate pointer clicks, keyboard focus, or
  live DxUi traversal must run in a foreground-capable desktop session. Hidden
  or background launches are not valid evidence for those cases because they
  can change focus routing, pane activation, and pointer targeting.
- Commands GUI self-tests that validate product routes based on
  `WindowFromPoint` or screen-coordinate hit testing must first bring the
  target dialog/window to the foreground/top of the z-order and verify the
  target point belongs to that window's root before sending synthetic wheel or
  pointer messages.
- Commands GUI self-tests must prefer UIA providers, retained debug hooks,
  message-pumped helpers, and direct target-window messages over desktop-global
  input. Any remaining `SetCursorPos`, `SendInput`, or global keyboard path
  must be justified by a product contract that samples real desktop input, must
  use the canonical shared `Tests/TestSupport/DirectedSelfTestInputWarning.h`
  for the movement/input window, and must restore the previous cursor/focus
  state on every exit path before the case returns. Test-family-local warning
  implementations are forbidden because they can drift in visibility, timing,
  and resource ownership.
- Commands GUI self-tests that observe DxUi popup menus must scope popup lookup
  and dismissal to the owning window when possible. A visible popup from another
  owner is not valid evidence for the current menu, and stale owned popups must
  be dismissed before opening a new popup for hover, keyboard, or paint-debug
  assertions.
- Commands GUI self-tests that navigate a native-menu-backed DxUi popup must
  resolve dynamic target indices after the root menu's initialization callback
  has completed. Directed foreground input must first reactivate the owning main
  window, and a navigation failure must report the refreshed target, observed
  keyboard/hover state, popup validity, and per-message delivery results.
- Commands GUI self-tests that mutate persisted or runtime settings, such as
  Connection Manager profiles, must snapshot and restore the touched settings
  before the case returns. Full-order keyboard or focus traversal coverage must
  seed deterministic data and must not depend on rows, profiles, or selection
  state left behind by earlier cases.
- A selftest that mutates fields in an optional settings block must materialize
  that block immediately before the mutation and restore its prior state when
  the fixture requires isolation. It must not use `optional::operator->` merely
  because an earlier case or retired persisted field used to keep the block
  engaged. The retired File Operations pre-calculation fields are not a reason
  to assume `Settings::fileOperations` is authoritative or permanently present.
- File Operations discovery coverage must prove one traversal feeds bounded execution and that a
  safe mutation may begin before traversal closure. `Phase5_DiscoverySingleTraversal` covers
  authoritative closure and the 256-record queue ceiling; `Fairstream_DiscoveryAheadOverlapsTransfer`
  proves mutation-before-close with exact discovered bytes/files; `Phase5_DiscoveryCancelLatencyLocal`
  proves bounded cancellation; `Phase5_DiscoverySkipContinues` proves the one-way mode switch on a
  Local traversal that is still producing (a provider that holds its whole plan, such as Dummy, closes
  before its first byte and leaves nothing to skip); and
  `Phase5_SwitchParallelToWaitDuringDiscovery` proves Skip does not bypass global Queue waiting.
  Popup/source-contract coverage must keep task, aggregate, compact-row, and taskbar progress plus ETA
  indeterminate while discovery is open and expose `Discovering as needed` after Skip. Tests must not
  restore retired pre-calculation threads, settings, callbacks, or `Calculating` task state.
- `FileOpsFamily_R4A19DiscoveryBaseline` owns the observation-fact, provider-control, and
  independent-volume fixtures. Its Local Copy and permanent-Delete phases must recreate bounded
  unique children beneath the acquired File Operations sandbox, set and restore the Local provider
  configuration, wait for typed task completion, verify exact bytes/files and destination/source
  state, and require the discovery service facts described in the performance specification. The
  Delete phase must also prove exactly one selected-root open-discovery mutation and zero provider
  progress callbacks; it must not manufacture a callback to satisfy the metric. The provider phase
  must rebuild a delayed/speed-limited Dummy tree and prove its direct route closes on the exact
  subtree before the first byte moves, create one test-only fake-MTP instance with ordinary
  representable child names, prove exact payloads and single-session serialization, and cancel while
  discovery and mutation are simultaneously live. The fixed-volume phase may run only
  when both C and D File Operations roots carry the ownership marker; it must compare isolated and
  concurrent copies and prove progress overlap. All provider configuration, fake-module lifetime,
  unique local children, and both volume roots are restored or removed through ordinary cleanup even
  after failure. A missing/unmarked alternate fixed volume is a recorded blocker, never a synthetic
  substitute or skip-green result.
- `cmd_pane_batchRename_collision_name_index_memory_gate` must build 65,536 destination names of
  240 UTF-16 units, prove case-folded and exact provider-key lookup semantics, and keep the retained
  collision-name index at or below 64 MiB without retaining a second raw-name set.
- Local permanent recursive Delete coverage must prove the same traversal supplies exact discovered
  and completed bytes, begins deletion before traversal closure, never retains more than one
  256-child mutation batch, and converges under Continue on error without repeatedly rediscovering a
  locked child. `Phase6_DeleteBytesMeaningful`, `Phase7_ParallelDeleteKnobs`, and
  `Fairstream_ParallelDeleteContinuesPastLockedChild` own those executable rows. The Commands source
  contract forbids restoring `FlattenDeleteDirectoryTree` or its 200,000-entry preflight pass.
- `Phase8_PerItemOrchestration` must prove that multiple selected Local leaf files contribute exact
  discovery bytes before their respective publications and that the task aggregate equals completed
  bytes at closure. Popup smoke coverage depends on that real leaf-discovery stream; it must not use
  a fabricated total merely to keep the card visible.
- File Operations identity coverage must exercise the host binding abstraction
  and the provider implementation together. The deterministic matrix includes
  local regular-file binding, hard-link equality, replacement-at-the-same-path,
  disappearance, no-follow junction identity, unsupported-provider results,
  malformed-success contract violations, object-ID/`IsSameObject` cross-checking,
  direct same-object versus hard-link alias, same-folder Move, destination-inside-source,
  destination junction ancestry, and source/destination swaps at final revalidation.
  The Fairstream alias case must prove the task is admitted without a filesystem preflight and is
  then stopped by the just-in-time no-follow guard before any aliased child is published. Path text,
  timestamps, size, and other basic metadata are never accepted as identity evidence in those cases.
- File Operations interlock coverage must prove the complete `ReadSource`/`WriteSource`/
  `PublishDestination` conflict matrix, classify Copy and Copy-only sources as reads and their
  destination as publication, permit only read/read overlap, and retain the existing exact
  endpoint/path-profile/authority-chain overlap checks for every pair containing a writer or
  publisher. The metric/source contract must keep role counts bounded per task rather than per
  descendant. `FileOps_ProviderCapabilityMatrix` must additionally prepare a 256-item local sibling
  selection, require exactly 256 read scopes plus 256 final-leaf publication scopes, count unique
  retained authority nodes across interned chains, and keep them within the selected-item count plus
  a small shared-ancestor allowance and the 1,024-node path-depth bound. Missing leaves retain an
  exact existing-ancestor anchor plus normalized suffix; same/ancestor suffixes overlap, distinct
  siblings do not, Indeterminate/contract failures reject before mutation, and Unsupported identity
  uses a conservative domain scope. The same case must compare two 256-scope sets of disjoint
  publication leaves, require a nonzero probe count below the 65,536 Cartesian bound without a false
  overlap, and emit the canonical `FileOps.Interlock.CheckUs`, `ScopeComparisons`, and
  `ActiveCandidates` evidence for that scan in test-enabled x64 Release. Separate conservative-scope
  coverage must prove the exact Cartesian fallback remains fail-safe.
- File Operations mutation-truth coverage must deterministically inject post-`CREATE_NEW`
  `FileIdInfo` Unsupported, all-zero file ID, rename/delete committed-but-reported-failed,
  publication failure, and owned-stage abort Unknown. The Local provider proof must show retained
  creation-handle authority still opens/aborts the exact stage and that zero IDs are never compared.
  `FileOps_ProviderCapabilityMatrix` and `Floodgate_LocalCopyOverwriteIsStaged` must prove final
  `NotPublished`, source `Retained`, item `Indeterminate`/996, independent stage `Unknown`, preserved
  original destination bytes, and one explicitly cleaned test artifact. Cleanup of that exact
  fixture-owned artifact uses a bounded retry because Windows can keep a successfully deleted name
  visible briefly while the exact path is already non-addressable. Final residue means an
  addressable stage path: an enumeration-only name whose exact `GetFileAttributesW` probe returns
  `ERROR_FILE_NOT_FOUND` or `ERROR_PATH_NOT_FOUND` is a delete-pending ghost, while every other
  probe/enumeration error and every addressable matching path remains a hard failure. Every fixture
  clears its environment hooks.
- Typed-result coverage must enforce one terminal store for retry-loop Cancel, Keep Both failure,
  requalification failure, TraversalLimit, verification/Continue-on-error, Skip, and stage cleanup.
  A later branch may not rewrite Published to NotPublished or lose a supplied bridge publication
  state. A cancellation or later failure after a committed publication must retain Published while
  reporting Canceled or Failed; mutation commitment alone cannot force Completed.
- Recycle escalation coverage must prove that a successful exact permanent delete replaces the stale
  Recycle failure for item completion, while retaining the original Recycle HRESULT only as historical
  prompt context. A Local recursive permanent-delete partial child failure must retain the selected
  root as a known non-commit even when already-admitted children were removed.
- The copy-merge popup fixture must inspect the live actionable conflict card before accepting its
  decision: `Waiting for your decision` and question/context are visible and unclipped, semantic
  prompt regions precede actions, loading is false, discovery/transfer indicators are absent, and
  the last-note region fits. Pure presentation coverage separately exercises metadata loading and
  open-discovery/aggregate cohort rebasing.
- Provider-path destination coverage must prove that ordinary transfers derive the leaf with the
  source path profile, join it with the destination profile, and preserve explicit mappings exactly.
  Execution and completion/cache/refresh consumers must use the same plan-owned resolver; source
  contracts reject a return to `std::filesystem` destination reconstruction or separator guessing.
- The cross-filesystem publication fixture must inject cryptographic stage-name entropy failure and
  prove failure before the first exclusive-writer call. PID, thread ID, clock, and counter fallback
  names are forbidden. Copy and Copy-only must share mandatory transferred-byte and committed-size
  validation; optional content verification remains plan-controlled and cannot authorize cleanup.
- `FileOps_ProviderCapabilityMatrix` must exercise Managed cleanup known-noncommit outcomes for
  sharing violation, access denied, path too long, and disk full with both Retry and Skip. Retry
  must issue only a second `DeleteIfUnchanged`; Skip must preserve the already-published bytes and
  source. Unknown outcome must remain one-shot and indeterminate. Source contracts must source-classify
  cleanup `ACCESS_DENIED` and forbid the retired host pathname/whole-tree cleanup walker.
- `Phase9_ConflictPrompt_KeepBothNestedCacheEligibility` must prove a cached top-level Keep Both
  grant never auto-applies to a descendant collision, while the descendant prompt still offers its
  own explicit Keep Both choice. Choosing Skip must retain the affected source ancestry without a
  second cleanup prompt.
- Typed conflict coverage must prove that file/file, read-only file/file, file/directory,
  directory/file, and exact destination-link collisions expose distinct action sets; Overwrite is
  absent for mismatches and links. A Local junction destination must carry a Junction-scoped
  Apply-to-all key and expose explicit Skip All, while a different source/destination kind pair or
  link kind cannot consume that cached decision.
- Phase 3 exact-receipt coverage must prove that Local Copy and Native Move Overwrite consume one
  retained destination authority, a selected-root Replace link receipt cannot overwrite a later
  colliding child, and a raced replacement object is never mutated. A provider without executable
  object binding must omit Overwrite/Replace actions while retaining non-destructive top-level Keep
  Both where supported. Folder-on-file and file-on-folder both remain typed mismatches.
- Deferred in-tree semantic-link publication that cannot map its retained target must enter the
  `TargetConflict` row, expose no mutation grant, and retain the source link plus dependent source
  ancestry after Skip. Coverage must separately exercise outside-root preservation, successful
  inside-root mapping, and an inside-root-unmappable sparse transform; the third case must never be
  reported as outside-root or publish source-root target spelling.
- Reparse-kind coverage must exercise at least one name-surrogate link and one known non-name-
  surrogate WOF/cloud/dedup-style object through Native qualification, bridge destination ensure,
  overwrite policy, and prompt refinement. Only the name surrogate may expose Replace link; unknown
  kind is non-destructive and one-shot.
- Permanent Delete replacement coverage must pause after confirmation/interlock binding, replace the
  pathname, and prove the substitute is preserved. Local must consume retained exact authority through
  `DeleteIfUnchanged`; unsupported unbound providers must fail before mutation. Recycle bulk behavior
  remains a separate shell-owned fixture.
- Permanent Delete Retry-cap coverage must inject two known non-commits at the retained
  `DeleteIfUnchanged` boundary. The first prompt offers Retry, the second omits it and reports the
  failed retry, Skip retains the exact source, and exactly two delete attempts occur without a
  pathname rebind. A selected regular-file exact delete must also prove that same-walk discovered
  size becomes completed bytes only after committed removal.
- Connection/profile performance overrides must prove they constrain admitted work without creating
  mutation authority. In particular, an unbound Dummy Permanent Delete remains rejected even when
  `deleteMaxConcurrency` is configured; provider-direct deletion is permitted only as explicit test
  teardown and is not evidence for host admission.
- The Local delete TOCTOU dir-to-junction fixture is enabled in every test-enabled configuration,
  including Release, and must prove that recursive provider cleanup does not follow a raced target.
- A nested selftest setup helper that records `Skip` must return setup failure/stop to its caller,
  and the caller must translate that stop back to a successful declared skip rather than an unreasoned
  failure; the truthy `Skip` return value is not resource-creation success. Release-only missing-export paths
  must never continue into null COM-interface use.
- A nested Local recursive collision that chooses Keep Both must create the sibling at the child
  location (`name (2)` in the original parent) and leave the selected top-level destination name
  unchanged. The test must fail if the provider reports Skip and the host instead retries the
  complete selected root. `Fairstream_MovedTreeKeepsLiteralLinks` also runs a parallel
  same-filesystem Local Copy with a forward junction targeting that colliding child; the copied
  junction keeps its stored (source) target text while the Keep Both sibling, the source tree, and
  the pre-existing blocker remain unchanged, and a 4,200-junction tree copies with no link ceiling.
- `FileOps_MoveMergeIntoExistingFolderSameVolume` must inspect the admitted immutable plans and
  prove that a Local regular-directory-on-existing-regular-directory Move is Managed while a
  destination-absent same-volume regular file remains Native. A destination-absent Local directory
  is Managed because Local cannot prove the Native semantic-link transform. The merge case must expose a colliding
  child, not the top-level folders, through the central `Exists` conflict surface and must complete
  after the explicit child decision without losing non-colliding destination children. When another
  writable fixed volume is available, the same case must submit a real host-admitted cross-volume
  Local Move, assert immutable strategy `Managed`, exact destination bytes, and removed source; the
  environment-gated absence of such a volume is reported as a skip, never simulated by injecting a
  plan strategy.
- `Fairstream_RerunCopyIdenticalCollisionPrompts` must prove that a byte-identical existing file
  still enters the typed Exists conflict surface and changes only after an explicit decision; no
  digest-derived silent Skip is accepted.
- `Fairstream_ManagedMovePreservesUncopiedNewFiles` must pause a real host-admitted Managed folder
  Move after selected-child publication and before exact source cleanup, create an unselected source
  child, and then require `S_FALSE` / `Copied; source kept`. The injected child remains only at the
  source, the selected child remains published and is removed from the source, and no cleanup
  conflict prompt is allowed. This deterministic pause is a race-correctness fixture, not a timing
  assumption.
- When another writable fixed volume is available,
  `FileOps.SelfTest.LocalProviderCrossVolumeMoveNoFallback` must call the Local provider Move entry
  point directly and prove it returns `ERROR_NOT_SAME_DEVICE`, retains the complete source, and
  publishes no destination. The corresponding host-admitted cross-volume scenario remains Managed;
  the direct provider route must never restore its retired copy/delete fallback.
- The `Cinderstar_LegacyWriter*` publication cases use a bound Local source and an unbound Dummy
  destination so admission is truthfully Copy-only. They exercise bounded final-path verification
  for the legacy atomic-final writer: one transient miss converges to `S_FALSE`, permanent miss and
  wrong size fail after publication, and cancellation interrupts retry backoff. Every outcome retains
  the source; no case may arm Managed cleanup or report a commit-size-proof fallback.
- Recursive File Operations coverage must exercise both host-bridge and Local-provider traversal.
  The wide bridge case records retained-entry, queued UTF-16 path-byte, ancestor/directory metadata,
  and walk-depth high-water and requires them to remain within 4,096 entries, 16 MiB, 8 MiB, and
  128 descendant directory edges below the selected root (root depth 0). A test-only reduced depth ceiling must stop a two-item task at the second item,
  preserve the first published destination, reject every descendant below the ceiling, and record
  the exact depth-limit hit even when normal item continuation could otherwise apply. The Local
  fixture must use a 96-level extended-length tree below the recursive ceiling, exceed `MAX_PATH`,
  exercise both `CopyItem` and the single-item `CopyItems` route, and retain directory timestamps only
  until post-order restoration. Commands
  source contracts reject any return of the host whole-tree traversal stack or the Local
  `createdDirectories` metadata list.
- Pack and Unpack delete-after-success GUI coverage must prove that the archive prompt's captured
  `ArchiveDeleteAfter` choice reaches a central permanent `DeletePlan`, that the archive disappears
  asynchronously through File Operations, and that no private pathname-delete helper remains.
- Serial and parallel File Operations fixtures must inject the same traversal-limit generation/status
  and require the same task-terminal result, cancellation cleanup, and prompt teardown. A source-
  contract guard must prevent either executor from dropping a terminal input used by the other.
- Local case-only rename coverage must inject entropy failure and a raced temporary sibling. Entropy
  failure leaves the original source name intact; the collision case completes through fresh entropy
  while the unowned colliding temporary object remains byte-for-byte untouched.
- `Floodgate_InlineF2WorkerQueueBypassAndInterlock` must prove that Inline Rename executes off the UI
  thread, bypasses global Queue waiting, still waits for an overlapping mutation owner, serializes
  both successful renames, and exposes bounded interlock-wait performance evidence. While the owner
  is paused immediately before mutation, every interlock scope must retain an exact bound authority
  chain and the observed chain must remain within the 1,024-node path-depth limit. A second round
  must cancel the waiter, prove that it never reaches the provider rename boundary, and leave its
  source name unchanged while the owning rename completes. Fake-clock rounds must prove clean-running
  reveal at exactly 500 ms, clean success without an active/completed card, immediate non-clean reveal,
  and once-visible-always-visible completion. The exact-destination round must accept one Overwrite
  decision backed by a rename/publication-capable bound receipt and complete conditionally; a second
  stable prompt after that grant is an immediate test failure, never a timeout or automatic retry.
- Commands cases that deliberately refresh a runtime plugin manager while
  module-backed work is in flight must treat deferred unload as case-owned
  state. After the work and owning UI are quiescent, the case must refresh the
  manager again and prove required built-in entries are loadable before it
  returns; unusable deferred placeholders must not leak into later cases.
- Tests for retained process-owned windows must obtain the HWND from an owning
  debug hook or another process-scoped lookup. A class-only `FindWindowW(...)`
  result is not valid teardown or identity evidence because another running app
  instance can own a window of the same class.
- Commands GUI fixtures that restore pane providers or paths must wait for the
  restored paths and their asynchronous enumerations to complete before the case
  returns. When a fixture changes the active pane, teardown must also restore and
  verify the original active-pane FolderView focus; scope-exit may remain only as
  the failure-path fallback, not the successful handoff boundary.
- Commands GUI self-tests that continue after activating a focused file must
  account for any transient top-level viewer, prompt, or tool window opened by
  that activation. The test must wait for or close that transient UI and
  re-establish the intended pane focus before validating the next keyboard mode
  or focus-sensitive command.
- Commands GUI self-tests that drive modal prompts from a worker thread must
  keep UI Automation reads bounded and must always leave a path that closes the
  prompt. A blocked UIA provider call must degrade to a failed assertion with a
  trace message, not leave the owner window disabled and the suite waiting
  forever.
- Commands navigation-shell cases that close an owned pane modal must require
  the actual FolderView HWND to regain focus, not only stable path, history,
  item, and selection state. Product prompt teardown must pair its synchronous
  best-effort focus restore with the owning window's deferred, foreground-safe
  folder-focus message; failure diagnostics must include retained folder focus,
  Win32 focus, active window, foreground window/process, and owner identities.
- Commands GUI self-tests running on an owning window thread must not call UI
  Automation client operations in-process on that same thread when the provider
  can marshal work back to the owner. They must run the UIA operation through a
  bounded worker while pumping the owning thread, and must use the shared
  timeout/cancellation/join path; detaching a timed-out UIA worker is forbidden.
- Commands UIA workers must use an `IUIAutomation2` client with connection and
  transaction timeouts shorter than the operation deadline. The declared total
  budget must reserve time after that deadline to request the worker stop and
  call `CoCancelCall(...)` for the registered COM call. A normal blocked-provider
  timeout therefore completes and returns a failed helper result inside the total
  budget; joining is allowed only after completion is observed, and a worker must
  never be detached into a later case. If Windows violates both its UIA client
  timeout and COM cancellation contracts, the harness must fail fast instead of
  performing an unbounded join. The deterministic blocked-operation case must
  ignore the stop token through the operation deadline, finish inside the reserved
  cancellation window, and prove the failed result returns within the declared
  total budget without requiring an unhealthy desktop provider.
- Commands self-tests that validate Win32 clipboard formats must treat the
  clipboard as desktop-global state: do not run clipboard GUI cases in
  parallel, use bounded `OpenClipboard` retries for reads, report the open owner
  and registered-format availability on failure, and allocate DROPFILES read
  buffers with room for the trailing null terminator.
- Commands GUI self-tests that dispatch pane commands through `WM_COMMAND` and
  depend on the focused pane must wait until the intended folder-view HWND is
  the actual focused folder view after transient UI cleanup. A single
  `SetActivePane`/`SetFocus` call is not durable evidence when queued focus
  restore or focus-loss messages may still run.
- Commands GUI self-tests that reactivate focus-exit modes such as integrated
  Quick Search must dispatch every initial and repeated activation through the
  real `WM_COMMAND` route, then reassert stable folder-view focus before sending
  synthetic `WM_CHAR` input. Direct command-method calls are model-level setup,
  not end-to-end command-routing evidence. `WM_KILLFOCUS` is a valid product exit
  path for those modes, so active debug state alone is not enough evidence that
  subsequent typed input is isolated from queued focus changes.
- Broad-order Quick Search integration must wait for both the pane path and the
  complete fixture item set with the same scaled 10-second settle budget before
  injecting input. Preferences retained-page navigation must likewise wait up
  to a scaled 5 seconds for the combined category, page title, resize, and scroll
  state after a real category click; these waits remain condition-based and
  return immediately when the observable state is ready.
- Commands GUI self-tests that call Win32 `SetFocus` must assert the resulting
  focused HWND, such as by polling `GetFocus()`, rather than comparing
  `SetFocus`'s return value with the target. `SetFocus` returns the previous
  focus HWND on success, so direct equality with the target is valid only when
  the target was already focused and is order-dependent test evidence.
- Standalone DxUiTests that assert real Win32 `GetFocus()`,
  `GetForegroundWindow()`, `GetGUIThreadInfo().hwndCaret`, or menu/popup
  foreground ownership must not issue unguarded focus/foreground probes or
  discarded `SetForegroundWindow(...)` calls in test bodies. Same-thread
  focus/caret assertions use `TryFocusDxUiTestWindow`; foreground-owned
  popup/menu assertions use `TryActivateDxUiTestWindow`. If the current session
  cannot provide the required interactive or foreground-capable desktop, the
  foreground-dependent test must print `SKIPPED:` with a reason and return.
  This is an environmental precondition, not quarantine or nonblocking flaky
  policy.
- `DxUiTests.exe --no-activate` is valid only with one explicitly selected
  focus-independent suite. Menu and NativeTextInput reject that switch. The
  default multi-suite developer run changes the host-window activation policy
  at each suite boundary, so only those two focus-sensitive suites may activate.
  ViewerPETests and ViewerSqliteTests likewise expose separate
  `--group=noninteractive` and `--group=interactive` entry points; exact
  focus-independent ViewerPE cases automatically install the guard, and a
  no-activation request for an activation-required case fails closed.
- ViewerSpace tooltip paint validation captures the frame immediately after its
  synchronous debug-show hook returns, before pumping queued hover dismissal.
  The witness retains nonempty tooltip state and an increased real paint count;
  a queued `WM_MOUSELEAVE` must not turn a completed paint into a false failure.
- Modal DxUI menu tests may use the test-only popup state probe to observe input
  during an unrelated owner-message flood. The modal loop must dispatch real
  mouse/keyboard input first, then state probes for its live popup chain, before
  ordinary owner traffic; otherwise the measurement can be starved even when
  the product input path is responsive.
- Native menu sessions opened from the keyboard must ignore an initial left- or
  right-button release that has no matching popup button-down. Menu regressions
  that inject such an orphaned release must remove any unexpected queued command
  before it can open a modal prompt, so a failed assertion cannot wedge the
  remaining broad selftest run.
- Self-tests that validate an external process by reading a marker file must
  wait for the expected marker content, not only for file existence. Process
  launch and shell redirection can create the file before the first line is
  flushed, especially in full-suite order.
- Self-tests that validate file-action macro expansion through an external
  marker must match the documented macro context. Macros expanded inside an
  `arguments` template are Windows command-line arguments and therefore include
  the quoting/escaping needed for safe argument boundaries.
- File-action selected-path lease lifecycle tests must poll to the bounded
  process/retention quiet point and retry transient filesystem status-query
  errors. A transient observation error is neither proof that the lease still
  exists nor proof that cleanup completed. Terminal failures must report the
  injected fault, exact lease path, final existence/error observation, and poll
  count so retention bugs can be distinguished from probe races.
- File-action selected-path format coverage must assert the complete UTF-16LE
  byte stream: BOM, original nonempty record order, exact CRLF after every
  record, Unicode and long records, duplicate preservation, and rejection of
  NUL/CR/LF before manifest creation or child launch. The registered creation,
  record, streaming, and child-parse cases are
  `file_action_selected_paths_creation_contract`,
  `file_action_selected_paths_record_contract`,
  `file_action_selected_paths_streaming_perf`, and
  `file_action_selected_paths_child_parse`. The representative SearchService
  child must parse the grammar and report `parsed:<count>`; an existence check,
  BOM-only check, or substring search is not consumer evidence.
- `file_action_selected_paths_identity_cleanup` must prove deletion of the exact
  live manifest and preservation of a same-path different-identity regular file,
  a same-path reparse point and its target, and both the old and replacement
  trees after the validated root is moved/replaced. Creation/root handles must be
  closed before launch so the representative child can read and deliberately
  replace the manifest; the replacement must survive lease cleanup.
- `file_action_selected_paths_recovery_bounds` must prove that only old regular
  manifests inside the validated private root are reclaimed. Fresh, locked/live,
  directory, reparse, shared `%TEMP%`-shaped, and out-of-root controls must
  survive. The case must exercise zero-time cutoff, a missing continuation
  cursor, entry and deletion cutoffs, and bounded multi-pass convergence while
  reporting inspected/deleted/skipped/error/bound/pass observations.
- `file_action_provider_record_rejected_in_command_route` must inject a
  provider-supplied CRLF record through the public viewer-command route and
  prove precise localized `ERROR_INVALID_DATA` feedback with no child process.
  `user_menu_populates_and_dispatches_configured_actions` must likewise send
  the real User Menu manifest to the exact child parser. Helper-only plan tests
  cannot replace these command-route boundaries.
- Preferences page-specific self-tests should navigate to their setup page by
  named category selection or the documented visible root-row helper. They must
  not encode `End` plus a fixed number of `Up` keys as a shortcut for a page,
  because expanded child nodes and visible root-order changes make that setup
  order-dependent. Dedicated category-tree tests remain responsible for
  validating real keyboard Home/End/Up/Down behavior by delivering paired
  `WM_KEYDOWN`/`WM_KEYUP` messages to the focused category-tree HWND. Calling the
  retained tree control's `OnKeyDown(...)` method directly is model coverage and
  cannot replace the HWND routing test.
- UI self-tests that assert no unrelated host repaint churn during a stimulus
  must first wait for the unrelated host's debug render count to settle after
  navigation and page setup. Pending renders from setup must not be charged to
  the action under test, but the stimulus window may still require zero extra
  unrelated renders, resize churn, or resize failures after that settled
  baseline. The settle wait must require a meaningful idle window with multiple
  unchanged samples; a single unchanged poll is not enough evidence that
  pending setup paints have drained.
- Repaint-churn baselines must be captured after foregrounding, hit-test
  preparation, `UpdateWindow`, and other setup that can pump messages or
  invalidate UI hosts. Those setup paints must not be measured as part of the
  stimulus being validated.
- UI self-tests that assert render-count budgets for the stimulated host itself
  must also measure from a settled pre-action baseline to a settled post-action
  snapshot. Render-count deltas captured around a single message-pump pass are
  order-dependent evidence because queued setup, hover, or paint work can be
  charged to the wrong action.
- Preferences category-switch coverage records both render and invalidation
  deltas across the same settled interval. It must require at least one render,
  bound the invalidation source, and prove renders are attributable to those
  invalidations plus at most the explicit page-switch redraw. A smaller fixed
  render ceiling that rejects fully attributable work under broad-process load
  is not a valid stability gate.
- UI self-tests that drive a control through debug-only helpers and then assert
  repaint evidence must make the stimulus deterministic: establish a scrollable
  starting position, verify the behavioral state change such as scroll offset,
  and force/update the target window's paint when the helper invalidates without
  going through the real message dispatch path.
- UI self-tests that assert hover-dependent state after pointer clicks must keep
  the real cursor position and synthetic mouse messages aligned. Sending
  `WM_MOUSEMOVE` without moving the cursor can be overwritten by normal message
  pumping and turn hover assertions into current-desktop artifacts.
- UI self-tests that wait on debug snapshots must format failure diagnostics
  from the post-wait snapshot, not from a snapshot captured before the wait.
  This is required for focus/scroll assertions where the observed state can
  change while the wait helper is polling.
- UI self-test layout snapshots consume the last completed normal render. A
  missing layout may request a future paint with ordinary invalidation and retry
  under a bounded deadline, but the snapshot helper or timer-driven test step
  must not force a synchronous nested redraw. In particular, it must not call
  `RedrawWindow(..., RDW_UPDATENOW, ...)` or an equivalent immediate paint from
  the active self-test callback: deep UI/test call stacks plus the graphics
  driver's `EndDraw` path can exhaust the UI-thread stack.
- Large UI self-test fixtures must remain in separately compiled phase dispatch
  helpers rather than one combined switch frame. Debug PDB frame evidence must
  retain material headroom beneath the default 1 MiB UI-thread stack for nested
  popup/rendering and crash-reporting calls; when a phase helper grows toward
  256 KiB, split it before adding more stack-resident fixtures. File Operations
  source-contract coverage must reject recombining its Phase 10–13 dispatchers.
- UI self-test debug snapshots that expose retained focus targets must report
  `None` when no DxUi control currently owns retained focus. Family-specific or
  optional controls must be pointer-guarded before comparison so a null retained
  focus cannot be misreported as a control that is absent on the active page.
- UI self-tests that expose both native focus ownership and retained DxUi focus
  targets must assert those states independently. Native focus on another child
  HWND, such as a dialog category tree, does not require a DxUi host's retained
  focus target to become `None` when the retained control still belongs to that
  host.
- UI self-tests that send native edit messages to a DxUi text-input target must
  wait for both the retained DxUi text-field focus and a valid focused Win32
  input target after any deferred rebuild, search refresh, or list rebind. A
  matching retained focus target alone is not enough evidence that native
  `WM_CHAR` / edit-message routing is ready to receive input; tests that probe
  compatibility `EM_SETSEL` / `EM_REPLACESEL` paths must drive the native DxUi
  host target and retained text session directly.
- UI self-test debug snapshots that expose a single `focusTarget` field and use
  it to drive keyboard input must report the active native-focus owner, not a
  retained-only DxUi control from a host that no longer owns native focus.
- GUI Tab-traversal self-tests must include retained focus, native focus
  ownership, and modifier-state evidence in failure diagnostics when those
  states can affect routing. Ordered traversal scripts must stop at the first
  failed step instead of continuing into cascaded focus traces that obscure the
  original failure.
- The shared Commands message pump must be cooperative and bounded. It must not
  drain the queue indefinitely because active DxUi hover, cursor, timer, or
  paint traffic can otherwise turn a bounded wait into a long-running queue
  drain and hide the real test step duration.

## Artifact Contract

Self-test artifacts must preserve enough detail to explain why a run passed, failed, or skipped:
- `results.json` records final case status and reason,
- `trace.txt` records supporting diagnostic context,
- `perf_metrics.jsonl` must be preserved when the scenario emits performance metrics,
- archived copies under `Specs/TestRuns/` must keep those files intact.

Archive-to-repo discovery must only accept a candidate repository root when it
contains all of:
- `RedSalamander.sln`,
- `Specs/TestRuns`,
- a `.git` directory or `.git` worktree file.

The parent walk from the executable directory must remain explicitly bounded.

Filesystem-heavy Commands and Compare selftests MUST use the shared
`SelfTestCommon` helpers for repository-root discovery, Local AppData lookup,
directory creation, UTF-8 file writes, stable device hashing, bounded JSON
unsigned-integer extraction, and MTP debug-export loading instead of maintaining
per-suite copies. Repository discovery honors `REDSALAMANDER_REPO_ROOT` and must
work from a normal checkout and a `.git` worktree file. The MTP loader applies
the product plugin-manager eligibility rules, including disabled/deferred paths,
pins the DLL with `LoadLibraryExW`, and validates a typed debug export before it
is called.

Source-contract tests MUST NOT pin product implementation wording, private member
order, helper spelling, or other source layout when a compiled behavioral test or
debug seam can verify the requirement. Source inspection remains appropriate only
for contracts that cannot be observed in the built artifact, and such tests must
state why behavioral coverage is unavailable.

## Menu and command-surface validation

- `main_menu_information_architecture` validates direct ownership and relative order for the Files, Edit, Commands, View, and Help menus; the intentional two-position Path from Other Pane contract; the Left/Right Terminal Pane placement; the popup Terminal command; Change Case/Attributes; registry mappings; and high-value default shortcuts.
- `command_surface_coverage_contract` requires every canonical command registry entry to have exactly one reviewed classification in `RedSalamander/SelfTest/Commands/CommandSurfaceCoverage.inc`. It opens the unfiltered Folder and Terminal Command Palette projections, requires one visible row per context-eligible palette-visible command, and fails if a palette-visible canonical command is unreachable from both contexts.
- `Tools/Tests/ResourceLocalizationContracts.Tests.ps1` validates base/satellite menu-tree parity, the exact reviewed FolderView/viewer/Monitor command sequences, and one unique actionable static access key per separator-delimited sibling group. The hidden ViewerText encoding catalog is picker data and is the only menu-resource mnemonic exemption.
- `viewer_text_encoding_picker_open_and_filter_full_catalog` exercises two consecutive F8 cycles and the process-local catalog cache; opens the localized picker with its current row selected and visible; proves Escape and close-during-filter teardown preserve state; proves the full catalog is present; exercises zero-result, codepage, and rapid replacement filtering; verifies disabled/reenabled OK state; commits Windows-1252; and confirms the main Encoding menu check state.

## DxUi Popup Validation

For command selftests that validate migrated app chrome using DxUi popup menus:
- owned `DxUi_ContextMenu` popup windows are an authoritative “menu opened” signal,
- popup dismissal may be validated by observing that owned popup window close, not only by `GUI_INMENUMODE`,
- tests must not rely on `GUI_INMENUMODE` alone once the validated surface routes through the shared DxUi popup path instead of a native Win32 menu loop.

## DxUi Pointer Validation

For command selftests that validate real pointer interaction on DxUi controls:
- a target inside a scrollable DxUi host must be scrolled into the viewport
  before the selftest sends mouse messages,
- debug host/client rectangles consumed by pointer tests must be in the target
  HWND's client coordinate space after any `ScrollPanel` offset has been
  applied,
- a control's semantic `IsVisible()` state is not enough to prove it is
  onscreen and clickable when the control lives in scrollable content.

## FolderView Current/Selection State Validation

The FolderView current-item and selection contract requires deterministic Commands coverage through these exact cases:

- `folder_view_focus_invariant_generic_rebuild`: cover surviving-current preservation, single and multiple disappearance, first-successor/nearest-predecessor repair under keyed sort and Sort None, filter/hide repair, empty/repopulation, and index/name/focused-flag agreement at every checkpoint.
- `folder_view_empty_background_keeps_current`: drive real left misses with every modifier combination and a pointer right-click miss; require selection clear, current/ownership retention, anchor reset, drag disarm, no item pointer target, and pending removal-focus eligibility unless a later gesture actually changes current. Keyboard context-menu routing must remain selected-or-current and must not use a synthetic point hit-test.
- `folder_view_focus_memory_bounded_lru`: cover insert/update/touch, 513th-entry eviction, 512 KiB payload eviction, oversized/unsupported no-cache behavior, deterministic LRU ordering, and both caps after every mutation.
- `folder_view_focus_memory_navigation_roundtrip`: cover A → B → A, Back/Forward, and identical path text under different provider/context identities; current restores independently while selection is empty after every actual navigation.
- `folder_view_selection_input_contract`: cover plain/Ctrl/Shift/Ctrl+Shift pointer and keyboard ranges, unmodified navigation, Alt+Up/Down selected-identity traversal, Space, Insert, Escape, Select All, inversion, masks, item/right/background clicks, current-on-button-down, anchor rules, notification ownership, and current/selection independence.
- `cmd_pane_selection_select_calculate_directory_size_next_keeps_navigation_shell_stable`: cover Space from both unselected and selected current folders, the immutable post-toggle selected snapshot, size start/cancellation and stale-result rejection, next/no-wrap movement, destination-not-selected behavior, final status state, navigation-shell quiet points, and Quick Search Space remaining text with no size request.
- `folderView_empty_folder_state`: require zero real/current/selected/anchor state plus a separate parent Invoke action; cover Enter/double-click/Backspace, root callback, and filter-empty/host-message/busy/cancel/error exclusions. The action is never an item command or drag source.
- `folder_view_selection_refresh_filter_contract`: cover surviving selection, removal/filter/hide no-resurrection, Select All followed by filter/unfilter, new-item-after-Select-All, ordinary disappearance/reappearance, rejected result preservation, and unselected automatic current repair.
- `cmd_pane_navigation_directory_impact_preserves_selection_across_chained_renames`: cover selected-only proven rename-chain transfer to the final identity and prove that an unselected rename chain stays unselected.
- `folder_view_drag_source_target_contract`: cover selected-or-current immutable drag snapshots, system drag-threshold rectangle, source membership after modifier transitions, Ctrl+Click deselection, lifecycle disarm, empty-background movement, double-click isolation, and agreement with command fallback.
- `folder_view_selection_visibility_restore_contract`: cover Hide Selected/Unselected plus Show, hidden/system visibility, filtered masks/inversion, explicit Save/Restore timing, and non-actionable rename provenance.
- Existing host-owned Delete/Move cases must also assert exactly-one-current and the selection-only ownership-epoch rule while retaining exact-per-source-`S_OK`, provider/folder/sort, newer-enumeration, and immutable-target gates. At least one keyed-sort fixture MUST deliver provider enumeration order different from resulting display order and prove that host removal uses the canonical successor/predecessor comparator, excludes every proven-removed identity, and never applies a captured display index to the raw payload.

These cases must use source-derived inventory and assertions. Do not hardcode mutable suite totals into normative text.

## NavigationView DxUi Text-Host Validation

For command selftests that validate NavigationView address-bar edit mode or full-path popup edit mode:
- `NavigationViewDebugSnapshot` is the authoritative contract for edit visibility, focus target, current edit text, selection range, `currentEditUsesNativeTextInput`, `currentEditHostHwnd`, the active backend-neutral `currentEditInputHwnd`, caret screen rectangle validity/coordinates, and active composition start/end state,
- tests must not enumerate descendant native `Edit` / `RICHEDIT50W` windows to prove edit mode, because NavigationView no longer exposes a visible native edit child and no longer installs NavigationView policy on a hidden child edit surface.
- path-region keyboard activation selftests must verify active native IME composition owns Enter/Escape/Tab before NavigationView submit/cancel/tab policy runs.
- edit-suggest keyboard-routing selftests must verify active native IME composition owns Up/Down before NavigationView suggestion-selection policy runs, so candidate/navigation keys do not change edit-suggest selection while composing.
- address-bar edit clipboard coverage must verify the focused DxUi host remains active and pane-level Select All, Copy, and Paste commands mutate/copy the edit text instead of falling through to FolderView command handling.
- invalid-path validation coverage must assert edit mode remains active and `NavigationViewDebugSnapshot::currentEditHelpText` exposes the rejected path text through the retained DxUi `TextField` HelpText contract.
- NavigationView pointer and region-keyboard selftests must not inherit the
  navigation-bar visibility left by earlier commands. Before using hit-test
  rectangles or `DebugFocusNavigationViewRegion(...)`, they must snapshot the
  previous navigation-bar visibility, reveal the target pane's NavigationView,
  wait until its child HWND is visible, and restore the original visibility
  before returning.

## Preview Pane and ViewerVLC Validation

For command selftests that validate embedded preview behavior:
- `FolderWindow::PreviewPaneDebugSnapshot` is the authoritative contract for preview activity, source/host panes, selected Folder/Preview tab, hosted plugin ID, tab-strip HWND and tab hit rectangles, Preview close visibility, tab-strip visible/pending tooltip text, content HWND, embedded viewer HWND, and embedded viewer instance identity.
- Preview tests must verify that embedded viewers do not take keyboard focus from the source pane, that focus changes reuse the current embedded viewer instance and HWND when the resolved plugin is unchanged, that stale content is cleared before the new file is reported as rendered, and that preview close/replacement persists changed plugin configuration.
- Preview source-contract tests must cover menu-bearing embedded viewers so Preview-appropriate menu options remain reachable from right-click context menus without requiring a visible embedded menubar or standalone filename/header chrome, while standalone-only actions and shortcut labels stay out of the embedded menu.
- Preview responsiveness coverage must verify that rapid same-plugin and cross-plugin preview switches do not wait for slow media/player teardown on the UI thread.
- Preview source-contract tests must cover embedded media audio-only paths so VLC visualizers cannot create top-level player or visualization windows outside the Preview host.
- When saved viewer associations are empty or resolve only the default text viewer, preview tests must cover the built-in embedded viewer defaults before `builtin/viewer-text` fallback.

For command selftests that validate `builtin/viewer-vlc`:
- `WndMsg::ViewerVlcDebugGetSnapshot` is the authoritative contract for HUD state, volume/mute state, snapshot dimensions, Fluent icon glyph usage, filled-button HUD styling, and outstanding asynchronous load-work count.
- Wheel-seek coverage must exercise both normal viewer surfaces and libVLC-owned child video windows so embedded preview and standalone playback keep the same wheel behavior.
- Slow-stop coverage must use `WndMsg::kViewerVlcDebugSetStopDelay` to ensure media-to-media preview switches, including video-to-audio transitions, avoid slow player retirement, media-to-image preview switches stay responsive while VLC player cleanup continues in the background, and VLC debug snapshots must assert that the embedded video surface remains a child of the preview viewer after same-plugin media navigation.
- Cleanup-gate retirement timing must first wait for the snapshot's outstanding load-work count to reach zero. Earlier opens may still own worker jobs under broad-suite load; those jobs are not evidence that the gated retirement cleanup failed.
- Async terminal-state coverage must use `WndMsg::kViewerVlcDebugSetAsyncControl` to force loader-submit failure, load-completion payload-post failure, delayed stale load completion, and close-completion payload-post failure. It must prove that loading terminates, no stale player attaches/plays, retained parent/video HWNDs stay hidden and valid until cleanup, every payload-post fallback is drained on the UI thread, each window identity is destroyed once, and `ViewerClosed` is delivered once.

## Search-Specific Coverage

Search coverage follows the same contract:
- local, fallback, indexed, and service search cases stay declared,
- ReFS validation stays declared even on machines without ReFS,
- a machine without a fixed ReFS volume records `skipped` with a reason instead of silently omitting the case.
- Direct SQLite cutover, prefilter, and injected SQLite failure cases require a
  readable live NTFS journal cursor for the requested root. When that cursor is
  unavailable, the case must remain declared and record `skipped` with that
  precondition as the reason.
- SQLite service no-wait query tests must distinguish a healthy service that
  answers by live scan from host fallback. `DEGRADED_NO_INDEX` means the service
  stayed healthy but could not use a ready/current database; only
  `SERVICE_UNAVAILABLE` means the host switched away from the service
  transport.
- Service tests whose assertion is not about the default ProgramData database
  must use isolated foreground-service storage paths so pre-existing local
  SQLite state cannot change the result.
- Service status tests must cover every persisted SQLite volume state that affects
  query cutover. `CURRENTNESS_UNPROVEN` must report cutover blocked before the
  first query while keeping the legacy-import migration count at zero.
- Startup-warmup status cases must inspect the persisted volume state and require
  aggregate service readiness to match it; they must not assume every test machine
  can prove live USN currentness. Query-path assertions branch on that persisted
  readiness and the readable-journal capability, while deterministic mixed-state
  coverage supplies the unconditional `CURRENTNESS_UNPROVEN` regression guard.
- External-generation rotation coverage must prove prior-generation runtime mode,
  store/sync/fallback state, root counters, and active root are cleared. An
  uninspectable same configured store may retain the last request execution mode,
  but must re-derive its remaining status from the failed inspection.

## Source Organization

Project-owned self-test implementation must live in `.cpp` source files, with optional `.h` declarations when cross-translation-unit declarations are needed.

Project code must not introduce `.inl` self-test implementation files.

The `--commands-selftest` suite may continue to include family source files from `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp`, but those included family files must still be `.cpp` files rather than `.inl` fragments.

`Commands.SelfTest.Preferences.cpp` may include Preferences chunk files for
navigation, but each `Commands.SelfTest.Preferences.*.cpp` chunk must own a
complete anonymous namespace wrapper. A chunk must not close a namespace opened
by the coordinator or leave a namespace open for the next family include.


Stale-run cleanup given a synthetic process snapshot must also receive an explicit `CandidateRunIds`
allowlist. Tests may remove only their own uniquely named run directories; injected PIDs never
authorize a sweep across the shared live root. Disk-audit fixtures assert their own entries and
schema consistency without assuming that concurrent runs or diagnostic root entries do not exist.

A native harness in a clean package extraction accepts an explicitly configured marked test root without a
repository beside its executable or current directory. TestSupport validates that root first; it discovers the
checkout only when choosing a default drive. Absolute-root, ownership-marker, traversal and reparse checks are
unchanged. The portable PluginContractTests local-writer proof exercises this path outside the checkout. The RequiresBuildToolchain Pester fixture in Tools/Tests/RelocatedTestSandbox.Tests.ps1 also builds a native harness outside any checkout, proves scratch writes with the marked root, and rejects outside-root/ADS/unconfigured requests. It uses the selected native configuration and the shared sanitizer runtime staging.


## Documentation screenshot scenarios

Application screenshots for specifications, documentation and UI reviews MUST come from deterministic
scenarios in the test harness. Static captures use the noninteractive activation guard. Explicitly
authorized interaction captures may borrow focus briefly through the existing warning and runner
desktop lease; they target only process-owned test windows and restore cursor/foreground state.
Computer Use and manual desktop screenshots are not substitutes. Captures MUST identify the scenario, theme, window size, DPI,
source/build identity and whether operation data is synthetic. Concept mockups are separate artifacts.

`FileOps_VisualGallery` is an explicit opt-in File Operations case. It is excluded from the ordinary
phase order and accepted by exact case selection. It injects presentation snapshots into the real
popup, reuses the engine conflict-action policy, captures owned menus and the custom speed prompt,
and exercises the graphics-independent fallback. It performs no provider transfers. Its synthetic
rates and receipts establish rendering coverage, not engine correctness or performance.

The canonical capture helper is `Tests/TestSupport/WindowScreenshot.h`. It captures the complete
process-owned HWND, including DirectComposition and window chrome, through Windows Graphics Capture
and writes PNG through WIC. It rejects foreign, hidden or minimized windows and output paths outside
the marked test sandbox. Frame readiness is event-driven and bounded. Missing capture support is a
failure, never permission to substitute a desktop screenshot or a partial renderer layer.

The focused runner command is documented in `Tests/README.md`. Runtime PNGs and the TSV scenario
manifest remain under the owned run's `artifacts/selftest/last_run/fileops/visual-gallery` directory.
Reviewed copies may be published into the requested spec/documentation folder, with a provenance
manifest and explicit coverage gaps. Screenshot inspection does not replace interaction, UIA,
assistive-technology, native DPI/locale, real-provider or performance qualification.

`FileOps_VisualGallery` loads French resources, exercises every enumerated physical monitor at its
actual native DPI, and records its work area. It includes long French paths, live confirmation overlays
with safe cancellation, populated/selected/scrolled Issues views, destination history and recovery menus.
These are synthetic presentation/diagnostic fixtures, not provider execution receipts.

`FileOps_VisualGalleryInteraction` is a separate exact-case, opt-in foreground lane, also excluded from
the broad phase order. Only that exact FileOps case obtains `RequiresInteractiveDesktop` and omits
`--selftest-no-activate`. The existing `DirectedSelfTestInputWarning` appears before activation. Its
`TryActivateOwnedWindow` requires a visible warning and a target on the calling process/UI thread;
foreground queue attachment is scoped to activation. Input is refused when the target lacks foreground
ownership. The case records Tab/Shift+Tab focus identities, real hover, pointer-down capture and a
same-HWND monitor round trip. Screenshot evidence is not full assistive-technology qualification.
