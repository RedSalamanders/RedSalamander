# SelfTest Specification

## Overview

RedSalamander ships three debug self-test suites:
- `--compare-selftest`
- `--commands-selftest`
- `--fileops-selftest`

This document defines the normative result contract shared by those suites.

Related documents:
- `Specs/Testing/Testing_TestCoverage.md` — comprehensive per-suite test case inventory
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/TestRuns/README.md`
- `Tests/README.md` — central test infrastructure index
- `Tools/Run-AllTests.ps1` — unified test runner with summary reporting

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
- Skipping is part of normal suite behavior for conditional coverage and must not make the suite fail by itself.

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
- In-product self-test suites must use the shared result-emission helper for suite-level case insertion, status counts, and first-failure propagation. Suites may keep distinct execution models, such as `SelfTest::RunCase` or FileOperations phase-state execution, but summary emission must stay centralized.
- In-product self-test suites must flush suite `results.json` through the shared case-result emission path after every recorded case. A crash or timeout must not erase the identities and statuses of earlier completed cases.
- The unified runner must compare each executed in-product suite's result names with runner-native `--selftest-list-cases` output for the same suite/filter and fail on duplicate expected names, duplicate actual names, missing declared results, or unexpected extra results. When `--selftest-repeat=N` is active, duplicate actual names are valid only when every expected repeated case appears exactly `N` times. CompareDirectories may emit an extra `setup` result for explicit setup failure reporting.
- `Tools/TestInventory.ps1` is the source-derived manifest for static Commands
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
- When `Tools/Run-AllTests.ps1` is invoked with a non-empty `-CaseFilter`, runner-native case listing must return at least one expected case for each executed self-test suite. A zero-expected, zero-actual filtered run is invalid evidence and must fail as result coverage drift.
- Runner-injected coverage failures must contribute to the effective suite status, displayed failure counts, runner aggregate `exit_code`, and process exit code even when the native self-test process exits successfully.
- Checked-in archived runs may contain skipped cases; that is valid when the skip reason documents the missing precondition.
- `Tools/Run-AllTests.ps1` must tolerate both direct suite JSON (`commands_results.json`, `compare_results.json`, `fileops_results.json`) and aggregated run JSON (`selftest_run_results.json`) when summarizing archived results.
- `Tools/Run-AllTests.ps1` must also write a runner-owned `run-all-tests-results.json` artifact for every invocation. Multi-suite runs must preserve each suite's status counts, case names, failure reasons, skipped reasons, wrapper exit code, canonical `test_root`, and per-run `run_id` in that aggregate file so later suites cannot overwrite earlier evidence. The aggregate must also preserve a non-blocking `test_sandbox_audit` object that records the current TestSandbox root, allowed run id, unexpected sibling run directories, unexpected direct children under the TestSandbox root, and any resolved legacy cleanup targets so full-run disk proof has durable evidence.
- `Tools/Run-AllTests.ps1` must also derive `run-all-tests-case-history.jsonl` and `run-all-tests-dashboard.md` from the aggregate summary. History rows must include `harness`, `case`, `duration_ms`, `status`, `reason`, `classification`, `seed`, `attempt`, and whether the row came from native case output or retry/shuffle evidence.
- Runner aggregate summaries must preserve blocking classification metadata: per-suite `classification`, `classification_reason`, `retry_attempts`, and top-level `classifications` counts for `flaky`, `regression`, `isolation_suspect`, and `unclassified_failure`.
- `Tools/Run-AllTests.ps1` must create a canonical `REDSALAMANDER_TEST_ROOT` sandbox base, defaulting to the repository-local `.build\TestSandbox`, create one invocation under `REDSALAMANDER_TEST_ROOT\runs\<runId>\`, set `REDSALAMANDER_TEST_RUN_ID` for child processes, and clear any inherited `REDSALAMANDER_SELFTEST_ROOT` before normal suite execution. Inherited legacy roots must not choose the normal runner context. Native self-test helpers must consume `REDSALAMANDER_TEST_ROOT` + `REDSALAMANDER_TEST_RUN_ID` directly and write their existing `last_run` tree under `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\selftest\last_run`. `REDSALAMANDER_SELFTEST_ROOT` remains a compatibility override for deliberate ad hoc or repair-lane launches only; unrelated roots must not silently mix artifact locations.
- Local capability-sensitive evidence may override `REDSALAMANDER_TEST_ROOT` to the short NTFS root `C:\RSPerf` when the repository-local default is not NTFS, is redirected/network-backed, or is too long for the path budget of the scenario under test. The override must still use the normal runner layout under `C:\RSPerf\runs\<runId>\scratch\...` and `C:\RSPerf\runs\<runId>\artifacts\...`, and the run record must name the override and reason. `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox` is diagnostic history only for path-sensitive scenarios because it can be too long for long-path fixtures; `C:\RST` is legacy cleanup debt and must not be used for new evidence.
- Native self-test scratch files must be acquired through `SelfTest::AcquireTestSandbox(...)`. Under the unified runner, that helper writes scratch data under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\<suite>\<case>\`, separate from `artifacts\selftest\last_run`; ad hoc legacy launches may fall back under `last_run\<suite>\scratch\<case>` only for compatibility. Commands plugin-config viewer perf and viewer close-roundtrip fixtures use `SelfTest::AcquireTestSandbox(SelfTestSuite::Commands, ...)` under `scratch\commands\<case>` for `viewer_text_hex_byte_color_perf`, `viewer_text_diff_perf`, `viewer_imgraw_close_roundtrip`, and `viewer_web_close_roundtrip`; BatchRename window fixture roots that need a concrete local-looking root use the same helper instead of fixed `C:\BatchRename*SelfTest` paths; ShellCommands long-path shortcut-save fallback uses the same helper under `scratch\commands\shell_shortcut_save_temp`; Compare dummy filesystem scratch for `dummy_content`, `normalized_name_collision_preserves_same_side_entries`, `deep_tree`, and `invalidate` uses `SelfTest::AcquireTestSandbox(SelfTestSuite::CompareDirectories, ...)` under `scratch\compare\<case>` instead of fixed `CompareSelfTest_*` drive-root paths. Direct process-temp APIs such as `GetTempPathW`, `GetTempFileNameW`, and `std::filesystem::temp_directory_path` must not be added to in-product self-test code when a `TestSandbox` path is sufficient.
- Standalone harnesses that cannot link the native self-test helper yet must still consume `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID` directly and use the same `runs\<runId>\scratch\<harness>\<case>\` layout. RedConfigureTests scratch roots live under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\redconfigure\<case>\`; ViewerSqliteTests database scratch roots live under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\viewer-sqlite\database`; CrashHandlingTests marker-file scratch roots live under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\crash-handling\marker-files`; PerformanceTests2 fixture roots live under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\performance-tests2\<case>`; ViewerPETests runtime fixture files live under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\viewer-pe\<case>`. DxUiTests generated default artifacts, including the control-gallery/button-contrast PNG defaults and local Animation/WindowHost perf JSONL sinks, live under `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\dxui`; explicit `--gallery-output`, `--gallery-output-directory`, `--button-audit-output`, and `--perf-jsonl` paths remain caller-owned outputs. Direct launches for those migrated harnesses fall back to `<cwd>\.build\TestSandbox` or the sibling `.build\TestSandbox` when launched from `.build\<platform>\<configuration>`. Migrated standalone harnesses must not acquire scratch through `std::filesystem::temp_directory_path`.
- Runner preflight cleanup must recognize both canonical runner IDs (`<UTC>-<pid>-<guid>`) and direct-harness fallback IDs (`<harness>-<pid>-<tick>`). It may remove only non-current/non-allowed run directories whose parsed owner PID is no longer live, or whose directory predates the current process that reused that PID; actual live-owner and unparseable manual directories remain protected. Process-start lookup must use a Windows process-metadata fallback when ordinary process APIs deny access, so protected system processes cannot accidentally shield an unrelated stale run after PID reuse.
- Tooling Pester tests that create scratch files, fake toolchains, or synthetic perf-run trees must use `Tools\TestRunPlan.ps1` `New-RSTestSandboxScratchDirectory(...)` under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\tools-pester\<case>\`. They must not allocate new scratch roots through `[System.IO.Path]::GetTempPath()` when the TestSandbox path is sufficient. `Tools\Clean-TestSandbox.ps1` must keep historical `%TEMP%` Pester roots in its legacy cleanup allow-list until those roots are no longer observed in local/CI machines.
- FileOperations real cross-volume coverage is the only sanctioned alternate-volume exception to the single-drive `REDSALAMANDER_TEST_ROOT` scratch rule. It must allocate through `SelfTest::AcquireTestSandboxOnVolume(...)` under `<AltDrive>:\RedSalamanderTestSandbox\runs\<runId>\scratch\<suite>\<case>\`, remove the case root after use, and prune empty run parents. New allocation must not create `RedSalamanderCrossVolumeSelfTest_*` roots; that legacy pattern is cleanup-only historical debt.
- FileOperations queue-state selftests must prove the queue state they assert. Tests that cancel a queued task must start the active holder first, wait until that task has entered operation and owns the active slot, then start the queued task, assert `IsWaitingInQueue()` before queued-cancel assertions, and fail fast if the queued task enters operation or completes before cancellation. Timeout diagnostics must include task ids, live/completed state, queue flags, progress, and HRESULTs for every task participating in the queue proof.
- Compare selftests that launch `RedSalamanderSearchService.exe --run-foreground` must create a JobObject with `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE`, start the child process suspended, assign it with `AssignProcessToJobObject`, and resume the main thread only after assignment succeeds. Foreground service processes must not be able to survive a selftest harness crash or timeout and hold SQLite/WAL or scratch-root handles into the next run.
- Compare selftests that gate foreground SearchService readiness must poll the explicit service status for the intended pipe and must diagnose early child-process exit with exit code plus captured stdout/stderr. Query-performance and parity tests must establish service readiness before the first warmup query; the first query is not an acceptable readiness probe.
- Foreground SearchService readiness polling uses a scaled, bounded 30-second default. Freshly linked images and cold SQLite stores can legitimately spend more than 10 seconds in image verification and initialization before creating their unique pipe; successful warm starts still return immediately, and cases with a tighter semantic budget may pass an explicit override.
- Foreground SearchService fault-injection and first-query gates must use the process-aware status helper with child output capture. A timeout, broken pipe, or early exit must retain the child exit code and captured stdout/stderr instead of reporting only the client-side pipe HRESULT.
- Compare selftests own foreground SearchService lifetime through the unique pipe: readiness requires a successful status response whose reported pipe matches that intended pipe, normal teardown sends the foreground-only graceful-shutdown request, and the child must exit within the case's bounded shutdown budget. `--max-requests` and incidental later status probes are not valid teardown mechanisms.
- Before child tests launch, `Tools\Run-AllTests.ps1` must sweep parseable sibling run directories whose names match `runs\<timestamp>-<pid>-<guid>` and whose owner PID is no longer live. The current run id, explicitly allowed run ids, live-PID sibling runs, and unparseable/manual directories must not be removed by this dead-PID sweeper. Removal failures must warn and return failed cleanup rows; they remain visible evidence and must not silently turn into a green disk proof.
- `Tools/Run-AllTests.ps1` must run `Tools/Clean-TestSandbox.ps1 -Apply -Confirm:$false` before build and child-test execution unless `-SkipLegacySandboxCleanup` is explicitly supplied for diagnosis. The cleanup script must be dry-run by default outside that runner call, require `-Apply` for removal, use `ShouldProcess`, and delete only resolved literal targets from the legacy allow-list: `%LOCALAPPDATA%\RedSalamander\SelfTest`, known `%TEMP%` standalone/perf test roots, and fixed-drive `RedSalamanderCrossVolumeSelfTest_*` roots. Individual removal failures such as locked or access-denied historical dumps must warn, return a `Status=Failed` row with the removal error, and must not abort `Run-AllTests.ps1` before build/test execution starts.
- `Tools/Run-AllTests.ps1 -Suite CI` is the GitHub Actions PR gate contract. It must route the gate through the runner, execute the three in-product self-test suites as separate processes, split DxUiTests by suite, preserve the explicit ViewerPE prompt cases, run deterministic FileSystemCurlTests, RedConfigureTests, PluginContractTests, SettingsSchemaTests, and CrashHandlingTests coverage, exclude only `RequiresBuildToolchain` Pester cases from artifact-only CI, classify failures, publish the runner summary to `GITHUB_STEP_SUMMARY` when GitHub provides it, and upload the runner-owned `REDSALAMANDER_TEST_ROOT` tree. RedSalamanderMonitorEtwLatency remains a broader closeout-only `-Suite Full` gate. Pull requests must also run the reusable build workflow for Debug ARM64 without attempting to execute ARM64 binaries on the x64 hosted runner.
- `PluginContractTests` must require each built-in filesystem plugin's `RunDebugSelfTests` export in Debug configurations. A missing expected export is a test failure, not a silent pass. Configurations that intentionally omit debug exports must report an explicit skip; if an optional export is present, the harness still executes it and requires its reported failures to remain zero.
- Portable ZIP creation must expand the completed archive into a new directory, validate the shared runtime-dependency manifest, start the packaged app through `--help`, and run the matching plugin harness there with `--package-smoke`. Package-smoke mode validates load/enumeration/exports/schemas/capabilities for every built-in plugin but deliberately excludes Debug-only selftests and runtime-refresh probes. ARM64 package execution may be skipped on an x64 host only after full file/dependency validation.
- `.github/workflows/asan.yml` owns the scheduled and high-risk-path x64 `ASan Debug` plugin-contract lane. Before accepting a green harness run, it must invoke the opt-in `--asan-seed-heap-overflow` probe, require a nonzero result with an AddressSanitizer diagnostic, and then run the ordinary contracts successfully. `_DISABLE_STL_ANNOTATION` remains documented technical debt for non-ASan vcpkg dependencies; the seeded proof covers ordinary heap detection, not STL container annotations.
- `.github/workflows/nightly-flake.yml` owns the expensive scheduled/manual shuffle-repeat lane. It must stay separate from PR CI, build the Debug self-test artifact, run `Tools/Run-AllTests.ps1 -Suite All -SkipBuild -SelfTestRepeat 5 -SelfTestShuffleSeed <seed> -ClassifyFailures`, and upload the runner-owned `REDSALAMANDER_TEST_ROOT` tree as `selftest-artifacts-nightly-shuffle`.
- Failure classification is blocking evidence, not a pass policy. `FLAKY` means an entry failed and passed on retry; for in-product suite failures it requires the failed cases to pass under `--selftest-case` and then pass shuffle triage. `REGRESSION` means retry evidence failed again, including any shuffle-triage seed that reproduces the failure. `ISOLATION_SUSPECT` means an in-product suite failed and failed cases passed under `--selftest-case`, but shuffle/order triage evidence is missing. Shuffle triage attempts must preserve their `shuffle_seed` in `retry_attempts`. Focused subset runs with `-CaseFilter` must still use failed-case retry plus shuffle triage, preserving the original subset filter for shuffle attempts instead of falling back to a whole-entry retry. When a native suite implements seeded/repeated order by dispatching one exact case at a time, that exact-case replay remains suite/shuffle context for classifier proof; only the runner's failed-case retry is isolated evidence. All non-`PASSED` classifications must keep the aggregate run exit code non-zero until the underlying test is fixed, replaced, or explicitly tracked through the reviewed quarantine process.
- Before each Commands seeded/repeated exact-case dispatch, the harness must close same-process non-main top-level windows, verify the main window is enabled, and reactivate it. A failed isolation boundary must be recorded as that case's failure instead of running the fixture against inherited UI state. Fail-fast bookkeeping may skip this cleanup after the first recorded failure because remaining cases do not execute.
- Reviewed temporary quarantine entries live in `Tools/test-quarantine.jsonl`, one JSON object per line. Empty/blank lines mean no active quarantine. Every nonblank entry must include `harness`, `name`, `owner`, `opened`, `expires`, `issue`, `root_cause_hypothesis`, and `fix_or_replace_plan`; `opened` and `expires` must use `YYYY-MM-DD`, `expires` must not be in the past, and an entry must not span more than 30 days. Invalid entries and valid active entries both keep `Run-AllTests.ps1` red; quarantine is a blocking repair ledger, not a nonblocking suppression list. Active entries must match a runnable harness adapter name from the current runner plan, and self-test entries must also match that adapter's `--selftest-list-cases` inventory. Matching self-test entries run in a separate repair lane with `--selftest-case=<name>` after the main gate, standalone entries rerun their matching adapter entry, and `run-all-tests-results.json` must preserve `quarantine.repair_attempts`, `repair_attempt_count`, and `repair_reproduced_count`.
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
- Builds and self-tests that use the same output configuration must be
  serialized. `build.ps1` may close an ordinary interactive process running
  from the exact target output path, but it must abort with PID, path, and
  command-line diagnostics instead of force-killing a process whose command
  line contains a `--...selftest...` flag. An exact-path process whose command
  line cannot be inspected is protected by the same fail-safe behavior.
- `build.ps1`, standalone `Tools\Invoke-SanitizedMsbuild.ps1`, deployment build
  tests, and the complete `Run-AllTests.ps1` lifecycle, including `-SkipBuild`,
  must hold the same repository-scoped exclusive lock file using `FileShare.None`.
  The protected test lifecycle starts before TestSandbox cleanup and ends only
  after result archival and disk audit, so another runner cannot invalidate the
  audit or load binaries while they are being replaced. Active ownership is
  recorded in `.build\artifact-operation-owner.json` with the root PID,
  operation, and UTC start time. Balanced nesting is limited to the owning
  managed thread. A direct child may reuse ownership only when the synchronous
  native launcher creates it inside the kill-on-close Job Object through
  `PROC_THREAD_ATTRIBUTE_JOB_LIST`, restricts inherited handles through
  `PROC_THREAD_ATTRIBUTE_HANDLE_LIST`, and then signals its one-use delegation;
  only the final root release removes the owner record. Delegation setup,
  temporary child-environment mutation, process creation, and restoration must
  share one exception-safe cleanup boundary on both PowerShell 7 and Windows
  PowerShell 5.1.
- Script-launched build and test process trees must use a kill-on-close Job
  Object. Before starting, the owner must reject residual `MSBuild.exe`,
  `cl.exe`, or `link.exe` processes whose command line targets the repository.
  A non-empty lock marker or stale owner record marks the shared artifacts as
  contaminated. Tests, direct MSBuild, clean-only builds, incremental builds,
  targeted rebuilds, and rebuilds for another configuration or platform remain
  blocked. Only a successful serialized full-solution `build.ps1 -Rebuild` for
  the recorded configuration and platform may clear the marker; a legacy marker
  without scope requires a successful full-solution rebuild. Source/build inputs
  must not be edited while an artifact operation is active.

## Self-Test Roots and UI Navigation

- `REDSALAMANDER_TEST_ROOT` is the runner-owned sandbox base for all harnesses.
  Local runner scripts may supply it as a relative path, but the runner must
  normalize it to an absolute path before launching child processes. The runner
  also supplies `REDSALAMANDER_TEST_RUN_ID`, and native self-tests resolve their
  artifact root as
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\selftest`. The legacy
  `REDSALAMANDER_SELFTEST_ROOT` variable is compatibility-only for deliberate
  override launches, including the reviewed quarantine repair lane; normal
  runner execution clears inherited values before child tests start.
- `C:\RSPerf` is the sanctioned short NTFS `REDSALAMANDER_TEST_ROOT` override
  for local evidence that needs NTFS semantics or a shorter absolute path than
  the workspace root provides. New evidence must not use
  `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox` or `C:\RST` for that role.
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
  use the canonical Commands translation-unit `DirectedSelfTestInputWarning`
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
- Preferences category-switch coverage keeps a three-render ceiling and records
  the category host's invalidation delta for the same settled interval. Each
  measured render must be attributable to those invalidations plus at most the
  explicit page-switch redraw; increasing the ceiling without that cause evidence
  is not an acceptable flake workaround.
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
