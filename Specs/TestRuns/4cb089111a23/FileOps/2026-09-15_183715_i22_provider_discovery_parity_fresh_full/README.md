# I22 provider discovery parity on direct routes: Fresh Full closeout

Source: `41f288e99`, a clean working tree.
Governed run: `20260915T183715Z-40020-dd9d631f292f4b3b9ea86d768f8b1e9e`.
Command: `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -AllowAlternateVolumeTestRoot`.
Profile: test-enabled x64 Debug; the full-solution build reported zero warnings and zero errors.
Accepted build receipt: `3ef95305a76395ca79e0fd99f531db4eb906c40e30ca44fa56a2340592bb9a37`.
Workspace snapshot: `8575087b590697a562c18185f1426dec677922eb20950e416fc9dc593cc5323b`.
Validation plan digest: `58336f513976e44ce58d3aabd7df254fc5020e40c27108305ad67ae98bf6d344`.

This is the gate the plan's record left outstanding after its own Fresh Full on `LT-PF5VDAGE` failed for
machine conditions (no symlink privilege, CRLF working copies of `eol=lf` test files, laptop loopback
timing). The record recommended finishing it on this machine, `4cb089111a23`, where I23's gate passed the
day before with the same command. The symlink privilege was confirmed by probe before launch.

## Verdict

Run status is **passed** in `Fresh` mode for suite `Full`. All 20 plan entries executed and were promoted; none were reused, provisional, invalidated or incomplete. The runner recorded
**2,144 passes, zero failures, 52 declared skips**, 2,196 total, in 68 minutes 7.4 seconds.
Flaky, regression, isolation-suspect and unclassified-failure classifications are all zero, and no
quarantine entry was active or repaired. This is one Fresh Full invocation, not a composite.

| Suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| Commands | 898 | 0 | 2 |
| CompareDirectories | 242 | 0 | 30 |
| FileOperations | 193 | 0 | 20 |
| ToolsPesterTests | 784 | 0 | 0 |
| ToolsPesterBuildToolchain | 12 | 0 | 0 |

FileOperations rose from I23's 187 to 193 passes for the six provider cases this plan adds.

## Required witnesses

Each case drives its provider's mutation entry points against the provider's fake backend through a
recording `IFileSystemOperationControl` and asserts that the cumulative stream never decreases, that
exactly one report closes it, that the closure is last, and that it closed on the exact expected totals;
then it runs one same-endpoint leaf transfer through the host on the admitted direct route and requires
the completed task to carry a closed exact record with no mutation completed while discovery was open.
All six are registered in `kFileOpsFamilyDiscoveryScope`, so an unfiltered FileOps or Full run executes them.

| Case | Result |
|---|---|
| `DiscoveryScope_DummyDirectApi` | Passed, 156 ms |
| `DiscoveryScope_CurlDirectApi` | Passed, 328 ms |
| `DiscoveryScope_S3DirectApi` | Passed, 688 ms |
| `DiscoveryScope_GDriveDirectApi` | Passed, 500 ms |
| `DiscoveryScope_GraphDirectApi` | Passed, 328 ms |
| `DiscoveryScope_MtpDirectApi` | Passed, 172 ms |

The failing-before witnesses for these cases are recorded in the plan under "Recording fixtures (D0)":
at the delivered code, S3, Google Drive, Microsoft Drive and MTP "reported no discovery at all" on every
call, and the 2026-09-14 Curl and Dummy bulk routes "closed on wrong totals" short by their roots.

## Controls that had to stay green

| Case | Result | Why it is the control |
|---|---|---|
| `DiscoveryScope_LocalDirectApi` | Passed, 94 ms | I23's Local contract, unchanged by the provider work. |
| `DiscoveryScope_LeafCopy` | Passed, 703 ms | I21's host-side leaf rule, unchanged. |
| `DiscoveryScope_ManagedDirectory` | Passed, 281 ms | Nested suppressed scopes still do not close the walker's root. |
| `DiscoveryScope_MixedRoots` | Passed, 532 ms | The host aggregate still closes once on the exact sum. |
| `R4A19_DiscoveryProviderControls` | Passed, 3,813 ms | Re-based Dummy control: closed exact record before the first byte; fake-MTP keeps the overlap facts. |
| `R4A19_DiscoveryIndependentVolumes` | Passed, 3,843 ms | Authorized Z: ReFS and D: NTFS roots overlap as before. |
| `Phase5_DiscoverySkipContinues` | Passed, 281 ms | Re-based onto the Local walker, whose traversal is still producing when Skip is pressed. |
| `Fairstream_DiscoveryAheadOverlapsTransfer` | Passed, 578 ms | Discovery did not become a preflight barrier. |
| `Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker` | Passed, 922 ms | Runs `FileSystem.dll`'s debug self-tests in-process; the case that cascaded on the laptop without the symlink privilege. |
| `C0_CurlCopyTraversalTruth` | Passed, 188,390 ms | The Curl native safety corpus, which the laptop's loopback timed out. |
| `Beeline_SameVolumeTreeMoveIsRename` | Passed, 2,812 ms | Same-endpoint Move behavior unchanged. |

## Suite-wide invariant

`perf/discovery-scope-metrics.jsonl` retains this run's `FileOps.Discovery.OpenUs`,
`FirstMutationBeforeClose` and `GrowthAfterClose` rows, the `FileOps.SelfTest.DiscoveryScope.*`
scenario rows and the `FileOps.Mtp.Discovery.SourceLookupUs` rows, 1,167 rows machine-bound to
profile `4cb089111a23`.

`FileOps.Discovery.GrowthAfterClose` counts discovery reports that added work after a task's totals
were already final. Closure is one-way, so a correct traversal never produces one. Across **382
recorded task completions in this run the value is 0 in every sample** — the whole FileOperations
suite, not only the discovery cases.

`FileOps.Mtp.Discovery.SourceLookupUs` records the cost of describing an MTP source from the backend's
path cache before the device mutates: six lookups in this run, 33 to 128 microseconds each, every one served from the cache and none from the device.

## Declared skips

All 52 skips are environment-gated and none were waived or deleted by this work.

- **CompareDirectories, 30.** Live journal cursor unavailable (nine), an NTFS-backed root (one), SMB (one),
  and absent remote connection profiles (nineteen: S3, FTP, SFTP, SCP, SharePoint, OneDrive Personal and
  Business).
- **FileOperations, 20.** Eighteen `Phase16_Remote*` cases for the same absent profiles, plus the two
  `Phase15_FileSystem7z*` cases owned by Compare.
- **Commands, 2.** The two explicit ViewerSpace opt-ins (`cmd_viewer_space_layout_20k_visible_optin`,
  `viewer_space_perf_large_optin`).

This is the same skip shape as I23's gate on this machine the day before, which is what a change confined
to provider discovery reporting should produce.

The runner also reports pre-existing test-root layout audit residuals. They are preserved in the
aggregate and were not deleted or waived here.

## Retained artifacts

| File | Contents |
|---|---|
| `results.json` | Full aggregate case statuses and skip reasons. |
| `run-all-tests-results.json` | Runner summary with per-entry durations. |
| `run-state.json` | Validation run identity, mode, build receipt, snapshot and entry promotion. |
| `validation-plan.json` | The executed Fresh plan. |
| `workspace-final.json` | Final workspace identity binding. |
| `trace.txt` | Whole-run trace. |
| `fileops/results.json`, `fileops/trace.txt` | FileOperations suite detail, including every `DiscoveryScope_*DirectApi` case. |
| `perf/discovery-scope-metrics.jsonl` | Discovery and MTP source-lookup metric rows for this run, machine-bound to profile `4cb089111a23`. |
