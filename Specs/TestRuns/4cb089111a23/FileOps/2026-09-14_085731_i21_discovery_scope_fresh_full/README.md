# I21 discovery scope and leaf progress: Fresh Full closeout

Source: `63d78f55bdcdeca458fc8818b4b503f840c7126e` plus the uncommitted I21 working tree.
Governed run: `20260914T085731Z-87148-52a2518b4e3b4a0cb5009b7d50e74e3b`.
Command: `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -TestRoot Z:\RedSalamander.Perf -AllowAlternateVolumeTestRoot`.
Profile: test-enabled x64 Debug; the full-solution rebuild reported zero warnings and zero errors.
Accepted build receipt: `4dafa4879fbb4c219065d67456e40af368cdbf9bef2c5bfefd4b0a1f4ef59dd6`.
Workspace snapshot: `4390c892c834208e1686a01e5af2fb89fec8f22f550c0198c1dd7c73f79ab813`.
Validation plan digest: `58336f513976e44ce58d3aabd7df254fc5020e40c27108305ad67ae98bf6d344`.

## Verdict

Run status is **passed** in `Fresh` mode for suite `Full`. All 20 plan entries executed and were
promoted; none were reused, provisional, invalidated or incomplete. The runner recorded **2,130
passes, zero failures, 54 declared skips**, 2,184 total, in 68 minutes 15.7 seconds. Flaky,
regression, isolation-suspect and unclassified-failure classifications are all zero, and no
quarantine entry was active or repaired. This is one Fresh Full invocation, not a composite.

| Suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| Commands | 896 | 0 | 4 |
| CompareDirectories | 242 | 0 | 30 |
| FileOperations | 186 | 0 | 20 |
| ToolsPesterTests | 779 | 0 | 0 |
| ToolsPesterBuildToolchain | 12 | 0 | 0 |

FileOperations rose from 180 to 186 passes because this change adds the six-case
`FileOpsFamily_DiscoveryScope` family.

## Required witnesses

Every case below is new in this change and registered in `kFileOpsFamilyDefinitions`, so an
unfiltered FileOps or Full run executes it.

| Case | Result | What it proves |
|---|---|---|
| `DiscoveryScope_LeafCopy` | Passed, 922 ms | At a checkpoint held before any payload, the live task reports one file of exactly 3,146,521 bytes with discovery closed, and the popup's own snapshot shows a fixed full-file denominator, one item, a determinate meter, no discovery activity and a known-total item header. |
| `DiscoveryScope_LeafMove` | Passed, 797 ms | The same through exact Move source cleanup; the source is removed only after the transfer, so an early closure is not permission to delete. |
| `DiscoveryScope_ManagedDirectory` | Passed, 313 ms | Nested exact source cleanup during a cross-provider managed directory Move does not close the walker's root. |
| `DiscoveryScope_NativeRenameMerge` | Passed, 250 ms | A same-endpoint directory rename refused by an occupied destination hands discovery to the merge walk instead of latching the root closed. |
| `DiscoveryScope_MixedRoots` | Passed, 750 ms | A leaf plus a directory, in both selection orders, close the aggregate once on the exact sum. |
| `DiscoveryScope_ZeroByteAndVerifiedLeaf` | Passed, 1,500 ms | A known empty file is a known zero total with finite determinate progress, and a verification-enabled same-endpoint Copy reaches the same leaf rule through a different bridge admission. |

## Controls that had to stay green

| Case | Result | Why it is the control |
|---|---|---|
| `Fairstream_DiscoveryAheadOverlapsTransfer` | Passed, 438 ms | Wide-tree overlap still latches `FirstMutationBeforeClose`; the leaf rule did not turn discovery into a preflight barrier. |
| `R4A19_DiscoveryMeasurementFacts` | Passed, 938 ms | Recursive-tree discovery totals and bounded queue depth unchanged. |
| `R4A19_DiscoveryProviderControls` | Passed, 3,578 ms | Delayed Dummy and serialized fake-MTP discovery controls unchanged. |
| `R4A19_DiscoveryIndependentVolumes` | Passed, 3,797 ms | Authorized Z: ReFS and D: NTFS roots overlap as before. |
| `Phase5_DiscoverySkipContinues` | Passed, 578 ms | The one-way Skip state survives. |
| `Phase8_PerItemOrchestration` | Passed, 19,047 ms | Per-item progress denominators unchanged. |
| `Beeline_SameVolumeTreeMoveIsRename` | Passed, 2,688 ms | Deferring the Native rename's closure did not change same-endpoint Move behavior. |
| `Beeline_RenameMergeSkipKeepsSourceFolder` | Passed, 266 ms | Rename-merge conflict handling unchanged. |

## Suite-wide invariant

`perf/discovery-scope-metrics.jsonl` retains every discovery metric row this run emitted for
`FileOps.Discovery.OpenUs`, `FirstMutationBeforeClose` and `GrowthAfterClose`, plus the new
`FileOps.SelfTest.DiscoveryScope.*` scenario rows.

`FileOps.Discovery.GrowthAfterClose` is the new observation-only counter for discovery reports that
added work after a task's totals were already final. Closure is one-way, so a correct traversal
never produces one. Across **376 recorded task completions in this run the value is 0 in every
sample**. That is the whole FileOperations suite, not only the six new cases.

For contrast, the paired failing-before baselines recorded on the same machine and profile with each
fix disabled in isolation read `growthAfterClose=5` on the managed-directory scenario and
`growthAfterClose=6` on the rename-merge scenario, and the leaf checkpoint was never reached.

## Declared skips

All 54 skips are environment-gated and none were waived or deleted by this work.

- **CompareDirectories, 30.** Remote connection profiles are absent: S3, OneDrive Personal and
  Business, SharePoint, FTP, SFTP, SCP, IMAP and SMB.
- **FileOperations, 20.** Eighteen `Phase16_Remote*` cases for the same absent profiles, plus the two
  `Phase15_FileSystem7z*` cases owned by Compare.
- **Commands, 4.** Two explicit ViewerSpace opt-ins, and two clipboard cases,
  `cmd_pane_selection_save_restore` and `cmd_pane_copy_text`, skipped because the OS clipboard was
  unavailable to the run (`OpenClipboard error=5`). The clipboard skips are an environment condition
  on this machine during this run, not a product change; nothing in this change touches clipboard
  code. The September 8 baseline recorded two Commands skips rather than four.

The runner also reports 30 pre-existing test-root layout audit residuals. They are preserved in the
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
| `fileops/results.json`, `fileops/trace.txt` | FileOperations suite detail, including every `DiscoveryScope_` case. |
| `perf/discovery-scope-metrics.jsonl` | Discovery metric rows for this run, machine-bound to profile `4cb089111a23`. |
