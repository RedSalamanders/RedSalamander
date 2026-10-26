# I23 direct provider API discovery contract: Fresh Full closeout

Source: `caad89d8`, a clean working tree.
Governed run: `20260914T181429Z-85540-4e37ec54d6284ff19dd756671a4bb017`.
Command: `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -AllowAlternateVolumeTestRoot`.
Profile: test-enabled x64 Debug; the full-solution build reported zero warnings and zero errors.
Accepted build receipt: `b38e89af5defde81ad2d3867b7dfc6f66681e2d9eceb1572eec2dc6c46d0e0f2`.
Workspace snapshot: `f60cbfa0f33feb421c3b22459e1b607c4e8895b79cdde3316eff9ad0a100e7f8`.
Validation plan digest: `58336f513976e44ce58d3aabd7df254fc5020e40c27108305ad67ae98bf6d344`.

## Verdict

Run status is **passed** in `Fresh` mode for suite `Full`. All 20 plan entries executed and were
promoted; none were reused, provisional, invalidated or incomplete. The runner recorded **2,136
passes, zero failures, 52 declared skips**, 2,188 total, in 66 minutes 27.7 seconds. Flaky,
regression, isolation-suspect and unclassified-failure classifications are all zero, and no
quarantine entry was active or repaired. This is one Fresh Full invocation, not a composite.

| Suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| Commands | 898 | 0 | 2 |
| CompareDirectories | 242 | 0 | 30 |
| FileOperations | 187 | 0 | 20 |
| ToolsPesterTests | 782 | 0 | 0 |
| ToolsPesterBuildToolchain | 12 | 0 | 0 |

FileOperations rose from I21's 186 to 187 passes for the one case this change adds.

## Required witness

| Case | Result | What it proves |
|---|---|---|
| `DiscoveryScope_LocalDirectApi` | Passed, 125 ms | Drives `CopyItems`, `MoveItems` and `DeleteItem` against the provider's own API with no host in the loop, through a recording `IFileSystemOperationControl`. For each call it asserts the cumulative stream never decreases, that there is exactly one closure, that the closure is the last event, and that it closed on the exact expected totals. |

Registered in `kFileOpsFamilyDiscoveryScope`, so an unfiltered FileOps or Full run executes it.

Against the original provider the same case named all three defects in a single run:
`CopyItems closed on wrong totals (bytes=32780/131115 files=4/7).` |
`MoveItems cumulative totals went backwards at report 1 (49165/1/0 after 0/0/1).` |
`DeleteItem closed its governed scope 0 times.`

## Controls that had to stay green

| Case | Result | Why it is the control |
|---|---|---|
| `DiscoveryScope_LeafCopy` | Passed, 750 ms | I21's host-side leaf rule is untouched by the provider change. |
| `DiscoveryScope_LeafMove` | Passed, 641 ms | Exact Move source cleanup unchanged. |
| `DiscoveryScope_ManagedDirectory` | Passed, 281 ms | Nested suppressed scopes still do not close the walker's root. |
| `DiscoveryScope_NativeRenameMerge` | Passed, 203 ms | Deferred closure on a provisional rename unchanged. |
| `DiscoveryScope_MixedRoots` | Passed, 625 ms | The host aggregate still closes once on the exact sum. |
| `DiscoveryScope_ZeroByteAndVerifiedLeaf` | Passed, 1,438 ms | Known-zero totals and the verification admission unchanged. |
| `R4A19_DiscoveryIndependentVolumes` | Passed, 4,032 ms | Authorized Z: ReFS and D: NTFS roots overlap as before. |
| `Fairstream_DiscoveryAheadOverlapsTransfer` | Passed, 625 ms | Discovery did not become a preflight barrier. |
| `Beeline_SameVolumeTreeMoveIsRename` | Passed, 2,844 ms | Same-endpoint Move behavior unchanged. |
| `C0_CurlCopyTraversalTruth` | Passed, 186,375 ms | The Curl native safety corpus, in line with its 186.0 s and 187.8 s baselines. |

## Suite-wide invariant

`perf/discovery-scope-metrics.jsonl` retains this run's `FileOps.Discovery.OpenUs`,
`FirstMutationBeforeClose` and `GrowthAfterClose` rows plus the `FileOps.SelfTest.DiscoveryScope.*`
scenario rows, 1,137 rows machine-bound to profile `4cb089111a23`.

`FileOps.Discovery.GrowthAfterClose` counts discovery reports that added work after a task's totals
were already final. Closure is one-way, so a correct traversal never produces one. Across **376
recorded task completions in this run the value is 0 in every sample** — the whole FileOperations
suite, not only the discovery cases. `OpenUs` and `FirstMutationBeforeClose` carry the same 376
samples, so the counter is present on every completion rather than only where it is convenient.

## Declared skips

All 52 skips are environment-gated and none were waived or deleted by this work.

- **CompareDirectories, 30.** Nine for an unavailable live journal cursor, one for an NTFS-backed
  root, one SMB, and nineteen for absent remote connection profiles (S3, FTP, SFTP, SCP, SharePoint,
  OneDrive Personal and Business).
- **FileOperations, 20.** Eighteen `Phase16_Remote*` cases for the same absent profiles, plus the two
  `Phase15_FileSystem7z*` cases owned by Compare.
- **Commands, 2.** The two explicit ViewerSpace opt-ins. I21's run additionally skipped two clipboard
  cases because the OS clipboard was unavailable to it; the clipboard was available here, which is
  why Commands reports 898 passes against I21's 896.

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
| `fileops/results.json`, `fileops/trace.txt` | FileOperations suite detail, including `DiscoveryScope_LocalDirectApi`. |
| `perf/discovery-scope-metrics.jsonl` | Discovery metric rows for this run, machine-bound to profile `4cb089111a23`. |
