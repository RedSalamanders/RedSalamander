# Operation Beeline: same-endpoint Move is a rename

- Date: 2026-09-08
- Machine profile: `4cb089111a23`
- Source base at capture: `c5a0a019aab2e7436fa4d68391aa30772f699fcc` plus the uncommitted Operation Beeline package under test
- Build: test-enabled x64 Debug, 0 warnings and 0 errors
- Build receipt: `0fabf2c625e565828cf8905b54aaaea1e935a11c710454f4d0984650761ce11f`
- Governed runner: `20260907T205843Z-46028-f4bfabc0e38e458cad88a4d92654fcac`
- Result: 166 passed / 2 failed / 29 documented skips (File Operations entry)
- Raw bounded metrics: `perf/perf_metrics.jsonl` (1,822 compacted rows, 688,789 bytes, SHA-256 `11e820f9461ff0dd596c55060ba4a428464276ac40618502d120a314ce68d9e5`), compacted from 542,478 original rows (SHA-256 `7102f727e4a017a296ceefecfa5d465daecb0fedbbad0ce8ccd7afdd8b2a1bdb`)

`fileops_results.json` is the complete per-case authority; `fileops_trace.txt` and `selftest_run_trace.txt`
retain execution diagnostics.

The Beeline claim is behavioral and route-shaped. A Move whose source and destination share one
qualified endpoint is admitted Native for every item shape (regular file, directory tree, link
object); a Native directory Move that meets an existing regular directory continues in the same
task as a rename merge (each child relocated by one provider Native rename, collisions through the
ordinary typed conflict surface, emptied source directories removed by exact cleanup); a Move never
applies the `Preserve links` / `Skip links` policy, which is now a Copy option; and the
`links.nativeMoveSemanticTransform` capability no longer exists. Copy `Skip links` skips
name-surrogate links only; a WOF-backed placeholder file copies.

Claim-bearing rows in `perf/perf_metrics.jsonl`:

- `FileOps.SelfTest.SameVolumeTreeMoveIsRename` (`shape=file+tree;identity-preserved`): the nested
  child's `BY_HANDLE_FILE_INFORMATION` identity is unchanged after the Move.
- `FileOps.SelfTest.SameVolumeTreeMoveThroughput`: `files=2000;bytes=16384000;copyMs=2016;moveMs=94`.
  The Copy of an identical tree is the byte cost the former Managed Move paid before its exact
  delete; the rename Move pays none of it. Same machine, same volume, same process, Debug build.
- `FileOps.SelfTest.LoopbackShareMoveIsRename` (`shape=file+tree;profile=local-win32-smb`): the SMB
  profile takes the same Native route when both paths live on the loopback administrative share.
- `FileOps.SelfTest.SameVolumeRenameMerge` (`shape=overwrite+skip+subtree;source-folder-kept`):
  two prompts, Overwrite relocates through the exact replace rename, Skip keeps its child and the
  source folder (`Moved; source folder kept`).
- `FileOps.SelfTest.NativeDirectoryRaceContinuesAsRenameMerge`: a destination directory created
  after admission returns the provider's known non-commit once and the item continues as a rename
  merge with the source removed.
- `FileOps.SelfTest.CopySkipLinksKeepsPlaceholders`: recorded as `skip=wof-unavailable` in this run.
  The sandbox volume is ReFS, which has no WOF driver, so the placeholder half of the Copy `Skip
  links` claim is code-covered by the case but not evidenced by this archive; the junction half is
  covered by `Phase12_ReparsePointPolicy`.
- `fileops.operation.strategy` buckets carry `move.native.same-root` for the same-endpoint trees and
  `move.managed.cross-root` only for the loopback-share and Dummy-provider destinations.

This archive does **not** claim Release-runtime, remote-credential, or production-workload
improvement; the throughput row is a same-machine Debug comparison of one 2,000-file, 16,384,000-byte
tree and its copy.

## Metric compaction

The original `perf/perf_metrics.jsonl` stream was compacted to the `FileOps.SelfTest.*`,
`fileops.operation.strategy`, `fileops.execute.native_directory_race_requalified_after_noncommit`,
`fileops.local.metadata.inspect_failed`, `FileOps.Bridge.FailurePhase`, and `FileOps.Bridge.StageDisposition`
rows in original order; per-item bridge timing counters and the high-frequency popup, progress,
identity, and discovery diagnostics are not load-bearing for this claim.
`perf/perf_metrics_summary.json` records the original and compacted row counts and digests.
