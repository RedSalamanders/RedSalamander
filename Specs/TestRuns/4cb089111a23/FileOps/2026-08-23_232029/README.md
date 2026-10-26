# File Operations P2.4 bounded traversal evidence

- Runner id: `20260823T211829Z-86072-d456d6accca64c90b8c951cb2cca717a`
- Command: `./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter Phase11_BridgeMultiFolderParallelCopyInFlightLines -SkipBuild -ValidationMode Fresh`
- Debug x64 build attestation: `d6a9596154822c6016aa445368806bea6be4944aca0075fcc8036f583c802667`
- Result: 3 passed, 0 failed, 0 skipped.

The wide-shallow task published 48 files from two folders with a queue high-water of 16,
21 retained work entries, 11,634 retained queued-path bytes, 9,674 retained metadata bytes,
and zero traversal-limit hits. Its transfer was intentionally throttled to 256 KiB/s so file
execution and continuing discovery overlapped deterministically.

The follow-on deterministic depth override stopped the second top-level item with exact HRESULT
`0x800703E9`, retained the already-published first item, published nothing below the rejected
depth, recorded one limit hit at observed depth two, and spent zero time in the conflict arbiter.

`perf/perf_metrics.jsonl` is a deterministic original-order compaction containing the 16
claim-bearing operation, traversal, and self-test rows. The uncompressed local stream had SHA-256
`9e017e583a5307f0205c103d7f14e428b15fadb79b5e6d0961b6b081e06bdc88`, 15,527,471 bytes, and
42,517 rows. The retained compact stream has SHA-256
`7fd350999b5280920777b53f282063f0c2d33f7c6ec758e84cebcb5bb426dbcd`, 7,975 bytes, and 16 rows.
