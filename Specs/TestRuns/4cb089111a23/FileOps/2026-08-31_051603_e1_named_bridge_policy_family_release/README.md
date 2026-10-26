# E1 named bridge policy family Release evidence

## Identity and result

- Scenario: complete Release
  `FileOpsFamily_Phase11_BridgeAndConnections` coverage for the E1 named
  bridge/per-item policy extraction.
- Machine hash: `4cb089111a23`.
- Runner ID: `20260831T031026Z-63536-56efdd730e384d2cbe5c5269ae0a2033`.
- Result: `9/9` passed, `0` failed, `0` skipped in `326,556 ms`; disk audit clean.
- Test-enabled Release build receipt:
  `00f7ac9fcbbb76b97d4fded3c9a86012ce7721d0bdd9df6e80f755cdb507bc46`.
- Source snapshot:
  `ab0582654de7cd6dc93bc0f90b3db1e66786c3e0aaf181e2222f9da69679bfa8`.
- The binary embeds repository HEAD
  `0e24444c392f4d00ba706cabf97cab5ff7ab6ef7`; the receipt source snapshot also
  binds the uncommitted second E1 extraction that produced this evidence.

## Claim-bearing measurements

The wide/many-file case created two folders and admitted all `48/48` files;
all 48 started before producer completion. Directory ensures were exactly `2`,
admission queue depth reached `23`, and the published high-water row recorded
`24` retained entries, `12,432` queued-path bytes, and `9,638` metadata bytes,
with zero traversal-limit hits. These remain far below the production ceilings
of `4,096` entries, `16,777,216` queued-path bytes, and `8,388,608` metadata
bytes. Terminal stage-retention and deferred-link-retention rows remained zero.

The exact pipeline case passed in `128,125 ms`; its in-run bridge baseline and
candidate slices were `127,782,000 us` and `281,000 us`. The companion exact
before/after archive in `../2026-08-31_050934_e1_named_bridge_policy_release`
contains the same-machine throughput comparison and diagnostic caveats.

## Compact artifact manifest

The runner's raw JSONL remains beneath the immutable local run directory at
`D:/RedSalamander.Perf/runs/20260831T031026Z-63536-56efdd730e384d2cbe5c5269ae0a2033/`.
It exceeded the repository's checked-in evidence limit and was compacted
deterministically:

| Artifact | Bytes | Rows | SHA-256 |
|---|---:|---:|---|
| Raw `perf_metrics.jsonl` | 34,539,803 | 92,840 | `cde31832178b39a13565349fb29a84a69b258821dce579e550a5089adbb1b3ff` |
| Checked-in compact JSONL | 969,550 | 2,266 | `36a95abd11dec7dda7f8af13dc3bd618f3fdfe63cd6c7cf04478bc850124eb33` |

Compaction preserves rows in original order and retains every
`FileOps.Bridge.*`, `FileOps.Scheduler.*`, and `FileOps.SelfTest.Bridge*` row,
plus the operation strategy, operation summary, and result-item rows. It removes
unrelated startup, rendering, popup-frame, and animation streams. The JSON
results and traces beside this README are copied unchanged from the same runner
artifact.
