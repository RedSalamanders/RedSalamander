# E1 named bridge policy Release evidence

## Identity and result

- Scenario: exact `Phase11_BridgePipelineDummyToDummyPerf` Release comparison
  for the E1 named bridge/per-item policy extraction.
- Machine hash: `4cb089111a23`.
- Runner ID: `20260831T030708Z-23064-e3ee95b1b9aa4dcd90102aef6fd68511`.
- Result: `3/3` passed, `0` failed, `0` skipped; disk audit clean.
- Test-enabled Release build receipt:
  `00f7ac9fcbbb76b97d4fded3c9a86012ce7721d0bdd9df6e80f755cdb507bc46`.
- Source snapshot:
  `ab0582654de7cd6dc93bc0f90b3db1e66786c3e0aaf181e2222f9da69679bfa8`.
- The binary embeds repository HEAD
  `0e24444c392f4d00ba706cabf97cab5ff7ab6ef7`; the receipt source snapshot also
  binds the uncommitted second E1 extraction that produced this evidence.

## Same-machine comparison

The before run is
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-27_011215`, built from
`3a652074ec08b4efbacd22a23cbb55f89e72d89f` on the same machine and with the
same Release scenario.

| Measurement | Before | E1 candidate | Delta |
|---|---:|---:|---:|
| Exact pipeline case | 130,890 ms | 128,344 ms | -2,546 ms (-1.9%) |
| Three-case run | 132,344 ms | 129,609 ms | -2,735 ms (-2.1%) |
| In-run bridge baseline | 130,484,000 us | 127,844,000 us | -2.0% |
| In-run candidate slice | 250,000 us | 281,000 us | +12.4% |
| Scheduler `ProcessIndexUs` p95 | 2,895,140 us | 3,145,325 us | +8.6% |

The whole scenario improved while all scenario hard gates passed. The in-run
candidate slice is a single sample, and scheduler p95 has only 18 samples, below
the repository's default p95 quality minimum. Those two increases are retained
as diagnostics, not treated as decision-grade regressions. The named policy adds
one virtual dispatch per item, not per copy chunk.

Bridge cardinalities were unchanged: scheduler dequeue attempts/successes and
effective concurrency remained `8/8/8`; admission depth remained `4`; directory
ensures remained `1`; all four files were admitted and started before producer
completion; traversal retained entries remained `4`; stage and deferred-link
retention remained zero. Queued-path and metadata high-water measurements were
effectively identical (`1,192`/`1,818` bytes in the first candidate pass).

## Correctness coverage

- Debug Phase 11 family:
  `20260831T022015Z-59396-357073a4d4cd49629e6ee23c494c6b06`, `9/9` passed.
- Debug broad FileOps:
  `20260831T022630Z-74156-5eaeffbad6494ac8b2f42b0da2c358b3`,
  `123/143` passed with `20` documented environmental skips.
- Release Phase 11 family:
  `20260831T024639Z-89372-1db801e718be4d359527b89ab26f2aa3`, `9/9` passed.
- Release broad FileOps:
  `20260831T025227Z-50900-9e6f613fecac4ae6b1430a3335768336`,
  `122/143` passed with `21` documented environmental/Debug-only skips.
- Full test-enabled Debug and Release solution builds completed with zero
  warnings/errors; receipts were
  `99c986f465fdd54eb2cf55a7f2ea646de0b19f94b867dc2a7d8002fc21bc84dc`
  and the Release receipt above.
- The compiled in-product source guard passed in Debug
  (`20260831T031902Z-68696-f55d88f2a71b4c88bb5133a59ec8f517`) and Release
  (`20260831T031902Z-88952-344258c6383f4709b6f364a98de2537a`), both with a
  clean disk audit.

## Artifact integrity

`perf/perf_metrics.jsonl` is the unmodified compact runner archive: `515,832`
bytes, `1,348` nonblank rows, SHA-256
`6681b31e0d9fbfb8eeb7ee191107e2c014eca2afccb48a6d1e63b6b75a9a3e94`.
The JSON results and traces beside this README are copied unchanged from the
same runner artifact.
