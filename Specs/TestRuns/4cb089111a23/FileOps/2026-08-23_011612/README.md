# File Operations P1.4 Debug closeout

This governed x64 Debug run closes P1.4 identity-owned staging and exact
publication on machine `4cb089111a23`.

- Result: 111 passed / 0 failed / 27 skipped.
- Duration: 921,970 ms.
- Build attestation: `962f2fd6b8a33c37187e7e8edc438a560344f69950e79aa9791ced10a51bb1b1`.
- Governed run id: `20260822T230027Z-79788-952f58990c5e42d797b4f4d3674c55a5`.
- Embedded source commit: `0653aaf25acf51d95c44de12cab129a7a3423a30`, plus the P1.4 working-tree implementation and specification diff archived by this commit.
- Suite: complete File Operations aggregate, x64 Debug.

The passing cases cover exact concurrent-replacement publication, local
identity-owned staged overwrite, CopyOnly source retention, and the invariant
that CopyOnly never enters source cleanup. Seven skips are reviewed P2/P3
deferrals, two are 7z cases owned by Compare coverage, and eighteen require
remote connection profiles that were not present. No throughput improvement is
claimed by this archive.

## Compact metric manifest

The mixed application metric stream exceeded the checked-in archive limit. The
original remains in the governed test sandbox at
`.build/TestSandbox/runs/20260822T230027Z-79788-952f58990c5e42d797b4f4d3674c55a5/artifacts/selftest/last_run/perf/perf_metrics.jsonl`
and is bound by `perf/perf_metrics_summary.json`.

| Artifact | Bytes | Rows | SHA-256 |
|---|---:|---:|---|
| Raw `perf_metrics.jsonl` | 134,294,886 | 361,850 | `bbcc3b960d0546e0552e976c31da6e8bd64130658af21a191a83ef185d3ad06d` |
| Checked-in claim rows | 5,952 | 12 | `e3227edcd2b2c03235ed365a13b7a800d0416bef1d329b4bcc3e8696b1f44451` |

`perf/perf_metrics_claims.jsonl` preserves the selected rows in their original
order. It contains P1.4 safety claims plus the Recycle batching measurements
used by the still-active regression case. Behavioral JSON results and traces
beside this README are copied unchanged from the same governed run.
