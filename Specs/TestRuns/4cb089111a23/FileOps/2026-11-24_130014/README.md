# I66 Preparing status and popup cadence

The test-enabled x64 Debug candidate at 9d4b49d71 used build receipt
43a396a5e2b5167cea0ff4a6b0182faf99f8f0ab718727869e833ec1e1e24d20. The fixed-clock Preparing status, short chip and queue precedence
passed in FileOps_PopupIdleCadence run 20261124T115346Z-19596-42d9378fb9594838b3962c7997e4d76a. The real Local
directory Delete reservation checks passed in FileOps_ReviewRemediationContracts
run 20261124T120125Z-12688-ad787f3e6bed48438c508063879e91db with the same receipt.
Each focused selection passed its main case, setup and cleanup. Repository coverage remained NOT_EVALUATED.

The same-machine x64 Debug cadence baseline at 0211122e8 came from
run 20261124T091911Z-69464-1f7f62f41fb24db186fcc04be6f8fa9d with receipt
30e70fc85e83d409cf324d7d60cb2f17aaf93b1d65b04b0fa9ecda6bd23a09ad.
These pairs use unchanged synthetic fixtures and the same timer and render counters.

| State | Baseline interval (ms) | Baseline ticks / renders | Candidate interval (ms) | Candidate ticks / renders |
|---|---:|---:|---:|---:|
| Running | 1210 | 11 / 11 | 1208 | 11 / 11 |
| Paused | 1209 | 1 / 1 | 1209 | 1 / 1 |
| Paused, unknown total | 1208 | 1 / 1 | 1208 | 1 / 1 |
| Prompting | 1211 | 1 / 1 | 1209 | 1 / 1 |
| Queued | 1209 | 1 / 1 | 1210 | 1 / 1 |
| Completed idle | 1211 | 0 / 0 | 1209 | 0 / 0 |
| Paused with path details | 1208 | 1 / 1 | 1209 | 1 / 1 |

FileOps.Admission.TaskLayoutBytes recorded 27088 bytes before and 27104 bytes
after. This is measured object size, separate from allocation count or total process memory.
The table is a single comparison of synthetic cadence. It leaves Release, live transfer,
allocation and credential wait validation open.

The marked sandbox retains both raw perf/perf_metrics.jsonl files under the run IDs above.
The baseline contains 2559766 bytes and 7264 rows, SHA-256 efb016afb32a251c11bbb22cf6121077ad4ff463af152bed6abd8fbd42409041.
It contributed 8 rows between original lines 728 and 2130.
The candidate contains 2553548 bytes and 7245 rows, SHA-256 3850be891ab1e1bcc56033641c3c682b271bb8ee4c152ab0c1499f02c1d03030.
It contributed 8 rows between original lines 715 and 2127.
The retained rows follow their original order within each run.
