# I66 Windows SMB exact Rename on x64 Release

The no-activation `FileOps_SmbExactRenameProvider` selection passed setup, the case and
cleanup on the approved Windows SMB test share in run
`20261125T090124Z-9520-4c4be152c458495aae45647c93c03920`. A full x64 Release
solution build at `65d02321c6e1b3ba47f6f7eda937c0162003801d` emitted zero
warnings and errors. The runner verified receipt
`7bac16ed76fa97a78e5af5c646bd79c69ee57ba2946a81011b63e26e7bfa4ef7`.

The case reopened nested free-name and case-only results to verify the retained identity,
stored spelling, source path and content. The occupied destination kept its original
content. A competing watcher returned `ERROR_SHARING_VIOLATION` with the source
intact. The owned fixture was removed and the prior machine resource file was restored
byte for byte.

| Original trace row | Event | Duration (µs) | Item count | Status |
|---:|---|---:|---:|---|
| 248 | Watcher refusal | 0 | 0 | `0x80070020` |
| 251 | Single guarded rename | 2002463 | 1 | `S_OK` |
| 353 | Batch of independent nested names | 1547817 | 100 | `S_OK` |

The timings include provider binding, mutation and verification. The
[x64 Debug run](../2026-11-24_161520/README.md) observed 4,961 µs for one item and
1,001,306 µs for 100. These are individual samples from different builds and
concurrent workloads. Repeated paired timing, peak handles and host cancellation
remain open.

The original perf trace is in `artifacts/selftest/last_run/perf/perf_metrics.jsonl`
inside the completed-run archive. It has 143,732 bytes, 392 rows, and SHA-256
`0a9ed1a84d9b0cb59647246993fdb5c7ee20cddd95fef96366fec516f3a89588`.
The table retains all three `FileOps.Smb.ExactRenameControlUs` rows in order.
