# I66 Curl Rename mutation receipts

The x64 Debug `C0_CurlCommittedMutationResponseLost` selection passed setup,
all 49 provider assertions and cleanup without desktop activation in
`20261125T100716Z-83176-dd6fb33aaa70463d88e1cc33c49c39cd`.
Full solution head `8f424320c` emitted zero warnings and errors, with receipt
`bdb4880060bd4a55be52218df64232a6911748bb29a06b34cdaab2730e9de5d2`.

| Loopback FTP scope | Scenarios | Mutation requests per item | Observed duration (µs) |
|---|---:|---:|---:|
| Delete, acknowledged or lost reply | 4 | 1 | 4,070–14,317 |
| Rename to a free name, acknowledged or lost reply | 4 | 1 | 4,275–5,857 |
| Rename with an existing destination, acknowledged or lost reply | 4 | 2 | 6,900–8,156 |

The twelve `FileOps.Curl.MutationReplyLoss` rows report the command counts
and one injected primary reply loss per loss scenario. Both overwrite Rename
shapes preserved the displaced bytes under the rollback sibling, including
when the primary result and source presence were unconfirmed. The tests
verified one item receipt, the final contents, the unchanged sibling,
and no duplicate mutation or incidental `NLST`.

The original `artifacts/selftest/last_run/perf/perf_metrics.jsonl` contains
264 rows and 97,406 bytes with SHA-256
`f699dfef5b9c686cefdb93ac4e3af1029bdf212db1b002c092d0d5e7543e82ea`.
These are single runs of individual test shapes with no paired timing claim.
The standalone plugin contract checks also passed with the same receipt in
`20261125T101146Z-82512-26f933adc1414654ba19c1415b4cd394`.
Live transport qualification and host Rename admission remain open.
