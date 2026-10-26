# C0 native Curl Delete repair — Debug candidate

- Source: `58f22d0e` plus working receipt repairs, native Curl root/cancel fixes,
  tests and owning specs. Build receipt:
  `dcc885a721edc04e4bf195d5d845b8b9598059db47b38a8fcd0bf0ee4ee66174`.
- Command: `Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-native-delete-candidate-debug-20260905`.
- Build: zero warnings/errors. Harness: 7/7 passed, zero skips. Native Curl case:
  29 checks passed, zero failed. It ran on an owned worker with its DLL retained until join.
- Results, trace, aggregate, and the complete 924,119-byte metric stream are unchanged
  copies from the run-owned `artifacts/selftest/last_run` directory.

Compare the same-machine Debug baseline `2026-09-05_c0_curl_native_delete_red_debug`:

| Protected behavior | Before | Candidate |
|---|---|---|
| Recursive server-root Delete, single/bulk, concurrency 1 | Four destructive commands; both files lost | Zero destructive commands; both files retained |
| Nonrecursive root Delete, single/bulk, concurrency 1/4 | One RMD command | Zero destructive commands |
| Parallel recursive root Delete | Refused without mutation | Still refused without mutation |
| Single Delete stalled first LIST | No cooperative return in 2,515 ms; fixture unblock required | Cooperative cancellation in 1,000 ms |
| Bulk Delete stalled first LIST | Cooperative cancellation in 1,000 ms | Cooperative cancellation in 984 ms |

All eight root rows pass. Eight new positive-control rows each prove successful
selected-subtree deletion (one DELE, two RMD) and a byte-checked surviving sibling.
Both listing-cancel rows report zero mutations. Differences near one callback tick
are not a percentile or throughput improvement claim. The earlier Google Drive and
Local mutation-receipt cases remain green in this expanded matrix.

This does not prove deep/wide traversal, real FTP/SFTP/SCP endpoints, safe ambiguous
mutation replay, C1–C5 completion, or final Full qualification. The disk audit reports
30 retained older-run issues; no older evidence was deleted and no clean-root claim is made.
