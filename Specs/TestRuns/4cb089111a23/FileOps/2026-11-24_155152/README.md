# I66 saved Queue/Parallel and bulk pause

The x64 Debug build at `6df121e23` completed with zero warnings and zero errors.
Build receipt `2f09b8ed3e12d094ef17694ef881bddc0bb5b36a2d8a47858e620161e05dc41f`
covers the selected RedSalamander project and its dependencies. The non-activating
`FileOps_PopupIdleCadence` run `20261124T144446Z-80952-5db2b2969dc742ec9a2d34ddbba5f4e2` passed setup, the case
and cleanup. It exercised the footer's settings write and a new owner restored from disk.
The private engine exercised group pause, prior individual pause, queue gates and later admission.

The [earlier x64 Debug run](../2026-11-24_130014/README.md) at `9d4b49d71` used the
same cadence fixtures. The paired checks retained identical synthetic tick and render counts.
The running engine recorded one additional render in this single observation.

| State | Earlier interval (ms) | Earlier ticks / renders | Current interval (ms) | Current ticks / renders |
|---|---:|---:|---:|---:|
| synthetic-running | 1208 | 11 / 11 | 1225 | 11 / 11 |
| synthetic-paused | 1209 | 1 / 1 | 1210 | 1 / 1 |
| synthetic-paused-unknown | 1208 | 1 / 1 | 1208 | 1 / 1 |
| synthetic-prompting | 1209 | 1 / 1 | 1209 | 1 / 1 |
| synthetic-queued | 1210 | 1 / 1 | 1210 | 1 / 1 |
| synthetic-completed-idle | 1209 | 0 / 0 | 1210 | 0 / 0 |
| synthetic-paused-path-detail | 1209 | 1 / 1 | 1202 | 1 / 1 |
| engine-prompting | 1208 | 1 / 1 | 1210 | 1 / 1 |
| engine-running | 1214 | 11 / 23 | 1216 | 11 / 24 |
| engine-paused | 2201 | 2 / 2 | 2204 | 2 / 2 |

| Metric | Action | Duration (µs) | First value |
|---|---|---:|---:|
| FileOps.Queue.BulkPauseUs | pause | 41 | 4 |
| FileOps.Queue.BulkPauseUs | resume | 4 | 5 |
| FileOps.Queue.BulkPauseUs | pause | 4 | 5 |
| FileOps.Settings.EnqueueUs | queue default | 1297 | 1 |
| FileOps.Settings.SaveUs | queue default | 7291 | 1 |
| FileOps.Settings.EnqueueUs | queue default | 909 | 1 |
| FileOps.Settings.SaveUs | queue default | 6481 | 2 |

The first value for bulk commands is the task count inspected. For deferred saves it is
the completed generation, and for enqueue it is the admission result. The single scoped
comparison leaves Release and interactive hit-target, focus and UIA validation open.

The raw traces remain in the marked sandbox. Retained rows follow their source order.

| Trace | Bytes | Rows | SHA-256 | Source lines |
|---|---:|---:|---|---|
| Baseline | 2553548 | 7245 | 3850be891ab1e1bcc56033641c3c682b271bb8ee4c152ab0c1499f02c1d03030 | 715, 851, 949, 987, 1157, 1183, 2113, 2687, 6433, 6758 |
| Candidate | 2658458 | 7458 | 1bb2bb3f4b63920120dfdd29bbdae63e7687de11d9a24bb2b197d14ba4d8be08 | 251, 252, 253, 260, 261, 264, 265, 695, 843, 953, 991, 1158, 1184, 2115, 2769, 6674, 6982 |
