# I66 saved Queue/Parallel and bulk pause, x64 Release

The test-enabled x64 Release build at `e3e5ac280` completed with zero warnings and zero errors.
Build receipt `c34c48eddd6e8908514110e803247e018bc87afa004158c0ad639329902ca70b`
covers the RedSalamander project and its runtime dependencies. The non-activating
`FileOps_PopupIdleCadence` run `20261124T162210Z-62860-6cd00aef04f241ceacedc1d1d503eb17`
passed setup, the case and cleanup. It verified Queue/Parallel restoration from disk and
the distinction between bulk and individual pause in the task engine.

| State | Interval (ms) | Ticks / renders |
|---|---:|---:|
| synthetic-running | 1213 | 11 / 11 |
| synthetic-paused | 1213 | 1 / 1 |
| synthetic-paused-unknown | 1215 | 1 / 1 |
| synthetic-prompting | 1213 | 1 / 1 |
| synthetic-queued | 1212 | 1 / 1 |
| synthetic-completed-idle | 1212 | 0 / 0 |
| synthetic-paused-path-detail | 1214 | 1 / 1 |
| engine-prompting | 1214 | 1 / 1 |
| engine-running | 1215 | 11 / 27 |
| engine-paused | 2211 | 2 / 2 |

The bulk pause, resume, and second pause took 11, 1, and 1 microseconds.
The two settings changes took 179 and 112 microseconds to enqueue, with deferred saves of
3,691 and 4,230 microseconds. This is one scoped Release observation. The matching profile
baseline, interactive walkthrough, and full repository verdict remain open.

The full 2,836,947-byte trace has 7,926 rows and SHA-256
`2e5cb46333c52a1204b173964111da6420b95c6f4dec37a97452598928106bb4`.
The archive retains source lines 250, 251, 252, 259, 260, 263, 264, 733, 857, 967, 1005,
1186, 1212, 2215, 2810, 7029, and 7382.
