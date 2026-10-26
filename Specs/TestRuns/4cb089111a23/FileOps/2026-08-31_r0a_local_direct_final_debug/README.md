# R0a Local direct-final Debug correctness

- Build: final test-enabled x64 Debug, receipt `dd1fc37bc23b0ceadcc8a2b1c4ebf2790e110c98571d6c8d5f015a868ed6c1e6`
- `race/`: `Floodgate_LocalCopyNewNameConcurrentReplacementSurvives`, PASS `3/0/0`, disk audit `0`
- `abort/`: `Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete`, PASS `3/0/0`, disk audit `0`
- `full/`: complete FileOps suite, PASS `125/0/20`, no classification or quarantine activity, disk audit `0`

The first aggregate attempt had one timing-sensitive `Phase7_ParallelCopyMoveKnobs` observer miss; the case passed immediately in isolation and the classified full rerun passed without needing a retry. The compact `full/` archive retains results and traces; its 283 MB general perf stream was intentionally not checked in.
