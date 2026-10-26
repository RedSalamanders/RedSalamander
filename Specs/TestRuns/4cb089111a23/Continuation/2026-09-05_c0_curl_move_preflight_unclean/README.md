# C0 native Move preflight — unclean initial baseline

Starting implementation: `75b20b3e`, with test/registration/spec additions only.
Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlMovePreflightTruth -FailFast -RunId c0-curl-move-preflight-red-debug-20260905`.

Build completed in 3:07, zero warnings/errors; receipt
`677fc9aed63413f2b4fd0ae391b7076e765bad0fa67a0370bc3a609a42d8802a`.
Source snapshot: `9039fe8ff1290a0290a7fb1d008509ff2a4a9c18274464fe27be4212029b77e5`.
The process (PID 126084) exited without a FileOps results artifact. The runner
correctly reports failure/no usable coverage; this is not a passed or completed
behavior baseline. The eight recorded fixture rows prove shallow and depth-80
single/bulk Move at concurrency 1/4 completed with exact contents and receipts.
The next scenario in source order is depth 129. No completion row exists for it.

No exception code/stack was recovered from the owned run or targeted marker/report
checks. Deep native preflight/Copy recursion is a suspect, not a proven diagnosis.
Do not invent passed/skipped case records for work the interrupted process never
reported. The next diagnostic run moves depth refusal last and adds a content-free
scenario-begin metric so the other safety witnesses can report first. Production
behavior remains unchanged for that diagnostic baseline.

The four raw files here are exact SHA-256-verified copies from the initialized
`C:/RedSalamander.Perf` run. No dump or test payload is promoted. Retained-run
disk-audit findings remain in the wrapper; no existing runs or user data were
deleted. The original run is retained and must not be represented as qualification.
