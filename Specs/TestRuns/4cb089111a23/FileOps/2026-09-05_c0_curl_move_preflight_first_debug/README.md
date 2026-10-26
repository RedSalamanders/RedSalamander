# C0 native Move preflight — first passing candidate

Production starts from regression checkpoint `062e416e` plus the iterative
preflight repair. Command:
`Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlMovePreflightTruth -FailFast -RunId c0-curl-move-preflight-candidate-debug-20260905`.

The focused harness passes 3/3 (setup, behavior, cleanup); all **97 native checks**
pass over 32 scenarios. Duration: 17,663 ms. Every depth-129 variant now returns
buffer-overflow with no mutation, rather than terminating without results.
Malformed, overlong and out-of-root records retain exact error/content/receipt
truth. Wide-tail and cancellation controls remain usable.

Build: 3:03, zero errors, four C4625/C4626/C5026/C5027 warnings for implicitly
deleted Frame copy/move members. The final candidate explicitly declares them;
this intermediate warning-bearing build is not final qualification.
Receipt `1a03f3698d644a7f07c7df59cdd02c0adc947892e580d7b49525457e2d5ec43d`;
snapshot `b98b8e8c9bc905b17834728e8ba0bd72605018e848293d066c7feef33da013ca`.

The five runtime files are exact SHA-256-verified copies from the initialized
`C:/RedSalamander.Perf` run. The full 172,001-byte / 468-row perf stream is retained,
SHA-256 `2d5eb6486f0d5ac50fc4c2f31478713d3ab32ab493262cb00871741c2471b23a`.
`build-warning.txt` preserves the actual build log. No data cleanup or push is
implied. Native Copy, accepted maximum-depth coverage, commitment-map memory and
real FTP/SSH session qualification remain open.
