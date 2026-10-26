# Curl Move-mode baseline — Debug

Base: `f0710f9ba47d3fc0a45413fffa3ffa84ac038074`. The exact four-file
test/spec delta and raw source hashes are in `tested.patch` and `tested-source.json`.
Full test-enabled Debug receipt:
`857bcbfec7fcd6b0004aa2dc1f443df96cd88ff53a4bc3e063d75d7a3178c131`;
snapshot `e266fcdc8c479e968b0a5a022a06beaadd1bb0f171c5bed94a4c205ff2995236`.
Build succeeded in 183,731 ms with three C4061 warnings in the new fixture switch.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlMovePreflightTruth -RunId c0-curl-movemode-baseline-debug-20260906`.

Result: 2 passed / 1 failed / 0 skipped; native assertions **153/241**, 88 failed.
The original 97 assertions remain green. Of the 48 new mode scenarios, all setup
checks and four same-endpoint file-Move controls pass. Forty refusal scenarios
fail content/result and callback checks (80 assertions), reproducing ignored
native-only/invalid/non-Move options. Four same-endpoint directory controls fail
both checks (8 assertions): the fake endpoint's RNFR/RNTO only recognize files.
This is a fixture limitation to repair, not evidence that production directory
rename is broken. The native case took 23,141 ms. No crash or hang occurred.

Retain this red baseline. The candidate must reject invalid modes before provider
I/O, refuse native-only cross-endpoint relay without mutation, retain independent
item truth, and keep valid same-endpoint native file/directory Move available.
It must explicitly handle the switch's enum values and repair directory rename
in the fake server. Whole-tree Move memory and real-server qualification remain open.

Raw copies and ordered metric selection are digest-bound by the accompanying
provenance manifests. Build-log normalization removes trailing horizontal
whitespace only. This is focused loopback evidence, not Full or live-server evidence.
