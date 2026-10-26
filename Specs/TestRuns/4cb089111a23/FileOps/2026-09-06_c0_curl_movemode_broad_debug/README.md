# Curl Move-mode repair — broad Debug regression

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_ -RunId c0-curl-movemode-broad-debug-20260906`.

Result: **17 passed / 0 failed / 0 skipped**. Native suite: 252,844 ms;
runner: approximately 259.8 seconds. Move-preflight 247/247, Copy traversal
142/142 (46 scenarios), entry lookup 294/294; deep Moves and both wide Copies pass.
Each wide Copy publishes all 4,097 files: concurrency 4 / 1 durations are
62,673,838 / 99,079,379 us. The deadline remains 120 seconds. The preceding
committed parser checkpoint measured 60,027,837 / 108,325,842 us for the same wide
corpus: directional single-run variation, not a throughput optimization claim.

Same frozen source and full Debug receipt as `../2026-09-06_c0_curl_movemode_candidate_debug/`,
which retains the exact tested patch, raw source hashes and complete build log.
This archive retains per-case verdicts, exact traces, build/source receipt identity
and digest-bound ordered metrics. The safety-mode changes preserve ordinary/default
routes and existing reply-loss, deletion, lookup and Copy truth tests.

This is loopback/fault regression evidence, not live server qualification or final
Fresh Full. C0's memory/route requirements and C1–C6 remain open.
