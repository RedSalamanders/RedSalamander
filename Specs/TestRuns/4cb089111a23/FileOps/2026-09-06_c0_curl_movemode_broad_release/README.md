# Curl Move-mode repair — broad Release regression

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_ -RunId c0-curl-movemode-broad-release-20260906`.

Result: **17 passed / 0 failed / 0 skipped**, native suite 100,344 ms;
runner approximately 107.0 seconds. Move-preflight 247/247, Copy traversal 142/142
over 46 scenarios and entry lookup 294/294 pass, including all deep Moves.
Both wide Copies publish all 4,097 files: concurrency 4 / 1 take 6,324,826 /
17,585,574 us. The preceding committed parser checkpoint recorded 6,184,419 /
18,162,982 us: directional single-run variation, not a throughput improvement claim.
No deadline increase, reduced corpus or isolated retry was used.

Same frozen source/Release receipt as `../2026-09-06_c0_curl_movemode_candidate_release/`;
the exact patch/raw source hashes are in the companion Debug candidate archive.
The Release build's one MSB3061 shared `z.dll` cleanup warning remains open under
C6 and is retained with its complete log in the Release candidate archive.

Exact traces, case results, receipt identity and ordered metric selection are
digest-bound. This broad loopback/fault run is not live server qualification or
final Fresh Full. Remaining C0 memory/route checks and C1–C6 are not closed.
