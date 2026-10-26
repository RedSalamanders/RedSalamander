# C0 native Move preflight — repeated Release qualification

Command: preserve the process `RSBuildEnableTests` value, set it to `true`, run
`Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter C0_ -SelfTestRepeat 3 -FailFast -RunId c0-curl-move-preflight-final-release-repeat3-20260905`,
then restore the prior environment value in `finally`.

**45/45 C0 cases pass**, zero skips: three complete repetitions, with 97/97 native
Move preflight assertions in each. The other C0 host/native cases also pass. This
is test-enabled Release evidence, not non-test Release staging or Fresh Full.

Build: 3:47, zero warnings/errors; suite 97,431 ms.
Receipt `46825a99e719b725377e64204a0a8c448066c10753eb27f4cf8104784c7c0144`;
code snapshot `7b36e420365f6c376e257c41a5e2a8b0bf05a41d1f2b8f0249f1400e31d8a4e2`,
identical to final Debug and ordinary FTP qualification.

All 96 scenario rows have exact status, contents and receipt truth. The twelve
blocked-LIST cancel samples are 56,427 / 60,513 / 57,737 / 59,478 / 63,810 /
55,005 / 55,555 / 57,815 / 52,038 / 51,568 / 54,766 / 54,322 us. No emergency
server stop supplies cancellation proof. Wide-tail cases take 454,850–691,557 us;
depth-80 controls take 615,987–947,733 us. There is no equivalent pre-fix Release
baseline; do not claim a cross-configuration throughput improvement.

All 96 traversal/retained pairs fit 129 frames, 81,024 path bytes and 6,612,024
metadata-reservation bytes. The whole-tree commitment map and opaque transport
memory/session limits are not qualified by these counters. Native Copy discovery
and success at the admitted maximum depth remain separate open work.

Four non-perf runtime files are exact SHA-256-verified copies from the initialized
`C:/RedSalamander.Perf` run. Its full 8,655,921-byte / 24,302-row perf stream stays
local; all 1,338 matching Curl/C0 rows are retained here in original order.
`metric-selection.json` binds the full original stream digest and line numbers.
Original failures and retained-run disk-audit findings are not erased.
