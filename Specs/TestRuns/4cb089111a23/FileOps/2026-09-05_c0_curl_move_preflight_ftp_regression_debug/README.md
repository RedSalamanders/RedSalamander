# Ordinary FTP workflow after Move preflight repair

Command:
`Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -FailFast -RunId c0-curl-move-preflight-ftp-regression-debug-20260905`.

**3/3 pass** (setup, ordinary FTP workflow, cleanup), zero skips. This preserves the
existing read/write/create/copy/rename/replace/delete regression; it is not the full
provider matrix or a Fresh Full repository verdict.

Build: 2:38, zero warnings/errors; suite 3,246 ms.
Receipt `4d4db3417bb99a0d17406e6c28cbf7a3002d3991bd1759724ab80ba90db8fe57`;
code snapshot `7b36e420365f6c376e257c41a5e2a8b0bf05a41d1f2b8f0249f1400e31d8a4e2`.

All five runtime files are exact SHA-256-verified copies from the initialized
`C:/RedSalamander.Perf` run. The full 692,167-byte / 1,899-row perf stream is retained,
SHA-256 `2855acfb70eaaadf0d6f9b05aad21a2532945bcabff2a3a8ee1a6edd51ab486c`.
The runner reports 65 retained-run disk-audit findings; no old evidence or user data
was deleted. Later documentation consolidation does not change the tested code.
