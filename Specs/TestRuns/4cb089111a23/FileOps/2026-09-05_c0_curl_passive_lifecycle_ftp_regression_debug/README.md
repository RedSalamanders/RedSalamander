# Ordinary FTP regression after passive-listener fixture repair

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -FailFast -RunId c0-curl-passive-lifecycle-ftp-regression-debug-20260905`.

Based on `61815385` plus the fixture-only repair/test registration described in
the [C0 archive](../2026-09-05_c0_curl_passive_lifecycle_debug/README.md).
Build: **02:55, zero warnings/errors**.
Receipt: `e57c70dca0cbc2a252acfe40bcfa5c245e1150e3b644a7e200d95ce4aa8b2e05`.
Workspace: `685cd7bebb72f4be804e5e7287bb867fcd286d6de986503cfa1c252ca54b6805`.

**3/3 passed, no skips**: Setup, R0fCurl_FakeFtpReadWriteCreateDelete, and cleanup.
Native case: 859 ms; native suite result: 2,282 ms; wrapper wall time about 10.6 s.
This qualifies ordinary fake-FTP read/write/create/delete, not the still-failed
wide-copy gate or real FTP/SSH infrastructure.

Full raw perf: 713,884 bytes / 1,963 rows; SHA-256
`f2cf603bd2219c46342583bfcf9fcf310be9ba5bec3ae35375a51bf7b8a17255`.
All results/reasons and traces are retained. Disk audit: 74 issues, none deleted.
