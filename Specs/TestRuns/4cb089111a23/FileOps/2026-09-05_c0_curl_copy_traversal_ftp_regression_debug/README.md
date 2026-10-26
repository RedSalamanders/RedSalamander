# Ordinary FTP regression after iterative Copy — passed

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -FailFast -RunId c0-curl-copy-traversal-ftp-regression-debug-20260905`.

Build 02:30, zero warnings/errors. Receipt:
`5472b956a125db72e173e7380af5c22002e6bfe37d8852fabfa9ebe9a98f5b0f`.
Workspace: `2d8e0ab053fb908e6580526cc73f650658015426fc14bf0c969ee69e44db2196`,
the same tested bytes as both candidate C0 runs. Full perf: 704,327 bytes / 1,938
rows; SHA-256 `cb35ac1b2b6df83bf23a8defd94e2d836642d77c831a7fedbe34447f8ef76700`.
The ordinary FTP workflow passes 3/3, no skips, about 3.3 s. This is a regression
gate for ordinary read/write/create/delete behavior, not proof that the new wide
Copy gate or the reliability plan is complete. Full results/reasons, traces and
the original perf stream are retained. See the adjacent candidate and broader C0
archives for the same implementation's remaining two wide-copy failures.
