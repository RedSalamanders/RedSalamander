# C0 bounded Curl Delete — ordinary FTP workflow regression

Same unchanged implementation/test/spec source as the C0 Debug and Release
qualification runs; checkpoint `5132d866` plus the recorded working changes.
Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -FailFast -RunId c0-curl-streaming-delete-ftp-regression-debug-20260905`.

Build: 2:48, zero warnings/errors. Receipt
`3e1dff4866794b80b43a0ead20433f0d54a863c42c36c6920905ba52c752dc8d`.
Runtime: 3/3, zero skips, 3.2 seconds. The actual host workflow still copies
to/from FTP with exact bytes and creates directories, renames, replaces, and deletes.
This checks ordinary usability after the shared command-operand correction.

Results/reasons, both traces, aggregate, and full ordered perf stream are copied
unchanged from the exact run-owned C:\RedSalamander.Perf artifacts. Retained-run
disk-audit limitations remain explicit and old evidence is preserved. This is
focused loopback integration, not live endpoint or Fresh Full qualification.
