# Curl Move-mode repair — ordinary FTP regression, Debug

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-movemode-crud-debug-20260906`.

Result: **3/0/0**, no skips. The ordinary read/write/create/delete case takes
1,047 ms; native suite 2,421 ms; runner approximately 9.4 seconds.
Same frozen source and full Debug receipt as `../2026-09-06_c0_curl_movemode_candidate_debug/`.
No live server was mutated. Exact traces, case results and ordered metrics retain
their raw provenance. This guards default operation compatibility, not Full closeout.
