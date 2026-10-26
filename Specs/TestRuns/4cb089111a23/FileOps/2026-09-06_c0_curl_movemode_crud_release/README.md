# Curl Move-mode repair — ordinary FTP regression, Release

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-movemode-crud-release-20260906`.

Result: **3/0/0**, no skips. The read/write/create/delete case takes 938 ms;
runner approximately 9.2 seconds. Same frozen source and full Release receipt as
`../2026-09-06_c0_curl_movemode_candidate_release/`, including the explicitly
retained C6 shared-DLL cleanup warning. No live server was mutated.
Exact results/traces and ordered metrics retain their raw provenance.
This is default-operation compatibility evidence, not final Full qualification.
