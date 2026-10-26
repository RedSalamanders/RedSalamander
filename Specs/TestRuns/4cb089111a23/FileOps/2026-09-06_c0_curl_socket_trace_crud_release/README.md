# Actual Curl socket trace — Release CRUD regression

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-socket-trace-crud-release-20260906`.

Native **3/0/0**, wrapper 2.9 s. Same receipt/source snapshot as the sibling lookup
archive. Exact native/wrapper results, traces, provenance and ordered metric
projection are retained. 143 old-run disk-audit findings remain visible; no
evidence was deleted. Local fake FTP only, not live-provider mutation evidence.
Broad C0 remains red; this focused result cannot close it.
