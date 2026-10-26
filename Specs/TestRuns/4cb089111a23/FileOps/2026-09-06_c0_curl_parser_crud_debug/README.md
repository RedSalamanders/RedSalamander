# C0 parser repair: Debug CRUD regression

Run: `c0-curl-parser-crud-debug-20260906`. Receipt:
`4ce3b8d38dad1e67204b067e6e2a6d1c594fa2da3c1972e53c9b8d91cab96f69`.
Exact candidate patch/hashes and build log are in
`../2026-09-06_c0_curl_parser_lookup_debug/`.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete
-RunId c0-curl-parser-crud-debug-20260906`.

Passed 3 / failed 0 / skipped 0; suite 3.4 seconds, runner 9.5 seconds.
This loopback CRUD witness is not a live endpoint mutation or full qualification.
Adjacent manifests bind raw copies and ordered metric selection.
