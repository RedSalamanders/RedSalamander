# C0 parser repair: Release CRUD regression

Run `c0-curl-parser-crud-release-20260906`; receipt
`4357fa43d22bade20de39b19f7d13e3a47f499fc3f1ff1b6388e600cc56e14fd`.
Exact candidate patch/hashes are in the adjacent Debug lookup archive and the
Release build log is in the Release lookup archive.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete
-RunId c0-curl-parser-crud-release-20260906`.

Passed 3 / failed 0 / skipped 0; suite 2.9 seconds, runner 8.7 seconds.
This loopback CRUD witness is not a live endpoint mutation or full qualification.
Adjacent manifests bind raw copies and ordered metric selection.
