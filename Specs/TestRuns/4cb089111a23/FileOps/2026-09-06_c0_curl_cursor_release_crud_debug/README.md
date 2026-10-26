# Bounded cursor release: Debug CRUD

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete
-RunId c0-curl-cursor-release-crud-debug-20260906`.

Result: 3 passed, zero failed/skipped, 3.5 seconds. The owning sequential batch
returns exit 1 because the preceding broad C0 run failed; this CRUD result is
independently green and does not erase that failure. Source/build match the sibling
Debug lookup archive and exact frozen Release candidate. Results/traces/fingerprints
are byte-exact; metric selection retains raw hashes and original row order.

Only owned loopback fixtures ran. No real profile mutation or normal settings change.
