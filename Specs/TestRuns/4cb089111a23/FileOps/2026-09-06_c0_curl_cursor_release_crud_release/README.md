# Bounded cursor release: Release CRUD

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete
-RunId c0-curl-cursor-release-crud-release-20260906`.

Result: 3 passed, zero failed/skipped, 3.0 seconds. This regression follows the
green broad C0 run on the same frozen full-solution Release receipt. Source,
tested patch and build log are retained by the sibling lookup Release archive.
Exact results/traces/fingerprints and the ordered selected metric stream are
digest-bound by `file-provenance.json` and `metric-selection.json`.

Owned loopback fixtures only. This is not live-profile mutation evidence or
final repository qualification.
