# Exclusive ownership — Debug lookup

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-exclusive-lookup-debug-20260906`.
Native 3/0/0, **273/273**, wrapper 23.5 s. Full Debug solution has zero warnings or
errors and the same source snapshot as the sibling Release lookup archive.
The complete build log and its transformation/hash provenance are retained here.
See that sibling README for scope, exact tested delta and outstanding broad failures.
148 old-run disk-audit findings remain recorded; no cleanup or live mutation.
