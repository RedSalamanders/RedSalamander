# Actual Curl socket trace — Release control

Diagnostic checkpoint based on `648b5d2a`, not a transport repair or C0 closeout.
The exact tested six-file delta and byte hashes are in `tested.patch` and
`tested-source.json`; later plan/metric-documentation changes are not runtime changes.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-socket-trace-lookup-release-20260906`.

Result: native 3/0/0, lookup **271/271**, wrapper 22.1 s. Eighteen parser/privacy
controls and the real refused-cursor witness pass. The two cursor failures retain
owned remote port 57103 and local ports 57104/57108; all six connection failures
(cursor 2, size 3, quote 1) report OS 10061. No raw diagnostic text is retained.
Policy 197/197, no skips. Full test-enabled Release build: 255638 ms, zero warnings
or errors; receipt `d2d3df18ffaff9b4082e3821cf687d5be8cfaa5c0773162c53f2b19a25de1750`,
source snapshot `8ee703718c2a62898de4495d1811ff4f1fd061f77c7b65bdb7e21ab3c066c6cc`.

The three archives retain exact native/wrapper results and ordered digest-bound
metric projections. The wrapper's 141 disk-audit findings are retained old run
directories, not a clean-sandbox claim; no evidence was deleted. Debug remains at
the previous 252-control checkpoint until rebuilt. See the sibling broad archive.
