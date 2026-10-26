# C0 parser repair: broad Release

Run `c0-curl-parser-broad-release-20260906`; full solution receipt
`4357fa43d22bade20de39b19f7d13e3a47f499fc3f1ff1b6388e600cc56e14fd`.
Exact candidate patch/hashes are in the adjacent Debug lookup archive; the
Release lookup archive retains the build log. Both builds share the same snapshot.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_
-RunId c0-curl-parser-broad-release-20260906`.

Passed 17 / failed 0 / skipped 0; Copy 142/142 across all 46 scenarios,
lookup 294/294. Both wide cases publish exactly 4,097 files; parallel/serial take
6,184,419 / 18,162,982 us. The previous committed Release run took
6,219,373 / 18,936,925 us. These same-machine single-run values are directional,
not a statistical performance guarantee. All deep Move controls remain green.
Suite time 95.9 seconds; runner wall time 101.7 seconds.

Adjacent manifests bind exact raw copies and ordered metric selection. This
proves the broad C0 regression selection, not all remaining C0 routes or final
Fresh Full. No deadline, safety admission, fixture size or machine setting changed.
