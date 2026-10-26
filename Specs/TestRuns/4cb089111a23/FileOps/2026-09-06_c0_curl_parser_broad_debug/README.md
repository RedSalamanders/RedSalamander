# C0 parser repair: broad Debug

Run: `c0-curl-parser-broad-debug-20260906-r1`, from `3917fef5` plus the exact
candidate patch/hashes in `../2026-09-06_c0_curl_parser_lookup_debug/`.
Receipt: `4ce3b8d38dad1e67204b067e6e2a6d1c594fa2da3c1972e53c9b8d91cab96f69`.
Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_
-RunId c0-curl-parser-broad-debug-20260906-r1`.

Passed 17 / failed 0 / skipped 0; Copy 142/142 across all 46 scenarios,
lookup 294/294. Both wide cases publish all 4,097 files. Parallel Copy takes
60,027,837 us; serial Copy takes 108,325,842 us, below its unchanged 120-second
deadline. The preceding committed broad Debug run copied 4,041/4,097 and timed
out at 119,997,810 us; this is a correctness improvement, not an equal-work
serial throughput comparison. Parallel previously took 64,002,003 us.
All deep Copy/Move controls remain green. The suite took 258.3 seconds; runner
wall time was 264.5 seconds. These are same-machine single-run measurements.

The parser-only cases validate 262,208 rows each: Unix present/missing
590,259 / 566,127 us; DOS present/missing 537,082 / 474,952 us. All pass.
No stale listing cache, admission relaxation, larger deadline, socket policy,
or fixture reduction was introduced.

The first launch attempt (without `-r1`) stopped before running tests because
the sandbox denied the runner's residual-process inspection. The successful
launch used approved process-inspection access and the unchanged safety guard.
This broad pass is C0 evidence, not final Fresh Full or closure of all C0 routes.
Adjacent manifests bind exact raw copies and the ordered compact metric stream.
