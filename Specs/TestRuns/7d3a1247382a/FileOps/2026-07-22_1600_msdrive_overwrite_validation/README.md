# Microsoft Drive overwrite mutation-time validation evidence

Scope: EG-MSD-01, x64 test-enabled Release, machine profile `7d3a1247382a`, 2026-07-22.

- `perf_metrics.jsonl` contains five same-binary samples. Each sample ran the complete
  `PluginContractTests.exe`; Microsoft Drive passed 202/202 on every iteration.
- `plugin-contract-debug.txt` is the complete contract-suite transcript from the final full-solution Debug rebuild;
  Microsoft Drive passes 202/202 and the executable exits successfully.
- The deterministic 16-child overwrite scenario compares the test-only pre-validation request shape with the safe
  candidate under fixed 3 ms metadata-GET latency. Median metadata time is 1,289,269 us baseline versus
  1,542,003 us candidate: 19.60% overhead, below the 30% advisory ceiling.
- Exact overwrite request counts are 83 logical / 84 total Graph GETs before validation and 99 / 100 after it: one
  added mutation-time metadata GET per overwritten child. Both variants retain 32 PATCHes and 16 cleanup DELETEs and
  produce identical destination/source identity results.
- The adjacent no-conflict merge contract remains 83 / 84 GETs in its legacy mode versus 67 / 68 with merge-local
  destination-hint reuse. Five-sample medians are 1,301,005 us versus 1,043,604 us, a 19.78% reduction, proving the
  overwrite guard does not tax missing/no-conflict children.
- Production aggregate rows record overwrite validation GET count/time and conflict count once per public move
  operation; the benchmark rows are `FileOps.MicrosoftDrive.OverwriteValidation.Baseline`, `.Candidate`, and
  `.Overhead`.
- Build logs `.build/logs/msbuild-20260722_155823_614.log` (test-enabled Release provider) and
  `.build/logs/msbuild-20260722_161458_386.log` (test-enabled Debug full solution) report 0 warnings / 0 errors.
  `.build/logs/msbuild-20260722_161313_843.log` is the clean tests-disabled production Release provider build;
  `dumpbin /exports` confirms `RedSalamanderMicrosoftDriveDebugSelfTests` is absent.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` passes 149/149 and pins the conditional mutation, race-test,
  exact request-count, and aggregate-metric wiring.

Failure-first evidence is archived separately at
`Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_1545_msdrive_mutation_red/`. The durable behavior and performance
contracts are in `Specs/FileSystem/FileSystem_MicrosoftDrive.md`.
