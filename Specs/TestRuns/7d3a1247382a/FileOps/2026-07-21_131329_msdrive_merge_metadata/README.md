# Microsoft Drive Merge Destination-Metadata Reuse Evidence

Five same-machine, same-binary test-enabled Release `PluginContractTests` runs compare the forced legacy destination refetch with the merge-local destination lookup hint. Every run passed the complete provider contract executable.

## Result

| Measure | Legacy median | Candidate median | Change |
|---|---:|---:|---:|
| Logical item-metadata GETs | 83 | 67 | -16 (-1 per child) |
| Total Graph GETs | 84 | 68 | -16 (-1 per child) |
| Graph PATCHes | 16 | 16 | unchanged |
| Fixed-latency elapsed time | 1,323,954 us | 1,059,139 us | -264,815 us (-20.00%) |

The source folder remains as documented safe cleanup debt, every source child lands at the destination with its identity preserved, and no prompt occurs in the no-conflict scenario. Focused tests in the same exported suite also cover existing-destination overwrite, Retry-to-missing, nested folders, and a late-created destination that must collide at PATCH without source loss or DELETE.

## Validation context

- Base commit: `ea1eed0697d0`; candidate source was in the working tree.
- Build: test-enabled x64 Release full-solution rebuild, `.build/logs/msbuild-20260721_130315_812.log`.
- Build result: 0 errors; 10 unrelated `ViewerPETests.cpp` unused-helper warnings (C5245/C5264). Focused Microsoft Drive Release builds were 0 warnings/errors.
- Per run: Microsoft Drive 137/137, S3 multipart 26/26, S3 directory-transfer probes 16/16, `PluginContractTests` exit 0.
- Raw evidence: `sample-01` through `sample-05` perf JSONL and assertion traces.

Only median and exact request-count claims are made. Five operation-level samples are not sufficient for a generic p95 claim under `Specs/Testing/Testing_PerformanceValidation.md`.
