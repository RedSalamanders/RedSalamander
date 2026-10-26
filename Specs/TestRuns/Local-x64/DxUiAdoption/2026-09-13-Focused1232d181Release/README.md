# Focused adoption regressions: 1232d181-x64-Release-focused

Local x64 Release at `1232d181e74ae7118d3fce7974196811b3cd4398`; no hosted CI.
Each declared case sequence in `summaries.json` runs ten times in its native self-test
suite. File Operations totals include setup/cleanup cases. Original native results
retain every case, repeat and failure reason.

| Sequence | Passed | Failed | Skipped | Run |
|---|---:|---:|---:|---|
| queue-after | 30 | 0 | 0 | 20260913T143402Z-39692-a2aca327ac874d46b00ff78045c77ba2 |
| dont-start | 30 | 0 | 0 | 20260913T143455Z-39692-bbd329bb5b3a4a22b0f3553c948f75d8 |
| navigation | 40 | 0 | 0 | 20260913T143544Z-39692-b4c573ce66704af0b135c4c5b45dad7c |

These are focused diagnostic runs, not Fresh Full qualification. A passing reduced
sequence does not explain the earlier Full failure or prove that broader process state
is clean. Expanded context and short sequences must not be presented as matched performance
workloads. `case-comparison.json` extracts counts/failures without modifying original receipts.

`retained-files.json` hashes 13 original receipts. Complete native archives stay
under `Z:/RedSalamander.Perf/evidence/I19-20260909/1232d181-x64-Release-focused`. The
[I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md)
owns the remaining x64 qualification; ARM64/ASan are explicitly user-deferred.
