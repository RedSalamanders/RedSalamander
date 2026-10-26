# Focused adoption regressions: a8350f4b-history-context

Local x64 Debug at `a8350f4bf3e8d18cd666c720e2628167b4eb3af3`; no hosted CI.
Each declared case sequence in `summaries.json` runs ten times in its native self-test
suite. File Operations totals include setup/cleanup cases. Original native results
retain every case, repeat and failure reason.

| Sequence | Passed | Failed | Skipped | Run |
|---|---:|---:|---:|---|
| history-context | 130 | 0 | 0 | 20260913T114613Z-105516-05431069a94d41e38d91275fc58d355c |

These are focused diagnostic runs, not Fresh Full qualification. A passing reduced
sequence does not explain the earlier Full failure or prove that broader process state
is clean. Expanded context and short sequences must not be presented as matched performance
workloads. `case-comparison.json` extracts counts/failures without modifying original receipts.

`retained-files.json` hashes 6 original receipts. Complete native archives stay
under `Z:/RedSalamander.Perf/evidence/I19-20260909/a8350f4b-history-context`. The
[I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md)
owns the remaining x64 qualification; ARM64/ASan are explicitly user-deferred.
