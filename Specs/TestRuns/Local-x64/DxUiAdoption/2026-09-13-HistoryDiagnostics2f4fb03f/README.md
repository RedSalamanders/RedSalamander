# Focused adoption regressions: 2f4fb03f-history-context

Local x64 Debug at `2f4fb03fffb9484783ac0badb47f2f9b9f61c739`; no hosted CI.
Each declared case sequence in `summaries.json` runs ten times in its native self-test
suite. File Operations totals include setup/cleanup cases. Original native results
retain every case, repeat and failure reason.

| Sequence | Passed | Failed | Skipped | Run |
|---|---:|---:|---:|---|
| history-context | 130 | 0 | 0 | 20260913T115212Z-89904-b5e179268b834f45a4f6e9d937b11e21 |

These are focused diagnostic runs, not Fresh Full qualification. A passing reduced
sequence does not explain the earlier Full failure or prove that broader process state
is clean. Expanded context and short sequences must not be presented as matched performance
workloads. `case-comparison.json` extracts counts/failures without modifying original receipts.

`retained-files.json` hashes 7 original receipts. Complete native archives stay
under `Z:/RedSalamander.Perf/evidence/I19-20260909/2f4fb03f-history-context`. The
[I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md)
owns the remaining x64 qualification; ARM64/ASan are explicitly user-deferred.
