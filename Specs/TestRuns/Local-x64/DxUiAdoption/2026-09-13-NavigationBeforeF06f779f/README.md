# Navigation failure diagnosis: f06f779f-navigation-before

Local x64 Debug at `f06f779fd45105d251c2761a82138e7c8215326e`; no hosted CI.
Each exact comma-separated case sequence in `summaries.json` runs ten times in one
Commands process. Original native results retain every case, repeat and failure reason.

| Sequence | Passed | Failed | Skipped | Run |
|---|---:|---:|---:|---|
| refresh-sequence | 40 | 0 | 0 | 20260913T100953Z-39304-082c576a1fd641a3a3e3fcb7f14609e3 |
| drive-menu-sequence | 40 | 0 | 0 | 20260913T101100Z-39304-10c768ee9ae049f5a51cec010a3b62ad |

These are focused diagnostic runs, not Fresh Full qualification. A passing reduced
sequence does not explain the earlier Full failure or prove that broader process state
is clean. Expanded context and short sequences must not be presented as matched performance
workloads. `case-comparison.json` extracts counts/failures without modifying original receipts.

`retained-files.json` hashes 7 original receipts. Complete native archives stay
under `Z:/RedSalamander.Perf/evidence/I19-20260909/f06f779f-navigation-before`. The
[I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md)
owns the remaining x64 qualification; ARM64/ASan are explicitly user-deferred.
