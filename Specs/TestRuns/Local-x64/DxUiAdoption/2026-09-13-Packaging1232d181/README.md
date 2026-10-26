# Corrected final portable package and rollback

Local A/B/A rehearsal: baseline master `0dd0bd6d`, corrected candidate `1232d181`, then
baseline again. Every stage passes clean-extraction application startup, plugin contract
and runtime dependency closure checks. All 11 candidate DxUi sidecars match the packaged
module bytes and pin `13788e95`. Baseline file names and payload hashes match before/after;
ZIP container hashes may differ. The candidate Release root build passes with
zero warnings/errors.

| Measure | Baseline | Candidate | Change |
|---|---:|---:|---:|
| ZIP bytes | 30,289,465 | 30,705,000 | +415,535 (1.37%) |
| Expanded file bytes | 78,647,558 | 79,787,905 | +1,140,347 (1.45%) |
| Main EXE bytes | 28,911,616 | 29,200,384 | +288,768 (1.00%) |
| Files | 171 | 182 | +11 (6.43%) |

Candidate ZIP SHA-256: `BBFA4B7EC8C4F1082CC5DDC9D1B0B0DBECA1C1387766D66141EF76DC3982CA9B`.
The eleven extra files are module provenance sidecars. This measures the test-enabled
static-library migration, including the corrected test diagnostics and pending-content fixture. It is not
a DLL-versus-static comparison, ordinary shipping-build measurement, or runtime-memory claim.

The six original receipts are hashed in `retained-files.json`. ZIPs stay under
`Z:/RedSalamander.Perf/evidence/I19-20260909/1232d181-portable-rollback`, outside Git.
Each stage uses process-local TEMP/LOCALAPPDATA isolation. This does not qualify installed
updates, persisted-user-settings migration, minimum-OS behavior or native ARM64 execution.
No hosted CI ran. The I19 closeout record binds the separate Fresh Full results; further
ARM64/ASan qualification is explicitly user-deferred, and the measured resource costs are accepted.
