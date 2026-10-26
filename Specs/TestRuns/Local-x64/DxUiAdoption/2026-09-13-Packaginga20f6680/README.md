# Corrected final portable package and rollback

Local A/B/A rehearsal: baseline master `0dd0bd6d`, corrected candidate `a20f6680`, then
baseline again. Every stage passes clean-extraction application startup, plugin contract
and runtime dependency closure checks. All 11 candidate DxUi sidecars match the packaged
module bytes and pin `13788e95`. Baseline file names and payload hashes match before/after;
ZIP container hashes may differ. The candidate Release root build passes in 4m48s with
zero warnings/errors.

| Measure | Baseline | Candidate | Change |
|---|---:|---:|---:|
| ZIP bytes | 30,289,465 | 30,701,720 | +412,255 (1.36%) |
| Expanded file bytes | 78,647,558 | 79,772,033 | +1,124,475 (1.43%) |
| Main EXE bytes | 28,911,616 | 29,184,512 | +272,896 (0.94%) |
| Files | 171 | 182 | +11 (6.43%) |

Candidate ZIP SHA-256: `CF17E924C6477A58B70FE3E8E88730C7EC3963172757772CFC0FE53ED8A45C2F`.
The eleven extra files are module provenance sidecars. This measures the test-enabled
static-library migration, including the corrected ViewerText test diagnostics. It is not
a DLL-versus-static comparison, ordinary shipping-build measurement, or runtime-memory claim.

The six original receipts are hashed in `retained-files.json`. ZIPs stay under
`Z:/RedSalamander.Perf/evidence/I19-20260909/a20f6680-portable-rollback`, outside Git.
Each stage uses process-local TEMP/LOCALAPPDATA isolation. This does not qualify installed
updates, persisted-user-settings migration, minimum-OS behavior or native ARM64 execution.
No hosted CI ran. Final Full and the remaining acceptance gates stay in the I19 checklist.
