# Final Release portable package and rollback rehearsal

RedSalamander `3a43186a3507ca48eb41e7ad8dec8585aed732f9`, pinned DxUi `13788e95`,
passes its local x64 Release root build in 4m48s with zero warnings/errors. The official
`Installer/zip/build-zip.ps1 -Configuration Release -Platform x64 -BuildNumber 100`
then runs serially for baseline `0dd0bd6d` (A1), candidate (B), and baseline again (A2).
Each stage passes clean-extraction application `--help`, plugin-contract and runtime-closure
checks. Separate process-local TEMP/LOCALAPPDATA directories isolate the rehearsal.
No installed application or normal user settings are changed.

| Test-enabled portable payload | Baseline | Candidate | Change |
| --- | ---: | ---: | ---: |
| ZIP bytes | 30,289,465 | 30,694,026 | +404,561 (+1.34%) |
| Expanded file bytes | 78,647,558 | 79,758,721 | +1,111,163 (+1.41%) |
| Files | 171 | 182 | +11 DxUi provenance sidecars |
| Main EXE bytes | 28,911,616 | 29,184,000 | +272,384 (+0.94%) |

All 11 candidate sidecars identify the exact pin, x64 and Release; independent SHA-256
checks match their module bytes inside the ZIP. A1 and A2 contain identical file names
and payload hashes. ZIP container hashes differ while all file payload hashes match.
The previously qualified baseline ZIP is also preserved before the rehearsal.

Candidate ZIP SHA-256:
`CB87D5FB294017FEA1EDBBF7239DA2881A7B0F7CD83F725B8E7244C001D5D11E`.
All three ZIPs and the previous baseline remain in
`Z:/RedSalamander.Perf/evidence/I19-20260909/3a43186a-portable-rollback`;
executables are not committed in this evidence folder. `retained-files.json` hashes
six original JSON/build/package receipts.

This measures the migration of a test-enabled static-library product. It is not an
isolated DLL-versus-static comparison or a shipping build without test support.
The rehearsal establishes portable extraction/rollback smoke only. Installed-product
update, persisted-settings migration, minimum-OS and native ARM64 execution are not
claimed. Final Full profile and resource acceptance remain in the I19 checklist.
