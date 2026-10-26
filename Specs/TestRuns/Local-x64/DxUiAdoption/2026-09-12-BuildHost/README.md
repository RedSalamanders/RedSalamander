# ARM64 cross-compiler host default

On this x64 Windows host, the unchanged consumer imports selected the x86 host for
all three ARM64 configurations. ARM64 Release failed with C1002 during whole-program
optimization; a later link also reported x86 address-space exhaustion. The conditional
root props default selects x64 while preserving explicit overrides and native ARM64
host selection.

`evaluation.json` retains real project evaluations before/after, the changed property
file hash and the 36 passing dependency/inventory/governance checks. Reproduce with
`Tools/Tests/DxUiDependency.Tests.ps1` under Pester 3.4. It evaluates all six real
consumer profiles and explicit overrides without building product artifacts.

The x64-host retry progressed past the heap failure but hit C1090 PDB-server errors
while another worktree was building. Sequential Release and ASan Debug retries
subsequently passed at `751437d2`; all six module identities are retained in the
[local qualification record](../2026-09-13-LocalQualification/README.md). These
property checks alone do not establish build success or native ARM64 runtime qualification.
Full local diagnostics remain in `Z:\RedSalamander.Perf\evidence\I19-20260909`, under
`751437d2-ARM64-Release*`, `arm64-host-default-tests.log`, and
`arm64-host-before-after.json`. No hosted CI was used.
