# Native build policy, console results and dependency paths

Status: normative intended contract
Last reviewed: 2026-10-28

## Compiler and linker policy

First-party C++ projects use MSVC 14.52.x/v145, Unicode, stdcpplatest, SDL and
conformance mode. WarningLevel is EnableAllWarnings (/Wall) and
TreatWarningAsError is false (/WX disabled), in every supported profile.
The build must finish with **zero warnings and zero errors** before a successful
receipt or green result is published. Disabling /WX allows collection of all
diagnostics; it does not permit accepting warnings.

Fix actionable warnings at their cause. Vendor/system headers may be classified
as external with external warning level zero; first-party code is not external.
NativeBuildPolicy.targets documents the reviewed optimizer-inline, padding and
optional Spectre notices. Any additional suppression requires a narrow scope,
rationale, owner and review date. Do not hide a new diagnostic to pass validation.

Ordinary Debug uses EditAndContinue (/ZI); ASan Debug uses ProgramDatabase (/Zi).
Release retains ProgramDatabase. The existing explicitly requested and attested
RedSalamander Embedded compiler-debug recovery mode remains a separate exception.
Release enables /Gy /GL and final /OPT:REF /LTCG /OPT:ICF, with incremental linking
disabled. Debug and ASan Debug explicitly disable COMDAT folding (/OPT:NOICF).
Static archives remain complete; final executable/DLL links perform elimination
and folding. Cross-builds do not establish native ARM64 runtime qualification.

Build/NativeBuildPolicy.targets owns common warning/debug/ICF metadata.
Build/PinnedDxUiPolicy.targets binds that policy to independently pinned DxUi
archives and direct helper builds. Consumers do not modify the SDK source clone.
Full native policy version and normalized target-file digests participate in
dependency fingerprints; an old archive cannot reuse the changed policy's output.

## Successful terminal output

Build/BuildConsole.psm1 implements the common result/diagnostic/link contract.
A native success summary is green only after native exit success, zero logged
warnings/errors and the entrypoint's existing provenance/attestation checks.
Failed builds remain failures. Clean/package-only operations preserve their
existing meaning and do not pretend to be native compilation.

Existing artifact and log paths become local file hyperlinks in terminals with
OSC 8 support. File URIs use localhost, proper URI escaping, and reject terminal
control characters. Plain/redirected output retains literal paths and a quoted
PowerShell launch command. NO_COLOR is honored; tests can explicitly select
Always/Never link and color modes on the formatter.

Links never launch automatically. Ctrl+click handling belongs to the terminal;
Windows Terminal supports OSC 8 and local file links. A host that cannot open
executable links can use the displayed launch command. ARM64 outputs on an x64
host are labelled for ARM64 execution rather than offered as a native Run action.
See https://github.com/microsoft/terminal/issues/5001.

## DxUi dependency directory identity

Consumer source clones are located at
.build/dependencies/DxUi/source/<first-16-commit-hex>/.
Compiled dependency roots use the first 16 hexadecimal characters of the full
SHA256 build fingerprint. These folder names are locators, not identity proofs.
The lock, Git HEAD, identity JSON, marker and provenance retain complete digests.

The normal consumer validator checks the complete source pin before reuse.
A verified output marker records the complete fingerprint and complete identity;
equal prefixes with different identities fail without resetting or deleting data.
An existing nonempty output without its verified marker is refused. Provenance
validates the actual resolved output root and the complete fingerprint.

Short-path changes leave old full-hash source/output directories in place.
Do not create compatibility junctions or clean historical output as a side effect
of this migration. Cleanup follows its separate inventory/preservation/process-use
policy. Source restore, MSBuild imports, helper invocations, provenance and tests
must agree on the same short folder convention.

## Reviewed ARM64 link-time notice

MSVC 14.52 ARM64 Release LTCG re-emits C4746 (the volatile-model notice) from
inlined code after local diagnostic scopes, including Windows SDK WRL.
The policy suppresses that notice only for ARM64 Release. The target's volatile
model and code behavior remain unchanged. Reviewed owned uses are stack rounding
barriers, ICF/benchmark observations, deliberate crash/ASan probes and page prefault
touches; shared crash ownership uses interlocked operations. New volatile uses
require their own semantic review and a local explanation. This exception does
not cover actionable conversion, lifetime, enum or concurrency diagnostics.
Owner: repository build maintainers. Review by 2027-01-31.
The evaluated-profile check rejects the exception in other profiles.
See https://learn.microsoft.com/en-us/cpp/error-messages/compiler-warnings/compiler-warning-c4746.
The green footer duration measures the public entrypoint through result publication, including restore, admission and attestation. Native compilation phase timings and workload performance receipts retain their separate meanings.
