# Native build policy, console results and dependency paths

Status: normative intended contract
Last reviewed: 2026-11-20

## Compiler and linker policy

First-party C++ projects use MSVC 14.52.x/v145, Unicode, stdcpplatest, SDL and
conformance mode. `Directory.Build.props` also enables the standard MSVC preprocessor with
`/Zc:preprocessor` and defines `WIN32_LEAN_AND_MEAN` and `NOMINMAX` with empty replacement lists.
Source and SDK consumers use the same first-party definitions. WarningLevel is EnableAllWarnings (/Wall) and
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

`Build/BuildConsole.psm1` binds the result, diagnostic, and link implementation to the exact clean CI engine checkout.
The diagnostic gate also refuses structured MSVC tool exception dumps.
A native success summary requires native exit success, the final MSBuild footer
with explicit zero warning and error counts, and the entrypoint's existing
provenance and attestation checks. Direct console mode asks the file logger for
the footer. Streamed execution asks the console logger for it.
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

The vcpkg installer mirror check uses the same source-path helper and verifies
the complete restored Git HEAD before importing DxUi's tool modules. A restored
16-character source locator must execute that check rather than silently skip it.

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

## Build-console fixture lifetime

The internal BuildConsoleCheck preserves its console/path assertions. Without explicit
ScratchRoot, borrow only an existing canonical Runner-format context under the selected
marked root. Validate the contained run directory before any context/scratch factory;
require its live owner (including a runner parent), rejecting a directory older than
that process's start under the canonical PID-reuse rule. Missing, fallback, unauthorized
or unavailable ownership fails before creation. Never archive a borrowed run.

Without an inherited id, own one canonical current-PID run and archive it on success or
failure. Report success only after Archived; retain primary and archive errors together,
including fatal descendants after the archive attempt. Explicit ScratchRoot/shared fake-root
fixtures keep caller ownership. Owning Pester preserves the original end-to-end assertions
and tests the lifecycle boundary plus a representative current canonical module-copy bp28
fixture. The copied-module fixture starts a contained no-window child using the same
PowerShell executable as the runner, through `TestSupport.psm1`, inheriting the marked run
and test temporary root. It verifies the actual helper's zero exit and success output, then
confirms none of the fixture's copied modules entered the Pester caller; this keeps later
name-based `InModuleScope` lookups unambiguous. That is not a historical copy-list claim;
TestSandbox already depends on Common.
