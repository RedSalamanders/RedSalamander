---
name: cpp-build
description: Build RedSalamander C++ project using PowerShell build script. Use when building, compiling, or rebuilding the solution or specific projects like RedSalamander, RedSalamanderMonitor, Common, or FileSystem.
metadata:
  author: RedSalamander
  version: "1.0"
---

# Building RedSalamander

## Quick Build Commands

```powershell
# Build entire solution (default)
.\build.ps1

# Build in Release configuration
.\build.ps1 -Configuration Release

# Build specific project
.\build.ps1 -ProjectName RedSalamander
.\build.ps1 -ProjectName RedSalamanderMonitor
.\build.ps1 -ProjectName Common

# Clean and rebuild
.\build.ps1 -Clean
.\build.ps1 -Rebuild

# Bound worker processes when the host cannot sustain MSBuild's default
.\build.ps1 -Rebuild -MaxCpuCount 4
```

## Parameters

| Parameter | Values | Default |
|-----------|--------|---------|
| `-Configuration` | Debug, Release, ASan Debug | Debug |
| `-Platform` | x64, ARM64 | x64 |
| `-ProjectName` | RedSalamander, RedSalamanderMonitor, Common, FileSystem | All projects |
| `-Clean` | Switch | False |
| `-Rebuild` | Switch | False |
| `-MaxCpuCount` | 0-256 | 0 (MSBuild default) |

## Output Locations

- Debug: `.build\x64\Debug\*.exe, *.dll`
- Release: `.build\x64\Release\*.exe, *.dll`
- Debug (ARM64): `.build\ARM64\Debug\*.exe, *.dll`
- Release (ARM64): `.build\ARM64\Release\*.exe, *.dll`

## Build Order (Dependencies)

1. **Common** - Shared library (no dependencies)
2. **RedSalamanderMonitor** - Monitoring app (depends: Common)
3. **RedSalamander** - Main app (depends: Common, FileSystem, RedSalamanderMonitor)
4. **FileSystem** - Plugin (no dependencies)

## Visual Studio Build

1. Open `RedSalamander.sln` in Visual Studio 2022+
2. Select configuration (Debug/Release) and platform (x64)
3. Build → Build Solution (Ctrl+Shift+B)

## vcpkg Integration

- Uses vcpkg for package management
- Dependencies defined in `vcpkg.json`
- The executable checkout is independently pinned by `vcpkg-tool.json`; local installs and CI must validate the
  exact Git commit
- Do not use global `vcpkg integrate install`; the repository supplies manifest/toolchain integration explicitly
- Keep both pin files intentional and reviewed

## Runtime Dependency Staging

- Root `RuntimeDependencies.props` is the canonical list of app-local plugin runtime DLLs
- `Directory.Build.targets` owns required-file failure, copying, and stale-output removal
- Do not add plugin-local `PostBuildEvent` or `xcopy` dependency batches; extend the canonical manifest
- `Installer/zip/build-zip.ps1` consumes the same manifest and runs a fresh-extraction app/plugin smoke
- See `Specs/Installer/Installer_PortableZip.md`

## Tooling Changes

Read `Specs/Testing/Testing_ToolingGovernance.md`, `Tools/README.md`, and
`.github/skills/tooling-governance/SKILL.md` before changing a build script or
helper. Keep stable build entrypoints at the Tools root, place reusable
implementation in explicit modules, update `Tools/tool-inventory.json`, and
prove that workflow cache keys include every imported module, patch, notice,
schema, and version input that can change produced artifacts.

### Operation Startrail build evidence

Build evidence Phase 1 is active. `build.ps1` invalidates the mutable receipt pointer
before artifact mutation, forces `Rebuild` for first/incompatible attestation, checks the
source snapshot before final publication, and writes content-addressed targeted/full
receipts only after all selected build/package operations succeed. Do not add an output
or post-build staging step without adding it to the attested closure or re-attesting it
before the final banner.

`Run-AllTests.ps1 -SkipBuild` is receipt-gated: it requires a test-enabled full-solution
receipt for the exact checkout/profile and verifies every artifact digest. CI/nightly
cross-job handoff must publish and verify the portable same-workflow attestation using
the reusable workflow's attestation/toolchain outputs, then rebind paths locally. Never
accept a local receipt's absolute diagnostic paths as portable authority. Fresh Full
remains the default and mandatory CI/nightly/release/final-closeout mode. Explicit local
exact Resume and Affected do not weaken build receipt verification.

An artifact-writer re-attestation is a build, not only a digest refresh. Apply the same
suite build environment used for the initial build and restore the caller environment
afterward. Full specifically requires `RSBuildEnableTests=true` even when invoked with
`-SkipBuild`; otherwise a valid receipt may bind a Monitor binary that cannot run the
declared Full selftest entry.

Build source identity now uses the canonical Phase 3 workspace snapshot: committed base,
index, and worktree bytes are distinct, dirty/untracked/delete/rename state is bound, and
mtime-only or generated evidence changes do not invalidate it. A validation entry
fingerprint consumes the verified receipt ID, declared outputs/runtime closure, and
current artifact epoch; it never substitutes for receipt compatibility. Compute the
complete receipt/runtime-closure binding digest once per run and bind that immutable
digest into every entry fingerprint; do not recanonicalize the full receipt per entry.

Validation examples:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -BuildNumber 1
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Resume -ResumeFrom <run-id> -SkipBuild
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Affected -ImpactBase origin/master
```

The executable build summary is printed exactly once after all selected build/package
work and receipt publication. Every `Run:` line must remain copy/paste executable.

## Shared MSBuild Defaults

- Shared first-party VC++ defaults live in `Directory.Build.props` and `Directory.Build.targets`
- Family-level overrides live in `Plugins/Directory.Build.props`, `Tests/Directory.Build.props`, and `PoC/Directory.Build.props`
- Prefer changing shared toolchain, output-path, versioning, and common compile defaults in those shared files instead of copying edits across individual `.vcxproj` files
- First-party projects default to `LanguageStandard=stdcpplatest` and `WarningLevel=EnableAllWarnings`
- First-party VC++ projects also default to `/FS`, so solution-targeted builds remain stable under the repo-wide `/MP` compiler settings
- Plugin DLLs and console test executables inherit shared external-header, include-path, and subsystem defaults from their family-level props files
- First-party `Application` and `DynamicLibrary` projects default to repo version stamping through `Directory.Build.targets` unless they explicitly opt out
- Proof-of-concept projects inherit `Level3` plus the shared `C4710/C4711` suppression from `PoC/Directory.Build.props`; keep any extra PoC warning deviations local only when truly project-specific

## Dependencies

- **WIL** - Windows Implementation Library (RAII wrappers)
- **fmt** - Modern C++ formatting library
- **DirectX** - Graphics and multimedia APIs (D2D, D3D11, DXGI)
- CI ARM64 builds install `Microsoft.VisualStudio.Component.VC.Tools.ARM64` with the Visual Studio 2026 VC++ workload and verify `Hostx64\arm64\cl.exe` before building.

## Build Script Features

- Automatically locates MSBuild (VS 2022 or later)
- Builds entire solution when no ProjectName specified
- Resolves most `-ProjectName` builds to the target `.vcxproj` directly; `RedSalamander` stays on the solution build graph so the solution can also build its bundled sibling outputs (`Plugins\*.dll`, `RedSalamanderMonitor.exe`, `RedSalamanderSearchService.exe`)
- Reuses the saved local beta build number by default, so ordinary repeated local builds stay incremental instead of forcing a fresh version stamp on every invocation
- Resolve-and-persist version state only through `Resolve-RSVersionContext`, which owns the exclusive repository-local version-state file lease and atomic counter/context publication. The file lease coordinates Windows sessions and checkout aliases and is released by process termination. Use pure `New-RSVersionContext` when a caller already has an explicit build number. `Read-RSVersionContext` is the locked, validated diagnostics-only reader; raw unlocked lock/read/save primitives are private, and the version lease never spans MSBuild or another native child.
- Plain consoles that support child-console output keep MSBuild's native console output/color while still capturing a `.build\logs\msbuild-*.log` file; Windows Terminal and redirected hosts such as Codex use the replay helper so build progress stays visible, with replayed errors/warnings/project completion colorized in the console
- MSBuild launches always use the sanitized environment helper so duplicate `Path`/`PATH` process-environment aliases are collapsed to canonical `Path` before MSBuild starts compiler tool tasks
- Artifact operations are exclusive by normalized repository root plus `Platform|Configuration`. Builds/tests for the same output profile serialize; disjoint profile artifact lifecycles and worktrees do not block one another merely because they are RedSalamander operations. `target` and `suite` are diagnostics, not lock partitions. Every production caller must pass both profile fields; the unscoped fallback is compatibility/test infrastructure only.
- `Run-AllTests.ps1` may execute only the canonical `RedSalamander.exe` beneath its selected repository/platform/configuration output. Its `-ExePath` parameter is compatibility syntax for explicitly naming that same normalized path, not permission to run a foreign or differently scoped artifact; mismatches fail before lock acquisition.
- MSBuild and direct vcpkg installation also hold the narrower repository+platform `shared-dependency` lock because Debug, Release, and ASan share `.build\vcpkg_installed\<triplet>`. It covers only dependency-capable phases, matches residual tools regardless of configuration, and never blocks the other platform or ordinary test execution. Multi-triplet installation acquires scopes in stable x64-then-ARM64 order. Explicit custom triplets must begin with `x64-` or `arm64-`; reject every architecture without a reviewed platform-lock mapping before mutating staging or install state. Apply the same mapping to path-qualified residual `vcpkg`, compiler, and linker evidence for custom installed/staging paths; a process name, foreign checkout, or opposite platform never suffices.
- The Ghostty runtime's `.build\TerminalEngine` state is shared across configurations and platforms within one repository, so its state mutex is derived only from normalized repository root; same-repository x64/ARM64 operations serialize, while different worktrees do not. A validated cache hit never takes the session-wide `R:` lease. Only an actual rebuild holds that lease, narrowly around mapping, contained overlay/build work, cleanup, and unmapping; a pre-existing mapping fails with inspection/removal guidance and is never cleared automatically.
- MSIX/MSI/ZIP/winget generation uses a repository-wide `packaging` scope only after compilation because platforms share installer inputs and `.build\AppPackages`. Standalone packagers that read compiled output acquire the artifact profile first and packaging second; `Installer/msix/build-msix.ps1` then takes its platform `shared-dependency` lease only around contained MSBuild and releases all three in reverse. Nested calls under `build.ps1` are balanced same-thread leases. Packaging abandonment fails closed for reviewed restoration rather than being cleared by a profile rebuild.
- Standalone MSI, symbols MSI, and ZIP calls require an explicit positive build number; Winget requires a complete version or positive build number plus both exact-version x64/ARM64 archive names; and MSIX requires a complete version. Never infer completed-artifact provenance from saved repository context or allocate a new number during packaging. Caller-selected package paths must remain beneath their guarded repository roots with every existing path component proven non-reparse before and after directory creation.
- Per-profile owner records authenticate PID, process start time, executable path, lock path, and token. Distinct same-thread nested scopes release LIFO. Only a same-thread nested call or an immediate child admitted by the one-use delegation handshake inside the owner's kill-on-close Job Object may reuse ownership; that child may reenter locally but cannot delegate a second generation. Never treat inherited environment variables alone as proof.
- Residual-tool checks require boundary-aware evidence for this repository and the same profile. A process name, an ambiguous substring, another configuration/platform, or another worktree is not enough to block or terminate anything.
- `Invoke-SanitizedMsbuild.ps1` requires explicit Configuration and Platform properties, including combined MSBuild property lists; it must never fall back to a repository-wide production lock.
- `build.ps1` never kills an independently launched application. If its executable path exactly equals an output the selected profile can replace, the build fails with PID/path/command-line diagnostics and asks the operator to close it. Automatic cleanup is limited to descendants launched and contained by the current operation.
- Interrupted launch trees are contained by a kill-on-close Job Object. An abandoned owner contaminates only its platform/configuration; run a full-solution `build.ps1 -Rebuild` for that same profile before trusting its incremental outputs or starting its tests.
- Shows build time and output file sizes
- Supports multi-processor builds (`/m`) and an explicit `-MaxCpuCount` stability bound
- Displays both executables when building full solution

## Required Validation Loop

For perf-sensitive work, building is only the first step.

After a successful build:

1. Run the deterministic selftest or scenario for the changed subsystem.
2. Confirm a new archived run appears under `Specs/TestRuns/`.
3. Compare the baseline and candidate runs before claiming an improvement.
4. If the work finishes a plan, move that plan to `Specs/Plans/Done/` and update the authoritative subsystem spec or repo guidance with the lasting requirement.

See:

- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/TestRuns/README.md`
- `.github/skills/perf-validation/SKILL.md`
