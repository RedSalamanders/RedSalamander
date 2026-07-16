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
```

## Parameters

| Parameter | Values | Default |
|-----------|--------|---------|
| `-Configuration` | Debug, Release, ASan Debug | Debug |
| `-Platform` | x64, ARM64 | x64 |
| `-ProjectName` | RedSalamander, RedSalamanderMonitor, Common, FileSystem | All projects |
| `-Clean` | Switch | False |
| `-Rebuild` | Switch | False |

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
- Keep vcpkg.json files up to date
- Pin versions for critical dependencies

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
- Plain consoles that support child-console output keep MSBuild's native console output/color while still capturing a `.build\logs\msbuild-*.log` file; Windows Terminal and redirected hosts such as Codex use the replay helper so build progress stays visible, with replayed errors/warnings/project completion colorized in the console
- MSBuild launches always use the sanitized environment helper so duplicate `Path`/`PATH` process-environment aliases are collapsed to canonical `Path` before MSBuild starts compiler tool tasks
- `build.ps1` and every `Run-AllTests.ps1` lane, including `-SkipBuild`, hold the same repository-scoped exclusive lock file. `.build\artifact-operation-owner.json` identifies the root PID, operation, and UTC start; only a direct child synchronously launched in a kill-on-close Job Object may reuse that ownership. Do not start unrelated operations in parallel or edit build inputs while either is running.
- Interrupted launch trees are contained by a kill-on-close Job Object. If abandoned ownership marks `.build\artifact-operation-contaminated.json`, run one serialized `build.ps1 -Rebuild` before trusting incremental outputs or starting tests.
- Shows build time and output file sizes
- Supports multi-processor builds (`/m`)
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
