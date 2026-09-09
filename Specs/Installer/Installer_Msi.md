# MSI Installer Specification

## Overview

This specification defines the MSI packaging flow for RedSalamander using WiX Toolset (v6+).
The MSI is intended to simplify installation on machines where MSIX is undesirable.

Primary files:
- `Installer/msi/Product.wxs`
- `Installer/msi/build-msi.ps1`

## Goals

- Produce an MSI installer for **Release x64** builds.
- Install to **`Program Files`** (per-machine install, requires admin).
- Include the **shipping runtime** from `.build\x64\Release\` (apps + plugins + runtime dependencies), excluding:
  - build artifacts (`*.pdb`, `*.lib`, `*.exp`, `*.ilk`, `*.iobj`, `*.ipdb`, `*.pch`)
  - non-shipping executables (PoC/test `*.exe`; the shipping executables are `RedSalamander.exe`, `RedSalamanderMonitor.exe`, and `RedSalamanderSearchService.exe`)
  - `asan.supp`
- Create Start Menu shortcuts for:
  - RedSalamander
  - RedSalamander Monitor
- Install `RedSalamanderSearchService.exe` as the optional `RedSalamanderSearchService` Windows service. The service component is selected by default, runs as LocalSystem, starts automatically, and can be opted out with `INSTALLSEARCHSERVICE=0`.

## Non-Goals

- No MSIX details in this document (see `Specs/Installer/Installer_Msix.md`).
- No auto-download of WiX in `build.ps1` (WiX must be installed on the build machine / CI).

## Tooling

- WiX Toolset CLI v6+ (provides `wix.exe`).
- WiX UI extension (`WixToolset.UI.wixext`) must be installed explicitly at a
  version matching the WiX CLI before packaging. MSI scripts do not run the
  machine-global `wix extension add -g` mutation.

Typical install:
- `winget install --exact --id WiXToolset.WiXCLI`

## Build

### Local build

- Build the solution and the MSI packages:
  - `build.ps1 -Msi`

Or build/package separately:
- `build.ps1 -Configuration Release`
- `Installer\msi\build-msi.ps1 -Configuration Release -Platform x64`
- `Installer\msi\build-msi-symbols.ps1 -Configuration Release -Platform x64`

Both standalone MSI entrypoints first hold the `Release|x64` artifact-profile
lease while reading compiled inputs, then enter the repository-wide packaging
coordination scope before publishing. `robocopy.exe` and WiX run as contained
children so interruption cannot leave an unowned packager continuing after the
leases close. Same-thread calls from `build.ps1 -Msi` reenter both locks with
balanced child leases.

An abandoned packaging owner records repository-wide contamination and the MSI
scripts fail closed with the marker path. Packaging does not clear this marker:
a reviewer must restore shared `Installer` inputs and `.build\AppPackages`, then
remove the marker explicitly.

`OutputDirectory` is normalized and accepted only at or beneath this
repository's `.build\AppPackages`; a repository-local lock must not imply
coordination for an arbitrary external path or another worktree.

MSI output is written to:
- `.build\AppPackages\`

## Symbols MSI (PDB)

In addition to the main installer, we produce a **Symbols MSI** that installs Release PDBs into the same install folder:
- WiX source: `Installer/msi/ProductSymbols.wxs`
- Build script: `Installer/msi/build-msi-symbols.ps1`

The symbols MSI includes:
- `RedSalamander.pdb`
- `RedSalamanderMonitor.pdb`
- `RedSalamanderSearchService.pdb`
- `Common.pdb`
- `Plugins\*.pdb`

## Versioning

MSI uses a 3-part version (`major.minor.build`) derived from `Common/Version.h`:
- `major` = `VERSINFO_MAJOR`
- `minor` = `VERSINFO_MINOR`
- `build` = the shared build number resolved by `Tools/Modules/Build/Versioning.psm1`

Both standalone MSI entrypoints require `-BuildNumber 1..65535`. They use the
pure `New-RSVersionContext` operation for their requested Release/x64 profile and
never read saved repository state or allocate a counter during packaging.
`build.ps1 -Msi` passes the build number that stamped the compiled profile.

Example:
- `7.0.183`

## Upgrades

The MSI is configured as a major-upgrade.
Older versions are removed early in the install sequence to avoid component GUID stability issues when the installed file set changes.
