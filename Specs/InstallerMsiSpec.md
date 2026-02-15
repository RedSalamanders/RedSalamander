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
  - non-shipping executables (PoC/test `*.exe` other than `RedSalamander.exe` and `RedSalamanderMonitor.exe`)
  - `asan.supp`
- Create Start Menu shortcuts for:
  - Red Salamander
  - Red Salamander Monitor

## Non-Goals

- No MSIX details in this document (see `Specs/InstallerMsixSpec.md`).
- No auto-download of WiX in `build.ps1` (WiX must be installed on the build machine / CI).

## Tooling

- WiX Toolset CLI v6+ (provides `wix.exe`).
- WiX UI extension (`WixToolset.UI.wixext`) is installed on-demand by `Installer/msi/build-msi.ps1`.

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

MSI output is written to:
- `.build\AppPackages\`

## Symbols MSI (PDB)

In addition to the main installer, we produce a **Symbols MSI** that installs Release PDBs into the same install folder:
- WiX source: `Installer/msi/ProductSymbols.wxs`
- Build script: `Installer/msi/build-msi-symbols.ps1`

The symbols MSI includes:
- `RedSalamander.pdb`
- `RedSalamanderMonitor.pdb`
- `Common.pdb`
- `Plugins\*.pdb`

## Versioning

MSI uses a 3-part version (`major.minor.build`) derived from `Common/Version.h`:
- `major` = `VERSINFO_MAJOR`
- `minor` = `VERSINFO_MINORA`
- `build` = `VERSINFO_BUILDNUMBER`

Example:
- `7.0.183`

## Upgrades

The MSI is configured as a major-upgrade.
Older versions are removed early in the install sequence to avoid component GUID stability issues when the installed file set changes.
