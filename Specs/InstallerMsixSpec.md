# MSIX Installer Specification

## Overview

This specification defines the MSIX packaging flow for RedSalamander. The installer is implemented as a Windows Application Packaging Project that bundles the Release outputs of the solution and produces an MSIX package suitable for per-user or per-machine deployment.

Primary files:
- `Installer/msix/RedSalamanderInstaller.wapproj`
- `Installer/msix/Package.appxmanifest`
- `Installer/msix/Assets/*`
- `Installer/msix/GenerateAssets.ps1` (generates `Installer/msix/Assets/*` from `RedSalamander/res/logo.png`)

## Goals

- Produce an MSIX installer for **Release x64** builds.
- Include **RedSalamander.exe**, **RedSalamanderMonitor.exe**, plugins, and runtime dependencies.
- Keep packaging **deterministic** and driven by the `x64\Release` output.
- Enable **optional signing** in CI with secrets.
- Support **per-user** install (default) and **per-machine** provisioning.

## Non-Goals

- No MSI details in this document (see `Specs/InstallerMsiSpec.md`).
- No automatic certificate issuance.
- No Store submission automation.

## Package Contents

The packaging project includes files from:
- `x64\Release\**\*` (drives plugins + runtime dependencies)

Excluded from the package:
- Debug and link artifacts (`*.pdb`, `*.lib`, `*.exp`, `*.ilk`, `*.iobj`, `*.ipdb`).
- Non-shipping executables (PoC/test `*.exe` other than `RedSalamander.exe` and `RedSalamanderMonitor.exe`).
- `asan.supp`

The package includes:
- `RedSalamander.exe`
- `RedSalamanderMonitor.exe`
- `Plugins\*.dll` and their copied runtime dependencies
- `Themes\*.theme.json5`
- `SettingsStore.schema.json`

## Manifest

`Installer/msix/Package.appxmanifest` declares two full-trust desktop apps:
- `RedSalamander` (main UI)
- `RedSalamanderMonitor`

Capabilities:
- `runFullTrust`

Versioning:
- The `Identity` version must match `Common/Version.h` (currently `7.0.0.183`).
- Update both when cutting a release.

## Build

### Local build

1. Build the solution and the MSIX:
   - `build.ps1 -Msix`

Or build/package separately:
- `build.ps1 -Configuration Release`
- `msbuild Installer\msix\RedSalamanderInstaller.wapproj /p:Configuration=Release /p:Platform=x64`

The MSIX output is written to:
- `AppPackages\`

### CI build

The GitHub workflow in `.github/workflows/release.yml`:
- Installs vcpkg dependencies
- Builds the solution in Release
- Builds the MSIX package
- Uploads the MSIX and a Release zip

## Signing

MSIX packages must be signed to install on most machines.

CI signing is **optional** and controlled by secrets:
- `MSIX_SIGNING_CERT`: base64-encoded PFX
- `MSIX_SIGNING_PASSWORD`: PFX password

When the secrets are present, the workflow signs every MSIX artifact via `signtool.exe`.

Note: the signing certificate subject must match `Installer/msix/Package.appxmanifest` → `Identity Publisher`.

## Installation Modes

- **Per-user install**: use `Add-AppxPackage` (or double-click if trusted) on the signed MSIX.
- **Per-machine install**: provision the package for all users with `Add-AppxProvisionedPackage` (admin required).

Notes:
- MSIX installs into `C:\Program Files\WindowsApps` (managed by Windows). This is expected for both per-user and per-machine installs.
- For enterprise deployment, use Intune or other provisioning systems.
