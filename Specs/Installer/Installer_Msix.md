# MSIX Installer Specification

## Overview

This specification defines the MSIX packaging flow for RedSalamander. The installer is implemented as a Windows Application Packaging Project that bundles the Release outputs of the solution and produces an MSIX package suitable for per-user or per-machine deployment.

Primary files:
- `Installer/msix/RedSalamanderInstaller.wapproj`
- `Installer/msix/Package.appxmanifest`
- `Installer/msix/Assets/*`
- `Installer/msix/GenerateAssets.ps1` (generates `Installer/msix/Assets/*` from `RedSalamander/res/logo.png`)
- `Installer/msix/UpdateManifestVersion.ps1` (stamps the generated package identity version and architecture)

## Goals

- Produce MSIX installers for requested **Release x64 and ARM64** builds.
- Include **RedSalamander.exe**, **RedSalamanderMonitor.exe**, plugins, and runtime dependencies.
- Keep packaging **deterministic** and driven by the `.build\x64\Release` output.
- Enable **optional signing** in CI with secrets.
- Support **per-user** install (default) and **per-machine** provisioning.

## Non-Goals

- No MSI details in this document (see `Specs/Installer/Installer_Msi.md`).
- No automatic certificate issuance.
- No Store submission automation.

## Package Contents

The packaging project includes files from:
- `.build\x64\Release\**\*` (drives plugins + runtime dependencies)

Excluded from the package:
- Debug and link artifacts (`*.pdb`, `*.lib`, `*.exp`, `*.ilk`, `*.iobj`, `*.ipdb`).
- Non-shipping executables (PoC/test `*.exe` other than `RedSalamander.exe` and `RedSalamanderMonitor.exe`).
- `asan.supp`

The package includes:
- `RedSalamander.exe`
- `RedSalamanderMonitor.exe`
- `Plugins\*.dll` and their copied runtime dependencies
- `Lang\*.dll` satellite resource assemblies
- `Themes\*.theme.json5`
- `SettingsStore.schema.json`

Directory-preserving harvest is required. Recursive content from `.build\<platform>\<configuration>` must not flatten subdirectories into the MSIX package root. `Installer/msix/RedSalamanderInstaller.wapproj` keeps the broad runtime dependency harvest for root-level files, but explicitly harvests `Plugins\`, `Lang\`, and `Themes\` with matching `PackagePath`, `Link`, and `TargetPath` metadata so plugin DLLs, language satellites, and themes keep their runtime-relative paths.

Validation:
- `Tools/Tests/WingetValidation.Tests.ps1` asserts the MSIX project preserves `Plugins\`, `Lang\`, and `Themes\` package paths and that the portable ZIP copies `Lang\`.
- Release package validation must list the built MSIX (or unpack it with `makeappx`) and confirm plugin DLLs and satellite resource DLLs appear under their expected subdirectories, not at the package root.

## Manifest

`Installer/msix/Package.appxmanifest` declares two full-trust desktop apps:
- `RedSalamander` (main UI)
- `RedSalamanderMonitor`

Capabilities:
- `runFullTrust`

Versioning:
- The source manifest keeps a neutral `7.0.0.0` identity version.
- Before packaging, `Installer/msix/UpdateManifestVersion.ps1` stamps the identity as `<major>.<minor>.<build>.0`, for example `7.0.183.0`.
- The same script stamps `ProcessorArchitecture` from the target platform.
- The build number comes from `Tools/Versioning.ps1`, using `GITHUB_RUN_NUMBER` in CI.

## Build

### Local build

1. Build the solution and the MSIX:
   - `build.ps1 -Msix`

Or build/package separately:
- `build.ps1 -Configuration Release`
- `msbuild Installer\msix\RedSalamanderInstaller.wapproj /p:Configuration=Release /p:Platform=x64`

The MSIX output is written to:
- `.build\AppPackages\`

### CI build

The GitHub workflow in `.github/workflows/release.yml`:
- Always builds the requested Release portable matrix: x64, plus ARM64 unless the explicit
  `build_arm64` input disables it.
- Builds MSIX only when the explicit `build_msix` input enables it. A disabled MSIX leg succeeds as a policy
  no-op; it is not treated as a missing artifact.
- Normalizes each enabled package to
  `RedSalamander-<major.minor.build>-<x64|ARM64>.msix` before upload.
- Fails before GitHub Release creation if any requested package job or download fails, if the exact expected
  ZIP/MSIX set is not present, or if any package is empty, duplicated, unexpectedly named, or for the wrong
  architecture.
- Validates MSIX identity `Name`, `Publisher`, four-part `Version`, and `ProcessorArchitecture` through
  `Tools/ReleaseArtifactPolicy.ps1`, then generates and revalidates the exact `checksums.sha256` entries.

The package identity contract is `Name="RedSalamander"`, `Publisher="CN=RedSalmanders"`, version
`<major>.<minor>.<build>.0`, and architecture `x64` or `arm64` matching the artifact filename. The workflow uses
ordinary successful-dependency semantics: it must never use `always()`, `!cancelled()`, ignored artifact-download
errors, or wildcard acceptance to publish a partial requested matrix.

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
