# Winget Integration Guide for RedSalamander

This document defines the supported Windows Package Manager publication flow for RedSalamander.

## Current Packaging Policy

RedSalamander publishes to Winget as a portable ZIP package.

Unsigned MSI packages are not part of the automated release or Winget workflow. Windows SmartScreen and installer trust prompts make unsigned MSI distribution a poor default, and the Winget path must not submit an MSI until a trusted code-signing pipeline exists.

MSIX remains optional and separate. MSIX packages also need signing before they are suitable for normal end-user installation.

## Release Artifact Contract

Official release versions use `major.minor.build`, where `major` and `minor` come from `Common/Version.h` and `build` comes from `GITHUB_RUN_NUMBER` in the release workflow.

The release workflow generates the release version and tag, so binary resources, ZIP names, GitHub release tag, and Winget manifest version stay aligned:

```text
v7.0.183
RedSalamander-7.0.183-x64-Portable.zip
PackageVersion: 7.0.183
```

The Winget workflow starts with the x64 and ARM64 portable ZIPs. Both assets are required for publication so Winget can select the native installer for the user's machine.

## Manifest Shape

RedSalamander uses a multi-file manifest under `Installer/winget/templates/`:

```text
RedSalamanders.RedSalamander.installer.yaml
RedSalamanders.RedSalamander.locale.en-US.yaml
RedSalamanders.RedSalamander.yaml
```

The installer manifest uses schema `1.12.0` and models the portable archive as:

```yaml
InstallerType: zip
NestedInstallerType: portable
NestedInstallerFiles:
  - RelativeFilePath: RedLauncher.exe
    PortableCommandAlias: RedSalamander
InstallationMetadata:
  Files:
    - RelativeFilePath: RedLauncher.exe
      FileType: launch
      DisplayName: RedSalamander
Installers:
  - Architecture: x64
    InstallerUrl: https://github.com/RedSalamanders/RedSalamander/releases/download/v{VERSION}/RedSalamander-{VERSION}-x64-Portable.zip
    InstallerSha256: {ZIP_SHA256}
  - Architecture: arm64
    InstallerUrl: https://github.com/RedSalamanders/RedSalamander/releases/download/v{VERSION}/RedSalamander-{VERSION}-ARM64-Portable.zip
    InstallerSha256: {ARM64_ZIP_SHA256}
```

Do not use `InstallerType: portable` for a ZIP archive. ZIP archives require `InstallerType: zip` plus `NestedInstallerType: portable`.

Declare `NestedInstallerFiles` at the manifest root because the executable path is identical in x64 and ARM64 ZIPs. The alias target is `RedLauncher.exe`, not `RedSalamander.exe`: Winget creates the command alias under its `Links` directory, and launching a symlinked executable directly can make the Windows loader search that alias directory before the package root. `RedLauncher.exe` is a dependency-free static-CRT Windows-subsystem shim that resolves its own final symlink target, finds the package root, and starts the real package-root `RedSalamander.exe` by absolute path so `Common.dll`, `yyjson.dll`, and other app-local DLLs are found without a console flash. Normal GUI launches, including diagnostic flags such as `--etw` and `--perf`, return control to the console immediately. The portable ZIP also includes `RedLauncherConsole.exe`, built from the same source as a console-subsystem companion, for foreground self-test automation that needs a wait and real app exit code. Keep `InstallationMetadata.Files` with `FileType: launch` pointing at `RedLauncher.exe` so Winget's repository validation has an explicit primary executable to locate after installation.

Use Winget's `Architecture` field for CPU selection: `x64` is the Intel/AMD 64-bit build, and `arm64` is the native Windows on ARM build. By default Winget chooses from the installers compatible with the current machine; users can override that choice with `winget install --architecture x64` or `winget install --architecture arm64`.

The portable ZIP bundles app-local Microsoft Visual C++ runtime DLLs from the latest installed Visual Studio Build Tools redistributable directory for the target architecture. Do not rely only on Winget `PackageDependencies` for the VC runtime: Winget repository validation launches the installed executable in a clean validation environment and may not install dependencies before that launch check.

Winget's terminal install experience does not render bitmap images from manifests. Console-visible branding is limited to text, so RedSalamander uses `InstallationNotes` for a short post-install message.

Winget has `Icons` metadata for PNG/JPEG/ICO package logos, but `winget validate` currently reports `Field usage requires verified publishers. [Icons]` for the public community manifest path and exits nonzero when RedSalamander adds that field. Do not add `Icons` to the submitted manifest until the publisher is accepted as verified by the Winget community repository. After verification, use `RedSalamander/res/logo64.png`, a version-pinned raw URL, `IconResolution: 64x64`, and a matching `IconSha256` generated from the checked-out release tag.

## Local Manifest Generation

Generate a manifest from an existing portable ZIP:

```powershell
.\Installer\winget\generate-manifest.ps1 `
  -Version 7.0.183 `
  -ZipPath .\RedSalamander-x64.zip `
  -Arm64ZipPath .\RedSalamander-ARM64.zip `
  -OutputDir .build\AppPackages\winget-manifest
```

Generate after a local release ZIP build:

```powershell
.\build.ps1 -Configuration Release -Zip -GenerateWingetManifest
```

Validate and smoke-test locally before submission:

```powershell
winget validate --manifest .build\AppPackages\winget-manifest
winget settings --enable LocalManifestFiles
winget install --manifest .build\AppPackages\winget-manifest
winget uninstall RedSalamanders.RedSalamander
```

`winget install --manifest` requires the `LocalManifestFiles` feature setting to be enabled before local manifest installation. Enabling that setting may require an administrator shell on developer machines.

`Installer\winget\WingetValidation.ps1` is the supported local helper for validating a generated manifest with the same legacy `winget.exe v1.11.x` schema-header warning policy used by the release workflow. The workflow intentionally keeps an inline copy of that policy because it checks out the release tag before manifest generation, and older release tags may not contain helper scripts added later.

## GitHub Actions Flow

`.github/workflows/winget-release.yml` runs on `release.published` and manual dispatch. Manual runs validate and install-test by default; set the `submit` input to `true` only when the run should open a `microsoft/winget-pkgs` pull request.

The workflow:

1. Resolves the release version from `Common/Version.h` plus `GITHUB_RUN_NUMBER`.
2. Checks out the matching release tag.
3. Reads the GitHub release asset list through the GitHub API.
4. Fails with the available asset names if either `RedSalamander-<version>-x64-Portable.zip` or `RedSalamander-<version>-ARM64-Portable.zip` is missing.
5. Downloads both ZIPs and computes their SHA256 values through `Installer/winget/generate-manifest.ps1`.
6. Runs a self-contained `winget validate --manifest` wrapper in the workflow. It is intentionally inline because the workflow checks out the release tag before generating the manifest, and older release tags may not contain helper scripts added later. The wrapper treats the known `winget.exe v1.11.x` schema-header warning for `ManifestVersion: 1.12.0` as non-fatal, but only when the manifest otherwise reports validation success and all warnings are that exact legacy schema-header warning.
7. Enables Winget `LocalManifestFiles`, runs `winget install --manifest winget-manifest` on the disposable runner, checks that `RedSalamander` appears in `winget list`, and then uninstalls it with `winget uninstall --id RedSalamanders.RedSalamander --purge`.
8. Submits the generated manifest directory with `wingetcreate submit` only for `release.published` runs or manual runs where `submit=true`.

`WINGET_TOKEN` must be a GitHub personal access token with the permissions required by WingetCreate to open a pull request against `microsoft/winget-pkgs`.

## Reintroducing MSI

MSI can return only after all of these are true:

- A trusted Authenticode code-signing certificate is available to CI.
- CI stores the certificate and password in secrets, for example `MSI_SIGNING_CERT` and `MSI_SIGNING_PASSWORD`.
- The release workflow signs every MSI with `signtool.exe sign /fd SHA256 /tr <timestamp-url> /td SHA256`.
- The workflow verifies every MSI with `signtool.exe verify /pa`.
- The Winget manifest uses `InstallerType: wix` or `InstallerType: msi`, includes the MSI `ProductCode`, and is validated locally before submission.

Until that exists, MSI scripts may remain for local experimentation, but automated release and Winget publication stay ZIP-only.

## Winget Package Options

Winget supports several installer types, including EXE, ZIP, Inno, Nullsoft, MSI, WiX, APPX, MSIX, Burn, portable, and font. For RedSalamander, the realistic options are:

- Portable ZIP: current recommended path; no installer trust prompt, simple release artifact, easiest to automate.
- Signed MSI or WiX: good future traditional installer path; requires code signing and ProductCode handling.
- Signed MSIX: strong Windows package identity model; requires signing and publisher identity alignment.
- EXE installer: possible if RedSalamander later adopts an installer bootstrapper with silent switches.

Script-based installers are not supported for the public Winget community repository.
