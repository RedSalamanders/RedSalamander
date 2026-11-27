# Winget Publication Contract

This document defines the supported Windows Package Manager publication flow for RedSalamander.

## Current Packaging Policy

RedSalamander distributes through portable ZIP archives and Winget manifests.
The x64 and ARM64 ZIPs are the only supported product package artifacts.

## Release Artifact Contract

Official product releases consume one original successful `Qualify main Release` run. Its matching receipts and retained x64 and ARM64 ZIPs are the authoritative publication inputs. The hosted publication operation promotes those exact bytes; it does not rebuild or restamp them.

Use the exact version, archive names, and SHA-256 values recorded for that qualification when generating the Winget manifest. Both architecture archives are required. Missing or expired artifacts require a fresh qualification, and conflicting tags, release assets, or Winget submissions are retained for review.

## Manifest Shape

The pinned CI engine renders the three-file manifest from the `winget` settings in `.ci/project.json`.

```text
RedSalamanders.RedSalamander.installer.yaml
RedSalamanders.RedSalamander.locale.en-US.yaml
RedSalamanders.RedSalamander.yaml
```

The installer manifest uses schema `1.12.0`. It declares the Windows minimum version and ZIP installer type. Each architecture entry owns its package URL, SHA-256, nested aliases and launch metadata.

Do not use `InstallerType: portable` for a ZIP archive. ZIP archives require `InstallerType: zip` plus `NestedInstallerType: portable`.

RedSalamander requires Windows 11 build 22000.2600 or later. Keep the Winget manifest's `MinimumOSVersion: 10.0.22000.2600` aligned with the runtime gates in the launch executables. The shared MSBuild Windows SDK target remains `10.0.26100.0` as the compile-time SDK baseline. Runtime gates compare the OS major/minor/build from `RtlGetVersion`; when the OS build equals `22000`, they also read the Windows `UBR` registry value to enforce the `.2600` update-build revision.

Each architecture declares `NestedInstallerFiles` using actual paths from that ZIP. The archive must contain one file for each configured launcher name. The alias targets are launcher shims, not `RedSalamander.exe`: Winget creates command aliases under its `Links` directory, and launching a symlinked executable directly can make the Windows loader search that alias directory before the package root. `RedLauncher.exe` is a dependency-free static-CRT console-subsystem shim with a detached console allocation manifest where supported; `red.exe` is the build-produced byte-for-byte alias copy of `RedLauncher.exe`. Both resolve their own final symlink target, find the package root, and start the real package-root `RedSalamander.exe` by absolute path so `Common.dll`, `yyjson.dll`, and other app-local DLLs are found without a console flash. Normal GUI launches, including diagnostic flags such as `--etw` and `--perf`, return control to the console immediately. Foreground self-test invocations use the same launcher implementation, wait for the app process, and propagate the real app exit code. Keep `InstallationMetadata.Files` with `FileType: launch` pointing at `RedLauncher.exe` and `red.exe` so Winget's repository validation has explicit launch executables to locate after installation.

Use Winget's `Architecture` field for CPU selection: `x64` is the Intel/AMD 64-bit build, and `arm64` is the native Windows on ARM build. By default Winget chooses from the installers compatible with the current machine; users can override that choice with `winget install --architecture x64` or `winget install --architecture arm64`.

The portable ZIP bundles app-local Microsoft Visual C++ runtime DLLs from the latest installed Visual Studio Build Tools redistributable directory for the target architecture. Do not rely only on Winget `PackageDependencies` for the VC runtime: Winget repository validation launches the installed executable in a clean validation environment and may not install dependencies before that launch check.

Winget's terminal install experience does not render bitmap images from manifests. Console-visible branding is limited to text, so RedSalamander uses `InstallationNotes` for a short post-install message.

Winget has `Icons` metadata for PNG/JPEG/ICO package logos, but `winget validate` currently reports `Field usage requires verified publishers. [Icons]` for the public community manifest path and exits nonzero when RedSalamander adds that field. Do not add `Icons` to the submitted manifest until the publisher is accepted as verified by the Winget community repository. After verification, use `RedSalamander/res/logo64.png`, a version-pinned raw URL, `IconResolution: 64x64`, and a matching `IconSha256` generated from the checked-out release tag.

## Local Manifest Generation

Run the root command with the exact portable package version.

```powershell
.\winget-manifest.ps1 -Version 7.0.183
```

The command looks for each versioned ZIP under `.build/packages` first, then under
`.build/AppPackages`. Supply `-X64ZipPath` and `-Arm64ZipPath` to select explicit files.
Each input must have the exact architecture and version in its portable ZIP name.
Outputs default to `.build/packages/winget-manifest/<version>`. The root command runs
`winget validate` on staged manifests by default. `-SkipValidation` permits rendering
on a host without WinGet.

The legacy `Installer/winget/generate-manifest.ps1` accepts the original explicit ZIP paths,
`-OutputDir` and either `-Version` or `-BuildNumber 1..65535`. It uses the same renderer
and defaults to `.build/AppPackages/winget-manifest`. `build.ps1 -GenerateWingetManifest`
passes the build's resolved version and uses the selected package layout. An explicit
`-ReleaseDate` uses calendar format `yyyy-MM-dd`. The local command defaults to today's date.

The repository-wide packaging lease covers input hashing, rendering, validation and publication.
Calls from `build.ps1 -GenerateWingetManifest` reenter that lease. The guarded output root is
`.build/packages` for the root command or `.build/AppPackages` for compatibility. Each existing
path component must be non-reparse. The renderer stages all three manifest files and validates
them before publication when validation is enabled. Each file is published through an atomic
same-directory replacement. Review the contamination marker after abandoned ownership.

## Hosted Publication Entry Point

The app-local publication adapter is `.github/workflows/publish.yml`. It runs only through `workflow_dispatch` on `main`, requires a `qualification_run` ID, and exposes `Prepare`, `Release`, or `Winget`; `Prepare` is the default. The workflow delegates to the reusable publication engine pinned in `.ci/project.json`, and the caller pin must match that engine reference.

Select the original successful `Qualify main Release` run. `Prepare` previews the selected run, `Release` publishes its qualified ZIPs, and `Winget` verifies both native installs before submitting the manifest. The manual publication workflow does not compile or create a new qualification. Its exact artifact, version, checksum, and publication checks are owned by the pinned engine and the project adapter.

Product releases publish portable ZIP assets only. The Winget manifest requires anonymous immutable release URLs, matching ZIP hashes, the configured nested executable, and the project license. The qualification and publication contract is defined in [Build_UnifiedCI.md](../Build/Build_UnifiedCI.md); local manifest shape and generation remain described above.

## Supported package type

RedSalamander publishes only portable ZIP archives through Winget.
