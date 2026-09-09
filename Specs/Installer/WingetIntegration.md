# Winget Publication Contract

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

The Winget workflow starts with the x64 and ARM64 portable ZIPs. Both assets are required for publication so Winget can select the native installer for the user's machine. An explicitly x64-only GitHub release remains supported, but `.github/workflows/release.yml` skips its Winget handoff before invoking the publication workflow; it must not publish a dual-architecture manifest whose ARM64 asset does not exist.

The upstream GitHub release is fail closed. `.github/workflows/release.yml` derives the exact package set from
`build_arm64` and `build_msix`; every requested portable leg is mandatory, while MSIX omission is allowed only when
that explicit input disables MSIX. `Tools/Modules/Packaging/ReleaseArtifactPolicy.psm1` requires exact deterministic filenames,
nonempty files, matching x64/ARM64 PE or MSIX architecture, matching MSIX identity metadata, and an exact verified
SHA256 manifest before release creation. A failed build, package, artifact upload/download, validation, or checksum
step prevents publication; a partial requested matrix is never a valid release.
The release workflow is serialized by its generated build/version identity. The Winget handoff runs only after the
requested release matrix and checksum file have been validated and the GitHub release has been created.

All active third-party GitHub Actions are pinned to reviewed full commit SHAs with exact version comments.
Dependabot owns the `github-actions` ecosystem update feed, and CODEOWNERS requires repository-owner review for
workflow, dependency-automation, and release-policy changes.

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
MinimumOSVersion: 10.0.22000.2600
NestedInstallerType: portable
NestedInstallerFiles:
  - RelativeFilePath: RedLauncher.exe
    PortableCommandAlias: RedSalamander
  - RelativeFilePath: red.exe
    PortableCommandAlias: red
InstallationMetadata:
  Files:
    - RelativeFilePath: RedLauncher.exe
      FileType: launch
      DisplayName: RedSalamander
    - RelativeFilePath: red.exe
      FileType: launch
      DisplayName: red
Installers:
  - Architecture: x64
    InstallerUrl: https://github.com/RedSalamanders/RedSalamander/releases/download/v{VERSION}/RedSalamander-{VERSION}-x64-Portable.zip
    InstallerSha256: {ZIP_SHA256}
  - Architecture: arm64
    InstallerUrl: https://github.com/RedSalamanders/RedSalamander/releases/download/v{VERSION}/RedSalamander-{VERSION}-ARM64-Portable.zip
    InstallerSha256: {ARM64_ZIP_SHA256}
```

Do not use `InstallerType: portable` for a ZIP archive. ZIP archives require `InstallerType: zip` plus `NestedInstallerType: portable`.

RedSalamander requires Windows 11 build 22000.2600 or later. Keep `MinimumOSVersion: 10.0.22000.2600` aligned with the MSIX package floor and the runtime gates in the launch executables. The shared MSBuild Windows SDK target remains `10.0.26100.0` as the compile-time SDK baseline. Runtime gates compare the OS major/minor/build from `RtlGetVersion`; when the OS build equals `22000`, they also read the Windows `UBR` registry value to enforce the `.2600` update-build revision.

Declare `NestedInstallerFiles` at the manifest root because the executable paths are identical in x64 and ARM64 ZIPs. The alias targets are launcher shims, not `RedSalamander.exe`: Winget creates command aliases under its `Links` directory, and launching a symlinked executable directly can make the Windows loader search that alias directory before the package root. `RedLauncher.exe` is a dependency-free static-CRT console-subsystem shim with a detached console allocation manifest where supported; `red.exe` is the build-produced byte-for-byte alias copy of `RedLauncher.exe`. Both resolve their own final symlink target, find the package root, and start the real package-root `RedSalamander.exe` by absolute path so `Common.dll`, `yyjson.dll`, and other app-local DLLs are found without a console flash. Normal GUI launches, including diagnostic flags such as `--etw` and `--perf`, return control to the console immediately. Foreground self-test invocations use the same launcher implementation, wait for the app process, and propagate the real app exit code. Keep `InstallationMetadata.Files` with `FileType: launch` pointing at `RedLauncher.exe` and `red.exe` so Winget's repository validation has explicit launch executables to locate after installation.

Use Winget's `Architecture` field for CPU selection: `x64` is the Intel/AMD 64-bit build, and `arm64` is the native Windows on ARM build. By default Winget chooses from the installers compatible with the current machine; users can override that choice with `winget install --architecture x64` or `winget install --architecture arm64`.

The portable ZIP bundles app-local Microsoft Visual C++ runtime DLLs from the latest installed Visual Studio Build Tools redistributable directory for the target architecture. Do not rely only on Winget `PackageDependencies` for the VC runtime: Winget repository validation launches the installed executable in a clean validation environment and may not install dependencies before that launch check.

Winget's terminal install experience does not render bitmap images from manifests. Console-visible branding is limited to text, so RedSalamander uses `InstallationNotes` for a short post-install message.

Winget has `Icons` metadata for PNG/JPEG/ICO package logos, but `winget validate` currently reports `Field usage requires verified publishers. [Icons]` for the public community manifest path and exits nonzero when RedSalamander adds that field. Do not add `Icons` to the submitted manifest until the publisher is accepted as verified by the Winget community repository. After verification, use `RedSalamander/res/logo64.png`, a version-pinned raw URL, `IconResolution: 64x64`, and a matching `IconSha256` generated from the checked-out release tag.

## Local Manifest Generation

Generate a manifest from an existing portable ZIP:

```powershell
.\Installer\winget\generate-manifest.ps1 `
  -Version 7.0.183 `
  -ZipPath .\.build\AppPackages\RedSalamander-7.0.183-x64-Portable.zip `
  -Arm64ZipPath .\.build\AppPackages\RedSalamander-7.0.183-ARM64-Portable.zip `
  -OutputDir .build\AppPackages\winget-manifest
```

Standalone manifest generation requires either the complete three-part `Version`
or `BuildNumber 1..65535`. It never reads saved repository context or allocates a
local counter; `build.ps1 -GenerateWingetManifest` passes its resolved build number,
while publication automation supplies the release version explicitly.
Both archive inputs are mandatory existing files whose basenames exactly match
`RedSalamander-<version>-x64-Portable.zip` and
`RedSalamander-<version>-ARM64-Portable.zip`; stale or mixed-version archives fail
before any manifest is written.

The standalone generator holds the repository-wide packaging coordination
scope from the first package hash through the last manifest write. It uses
`Release|x64` as required diagnostic evidence; the packaging key itself is
repository-wide because the manifest consumes both architectures. Calls nested
under `build.ps1 -GenerateWingetManifest` reenter safely. Abandoned ownership
records contamination and fails closed with the marker path; the generator does
not clear it automatically.
`OutputDir` is normalized and accepted only beneath this repository's
`.build\AppPackages`. Every existing component from the repository through the
destination must be non-reparse, including after directory creation, so lexical
containment cannot redirect writes outside the guarded root. The repository-wide
lease therefore covers every generated manifest destination.

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

`.github/workflows/winget-release.yml` is a reusable publication workflow. The release
workflow invokes it directly after the complete GitHub release has been published;
`workflow_dispatch` remains available for validation and recovery. This direct edge is
required because a release created with the repository `GITHUB_TOKEN` does not start a
second workflow from the resulting `release.published` event. Manual runs validate and
install-test by default; set the `submit` input to `true` only when the run should open
a `microsoft/winget-pkgs` pull request.

Submission is serialized per package version with a normalized workflow concurrency group. Release events and
direct/manual calls therefore use the same `v<major>.<minor>.<build>` key. Direct/manual inputs are unprefixed
three-part versions; release events require the lowercase `v` tag. Two automation attempts for the same package
version cannot run concurrently.

`Tools/Modules/Packaging/WingetPrPublication.psm1` owns the fail-closed publication state machine. It computes a stable publication
identity from the package, version, exact three manifest paths, and their SHA256 hashes. WingetCreate 1.12.8.0
chooses a GUID-suffixed branch, so the workflow supplies an initial title containing the stable identity marker;
the shared PATCH path moves that marker into the finalized body. A PR is adoptable only when all of these match:

- the token owner's exact GitHub login and `<login>/winget-pkgs` fork;
- WingetCreate's reviewed exact `<package>-<version>-<GUID>` head-branch shape;
- the open state and the stable publication marker in the initial title or finalized body;
- exactly the three expected upstream manifest paths, with only added/modified statuses;
- byte-exact SHA256 content for every generated local manifest.

The state machine creates only when the bounded direct upstream list proves there is no exact-version PR. Exactly
one automation-owned open match resumes through the same idempotent PATCH-and-verify path. External, ambiguous,
multiple, closed, extra-file, or content-mismatched candidates stop without mutation. A submit failure or malformed
WingetCreate output is not authoritative: bounded direct-list retries recover a newly created exact PR after GitHub
eventual consistency. If stable ownership cannot be proven, publication stops; titles are never fuzzy-searched.

The workflow:

1. Resolves the release version from `Common/Version.h` plus `GITHUB_RUN_NUMBER`.
2. Checks out the matching release tag.
3. Reads the GitHub release asset list through the GitHub API.
4. Fails with the available asset names if either `RedSalamander-<version>-x64-Portable.zip` or `RedSalamander-<version>-ARM64-Portable.zip` is missing.
5. Downloads both ZIPs and computes their SHA256 values through `Installer/winget/generate-manifest.ps1`.
6. Runs a self-contained `winget validate --manifest` wrapper in the workflow. It is intentionally inline because the workflow checks out the release tag before generating the manifest, and older release tags may not contain helper scripts added later. The wrapper treats the known `winget.exe v1.11.x` schema-header warning for `ManifestVersion: 1.12.0` as non-fatal, but only when the manifest otherwise reports validation success and all warnings are that exact legacy schema-header warning.
7. Enables Winget `LocalManifestFiles`, runs `winget install --manifest .build\AppPackages\winget-manifest` on the disposable runner, checks that `RedSalamander` appears in `winget list`, and then uninstalls it with `winget uninstall --id RedSalamanders.RedSalamander --purge`.
8. For direct release calls or manual runs where `submit=true`, resolves the token owner, computes the manifest-bound publication marker, and directly lists upstream PRs and changed files to classify the exact-version state.
9. Installs the reviewed WingetCreate package version `1.12.8.0`, resolves its absolute
   executable path, and verifies the executable version before the publication secret
   is exposed. It then submits the generated manifest directory with
   `wingetcreate submit` using Winget's
   `Update: RedSalamanders.RedSalamander to <version>` title format plus the stable creation marker. It does not
   trust or log WingetCreate's PR URL output; the bounded direct-list state machine finds and verifies the PR.
10. Creates or resumes exactly one automation-owned PR, then uses the same idempotent PATCH-and-verify path to set
    the normal title and a marker-bearing completed Winget checklist plus the release URL, installer type,
    architectures, commands, minimum OS, manifest schema, and validation/install/list/uninstall evidence. The
    workflow emits the one verified PR URL returned by this path.

`WINGET_TOKEN` must be a GitHub personal access token with the permissions required by WingetCreate to open a pull request against `microsoft/winget-pkgs` and update that PR's title and body. The workflow exposes it only to the verified submit step as `WINGET_CREATE_GITHUB_TOKEN`; it must never appear in a command-line argument, generated script, log, summary, or artifact. Publication recovery tests use injected repository and submit operations and never use a production token.

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
