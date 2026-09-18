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
identity from the package, version, exact three manifest paths, and the SHA256 of each manifest's canonical content.
WingetCreate 1.12.8.0 chooses a GUID-suffixed branch, so the workflow supplies an initial title containing the stable
identity marker; the shared PATCH path moves that marker into the finalized body. A PR is adoptable only when all of
these match:

- the token owner's exact GitHub login and `<login>/winget-pkgs` fork;
- WingetCreate's reviewed exact `<package>-<version>-<GUID>` head-branch shape;
- the open state and the stable publication marker in the initial title or finalized body;
- exactly the three expected upstream manifest paths, with only added/modified statuses;
- identical canonical content for every generated local manifest.

Canonical content, not bytes, because `wingetcreate submit` re-serializes manifests before committing them: it
prepends `# Created using wingetcreate <version>`, writes CRLF, re-indents sequences to column zero, and reorders
keys (`ReleaseDate` moves after `Installers`). `ConvertTo-RSWingetCanonicalManifest` drops comment lines, blank
lines, indentation, `- ` sequence markers, and line endings, then sorts the remaining lines ordinally;
`Get-RSWingetManifestContentHash` hashes that form after stripping a UTF-8 BOM. Every scalar line survives, so a
changed installer URL or SHA256, version, alias, date, or a missing line still changes the hash, while the generated
file and its committed re-serialization hash identically (verified against the v7.0.59 manifests committed in
[winget-pkgs#436389](https://github.com/microsoft/winget-pkgs/pull/436389)). The workflow's `GetManifestContentHash`
operation applies the same helper to the PR head's file contents.

The identity marker embeds those hashes, so the generated manifests must be identical across reruns. The manifest
`ReleaseDate` therefore comes from the GitHub release's `published_at` date (UTC, `yyyy-MM-dd`), which the workflow
passes to `generate-manifest.ps1 -ReleaseDate`; the generator validates the format and stamps today's date only
when no date is supplied (bare local runs). Release tags whose generator predates `-ReleaseDate` fall back to today
with a workflow warning, and a rerun of such a tag on a later day cannot resume its earlier PR. Marker mismatches on
the token owner's own exact-version PR stop as `not owned`; the earlier PR must be closed or finalized by hand.

Before any pull request is listed, the state machine reads upstream `master`: it resolves the head commit and lists
the exact `manifests/<p>/<Publisher>/<Name>/<version>` directory (a 404 means not published). When the directory holds
exactly the three expected files with identical canonical content, the state is `Published`: publication reports the
manifest tree URL with `Published = $true`, submits nothing, and patches nothing. A directory with a different file set
or different canonical content is a `Conflict`. This check exists because the pull-request list covers only the 100
most recently updated upstream PRs, so a merged publication drops out of it within hours; without it, the v7.0.60 rerun
re-submitted an already merged version and upstream closed the empty PR as "does not update any files". The same
check runs during the visibility retries, so a submission that upstream merges before the list ever shows it also
ends as `Published`.

The state machine creates only when upstream master lacks the version and the bounded direct upstream list proves
there is no exact-version PR. Exactly one automation-owned open match resumes through the same idempotent
PATCH-and-verify path. External, ambiguous, multiple, closed, extra-file, or content-mismatched candidates stop
without mutation; an owned candidate whose changed-file list is still empty is `Pending`, because GitHub lists a
fresh PR's files a few seconds after the PR itself exists. A submit failure or malformed WingetCreate output is not
authoritative: bounded direct-list retries recover a newly created exact PR after GitHub eventual consistency. If
stable ownership cannot be proven, publication stops; titles are never fuzzy-searched.

Each state check lists the 100 most recently updated upstream PRs (two pages of 50), but it reads a changed-file list
(one request per PR) only for candidate PRs: those authored by the token owner, those carrying the publication marker,
and those naming the package identifier (case-insensitive) in their title, body, or head branch. Every WingetCreate,
Komac, and bot submission names the package in at least its head branch, so exact-version conflicts from other
submitters are still inspected, while the ~100 unrelated PRs that dominate the list cost nothing beyond the listing.
Ownership and conflict decisions are unchanged: a candidate is still adoptable only under the exact rules above, and
the token owner's PRs are always inspected even when a maintainer renamed the title or branch, in which case they
stop as not owned rather than being resumed. Before this filter a single check cost about 100 requests and a stopped
publication about 700.

Because a check no longer takes most of a minute, the visibility waits carry GitHub's eventual-consistency window
instead: the workflow and the module default sleep 15, 30, 45, 60, and 60 seconds between the six attempts (about
3.5 minutes in total). A fresh PR's file list and `contents?ref=<sha>` reads can lag its creation by minutes; the
v7.0.59 rerun created [winget-pkgs#436389](https://github.com/microsoft/winget-pkgs/pull/436389) and then stopped
because the manifest content stayed unreadable for all six attempts. A `Pending` reason for an unreadable file list
or manifest content therefore names the pull request, the path and head commit, and the upstream exception message,
so a 404 that never clears and a 403 rate limit are distinguishable in the final failure.

Every paged upstream read enumerates each page before counting it. `Invoke-RestMethod` returns a JSON array as a
single object, so a page collected without enumeration looks like one unusable result: paging stops after the first
page and the scan inspects the page array instead of each pull request. Publication tests therefore return page
results unenumerated, exactly as `Invoke-RestMethod` does, so this shape is covered rather than assumed.

When publication stops because no exact automation-owned pull request became visible, the failure names the last
state kind and reason and states whether submission ever ran. A state that stays unreadable must not be reported as
a bare visibility timeout, because that hides the upstream error that actually blocked the release.

For the same reason, a failed submission must not be reported as a bare `wingetcreate submit` exit code. When
submission ran and publication still stops, the failure carries WingetCreate's captured output as diagnostics:
blank lines are dropped, token-shaped values (`ghp_`/`gho_`/`ghu_`/`ghs_`/`ghr_` and `github_pat_` tokens, and any
`Bearer`/`token`/`password` credential value) are redacted to `***`, each line is capped at 500 characters, and only
the last 40 lines are kept with a count of the omitted earlier lines. The output is never parsed for the PR URL or
used to decide state; the direct-list state machine remains the only authority. The workflow itself never echoes the
raw output. Because the v7.0.59 release stopped with only `wingetcreate submit failed with exit code 1` while the
automation fork was 25,181 commits behind `microsoft/winget-pkgs`, the fork-sync error that actually blocked the
release was invisible; this contract exists so that cannot recur.

The workflow:

1. Resolves the release version from `Common/Version.h` plus `GITHUB_RUN_NUMBER`.
2. Checks out the matching release tag.
3. Reads the GitHub release asset list through the GitHub API.
4. Fails with the available asset names if either `RedSalamander-<version>-x64-Portable.zip` or `RedSalamander-<version>-ARM64-Portable.zip` is missing.
5. Downloads both ZIPs, records the release's `published_at` date, and computes the ZIP SHA256 values through
   `Installer/winget/generate-manifest.ps1 -ReleaseDate <published date>` when the release tag's generator accepts
   that parameter.
6. Runs a self-contained `winget validate --manifest` wrapper in the workflow. It is intentionally inline because the workflow checks out the release tag before generating the manifest, and older release tags may not contain helper scripts added later. The wrapper treats the known `winget.exe v1.11.x` schema-header warning for `ManifestVersion: 1.12.0` as non-fatal, but only when the manifest otherwise reports validation success and all warnings are that exact legacy schema-header warning.
7. Enables Winget `LocalManifestFiles`, runs `winget install --manifest .build\AppPackages\winget-manifest` on the disposable runner, checks that `RedSalamander` appears in `winget list`, and then uninstalls it with `winget uninstall --id RedSalamanders.RedSalamander --purge`. That cleanup always runs, including when an earlier step failed before installing, so it first lists the package and reports nothing to clean up when it is absent rather than failing the job on a `winget uninstall` that found no match. Only Winget's no-match result (`0x8A150014`) counts as absent; any other `winget list` failure fails the step, so a source or service error can never silently leave an installed package behind.
8. For direct release calls or manual runs where `submit=true`, resolves the token owner and logs its login, its
   `<login>/winget-pkgs` head repository, and the classic scope names reported by `X-OAuth-Scopes` (never the token
   value). WingetCreate needs a classic token carrying `public_repo` (or `repo`); a fine-grained token reports no
   classic scopes and is unsupported by winget-create 1.12.8.0, so a missing scope is named in a warning before
   submission instead of surfacing later as an opaque submit failure. It then computes the manifest-bound publication
   marker and directly lists upstream PRs and changed files to classify the exact-version state.
9. Installs the reviewed WingetCreate package version `1.12.8.0`, resolves its absolute
   executable path, and verifies the executable version before the publication secret
   is exposed. Verification reads the version from the `WingetCreateCLI <version>+<commit>`
   banner line, because `wingetcreate --version` also prints its verb help; the whole
   captured output is never compared against the reviewed version. The match must reach the
   `+<commit>` delimiter, so a longer or pre-release version whose leading components equal
   the reviewed version cannot satisfy the exact gate. Because `wingetcreate --version` exits
   non-zero when it appends that help, the step clears the probe's exit code and ends with an
   explicit success so the version probe cannot fail a step whose install and banner gates both
   passed. The reviewed-version install itself is still checked against its own exit code.
   It then submits the generated manifest directory with
   `wingetcreate submit` using Winget's
   `Update: RedSalamanders.RedSalamander to <version>` title format plus the stable creation marker. It does not
   trust WingetCreate's PR URL output; the bounded direct-list state machine finds and verifies the PR. WingetCreate's
   output is kept only as redacted, bounded diagnostics for a publication that stops.
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
