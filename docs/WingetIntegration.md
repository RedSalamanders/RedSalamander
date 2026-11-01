# Winget Integration

RedSalamander keeps Winget manifest templates under `Installer/winget/`.
Generated manifests are written to `.build\AppPackages\winget-manifest` by
default.

## Files

| Path | Purpose |
| --- | --- |
| `Installer/winget/generate-manifest.ps1` | Generates versioned manifests from templates and package hashes. |
| `Installer/winget/WingetValidation.ps1` | Wraps `winget validate` and normalizes known schema-header warning behavior. |
| `Installer/winget/templates/RedSalamanders.RedSalamander.yaml` | Package version manifest template. |
| `Installer/winget/templates/RedSalamanders.RedSalamander.locale.en-US.yaml` | English package metadata template. |
| `Installer/winget/templates/RedSalamanders.RedSalamander.installer.yaml` | Installer architecture and archive metadata template. |

## Generate manifests

Build or obtain the x64 and ARM64 portable archives first, then run:

```powershell
.\Installer\winget\generate-manifest.ps1 `
  -Version 7.0.183 `
  -ZipPath .build\AppPackages\RedSalamander-7.0.183-x64-Portable.zip `
  -Arm64ZipPath .build\AppPackages\RedSalamander-7.0.183-ARM64-Portable.zip
```

If `-Version` is omitted, pass a positive `-BuildNumber`; the script constructs
the version from `Common\Version.h` without reading or allocating repository
version state. Both ZIP inputs are mandatory and must use the exact selected
version in their x64/ARM64 portable filenames.

The generator replaces:

- `{VERSION}`
- `{ZIP_SHA256}`
- `{ARM64_ZIP_SHA256}`
- `{RELEASE_DATE}`

## Validate

Use the helper when possible:

```powershell
. .\Installer\winget\WingetValidation.ps1
Invoke-RSWingetManifestValidation -ManifestPath .build\AppPackages\winget-manifest
```

Or call Winget directly:

```powershell
winget validate --manifest .build\AppPackages\winget-manifest
winget install --manifest .build\AppPackages\winget-manifest
```

## Hosted publication

The app-local entrypoint is `.github/workflows/publish.yml`. It runs through
`workflow_dispatch` on `main`, requires the ID of an original successful
`Qualify main Release` run, and offers `Prepare` (the default), `Release`, and
`Winget` operations. It delegates to the reusable publication engine pinned in
`.ci/project.json`; keep the workflow caller and project pin aligned.

`Prepare` previews the selected qualification, `Release` publishes its retained
ZIPs, and `Winget` verifies both native installs before submitting the manifest.
Publication promotes the selected run's qualified artifacts without rebuilding
or restamping them. Both x64 and ARM64 ZIPs are required. See the normative
[Winget publication contract](../Specs/Installer/WingetIntegration.md) and
[hosted CI policy](../Specs/Build/Build_UnifiedCI.md).

Before requesting publication, validate a locally generated manifest with the
current Winget client and confirm its version, public immutable archive URLs,
and SHA-256 values match the selected qualification's release artifacts. Local
validation does not replace the hosted qualification or publication gates.
