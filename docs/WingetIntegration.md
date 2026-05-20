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
  -ZipPath .build\AppPackages\RedSalamander-7.0.183-x64.zip `
  -Arm64ZipPath .build\AppPackages\RedSalamander-7.0.183-arm64.zip
```

If `-Version` is omitted, the script derives the package version through
`Tools\Versioning.ps1`.

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

## Submission checklist

- Confirm the generated `PackageVersion` matches the GitHub release tag.
- Confirm both archive URLs are public and immutable.
- Confirm x64 and ARM64 SHA256 values came from the final release archives.
- Confirm `DocumentUrl` points at `docs/README.md`.
- Validate locally with the current Winget client.
- Submit the generated manifest directory to `microsoft/winget-pkgs` following
  the repository contribution process.
