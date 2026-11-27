# Winget Integration

Build or obtain matching x64 and ARM64 portable ZIPs. The root command finds
each archive in `.build/packages` or `.build/AppPackages` and generates the
manifest for the specified version.

```powershell
.\winget-manifest.ps1 -Version 7.0.183
```

The command validates the staged manifest with WinGet and writes the three files to
`.build/packages/winget-manifest/7.0.183`. Use `-X64ZipPath` and `-Arm64ZipPath`
to select archives elsewhere, or `-SkipValidation` to generate files on a host
without WinGet. The ZIP filenames must include the selected version and architecture.

`Installer/winget/generate-manifest.ps1` retains its existing arguments and
defaults to `.build/AppPackages/winget-manifest`.

For a local install check, enable LocalManifestFiles on the test host, then point
WinGet at the generated version directory.

```powershell
winget install --manifest .build/packages/winget-manifest/7.0.183
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
