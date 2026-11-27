<#
.SYNOPSIS
Generates the WinGet manifests for both portable Release ZIPs.
.PARAMETER Version
Exact three-part package version with a positive build number.
.PARAMETER X64ZipPath
x64 package path. By default, checks .build/packages then .build/AppPackages.
.PARAMETER Arm64ZipPath
ARM64 package path. By default, checks the same package roots.
.PARAMETER ReleaseDate
Manifest release date in yyyy-MM-dd. The local date is the default.
.PARAMETER OutputDirectory
Directory beneath .build/packages. Defaults to .build/packages/winget-manifest/<version>.
.PARAMETER SkipValidation
Generates the manifests without invoking the installed WinGet validator.
.OUTPUTS
Three WinGet manifest files and console validation results.
.NOTES
Requires the pinned .ci-engine checkout. WinGet is required for default validation.
.EXAMPLE
./winget-manifest.ps1 -Version 7.0.184
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Version,
    [Alias('ZipPath')]
    [string]$X64ZipPath,
    [string]$Arm64ZipPath,
    [string]$ReleaseDate,
    [Alias('OutputDir')]
    [string]$OutputDirectory,
    [switch]$SkipValidation
)

$ErrorActionPreference = 'Stop'
$arguments = @{
    Version = $Version
    ZipPath = $X64ZipPath
    Arm64ZipPath = $Arm64ZipPath
    ReleaseDate = $ReleaseDate
    OutputDir = $OutputDirectory
    CanonicalPackageDirectory = $true
    ValidateManifest = -not $SkipValidation
}
& (Join-Path $PSScriptRoot 'Installer/winget/generate-manifest.ps1') @arguments
if (-not $?) { exit $LASTEXITCODE }
