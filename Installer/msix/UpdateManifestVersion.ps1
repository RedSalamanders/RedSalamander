<#
.SYNOPSIS
    Updates the MSIX package identity version and architecture.
.PARAMETER Version
    Three-part package version, for example 7.0.184.
.PARAMETER Platform
    Target platform, x64 or ARM64.
.PARAMETER ManifestPath
    Path to Package.appxmanifest.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [ValidateSet("x64", "ARM64")]
    [string]$Platform,

    [string]$ManifestPath = (Join-Path $PSScriptRoot "Package.appxmanifest")
)

$ErrorActionPreference = "Stop"

if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "MSIX version input must be a three-part version such as 7.0.184: $Version"
}

if (-not (Test-Path $ManifestPath)) {
    throw "MSIX manifest not found: $ManifestPath"
}

$identityVersion = "$Version.0"
$architecture = if ($Platform -eq "ARM64") { "arm64" } else { "x64" }

$manifest = [xml](Get-Content -Path $ManifestPath -Raw)
$identity = $manifest.Package.Identity
if (-not $identity) {
    throw "MSIX manifest Identity element not found: $ManifestPath"
}

$identity.Version = $identityVersion
$identity.ProcessorArchitecture = $architecture

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$settings = [System.Xml.XmlWriterSettings]::new()
$settings.Encoding = $utf8NoBom
$settings.Indent = $true

$writer = [System.Xml.XmlWriter]::Create($ManifestPath, $settings)
try {
    $manifest.Save($writer)
}
finally {
    $writer.Dispose()
}

Write-Host "Updated MSIX manifest: Version=$identityVersion ProcessorArchitecture=$architecture"
