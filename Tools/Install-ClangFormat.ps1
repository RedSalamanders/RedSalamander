<#
.SYNOPSIS
Installs the checksum-pinned native formatter without Python.
.DESCRIPTION
Downloads the pinned Windows x64 wheel, verifies SHA-256 and extracts only clang-format.exe.
Uses the repository-local .build/format cache and preserves the existing formatter on failure.
.PARAMETER WheelPath
Optional existing wheel for an offline install. The same pinned SHA-256 is required.
.OUTPUTS
System.String. The installed native executable path. WhatIf performs no installation or download.
.NOTES
Requires PowerShell 7. Downloads from the package index when no verified cached wheel is available.
Writes only .build/format. Does not require Python, pip, or a native product build.
CI invokes this command before checking changed native files. Failures terminate the command.
.EXAMPLE
./Tools/Install-ClangFormat.ps1 -WhatIf
.EXAMPLE
./Tools/Install-ClangFormat.ps1 -WheelPath C:/Downloads/clang-format.whl
#>
#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param([string]$WheelPath = '')
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Modules/Build/ClangFormat.psm1') -Force
$parameters = @{RepositoryRoot=(Split-Path $PSScriptRoot); WheelPath=$WheelPath}
foreach ($name in @('WhatIf','Confirm')) {
    if ($PSBoundParameters.ContainsKey($name)) { $parameters[$name] = $PSBoundParameters[$name] }
}
Install-RSClangFormat @parameters
$global:LASTEXITCODE = 0
