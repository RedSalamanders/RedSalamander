<#
.SYNOPSIS
Installs the checksum-pinned native formatter without Python.
.DESCRIPTION
Downloads the pinned Windows x64 wheel, verifies SHA-256 and extracts only clang-format.exe.
Uses the repository-local .build/format cache and preserves the existing formatter on failure.
.PARAMETER WheelPath
Optional existing wheel for an offline install. Relative paths use the PowerShell location.
The same pinned SHA-256 is required; hashing and extraction share one open input stream.
.OUTPUTS
System.String. The installed native executable path. WhatIf performs no installation or download.
.NOTES
Requires PowerShell 7. Downloads from the package index when no verified cached wheel is available.
Writes only .build/format. Does not require Python, pip, or a native product build.
Rejects reparse points throughout the output chain. Uses guarded atomic publication and an
exclusive .build/format/.install.lock lease; a competing install fails immediately. The empty
lock file remains in the cache after the lease closes. WhatIf creates no cache or lock file.
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
