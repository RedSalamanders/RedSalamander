<#
.SYNOPSIS
Builds and verifies a portable Release ZIP for RedSalamander.
.PARAMETER Platform
Target x64 or ARM64 architecture. Defaults to x64.
.PARAMETER BuildNumber
Optional build number for the compiled outputs and archive. Defaults to the build command's current version policy.
.OUTPUTS
A clean-extraction checked ZIP under .build/packages and the build receipt.
.NOTES
Uses the native Release build and package pipeline. Run on the target architecture for execution evidence.
.EXAMPLE
./package.ps1 -Platform x64
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform = 'x64',

    [ValidateRange(0, 65535)]
    [int]$BuildNumber = 0
)

$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'build.ps1') -Configuration Release -Platform $Platform -BuildNumber $BuildNumber -Zip -CanonicalPackageDirectory
if (-not $?) { exit $LASTEXITCODE }
