<#
.SYNOPSIS Verifies the product toolchain through the pinned shared MSVC setup.
.PARAMETER Platform Native target architecture.
.PARAMETER VerifyOnly Performs the shared read-only toolchain check.
.DESCRIPTION Hosted provisioning uses the contained process runner and the existing 45-minute installer budget.
#>
[CmdletBinding()]
param([ValidateSet('x64','ARM64')][string]$Platform='x64',[switch]$VerifyOnly)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
$engine = Join-Path $root '.ci-engine'
$sourceInputs = Join-Path $engine 'Tools/Modules/Ci.SourceInputs.psm1'
if (-not [IO.File]::Exists($sourceInputs)) {
    throw 'Set up .ci-engine at the commit in .ci/project.json before verifying or provisioning MSVC.'
}
Import-Module $sourceInputs -Scope Local
[void](Get-CiPinnedEngineInputsIdentity -Root $root)

$parameters = @{ ProjectRoot=$root; Platform=$Platform; VerifyOnly=$VerifyOnly; InstallerTimeoutSeconds=2700 }
if (-not $VerifyOnly) {
    Import-Module (Join-Path $root 'Tools/Modules/Build/ProcessStreaming.psm1') -Scope Local
    $parameters.InstallerRunner = (Get-Command Invoke-RSStreamingProcess -Module ProcessStreaming -ErrorAction Stop).ScriptBlock
}
& (Join-Path $engine '.github/actions/setup-msvc/Setup.ps1') @parameters
