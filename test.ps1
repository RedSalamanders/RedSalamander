<#
.SYNOPSIS
Runs affected RedSalamander tests or explicitly selected full coverage.
.DESCRIPTION
Reuses Run-AllTests impact, receipts and explicit exact Resume. Default iteration is
Affected. PrePush keeps all native obligations local because the PR gate runs
compiler-free checks.
.PARAMETER Mode
Affected (default), Full, or PrePush.
.PARAMETER Full
Selects the full local obligation. Equivalent to Mode Full.
.PARAMETER BaseRef
Local comparison revision; defaults to origin/HEAD, the remote's recorded default branch. No fetch is performed.
.PARAMETER Configuration
Debug, Release or ASan Debug.
.PARAMETER Platform
x64 or ARM64; runtime requires native execution.
.PARAMETER Explain
Publishes the canonical plan without building or launching tests.
.PARAMETER SkipBuild
Requires the canonical runner's byte-verified solution receipt or its exact governed
Affected/focused project union. No-build/zero-entry changes need no native receipt.
.PARAMETER BuildNumber
Build number forwarded to the canonical runner for its build.
.PARAMETER Suites
Canonical test entry IDs selected by the shared scope planner. Runs only those
entries without desktop admission and projects the completed archive to suite reports.
.PARAMETER SkipTooling
Compatibility switch for the shared runner. The Suites adapter runs only the
explicit list, including any named tooling entries in their canonical order.
.PARAMETER ResumeFrom
Explicit exact promoted Full origin; valid only with Mode Full. No automatic widening.
.PARAMETER Interactive
Allows focus-taking entries after the person agrees to the time. Without this switch,
those entries remain pending and the repository verdict is NOT_EVALUATED.
.PARAMETER TestRoot
Existing marked X:\RedSalamander.Perf sandbox. This parameter does not authorize creation.
.PARAMETER AllowAlternateVolumeTestRoot
Explicitly opt in to the canonical runner's alternate-volume File Operations coverage.
The selected sandbox's enabled path.local.alternate resource restricts the exact root.
The default remains disabled; this switch does not authorize initializing a new drive-root workspace.
.PARAMETER Suite
Optional focused Commands, Compare or FileOps suite; valid only in Affected mode.
.PARAMETER CaseFilter
Optional case/prefix within a focused suite. Never qualifies the complete repository.
.PARAMETER CommandsFamily
Optional governed Commands family; valid with Suite Commands.
.OUTPUTS
Canonical evidence stays under the marked sandbox. The Suites adapter also writes
shared suite report metadata under .build/reports. Exit 0 on this path means verified
selected execution. The archived repository verdict remains NOT_EVALUATED.
Ordinary commands keep the canonical runner's exit categories, including exit 2 for
incomplete repository coverage.
.NOTES
Requires PowerShell 7, Git, the native toolchain and a marked sandbox. Builds and tests
through the canonical runner; no second evidence cache or desktop lease is introduced.
.EXAMPLE
./test.ps1 -Explain
.EXAMPLE
./test.ps1 -Suite Commands -CommandsFamily file-operations -SkipBuild
.EXAMPLE
./test.ps1 -Mode PrePush -Configuration Debug
.EXAMPLE
./test.ps1 -Full -Interactive -AllowAlternateVolumeTestRoot -TestRoot E:\RedSalamander.Perf
#>
[CmdletBinding(DefaultParameterSetName='Mode')]
param(
    [Parameter(ParameterSetName='Full',Mandatory)][switch]$Full,
    [Parameter(ParameterSetName='Mode')][ValidateSet('Affected','Full','PrePush')][string]$Mode='Affected',
    [Parameter(ParameterSetName='Selected',Mandatory)][ValidateNotNullOrEmpty()][string[]]$Suites,
    [Parameter(ParameterSetName='Selected')][switch]$SkipTooling,
    [Parameter(ParameterSetName='Mode')][Parameter(ParameterSetName='Full')][string]$BaseRef='origin/HEAD',
    [ValidateSet('Debug','Release','ASan Debug')][string]$Configuration='Debug',
    [ValidateSet('x64','ARM64')][string]$Platform='x64',
    [Parameter(ParameterSetName='Mode')][Parameter(ParameterSetName='Full')][switch]$Explain,
    [switch]$SkipBuild,
    [ValidateRange(0, [int]::MaxValue)][int]$BuildNumber = 0,
    [Parameter(ParameterSetName='Mode')][Parameter(ParameterSetName='Full')][string]$ResumeFrom='',
    [Parameter(ParameterSetName='Mode')][Parameter(ParameterSetName='Full')][switch]$Interactive,
    [string]$TestRoot=$env:REDSALAMANDER_TEST_ROOT,
    [Parameter(ParameterSetName='Mode')][Parameter(ParameterSetName='Full')][switch]$AllowAlternateVolumeTestRoot,
    [Parameter(ParameterSetName='Mode')][ValidateSet('','Commands','Compare','FileOps')][string]$Suite='',
    [Parameter(ParameterSetName='Mode')][string]$CaseFilter='',
    [Parameter(ParameterSetName='Mode')][string]$CommandsFamily=''
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if ($Full) { $Mode = 'Full' }
# Pending fragments are folded on the default branch after merge. The pinned validator checks the exact fold or generated release transition.
# BaseRef defaults to origin/HEAD for the impact comparison; the changelog is compared with origin/main unless a BaseRef is supplied.
if (-not $Explain) {
    $changelogArguments = @{}
    if ($Mode -eq 'PrePush') {
        $changelogArguments.Base = if ($PSBoundParameters.ContainsKey('BaseRef') -and $BaseRef) { $BaseRef } else { 'origin/main' }
    }
    & (Join-Path $PSScriptRoot 'Tools/Test-Changelog.ps1') @changelogArguments
}
Import-Module (Join-Path $PSScriptRoot 'Tools/Modules/Testing/ScopedTesting.psm1') -Force
Write-Host "Native test naming: $(Assert-ScopedTestNames $PSScriptRoot) files checked."
if ($Suite -and $Mode -ne 'Affected') {throw 'Focused Suite is valid only for Affected iteration.'}
if ($ResumeFrom -and ($Mode -ne 'Full' -or -not $Interactive)) {throw 'Exact Full Resume requires Mode Full and the same interactive coverage as its origin.'}
$arguments = @{
    Suite = 'Full'
    Configuration = $Configuration
    Platform = $Platform
    ImpactBase = $BaseRef
    SkipBuild = $SkipBuild
    NonInteractive = -not $Interactive
    ValidationMode = 'Fresh'
    TestRoot = $TestRoot
    AllowAlternateVolumeTestRoot = $AllowAlternateVolumeTestRoot
}
if ($PSBoundParameters.ContainsKey('BuildNumber')) { $arguments.BuildNumber = $BuildNumber }
if ($PSCmdlet.ParameterSetName -eq 'Selected') {
    $arguments.ScopedEntryId = $Suites
} elseif ($Mode -eq 'Affected') {
    $arguments.ValidationMode = 'Affected'
}
if ($Suite) {
    $arguments.Suite = $Suite
    $arguments.ValidationMode = 'Fresh'
    $arguments.CaseFilter = $CaseFilter
    $arguments.CommandsFamily = $CommandsFamily
} elseif ($CaseFilter -or $CommandsFamily) {
    throw 'CaseFilter and CommandsFamily require a focused Suite.'
}
if ($ResumeFrom) {
    $arguments.ValidationMode = 'Resume'
    $arguments.ResumeFrom = $ResumeFrom
}
if ($Mode -eq 'PrePush') {
    Write-Host 'PrePush selects the full local obligation. GitHub PR checks are compiler-free.'
}
if ($Explain) {$arguments.ExplainPlan=$true}
& (Join-Path $PSScriptRoot 'Tools/Run-AllTests.ps1') @arguments
exit $LASTEXITCODE
