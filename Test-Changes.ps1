<#
.SYNOPSIS
Runs the governed affected plan or the local portion of a full PR gate.
.DESCRIPTION
Reuses Run-AllTests impact, receipts and explicit exact Resume. Default iteration is
Affected. PrePush defers only entries whose complete portable contracts match Suite PR
in the enabled GitHub workflow with reviewed candidate bytes and its complete repository-owned
workflow/action/tooling execution digest. Missing or changed dependencies keep work local.
Nightly/weekly coverage is never delegated.
.PARAMETER Mode
Affected (default), Full, or PrePush.
.PARAMETER BaseRef
Local comparison revision; defaults to origin/HEAD, the remote's recorded default branch. No fetch is performed.
.PARAMETER Configuration
Debug, Release or ASan Debug.
.PARAMETER Platform
x64 or ARM64; runtime requires native execution.
.PARAMETER Explain
Publishes the canonical plan without building or launching tests.
.PARAMETER SkipBuild
Requires the canonical runner's byte-verified solution receipt.
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
Canonical evidence under the marked sandbox. Exit 2 is successful selected work with
incomplete repository coverage; failures and invalid/blocked/corrupt evidence retain
the canonical runner's other exit categories. PrePush never reports a full pass.
.NOTES
Requires PowerShell 7, Git, the native toolchain and a marked sandbox. Builds and tests
through the canonical runner; no second evidence cache or desktop lease is introduced.
.EXAMPLE
./Test-Changes.ps1 -Explain
.EXAMPLE
./Test-Changes.ps1 -Suite Commands -CommandsFamily file-operations -SkipBuild
.EXAMPLE
./Test-Changes.ps1 -Mode PrePush -Configuration Debug
.EXAMPLE
./Test-Changes.ps1 -Mode Full -Interactive -AllowAlternateVolumeTestRoot -TestRoot E:\RedSalamander.Perf
#>
[CmdletBinding()]
param(
    [ValidateSet('Affected','Full','PrePush')][string]$Mode='Affected',
    [string]$BaseRef='origin/HEAD',
    [ValidateSet('Debug','Release','ASan Debug')][string]$Configuration='Debug',
    [ValidateSet('x64','ARM64')][string]$Platform='x64',
    [switch]$Explain,
    [switch]$SkipBuild,
    [string]$ResumeFrom='',
    [switch]$Interactive,
    [string]$TestRoot=$env:REDSALAMANDER_TEST_ROOT,
    [switch]$AllowAlternateVolumeTestRoot,
    [ValidateSet('','Commands','Compare','FileOps')][string]$Suite='',
    [string]$CaseFilter='',
    [string]$CommandsFamily=''
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
# Pending fragments are expected: they are folded on the default branch after merge. CHANGELOG.md may change only by such a fold.
# BaseRef defaults to origin/HEAD for the impact comparison; the changelog is compared with origin/main unless a BaseRef is supplied.
if (-not $Explain -and $Mode -eq 'PrePush') { & (Join-Path $PSScriptRoot 'Tools/Test-Changelog.ps1') -Base $(if ($PSBoundParameters.ContainsKey('BaseRef') -and $BaseRef) { $BaseRef } else { 'origin/main' }) }
Import-Module (Join-Path $PSScriptRoot 'Tools/Modules/Testing/ScopedTesting.psm1') -Force
Write-Host "Native test naming: $(Assert-ScopedTestNames $PSScriptRoot) files checked."
if ($Suite -and $Mode -ne 'Affected') {throw 'Focused Suite is valid only for Affected iteration.'}
if ($ResumeFrom -and ($Mode -ne 'Full' -or -not $Interactive)) {throw 'Exact Full Resume requires Mode Full and the same interactive coverage as its origin.'}
$arguments=@{Suite='Full';Configuration=$Configuration;Platform=$Platform;ImpactBase=$BaseRef;SkipBuild=$SkipBuild;NonInteractive=(-not $Interactive);ValidationMode='Fresh';TestRoot=$TestRoot;AllowAlternateVolumeTestRoot=$AllowAlternateVolumeTestRoot}
if ($Mode -eq 'Affected') {$arguments.ValidationMode='Affected'}
if ($Suite) {$arguments.Suite=$Suite;$arguments.ValidationMode='Fresh';$arguments.CaseFilter=$CaseFilter;$arguments.CommandsFamily=$CommandsFamily}
elseif ($CaseFilter -or $CommandsFamily) {throw 'CaseFilter and CommandsFamily require a focused Suite.'}
if ($ResumeFrom) {$arguments.ValidationMode='Resume';$arguments.ResumeFrom=$ResumeFrom}
if ($Mode -eq 'PrePush') {
    $manifest=Read-ScopedTestManifest $PSScriptRoot
    $covered=@(Get-ScopedPrCoverage $PSScriptRoot $manifest $Platform $Configuration)
    if ('PR' -in $covered) {$arguments.DeferPrGate=$true;Write-Host 'PR-equivalent entries are delegated to GitHub. CI success remains pending.'}
    else {Write-Host 'PR equivalence could not be verified; full local obligations remain selected.'}
}
if ($Explain) {$arguments.ExplainPlan=$true}
& (Join-Path $PSScriptRoot 'Tools/Run-AllTests.ps1') @arguments
exit $LASTEXITCODE
