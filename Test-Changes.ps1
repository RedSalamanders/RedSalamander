<#
.SYNOPSIS
Compatibility entry point for RedSalamander scoped tests.
.DESCRIPTION
Forwards the existing parameters to the root test.ps1, preserving single-profile
execution and returning its exit code. Shared scope options select the pinned engine.
.PARAMETER Root
Checkout containing the RedSalamander root test command. Defaults to this checkout.
.PARAMETER Mode
Affected (default), Full, or PrePush.
.PARAMETER BaseRef
Local comparison revision. The canonical runner defaults to origin/HEAD. The shared runner
uses the default branch in the scope manifest. The comparison uses local refs.
.PARAMETER Scopes
Runs explicit shared scopes in Affected mode. Accepts names or comma-separated strings.
.PARAMETER Force
Uses the shared runner and bypasses successful scope reuse and its iteration baseline.
.PARAMETER ParallelSuites
Selects the shared runner. RedSalamander's product-owned native profiles require one lane.
.PARAMETER Configuration
Debug, Release or ASan Debug.
.PARAMETER Platform
x64 or ARM64; runtime requires native execution.
.PARAMETER Explain
Publishes the canonical plan without building or launching tests.
.PARAMETER SkipBuild
Requires the canonical runner's byte-verified solution receipt or its exact governed
Affected/focused project union. No-build/zero-entry changes need no native receipt.
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
the canonical runner's other exit categories.
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
./Test-Changes.ps1 -Scopes CompilerFreeValidation -Explain
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
    [string]$CommandsFamily='',
    [string]$Root=$PSScriptRoot,
    [string[]]$Scopes=@(),
    [switch]$Force,
    [ValidateRange(1,8)][int]$ParallelSuites=1
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$checkout = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)
if ($PSBoundParameters.Remove('Root')) {
    $project = Get-Content -Raw -LiteralPath (Join-Path $checkout '.ci/project.json') | ConvertFrom-Json
    if ($project.repository -cne 'RedSalamanders/RedSalamander') { throw 'Root must be a RedSalamander checkout.' }
}
if (-not ($Scopes -or $Force -or $PSBoundParameters.ContainsKey('ParallelSuites'))) {
    $PSBoundParameters['CanonicalProfile'] = $true
}
& (Join-Path $checkout 'test.ps1') @PSBoundParameters
exit $LASTEXITCODE
