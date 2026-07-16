<#
.SYNOPSIS
    Validates changed checked-in TestRuns artifacts.
#>

[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [long]$MaxFileBytes = 2MB,
    [long]$MaxRunBytes = 5MB
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'TestRunArchive.psm1') -Force

$paths = @(Get-RSChangedTestRunArchivePaths -RepoRoot $RepoRoot)
$violations = @(Get-RSTestRunArchiveViolations `
    -RepoRoot $RepoRoot `
    -Paths $paths `
    -MaxFileBytes $MaxFileBytes `
    -MaxRunBytes $MaxRunBytes)

if ($violations.Count -ne 0) {
    foreach ($violation in $violations) {
        Write-Error "$($violation.Kind): $($violation.Path): $($violation.Message)" -ErrorAction Continue
    }
    exit 1
}

Write-Host "TestRuns archive contract passed for $($paths.Count) changed file(s)."
exit 0
