<#
.SYNOPSIS
    Lists or removes one exact RedSalamander test run directory.

.DESCRIPTION
    Resolves only X:\RedSalamander.Perf\runs\<RunId> beneath the selected marked
    test root. Dry-run is the default. Pass -Apply to remove that exact run;
    ShouldProcess still honors -WhatIf and -Confirm. The command never enumerates,
    writes, or removes historical test locations outside the selected root.

.PARAMETER RunId
    Exact direct-child run identifier below X:\RedSalamander.Perf\runs.

.PARAMETER TestRoot
    Exact marked X:\RedSalamander.Perf root. The repository drive is used when
    neither this parameter nor REDSALAMANDER_TEST_ROOT selects another fixed drive.

.PARAMETER Apply
    Removes the exact resolved run directory. Without this switch the command only
    reports the target.

.OUTPUTS
    One PSCustomObject target record when the selected run exists.

.NOTES
    Prerequisites: an initialized root with a valid ownership marker. Side effects:
    -Apply recursively removes only the exact selected run below the marked root and
    honors ShouldProcess. It never cleans legacy paths outside X:\RedSalamander.Perf.

.EXAMPLE
    .\Tools\Clean-TestSandbox.ps1 -RunId 20260830T120000Z-1234-abcdef
    Reports the exact run directory without removing it.

.EXAMPLE
    .\Tools\Clean-TestSandbox.ps1 -RunId 20260830T120000Z-1234-abcdef -Apply -Confirm:$false
    Removes only the exact selected run directory.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$RunId,

    [string]$TestRoot = $env:REDSALAMANDER_TEST_ROOT,

    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
Import-Module (Join-Path $scriptRoot 'Modules\Testing\TestRunPlan.psm1') -Force -ErrorAction Stop

$context = New-RSTestRunContext `
    -RepoRoot $repoRoot `
    -RunId $RunId `
    -TestRootOverride $TestRoot `
    -SelfTestRootOverride ''
$runRoot = Assert-RSTestSandboxContainedPath `
    -TestRoot $context.TestRoot `
    -Path $context.RunRoot `
    -RequireDescendant

if (-not (Test-Path -LiteralPath $runRoot -PathType Container)) {
    return
}

$result = [pscustomobject]@{
    Path = $runRoot
    RunId = $context.RunId
    Category = 'exact-test-run-dir'
    Status = 'Planned'
    Error = $null
}

if (-not $Apply) {
    Write-Host 'Dry run: pass -Apply to remove this exact test run directory.' -ForegroundColor Yellow
    $result
    return
}

if (-not $PSCmdlet.ShouldProcess($runRoot, 'Remove exact RedSalamander test run')) {
    $result.Status = 'Skipped'
    $result
    return
}

try {
    Remove-RSTestSandboxExactRunDirectory -TestRoot $context.TestRoot -RunId $context.RunId
    $result.Status = 'Removed'
} catch {
    $result.Status = 'Failed'
    $result.Error = $_.Exception.Message
    Write-Warning "Failed to remove exact RedSalamander test run '$runRoot': $($result.Error)"
}

$result
