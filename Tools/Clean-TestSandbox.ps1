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

    Apply refuses a parseable run whose owner is still live, using the same
    PID-reuse checks as the runner. Review manually named runs and preserve useful
    results before applying cleanup; an unparseable name is not ownership proof.

.OUTPUTS
    One PSCustomObject target record when the selected run exists.

.NOTES
    Prerequisites: an initialized root with a valid ownership marker. Side effects:
    -Apply recursively removes only the exact selected run below the marked root and
    honors ShouldProcess. Before any walk it looks for test-owned VHDX receipts under
    runs\<RunId>\artifacts\test-owned-vhdx: elevated, it unmounts, detaches and deletes each
    recorded disk; not elevated, it deletes a VHDX whose receipt is 'detached' (kept after a
    failed FileOps suite) and otherwise refuses the whole run (Status Failed) and reports the
    receipt path with the exact recovery commands. The removal never follows reparse points:
    junctions and directory symlinks are unlinked, and a run containing a volume mount point
    is refused whole (Status Failed). It never cleans legacy paths outside X:\RedSalamander.Perf.
    Exit code is 0 for preview, WhatIf, an absent target, or successful removal;
    invalid invocation/path authority and failed removal return a nonzero exit.
    A failed removal also emits its structured Status Failed record and warning.

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
    exit 0
}

$result = [pscustomobject]@{
    Path = $runRoot
    RunId = $context.RunId
    Category = 'exact-test-run-dir'
    Status = 'Planned'
    Error = $null
}

if (-not $Apply) {
    foreach ($receiptPath in @(Get-RSTestOwnedVirtualDiskReceiptPaths -RunRoot $runRoot)) {
        Write-Warning "The run records a test-owned VHDX ($receiptPath); -Apply deletes a detached one and tears down any other only from an elevated prompt."
    }
    Write-Host 'Dry run: pass -Apply to remove this exact test run directory.' -ForegroundColor Yellow
    $result
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($runRoot, 'Remove exact RedSalamander test run')) {
    $result.Status = 'Skipped'
    $result
    exit 0
}

try {
    # Reuse the runner's live-owner/PID-reuse policy rather than reinterpreting a PID.
    $runInfo = ConvertFrom-RSTestRunId -RunId $context.RunId
    if ($null -ne $runInfo) {
        $inactiveTarget = @(Resolve-RSTestSandboxStaleRunTargets -TestRoot $context.TestRoot |
                Where-Object { [string]::Equals($_.RunId, $context.RunId, [StringComparison]::OrdinalIgnoreCase) })
        if ($inactiveTarget.Count -ne 1) {
            throw "Refusing to remove TestSandbox run '$runRoot': live owner $($runInfo.ProcessId) remains protected."
        }
    }
    Remove-RSTestSandboxExactRunDirectory -TestRoot $context.TestRoot -RunId $context.RunId
    $result.Status = 'Removed'
} catch {
    $result.Status = 'Failed'
    $result.Error = $_.Exception.Message
    Write-Warning "Failed to remove exact RedSalamander test run '$runRoot': $($result.Error)"
}

$result
if ($result.Status -eq 'Failed') { exit 1 }
exit 0
