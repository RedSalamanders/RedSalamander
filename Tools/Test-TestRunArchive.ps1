<#
.SYNOPSIS
    Validates TestRuns archives and paired Terminal VT upgrade evidence.

.DESCRIPTION
    Applies the canonical TestRuns archive path, area/version schema, generated-file, identity-binding, and size policies to changed files, explicitly selected run directories, or the complete checked-in inventory. With BaselineRunPath and CandidateRunPath, also recomputes Terminal VT nearest-rank percentiles and enforces the same-machine, same-corpus, reviewed-delta, and baseline x1.10 admission contract.

.PARAMETER RepoRoot
    Specifies the repository whose Specs/TestRuns archive is validated.

.PARAMETER RunPath
    Selects one or more run directories for explicit pre-stage validation.

.PARAMETER Inventory
    Validates every archive added or modified since its immutable contract baseline.
    The history-free import uses its parentless pre-contract snapshot. Missing or
    unrelated baseline ancestry validates every tracked archive instead.

.PARAMETER BaselineRunPath
    Selects the Product/baseline Terminal VT run for paired admission comparison. CandidateRunPath is required.

.PARAMETER CandidateRunPath
    Selects the candidate Terminal VT run for paired admission comparison. BaselineRunPath is required.

.PARAMETER ZigToolchainTransition
    Declares two distinct reviewed Zig versions in baseline,candidate order for a complete-upgrade comparison. Both recorded versions must match exactly. All other gates remain mandatory. Requires paired comparison; omission requires identical Zig versions.

.PARAMETER MaxFileBytes
    Sets the maximum allowed size for one archived file.

.PARAMETER MaxRunBytes
    Sets the maximum total size for one archived run.

.OUTPUTS
    No supported pipeline objects. Writes violations as errors or one success message; exit code is 1 on any violation and 0 otherwise.

.NOTES
    Prerequisites: Git for changed-file mode and the canonical TestRuns archive module. Side effects: none; validation is read-only. Inventory and RunPath are mutually exclusive. Primary consumers: pre-stage/CI evidence checks and performance closeout.

.EXAMPLE
    .\Tools\Test-TestRunArchive.ps1

.EXAMPLE
    .\Tools\Test-TestRunArchive.ps1 -Inventory

.EXAMPLE
    ./Tools/Test-TestRunArchive.ps1 -BaselineRunPath <baseline-run> -CandidateRunPath <candidate-run> -ZigToolchainTransition 0.16.0,0.17.0
    Compares a complete reviewed compiler/runtime upgrade; both recorded Zig versions must match exactly.

.EXAMPLE
    .\Tools\Test-TestRunArchive.ps1 -BaselineRunPath .\Specs\TestRuns\4cb089111a23\Terminal\20260826_071812_pair2-product -CandidateRunPath .\Specs\TestRuns\4cb089111a23\Terminal\20260826_071820_pair2-candidate
#>

[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string[]]$RunPath = @(),
    [switch]$Inventory,
    [string]$BaselineRunPath = '',
    [string]$CandidateRunPath = '',
    [string[]]$ZigToolchainTransition = @(),
    [long]$MaxFileBytes = 2MB,
    [long]$MaxRunBytes = 5MB
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'Modules\Testing\TestRunArchive.psm1') -Force -ErrorAction Stop

$pairRequested = -not [string]::IsNullOrWhiteSpace($BaselineRunPath) -or
    -not [string]::IsNullOrWhiteSpace($CandidateRunPath)
if ($pairRequested -and ([string]::IsNullOrWhiteSpace($BaselineRunPath) -or
        [string]::IsNullOrWhiteSpace($CandidateRunPath))) {
    throw '-BaselineRunPath and -CandidateRunPath must be supplied together.'
}
if (-not $pairRequested -and $ZigToolchainTransition.Count -ne 0) {
    throw '-ZigToolchainTransition requires paired Terminal VT comparison.'
}
if (($Inventory -and $RunPath.Count -ne 0) -or ($pairRequested -and ($Inventory -or $RunPath.Count -ne 0))) {
    throw '-Inventory, -RunPath, and paired Terminal VT comparison are mutually exclusive.'
}

$mode = 'changed'
if ($pairRequested) {
    $mode = 'terminal-vt-pair'
    $paths = @(Get-RSTestRunArchivePathsForRuns `
        -RepoRoot $RepoRoot `
        -RunPath @($BaselineRunPath, $CandidateRunPath))
} elseif ($Inventory) {
    $mode = 'whole-inventory'
    $paths = @(Get-RSTestRunArchiveInventoryPaths -RepoRoot $RepoRoot)
} elseif ($RunPath.Count -ne 0) {
    $mode = 'explicit-pre-stage'
    $paths = @(Get-RSTestRunArchivePathsForRuns -RepoRoot $RepoRoot -RunPath $RunPath)
} else {
    $paths = @(Get-RSChangedTestRunArchivePaths -RepoRoot $RepoRoot)
}

$violations = @(Get-RSTestRunArchiveViolations `
    -RepoRoot $RepoRoot `
    -Paths $paths `
    -MaxFileBytes $MaxFileBytes `
    -MaxRunBytes $MaxRunBytes)

$comparison = $null
if ($pairRequested -and $violations.Count -eq 0) {
    $comparison = Compare-RSTerminalVtUpgradeEvidence `
        -RepoRoot $RepoRoot `
        -BaselineRunPath $BaselineRunPath `
        -CandidateRunPath $CandidateRunPath `
        -ZigToolchainTransition $ZigToolchainTransition
    $violations = @($comparison.Violations)
}

if ($violations.Count -ne 0) {
    foreach ($violation in $violations) {
        Write-Error "$($violation.Kind): $($violation.Path): $($violation.Message)" -ErrorAction Continue
    }
    exit 1
}

if ($null -ne $comparison) {
    if ($comparison.ZigToolchainTransition.Count -eq 2) {
        Write-Host ("Reviewed Zig toolchain transition: {0} -> {1}; results compare the complete upgrade." -f
            $comparison.ZigToolchainTransition[0], $comparison.ZigToolchainTransition[1])
    }
    foreach ($measurement in $comparison.Measurements) {
        Write-Host ("{0}: baseline p95={1}us, candidate p95={2}us, ceiling={3}us" -f `
            $measurement.Name, $measurement.BaselineP95, $measurement.CandidateP95, $measurement.CandidateMaximumP95)
    }
}
Write-Host "TestRuns archive contract passed in $mode mode for $($paths.Count) file(s)."
exit 0
