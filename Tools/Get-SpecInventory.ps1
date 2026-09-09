<#
.SYNOPSIS
Validates specification authority, links, WIP ownership, and consistency evidence.

.DESCRIPTION
Builds source-derived JSON and Markdown inventories for the repository specification tree, validates normative consistency records and protected completed-plan history, and also emits the generated test inventory.

.PARAMETER Root
Selects the specification root to inventory. The default is Specs.

.PARAMETER OutputRoot
Selects the output directory, which must resolve below the repository .build directory.

.PARAMETER ConsistencyLedgerPath
Selects the normative consistency ledger relative to the repository root.

.PARAMETER ProtectedDoneBaseCommit
Selects the reviewed Git baseline used to detect unexpected completed-plan changes.

.PARAMETER FailOnFindings
Throws when the generated inventory contains blocking findings.

.PARAMETER PassThru
Returns the complete inventory object after writing reports.

.OUTPUTS
With -PassThru, a PSCustomObject containing files, links, consistency records, and findings. Always writes spec-inventory.json, normative-consistency.json, spec-inventory.md, and test-inventory.json below OutputRoot.

.NOTES
Prerequisites: Git and a repository checkout containing the consistency ledger. Side effects: creates or replaces disposable reports below .build only. Exit is nonzero for invalid output paths, malformed contracts, or blocking findings when -FailOnFindings is set. Primary consumers: specification closeout, CI policy tests, and maintainers reconciling WIP/spec authority.

.EXAMPLE
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
#>
[CmdletBinding()]
param(
    [string]$Root = 'Specs',

    [string]$OutputRoot = '.build/SpecInventory',

    [string]$ConsistencyLedgerPath = 'Specs/NormativeConsistency.json',

    [string]$ProtectedDoneBaseCommit = '6aecfde8e',

    [switch]$FailOnFindings,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $PSScriptRoot 'Modules\Tooling\SpecInformationArchitecture.psm1'
Import-Module $modulePath -Force -ErrorAction Stop

$outputFull = if ([IO.Path]::IsPathRooted($OutputRoot)) {
    [IO.Path]::GetFullPath($OutputRoot)
} else {
    [IO.Path]::GetFullPath((Join-Path $repoRoot $OutputRoot))
}
$repositoryFull = [IO.Path]::GetFullPath($repoRoot).TrimEnd('\', '/')
$outputPrefix = $repositoryFull + [IO.Path]::DirectorySeparatorChar
if (-not $outputFull.StartsWith($outputPrefix, [StringComparison]::OrdinalIgnoreCase) -or
    -not $outputFull.StartsWith((Join-Path $repositoryFull '.build') + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "OutputRoot must resolve beneath the repository .build directory: $OutputRoot"
}

$inventory = Get-RSSpecInventory `
    -RepositoryRoot $repoRoot `
    -Root $Root `
    -ConsistencyLedgerPath $ConsistencyLedgerPath `
    -ProtectedDoneBaseCommit $ProtectedDoneBaseCommit

New-Item -ItemType Directory -Path $outputFull -Force | Out-Null
$inventoryPath = Join-Path $outputFull 'spec-inventory.json'
$consistencyPath = Join-Path $outputFull 'normative-consistency.json'
$reportPath = Join-Path $outputFull 'spec-inventory.md'
$testInventoryPath = Join-Path $outputFull 'test-inventory.json'

$inventory | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $inventoryPath -Encoding utf8
@($inventory.Files | Where-Object { $_.AuthorityClass -in @('Normative', 'NormativeCatalog') } | ForEach-Object {
        [pscustomobject]@{
            path = $_.Path
            status = if ($null -ne $_.Consistency) { $_.Consistency.status } else { 'UNVERIFIED' }
            contractOwner = if ($null -ne $_.Consistency) { $_.Consistency.contractOwner } else { $null }
            implementationAnchors = if ($null -ne $_.Consistency) { @($_.Consistency.implementationAnchors) } else { @() }
            testAnchors = if ($null -ne $_.Consistency) { @($_.Consistency.testAnchors) } else { @() }
            siblingSpecs = if ($null -ne $_.Consistency) { @($_.Consistency.siblingSpecs) } else { @() }
            reviewerNote = if ($null -ne $_.Consistency) { $_.Consistency.reviewerNote } else { $null }
            discrepancyIds = if ($null -ne $_.Consistency) { @($_.Consistency.discrepancyIds) } else { @() }
        }
    }) | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $consistencyPath -Encoding utf8

$report = [System.Collections.Generic.List[string]]::new()
$null = $report.Add('# Specification inventory')
$null = $report.Add('')
$null = $report.Add("- Repository commit: ``$($inventory.RepositoryCommit)``")
$null = $report.Add("- Generated UTC: ``$($inventory.GeneratedAtUtc)``")
$null = $report.Add("- Files: $($inventory.Summary.Files)")
$null = $report.Add("- Normative files: $($inventory.Summary.NormativeFiles)")
$null = $report.Add("- Direct WIP files/index entries: $($inventory.Summary.DirectWipFiles)/$($inventory.Summary.WipIndexEntries)")
$null = $report.Add("- Blocking findings: $($inventory.Summary.BlockingFindings)")
$null = $report.Add('')
$null = $report.Add('## Authority summary')
$null = $report.Add('')
$null = $report.Add('| Class | Count |')
$null = $report.Add('|---|---:|')
foreach ($group in @($inventory.Files | Where-Object IsInScope | Group-Object AuthorityClass | Sort-Object Name)) {
    $null = $report.Add("| $($group.Name) | $($group.Count) |")
}
$null = $report.Add('')
$null = $report.Add('## Normative consistency')
$null = $report.Add('')
$null = $report.Add('| Path | Status | Implementation anchors | Test anchors |')
$null = $report.Add('|---|---|---:|---:|')
foreach ($file in @($inventory.Files | Where-Object { $_.AuthorityClass -in @('Normative', 'NormativeCatalog') } | Sort-Object Path)) {
    $status = if ($null -ne $file.Consistency) { $file.Consistency.status } else { 'UNVERIFIED' }
    $implementationCount = if ($null -ne $file.Consistency) { @($file.Consistency.implementationAnchors).Count } else { 0 }
    $testCount = if ($null -ne $file.Consistency) { @($file.Consistency.testAnchors).Count } else { 0 }
    $null = $report.Add("| ``$($file.Path)`` | ``$status`` | $implementationCount | $testCount |")
}
$null = $report.Add('')
$null = $report.Add('## Findings')
$null = $report.Add('')
if ($inventory.Findings.Count -eq 0) {
    $null = $report.Add('None.')
} else {
    $null = $report.Add('| Severity | Code | Path | Line | Message |')
    $null = $report.Add('|---|---|---|---:|---|')
    foreach ($finding in @($inventory.Findings | Sort-Object Severity, Code, Path, Line)) {
        $message = ([string]$finding.Message).Replace('|', '\|')
        $null = $report.Add("| $($finding.Severity) | ``$($finding.Code)`` | ``$($finding.Path)`` | $($finding.Line) | $message |")
    }
}
$report | Set-Content -LiteralPath $reportPath -Encoding utf8

$getTestInventory = Join-Path $PSScriptRoot 'Get-TestInventory.ps1'
& $getTestInventory -RepoRoot $repoRoot -Format Json | Set-Content -LiteralPath $testInventoryPath -Encoding utf8

Write-Host "Spec inventory: $inventoryPath"
Write-Host "Normative consistency: $consistencyPath"
Write-Host "Markdown report: $reportPath"
Write-Host "Test inventory: $testInventoryPath"
Write-Host "Blocking findings: $($inventory.Summary.BlockingFindings)"

if ($FailOnFindings -and $inventory.Summary.BlockingFindings -ne 0) {
    throw "Specification inventory found $($inventory.Summary.BlockingFindings) blocking issue(s). See $reportPath."
}

if ($PassThru) {
    return $inventory
}
