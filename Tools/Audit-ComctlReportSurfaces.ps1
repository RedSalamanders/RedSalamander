<#
.SYNOPSIS
    Verifies the classified list-view and tooltip common-control source surface.

.DESCRIPTION
    Scans the product source roots for report-surface common-control identifiers and
    fails when the discovered file set differs from the command-owned classification.
    A successful run prints the classification table and audited roots.

.OUTPUTS
    Human-readable classification and status text. No supported pipeline objects.

.NOTES
    Prerequisites: repository product source. Side effects: none; scanning is read-only. Any unclassified or stale report-surface path throws and produces a nonzero exit. Primary consumers: CI and documentation-drift policy tests.

.EXAMPLE
    .\Tools\Audit-ComctlReportSurfaces.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'Modules\Auditing\RepositorySourceScanner.psm1') -Force

$auditedFiles = @()

$sourceRoots = @(
    "Common",
    "Plugins",
    "RedSalamander",
    "RedSalamanderMonitor",
    "RedConfigure"
)

$sourceExtensions = @(".cpp", ".h", ".rc", ".rc2", ".vcxproj", ".props", ".targets", ".manifest")
$reportSurfacePattern = "SysListView32|WC_LISTVIEWW|TOOLTIPS_CLASSW|TOOLINFOW|NMTTDISPINFOW|TTM_|TTN_|TTS_|TTF_|ListView_|LVS_|LVN_|LVIF_|LVIS_|LVNI_"

$actualFiles = @(Find-RSRepositorySourceMatches `
        -RepositoryRoot $repoRoot `
        -SourceRoots $sourceRoots `
        -Extensions $sourceExtensions `
        -Patterns @($reportSurfacePattern) `
        -PathSeparator Slash |
    Select-Object -ExpandProperty Path -Unique |
    Sort-Object -Unique)

$expectedFiles = $auditedFiles.File | Sort-Object -Unique
$missingFiles = @($expectedFiles | Where-Object { $_ -notin $actualFiles })
$unexpectedFiles = @($actualFiles | Where-Object { $_ -notin $expectedFiles })

if ($missingFiles.Count -gt 0 -or $unexpectedFiles.Count -gt 0)
{
    if ($missingFiles.Count -gt 0)
    {
        Write-Host "Missing audited comctl report-surface files:" -ForegroundColor Red
        $missingFiles | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    }

    if ($unexpectedFiles.Count -gt 0)
    {
        Write-Host "Unexpected comctl report-surface files requiring classification:" -ForegroundColor Red
        $unexpectedFiles | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    }

    throw "Comctl report-surface audit is out of date."
}

Write-Host "# Comctl Report-Surface Audit" -ForegroundColor Green
Write-Host ""
Write-Host "| File | Surface | Bucket | Current Path | Notes |"
Write-Host "| --- | --- | --- | --- | --- |"

foreach ($row in ($auditedFiles | Sort-Object File, Surface))
{
    Write-Host "| $($row.File) | $($row.Surface) | $($row.Bucket) | $($row.CurrentPath) | $($row.Notes) |"
}

Write-Host ""
Write-Host ""
Write-Host "Audited roots: $($sourceRoots -join ', ')" -ForegroundColor Green
Write-Host "Audit passed: every remaining visible listview/tooltip comctl report-surface reference is classified." -ForegroundColor Green
