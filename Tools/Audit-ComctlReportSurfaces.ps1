param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

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

$actualFiles = @(
    foreach ($root in $sourceRoots)
    {
        Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object {
            $_.Extension -in $sourceExtensions
        } | Select-String -Pattern $reportSurfacePattern | Select-Object -ExpandProperty Path -Unique
    }
) | ForEach-Object { (Resolve-Path -LiteralPath $_ -Relative).TrimStart('.', '\', '/').Replace('\', '/') } | Sort-Object -Unique

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
