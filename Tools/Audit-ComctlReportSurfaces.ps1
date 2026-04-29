param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$auditedFiles = @()

$actualFiles = @(
    foreach ($root in @("RedSalamander", "Plugins"))
    {
        Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object {
            $_.Extension -in @(".cpp", ".h", ".rc")
        } | Select-String -Pattern "WC_LISTVIEWW|SysListView32" | Select-Object -ExpandProperty Path -Unique
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
Write-Host "Audit passed: every remaining SysListView32/WC_LISTVIEWW reference is classified." -ForegroundColor Green
