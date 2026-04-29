param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$auditedFiles = @()

$actualFiles = @(
    foreach ($root in @("RedSalamander", "Plugins", "Common"))
    {
        Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object {
            $_.Extension -in @(".cpp", ".h", ".rc")
        } | Select-String -Pattern "STATUSCLASSNAMEW|msctls_progress32|SysTreeView32|WC_TREEVIEWW|ToolbarWindow32|TOOLBARCLASSNAME|SysHeader32|WC_HEADER|SysTabControl32|WC_TABCONTROL|msctls_trackbar32|TRACKBAR_CLASSW" |
            Select-Object -ExpandProperty Path -Unique
    }
) | ForEach-Object { (Resolve-Path -LiteralPath $_ -Relative).TrimStart('.', '\', '/').Replace('\', '/') } | Where-Object {
    $_ -ne "RedSalamander/Commands.SelfTest.cpp" -and
        $_ -ne "RedSalamander/Preferences.Dialog.cpp" -and
        $_ -ne "RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp"
} | Sort-Object -Unique

$expectedFiles = $auditedFiles.File | Sort-Object -Unique
$missingFiles = @($expectedFiles | Where-Object { $_ -notin $actualFiles })
$unexpectedFiles = @($actualFiles | Where-Object { $_ -notin $expectedFiles })

if ($missingFiles.Count -gt 0 -or $unexpectedFiles.Count -gt 0)
{
    if ($missingFiles.Count -gt 0)
    {
        Write-Host "Missing audited visible-native files:" -ForegroundColor Red
        $missingFiles | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    }

    if ($unexpectedFiles.Count -gt 0)
    {
        Write-Host "Unexpected visible-native files requiring classification:" -ForegroundColor Red
        $unexpectedFiles | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    }

    throw "Visible-native audit is out of date."
}

Write-Host "# Visible Native Surface Audit" -ForegroundColor Green
Write-Host ""
Write-Host "| File | Surface | Native Class | Bucket | Notes |"
Write-Host "| --- | --- | --- | --- | --- |"

foreach ($row in ($auditedFiles | Sort-Object File, Surface))
{
    Write-Host "| $($row.File) | $($row.Surface) | $($row.NativeClass) | $($row.Bucket) | $($row.Notes) |"
}

Write-Host ""
Write-Host "Audit passed: every remaining non-report visible-native surface reference is classified." -ForegroundColor Green
