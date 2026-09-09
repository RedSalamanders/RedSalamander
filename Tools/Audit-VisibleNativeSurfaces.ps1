<#
.SYNOPSIS
    Verifies the classified visible native common-control source surface.

.DESCRIPTION
    Scans product source for non-report visible native control identifiers, excluding
    the explicitly test-only or fallback files owned by this command. The command fails
    when the discovered file set differs from its classification table.

.OUTPUTS
    Human-readable classification and status text. No supported pipeline objects.

.NOTES
    Prerequisites: repository product source. Side effects: none; scanning is read-only. Any unclassified or stale visible-native path throws and produces a nonzero exit. Primary consumers: CI and documentation-drift policy tests.

.EXAMPLE
    .\Tools\Audit-VisibleNativeSurfaces.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'Modules\Auditing\RepositorySourceScanner.psm1') -Force

$auditedFiles = @()

$actualFiles = @(Find-RSRepositorySourceMatches `
        -RepositoryRoot $repoRoot `
        -SourceRoots @('RedSalamander', 'Plugins', 'Common') `
        -Extensions @('.cpp', '.h', '.rc') `
        -Patterns @('STATUSCLASSNAMEW|msctls_progress32|SysTreeView32|WC_TREEVIEWW|ToolbarWindow32|TOOLBARCLASSNAME|SysHeader32|WC_HEADER|SysTabControl32|WC_TABCONTROL|msctls_trackbar32|TRACKBAR_CLASSW') `
        -ExcludedRelativePaths @(
            'RedSalamander\Commands.SelfTest.cpp',
            'RedSalamander\Preferences.Dialog.cpp',
            'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ChromeAndPlugins.cpp'
        ) `
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
