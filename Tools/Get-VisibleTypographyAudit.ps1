<#
.SYNOPSIS
    Reports remaining native-font and GDI typography references in visible UI source.

.DESCRIPTION
    Scans the owned UI source roots for native font creation, default GUI fonts,
    hard-coded Segoe UI DirectWrite formats, and GDI text measurement. Results are
    returned as objects by default or grouped Markdown with -AsMarkdown.

.PARAMETER RepoRoot
    Repository root to scan. Defaults to the parent of the Tools directory.

.PARAMETER AsMarkdown
    Emits grouped Markdown tables instead of structured result objects.

.OUTPUTS
    PSCustomObject findings by default or Markdown strings with AsMarkdown.

.NOTES
    Prerequisites: repository UI source. Side effects: none; scanning is read-only. This is an inventory command, so findings are output rather than treated as failure. Primary use is typography migration review.

.EXAMPLE
    .\Tools\Get-VisibleTypographyAudit.ps1

.EXAMPLE
    .\Tools\Get-VisibleTypographyAudit.ps1 -AsMarkdown
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [switch]$AsMarkdown
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Modules\Auditing\RepositorySourceScanner.psm1') -Force

$sourceRoots = @(
    'Common',
    'RedSalamander',
    'Plugins'
)

$patterns = @(
    @{ Name = 'DEFAULT_GUI_FONT'; Pattern = 'DEFAULT_GUI_FONT' },
    @{ Name = 'NONCLIENTMETRICS.lfMenuFont'; Pattern = 'lfMenuFont' },
    @{ Name = 'CreateFontW'; Pattern = 'CreateFontW\(' },
    @{ Name = 'CreateFontIndirectW'; Pattern = 'CreateFontIndirectW\(' },
    @{ Name = 'Hardcoded Segoe UI DWrite'; Pattern = 'CreateTextFormat\([^;\r\n]*L"Segoe UI"' },
    @{ Name = 'Menu/GDI text measurement'; Pattern = 'GetTextExtentPoint32W|GetTextMetricsW' }
)

$results = @(Find-RSRepositorySourceMatches `
    -RepositoryRoot $RepoRoot `
    -SourceRoots $sourceRoots `
    -Extensions @('.cpp', '.h') `
    -Patterns $patterns |
    ForEach-Object {
        [pscustomobject]@{
            Category = $_.PatternName
            Path = $_.Path
            LineNumber = $_.LineNumber
            Line = $_.Line
        }
    })

if ($AsMarkdown)
{
    $grouped = $results | Group-Object Category | Sort-Object Name
    foreach ($group in $grouped)
    {
        Write-Output "## $($group.Name)"
        Write-Output ''
        Write-Output '| Path | Line | Snippet |'
        Write-Output '|------|------|---------|'
        foreach ($item in ($group.Group | Sort-Object Path, LineNumber))
        {
            $snippet = ($item.Line -replace '\|', '\|')
            Write-Output "| ``$($item.Path)`` | $($item.LineNumber) | $snippet |"
        }
        Write-Output ''
    }
}
else
{
    $results | Sort-Object Category, Path, LineNumber
}
