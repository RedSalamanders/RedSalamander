param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [switch]$AsMarkdown
)

$ErrorActionPreference = 'Stop'

$sourceRoots = @(
    'Common/DxUi',
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

function Get-RelativeRepoPath
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $baseFull   = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $BasePath).Path)
    $targetFull = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $TargetPath).Path)

    if (! $baseFull.EndsWith([IO.Path]::DirectorySeparatorChar))
    {
        $baseFull += [IO.Path]::DirectorySeparatorChar
    }

    $baseUri     = [Uri]$baseFull
    $targetUri   = [Uri]$targetFull
    $relativeUri = $baseUri.MakeRelativeUri($targetUri)

    return [Uri]::UnescapeDataString($relativeUri.ToString()).Replace('/', '\')
}

$files = foreach ($root in $sourceRoots)
{
    Get-ChildItem (Join-Path $RepoRoot $root) -Recurse -Include *.cpp,*.h -File
}

$results = foreach ($pattern in $patterns)
{
    $matches = $files | Select-String -Pattern $pattern.Pattern
    foreach ($match in $matches)
    {
        [PSCustomObject]@{
            Category   = $pattern.Name
            Path       = Get-RelativeRepoPath -BasePath $RepoRoot -TargetPath $match.Path
            LineNumber = $match.LineNumber
            Line       = $match.Line.Trim()
        }
    }
}

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
