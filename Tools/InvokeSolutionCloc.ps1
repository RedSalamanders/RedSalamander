<#
.SYNOPSIS
    Runs cloc against directories that are part of RedSalamander.sln.

.DESCRIPTION
    This script parses the solution file and invokes cloc only for project directories and
    non-root shared source directories referenced by solution items. That keeps cloc from
    scanning unrelated folders such as Docs, vcpkg_installed, build outputs, and other
    repo content that is not part of the solution codebase.

.PARAMETER SolutionPath
    Path to the solution file. Defaults to ..\RedSalamander.sln relative to this script.

.PARAMETER ClocPath
    cloc executable name or full path. Defaults to 'cloc'.

.PARAMETER ClocArgs
    Extra arguments passed through to cloc.

.EXAMPLE
    .\Tools\InvokeSolutionCloc.ps1

.EXAMPLE
    .\Tools\InvokeSolutionCloc.ps1 --by-file --xml
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SolutionPath = (Join-Path $PSScriptRoot '..\RedSalamander.sln'),

    [Parameter(Mandatory = $false)]
    [string]$ClocPath = 'cloc',

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ClocArgs
)

$ErrorActionPreference = 'Stop'

function Resolve-ExistingPath([string]$PathText) {
    $resolved = Resolve-Path -LiteralPath $PathText -ErrorAction Stop
    return $resolved.Path
}

function Normalize-RelativePath([string]$PathText) {
    $combined = Join-Path $script:SolutionRoot $PathText
    $fullPath = [System.IO.Path]::GetFullPath($combined)
    $relative = [System.IO.Path]::GetRelativePath($script:RepoRoot, $fullPath)
    return $relative.Replace('/', '\').TrimEnd('\')
}

function Add-RelativeDirectory([System.Collections.Generic.HashSet[string]]$Set, [string]$RelativePath) {
    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        return
    }

    $trimmed = $RelativePath.Trim().TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed -eq '.') {
        return
    }

    if ($trimmed.StartsWith('..', [System.StringComparison]::Ordinal)) {
        return
    }

    [void]$Set.Add($trimmed)
}

function Test-IsSameOrDescendant([string]$Candidate, [string]$Ancestor) {
    if ($Candidate.Equals($Ancestor, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    return $Candidate.StartsWith($Ancestor + '\', [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-SolutionDirectories {
    $paths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $lines = Get-Content -LiteralPath $script:ResolvedSolutionPath -ErrorAction Stop
    $insideSolutionItems = $false

    foreach ($line in $lines) {
        if ($line -match '^Project\("[^"]+"\)\s*=\s*"[^"]+",\s*"(?<projectPath>[^"]+)",\s*"\{[^"]+\}"$') {
            $projectPath = $Matches.projectPath
            if (-not [string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($projectPath))) {
                $projectDirectory = Split-Path -Path $projectPath -Parent
                Add-RelativeDirectory -Set $paths -RelativePath (Normalize-RelativePath $projectDirectory)
            }
        }

        if ($line -match '^\s*ProjectSection\(SolutionItems\)\s*=\s*preProject\s*$') {
            $insideSolutionItems = $true
            continue
        }

        if ($insideSolutionItems -and $line -match '^\s*EndProjectSection\s*$') {
            $insideSolutionItems = $false
            continue
        }

        if (-not $insideSolutionItems) {
            continue
        }

        if ($line -match '^\s*(?<itemPath>[^=]+?)\s*=\s*.+$') {
            $itemPath = $Matches.itemPath.Trim()
            $itemDirectory = Split-Path -Path $itemPath -Parent
            if (-not [string]::IsNullOrWhiteSpace($itemDirectory) -and $itemDirectory -ne '.') {
                Add-RelativeDirectory -Set $paths -RelativePath (Normalize-RelativePath $itemDirectory)
            }
        }
    }

    $sortedPaths = @($paths) | Sort-Object -Property @{ Expression = { $_.Length } }, @{ Expression = { $_ } }
    $filteredPaths = [System.Collections.Generic.List[string]]::new()

    foreach ($candidate in $sortedPaths) {
        $hasAncestor = $false
        foreach ($existing in $filteredPaths) {
            if (Test-IsSameOrDescendant -Candidate $candidate -Ancestor $existing) {
                $hasAncestor = $true
                break
            }
        }

        if (-not $hasAncestor) {
            [void]$filteredPaths.Add($candidate)
        }
    }

    return $filteredPaths
}

$script:ResolvedSolutionPath = Resolve-ExistingPath $SolutionPath
$script:SolutionRoot = Split-Path -Path $script:ResolvedSolutionPath -Parent
$script:RepoRoot = $script:SolutionRoot

$clocCommand = Get-Command -Name $ClocPath -ErrorAction SilentlyContinue
if ($null -eq $clocCommand) {
    throw "Unable to find cloc executable '$ClocPath'. Install cloc or pass -ClocPath with the full executable path."
}

$solutionDirectories = Get-SolutionDirectories
if ($solutionDirectories.Count -eq 0) {
    throw "No solution-backed directories were found in $script:ResolvedSolutionPath."
}

$clocTargets = @()
foreach ($relativeDirectory in $solutionDirectories) {
    $clocTargets += (Join-Path $script:RepoRoot $relativeDirectory)
}

Write-Host 'Running cloc on solution-backed directories:' -ForegroundColor Cyan
foreach ($target in $solutionDirectories) {
    Write-Host "  $target"
}

& $clocCommand.Source @ClocArgs @clocTargets
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}