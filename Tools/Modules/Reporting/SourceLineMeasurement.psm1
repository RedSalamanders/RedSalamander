Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-RSSolutionRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SolutionRoot,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $combined = Join-Path $SolutionRoot $Path
    $fullPath = [IO.Path]::GetFullPath($combined)
    return [IO.Path]::GetRelativePath($RepositoryRoot, $fullPath).Replace('/', '\').TrimEnd('\')
}

function Add-RSSolutionRelativeDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$Set,

        [AllowEmptyString()]
        [string]$RelativePath
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        return
    }

    $trimmed = $RelativePath.Trim().TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($trimmed) -or
        $trimmed -eq '.' -or
        $trimmed.StartsWith('..', [StringComparison]::Ordinal)) {
        return
    }

    [void]$Set.Add($trimmed)
}

function Test-RSSameOrDescendantPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Candidate,

        [Parameter(Mandatory = $true)]
        [string]$Ancestor
    )

    if ($Candidate.Equals($Ancestor, [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    return $Candidate.StartsWith($Ancestor + '\', [StringComparison]::OrdinalIgnoreCase)
}

function Get-RSSolutionSourceDirectories {
    <#
    .SYNOPSIS
        Returns the minimal repository-relative directory set represented by a solution.

    .DESCRIPTION
        Parses project paths and non-root solution-item directories, rejects paths outside
        the repository, then removes descendants already covered by an ancestor target.
        The returned paths are suitable as cloc directory arguments.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SolutionPath,

        [string]$RepositoryRoot
    )

    $resolvedSolutionPath = (Resolve-Path -LiteralPath $SolutionPath -ErrorAction Stop).Path
    $solutionRoot = Split-Path -Path $resolvedSolutionPath -Parent
    if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
        $RepositoryRoot = $solutionRoot
    }
    $resolvedRepositoryRoot = (Resolve-Path -LiteralPath $RepositoryRoot -ErrorAction Stop).Path

    $paths = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)
    $insideSolutionItems = $false

    foreach ($line in (Get-Content -LiteralPath $resolvedSolutionPath -ErrorAction Stop)) {
        if ($line -match '^Project\("[^"]+"\)\s*=\s*"[^"]+",\s*"(?<projectPath>[^"]+)",\s*"\{[^"]+\}"$') {
            $projectPath = $Matches.projectPath
            if (! [string]::IsNullOrWhiteSpace([IO.Path]::GetExtension($projectPath))) {
                $projectDirectory = Split-Path -Path $projectPath -Parent
                $relativeDirectory = ConvertTo-RSSolutionRelativePath `
                    -SolutionRoot $solutionRoot `
                    -RepositoryRoot $resolvedRepositoryRoot `
                    -Path $projectDirectory
                Add-RSSolutionRelativeDirectory -Set $paths -RelativePath $relativeDirectory
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

        if (! $insideSolutionItems) {
            continue
        }

        if ($line -match '^\s*(?<itemPath>[^=]+?)\s*=\s*.+$') {
            $itemDirectory = Split-Path -Path $Matches.itemPath.Trim() -Parent
            if (! [string]::IsNullOrWhiteSpace($itemDirectory) -and $itemDirectory -ne '.') {
                $relativeDirectory = ConvertTo-RSSolutionRelativePath `
                    -SolutionRoot $solutionRoot `
                    -RepositoryRoot $resolvedRepositoryRoot `
                    -Path $itemDirectory
                Add-RSSolutionRelativeDirectory -Set $paths -RelativePath $relativeDirectory
            }
        }
    }

    $sortedPaths = @($paths) | Sort-Object `
        -Property @{ Expression = { $_.Length } }, @{ Expression = { $_ } }
    $filteredPaths = [System.Collections.Generic.List[string]]::new()

    foreach ($candidate in $sortedPaths) {
        $hasAncestor = $false
        foreach ($existing in $filteredPaths) {
            if (Test-RSSameOrDescendantPath -Candidate $candidate -Ancestor $existing) {
                $hasAncestor = $true
                break
            }
        }

        if (! $hasAncestor) {
            $filteredPaths.Add($candidate)
        }
    }

    return @($filteredPaths)
}

Export-ModuleMember -Function 'Get-RSSolutionSourceDirectories'
