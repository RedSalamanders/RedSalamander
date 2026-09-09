Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-RSRepositoryRelativePath {
    <#
    .SYNOPSIS
        Converts an existing path to a repository-relative path.

    .DESCRIPTION
        Resolves both paths before computing the relative path and returns a stable
        separator style for reports and source-contract comparisons.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [ValidateSet('Backslash', 'Slash')]
        [string]$PathSeparator = 'Backslash'
    )

    $baseFull = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $RepositoryRoot).Path)
    $targetFull = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Path).Path)

    if (! $baseFull.EndsWith([IO.Path]::DirectorySeparatorChar)) {
        $baseFull += [IO.Path]::DirectorySeparatorChar
    }

    $baseUri = [Uri]$baseFull
    $targetUri = [Uri]$targetFull
    $relative = [Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString())

    if ($PathSeparator -eq 'Slash') {
        return $relative.Replace('\', '/')
    }

    return $relative.Replace('/', '\')
}

function Get-RSRepositorySourceFiles {
    <#
    .SYNOPSIS
        Enumerates source files from selected repository roots.

    .DESCRIPTION
        Centralizes recursive source discovery, extension filtering, excluded tree
        filtering, and stable ordering for repository audit commands.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$SourceRoots,

        [Parameter(Mandatory = $true)]
        [string[]]$Extensions,

        [string[]]$ExcludedPathFragments = @(),

        [string[]]$ExcludedRelativePaths = @()
    )

    $resolvedRoot = (Resolve-Path -LiteralPath $RepositoryRoot).Path
    $extensionSet = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($extension in $Extensions) {
        $normalizedExtension = $extension.Trim().TrimStart('*')
        if (! $normalizedExtension.StartsWith('.')) {
            $normalizedExtension = ".$normalizedExtension"
        }
        [void]$extensionSet.Add($normalizedExtension)
    }

    $excludedRelativeSet = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($relativePath in $ExcludedRelativePaths) {
        [void]$excludedRelativeSet.Add($relativePath.Replace('/', '\').TrimStart('\'))
    }

    $files = [System.Collections.Generic.List[IO.FileInfo]]::new()
    foreach ($sourceRoot in $SourceRoots) {
        $rootPath = Join-Path $resolvedRoot $sourceRoot
        if (! (Test-Path -LiteralPath $rootPath -PathType Container)) {
            continue
        }

        foreach ($file in (Get-ChildItem -LiteralPath $rootPath -Recurse -File)) {
            if (! $extensionSet.Contains($file.Extension)) {
                continue
            }

            $excluded = $false
            $fullPath = [IO.Path]::GetFullPath($file.FullName)
            foreach ($fragment in $ExcludedPathFragments) {
                if ($fullPath.IndexOf($fragment, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $excluded = $true
                    break
                }
            }
            if ($excluded) {
                continue
            }

            if ($excludedRelativeSet.Count -gt 0) {
                $relativePath = ConvertTo-RSRepositoryRelativePath `
                    -RepositoryRoot $resolvedRoot `
                    -Path $file.FullName
                if ($excludedRelativeSet.Contains($relativePath)) {
                    continue
                }
            }

            $files.Add($file)
        }
    }

    return @($files | Sort-Object FullName -Unique)
}

function Find-RSRepositorySourceMatches {
    <#
    .SYNOPSIS
        Finds regex matches in selected repository source files.

    .DESCRIPTION
        Combines the shared file discovery policy with Select-String and emits a
        normalized record for each source match. Pattern definitions may be strings
        or hashtables with Pattern plus either Name or Category metadata.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$SourceRoots,

        [Parameter(Mandatory = $true)]
        [string[]]$Extensions,

        [Parameter(Mandatory = $true)]
        [object[]]$Patterns,

        [string[]]$ExcludedPathFragments = @(),

        [string[]]$ExcludedRelativePaths = @(),

        [ValidateSet('Backslash', 'Slash')]
        [string]$PathSeparator = 'Backslash'
    )

    $files = @(Get-RSRepositorySourceFiles `
        -RepositoryRoot $RepositoryRoot `
        -SourceRoots $SourceRoots `
        -Extensions $Extensions `
        -ExcludedPathFragments $ExcludedPathFragments `
        -ExcludedRelativePaths $ExcludedRelativePaths)

    foreach ($definition in $Patterns) {
        $pattern = $null
        $patternName = ''
        if ($definition -is [string]) {
            $pattern = [string]$definition
        }
        elseif ($definition -is [System.Collections.IDictionary]) {
            $pattern = [string]$definition['Pattern']
            if ($definition.Contains('Category')) {
                $patternName = [string]$definition['Category']
            }
            elseif ($definition.Contains('Name')) {
                $patternName = [string]$definition['Name']
            }
        }
        else {
            $pattern = [string]$definition.Pattern
            if ($definition.PSObject.Properties.Match('Category').Count -gt 0) {
                $patternName = [string]$definition.Category
            }
            elseif ($definition.PSObject.Properties.Match('Name').Count -gt 0) {
                $patternName = [string]$definition.Name
            }
        }

        if ([string]::IsNullOrWhiteSpace($pattern)) {
            throw 'Repository source scan pattern definitions must provide a non-empty Pattern.'
        }

        foreach ($match in @($files | Select-String -Pattern $pattern)) {
            [pscustomobject]@{
                PatternName = $patternName
                Path = ConvertTo-RSRepositoryRelativePath `
                    -RepositoryRoot $RepositoryRoot `
                    -Path $match.Path `
                    -PathSeparator $PathSeparator
                LineNumber = $match.LineNumber
                Line = $match.Line.Trim()
            }
        }
    }
}

Export-ModuleMember -Function @(
    'ConvertTo-RSRepositoryRelativePath',
    'Get-RSRepositorySourceFiles',
    'Find-RSRepositorySourceMatches'
)
