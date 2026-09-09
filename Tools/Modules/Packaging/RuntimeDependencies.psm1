# Private packaging module; import through its reviewed callers.
Set-StrictMode -Version Latest

function Resolve-RSRuntimeOutputPath {
    param(
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][ValidateSet('Plugins', 'BuildOutput')][string]$OutputRoot,
        [Parameter(Mandatory = $true)][string]$OutputName
    )

    if ([string]::IsNullOrWhiteSpace($OutputName) -or [IO.Path]::IsPathRooted($OutputName) -or
        $OutputName -match '[:*?]' -or $OutputName -match '[\$%@]\(' -or
        $OutputName -match '(^|[\\/])\.{1,2}([\\/]|$)' -or $OutputName -match '(^|[\\/])([\\/]|$)' -or
        $OutputName.EndsWith('\') -or $OutputName.EndsWith('/')) {
        throw "Runtime dependency OutputName must be a relative path: '$OutputName'"
    }
    $buildOutputFull = [IO.Path]::GetFullPath($BuildOutputDir).TrimEnd('\', '/')
    $selectedRoot = if ($OutputRoot -eq 'Plugins') { Join-Path $buildOutputFull 'Plugins' } else { $buildOutputFull }
    $selectedRoot = [IO.Path]::GetFullPath($selectedRoot).TrimEnd('\', '/')
    $resolved = [IO.Path]::GetFullPath((Join-Path $selectedRoot $OutputName))
    $prefix = $selectedRoot + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Runtime dependency OutputName escapes its $OutputRoot root: '$OutputName'"
    }
    return $resolved
}

function Get-RSRuntimeDependencyManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Debug', 'ASan Debug', 'Release')]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform
    )

    $manifestPath = Join-Path $RepoRoot 'RuntimeDependencies.props'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "Runtime dependency manifest not found: $manifestPath"
    }

    [xml]$document = Get-Content -LiteralPath $manifestPath -Raw
    $flavor = if ($Configuration -eq 'Release') { 'Release' } else { 'Debug' }
    $dependencies = foreach ($node in @($document.Project.ItemGroup.RSRuntimeDependency)) {
        if ([string]$node.Flavor -notin @('Any', $flavor)) { continue }
        if ([string]$node.Platform -notin @('Any', $Platform)) { continue }

        $sourceRootNode = $node.SelectSingleNode('SourceRoot')
        $sourceRoot = if ($null -eq $sourceRootNode -or [string]::IsNullOrWhiteSpace([string]$sourceRootNode.InnerText)) {
            'VcpkgTriplet'
        } else {
            [string]$sourceRootNode.InnerText
        }

        if ($sourceRoot -notin @('VcpkgTriplet', 'RepoRoot')) {
            throw "Runtime dependency '$([string]$node.Include)' has unsupported SourceRoot '$sourceRoot'."
        }

        $outputRootNode = $node.SelectSingleNode('OutputRoot')
        $outputRoot = if ($null -eq $outputRootNode -or [string]::IsNullOrWhiteSpace([string]$outputRootNode.InnerText)) {
            'Plugins'
        } else {
            [string]$outputRootNode.InnerText
        }
        if ($outputRoot -notin @('Plugins', 'BuildOutput')) {
            throw "Runtime dependency '$([string]$node.Include)' has unsupported OutputRoot '$outputRoot'."
        }

        $source = [string]$node.Source
        if ($sourceRoot -eq 'RepoRoot' -and
            ([string]::IsNullOrWhiteSpace($source) -or [IO.Path]::IsPathRooted($source) -or $source -match '[:*?]' -or
             $source -match '[\$%@]\(' -or $source -match '(^|[\\/])\.{1,2}([\\/]|$)' -or $source -match '(^|[\\/])([\\/]|$)')) {
            throw "Runtime dependency '$([string]$node.Include)' has an unsafe RepoRoot Source '$source'."
        }

        $outputName = [string]$node.OutputName
        $null = Resolve-RSRuntimeOutputPath -BuildOutputDir (Join-Path $RepoRoot '.runtime-manifest-validation') `
            -OutputRoot $outputRoot -OutputName $outputName

        [pscustomobject]@{
            Id = [string]$node.Include
            Projects = @(([string]$node.Projects).Split(';', [System.StringSplitOptions]::RemoveEmptyEntries))
            SourceRoot = $sourceRoot
            Source = $source
            OutputRoot = $outputRoot
            OutputName = $outputName
            Required = ([string]$node.Required -ne 'false')
        }
    }

    $remove = foreach ($node in @($document.Project.ItemGroup.RSRuntimeDependencyRemove)) {
        if ([string]$node.Flavor -notin @('Any', $flavor)) { continue }
        if ([string]$node.Platform -notin @('Any', $Platform)) { continue }
        [string]$node.OutputName
    }

    return [pscustomobject]@{
        Path = $manifestPath
        Dependencies = @($dependencies)
        Remove = @($remove | Sort-Object -Unique)
    }
}

function Assert-RSRuntimeDependenciesInOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$BuildOutputDir,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Debug', 'ASan Debug', 'Release')]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform
    )

    $manifest = Get-RSRuntimeDependencyManifest -RepoRoot $RepoRoot -Configuration $Configuration -Platform $Platform
    $pluginsDir = Join-Path $BuildOutputDir 'Plugins'
    if (-not (Test-Path -LiteralPath $pluginsDir -PathType Container)) {
        throw "Plugin output directory not found: $pluginsDir"
    }

    $missing = @($manifest.Dependencies |
        Where-Object Required |
        ForEach-Object { Resolve-RSRuntimeOutputPath -BuildOutputDir $BuildOutputDir -OutputRoot $_.OutputRoot -OutputName $_.OutputName } |
        Sort-Object -Unique |
        Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($missing.Count -gt 0) {
        throw "Required runtime dependencies are missing from the build output: $($missing -join ', ')"
    }

    $stale = @($manifest.Remove |
        ForEach-Object { Resolve-RSRuntimeOutputPath -BuildOutputDir $BuildOutputDir -OutputRoot Plugins -OutputName $_ } |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })
    if ($stale.Count -gt 0) {
        throw "Forbidden stale runtime dependencies remain in the build output: $($stale -join ', ')"
    }

    return $manifest
}

Export-ModuleMember -Function Get-RSRuntimeDependencyManifest, Assert-RSRuntimeDependenciesInOutput
