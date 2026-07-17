Set-StrictMode -Version Latest

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

        [pscustomobject]@{
            Id = [string]$node.Include
            Projects = @(([string]$node.Projects).Split(';', [System.StringSplitOptions]::RemoveEmptyEntries))
            Source = [string]$node.Source
            OutputName = [string]$node.OutputName
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
        ForEach-Object { Join-Path $pluginsDir $_.OutputName } |
        Sort-Object -Unique |
        Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($missing.Count -gt 0) {
        throw "Required runtime dependencies are missing from the build output: $($missing -join ', ')"
    }

    $stale = @($manifest.Remove |
        ForEach-Object { Join-Path $pluginsDir $_ } |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })
    if ($stale.Count -gt 0) {
        throw "Forbidden stale runtime dependencies remain in the build output: $($stale -join ', ')"
    }

    return $manifest
}

Export-ModuleMember -Function Get-RSRuntimeDependencyManifest, Assert-RSRuntimeDependenciesInOutput
