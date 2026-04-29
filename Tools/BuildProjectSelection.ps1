Set-StrictMode -Version Latest

function Get-RSSolutionProjects {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SolutionPath
    )

    $solutionText = Get-Content -Path $SolutionPath -Raw
    $projectPattern = 'Project\("\{[0-9A-Fa-f-]+\}"\)\s*=\s*"(?<name>[^"]+)",\s*"(?<path>[^"]+)",\s*"\{(?<guid>[0-9A-Fa-f-]+)\}"(?<body>.*?)(?:\r?\n)EndProject'
    $matches = [Regex]::Matches($solutionText, $projectPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    $projects = @()
    foreach ($match in $matches) {
        $relativePath = $match.Groups['path'].Value
        $fullPath = Join-Path -Path (Split-Path -Parent $SolutionPath) -ChildPath $relativePath
        $resolvedPath = $null
        if (Test-Path $fullPath) {
            $resolvedPath = (Resolve-Path $fullPath).Path
        }

        $projects += [pscustomobject]@{
            Name = $match.Groups['name'].Value
            Guid = $match.Groups['guid'].Value.ToUpperInvariant()
            RelativePath = $relativePath
            Path = $resolvedPath
            Body = $match.Groups['body'].Value
        }
    }

    return $projects
}

function Resolve-ProjectFileFromSolution {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SolutionPath,

        [Parameter(Mandatory = $true)]
        [string]$ProjectName
    )

    $project = Get-RSSolutionProjects -SolutionPath $SolutionPath | Where-Object { $_.Name -eq $ProjectName } | Select-Object -First 1
    if (-not $project) {
        throw "Project '$ProjectName' not found in solution '$SolutionPath'."
    }

    if (-not $project.Path) {
        throw "Project file not found: $($project.RelativePath)"
    }

    return $project.Path
}

function Get-RSBuildSelection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SolutionPath,

        [Parameter(Mandatory = $true)]
        [string]$SolutionDir,

        [Parameter(Mandatory = $true)]
        [string]$ProjectName,

        [switch]$Rebuild
    )

    $resolvedProjectPath = Resolve-ProjectFileFromSolution -SolutionPath $SolutionPath -ProjectName $ProjectName
    $mainAppProjectPath = (Resolve-Path (Join-Path $SolutionDir 'RedSalamander\RedSalamander.vcxproj')).Path
    $buildProjectDirectly = -not [System.String]::Equals(
        $resolvedProjectPath,
        $mainAppProjectPath,
        [System.StringComparison]::OrdinalIgnoreCase)

    $buildInput = if ($buildProjectDirectly) { $resolvedProjectPath } else { $SolutionPath }
    $msbuildTarget = if ($buildProjectDirectly) {
        if ($Rebuild) { 'Rebuild' } else { 'Build' }
    } elseif ($Rebuild) {
        '{0}:Rebuild' -f $ProjectName
    } else {
        $ProjectName
    }

    $cleanTarget = if ($buildProjectDirectly) { 'Clean' } else { '{0}:Clean' -f $ProjectName }

    return [pscustomobject]@{
        ResolvedProjectPath = $resolvedProjectPath
        BuildProjectDirectly = $buildProjectDirectly
        BuildInput = $buildInput
        MSBuildTarget = $msbuildTarget
        CleanTarget = $cleanTarget
    }
}
