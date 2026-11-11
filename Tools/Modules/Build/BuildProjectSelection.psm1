# Private build module; import through its reviewed callers.
Set-StrictMode -Version Latest

function Get-RSSolutionProjects {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SolutionPath
    )

    $solutionText = Get-Content -Path $SolutionPath -Raw
    $projectPattern = 'Project\("\{(?<type>[0-9A-Fa-f-]+)\}"\)\s*=\s*"(?<name>[^"]+)",\s*"(?<path>[^"]+)",\s*"\{(?<guid>[0-9A-Fa-f-]+)\}"(?<body>.*?)(?:\r?\n)EndProject'
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
            ProjectType = $match.Groups['type'].Value.ToUpperInvariant()
            SolutionTarget = $null
        }
    }

    $byGuid = @{}
    foreach ($project in $projects) {
        if ($byGuid.ContainsKey($project.Guid)) { throw "Duplicate solution project GUID: $($project.Guid)" }
        $byGuid.Add($project.Guid, $project)
    }
    $folderType = '2150E333-8FDC-42A3-9474-1A3956D46DE8'
    $parents = @{}
    $sections = [regex]::Matches($solutionText, '(?s)GlobalSection\(NestedProjects\)\s*=\s*preSolution\s*\r?\n(?<body>.*?)EndGlobalSection')
    if ($sections.Count -gt 1) { throw 'Duplicate NestedProjects solution sections.' }
    if ($sections.Count) {
        foreach ($line in ($sections[0].Groups['body'].Value -split '\r?\n')) {
            if (-not $line.Trim()) { continue }
            if ($line -notmatch '^\s*\{(?<child>[0-9A-Fa-f-]+)\}\s*=\s*\{(?<parent>[0-9A-Fa-f-]+)\}\s*$') {
                throw 'Malformed NestedProjects solution entry.'
            }
            $child = $Matches.child.ToUpperInvariant(); $parent = $Matches.parent.ToUpperInvariant()
            if (-not $byGuid.ContainsKey($child) -or -not $byGuid.ContainsKey($parent) -or
                $byGuid[$parent].ProjectType -cne $folderType -or $parents.ContainsKey($child)) {
                throw "Unknown, non-folder or duplicate solution nesting: $child -> $parent"
            }
            $parents.Add($child, $parent)
        }
    }
    $targets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($project in $projects) {
        $names = [Collections.Generic.List[string]]::new()
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $current = $project
        while ($true) {
            if (-not $seen.Add($current.Guid)) { throw "Cyclic solution nesting for $($project.Name)." }
            $names.Insert(0, $current.Name)
            if (-not $parents.ContainsKey($current.Guid)) { break }
            $current = $byGuid[$parents[$current.Guid]]
        }
        $target = $names -join '\'
        # Match MSBuild ProjectInSolution's unique target naming, including folder
        # prefixes and reserved entrypoints; filesystem paths are not target names.
        foreach ($character in @('%', '$', '@', ';', '.', '(', ')', "'")) { $target = $target.Replace($character, '_') }
        if ($target -in @('Build', 'Rebuild', 'Clean', 'Publish')) { $target = 'Solution:' + $target }
        if ($project.ProjectType -cne $folderType -and -not $targets.Add($target)) {
            throw "Colliding MSBuild solution target: $target"
        }
        $project.SolutionTarget = $target
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

Export-ModuleMember -Function @(
    'Get-RSSolutionProjects',
    'Get-RSBuildSelection'
)
