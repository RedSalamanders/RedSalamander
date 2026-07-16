Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$buildScript = Join-Path $repoRoot 'build.ps1'
$solutionPath = Join-Path $repoRoot 'RedSalamander.sln'
$outputDir = Join-Path $repoRoot '.build\x64\Debug'
$pluginDir = Join-Path $repoRoot '.build\x64\Debug\Plugins'
$monitorExe = Join-Path $outputDir 'RedSalamanderMonitor.exe'
$searchServiceExe = Join-Path $outputDir 'RedSalamanderSearchService.exe'
$buildLogDir = Join-Path $repoRoot '.build\logs'
$sanitizedEnvironmentScript = Join-Path $repoRoot 'Tools\SanitizedEnvironment.ps1'
$artifactOperationLockScript = Join-Path $repoRoot 'Tools\ArtifactOperationLock.ps1'
if (-not (Test-Path -LiteralPath $sanitizedEnvironmentScript)) {
    throw "Sanitized environment helper not found: $sanitizedEnvironmentScript"
}
. $sanitizedEnvironmentScript
if (-not (Test-Path -LiteralPath $artifactOperationLockScript)) {
    throw "Artifact operation lock helper not found: $artifactOperationLockScript"
}
. $artifactOperationLockScript

function Get-RSSolutionProjects {
    $projects = @()
    foreach ($line in (Get-Content -LiteralPath $solutionPath)) {
        $match = [regex]::Match($line, '^Project\("\{[^"]+\}"\)\s*=\s*"(?<name>[^"]+)",\s*"(?<path>[^"]+)",\s*"\{(?<guid>[^}]+)\}"')
        if (-not $match.Success) {
            continue
        }

        $projects += [pscustomobject]@{
            Name = $match.Groups['name'].Value
            RelativePath = $match.Groups['path'].Value
            Guid = $match.Groups['guid'].Value.ToUpperInvariant()
        }
    }

    return $projects
}

function Get-RSRedSalamanderDependencyGuids {
    $dependencies = @()
    $insideRedSalamander = $false
    $insideDependencies = $false

    foreach ($line in (Get-Content -LiteralPath $solutionPath)) {
        $projectMatch = [regex]::Match($line, '^Project\("\{[^"]+\}"\)\s*=\s*"(?<name>[^"]+)",\s*"(?<path>[^"]+)",\s*"\{(?<guid>[^}]+)\}"')
        if ($projectMatch.Success) {
            $insideRedSalamander = ($projectMatch.Groups['name'].Value -eq 'RedSalamander')
            $insideDependencies = $false
            continue
        }

        if (-not $insideRedSalamander) {
            continue
        }

        if ($line -match '^\s*ProjectSection\(ProjectDependencies\)\s*=\s*postProject') {
            $insideDependencies = $true
            continue
        }

        if ($insideDependencies -and $line -match '^\s*EndProjectSection\b') {
            return $dependencies
        }

        if ($line -match '^EndProject\b') {
            return $dependencies
        }

        if ($insideDependencies) {
            $dependencyMatch = [regex]::Match($line, '\{(?<guid>[0-9A-Fa-f-]+)\}')
            if ($dependencyMatch.Success) {
                $dependencies += $dependencyMatch.Groups['guid'].Value.ToUpperInvariant()
            }
        }
    }

    return $dependencies
}

function Stop-RSTestProcessTree {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ProcessId
    )

    $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId = $ProcessId" -ErrorAction SilentlyContinue)
    foreach ($child in $children) {
        Stop-RSTestProcessTree -ProcessId ([int]$child.ProcessId)
    }

    Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
}

function Invoke-RSTargetedBuildForDeploymentTest {
    [void](New-Item -ItemType Directory -Path $buildLogDir -Force)

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $stdoutLog = Join-Path $buildLogDir "pester-targeted-plugin-deployment-$timestamp.out.log"
    $stderrLog = Join-Path $buildLogDir "pester-targeted-plugin-deployment-$timestamp.err.log"
    $timeoutSeconds = 900
    $process = $null
    $stdoutTask = $null
    $stderrTask = $null
    $previousEnableTests = [Environment]::GetEnvironmentVariable('RSBuildEnableTests', 'Process')

    try {
        [Environment]::SetEnvironmentVariable('RSBuildEnableTests', 'true', 'Process')
        $psi = New-RSProcessStartInfo `
            -FilePath 'powershell.exe' `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $buildScript, '-ProjectName', 'RedSalamander', '-Configuration', 'Debug', '-Platform', 'x64') `
            -WorkingDirectory $repoRoot
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $process = Start-RSContainedProcess `
            -ProcessStartInfo $psi `
            -DelegateArtifactOperation

        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()

        if (-not $process.WaitForExit($timeoutSeconds * 1000)) {
            Close-RSContainedProcess -Process $process
            $process = $null
            throw "Timed out after $timeoutSeconds seconds running targeted RedSalamander build. stdout: $stdoutLog stderr: $stderrLog"
        }
        $process.WaitForExit()

        [System.IO.File]::WriteAllText($stdoutLog, $stdoutTask.Result)
        [System.IO.File]::WriteAllText($stderrLog, $stderrTask.Result)

        $process.Refresh()
        $exitCode = if ($null -ne $process.ExitCode) { [int]$process.ExitCode } else { 0 }
        if ($exitCode -ne 0) {
            throw "Targeted RedSalamander build failed with exit code $exitCode. stdout: $stdoutLog stderr: $stderrLog"
        }
    }
    finally {
        [Environment]::SetEnvironmentVariable('RSBuildEnableTests', $previousEnableTests, 'Process')
        if ($null -ne $process) {
            Close-RSContainedProcess -Process $process
        }
    }
}

Describe 'RedSalamander targeted plugin deployment' -Tag RequiresBuildToolchain {
    BeforeAll {
        $script:deploymentArtifactOperationLock = Enter-RSArtifactOperationLock `
            -RepoRoot $repoRoot `
            -Operation 'Pester targeted plugin deployment Debug|x64' `
            -Scope @{
                kind = 'build'
                target = 'RedSalamander-targeted-plugin-deployment'
                configuration = 'Debug'
                platform = 'x64'
            }
        if ($script:deploymentArtifactOperationLock.WasAbandoned) {
            [void](Set-RSArtifactOperationContaminated `
                    -RepoRoot $repoRoot `
                    -Reason 'The previous build/test owner exited without clearing the exclusive artifact-operation lock.' `
                    -AbandonedOwner $script:deploymentArtifactOperationLock.AbandonedOwner)
        }
        if (Test-RSArtifactOperationContaminated -RepoRoot $repoRoot) {
            throw "Targeted deployment testing cannot use contaminated shared artifacts. Run a matching full-solution build.ps1 -Rebuild first."
        }
        Assert-RSNoResidualArtifactToolProcesses -RepoRoot $repoRoot
    }

    AfterAll {
        Exit-RSArtifactOperationLock -Lock $script:deploymentArtifactOperationLock
        $script:deploymentArtifactOperationLock = $null
    }

    It 'repopulates bundled sibling binaries when build.ps1 targets RedSalamander' {
        if (Test-Path $pluginDir) {
            Remove-Item -LiteralPath $pluginDir -Recurse -Force
        }

        foreach ($path in @($monitorExe, $searchServiceExe)) {
            if (Test-Path $path) {
                Remove-Item -LiteralPath $path -Force
            }
        }

        New-Item -ItemType Directory -Path $pluginDir -Force | Out-Null

        Invoke-RSTargetedBuildForDeploymentTest

        foreach ($path in @($monitorExe, $searchServiceExe)) {
            (Test-Path $path) | Should Be $true
        }

        $expectedPluginNames = @(
            Get-ChildItem -Path (Join-Path $repoRoot 'Plugins') -Recurse -Filter '*.vcxproj' |
                Where-Object { $_.FullName -notmatch '\\Lang\\' } |
                ForEach-Object { '{0}.dll' -f $_.BaseName } |
                Sort-Object -Unique
        )
        $pluginNames = @(Get-ChildItem -Path $pluginDir -File | Select-Object -ExpandProperty Name)
        $missingPlugins = @($expectedPluginNames | Where-Object { $_ -notin $pluginNames })

        if ($missingPlugins.Count -gt 0) {
            throw "Missing bundled plugin(s): $($missingPlugins -join ', ')"
        }

        $langDir = Join-Path $outputDir 'Lang'
        $solutionProjects = @(Get-RSSolutionProjects)
        $redSalamanderDependencyGuids = @(Get-RSRedSalamanderDependencyGuids)
        $expectedPluginLanguageNames = @(
            $solutionProjects |
                Where-Object { $_.Guid -in $redSalamanderDependencyGuids -and $_.RelativePath -match '^Plugins\\.*\\Lang\\.*\.vcxproj$' } |
                ForEach-Object { '{0}.dll' -f [System.IO.Path]::GetFileNameWithoutExtension($_.RelativePath) } |
                Sort-Object -Unique
        )
        $languageNames = @(Get-ChildItem -Path $langDir -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
        $missingPluginLanguages = @($expectedPluginLanguageNames | Where-Object { $_ -notin $languageNames })

        if ($missingPluginLanguages.Count -gt 0) {
            throw "Missing bundled plugin language resource(s): $($missingPluginLanguages -join ', ')"
        }
    }
}
