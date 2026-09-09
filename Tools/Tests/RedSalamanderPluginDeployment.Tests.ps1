Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$script:selectedPlatform = $null
$script:selectedConfiguration = $null
$buildScript = Join-Path $repoRoot 'build.ps1'
$solutionPath = Join-Path $repoRoot 'RedSalamander.sln'
$script:outputDir = $null
$script:pluginDir = $null
$script:monitorExe = $null
$script:searchServiceExe = $null
$buildLogDir = Join-Path $repoRoot '.build\logs'
$processStreamingModule = Join-Path $repoRoot 'Tools\Modules\Build\ProcessStreaming.psm1'
$artifactOperationLockModule = Join-Path $repoRoot 'Tools\Modules\Build\ArtifactOperationLock.psm1'

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

function Invoke-RSTargetedBuildForDeploymentTest {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform,
        [Parameter(Mandatory = $true)][ValidateSet('Debug', 'Release', 'ASan Debug')][string]$Configuration
    )
    [void](New-Item -ItemType Directory -Path $buildLogDir -Force)

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $stdoutLog = Join-Path $buildLogDir "pester-targeted-plugin-deployment-$timestamp.out.log"
    $stderrLog = Join-Path $buildLogDir "pester-targeted-plugin-deployment-$timestamp.err.log"
    try {
        $exitCode = Invoke-RSStreamingProcess `
            -FilePath 'pwsh.exe' `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $buildScript, '-ProjectName', 'RedSalamander', '-Configuration', $Configuration, '-Platform', $Platform, '-MaxCpuCount', '1') `
            -WorkingDirectory $repoRoot `
            -LogPath $stdoutLog `
            -ErrorLogPath $stderrLog `
            -AdditionalEnvironment @{ RSBuildEnableTests = 'true' } `
            -TimeoutMilliseconds (15 * 60 * 1000) `
            -OutputDrainTimeoutMilliseconds 5000 `
            -DelegateArtifactOperation `
            -OutputLineCallback {
            param(
                [string]$Line,
                [bool]$IsError
            )
            }
    }
    catch {
        $originalException = $_.Exception
        $message = "Targeted RedSalamander build did not complete: $($originalException.Message) stdout: $stdoutLog stderr: $stderrLog"
        throw [System.InvalidOperationException]::new($message, $originalException)
    }

    if ($exitCode -ne 0) {
        throw "Targeted RedSalamander build failed with exit code $exitCode. stdout: $stdoutLog stderr: $stderrLog"
    }
}

Describe 'RedSalamander targeted plugin deployment' -Tag RequiresBuildToolchain {
    BeforeAll {
        $script:selectedPlatform = [Environment]::GetEnvironmentVariable('RS_VALIDATION_PLATFORM', 'Process')
        $script:selectedConfiguration = [Environment]::GetEnvironmentVariable('RS_VALIDATION_CONFIGURATION', 'Process')
        if ($selectedPlatform -notin @('x64', 'ARM64')) {
            throw 'RS_VALIDATION_PLATFORM must explicitly select x64 or ARM64 for artifact-mutating deployment coverage.'
        }
        if ($selectedConfiguration -notin @('Debug', 'Release', 'ASan Debug') -or
            ($selectedPlatform -eq 'ARM64' -and $selectedConfiguration -eq 'ASan Debug')) {
            throw "Unsupported artifact-mutating validation profile: $selectedPlatform|$selectedConfiguration"
        }
        $script:outputDir = Join-Path $repoRoot ('.build\{0}\{1}' -f $selectedPlatform, $selectedConfiguration)
        $script:pluginDir = Join-Path $outputDir 'Plugins'
        $script:monitorExe = Join-Path $outputDir 'RedSalamanderMonitor.exe'
        $script:searchServiceExe = Join-Path $outputDir 'RedSalamanderSearchService.exe'

        if (-not (Test-Path -LiteralPath $artifactOperationLockModule)) {
            throw "Artifact operation lock helper not found: $artifactOperationLockModule"
        }
        # The contained-process module resolves delegation hooks from the global
        # command table, so the lock module must be visible before the launch.
        Import-Module $artifactOperationLockModule -Force -Global -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $processStreamingModule)) {
            throw "Process streaming helper not found: $processStreamingModule"
        }
        Import-Module $processStreamingModule -Force -ErrorAction Stop

        $script:deploymentArtifactOperationScope = @{
            kind = 'build'
            target = 'RedSalamander-targeted-plugin-deployment'
            configuration = $selectedConfiguration
            platform = $selectedPlatform
        }
        $script:deploymentArtifactOperationLock = Enter-RSArtifactOperationLock `
            -RepoRoot $repoRoot `
            -Operation "Pester targeted plugin deployment $selectedConfiguration|$selectedPlatform" `
            -Scope $script:deploymentArtifactOperationScope
        if ($script:deploymentArtifactOperationLock.WasAbandoned) {
            [void](Set-RSArtifactOperationContaminated `
                    -RepoRoot $repoRoot `
                    -Reason 'The previous build/test owner exited without clearing the exclusive artifact-operation lock.' `
                    -AbandonedOwner $script:deploymentArtifactOperationLock.AbandonedOwner `
                    -Scope $script:deploymentArtifactOperationScope)
        }
        if (Test-RSArtifactOperationContaminated `
                -RepoRoot $repoRoot `
                -Scope $script:deploymentArtifactOperationScope) {
            throw "Targeted deployment testing cannot use contaminated shared artifacts. Run a matching full-solution build.ps1 -Rebuild first."
        }
        Assert-RSNoResidualArtifactToolProcesses `
            -RepoRoot $repoRoot `
            -Scope $script:deploymentArtifactOperationScope
    }

    AfterAll {
        Exit-RSArtifactOperationLock -Lock $script:deploymentArtifactOperationLock
        $script:deploymentArtifactOperationLock = $null
        $script:deploymentArtifactOperationScope = $null
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

        Invoke-RSTargetedBuildForDeploymentTest -Platform $selectedPlatform -Configuration $selectedConfiguration

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
