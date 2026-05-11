Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$buildScript = Join-Path $repoRoot 'build.ps1'
$outputDir = Join-Path $repoRoot '.build\x64\Debug'
$pluginDir = Join-Path $repoRoot '.build\x64\Debug\Plugins'
$monitorExe = Join-Path $outputDir 'RedSalamanderMonitor.exe'
$searchServiceExe = Join-Path $outputDir 'RedSalamanderSearchService.exe'
$buildLogDir = Join-Path $repoRoot '.build\logs'

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

    try {
        $process = Start-Process `
            -FilePath 'powershell.exe' `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $buildScript, '-ProjectName', 'RedSalamander', '-Configuration', 'Debug', '-Platform', 'x64') `
            -RedirectStandardOutput $stdoutLog `
            -RedirectStandardError $stderrLog `
            -WindowStyle Hidden `
            -PassThru

        if (-not $process.WaitForExit($timeoutSeconds * 1000)) {
            Stop-RSTestProcessTree -ProcessId $process.Id
            throw "Timed out after $timeoutSeconds seconds running targeted RedSalamander build. stdout: $stdoutLog stderr: $stderrLog"
        }

        $process.Refresh()
        $exitCode = if ($null -ne $process.ExitCode) { [int]$process.ExitCode } else { 0 }
        if ($exitCode -ne 0) {
            throw "Targeted RedSalamander build failed with exit code $exitCode. stdout: $stdoutLog stderr: $stderrLog"
        }
    }
    finally {
        if ($null -ne $process) {
            $process.Dispose()
        }
    }
}

Describe 'RedSalamander targeted plugin deployment' -Tag RequiresBuildToolchain {
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
        $expectedPluginLanguageNames = @(
            Get-ChildItem -Path (Join-Path $repoRoot 'Plugins') -Recurse -Filter '*.vcxproj' |
                Where-Object { $_.FullName -match '\\Lang\\' } |
                ForEach-Object { '{0}.dll' -f $_.BaseName } |
                Sort-Object -Unique
        )
        $languageNames = @(Get-ChildItem -Path $langDir -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
        $missingPluginLanguages = @($expectedPluginLanguageNames | Where-Object { $_ -notin $languageNames })

        if ($missingPluginLanguages.Count -gt 0) {
            throw "Missing bundled plugin language resource(s): $($missingPluginLanguages -join ', ')"
        }
    }
}
