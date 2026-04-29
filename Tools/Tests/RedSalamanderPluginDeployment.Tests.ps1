Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$buildScript = Join-Path $repoRoot 'build.ps1'
$outputDir = Join-Path $repoRoot '.build\x64\Debug'
$pluginDir = Join-Path $repoRoot '.build\x64\Debug\Plugins'
$monitorExe = Join-Path $outputDir 'RedSalamanderMonitor.exe'
$searchServiceExe = Join-Path $outputDir 'RedSalamanderSearchService.exe'

Describe 'RedSalamander targeted plugin deployment' {
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

        $process = Start-Process `
            -FilePath 'powershell.exe' `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $buildScript, '-ProjectName', 'RedSalamander', '-Configuration', 'Debug', '-Platform', 'x64') `
            -Wait `
            -PassThru `
            -NoNewWindow

        $process.ExitCode | Should Be 0

        foreach ($path in @($monitorExe, $searchServiceExe)) {
            (Test-Path $path) | Should Be $true
        }

        $expectedPluginNames = @(
            Get-ChildItem -Path (Join-Path $repoRoot 'Plugins') -Recurse -Filter '*.vcxproj' |
                ForEach-Object { '{0}.dll' -f $_.BaseName } |
                Sort-Object -Unique
        )
        $pluginNames = @(Get-ChildItem -Path $pluginDir -File | Select-Object -ExpandProperty Name)
        $missingPlugins = @($expectedPluginNames | Where-Object { $_ -notin $pluginNames })

        if ($missingPlugins.Count -gt 0) {
            throw "Missing bundled plugin(s): $($missingPlugins -join ', ')"
        }
    }
}
