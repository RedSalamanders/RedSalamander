Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools\RuntimeDependencies.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\PortablePackageSmoke.psm1') -Force
. (Join-Path $repoRoot 'Tools\VcpkgToolIdentity.ps1')

function Assert-RSThrowsMatching {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    $message = $null
    try {
        $null = & $ScriptBlock
    } catch {
        $message = $_.Exception.Message
    }

    ($null -ne $message) | Should Be $true
    $message | Should Match $Pattern
}

function New-RSFakePortablePackageTree {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$Platform
    )

    $plugins = Join-Path $Root 'Plugins'
    New-Item -ItemType Directory -Path $plugins -Force | Out-Null

    $requiredFiles = @(
        'RedLauncher.exe',
        'red.exe',
        'RedSalamander.exe',
        'RedSalamanderMonitor.exe',
        'Plugins\FileSystem.dll',
        'Plugins\FileSystem7z.dll',
        'Plugins\FileSystemCurl.dll',
        'Plugins\FileSystemDummy.dll',
        'Plugins\FileSystemGoogleDrive.dll',
        'Plugins\FileSystemMicrosoftDrive.dll',
        'Plugins\FileSystemS3.dll',
        'Plugins\ViewerText.dll',
        'Plugins\ViewerSqlite.dll',
        'Plugins\ViewerSpace.dll',
        'Plugins\ViewerImgRaw.dll',
        'Plugins\ViewerVLC.dll',
        'Plugins\ViewerPE.dll',
        'Plugins\ViewerWeb.dll'
    )
    foreach ($relativePath in $requiredFiles) {
        [System.IO.File]::WriteAllBytes((Join-Path $Root $relativePath), [byte[]](0x52, 0x53))
    }

    $manifest = Get-RSRuntimeDependencyManifest -RepoRoot $repoRoot -Configuration $Configuration -Platform $Platform
    foreach ($outputName in @($manifest.Dependencies | Where-Object Required | Select-Object -ExpandProperty OutputName -Unique)) {
        [System.IO.File]::WriteAllBytes((Join-Path $plugins $outputName), [byte[]](0x52, 0x53))
    }
}

Describe 'Declarative runtime dependency graph' {
    It 'uses one complete manifest for every supported flavor and platform' {
        [xml]$document = Get-Content -LiteralPath (Join-Path $repoRoot 'RuntimeDependencies.props') -Raw
        $nodes = @($document.Project.ItemGroup.RSRuntimeDependency)
        $ids = @($nodes | ForEach-Object { [string]$_.Include })
        $ids.Count | Should Be (@($ids | Sort-Object -Unique).Count)

        foreach ($configuration in 'Debug', 'ASan Debug', 'Release') {
            foreach ($platform in 'x64', 'ARM64') {
                $manifest = Get-RSRuntimeDependencyManifest -RepoRoot $repoRoot -Configuration $configuration -Platform $platform
                ($manifest.Dependencies.Count -gt 0) | Should Be $true
                @($manifest.Dependencies | Where-Object {
                        [string]::IsNullOrWhiteSpace($_.Id) -or
                        $_.Projects.Count -eq 0 -or
                        [string]::IsNullOrWhiteSpace($_.Source) -or
                        [string]::IsNullOrWhiteSpace($_.OutputName)
                    }).Count | Should Be 0
            }
        }
    }

    It 'replaces plugin-local xcopy batches with fail-closed MSBuild tasks' {
        $projects = @(
            'Plugins\FileSystem7z\FileSystem7z.vcxproj',
            'Plugins\FileSystemCurl\FileSystemCurl.vcxproj',
            'Plugins\FileSystemGoogleDrive\FileSystemGoogleDrive.vcxproj',
            'Plugins\FileSystemS3\FileSystemS3.vcxproj',
            'Plugins\ViewerImgRaw\ViewerImgRaw.vcxproj',
            'Plugins\ViewerWeb\ViewerWeb.vcxproj'
        )
        foreach ($project in $projects) {
            $source = Get-Content -LiteralPath (Join-Path $repoRoot $project) -Raw
            $source | Should Not Match '(?i)\bxcopy\b'
            $source | Should Not Match '<PostBuildEvent>'
        }

        $targets = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.targets') -Raw
        $targets | Should Match 'MSBuildThisFileDirectory\)RuntimeDependencies\.props'
        $targets | Should Match '<Error\s+Condition='
        $targets | Should Match '<Copy\s+SourceFiles='
        $targets | Should Match '<Delete\s+Files='
    }

    It 'honors test-harness definitions in Release and permits bounded build parallelism' {
        $props = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.props') -Raw
        $monitorProject = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\MonitorTest\MonitorTest.vcxproj') -Raw
        $buildScript = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $monitorProject | Should Match '<RSTestAdditionalPreprocessorDefinitions>ENABLE_TESTS;'
        $props | Should Match '<PreprocessorDefinitions>\$\(RSTestAdditionalPreprocessorDefinitions\)%\(PreprocessorDefinitions\)</PreprocessorDefinitions>'
        $buildScript | Should Match '\[int\]\$MaxCpuCount\s*=\s*0'
        $buildScript | Should Match '"/m:\$MaxCpuCount"'
    }

    It 'rejects a missing required staged dependency by exact output name' {
        $output = Join-Path $TestDrive "runtime-output-$([Guid]::NewGuid().ToString('N'))"
        New-RSFakePortablePackageTree -Root $output -Configuration Debug -Platform x64
        $null = Assert-RSRuntimeDependenciesInOutput -RepoRoot $repoRoot -BuildOutputDir $output -Configuration Debug -Platform x64

        $missing = 'libssh2.dll'
        Remove-Item -LiteralPath (Join-Path $output "Plugins\$missing")
        Assert-RSThrowsMatching {
            Assert-RSRuntimeDependenciesInOutput -RepoRoot $repoRoot -BuildOutputDir $output -Configuration Debug -Platform x64
        } -Pattern ([regex]::Escape($missing))
    }

    It 'validates a clean extracted portable package from the shared manifest' {
        $stage = Join-Path $TestDrive "package-stage-$([Guid]::NewGuid().ToString('N'))"
        $zip = Join-Path $TestDrive "package-$([Guid]::NewGuid().ToString('N')).zip"
        New-RSFakePortablePackageTree -Root $stage -Configuration Release -Platform x64
        Compress-Archive -Path "$stage\*" -DestinationPath $zip

        $result = Test-RSPortablePackage `
            -RepoRoot $repoRoot `
            -ZipPath $zip `
            -BuildOutputDir $stage `
            -Configuration Release `
            -Platform x64 `
            -ScratchRoot (Join-Path $TestDrive 'smoke') `
            -SkipExecution

        $result.ExecutionSkipped | Should Be $true
        $result.RequiredFileCount | Should Be 18
    }

    It 'makes fresh-extraction app and plugin smoke part of portable packaging' {
        $packageScript = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\zip\build-zip.ps1') -Raw
        $releaseWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\release.yml') -Raw
        $packageScript | Should Match 'PortablePackageSmoke\.psm1'
        $packageScript | Should Match 'Test-RSPortablePackage'
        $packageScript | Should Match 'Compress-Archive[\s\S]+Test-RSPortablePackage'
        $smokeModule = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\PortablePackageSmoke.psm1') -Raw
        $smokeModule | Should Match "ArgumentList @\('--package-smoke'\)"
        $releaseWorkflow | Should Match 'Installer\\zip\\build-zip\.ps1'
    }
}

Describe 'Pinned build-tool and CI identity' {
    It 'separately pins and validates the vcpkg executable revision locally and in CI' {
        $pin = Get-RSVcpkgToolPin -RepoRoot $repoRoot
        $pin.Repository | Should Match '^https://'
        $pin.Commit | Should Match '^[0-9a-f]{40}$'

        $localInstall = Get-Content -LiteralPath (Join-Path $repoRoot 'vcpkg-install.ps1') -Raw
        $ci = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw
        $localInstall | Should Match 'Assert-RSVcpkgToolIdentity'
        $ci | Should Match "hashFiles\('vcpkg\.json', 'vcpkg-tool\.json'\)"
        $ci | Should Match 'git checkout --detach \$toolCommit'
        $ci | Should Match '\$actualToolCommit -ne \$toolCommit'
        $ci | Should Not Match 'vcpkg(?:\.exe)?\s+integrate\s+install'
        $ci | Should Match '\$env:ForceImportBeforeCppTargets\s*=\s*\$vcpkgProps'
        $ci | Should Match '\$env:ForceImportAfterCppTargets\s*=\s*\$vcpkgTargets'
        $ci | Should Match '\$isDebugConfiguration\s*=\s*"\$\{\{ inputs\.configuration \}\}"\s+-like\s+"\*Debug\*"'
        $ci | Should Match '\$env:VcpkgConfiguration\s*=\s*\$vcpkgConfiguration'
        $ci | Should Match 'if \(\$isDebugConfiguration\) \{ "debug\\lib" \} else \{ "lib" \}'
        $ci | Should Match 'if \(\$isDebugConfiguration\) \{ "debug\\bin" \} else \{ "bin" \}'
        $ci | Should Match '\$env:VcpkgManifestInstall\s*=\s*"false"'
        $ci | Should Match '\$env:VCPkgLocalAppDataDisabled\s*=\s*"true"'
        $ci | Should Match 'vcpkg\\scripts\\buildsystems\\msbuild'
        $ci.IndexOf('$env:ForceImportBeforeCppTargets') | Should BeLessThan $ci.IndexOf('& $buildScript @buildArgs')
        $ci.IndexOf('$env:ForceImportAfterCppTargets') | Should BeLessThan $ci.IndexOf('& $buildScript @buildArgs')
    }

    It 'keeps ARM64 compilation and critical contract suites in the PR gate' {
        $ci = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\ci.yml') -Raw
        $subclassGuard = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Verify-NoSubclassManager.ps1') -Raw
        $selfTests = Get-Content -LiteralPath (Join-Path $repoRoot 'Specs\Testing\Testing_SelfTests.md') -Raw
        $ci | Should Match 'push:\s*\r?\n\s*branches:\s*\[main, master\]'
        $ci | Should Match 'pull_request:\s*\r?\n\s*branches:\s*\[main, master\]'
        $ci | Should Match 'id:\s*cpp_changes'
        $ci | Should Match 'git diff --name-only "\$env:BASE_SHA\.\.\.HEAD"'
        @($ci | Select-String -Pattern "if: steps\.cpp_changes\.outputs\.changed == 'true'" -AllMatches).Matches.Count | Should Be 3
        $subclassGuard | Should Match "Get-Command 'rg' -ErrorAction SilentlyContinue"
        $subclassGuard | Should Match 'Select-String -SimpleMatch -Pattern \$pattern'
        $subclassGuard | Should Match '\$global:LASTEXITCODE\s*=\s*0\s*$'
        $ci | Should Match 'platform:\s*ARM64'
        $ci | Should Match 'configuration:\s*Debug'
        $selfTests | Should Match 'PluginContractTests, SettingsSchemaTests, and CrashHandlingTests'
        $selfTests | Should Match 'RedSalamanderMonitorEtwLatency remains a broader closeout-only `-Suite Full` gate'
    }

    It 'runs scheduled and high-risk ASan with a seeded detector proof before green contracts' {
        $asan = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\asan.yml') -Raw
        $reusable = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw
        $harness = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\PluginContractTests\PluginContractTests.cpp') -Raw
        $asan | Should Match 'schedule:'
        $asan | Should Match 'pull_request:'
        $asan | Should Match '\*\*/\*\.cpp'
        $asan | Should Match 'configuration:\s*ASan Debug'
        $reusable | Should Match '--asan-seed-heap-overflow'
        $reusable | Should Match 'not rejected by AddressSanitizer'
        $reusable | Should Match 'AddressSanitizer diagnostic'
        $harness | Should Match 'RS_ASAN_DEBUG_BUILD'
        $harness | Should Match 'raw\[16\]\s*=\s*0x5Au'
    }
}
