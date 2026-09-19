Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools\Modules\Packaging\RuntimeDependencies.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Packaging\PortablePackageSmoke.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Build\VcpkgToolIdentity.psm1') -Force -ErrorAction Stop

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
        'Plugins\ViewerWeb.dll',
        'Plugins\Terminal.dll'
    )
    foreach ($relativePath in $requiredFiles) {
        [System.IO.File]::WriteAllBytes((Join-Path $Root $relativePath), [byte[]](0x52, 0x53))
    }

    $manifest = Get-RSRuntimeDependencyManifest -RepoRoot $repoRoot -Configuration $Configuration -Platform $Platform
    foreach ($dependency in @($manifest.Dependencies | Where-Object Required)) {
        $destinationRoot = if ($dependency.OutputRoot -eq 'BuildOutput') { $Root } else { $plugins }
        $outputPath = Join-Path $destinationRoot $dependency.OutputName
        New-Item -ItemType Directory -Path (Split-Path -Parent $outputPath) -Force | Out-Null
        [System.IO.File]::WriteAllBytes($outputPath, [byte[]](0x52, 0x53))
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
                        $_.OutputRoot -notin @('Plugins', 'BuildOutput') -or
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
        $targets = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.targets') -Raw
        $testTargets = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\Directory.Build.targets') -Raw
        $monitorProject = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\MonitorTest\MonitorTest.vcxproj') -Raw
        $buildScript = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $monitorProject | Should Match '<RSTestAdditionalPreprocessorDefinitions>ENABLE_TESTS;'
        $props | Should Match '<PreprocessorDefinitions>\$\(RSTestAdditionalPreprocessorDefinitions\)%\(PreprocessorDefinitions\)</PreprocessorDefinitions>'
        $props | Should Match "RSBuildEnableTests.*Configuration.*Debug.*ASan Debug.*true"
        $props | Should Match "RSBuildEnableTests.*false"
        $props | Should Not Match "ItemDefinitionGroup[^>]+RSDefaultEnableTests"
        $buildScript | Should Match '\[int\]\$MaxCpuCount\s*=\s*0'
        $buildScript | Should Match '"/m:\$MaxCpuCount"'
        $buildScript | Should Match '"/p:RSBuildEnableTests=\$\(\$effectiveTestsEnabled'
        $buildScript | Should Match 'tests-enabled=\$\(\$effectiveTestsEnabled'
        $buildScript | Should Match "RSBuildCompilerDebugInformation must be 'Embedded' or 'ProgramDatabase'"
        $buildScript | Should Match 'compiler-debug-information=\$\(\$effectiveCompilerDebugInformation\.ToLowerInvariant\(\)\)'
        $buildScript | Should Match '"/p:RSBuildCompilerDebugInformation=\$effectiveCompilerDebugInformation"'
        $buildScript | Should Match 'Test-RSBuildArtifactManifestTestsEnabled'
        $targets | Should Match 'Name="RSWriteBuildArtifactManifest"'
        $targets | Should Match 'DependsOnTargets="RSCopyAsanRuntime"'
        $targets | Should Match 'tests_enabled=\$\(RSBuildEnableTests\)'
        $targets | Should Match 'vcpkg_tlog=\$\(TLogLocation\)\$\(ProjectName\)\.write\.1u\.tlog'
        $targets | Should Match 'runtime=\$\(TargetDir\)\$\(RSAsanRuntimeFileName\)'
        $targets | Should Match "EnableASAN.*ConfigurationType.*Application.*RSAsanRuntimeSourcePath"
        $targets | Should Match '_RSLibraryPathForBuildEnvironment'
        $targets | Should Match 'SetBuildDefaultEnvironmentVariablesDependsOn'
        $targets | Should Match 'Condition="Exists\(''%\(Identity\)''\)"'
        $targets | Should Match '<ItemDefinitionGroup Condition="[^\"]+RSBuildCompilerDebugInformation[^\"]+Embedded'
        $targets | Should Match '<DebugInformationFormat>OldStyle</DebugInformationFormat>'
        $targets | Should Match '<ClCompile Update="\.\.\\Common\\SearchTextHelpers\.cpp;\.\.\\\.\.\\Common\\SearchTextHelpers\.cpp">'
        $targets | Should Match '<MultiProcessorCompilation>false</MultiProcessorCompilation>'
        $testTargets | Should Match 'RSParentDirectoryBuildTargets'
        $testTargets | Should Match "RSIsTestProject.*RSBuildEnableTests.*!='true'"
        $testTargets | Should Match "MSBuildProjectName.*!='PluginContractTests'"
        $testTargets | Should Match '<ConfigurationType>Utility</ConfigurationType>'
        $testTargets | Should Match '(?s)<PropertyGroup Condition="[^"]*RSIsTestProject[^"]*RSBuildEnableTests[^"]*">.*?<CppCleanDependsOn>RSCleanDisabledTestOutputs</CppCleanDependsOn>'
        $testTargets | Should Match '<CppCleanDependsOn>RSCleanDisabledTestOutputs</CppCleanDependsOn>'
        $testTargets | Should Match 'Name="RSCleanDisabledTestOutputs"'
        $testTargets | Should Match '(?s)<Target Name="RSCleanDisabledTestOutputs"\s+Condition="[^"]*RSIsTestProject[^"]*RSBuildEnableTests[^"]*">'
        $testTargets | Should Match '<Delete Files="@\(_RSDisabledTestOutput\)"'

        foreach ($project in @(
            'Tests\FileSystemCurlTests\FileSystemCurlTests.vcxproj',
            'Tests\RedConfigureTests\RedConfigureTests.vcxproj',
            'Tests\ViewerPETests\ViewerPETests.vcxproj'
        )) {
            (Get-Content -LiteralPath (Join-Path $repoRoot $project) -Raw) |
                Should Not Match '<RSBuildEnableTests'
        }

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

    It 'stages and validates the SearchService SQLite runtime at the profile root' {
        foreach ($configuration in @('Debug', 'Release')) {
            foreach ($platform in @('x64', 'ARM64')) {
                $manifest = Get-RSRuntimeDependencyManifest -RepoRoot $repoRoot -Configuration $configuration -Platform $platform
                $sqlite = @($manifest.Dependencies | Where-Object { $_.Projects -contains 'RedSalamanderSearchService' })
                $sqlite.Count | Should Be 1
                $sqlite[0].OutputRoot | Should Be 'BuildOutput'
                $sqlite[0].OutputName | Should Be 'sqlite3.dll'
                $expectedSource = if ($configuration -eq 'Release') { '^bin\\sqlite3\.dll$' } else { '^debug\\bin\\sqlite3\.dll$' }
                $sqlite[0].Source | Should Match $expectedSource
            }
        }

        $output = Join-Path $TestDrive "search-service-runtime-$([Guid]::NewGuid().ToString('N'))"
        New-RSFakePortablePackageTree -Root $output -Configuration Release -Platform x64
        Test-Path -LiteralPath (Join-Path $output 'sqlite3.dll') -PathType Leaf | Should Be $true
        $null = Assert-RSRuntimeDependenciesInOutput -RepoRoot $repoRoot -BuildOutputDir $output -Configuration Release -Platform x64
        Remove-Item -LiteralPath (Join-Path $output 'sqlite3.dll')
        Assert-RSThrowsMatching {
            Assert-RSRuntimeDependenciesInOutput -RepoRoot $repoRoot -BuildOutputDir $output -Configuration Release -Platform x64
        } -Pattern 'sqlite3\.dll'
    }

    It 'keeps Terminal runtime sources literal, platform-specific, and nested under the plugin output' {
        $expectedOutputs = @(
            'TerminalRuntime\ghostty-vt.dll'
            'TerminalRuntime\terminal-runtime-identity.json'
            'TerminalRuntime\LICENSE-Ghostty.txt'
            'TerminalRuntime\LICENSE-uucode.txt'
            'TerminalRuntime\LICENSE-uucode-Bjoern-Hoehrmann.txt'
            'TerminalRuntime\LICENSE-Unicode.txt'
            'TerminalRuntime\LICENSE-Wuffs.txt'
            'TerminalRuntime\LICENSE-Zig.txt'
        )
        foreach ($platform in @('x64', 'ARM64')) {
            $manifest = Get-RSRuntimeDependencyManifest -RepoRoot $repoRoot -Configuration Debug -Platform $platform
            $terminalDependencies = @($manifest.Dependencies | Where-Object { $_.Projects -contains 'Terminal' })
            $terminalDependencies.Count | Should Be $expectedOutputs.Count
            @($terminalDependencies.OutputName | Sort-Object) -join "`n" |
                Should Be (@($expectedOutputs | Sort-Object) -join "`n")
            foreach ($dependency in $terminalDependencies) {
                $dependency.SourceRoot | Should Be 'RepoRoot'
                $dependency.Source | Should Not Match '[\$%*?]'
                $dependency.Source | Should Match ([regex]::Escape("Product\$platform\"))
                $dependency.OutputName | Should Match '^TerminalRuntime\\[^\\]+$'
            }
        }
    }

    It 'validates a clean extracted portable package from the shared manifest' {
        Mock -CommandName Test-RSPortableGhosttyRuntime -ModuleName PortablePackageSmoke -MockWith {
            [pscustomobject]@{
                DllSha256 = 'synthetic'
                IdentitySha256 = 'synthetic'
                Machine = 'AMD64'
                RuntimeLoadSkipped = $true
            }
        }

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
        $result.RequiredFileCount | Should Be 19
        $result.GhosttyRuntime.Machine | Should Be 'AMD64'
        Assert-MockCalled -CommandName Test-RSPortableGhosttyRuntime -ModuleName PortablePackageSmoke -Times 1 -Exactly -Scope It
    }

    It 'makes fresh-extraction app and plugin smoke part of portable packaging' {
        $packageScript = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\zip\build-zip.ps1') -Raw
        $releaseWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\release.yml') -Raw
        $testTargets = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\Directory.Build.targets') -Raw
        $pluginContractHarness = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\PluginContractTests\PluginContractTests.cpp') -Raw
        $packageScript | Should Match 'PortablePackageSmoke\.psm1'
        $packageScript | Should Match 'Test-RSPortablePackage'
        $packageScript | Should Match 'Compress-Archive[\s\S]+Test-RSPortablePackage'
        $smokeModule = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Modules\Packaging\PortablePackageSmoke.psm1') -Raw
        $smokeModule | Should Match "ArgumentList @\('--package-smoke'\)"
        $smokeModule | Should Match 'Test-RSPortableGhosttyRuntime'
        $smokeModule | Should Match 'Portable Ghostty runtime identity differs from the receipt-bound shipping-profile identity'
        $smokeModule | Should Not Match 'NativeLibrary\]::(Load|Free)'
        $pluginContractHarness | Should Match 'TestPackagedGhosttyRuntimeLoad'
        $pluginContractHarness | Should Match 'Plugins\\\\TerminalRuntime\\\\ghostty-vt\.dll'
        $pluginContractHarness | Should Match 'LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR \| LOAD_LIBRARY_SEARCH_APPLICATION_DIR \| LOAD_LIBRARY_SEARCH_SYSTEM32'
        $pluginContractHarness | Should Match 'wil::unique_hmodule runtime'
        $testTargets | Should Match "MSBuildProjectName.*!='PluginContractTests'"
        $packageScript | Should Match 'PluginContractTests\.exe'
        $pluginContractHarness | Should Match 'TestTerminalSizedRecords\(success, ! packageSmoke\)'
        $pluginContractHarness | Should Match 'TestTerminalLocalizedContracts\(success, ! packageSmoke\)'
        $smokeModule | Should Match 'Build\\SanitizedEnvironment\.psm1'
        $smokeModule | Should Match '(?s)function Invoke-RSPortableSmokeProcess.+?Invoke-RSProcess.+?-Arguments \$ArgumentList'
        $smokeModule | Should Not Match '(?m)^\s*&\s+\$FilePath\s+@ArgumentList'
        $releaseWorkflow | Should Match 'Installer\\zip\\build-zip\.ps1'
    }

    It 'publishes portable ZIPs through guarded same-directory atomic replacement' {
        $packageScript = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\zip\build-zip.ps1') -Raw
        $packageScript | Should Match '(?s)Resolve-RSGuardedRepositoryPath.+?-GuardedRelativeRoot\s+''\.build\\AppPackages''.+?-CreateDirectory'
        $packageScript | Should Match '(?s)\$StagedZipPath\s*=\s*Join-Path\s+\$PackageOutputDir.+?\.tmp\.zip.+?NewGuid'
        $packageScript | Should Match '(?s)Compress-Archive.+?-DestinationPath\s+\$StagedZipPath.+?Test-RSPortablePackage.+?-ZipPath\s+\$StagedZipPath'
        $packageScript | Should Match '(?s)Publish-RSGuardedRepositoryFile.+?-StagedPath\s+\$StagedZipPath.+?-DestinationPath\s+\$ZipPath'
        $packageScript | Should Not Match '(?m)^\s*Compress-Archive[^\r\n]+-DestinationPath\s+\$ZipPath'
    }

    It 'runs both fresh-extraction smoke executables through contained process support' {
        Mock -CommandName Invoke-RSProcess -ModuleName PortablePackageSmoke -MockWith { 0 }
        Mock -CommandName Test-RSPortableGhosttyRuntime -ModuleName PortablePackageSmoke -MockWith {
            [pscustomobject]@{
                DllSha256 = 'synthetic'
                IdentitySha256 = 'synthetic'
                Machine = 'AMD64'
                RuntimeLoadSkipped = $false
            }
        }

        $stage = Join-Path $TestDrive "contained-smoke-stage-$([Guid]::NewGuid().ToString('N'))"
        $zip = Join-Path $TestDrive "contained-smoke-$([Guid]::NewGuid().ToString('N')).zip"
        New-RSFakePortablePackageTree -Root $stage -Configuration Release -Platform x64
        [System.IO.File]::WriteAllBytes(
            (Join-Path $stage 'PluginContractTests.exe'),
            [byte[]](0x52, 0x53))
        Compress-Archive -Path "$stage\*" -DestinationPath $zip

        $result = Test-RSPortablePackage `
            -RepoRoot $repoRoot `
            -ZipPath $zip `
            -BuildOutputDir $stage `
            -Configuration Release `
            -Platform x64 `
            -ScratchRoot (Join-Path $TestDrive 'contained-smoke')

        $result.ExecutionSkipped | Should Be $false
        Assert-MockCalled -CommandName Invoke-RSProcess -ModuleName PortablePackageSmoke -Times 2 -Exactly -Scope It
        Assert-MockCalled -CommandName Test-RSPortableGhosttyRuntime -ModuleName PortablePackageSmoke -Times 1 -Exactly -Scope It
    }
}

Describe 'MSI and MSIX package source contracts' {
    It 'requires explicit standalone package identities and never allocates from repository version state' {
        foreach ($relativePath in @(
                'Installer/msi/build-msi.ps1',
                'Installer/msi/build-msi-symbols.ps1',
                'Installer/zip/build-zip.ps1')) {
            $source = Get-Content -LiteralPath (Join-Path $repoRoot $relativePath) -Raw
            $source | Should Match '\$BuildNumber\s+-le\s+0'
            $source | Should Match 'New-RSVersionContext'
            $source | Should Not Match 'Resolve-RSVersionContext|Read-RSVersionContext'
        }

        $winget = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer/winget/generate-manifest.ps1') -Raw
        $winget | Should Match '(?s)IsNullOrWhiteSpace\(\$Version\).+?\$BuildNumber\s+-le\s+0'
        $winget | Should Match 'New-RSVersionContext'
        $winget | Should Not Match 'Resolve-RSVersionContext|Read-RSVersionContext'

        $buildScript = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $resolver = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/ResolveVersionForMsbuild.ps1') -Raw
        $buildScript | Should Match 'Resolve-RSVersionContext'
        $resolver | Should Match 'Resolve-RSVersionContext'
        $releaseWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/release.yml') -Raw
        $releaseWorkflow | Should Match 'Resolve-RSVersionContext'
        $releaseWorkflow | Should Not Match 'Get-RSVersionContext'
        $buildScript | Should Not Match 'Use-RSVersionStateLock|Save-RSVersionContext'
        $resolver | Should Not Match 'Use-RSVersionStateLock|Save-RSVersionContext'
    }

    It 'keeps standalone package mutators under balanced fail-closed coordination' {
        $contracts = @(
            @{ Path = 'Installer/msi/build-msi.ps1'; ReadsProfile = $true; ContainedProcesses = 3 },
            @{ Path = 'Installer/msi/build-msi-symbols.ps1'; ReadsProfile = $true; ContainedProcesses = 1 },
            @{ Path = 'Installer/msix/build-msix.ps1'; ReadsProfile = $true; ContainedProcesses = 1 },
            @{ Path = 'Installer/zip/build-zip.ps1'; ReadsProfile = $true; ContainedProcesses = 0 },
            @{ Path = 'Installer/winget/generate-manifest.ps1'; ReadsProfile = $false; ContainedProcesses = 0 },
            @{ Path = 'Installer/msix/GenerateAssets.ps1'; ReadsProfile = $false; ContainedProcesses = 0 },
            @{ Path = 'Installer/msix/UpdateManifestVersion.ps1'; ReadsProfile = $false; ContainedProcesses = 0 }
        )

        foreach ($contract in $contracts) {
            $source = Get-Content -LiteralPath (Join-Path $repoRoot $contract.Path) -Raw
            $source | Should Match 'ArtifactOperationLock\.psm1'
            $source | Should Match "coordination\s*=\s*'packaging'"
            $source | Should Match 'configuration\s*='
            $source | Should Match 'platform\s*='
            $source | Should Match 'Enter-RSArtifactOperationLock'
            $source | Should Match '\.WasAbandoned'
            $source | Should Match 'Set-RSArtifactOperationContaminated'
            $source | Should Match 'Read-RSArtifactOperationContamination|Test-RSArtifactOperationContaminated'
            $source | Should Match 'Get-RSArtifactContaminationMarkerPath'
            $source | Should Match '(?s)finally\s*\{.*?Exit-RSArtifactOperationLock\s+-Lock\s+\$[A-Za-z]*PackagingLock'
            $source | Should Match '(?i)review.+restore.+Installer inputs.+\.build\\AppPackages.+remove the marker explicitly'
            $source | Should Not Match 'Clear-RSArtifactOperationContaminated'
            $source | Should Match '(?im)^\s*\.OUTPUTS\s*$'
            $source | Should Match '(?im)^\s*\.NOTES\s*$'

            if ($contract.ReadsProfile) {
                $profileEnter = $source.IndexOf(
                    '$ArtifactProfileLock = Enter-RSArtifactOperationLock',
                    [StringComparison]::OrdinalIgnoreCase)
                $packagingEnter = $source.IndexOf(
                    '$PackagingLock = Enter-RSArtifactOperationLock',
                    [StringComparison]::OrdinalIgnoreCase)
                $packagingExit = $source.LastIndexOf(
                    'Exit-RSArtifactOperationLock -Lock $PackagingLock',
                    [StringComparison]::OrdinalIgnoreCase)
                $profileExit = $source.LastIndexOf(
                    'Exit-RSArtifactOperationLock -Lock $ArtifactProfileLock',
                    [StringComparison]::OrdinalIgnoreCase)
                ($profileEnter -ge 0) | Should Be $true
                ($profileEnter -lt $packagingEnter) | Should Be $true
                ($packagingEnter -lt $packagingExit) | Should Be $true
                ($packagingExit -lt $profileExit) | Should Be $true
                $source | Should Match '(?s)Test-RSArtifactOperationContaminated.+?-Scope\s+\$[A-Za-z]*ArtifactProfileScope'
                $residualCheck = $source.IndexOf('Assert-RSNoResidualArtifactToolProcesses', [StringComparison]::OrdinalIgnoreCase)
                ($profileEnter -lt $residualCheck) | Should Be $true
                ($residualCheck -lt $packagingEnter) | Should Be $true
                $source | Should Match '(?s)Assert-RSNoResidualArtifactToolProcesses.+?-Scope\s+\$[A-Za-z]*ArtifactProfileScope'
            }

            if ($contract.ContainedProcesses -gt 0) {
                $source | Should Match 'SanitizedEnvironment\.psm1'
                ([regex]::Matches($source, 'Invoke-RSProcess').Count -ge $contract.ContainedProcesses) | Should Be $true
                $source | Should Not Match '(?m)^\s*&\s+\$wixExe\s+@wixArgs'
            }
        }

        $mainMsi = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer/msi/build-msi.ps1') -Raw
        ([regex]::Matches($mainMsi, 'Invoke-RSProcess').Count -ge 3) | Should Be $true
        $mainMsi | Should Not Match '(?m)^\s*&\s+robocopy\.exe'
        foreach ($msiScript in @(
                $mainMsi,
                (Get-Content -LiteralPath (Join-Path $repoRoot 'Installer/msi/build-msi-symbols.ps1') -Raw))) {
            $msiScript | Should Not Match '(?i)extension\s+add'
            $msiScript | Should Match 'Packaging does not mutate machine-global WiX extensions'
            $msiScript | Should Match '(?s)Resolve-RSGuardedRepositoryPath.+?-GuardedRelativeRoot\s+''\.build\\AppPackages''.+?-CreateDirectory'
            $msiScript | Should Match '(?s)\$workRoot\s*=.+?NewGuid.+?try\s*\{.+?finally\s*\{\s*if\s*\(Test-Path\s+-LiteralPath\s+\$workRoot\)\s*\{\s*Remove-Item\s+-LiteralPath\s+\$workRoot\s+-Recurse\s+-Force'
            $msiScript | Should Match '(?m)^\s*New-Item\s+-ItemType\s+Directory\s+-Path\s+\$workRoot\s*\|\s*Out-Null\s*$'
            $msiScript | Should Not Match '(?m)^\s*New-Item\s+-ItemType\s+Directory\s+-Path\s+\$workRoot\s+-Force'
            $cleanupIndex = $msiScript.LastIndexOf('Remove-Item -LiteralPath $workRoot', [StringComparison]::OrdinalIgnoreCase)
            $lockExitIndex = $msiScript.LastIndexOf('Exit-RSArtifactOperationLock -Lock $packagingLock', [StringComparison]::OrdinalIgnoreCase)
            ($cleanupIndex -ge 0) | Should Be $true
            ($cleanupIndex -lt $lockExitIndex) | Should Be $true
        }

        $wingetGenerator = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer/winget/generate-manifest.ps1') -Raw
        $wingetGenerator | Should Match '(?s)Resolve-RSGuardedRepositoryPath.+?-GuardedRelativeRoot\s+''\.build\\AppPackages''.+?-CreateDirectory'
        $manifestStamper = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer/msix/UpdateManifestVersion.ps1') -Raw
        $manifestStamper | Should Match '(?s)Resolve-RSGuardedRepositoryPath.+?-GuardedRelativeRoot\s+''Installer\\msix''.+?-Path\s+\$ManifestPath'
        $manifestStamper | Should Match '(?s)Resolve-RSGuardedRepositoryPath.+?-GuardedRelativeRoot\s+''\.build\\AppPackages''.+?-Path\s+\(Split-Path \$OutputManifestPath -Parent\).+?-CreateDirectory'
        $manifestStamper | Should Match '\[IO\.File\]::Open\(\$OutputManifestPath, \[IO\.FileMode\]::CreateNew'
        $manifestStamper | Should Not Match 'XmlWriter\]::Create\(\$ManifestPath'

        $guardModule = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Build/ArtifactOperationLock.psm1') -Raw
        $guardModule | Should Match 'function\s+Resolve-RSGuardedRepositoryPath'
        $guardModule | Should Match 'FileAttributes\]::ReparsePoint'
        $guardModule | Should Match "'Resolve-RSGuardedRepositoryPath'"
    }

    It 'validates exact Winget versions and portable archive identities before hashing' {
        $scriptPath = Join-Path $repoRoot 'Installer/winget/generate-manifest.ps1'
        $tokens = $null
        $parseErrors = $null
        $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $scriptPath,
            [ref]$tokens,
            [ref]$parseErrors)
        $parseErrors.Count | Should Be 0
        foreach ($functionName in @('Assert-WingetManifestVersion', 'Resolve-WingetPortableZip')) {
            $functionAst = $scriptAst.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq $functionName
            }, $true)
            ($null -ne $functionAst) | Should Be $true
            Invoke-Expression $functionAst.Extent.Text
        }

        { Assert-WingetManifestVersion -CandidateVersion '7.0' } |
            Should Throw 'exactly three numeric components'
        { Assert-WingetManifestVersion -CandidateVersion '7.0.0' } |
            Should Throw 'build component must be positive'
        { Assert-WingetManifestVersion -CandidateVersion '65536.0.1' } |
            Should Throw 'range 0..65535'
        Assert-WingetManifestVersion -CandidateVersion '7.0.184'

        $expectedLeaf = 'RedSalamander-7.0.184-x64-Portable.zip'
        $zipPath = Join-Path $TestDrive $expectedLeaf
        [System.IO.File]::WriteAllBytes($zipPath, [byte[]](0x52, 0x53))
        (Resolve-WingetPortableZip `
                -CandidatePath $zipPath `
                -ExpectedLeaf $expectedLeaf `
                -Label 'x64') | Should Be $zipPath
        {
            Resolve-WingetPortableZip `
                -CandidatePath $zipPath `
                -ExpectedLeaf 'RedSalamander-7.0.183-x64-Portable.zip' `
                -Label 'x64'
        } | Should Throw 'leaf must be exactly'
        {
            Resolve-WingetPortableZip `
                -CandidatePath (Join-Path $TestDrive 'missing.zip') `
                -ExpectedLeaf 'missing.zip' `
                -Label 'ARM64'
        } | Should Throw 'must be an existing file'
    }

    It 'keeps each release MSIX build in one contained packaging transaction' {
        $workflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\release.yml') -Raw
        $workflow | Should Match '(?s)Installer\\msix\\build-msix\.ps1.+?-Version "\$\{\{ needs\.version-info\.outputs\.release_version \}\}".+?-Configuration Release.+?-Platform "\$\{\{ matrix\.platform \}\}".+?-MSBuildPath \$env:MSBUILD_EXE_PATH'
        $workflow | Should Not Match '(?m)^\s+MSBuild\.exe\s+@msixParams'

        $wrapper = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer/msix/build-msix.ps1') -Raw
        $wrapper | Should Match '(?s)& \$generateAssetsScript.+?& \$updateManifestScript.+?Invoke-RSProcess'
        $wrapper | Should Match '"/p:RSAppxManifestPath=\$stagedManifestPath"'
        $wrapper | Should Match '(?s)finally\s*\{.+?Remove-Item -LiteralPath \$stagedManifestDirectory -Recurse -Force.+?Exit-RSArtifactOperationLock'
        $profileEnter = $wrapper.IndexOf('$artifactProfileLock = Enter-RSArtifactOperationLock', [StringComparison]::OrdinalIgnoreCase)
        $packagingEnter = $wrapper.IndexOf('$packagingLock = Enter-RSArtifactOperationLock', [StringComparison]::OrdinalIgnoreCase)
        $dependencyEnter = $wrapper.IndexOf('$sharedDependencyLock = Enter-RSArtifactOperationLock', [StringComparison]::OrdinalIgnoreCase)
        $msbuildInvoke = $wrapper.IndexOf('$exitCode = Invoke-RSProcess', [StringComparison]::OrdinalIgnoreCase)
        $dependencyExit = $wrapper.LastIndexOf('Exit-RSArtifactOperationLock -Lock $sharedDependencyLock', [StringComparison]::OrdinalIgnoreCase)
        $packagingExit = $wrapper.LastIndexOf('Exit-RSArtifactOperationLock -Lock $packagingLock', [StringComparison]::OrdinalIgnoreCase)
        $profileExit = $wrapper.LastIndexOf('Exit-RSArtifactOperationLock -Lock $artifactProfileLock', [StringComparison]::OrdinalIgnoreCase)
        ($profileEnter -lt $packagingEnter) | Should Be $true
        ($packagingEnter -lt $dependencyEnter) | Should Be $true
        ($dependencyEnter -lt $msbuildInvoke) | Should Be $true
        ($msbuildInvoke -lt $dependencyExit) | Should Be $true
        ($dependencyExit -lt $packagingExit) | Should Be $true
        ($packagingExit -lt $profileExit) | Should Be $true
        $wrapper | Should Match "coordination\s*=\s*'shared-dependency'"
        $wrapper | Should Match '(?s)Read-RSArtifactOperationContamination.+?-Scope\s+\$sharedDependencyScope'
        $wrapper | Should Match '(?s)Assert-RSNoResidualArtifactToolProcesses.+?-Scope\s+\$sharedDependencyScope'

        $buildScript = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $buildScript | Should Match '(?s)& \$msixAssetsScript\s+-Configuration \$Configuration\s+-Platform \$Platform'
        $buildScript | Should Match '"/p:RSAppxManifestPath=\$stagedMsixManifest"'
        $buildScript | Should Match 'Remove-Item -LiteralPath \$stagedMsixManifestDirectory -Recurse -Force'
        $buildScript | Should Match '(?s)& \$msixVersionScript.+?-Version \$versionContext\.PackagingVersion.+?-Platform \$Platform.+?-Configuration \$Configuration'
    }

    It 'keeps MSI staging aligned with the three shipping executables and optional service component' {
        $buildMsi = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\msi\build-msi.ps1') -Raw
        $product = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\msi\Product.wxs') -Raw

        $buildMsi | Should Match '\$shippingExes\s*=\s*@\("RedSalamander\.exe",\s*"RedSalamanderMonitor\.exe",\s*"RedSalamanderSearchService\.exe"\)'
        $buildMsi | Should Match 'Expected shipping executable not found'
        $product | Should Match '<MajorUpgrade[\s\S]+Schedule="afterInstallInitialize"'
        $product | Should Match '<StandardDirectory Id="ProgramFiles64Folder">'
        $product | Should Match 'Condition="INSTALLSEARCHSERVICE&lt;&gt;&quot;0&quot;"'
        $product | Should Match '<ServiceInstall[\s\S]+Name="RedSalamanderSearchService"[\s\S]+Start="auto"[\s\S]+Account="LocalSystem"'
        $product | Should Match 'Target="\[INSTALLFOLDER\]RedSalamander\.exe"'
        $product | Should Match 'Target="\[INSTALLFOLDER\]RedSalamanderMonitor\.exe"'
    }

    It 'keeps the symbols MSI allowlist aligned with shipped binaries and plugin PDBs' {
        $symbols = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\msi\build-msi-symbols.ps1') -Raw

        foreach ($pdb in @(
                'RedSalamander.pdb',
                'RedSalamanderMonitor.pdb',
                'RedSalamanderSearchService.pdb',
                'Common.pdb'
            )) {
            $symbols | Should Match ([regex]::Escape(('"{0}"' -f $pdb)))
        }
        $symbols | Should Match 'Copy-Item\s+-Path\s+\(Join-Path\s+\$pluginsSrc\s+"\*\.pdb"\)\s+-Destination\s+\$pluginsDst'
    }

    It 'keeps MSIX identity applications architecture stamping and directory-preserving harvest aligned' {
        $project = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\msix\RedSalamanderInstaller.wapproj') -Raw
        $manifest = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\msix\Package.appxmanifest') -Raw
        $stamper = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\msix\UpdateManifestVersion.ps1') -Raw

        $manifest | Should Match '<Identity Name="RedSalamander" Publisher="CN=RedSalmanders" Version="7\.0\.0\.0" ProcessorArchitecture="x64"'
        ([regex]::Matches($manifest, '<Application Id="')).Count | Should Be 2
        $manifest | Should Match '<Application Id="RedSalamander" Executable="RedSalamander\.exe"'
        $manifest | Should Match '<Application Id="RedSalamanderMonitor" Executable="RedSalamanderMonitor\.exe"'
        $project | Should Match '<TargetPlatformMinVersion>10\.0\.22000\.2600</TargetPlatformMinVersion>'
        $project | Should Match '<RSAppxManifestPath Condition="''\$\(RSAppxManifestPath\)''==''''">\$\(MSBuildThisFileDirectory\)Package\.appxmanifest</RSAppxManifestPath>'
        $project | Should Match '<AppxManifest Include="\$\(RSAppxManifestPath\)">'
        foreach ($directory in @('Plugins', 'Lang', 'Themes')) {
            $project | Should Match ("<PackagePath>{0}\\%\(RecursiveDir\)%\(Filename\)%\(Extension\)</PackagePath>" -f $directory)
        }
        $stamper | Should Match '\$identity\.Version\s*=\s*\$identityVersion'
        $stamper | Should Match '\$identity\.ProcessorArchitecture\s*=\s*\$architecture'
    }

    It 'stages MSIX identity changes without changing the tracked manifest' {
        $sourceManifest = Join-Path $repoRoot 'Installer\msix\Package.appxmanifest'
        $sourceDigest = (Get-FileHash -LiteralPath $sourceManifest -Algorithm SHA256).Hash
        $stageRoot = Join-Path $repoRoot ('.build\AppPackages\ManifestStagingTests\{0}' -f [Guid]::NewGuid().ToString('N'))
        $stagePath = Join-Path $stageRoot 'Package.appxmanifest'
        try {
            $staged = & (Join-Path $repoRoot 'Installer\msix\UpdateManifestVersion.ps1') `
                -Version '7.0.184' -Platform ARM64 -OutputManifestPath $stagePath
            $staged | Should Be $stagePath
            (Get-FileHash -LiteralPath $sourceManifest -Algorithm SHA256).Hash | Should Be $sourceDigest
            $xml = [xml](Get-Content -LiteralPath $stagePath -Raw)
            $xml.Package.Identity.Version | Should Be '7.0.184.0'
            $xml.Package.Identity.ProcessorArchitecture | Should Be 'arm64'
            {
                & (Join-Path $repoRoot 'Installer\msix\UpdateManifestVersion.ps1') `
                    -Version '7.0.185' -Platform x64 -OutputManifestPath $stagePath
            } | Should Throw 'unique path is required'
            (Get-FileHash -LiteralPath $sourceManifest -Algorithm SHA256).Hash | Should Be $sourceDigest
            $xmlAfterFailure = [xml](Get-Content -LiteralPath $stagePath -Raw)
            $xmlAfterFailure.Package.Identity.Version | Should Be '7.0.184.0'
        }
        finally {
            if (Test-Path -LiteralPath $stageRoot -PathType Container) {
                Remove-Item -LiteralPath $stageRoot -Recurse -Force
            }
        }
    }
}

Describe 'Pinned build-tool and CI identity' {
    It 'separately pins and validates the vcpkg executable revision locally and in CI' {
        $pin = Get-RSVcpkgToolPin -RepoRoot $repoRoot
        $pin.Repository | Should Match '^https://'
        $pin.Commit | Should Match '^[0-9a-f]{40}$'

        $localInstall = Get-Content -LiteralPath (Join-Path $repoRoot 'vcpkg-install.ps1') -Raw
        $ci = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw
        $props = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.props') -Raw
        $localInstall | Should Match 'Assert-RSVcpkgToolIdentity'
        $props | Should Match '<RSVcpkgPlatformScope[^>]*>x64</RSVcpkgPlatformScope>'
        $props | Should Match '<RSVcpkgPlatformScope[^>]*>arm64</RSVcpkgPlatformScope>'
        $props | Should Match '<VcpkgInstalledDir[^>]*>\$\(RSBuildRoot\)vcpkg_installed\\\$\(RSVcpkgPlatformScope\)\\</VcpkgInstalledDir>'
        $ci | Should Match 'VCPKG_INSTALLED_DIR:.*\\vcpkg_installed\\\$\{\{ inputs\.platform'
        $ci | Should Match '--x-install-root \$env:VCPKG_INSTALLED_DIR'
        $ci | Should Match '\$vcpkgRoot\s*=\s*Join-Path \$env:VCPKG_INSTALLED_DIR \$triplet'
        $ci | Should Match "hashFiles\('vcpkg\.json', 'vcpkg-tool\.json'\)"
        $ci | Should Match 'id:\s*toolchain_identity'
        $ci | Should Match 'Get-RSBuildToolchainIdentity'
        @($ci | Select-String -Pattern 'steps\.toolchain_identity\.outputs\.digest' -AllMatches).Matches.Count | Should Be 2
        $ci | Should Not Match 'Record exact resolved toolchain identity'
        $ci | Should Not Match 'compilerBv'
        $ci | Should Match 'git checkout --detach \$toolCommit'
        $ci | Should Match '\$actualToolCommit -ne \$toolCommit'
        $ci | Should Match '\$toolMetadata = git show "\$\{toolCommit\}:scripts/vcpkg-tools\.json" \| ConvertFrom-Json'
        $ci | Should Match '\$baselineMetadata = git show "\$\{baseline\}:scripts/vcpkg-tools\.json" \| ConvertFrom-Json'
        $ci | Should Match '\$toolSchemaVersion -lt \$baselineSchemaVersion'
        $ci | Should Match 'Update vcpkg-tool\.json with the baseline'
        $ci.IndexOf('$toolSchemaVersion -lt $baselineSchemaVersion') | Should BeLessThan $ci.IndexOf('git checkout --detach $toolCommit')
        $ci | Should Not Match 'vcpkg(?:\.exe)?\s+integrate\s+install'
        $ci | Should Match '\$env:ForceImportBeforeCppTargets\s*=\s*\$vcpkgProps'
        $ci | Should Match '\$env:ForceImportAfterCppTargets\s*=\s*\$vcpkgTargets'
        $ci | Should Match '\$isDebugConfiguration\s*=\s*"\$\{\{ inputs\.configuration \}\}"\s+-like\s+"\*Debug\*"'
        $ci | Should Match '\$env:VcpkgConfiguration\s*=\s*\$vcpkgConfiguration'
        $ci | Should Match 'if \(\$isDebugConfiguration\) \{ "debug\\lib" \} else \{ "lib" \}'
        $ci | Should Match 'if \(\$isDebugConfiguration\) \{ "debug\\bin" \} else \{ "bin" \}'
        $ci | Should Match '\$env:VcpkgManifestInstall\s*=\s*"false"'
        $ci | Should Match '\$env:VcpkgXUseBuiltInApplocalDeps\s*=\s*"false"'
        $ci | Should Match '\$env:VCPkgLocalAppDataDisabled\s*=\s*"true"'
        $ci | Should Match '\$vcpkgMsbuildRoot\s*=\s*Join-Path \$env:VCPKG_ROOT "scripts\\buildsystems\\msbuild"'
        $ci | Should Match '\$vcpkgProps\s*=\s*Join-Path \$vcpkgMsbuildRoot "vcpkg\.props"'
        $ci | Should Match '\$vcpkgTargets\s*=\s*Join-Path \$vcpkgMsbuildRoot "vcpkg\.targets"'
        $ci | Should Match 'VCPKG_ROOT:.*\\\.build\\vcpkg-tool'
        $ci | Should Match '\.build/vcpkg-tool/downloads'
        $ci | Should Match 'Resolve-RSVcpkgSafeChildPath -Root \$env:GITHUB_WORKSPACE -Child ''\.build\\vcpkg-tool'''
        $ci | Should Not Match 'git clone \$toolRepository \.\\\\vcpkg'
        $ci.IndexOf('$env:ForceImportBeforeCppTargets') | Should BeLessThan $ci.IndexOf('& $buildScript @buildArgs')
        $ci.IndexOf('$env:ForceImportAfterCppTargets') | Should BeLessThan $ci.IndexOf('& $buildScript @buildArgs')
        $ci.IndexOf('$env:VcpkgXUseBuiltInApplocalDeps') | Should BeLessThan $ci.IndexOf('& $buildScript @buildArgs')
        $ci.IndexOf('Resolve canonical toolchain identity') | Should BeLessThan $ci.IndexOf('Cache vcpkg')

        # Restore and save are separate so a successful dependency build is cached even when the
        # build or tests fail; the combined action saved only on job success and left red runs
        # rebuilding vcpkg and the Terminal runtime every time.
        $ci | Should Match 'id:\s*vcpkg_cache\s*\r?\n\s+uses:\s*actions/cache/restore@'
        $ci | Should Match 'id:\s*terminal_runtime_cache\s*\r?\n\s+uses:\s*actions/cache/restore@'
        $ci | Should Not Match 'uses:\s*actions/cache@'
        $ci | Should Match "(?ms)- name: Save vcpkg cache\s+if: steps\.vcpkg_cache\.outputs\.cache-hit != 'true'\s+uses: actions/cache/save@.+?key: \`$\{\{ steps\.vcpkg_cache\.outputs\.cache-primary-key \}\}"
        $ci | Should Match "(?ms)- name: Save pinned Terminal runtime cache\s+if: steps\.terminal_runtime_cache\.outputs\.cache-hit != 'true'\s+uses: actions/cache/save@.+?key: \`$\{\{ steps\.terminal_runtime_cache\.outputs\.cache-primary-key \}\}"
        $ci.IndexOf('Normalize vcpkg install root') | Should BeLessThan $ci.IndexOf('Save vcpkg cache')
        $ci.IndexOf('Save vcpkg cache') | Should BeLessThan $ci.IndexOf('- name: Build solution')
        $ci.IndexOf('- name: Build solution') | Should BeLessThan $ci.IndexOf('Save pinned Terminal runtime cache')
        $ci.IndexOf('Save pinned Terminal runtime cache') | Should BeLessThan $ci.IndexOf('Run native validation suite')

        $buildScript = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $buildScript | Should Match '(?s)Get-RSBuildIdentityBundle.+?-Platform\s+\$Platform\s+-Configuration\s+\$Configuration'
        $buildScript | Should Match '(?s)GITHUB_ACTIONS.+?MSBUILD_EXE_PATH.+?GetFullPath.+?return\s+@\{.+?Method\s*=\s*"MSBUILD_EXE_PATH"'
        $buildScript | Should Match 'Publish-RSBuildToolchainIdentity'
        $buildScript | Should Match 'ToolchainIdentity\.schema\.json'
        $buildScript | Should Match '(?s)New-RSBuildArtifactRecord.+?-LiteralPath\s+\$toolchainIdentityPath'
    }

    It 'keeps ARM64 compilation and critical contract suites in the PR gate' {
        $ci = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\ci.yml') -Raw
        $subclassGuard = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Verify-NoSubclassManager.ps1') -Raw
        $selfTests = Get-Content -LiteralPath (Join-Path $repoRoot 'Specs\Testing\Testing_SelfTests.md') -Raw
        # The run title names the suite, the profiles, and the trigger instead of the PR title.
        $ci | Should Match "run-name: >-\s*\r?\n\s*\$\{\{ github\.event_name == 'pull_request'\s*\r?\n\s*&& 'Suite PR on x64 Debug \+ ARM64 Debug \(pull request\)'"
        $ci | Should Match "\|\| format\('Suite PR on x64 \+ ARM64, Debug \+ Release \(\{0\} \{1\}\)', github\.event_name == 'push' && 'push to' \|\| 'dispatch on', github\.ref_name\)"
        $ci | Should Match 'push:\s*\r?\n\s*branches:\s*\[main, master\]'
        $ci | Should Match 'pull_request:\s*\r?\n\s*branches:\s*\[main, master\]'
        $ci | Should Match 'contents:\s*read'
        $ci | Should Not Match 'contents:\s*write|git commit|git push|format-all\.ps1'
        $ci | Should Match 'git diff --name-only --diff-filter=ACMR'
        $ci | Should Match '--require-hashes --only-binary=:all: --no-deps'
        $ci | Should Match '--dry-run --Werror --style=file'
        $ci | Should Match '\^\(Common\|Plugins\|PoC\|Red\[\^/\]\+\|Tests\|Tools\)'
        $subclassGuard | Should Match "Get-Command 'rg' -ErrorAction SilentlyContinue"
        $subclassGuard | Should Match 'Select-String -SimpleMatch -Pattern \$pattern'
        $subclassGuard | Should Match '\$global:LASTEXITCODE\s*=\s*0\s*$'
        $ci | Should Match '"platform": "ARM64", "runner": "windows-11-vs2026-arm", "configuration": "Debug"'
        $selfTests | Should Match 'Suite PR is the ten-minute subset of Suite CI: .*PluginContractTests, SettingsSchemaTests, CrashHandlingTests'
        $ci | Should Match 'run_full_tests:\s*true'
        # Pull requests and pushes to main run Suite PR; Suite CI runs nightly from nightly-ci.yml;
        # Full stays the local closeout gate.
        $ci | Should Match 'test_suite:\s*PR\s*\r?\n'
        $ci | Should Not Match 'test_suite:\s*\$\{\{'
        $nightly = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\nightly-ci.yml') -Raw
        $nightly | Should Match 'schedule:\s*\r?\n\s*- cron: "47 2 \* \* \*"'
        $nightly | Should Match 'workflow_dispatch:'
        $nightly | Should Match "run-name: >-\s*\r?\n\s*\$\{\{ format\('Suite CI on x64 \+ ARM64, Debug \+ Release \(\{0\} \{1\}\)', github\.event_name == 'schedule' && 'nightly on' \|\| 'dispatch on', github\.ref_name\) \}\}"
        $nightly | Should Not Match 'pull_request:|push:'
        $nightly | Should Match 'test_suite:\s*CI\s*\r?\n'
        $nightly | Should Match 'run_full_tests:\s*true'
        $nightly | Should Not Match '"configuration": "ASan Debug"|configuration: ASan Debug'
        $nightly | Should Match 'workflows/nightly-ci\.yml/runs\?branch='
        # The run-list call needs the Actions read scope; an API failure runs the suite instead of skipping it.
        $nightly | Should Match '(?ms)^  changes:\s*\r?\n(?:.*?\r?\n)*?    permissions:\s*\r?\n\s*actions: read\s*\r?\n\s*contents: read'
        $nightly | Should Match 'if ! last=\$\(gh api'
        $nightly | Should Match '::warning::Could not list the completed nightly runs'
        $nightly | Should Match "if: needs\.changes\.outputs\.run == 'true'"
        foreach ($profile in @('x64, runner: windows-2025-vs2026, configuration: Debug',
                'x64, runner: windows-2025-vs2026, configuration: Release',
                'ARM64, runner: windows-11-vs2026-arm, configuration: Debug',
                'ARM64, runner: windows-11-vs2026-arm, configuration: Release')) {
            $nightly | Should Match ([regex]::Escape("- { platform: $profile }"))
        }
        $nightly | Should Match '(?ms)^concurrency:\s+group:\s*nightly-ci-\$\{\{ github\.ref \}\}\s+cancel-in-progress:\s*false'
        $ci | Should Not Match '"configuration": "ASan Debug"'
        $ci | Should Match '(?ms)^concurrency:\s+group:\s*ci-\$\{\{ github\.event\.pull_request\.number \|\| github\.ref \}\}\s+cancel-in-progress:\s*true'
    }

    It 'runs scheduled and high-risk ASan with a seeded detector proof before green contracts' {
        $asan = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\asan.yml') -Raw
        $reusable = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw
        $harness = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\PluginContractTests\PluginContractTests.cpp') -Raw
        $asan | Should Match "run-name: >-\s*\r?\n\s*\$\{\{ github\.event_name == 'pull_request'\s*\r?\n\s*&& 'ASan Debug, Suite PR on x64 \(pull request\)'"
        $asan | Should Match "\|\| format\('ASan Debug, Suite CI on x64 \+ ARM64 \(\{0\}\)', github\.event_name == 'schedule' && 'weekly' \|\| 'dispatch'\)"
        $asan | Should Match 'schedule:'
        $asan | Should Match 'pull_request:\s*\r?\n\s*branches:\s*\[main, master\]'
        $asan | Should Match '\*\*/\*\.cpp'
        $asan | Should Match 'configuration:\s*ASan Debug'
        $asan | Should Match 'run_asan_plugin_contracts:\s*true'
        $asan | Should Match 'run_full_tests:\s*true'
        # Pull requests run the ten-minute Suite PR on x64 only; schedule and dispatch run Suite CI on both.
        $asan | Should Match "test_suite: \$\{\{ github\.event_name == 'pull_request' && 'PR' \|\| 'CI' \}\}"
        $asan | Should Match "github\.event_name == 'pull_request'\s*\r?\n\s*&& '\[\{`"platform`": `"x64`", `"runner`": `"windows-2025-vs2026`"\}\]'"
        $asan | Should Match '"platform": "ARM64", "runner": "windows-11-vs2026-arm"'
        $reusable | Should Match '--asan-seed-heap-overflow'
        $reusable | Should Match 'not rejected by AddressSanitizer'
        $reusable | Should Match 'AddressSanitizer diagnostic'
        $harness | Should Match 'RS_ASAN_DEBUG_BUILD'
        $harness | Should Match 'raw\[16\]\s*=\s*0x5Au'
    }
}
