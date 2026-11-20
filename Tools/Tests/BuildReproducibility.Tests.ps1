Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools\Modules\Packaging\RuntimeDependencies.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Packaging\PortablePackageSmoke.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Build\VcpkgToolIdentity.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\TestSandbox.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Tests\TestSupport.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Build\SanitizedEnvironment.psm1') -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Build\ProcessStreaming.psm1') -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Build\MSBuildInvocation.psm1') -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Build\BuildEvidence.psm1') -ErrorAction Stop

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
        'SettingsStore.schema.json',
        'RedMonitor-LICENSE.txt',
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
        New-Item -ItemType Directory -Path (Split-Path (Join-Path $Root $relativePath)) -Force | Out-Null
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

    It 'declares vcpkg link inputs and stages the previously implicit runtime DLLs' {
        $props = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.props') -Raw
        $targets = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.targets') -Raw
        $linkInputs = Get-Content -LiteralPath (Join-Path $repoRoot 'Build/VcpkgLinkDependencies.props') -Raw
        $props | Should Match '<VcpkgEnabled>false</VcpkgEnabled>'
        $props | Should Match '<VcpkgAutoLink>false</VcpkgAutoLink>'
        $targets | Should Match 'Build\\VcpkgLinkDependencies\.props'
        $linkInputs | Should Not Match '\*\.lib|RedLauncher'
        $linkInputs | Should Match 'libcurl-d\.lib'
        $linkInputs | Should Match 'raw_rd\.lib'

        foreach ($platform in 'x64', 'ARM64') {
            foreach ($configuration in 'Debug', 'ASan Debug', 'Release') {
                $manifest = Get-RSRuntimeDependencyManifest -RepoRoot $repoRoot -Configuration $configuration -Platform $platform
                foreach ($dependency in @(
                    @{ Project = 'Common'; Dll = 'blake3.dll'; Root = 'BuildOutput' },
                    @{ Project = 'Common'; Dll = 'yyjson.dll'; Root = 'BuildOutput' },
                    @{ Project = 'ViewerSqlite'; Dll = 'sqlite3.dll'; Root = 'Plugins' },
                    @{ Project = 'ViewerSqlite'; Dll = 'yyjson.dll'; Root = 'Plugins' },
                    @{ Project = 'ViewerPE'; Dll = 'pe-parse.dll'; Root = 'Plugins' },
                    @{ Project = 'ViewerImgRaw'; Dll = 'spng.dll'; Root = 'Plugins' },
                    @{ Project = 'FileSystemGoogleDrive'; Dll = 'libssh2.dll'; Root = 'Plugins' }
                )) {
                    @($manifest.Dependencies | Where-Object {
                        $_.Projects -contains $dependency.Project -and
                        $_.OutputName -ceq $dependency.Dll -and
                        $_.OutputRoot -ceq $dependency.Root -and $_.Required
                    }).Count | Should Be 1
                }
                @($manifest.Dependencies | Where-Object { $_.Projects -contains 'RedLauncher' }).Count | Should Be 0
            }
        }
    }

    It 'honors test-harness definitions in Release and permits bounded build parallelism' {
        $props = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.props') -Raw
        $targets = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.targets') -Raw
        $testTargets = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\Directory.Build.targets') -Raw
        $buildScript = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
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
        $targets | Should Not Match 'vcpkg_tlog='
        $targets | Should Match 'runtime=\$\(TargetDir\)\$\(RSAsanRuntimeFileName\)'
        $targets | Should Match "EnableASAN.*ConfigurationType.*Application.*RSAsanRuntimeSourcePath"
        $targets | Should Match '_RSLibraryPathForBuildEnvironment'
        $targets | Should Match 'SetBuildDefaultEnvironmentVariablesDependsOn'
        $targets | Should Match 'Condition="Exists\(''%\(Identity\)''\)"'
        $targets | Should Match '<ItemDefinitionGroup Condition="[^\"]+RSBuildCompilerDebugInformation[^\"]+Embedded'
        $targets | Should Match '<DebugInformationFormat>OldStyle</DebugInformationFormat>'
        $targets | Should Match '<ClCompile Update="@\(ClCompile\)">'
        $targets | Should Match '_RSEmbeddedSearchTextPath'
        $targets | Should Match 'WithMetadataValue\(''FullPath'''
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
            'TerminalRuntime\LICENSE-Aro.txt'
            'TerminalRuntime\LICENSE-Pixels.txt'
            'TerminalRuntime\LICENSE-uucode.txt'
            'TerminalRuntime\LICENSE-uucode-Bjoern-Hoehrmann.txt'
            'TerminalRuntime\LICENSE-Unicode.txt'
            'TerminalRuntime\LICENSE-Wuffs.txt'
            'TerminalRuntime\LICENSE-Zig.txt'
        )
        foreach ($platform in @('x64', 'ARM64')) {
            $manifest = Get-RSRuntimeDependencyManifest -RepoRoot $repoRoot -Configuration Debug -Platform $platform
            $terminalDependencies = @($manifest.Dependencies | Where-Object {
                $_.Projects -contains 'Terminal' -and $_.OutputName -like 'TerminalRuntime\*'
            })
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
        $result.RequiredFileCount | Should Be 20
        $result.GhosttyRuntime.Machine | Should Be 'AMD64'
        Assert-MockCalled -CommandName Test-RSPortableGhosttyRuntime -ModuleName PortablePackageSmoke -Times 1 -Exactly -Scope It
    }

    It 'makes fresh-extraction app and plugin smoke part of portable packaging' {
        $packageScript = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\zip\build-zip.ps1') -Raw
        $releaseWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1') -Raw
        $testTargets = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\Directory.Build.targets') -Raw
        $pluginContractHarness = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\PluginContractTests\PluginContract.Tests.Runner.cpp') -Raw
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
        $releaseWorkflow | Should Match 'Installer[\\/]zip[\\/]build-zip\.ps1'
    }

    It 'publishes portable ZIPs through guarded same-directory atomic replacement' {
        $packageScript = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\zip\build-zip.ps1') -Raw
        $packageScript | Should Match '(?s)Resolve-RSGuardedRepositoryPath.+?-GuardedRelativeRoot\s+''\.build\\AppPackages''.+?-CreateDirectory'
        $packageScript | Should Match '(?s)\$StagedZipPath\s*=\s*Join-Path\s+\$PackageOutputDir.+?\.tmp\.zip.+?NewGuid'
        $packageScript | Should Match '(?s)Compress-Archive.+?-DestinationPath\s+\$StagedZipPath.+?Test-RSPortablePackage.+?-ZipPath\s+\$StagedZipPath'
        $packageScript | Should Match '(?s)Publish-RSGuardedRepositoryFile.+?-StagedPath\s+\$StagedZipPath.+?-DestinationPath\s+\$ZipPath'
        $packageScript | Should Not Match '(?m)^\s*Compress-Archive[^\r\n]+-DestinationPath\s+\$ZipPath'
    }

    It 'runs app and plugin extraction probes through contained process support' {
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

function New-RSFormatCommandFixture {
    <# A synthetic checkout with the shared engine's real policy module and project schema. Only
       Install-CiFormatter is a stub that returns the owned formatter probe, so nothing is downloaded. #>
    param([Parameter(Mandatory = $true)][string]$RepositoryRoot, [Parameter(Mandatory = $true)][string]$Root)

    $engine = Join-Path $Root '.ci-engine'
    foreach ($directory in @('.ci', 'Owned/Nested', 'OwnedSibling', 'Outside', 'Tools/Sealed', '.ci-engine/Tools/Modules', '.ci-engine/Schemas')) {
        [void][IO.Directory]::CreateDirectory((Join-Path $Root $directory))
    }
    Copy-Item -LiteralPath (Join-Path $RepositoryRoot '.ci-engine/Tools/Modules/Ci.Policy.psm1') -Destination (Join-Path $engine 'Tools/Modules/Ci.Policy.psm1')
    Copy-Item -LiteralPath (Join-Path $RepositoryRoot '.ci-engine/Schemas/project.schema.json') -Destination (Join-Path $engine 'Schemas/project.schema.json')
    [IO.File]::WriteAllText((Join-Path $engine 'Tools/Modules/Ci.Format.psm1'), @'
Set-StrictMode -Version Latest
function Install-CiFormatter {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)
    [IO.File]::AppendAllText($env:RS_FORMATTER_INSTALL_LOG, $Root + [Environment]::NewLine)
    return $env:RS_FORMATTER_EXECUTABLE
}
Export-ModuleMember -Function Install-CiFormatter
'@)

    foreach ($repository in @($engine, $Root)) {
        [void](Invoke-RSTestFixtureGit -Arguments @('-C'; $repository; 'init'; '-q'; '-b'; 'main') 2>&1)
        if ($LASTEXITCODE -ne 0) { throw "Fixture Git failed with exit $LASTEXITCODE." }
        foreach ($setting in @(@('user.name', 'Formatter fixture'), @('user.email', 'format@fixture.invalid'),
                @('core.hooksPath', (Join-Path $Root '.no-hooks')), @('core.autocrlf', 'false'), @('commit.gpgsign', 'false'))) {
            [void](Invoke-RSTestFixtureGit -Arguments @('-C'; $repository; 'config'; $setting) 2>&1)
            if ($LASTEXITCODE -ne 0) { throw "Fixture Git failed with exit $LASTEXITCODE." }
        }
    }
    [void](Invoke-RSTestFixtureGit -Arguments @('-C'; $engine; 'add'; '--all') 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Fixture Git failed with exit $LASTEXITCODE." }
    [void](Invoke-RSTestFixtureGit -Arguments @('-C'; $engine; 'commit'; '-q'; '-m'; 'Fixture engine') 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Fixture Git failed with exit $LASTEXITCODE." }
    $engineHead = ([string]@(Invoke-RSTestFixtureGit -Arguments @('-C'; $engine; 'rev-parse'; 'HEAD'))[0]).Trim()
    if ($LASTEXITCODE -ne 0) { throw "Fixture Git failed with exit $LASTEXITCODE." }
    if ($engineHead -notmatch '^[0-9a-f]{40}$') { throw 'Fixture engine commit is not a full SHA-1.' }

    $project = Get-Content -LiteralPath (Join-Path $RepositoryRoot '.ci/project.json') -Raw | ConvertFrom-Json -AsHashtable
    $project.engine.ref = $engineHead
    $project.formatRoots = @('Owned', 'Tools')
    $project.formatExclude = @('Tools/Sealed/**')
    $projectText = $project | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText((Join-Path $Root '.ci/project.json'), $projectText)

    $marker = '/* fixture: needs formatting */'
    $files = [ordered]@{
        '.gitignore' = "/.ci-engine/`n*.generated.h`n"
        '.clang-format' = "BasedOnStyle: LLVM`n"
        'Owned/sample with spaces.cpp' = "int owned(); $marker`n"
        'Owned/Nested/Header.HPP' = "int header();`n"
        'Owned/notes.txt' = "$marker`n"
        'Tools/current.cc' = "int tool(); $marker`n"
        'Tools/Sealed/frozen.cpp' = "int sealed(); $marker`n"
        'OwnedSibling/prefix.cpp' = "int sibling(); $marker`n"
        'Outside/other.cpp' = "int outside(); $marker`n"
    }
    foreach ($file in $files.GetEnumerator()) { [IO.File]::WriteAllText((Join-Path $Root $file.Key), $file.Value) }
    foreach ($script in @('format.ps1', 'format-all.ps1')) { Copy-Item -LiteralPath (Join-Path $RepositoryRoot $script) -Destination (Join-Path $Root $script) }
    [void](Invoke-RSTestFixtureGit -Arguments @('-C'; $Root; 'add'; '--all') 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Fixture Git failed with exit $LASTEXITCODE." }
    # Added after staging: one new file CI formats and one file Git ignores.
    [IO.File]::WriteAllText((Join-Path $Root 'Owned/new.hxx'), "int added();`n")
    [IO.File]::WriteAllText((Join-Path $Root 'Owned/build.generated.h'), "int generated(); $marker`n")

    [pscustomobject]@{
        Root = $Root
        EngineHead = $engineHead
        ProjectText = $projectText
        Marker = $marker
        Selected = 'Owned/Nested/Header.HPP|Owned/new.hxx|Owned/sample with spaces.cpp|Tools/current.cc'
    }
}

function Invoke-RSFormatFixtureCommand {
    <# Runs one fixture command in a fresh PowerShell from a foreign directory. The bootstrap prints a
       terminating error's message unwrapped, so assertions do not depend on the host's error view. #>
    param(
        [Parameter(Mandatory = $true)][object]$Context,
        [Parameter(Mandatory = $true)][string]$Script,
        [hashtable]$Parameters = @{}
    )

    $probe = $Context.Probe
    $env:RS_FORMATTER_LOG = Join-Path $probe.Root ('args-' + [guid]::NewGuid().ToString('N') + '.jsonl')
    $env:RS_FORMATTER_INSTALL_LOG = Join-Path $probe.Root ('install-' + [guid]::NewGuid().ToString('N') + '.log')
    $auditBefore = if (Test-Path -LiteralPath $probe.AuditPath) { @(Get-Content -LiteralPath $probe.AuditPath).Count } else { 0 }
    $arguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', $Context.Bootstrap, '-Script', (Join-Path $Context.Fixture.Root $Script))
    foreach ($name in @('Check', 'WhatIf')) { if ($Parameters.ContainsKey($name)) { $arguments += "-Use$name" } }
    if ($Parameters.ContainsKey('RepositoryRoot')) { $arguments += @('-RepositoryRoot', $Parameters.RepositoryRoot) }
    $records = @(Invoke-RSTestFixtureNative -FilePath (Join-Path $PSHOME 'pwsh.exe') -Arguments $arguments -WorkingDirectory $Context.Foreign 2>&1)
    $exitCode = $LASTEXITCODE

    $invocations = @(if (Test-Path -LiteralPath $env:RS_FORMATTER_LOG -PathType Leaf) {
            foreach ($line in @(Get-Content -LiteralPath $env:RS_FORMATTER_LOG | Where-Object { $_ })) {
                $values = @($line | ConvertFrom-Json)
                $separator = [Array]::IndexOf([object[]]$values, '--')
                [pscustomobject]@{ Options = (@($values | Select-Object -First $separator) -join ' '); Files = @($values | Select-Object -Skip ($separator + 1)) }
            }
        })
    [string[]]$files = @($invocations | ForEach-Object { $_.Files })
    [Array]::Sort($files, [StringComparer]::Ordinal)
    $audit = @(if (Test-Path -LiteralPath $probe.AuditPath) {
            @(Get-Content -LiteralPath $probe.AuditPath) | Select-Object -Skip $auditBefore | ForEach-Object { $_ | ConvertFrom-Json }
        })
    [pscustomobject]@{
        ExitCode = $exitCode
        Output = (@($records | ForEach-Object { [string]$_ }) -join "`n")
        Invocations = $invocations
        Files = ($files -join '|')
        Installs = @(if (Test-Path -LiteralPath $env:RS_FORMATTER_INSTALL_LOG) { Get-Content -LiteralPath $env:RS_FORMATTER_INSTALL_LOG | Where-Object { $_ } })
        Audit = $audit
    }
}

function Get-RSFormatFixtureHashes {
    param([Parameter(Mandatory = $true)][string]$Root)
    $prefix = $Root.TrimEnd('\') + '\'
    @(Get-ChildItem -LiteralPath $Root -Recurse -File -Force | Where-Object { $_.FullName -notmatch '\\\.git\\' } | Sort-Object FullName |
        ForEach-Object { $_.FullName.Substring($prefix.Length) + '=' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join "`n"
}

Describe 'Root format command and deprecated format-all wrapper' {
    BeforeAll {
        $script:formatEnvironment = @{}
        foreach ($name in @('RS_FORMATTER_EXECUTABLE', 'RS_FORMATTER_INSTALL_LOG', 'RS_FORMATTER_LOG', 'RS_FORMATTER_UNFORMATTED',
                'RS_FORMATTER_STDOUT', 'RS_FORMATTER_EXIT', 'RS_FORMATTER_STDERR_COUNT', 'RS_FORMATTER_VERSION')) {
            $script:formatEnvironment[$name] = [Environment]::GetEnvironmentVariable($name)
            # PowerShell passes $null to a .NET string as an empty value, which would leave an empty variable.
            Remove-Item -LiteralPath "Env:$name" -ErrorAction SilentlyContinue
        }
        $case = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'fmt' -Case ([guid]::NewGuid().ToString('N'))
        $bootstrap = Join-Path $case 'invoke-format.ps1'
        [IO.File]::WriteAllText($bootstrap, @'
param([Parameter(Mandatory)][string]$Script, [switch]$UseCheck, [switch]$UseWhatIf, [string]$RepositoryRoot)
$parameters = @{}
if ($UseCheck) { $parameters.Check = $true }
if ($UseWhatIf) { $parameters.WhatIf = $true }
if ($PSBoundParameters.ContainsKey('RepositoryRoot')) { $parameters.RepositoryRoot = $RepositoryRoot }
try { & $Script @parameters } catch { [Console]::Error.WriteLine('FAILED: ' + $_.Exception.Message); exit 1 }
exit 0
'@)
        $foreign = Join-Path $case 'foreign'
        [void][IO.Directory]::CreateDirectory($foreign)
        $script:formatCommand = [pscustomobject]@{
            Fixture = (New-RSFormatCommandFixture -RepositoryRoot $repoRoot -Root (Join-Path $case 'checkout'))
            Probe = (New-RSTestFormatterFixture)
            Bootstrap = $bootstrap
            Foreign = $foreign
        }
        $env:RS_FORMATTER_EXECUTABLE = $script:formatCommand.Probe.Executable
    }

    AfterAll {
        foreach ($entry in $script:formatEnvironment.GetEnumerator()) {
            if ($null -eq $entry.Value) { Remove-Item -LiteralPath "Env:$($entry.Key)" -ErrorAction SilentlyContinue }
            else { Set-Item -LiteralPath "Env:$($entry.Key)" -Value $entry.Value }
        }
    }

    It 'refuses an engine checkout at another commit than the pin before acquiring the formatter' {
        $fixture = $script:formatCommand.Fixture
        $projectPath = Join-Path $fixture.Root '.ci/project.json'
        try {
            [IO.File]::WriteAllText($projectPath, $fixture.ProjectText.Replace($fixture.EngineHead, ('0' * 40)))
            $result = Invoke-RSFormatFixtureCommand -Context $script:formatCommand -Script 'format.ps1' -Parameters @{ Check = $true }
        } finally {
            [IO.File]::WriteAllText($projectPath, $fixture.ProjectText)
        }
        $result.ExitCode | Should Be 1
        $result.Output | Should Match ([regex]::Escape("FAILED: .ci-engine is at $($fixture.EngineHead), but .ci/project.json pins $('0' * 40)."))
        $result.Installs.Count | Should Be 0
        $result.Invocations.Count | Should Be 0
    }

    It 'refuses a root style that can inherit a style from outside the checkout' {
        $stylePath = Join-Path $script:formatCommand.Fixture.Root '.clang-format'
        try {
            [IO.File]::WriteAllText($stylePath, "BasedOnStyle: InheritParentConfig`n")
            $result = Invoke-RSFormatFixtureCommand -Context $script:formatCommand -Script 'format.ps1' -Parameters @{ Check = $true }
        } finally {
            [IO.File]::WriteAllText($stylePath, "BasedOnStyle: LLVM`n")
        }
        $result.ExitCode | Should Be 1
        $result.Output | Should Match 'FAILED: Formatting requires the repository''s own root \.clang-format, which must not inherit a style from outside the checkout\.'
        $result.Installs.Count | Should Be 0
    }

    It 'checks exactly the CI file selection without writing and lists every file that needs formatting' {
        $fixture = $script:formatCommand.Fixture
        $before = Get-RSFormatFixtureHashes -Root $fixture.Root
        $env:RS_FORMATTER_UNFORMATTED = $fixture.Marker
        try {
            $result = Invoke-RSFormatFixtureCommand -Context $script:formatCommand -Script 'format.ps1' -Parameters @{ Check = $true }
        } finally { $env:RS_FORMATTER_UNFORMATTED = $null }

        $result.ExitCode | Should Be 1
        $result.Output | Should Match 'Formatter: clang-format version 22\.1\.3 \('
        $needs = @([regex]::Matches($result.Output, '(?m)^Needs formatting: (.+?)\s*$') | ForEach-Object { $_.Groups[1].Value })
        ($needs -join '|') | Should Be 'Owned/sample with spaces.cpp|Tools/current.cc'
        $result.Output | Should Match 'FAILED: 2 of 4 files need formatting\. Run \./format\.ps1 to format them\.'
        $result.Files | Should Be $fixture.Selected
        $result.Invocations.Count | Should BeGreaterThan 0
        foreach ($invocation in $result.Invocations) { $invocation.Options | Should Be '--style=file --Werror --dry-run' }
        ($result.Installs -join '|') | Should Be $fixture.Root
        $result.Audit.Count | Should Be ($result.Invocations.Count + 1)
        @($result.Audit | Where-Object { @($_.arguments).Count -eq 1 -and $_.arguments[0] -ceq '--version' }).Count | Should Be 1
        foreach ($row in $result.Audit) {
            $row.workingDirectory | Should Be $fixture.Root
            $row.consoleVisible | Should Be $false
        }
        Get-RSFormatFixtureHashes -Root $fixture.Root | Should Be $before
    }

    It 'formats the same selection in place and then reports a clean check' {
        $formatted = Invoke-RSFormatFixtureCommand -Context $script:formatCommand -Script 'format.ps1'
        $formatted.ExitCode | Should Be 0
        $formatted.Output | Should Match '(?m)^Formatted 4 files\.\s*$'
        $formatted.Files | Should Be $script:formatCommand.Fixture.Selected
        foreach ($invocation in $formatted.Invocations) { $invocation.Options | Should Be '--style=file --Werror -i' }
        foreach ($row in $formatted.Audit) { $row.workingDirectory | Should Be $script:formatCommand.Fixture.Root }

        $checked = Invoke-RSFormatFixtureCommand -Context $script:formatCommand -Script 'format.ps1' -Parameters @{ Check = $true }
        $checked.ExitCode | Should Be 0
        $checked.Output | Should Match '(?m)^Checked 4 files\.\s*$'
        $checked.Output | Should Not Match 'Needs formatting'
    }

    It 'keeps format-all.ps1 as a deprecated wrapper that maps -WhatIf to the check and refuses another checkout' {
        $fixture = $script:formatCommand.Fixture
        (Get-Help -Name (Join-Path $repoRoot 'format-all.ps1')).Synopsis | Should Match '^Deprecated compatibility wrapper for format\.ps1\. It will be removed after one release\.'

        $env:RS_FORMATTER_UNFORMATTED = $fixture.Marker
        try {
            $preview = Invoke-RSFormatFixtureCommand -Context $script:formatCommand -Script 'format-all.ps1' -Parameters @{ WhatIf = $true }
        } finally { $env:RS_FORMATTER_UNFORMATTED = $null }
        $preview.ExitCode | Should Be 1
        $preview.Output | Should Match 'format-all\.ps1 is deprecated and will be removed after one release; use format\.ps1'
        $preview.Output | Should Match '(?m)^Needs formatting: Tools/current\.cc\s*$'
        $preview.Output | Should Match 'FAILED: 2 of 4 files need formatting\.'
        $preview.Files | Should Be $fixture.Selected
        foreach ($invocation in $preview.Invocations) { $invocation.Options | Should Be '--style=file --Werror --dry-run' }

        $refused = Invoke-RSFormatFixtureCommand -Context $script:formatCommand -Script 'format-all.ps1' -Parameters @{ RepositoryRoot = $script:formatCommand.Foreign }
        $refused.ExitCode | Should Be 1
        $refused.Output | Should Match ([regex]::Escape("FAILED: format-all.ps1 forwards to format.ps1, which formats only its own checkout. Run format.ps1 in '$($script:formatCommand.Foreign)' instead."))
        $refused.Installs.Count | Should Be 0
        $refused.Audit.Count | Should Be 0

        $formatted = Invoke-RSFormatFixtureCommand -Context $script:formatCommand -Script 'format-all.ps1' -Parameters @{ RepositoryRoot = ($fixture.Root + '\') }
        $formatted.ExitCode | Should Be 0
        $formatted.Output | Should Match '(?m)^Formatted 4 files\.\s*$'
        $formatted.Files | Should Be $fixture.Selected
        foreach ($invocation in $formatted.Invocations) { $invocation.Options | Should Be '--style=file --Werror -i' }
    }
}

Describe 'ZIP and winget packaging source contracts' {
    It 'rejects ZIP packaging with ProductionOnly before importing build modules' {
        $buildPath = Join-Path $repoRoot 'build.ps1'
        $tokens = $null
        $parseErrors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($buildPath, [ref]$tokens, [ref]$parseErrors)
        @($parseErrors).Count | Should Be 0
        $guard = @($ast.FindAll({
                    param($node)
                    $node -is [Management.Automation.Language.IfStatementAst] -and
                    $node.Extent.Text -match 'ProductionOnly cannot create a ZIP because it excludes the package smoke-test host'
                }, $true))
        $guard.Count | Should Be 1
        $firstImport = @($ast.FindAll({
                    param($node)
                    $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Import-Module'
                }, $true) | Sort-Object { $_.Extent.StartOffset } | Select-Object -First 1)
        $firstImport.Count | Should Be 1
        ($guard[0].Extent.EndOffset -lt $firstImport[0].Extent.StartOffset) | Should Be $true

        $ProductionOnly = $true
        $Zip = $true
        { . ([scriptblock]::Create($guard[0].Extent.Text)) } |
            Should Throw 'ProductionOnly cannot create a ZIP because it excludes the package smoke-test host'

        $Zip = $false
        $allowed = Invoke-RSTestBlock { . ([scriptblock]::Create($guard[0].Extent.Text)) }
        $allowed.Succeeded | Should Be $true
        $ProductionOnly = $false
        $Zip = $true
        $allowed = Invoke-RSTestBlock { . ([scriptblock]::Create($guard[0].Extent.Text)) }
        $allowed.Succeeded | Should Be $true
    }

    It 'uses explicit synchronized versions for portable ZIP and winget publication' {
        $zip = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer/zip/build-zip.ps1') -Raw
        $zip | Should Match '\$BuildNumber\s+-le\s+0'
        $zip | Should Match 'New-RSVersionContext'
        $zip | Should Not Match 'Resolve-RSVersionContext|Read-RSVersionContext'

        $winget = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer/winget/generate-manifest.ps1') -Raw
        $winget | Should Match '(?s)IsNullOrWhiteSpace\(\$Version\).+?\$BuildNumber\s+-le\s+0'
        $winget | Should Match 'New-RSVersionContext'
        $winget | Should Not Match 'Resolve-RSVersionContext|Read-RSVersionContext'

        $buildScript = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $resolver = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/ResolveVersionForMsbuild.ps1') -Raw
        $buildScript | Should Match 'Resolve-RSVersionContext'
        $resolver | Should Match 'Resolve-RSVersionContext'
        $releaseWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1') -Raw
        $releaseWorkflow | Should Match 'context\.batch\.buildNumber'
        $releaseWorkflow | Should Not Match 'Get-RSVersionContext'
        $buildScript | Should Not Match 'Use-RSVersionStateLock|Save-RSVersionContext'
        $resolver | Should Not Match 'Use-RSVersionStateLock|Save-RSVersionContext'
    }

    It 'keeps the ZIP and winget mutations inside balanced packaging coordination' {
        $contracts = @(
            @{ Path = 'Installer/zip/build-zip.ps1'; ReadsProfile = $true },
            @{ Path = 'Installer/winget/generate-manifest.ps1'; ReadsProfile = $false }
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
                $profileEnter = $source.IndexOf('$ArtifactProfileLock = Enter-RSArtifactOperationLock', [StringComparison]::OrdinalIgnoreCase)
                $packagingEnter = $source.IndexOf('$PackagingLock = Enter-RSArtifactOperationLock', [StringComparison]::OrdinalIgnoreCase)
                $packagingExit = $source.LastIndexOf('Exit-RSArtifactOperationLock -Lock $PackagingLock', [StringComparison]::OrdinalIgnoreCase)
                $profileExit = $source.LastIndexOf('Exit-RSArtifactOperationLock -Lock $ArtifactProfileLock', [StringComparison]::OrdinalIgnoreCase)
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
        }

        $guardModule = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Build/ArtifactOperationLock.psm1') -Raw
        $guardModule | Should Match 'function\s+Resolve-RSGuardedRepositoryPath'
        $guardModule | Should Match 'FileAttributes\]::ReparsePoint'
    }

    It 'validates exact winget versions and portable archive identities before hashing' {
        $scriptPath = Join-Path $repoRoot 'Installer/winget/generate-manifest.ps1'
        $tokens = $null
        $parseErrors = $null
        $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
        $parseErrors.Count | Should Be 0
        foreach ($functionName in @('Assert-WingetManifestVersion', 'Resolve-WingetPortableZip')) {
            $functionAst = $scriptAst.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
            }, $true)
            ($null -ne $functionAst) | Should Be $true
            Invoke-Expression $functionAst.Extent.Text
        }

        { Assert-WingetManifestVersion -CandidateVersion '7.0' } | Should Throw 'exactly three numeric components'
        { Assert-WingetManifestVersion -CandidateVersion '7.0.0' } | Should Throw 'build component must be positive'
        { Assert-WingetManifestVersion -CandidateVersion '65536.0.1' } | Should Throw 'range 0..65535'
        Assert-WingetManifestVersion -CandidateVersion '7.0.184'

        $expectedLeaf = 'RedSalamander-7.0.184-x64-Portable.zip'
        $zipPath = Join-Path $TestDrive $expectedLeaf
        [IO.File]::WriteAllBytes($zipPath, [byte[]](0x52, 0x53))
        (Resolve-WingetPortableZip -CandidatePath $zipPath -ExpectedLeaf $expectedLeaf -Label 'x64') | Should Be $zipPath
        { Resolve-WingetPortableZip -CandidatePath $zipPath -ExpectedLeaf 'RedSalamander-7.0.183-x64-Portable.zip' -Label 'x64' } |
            Should Throw 'leaf must be exactly'
        { Resolve-WingetPortableZip -CandidatePath (Join-Path $TestDrive 'missing.zip') -ExpectedLeaf 'missing.zip' -Label 'ARM64' } |
            Should Throw 'must be an existing file'
    }
}

$compilerPolicyCases = @(foreach ($platform in @('x64', 'ARM64')) {
    foreach ($configuration in @('Debug', 'Release', 'ASan Debug')) {
        foreach ($mode in @('ProgramDatabase', 'Embedded')) {
            @{ platform=$platform; configuration=$configuration; mode=$mode }
        }
    }
})

Describe 'Effective invocation-wide compiler metadata' -Tag RequiresBuildToolchain {
    BeforeAll {
        $script:compilerPolicyRepo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        Import-Module (Join-Path $compilerPolicyRepo 'Tools/Modules/Testing/TestSandbox.psm1') -ErrorAction Stop
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installation = (& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath).Trim()
        $hostFolder = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64') { 'arm64' } else { 'amd64' }
        $script:compilerPolicyMSBuild = Join-Path $installation "MSBuild/Current/Bin/$hostFolder/MSBuild.exe"
    }

    It 'evaluates the toolset guard for VC++ projects only' {
        $scratch = New-RSTestSandboxScratchDirectory -RepoRoot $compilerPolicyRepo -Harness 'MsvcVersionGuard' -Case ([guid]::NewGuid().ToString('N'))
        $targetsPath = [IO.Path]::GetFullPath((Join-Path $compilerPolicyRepo 'Directory.Build.targets'))
        $escapedTargetsPath = [Security.SecurityElement]::Escape($targetsPath)

        function Invoke-ToolsetGuardProbe([string]$Extension, [string]$ToolsetVersion) {
            $projectPath = Join-Path $scratch ('guard-probe' + $Extension)
            $projectContents = @"
<Project>
  <PropertyGroup>
    <RSIsFirstPartyProject>true</RSIsFirstPartyProject>
    <RequiredMsvcToolsVersion>14.52</RequiredMsvcToolsVersion>
    <VCToolsVersion>$ToolsetVersion</VCToolsVersion>
  </PropertyGroup>
  <Import Project="$escapedTargetsPath" />
</Project>
"@
            [IO.File]::WriteAllText($projectPath, $projectContents, [Text.UTF8Encoding]::new($false))
            $output = @(& $compilerPolicyMSBuild $projectPath '/nologo' '/noautoresponse' '/nr:false' '/t:ValidateMsvcToolsVersion' 2>&1)
            return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = ($output -join "`n") }
        }

        $wrongVcx = Invoke-ToolsetGuardProbe '.vcxproj' '14.51.37025'
        $wrongVcx.ExitCode | Should Not Be 0
        $wrongVcx.Output | Should Match 'requires MSVC build tools 14\.52\.x'
        $matchingVcx = Invoke-ToolsetGuardProbe '.vcxproj' '14.52.37025'
        $matchingVcx.ExitCode | Should Be 0
    }

    foreach ($case in $compilerPolicyCases) {
        It "evaluates $($case.platform) $($case.configuration) $($case.mode) after every project import" -TestCases @($case) {
            param($platform, $configuration, $mode)
            $scratch = New-RSTestSandboxScratchDirectory -RepoRoot $compilerPolicyRepo -Harness 'EffectiveCompilerPolicy' -Case ([guid]::NewGuid().ToString('N'))
            $expected = if ($mode -eq 'Embedded') { 'OldStyle' } elseif ($configuration -eq 'Debug') { 'EditAndContinue' } else { 'ProgramDatabase' }
            $sharedPath = [IO.Path]::GetFullPath((Join-Path $compilerPolicyRepo 'Common/SearchTextHelpers.cpp'))
            $sharedCount = 0
            foreach ($project in @('Common/Common/Common.vcxproj', 'Plugins/Terminal/Terminal.vcxproj', 'RedSalamander/RedSalamander.vcxproj', 'Tests/PluginContractTests/PluginContractTests.vcxproj')) {
                $output = Join-Path $scratch (([IO.Path]::GetFileNameWithoutExtension($project)) + '.json')
                & $compilerPolicyMSBuild (Join-Path $compilerPolicyRepo $project) /nologo /noautoresponse /nr:false `
                    "/p:Configuration=$configuration" "/p:Platform=$platform" "/p:SolutionDir=$compilerPolicyRepo\" `
                    /p:RSBuildEnableTests=true "/p:RSBuildCompilerDebugInformation=$mode" '-getItem:ClCompile' > $output
                $LASTEXITCODE | Should Be 0
                $items = @((Get-Content -LiteralPath $output -Raw | ConvertFrom-Json).Items.ClCompile)
                $items.Count | Should BeGreaterThan 0
                foreach ($item in $items) {
                    $item.DebugInformationFormat | Should Be $expected
                    if ($mode -eq 'Embedded') {
                        $item.AdditionalOptions | Should Not Match '(?i)/Zi(?:\s|$)'
                        if ($item.FullPath -ieq $sharedPath) {
                            $sharedCount++
                            $item.MultiProcessorCompilation | Should Be 'false'
                        }
                    }
                }
            }
            if ($mode -eq 'Embedded') { $sharedCount | Should BeGreaterThan 0 }
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
        $ci | Should Match 'vcpkg-install\.ps1 -Platform \$platform .*?-VcpkgRoot \$env:VCPKG_REGISTRY_ROOT -ReuseCurrent'
        $ci | Should Not Match 'Normalize vcpkg install root|--x-install-root'
        $ci | Should Match "hashFiles\('vcpkg\.json',\ 'vcpkg-configuration\.json',\ 'vcpkg-tool\.json',\ 'vcpkg-install\.ps1',\ 'Tools/Modules/Build/VcpkgInstallSafety\.psm1',\ 'Tools/Modules/Build/VcpkgToolIdentity\.psm1',\ 'Tools/Modules/Build/MSBuildInvocation\.psm1',\ 'Tools/Modules/Build/ArtifactOperationLock\.psm1',\ 'Tools/Modules/Build/ProcessStreaming\.psm1',\ 'Tools/Modules/Build/SanitizedEnvironment\.psm1',\ '\.ci/project\.json',\ 'Build/vcpkg-ports/\*\*'\)"
        $ci | Should Match 'id:\s*toolchain_identity'
        $ci | Should Match 'Get-RSBuildToolchainIdentity'
        @($ci | Select-String -Pattern 'steps\.toolchain_identity\.outputs\.digest' -AllMatches).Matches.Count | Should Be 2
        $ci | Should Not Match 'Record exact resolved toolchain identity'
        $ci | Should Not Match 'compilerBv'
        $ci | Should Match 'Initialize-RSVcpkgToolContext -RepoRoot \$env:GITHUB_WORKSPACE'
        $bootstrap = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -Raw
        # Real Git/bootstrap negative controls live in VcpkgBootstrap.Tests.ps1.
        $bootstrap | Should Match '\$toolSchema -lt \$portsSchema'
        $bootstrap.IndexOf('$toolSchema -lt $portsSchema') | Should BeLessThan $bootstrap.IndexOf("@('checkout', '--detach', " + '$commit)')
        $bootstrap | Should Match "Update the pins together"
        $ci | Should Not Match 'vcpkg(?:\.exe)?\s+integrate\s+install'
        $ci | Should Not Match 'ForceImportBeforeCppTargets|ForceImportAfterCppTargets|VcpkgXUseBuiltInApplocalDeps'
        $ci | Should Match 'Get-RSVcpkgBuildEnvironment .*?-RuntimeOnly'
        $ci | Should Match 'VCPKG_ROOT:.*\\\.build\\vcpkg-tool'
        $ci | Should Match '\.build/vcpkg-registry/downloads'
        $bootstrap | Should Match 'Resolve-RSVcpkgSafeChildPath -Root \$RepoRoot -Child ''\.build/vcpkg-tool'''
        $ci | Should Not Match 'git clone \$toolRepository \.\\\\vcpkg'
        $ci.IndexOf('Get-RSVcpkgBuildEnvironment') | Should BeLessThan $ci.IndexOf('& $buildScript @buildArgs')
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
        $buildScript | Should Match '\$workflowSelected = Get-RSWorkflowSelectedMSBuild'
        $buildScript | Should Match 'if \(\$null -ne \$workflowSelected\) \{ return \$workflowSelected \}'
        $buildScript | Should Match 'Publish-RSBuildToolchainIdentity'
        $buildScript | Should Match 'ToolchainIdentity\.schema\.json'
        $buildScript | Should Match '(?s)New-RSBuildArtifactRecord.+?-LiteralPath\s+\$toolchainIdentityPath'
    }

    It 'installs CI dependencies from a registry checkout of the manifest baseline beside the pinned executable' {
        # The tool commit's own version database predates the versions vcpkg.json requires; with one checkout vcpkg fails
        # with "no version database entry" (Build_Toolchain.md).
        $ci = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw
        $ci | Should Match 'Initialize-RSVcpkgToolContext -RepoRoot \$env:GITHUB_WORKSPACE'
        $bootstrap = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -Raw
        $bootstrap | Should Match '\.build/vcpkg-registry'
        $bootstrap | Should Match '''worktree'', ''add'', ''--detach'', \$registryRoot, \$baseline'
        $bootstrap | Should Match '''rev-parse'', ''HEAD''\)\) -cne \$baseline'
        $ci | Should Match '-VcpkgRoot \$env:VCPKG_REGISTRY_ROOT'
        $ci | Should Match '\.build/vcpkg-registry/downloads'
        $ci.IndexOf('Initialize-RSVcpkgToolContext') | Should BeLessThan $ci.IndexOf('-VcpkgRoot $env:VCPKG_REGISTRY_ROOT')
    }

    It 'keeps native ARM64 regression behind main admission and PRs compiler-free' {
        $ci = Get-Content (Join-Path $repoRoot '.github/workflows/ci.yml') -Raw
        $native = Get-Content (Join-Path $repoRoot '.ci-engine/.github/workflows/native.yml') -Raw
        $pr = Get-Content (Join-Path $repoRoot '.ci-engine/.github/workflows/pr.yml') -Raw
        $adapter = Get-Content (Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1') -Raw
        $ci | Should Match 'branches: \[main\]'
        $ci | Should Not Match 'run_full_tests|test_suite:|build-reusable'
        $pr | Should Match "CI_NATIVE_BUILDS_DISABLED: 'true'"
        $pr | Should Not Match 'setup-msvc|vcpkg-install|build.ps1'
        $native | Should Match 'platform: x64'
        $native | Should Match 'platform: ARM64'
        $native | Should Match 'runner: windows-11-vs2026-arm'
        $native | Should Match 'ReleaseTestBuild'
        $native | Should Match 'ReleaseTest -Platform'
        $native | Should Match 'needs: \[admission, test-gate\]'
        $adapter | Should Match "else \{ 'CI' \}"
        $adapter | Should Match 'Invoke-RSWithHostedTestDesktop'
        $adapter | Should Match 'QualificationScope -eq ''Full'''
        $adapter | Should Match 'Import-Module.*HostedTestDesktop'
        $nightly = Get-Content (Join-Path $repoRoot '.github/workflows/nightly-ci.yml') -Raw
        $nightly | Should Match 'workflow_dispatch:'
        $scopeInput = '(?m-i)^      scope:\r?\n        description:[^\r\n]+\r?\n        type: choice\r?\n        options: \[Full, CommandsHandoff, CommandsUiAttribution, FileOpsVolumes, ViewerSpaceCapture, R0ePaired, DesktopProbe\]\r?\n        default: Full\s*$'
        $platformInput = '(?m-i)^      platform:\r?\n        description:[^\r\n]+\r?\n        type: choice\r?\n        options: \[All, x64, ARM64\]\r?\n        default: All\s*$'
        $platformForward = '(?m-i)^      platform: \$\{\{ inputs\.platform \}\}\s*$'
        $nightly | Should Match $scopeInput
        $nightly | Should Match $platformInput
        $nightly | Should Match $platformForward
        $nightly.Replace(', CommandsUiAttribution', '') | Should Not Match $scopeInput
        $nightly.Replace('DesktopProbe]', 'DesktopProbe, UnreviewedScope]') | Should Not Match $scopeInput
        $nightly.Replace('All, x64, ARM64]', 'All, x64, ARM64, Win32]') | Should Not Match $platformInput
        $nightly.Replace('default: All', 'default: ARM64') | Should Not Match $platformInput
        $nightly.Replace('platform: '+[char]36+'{{ inputs.platform }}', 'platform: ARM64') | Should Not Match $platformForward
        $nightly | Should Match '(?m)^        default: Full\s*$'
        $nightly | Should Match '(?m)^      scope: \$\{\{ inputs\.scope \}\}\s*$'
        $nightly | Should Not Match 'schedule:|pull_request:|push:'
        @((Get-Content (Join-Path $repoRoot 'Tests/test-scopes.json') -Raw|ConvertFrom-Json).prCoverage).Count | Should Be 0
    }

    It 'refuses desktop diagnostics outside hosted Qualification before engine or sandbox work' {
        $adapterPath = Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1'
        $adapter = [IO.File]::ReadAllText($adapterPath)
        $prior = [Environment]::GetEnvironmentVariable('GITHUB_ACTIONS', 'Process')
        try {
            $env:GITHUB_ACTIONS = 'false'
            Assert-RSThrowsMatching -ScriptBlock { & $adapterPath -Phase Static -EngineRoot 'does-not-exist' -QualificationScope DesktopProbe } -Pattern 'only valid for explicit Qualification'
            Assert-RSThrowsMatching -ScriptBlock { & $adapterPath -Phase Qualification -EngineRoot 'does-not-exist' -QualificationScope DesktopProbe } -Pattern 'restricted to GitHub Actions'
        } finally { [Environment]::SetEnvironmentVariable('GITHUB_ACTIONS', $prior, 'Process') }
        $adapter.IndexOf("if (`$QualificationScope -ceq 'DesktopProbe')") | Should BeLessThan $adapter.IndexOf('function Test-PinnedDxUiMatrix')
        $adapter | Should Match "repository_verdict='NOT_EVALUATED'; desktop=\(Get-RSHostedDesktopDiagnostics\)"
    }

    It 'keeps manual ASan on both native architectures with seeded detector proof' {
        $asan = Get-Content (Join-Path $repoRoot '.github/workflows/asan.yml') -Raw
        $qualification = Get-Content (Join-Path $repoRoot '.ci-engine/.github/workflows/qualification.yml') -Raw
        $adapter = Get-Content (Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1') -Raw
        $harness = Get-Content (Join-Path $repoRoot 'Tests/PluginContractTests/PluginContract.Tests.Runner.cpp') -Raw
        $asan | Should Match 'workflow_dispatch:'
        $asan | Should Not Match 'schedule:|pull_request:|push:'
        $asan | Should Match 'configuration: ASan Debug'
        $qualification | Should Match 'default: All'
        $qualification | Should Match 'matrix: \$\{\{ fromJSON\(needs.select-platforms.outputs.matrix\) \}\}'
        # Execute the actual pinned selector; static matrix spelling is not the architecture contract.
        $body=[regex]::Match($qualification,'(?m)^      - name: Select the requested qualification architectures\r?\n[\s\S]*?^        run: \|\r?\n(?<body>(?:          [^\r\n]*\r?\n)+)').Groups['body'].Value
        $body.Length | Should BeGreaterThan 0
        $selector=($body -split '\r?\n'|ForEach-Object{$_ -replace '^          ',''}) -join "`n"
        $oldPlatform=$env:CI_QUALIFICATION_PLATFORM; $oldOutput=$env:GITHUB_OUTPUT
        try {
            foreach($selection in @('All','x64','ARM64')) {
                $env:CI_QUALIFICATION_PLATFORM=$selection
                $env:GITHUB_OUTPUT=Join-Path $TestDrive ('qualification-'+[guid]::NewGuid().ToString('N')+'.txt')
                & ([scriptblock]::Create($selector))
                $lines=@(Get-Content -LiteralPath $env:GITHUB_OUTPUT)
                $lines.Count | Should Be 1
                $matrix=($lines[0] -replace '^matrix=','')|ConvertFrom-Json
                $expected=if($selection -ceq 'All'){@('x64','ARM64')}else{@($selection)}
                (@($matrix.include.platform) -join ',') | Should Be ($expected -join ',')
                foreach($entry in $matrix.include) {
                    $entry.runner | Should Be $(if($entry.platform -ceq 'x64'){'windows-2025-vs2026'}else{'windows-11-vs2026-arm'})
                }
            }
            foreach ($invalidSelection in @('arm64','unknown')) {
                $env:CI_QUALIFICATION_PLATFORM=$invalidSelection
                { & ([scriptblock]::Create($selector)) } | Should Throw 'Qualification platform must be All, x64 or ARM64.'
            }
        } finally {
            [Environment]::SetEnvironmentVariable('CI_QUALIFICATION_PLATFORM',$oldPlatform,'Process')
            [Environment]::SetEnvironmentVariable('GITHUB_OUTPUT',$oldOutput,'Process')
        }
        $adapter | Should Match '--asan-seed-heap-overflow'
        $adapter | Should Match 'seed.exitCode -eq 0'
        $adapter | Should Match 'not rejected with an AddressSanitizer diagnostic'
        $adapter.IndexOf('--asan-seed-heap-overflow') | Should BeLessThan $adapter.IndexOf('Invoke-RSWithHostedTestDesktop -Action')
        $harness | Should Match 'RS_ASAN_DEBUG_BUILD'
        $harness | Should Match 'raw\[16\]\s*=\s*0x5Au'
    }
}

Describe 'Workflow-selected MSBuild identity' {
    It 'accepts a relocated VS18 MSBuild and rejects wrong layouts and non-MSBuild binaries' {
        $tokens = $null
        $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'build.ps1'),[ref]$tokens,[ref]$errors)
        $errors.Count | Should Be 0
        $functions = @($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Find-MSBuild'},$true))
        $functions.Count | Should Be 1
        $find = $functions[0].Body.GetScriptBlock()
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installation = @(Invoke-RSTestFixtureNative -FilePath $vswhere -Arguments @('-latest'; '-prerelease'; '-products'; '*'; '-requires'; 'Microsoft.Component.MSBuild'; '-property'; 'installationPath'))[0]
        $hostFolder = if([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64'){'arm64'}else{'amd64'}
        $real = Join-Path $installation ('MSBuild/Current/Bin/'+$hostFolder+'/MSBuild.exe')
        (Get-Item -LiteralPath $real).VersionInfo.FileMajorPart | Should BeGreaterThan 17
        $fixture = Join-Path $TestDrive ('relocated VS é日/MSBuild/Current/Bin/'+$hostFolder+'/MSBuild.exe')
        [void][IO.Directory]::CreateDirectory((Split-Path $fixture))
        Copy-Item -LiteralPath $real -Destination $fixture
        $savedActions = $env:GITHUB_ACTIONS
        $savedMsbuild = $env:MSBUILD_EXE_PATH
        try {
            $env:GITHUB_ACTIONS = 'true'
            $env:MSBUILD_EXE_PATH = $fixture
            $selected = & $find
            $selected.Path | Should Be ([IO.Path]::GetFullPath($fixture))
            $selected.Method | Should Be 'MSBUILD_EXE_PATH'
            $wrongLayout = Join-Path $TestDrive 'outside-install-layout/MSBuild.exe'
            [void][IO.Directory]::CreateDirectory((Split-Path $wrongLayout))
            Copy-Item -LiteralPath $real -Destination $wrongLayout
            $env:MSBUILD_EXE_PATH = $wrongLayout
            Assert-RSThrowsMatching { & $find } 'not an existing Visual Studio 2026 executable'
            # Keep valid version metadata but use a different binary identity.
            # Native ARM64 installs can keep task hosts/managed assemblies in the
            # parent Bin directory. Require a real VS18 counterexample there too.
            $otherCandidates = @(
                foreach ($folder in @((Split-Path $real), (Split-Path (Split-Path $real)))) {
                    foreach ($name in @('MSBuildTaskHost.exe', 'Microsoft.Build.dll')) {
                        $candidate = Join-Path $folder $name
                        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                            $identity = (Get-Item -LiteralPath $candidate).VersionInfo
                            if ($identity.FileMajorPart -ge 18 -and $identity.OriginalFilename -ine 'MSBuild.exe') { $candidate }
                        }
                    }
                }
            )
            $otherCandidates.Count | Should BeGreaterThan 0
            $other = $otherCandidates[0]
            (Get-Item -LiteralPath $other).VersionInfo.FileMajorPart | Should BeGreaterThan 17
            (Get-Item -LiteralPath $other).VersionInfo.OriginalFilename | Should Not Be 'MSBuild.exe'
            Copy-Item -LiteralPath $other -Destination $fixture -Force
            $env:MSBUILD_EXE_PATH = $fixture
            Assert-RSThrowsMatching { & $find } 'not an existing Visual Studio 2026 executable'
            $env:MSBUILD_EXE_PATH = Join-Path $TestDrive 'missing/MSBuild.exe'
            Assert-RSThrowsMatching { & $find } 'not an existing Visual Studio 2026 executable'
        } finally {
            $env:GITHUB_ACTIONS = $savedActions
            $env:MSBUILD_EXE_PATH = $savedMsbuild
        }
    }
}

Describe 'Vcpkg captured-launch cache closure' {
    BeforeAll {
        $script:vcpkgCommonInputs = @(
            'vcpkg.json'
            'vcpkg-configuration.json'
            'vcpkg-tool.json'
            'vcpkg-install.ps1'
            'Tools/Modules/Build/VcpkgInstallSafety.psm1'
            'Tools/Modules/Build/VcpkgToolIdentity.psm1'
            'Tools/Modules/Build/MSBuildInvocation.psm1'
            'Tools/Modules/Build/ArtifactOperationLock.psm1'
            'Tools/Modules/Build/ProcessStreaming.psm1'
            'Tools/Modules/Build/SanitizedEnvironment.psm1'
            '.ci/project.json'
        )
    }
    BeforeEach {
        Mock Get-RSBuildToolchainIdentity -ModuleName BuildEvidence { [pscustomobject]@{ Record=@{}; Digest=('0' * 64) } }
        Mock Get-RSGitSourcePaths -ModuleName BuildEvidence { @() }
        Mock Get-RSGhosttyRuntimeIdentityDigest -ModuleName BuildEvidence { '0' * 64 }
    }

    It 'keeps the local common inputs and workflow cache bound to the full installer import closure' {
        $tokens=$null; $errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'Tools/Modules/Build/BuildEvidence.psm1'),[ref]$tokens,[ref]$errors)
        $errors.Count | Should Be 0
        $function=$ast.Find({param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Get-RSBuildIdentityBundle'
        },$true)
        $assignment=$function.Find({param($node)
            $node -is [Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left -is [Management.Automation.Language.VariableExpressionAst] -and
            $node.Left.VariablePath.UserPath -ceq 'vcpkgInputs' -and $node.Operator -eq 'Equals'
        },$true)
        $actual=@($assignment.Right.FindAll({param($node)
            $node -is [Management.Automation.Language.StringConstantExpressionAst]
        },$true) | ForEach-Object Value)
        ($actual -join '|') | Should Be ($script:vcpkgCommonInputs -join '|')
        $workflow=[IO.File]::ReadAllText((Join-Path $repoRoot '.github/workflows/build-reusable.yml'))
        $cache=[regex]::Match($workflow,'(?m)^\s*key:.*-vcpkg-.*hashFiles\((?<inputs>[^\r\n]+?)\)')
        $cache.Success | Should Be $true
        $cacheInputs=@([regex]::Matches($cache.Groups['inputs'].Value,"'(?<path>[^']+)'") | ForEach-Object { $_.Groups['path'].Value })
        ($cacheInputs -join '|') | Should Be ((@($script:vcpkgCommonInputs)+@('Build/vcpkg-ports/**')) -join '|')
    }

    It 'changes the actual vcpkg identity when any installer launcher dependency or pin changes' {
        $fixture=Join-Path $TestDrive 'vcpkg-launch-closure'
        foreach($relative in $script:vcpkgCommonInputs) {
            $path=Join-Path $fixture $relative
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
            [IO.File]::WriteAllText($path,'initial fixture input')
        }
        $baseline=(Get-RSBuildIdentityBundle -RepoRoot $fixture -MSBuildPath 'fixture-not-executed.exe' -Platform x64 -Configuration Debug).VcpkgIdentityDigest
        foreach($relative in $script:vcpkgCommonInputs) {
            $path=Join-Path $fixture $relative
            try {
                [IO.File]::WriteAllText($path,'changed fixture input')
                $changed=(Get-RSBuildIdentityBundle -RepoRoot $fixture -MSBuildPath 'fixture-not-executed.exe' -Platform x64 -Configuration Debug).VcpkgIdentityDigest
                $changed | Should Not Be $baseline
            } finally {
                [IO.File]::WriteAllText($path,'initial fixture input')
            }
        }
        (Get-RSBuildIdentityBundle -RepoRoot $fixture -MSBuildPath 'fixture-not-executed.exe' -Platform x64 -Configuration Debug).VcpkgIdentityDigest | Should Be $baseline
        Assert-MockCalled Get-RSBuildToolchainIdentity -ModuleName BuildEvidence -Times 13 -Exactly -Scope It
    }
}
