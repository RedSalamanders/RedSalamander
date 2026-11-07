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
        $targets | Should Match 'vcpkg_tlog=\$\(TLogLocation\)\$\(ProjectName\)\.write\.1u\.tlog'
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

Describe 'Root format-all command ownership and preview' {
    It 'formats only tracked owned files with the pinned installer from a foreign working directory' {
        $fixture = Join-Path $TestDrive ('format-all-' + [guid]::NewGuid().ToString('N'))
        foreach ($directory in @('Owned', 'ThirdParty/RedMonitor/include/RedMonitor', 'Specs/TestRuns', '.codex', 'Tools', 'Dependencies', 'Build')) {
            [void][IO.Directory]::CreateDirectory((Join-Path $fixture $directory))
        }
        $sourcePath = Join-Path $fixture 'Owned/sample with spaces.cpp'
        $vendorPath = Join-Path $fixture 'ThirdParty/RedMonitor/include/RedMonitor/Client.h'
        $archivedPath = Join-Path $fixture 'Specs/TestRuns/old.cpp'
        $agentPath = Join-Path $fixture '.codex/fixture.h'
        [IO.File]::WriteAllText($sourcePath, 'int owned();' + "`n")
        [IO.File]::WriteAllText($vendorPath, '// pinned vendor bytes' + "`n")
        [IO.File]::WriteAllText($archivedPath, '// archived evidence' + "`n")
        [IO.File]::WriteAllText($agentPath, '// agent file' + "`n")
        [IO.File]::WriteAllText((Join-Path $fixture '.clang-format'), 'BasedOnStyle: LLVM' + "`n")
        [IO.File]::WriteAllText((Join-Path $fixture 'Build/clang-format.json'), '{"version":"22.1.3","wheelSha256":"' + ('0' * 64) + '"}')
        $vendorHash = (Get-FileHash -LiteralPath $vendorPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $lockJson = [ordered]@{
            files = @(
                [ordered]@{
                    path = 'include/RedMonitor/Client.h'
                    sha256 = $vendorHash
                }
            )
        } | ConvertTo-Json -Compress -Depth 4
        [IO.File]::WriteAllText((Join-Path $fixture 'Dependencies/RedMonitor.lock.json'),
            $lockJson)

        $gitOutput = @(& git -C $fixture init -q -b main 2>&1)
        $LASTEXITCODE | Should Be 0
        foreach ($setting in @(@('user.name', 'Formatter fixture'), @('user.email', 'format@fixture.invalid'),
                @('core.hooksPath', (Join-Path $fixture '.no-hooks')), @('core.autocrlf', 'false'))) {
            [void](& git -C $fixture config @setting 2>&1)
            $LASTEXITCODE | Should Be 0
        }
        [void](& git -C $fixture add --all 2>&1)
        $LASTEXITCODE | Should Be 0

        $formatterPath = Join-Path $fixture 'fake-clang-format.cmd'
        $formatterLog = Join-Path $fixture 'formatter-invocations.txt'
        [IO.File]::WriteAllText($formatterPath,
            '@echo off' + "`r`n" +
            'if "%1"=="--version" (' + "`r`n" +
            '  echo clang-format version 22.1.3' + "`r`n" +
            '  exit /b 0' + "`r`n" +
            ')' + "`r`n" +
            'echo %*>>"%RS_FORMATTER_LOG%"' + "`r`n" +
            'exit /b 0' + "`r`n")
        $installerLog = Join-Path $fixture 'installer-invocations.txt'
        $installer = Join-Path $fixture 'Tools/Install-ClangFormat.ps1'
        [IO.File]::WriteAllText($installer,
            '[IO.File]::AppendAllText($env:RS_FORMATTER_INSTALL_LOG, "called")' + "`n" +
            'Write-Output $env:RS_FORMATTER_EXECUTABLE' + "`n" +
            '$global:LASTEXITCODE = 0' + "`n")

        $oldFormatterLog = $env:RS_FORMATTER_LOG
        $oldInstallerLog = $env:RS_FORMATTER_INSTALL_LOG
        $oldFormatterExe = $env:RS_FORMATTER_EXECUTABLE
        try {
            $env:RS_FORMATTER_LOG = $formatterLog
            $env:RS_FORMATTER_INSTALL_LOG = $installerLog
            $env:RS_FORMATTER_EXECUTABLE = $formatterPath
            $foreignDirectory = Join-Path $TestDrive ('foreign-' + [guid]::NewGuid().ToString('N'))
            [void][IO.Directory]::CreateDirectory($foreignDirectory)
            Push-Location $foreignDirectory
            try {
                & (Join-Path $repoRoot 'format-all.ps1') -RepositoryRoot $fixture -WhatIf
                Test-Path -LiteralPath $installerLog | Should Be $false
                Test-Path -LiteralPath $formatterLog | Should Be $false
                (Get-FileHash -LiteralPath $vendorPath -Algorithm SHA256).Hash.ToLowerInvariant() | Should Be $vendorHash
                (Get-FileHash -LiteralPath $vendorPath -Algorithm SHA256).Hash.ToLowerInvariant() |
                    Should Be ((Get-Content -LiteralPath (Join-Path $fixture 'Dependencies/RedMonitor.lock.json') -Raw | ConvertFrom-Json).files[0].sha256)

                & (Join-Path $repoRoot 'format-all.ps1') -RepositoryRoot $fixture -Confirm:$false
                Test-Path -LiteralPath $installerLog | Should Be $true
                $invocations = Get-Content -LiteralPath $formatterLog -Raw
                $invocations | Should Match 'sample with spaces\.cpp'
                $invocations | Should Not Match 'Client\.h|TestRuns|\.codex'
                @($invocations -split '\r?\n' | Where-Object { $_ }).Count | Should Be 1
                (Get-FileHash -LiteralPath $vendorPath -Algorithm SHA256).Hash.ToLowerInvariant() | Should Be $vendorHash
            } finally {
                Pop-Location
            }
        } finally {
            $env:RS_FORMATTER_LOG = $oldFormatterLog
            $env:RS_FORMATTER_INSTALL_LOG = $oldInstallerLog
            $env:RS_FORMATTER_EXECUTABLE = $oldFormatterExe
        }
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
        # The installer splats its arguments; VcpkgInstallSafety also executes
        # the live workflow's platform, overlay and native-failure branches.
        $ci | Should Match '\$vcpkgArguments\s*=\s*@\(\s*''install'','
        $ci | Should Match '''--x-install-root'',\s*\$env:VCPKG_INSTALLED_DIR'
        $ci | Should Match '&\s+"\$env:VCPKG_ROOT\\vcpkg\.exe"\s+@vcpkgArguments'
        $ci | Should Match '\$vcpkgRoot\s*=\s*Join-Path \$env:VCPKG_INSTALLED_DIR \$triplet'
        $ci | Should Match "hashFiles\('vcpkg\.json', 'vcpkg-tool\.json', 'vcpkg-install\.ps1', 'Tools/Modules/Build/VcpkgInstallSafety\.psm1', 'Build/vcpkg-ports/\*\*'\)"
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
        $bootstrap | Should Match 'Resolve-RSVcpkgSafeChildPath -Root \$RepoRoot -Child ''\.build/vcpkg-tool'''
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

    It 'installs CI dependencies from a registry checkout of the manifest baseline beside the pinned executable' {
        # The tool commit's own version database predates the versions vcpkg.json requires; with one checkout vcpkg fails
        # with "no version database entry" (Build_Toolchain.md).
        $ci = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw
        $ci | Should Match 'Initialize-RSVcpkgToolContext -RepoRoot \$env:GITHUB_WORKSPACE'
        $bootstrap = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -Raw
        $bootstrap | Should Match '\.build/vcpkg-registry'
        $bootstrap | Should Match '''worktree'', ''add'', ''--detach'', \$registryRoot, \$baseline'
        $bootstrap | Should Match '''rev-parse'', ''HEAD''\)\) -cne \$baseline'
        $ci | Should Match '"--vcpkg-root=\$env:VCPKG_REGISTRY_ROOT"'
        $ci | Should Match '"--downloads-root=\$downloadsRoot"'
        $ci.IndexOf('Initialize-RSVcpkgToolContext') | Should BeLessThan $ci.IndexOf('"--vcpkg-root=$env:VCPKG_REGISTRY_ROOT"')
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
        $nightly | Should Match 'options: \[Full, CommandsHandoff, FileOpsVolumes, ViewerSpaceCapture, R0ePaired, DesktopProbe\]'
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
        $qualification | Should Match 'platform: x64'
        $qualification | Should Match 'platform: ARM64'
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
        $installation = @(& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath)[0]
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
            Assert-RSThrowsMatching { & $find } 'not a Visual Studio 2026 executable'
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
            Assert-RSThrowsMatching { & $find } 'not a Visual Studio 2026 executable'
            $env:MSBUILD_EXE_PATH = Join-Path $TestDrive 'missing/MSBuild.exe'
            Assert-RSThrowsMatching { & $find } 'does not exist'
        } finally {
            $env:GITHUB_ACTIONS = $savedActions
            $env:MSBUILD_EXE_PATH = $savedMsbuild
        }
    }
}
