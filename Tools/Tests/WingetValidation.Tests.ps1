Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRunPlanScript = Join-Path $repoRoot 'Tools\TestRunPlan.ps1'
$helperScript = Join-Path $repoRoot 'Installer\winget\WingetValidation.ps1'
. $testRunPlanScript

Describe 'Winget validation helper' {
    BeforeAll {
        . $helperScript
    }

    function New-RSTemporaryWingetManifestRoot {
        return (New-RSTestSandboxScratchDirectory `
                -RepoRoot $repoRoot `
                -Harness 'tools-pester' `
                -Case "winget-validation-$([Guid]::NewGuid().ToString('N'))")
    }

    function New-RSFakeWingetCommand {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Root,

            [Parameter(Mandatory = $true)]
            [string]$VersionText,

            [Parameter(Mandatory = $true)]
            [string[]]$ValidationOutput,

            [Parameter(Mandatory = $true)]
            [int]$ValidationExitCode
        )

        $scriptPath = Join-Path $Root 'winget.cmd'
        $outputCommands = @($ValidationOutput | ForEach-Object { "echo $_" }) -join "`r`n"
        Set-Content -Path $scriptPath -Encoding ASCII -Value @"
@echo off
if "%1"=="--version" (
  echo $VersionText
  exit /b 0
)
if "%1"=="validate" (
  $outputCommands
  exit /b $ValidationExitCode
)
exit /b 99
"@
        return $scriptPath
    }

    It 'allows only the known winget 1.11 schema-header warning for 1.12 manifests' {
        $output = @(
            'Manifest validation succeeded with warnings.'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.installer.1.12.0.schema.json Line: 1, Column: 25 File: RedSalamanders.RedSalamander.installer.yaml'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.defaultLocale.1.12.0.schema.json Line: 1, Column: 25 File: RedSalamanders.RedSalamander.locale.en-US.yaml'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.version.1.12.0.schema.json Line: 1, Column: 25 File: RedSalamanders.RedSalamander.yaml'
        )

        Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion 'v1.11.510' -OutputLines $output | Should Be $true
    }

    It 'rejects unrelated warnings' {
        $output = @(
            'Manifest validation succeeded with warnings.'
            'Manifest Warning: Field usage requires verified publishers. [Icons]'
        )

        Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion 'v1.11.510' -OutputLines $output | Should Be $false
    }

    It 'does not suppress warnings from current winget validators' {
        $output = @(
            'Manifest validation succeeded with warnings.'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.installer.1.12.0.schema.json'
        )

        Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion 'v1.28.240' -OutputLines $output | Should Be $false
    }

    It 'fails clearly when the winget executable is unavailable' {
        $tempRoot = New-RSTemporaryWingetManifestRoot
        try {
            $missingWinget = Join-Path $tempRoot 'missing-winget.exe'

            { Invoke-RSWingetManifestValidation -ManifestPath $tempRoot -WingetCommand $missingWinget } |
                Should Throw "winget executable not found: $missingWinget"
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'propagates real validation failures with the winget exit code' {
        $tempRoot = New-RSTemporaryWingetManifestRoot
        try {
            $fakeWinget = New-RSFakeWingetCommand `
                -Root $tempRoot `
                -VersionText 'v1.12.0' `
                -ValidationOutput @('Manifest validation failed.', 'Manifest Error: PackageVersion is required.') `
                -ValidationExitCode 23

            { Invoke-RSWingetManifestValidation -ManifestPath $tempRoot -WingetCommand $fakeWinget } |
                Should Throw 'winget validate failed with exit code 23.'
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'does not suppress unrelated validation failures from legacy winget validators' {
        $tempRoot = New-RSTemporaryWingetManifestRoot
        try {
            $fakeWinget = New-RSFakeWingetCommand `
                -Root $tempRoot `
                -VersionText 'v1.11.510' `
                -ValidationOutput @('Manifest validation failed.', 'Manifest Error: InstallerSha256 is invalid.') `
                -ValidationExitCode 7

            { Invoke-RSWingetManifestValidation -ManifestPath $tempRoot -WingetCommand $fakeWinget } |
                Should Throw 'winget validate failed with exit code 7.'
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Winget release workflow' {
    BeforeAll {
        $workflowPath = Join-Path $repoRoot '.github\workflows\winget-release.yml'
        $workflow = Get-Content -Path $workflowPath -Raw
    }

    It 'enables local manifest installs before testing the generated manifest' {
        $enableIndex = $workflow.IndexOf('winget settings --enable LocalManifestFiles')
        $installMatch = [regex]::Match($workflow, '(?ms)winget install\s+`\r?\n\s+--manifest "winget-manifest"')

        if ($enableIndex -lt 0) {
            throw 'winget-release.yml must enable LocalManifestFiles before winget install --manifest.'
        }

        if (-not $installMatch.Success) {
            throw 'winget-release.yml must install the generated winget-manifest.'
        }

        ($enableIndex -lt $installMatch.Index) | Should Be $true
    }
}

Describe 'Winget manifest template' {
    BeforeAll {
        $installerTemplatePath = Join-Path $repoRoot 'Installer\winget\templates\RedSalamanders.RedSalamander.installer.yaml'
        $installerTemplate = Get-Content -Path $installerTemplatePath -Raw
        $msixManifestPath = Join-Path $repoRoot 'Installer\msix\Package.appxmanifest'
        $msixManifest = Get-Content -Path $msixManifestPath -Raw
        $msixProjectPath = Join-Path $repoRoot 'Installer\msix\RedSalamanderInstaller.wapproj'
        $msixProject = Get-Content -Path $msixProjectPath -Raw
        $directoryBuildPropsPath = Join-Path $repoRoot 'Directory.Build.props'
        $directoryBuildProps = Get-Content -Path $directoryBuildPropsPath -Raw
        $zipScriptPath = Join-Path $repoRoot 'Installer\zip\build-zip.ps1'
        $zipScript = Get-Content -Path $zipScriptPath -Raw
        $minimumOsHeaderPath = Join-Path $repoRoot 'Common\MinimumOsVersion.h'
        $minimumOsHeader = Get-Content -Path $minimumOsHeaderPath -Raw
        $minimumOsSourcePath = Join-Path $repoRoot 'Common\MinimumOsVersion.cpp'
        $minimumOsSource = Get-Content -Path $minimumOsSourcePath -Raw
    }

    It 'declares the dependency-free launcher aliases as portable alias targets on the Windows 11 build 22000.2600 floor' {
        $installerTemplate | Should Match '(?m)^MinimumOSVersion:\s+10\.0\.22000\.2600\r?$'
        $installerTemplate | Should Not Match '10\.0\.26100\.0'
        $installerTemplate | Should Not Match '10\.0\.19041\.0'
        $msixManifest | Should Match 'MinVersion="10\.0\.22000\.2600"'
        $msixProject | Should Match '<TargetPlatformMinVersion>10\.0\.22000\.2600</TargetPlatformMinVersion>'
        $directoryBuildProps | Should Match '<WindowsTargetPlatformVersion>10\.0\.26100\.0</WindowsTargetPlatformVersion>'
        $directoryBuildProps | Should Match 'NTDDI_VERSION=NTDDI_WIN11_GE'
        $minimumOsHeader | Should Match 'kMinimumWindowsBuildNumber\s*=\s*22000'
        $minimumOsHeader | Should Match 'kMinimumWindowsBuildRevision\s*=\s*2600'
        $minimumOsSource | Should Match 'TryGetWindowsBuildRevision'
        $minimumOsSource | Should Match 'CurrentVersion'
        $zipScript | Should Match 'Windows 11 build 22000\.2600'
        [regex]::Matches($installerTemplate, '(?m)^NestedInstallerFiles:\r?$').Count | Should Be 1
        [regex]::Matches($installerTemplate, '(?m)^\s+- RelativeFilePath: RedLauncher\.exe\r?\n\s+PortableCommandAlias: RedSalamander\r?$').Count | Should Be 1
        [regex]::Matches($installerTemplate, '(?m)^\s+- RelativeFilePath: red\.exe\r?\n\s+PortableCommandAlias: red\r?$').Count | Should Be 1
        $installerTemplate | Should Match '(?ms)^Commands:\s*\r?\n\s+- RedSalamander\r?\n\s+- red\r?\n'
    }

    It 'marks the launcher executables as launch files in installation metadata' {
        $installerTemplate | Should Match '(?ms)^InstallationMetadata:\s*\r?\n\s+Files:\s*\r?\n\s+- RelativeFilePath: RedLauncher\.exe\r?\n\s+FileType: launch\r?\n\s+DisplayName: RedSalamander'
        $installerTemplate | Should Match '(?ms)^InstallationMetadata:.*\r?\n\s+- RelativeFilePath: red\.exe\r?\n\s+FileType: launch\r?\n\s+DisplayName: red'
    }

    It 'copies RedLauncher.exe and its red.exe alias into the portable ZIP root' {
        $zipScript | Should Match 'Copy-Item \(Join-Path \$BuildOutputDir "RedLauncher\.exe"\) \$TempDir'
        $zipScript | Should Match 'Copy-Item \(Join-Path \$BuildOutputDir "red\.exe"\) \$TempDir'
        $zipScript | Should Not Match 'RedLauncherConsole\.exe'
    }

    It 'preserves package subdirectories for plugins, language satellites, and themes' {
        $zipScript | Should Match '\$LangSource = Join-Path \$BuildOutputDir "Lang"'
        $zipScript | Should Match 'Copy-Item \$LangSource -Destination \$LangDest -Recurse -Force'

        $catchAllInclude = 'Content Include="..\..\.build\$(Platform)\$(Configuration)\**\*"'
        $catchAllIndex = $msixProject.IndexOf($catchAllInclude, [System.StringComparison]::Ordinal)
        $catchAllIndex | Should Not Be -1
        $catchAllEnd = $msixProject.IndexOf('>', $catchAllIndex)
        $catchAllLine = $msixProject.Substring($catchAllIndex, $catchAllEnd - $catchAllIndex)

        foreach ($folder in @('Plugins', 'Lang', 'Themes')) {
            $include = 'Content Include="..\..\.build\$(Platform)\$(Configuration)\{0}\**\*"' -f $folder
            $exclude = '..\..\.build\$(Platform)\$(Configuration)\{0}\**\*' -f $folder
            $packagePath = '<PackagePath>{0}\%(RecursiveDir)%(Filename)%(Extension)</PackagePath>' -f $folder
            $link = '<Link>{0}\%(RecursiveDir)%(Filename)%(Extension)</Link>' -f $folder
            $targetPath = '<TargetPath>{0}\%(RecursiveDir)%(Filename)%(Extension)</TargetPath>' -f $folder

            $msixProject.Contains($include) | Should Be $true
            $catchAllLine.Contains($exclude) | Should Be $true
            $msixProject.Contains($packagePath) | Should Be $true
            $msixProject.Contains($link) | Should Be $true
            $msixProject.Contains($targetPath) | Should Be $true
        }
    }
}

Describe 'RedLauncher project' {
    BeforeAll {
        $launcherProjectPath = Join-Path $repoRoot 'RedLauncher\RedLauncher.vcxproj'
        $consoleLauncherProjectPath = Join-Path $repoRoot 'RedLauncher\RedLauncherConsole.vcxproj'
        $launcherManifestPath = Join-Path $repoRoot 'RedLauncher\res\exe.manifest'
        $launcherSourcePath = Join-Path $repoRoot 'RedLauncher\Main.cpp'
        $redSalamanderSourcePath = Join-Path $repoRoot 'RedSalamander\RedSalamander.cpp'
        $monitorSourcePath = Join-Path $repoRoot 'RedSalamanderMonitor\RedSalamanderMonitor.cpp'
        $configureSourcePath = Join-Path $repoRoot 'RedConfigure\Main.cpp'
        $searchServiceSourcePath = Join-Path $repoRoot 'RedSalamanderSearchService\Main.cpp'
        $solutionPath = Join-Path $repoRoot 'RedSalamander.sln'
        $solution = Get-Content -Path $solutionPath -Raw
        $buildScriptPath = Join-Path $repoRoot 'build.ps1'
        $buildScript = Get-Content -Path $buildScriptPath -Raw
    }

    It 'is a first-party static-CRT detached console executable project for the WinGet alias' {
        Test-Path $launcherProjectPath | Should Be $true
        Test-Path $consoleLauncherProjectPath | Should Be $false
        Test-Path $launcherManifestPath | Should Be $true

        $launcherProject = Get-Content -Path $launcherProjectPath -Raw
        $launcherManifest = Get-Content -Path $launcherManifestPath -Raw
        $launcherProject | Should Match '<ClCompile Include="Main\.cpp" />'
        $launcherProject | Should Match '<SubSystem>Console</SubSystem>'
        $launcherProject | Should Match '<AdditionalManifestFiles>\$\(ProjectDir\)res\\exe\.manifest</AdditionalManifestFiles>'
        $launcherManifest | Should Match 'consoleAllocationPolicy'
        $launcherManifest | Should Match '>detached<'
        $launcherProject | Should Not Match 'REDLAUNCHER_GUI'
        $launcherProject | Should Not Match 'REDLAUNCHER_CONSOLE'
        $launcherProject | Should Match '<RuntimeLibrary>MultiThreaded</RuntimeLibrary>'
        $launcherProject | Should Match '<RuntimeLibrary>MultiThreadedDebug</RuntimeLibrary>'
        $launcherProject | Should Match '<Target Name="CopyRedLauncherAlias"'
        $launcherProject | Should Match 'SourceFiles="\$\(TargetPath\)"'
        $launcherProject | Should Match 'DestinationFiles="\$\(OutDir\)red\.exe"'
        $launcherProject | Should Not Match '<ProjectReference'
        $launcherProject | Should Not Match 'Common\.vcxproj'
        $solution | Should Not Match 'RedLauncherConsole'
        $buildScript | Should Not Match 'RedLauncherConsole'
    }

    It 'resolves the WinGet alias symlink before launching RedSalamander.exe' {
        Test-Path $launcherSourcePath | Should Be $true

        $launcherSource = Get-Content -Path $launcherSourcePath -Raw
        $launcherSource | Should Match 'GetFinalPathNameByHandleW'
        $launcherSource | Should Match 'RedSalamander\.exe'
        $launcherSource | Should Match 'CreateProcessW'
        $launcherSource | Should Match 'RtlGetVersion'
        $launcherSource | Should Match 'TryGetWindowsBuildRevision'
        $launcherSource | Should Match 'kMinimumWindowsBuildNumber\s*=\s*22000'
        $launcherSource | Should Match 'kMinimumWindowsBuildRevision\s*=\s*2600'
        $launcherSource | Should Match 'Windows 11 build 22000\.2600'
        (Get-Content -Path $redSalamanderSourcePath -Raw) | Should Match 'EnsureCurrentWindowsVersionSupported'
        (Get-Content -Path $monitorSourcePath -Raw) | Should Match 'EnsureCurrentWindowsVersionSupported'
        (Get-Content -Path $configureSourcePath -Raw) | Should Match 'EnsureCurrentWindowsVersionSupported'
        (Get-Content -Path $searchServiceSourcePath -Raw) | Should Match 'EnsureCurrentWindowsVersionSupported'
    }

    It 'waits only for foreground self-test invocations' {
        Test-Path $launcherSourcePath | Should Be $true

        $launcherSource = Get-Content -Path $launcherSourcePath -Raw
        $launcherSource | Should Match 'ShouldWaitForTargetExit'
        $launcherSource | Should Match 'wmain'
        $launcherSource | Should Not Match 'wWinMain'
        $launcherSource | Should Match '--selftest'
        $launcherSource | Should Match '--compare-selftest'
        $launcherSource | Should Match '--commands-selftest'
        $launcherSource | Should Match '--fileops-selftest'
        $launcherSource | Should Match '--selftest-list-cases'
        $launcherSource | Should Not Match 'argc\s*>\s*1'
        $launcherSource | Should Not Match '--etw'
        $launcherSource | Should Not Match '--perf'
    }

}

Describe 'VC runtime ZIP helper' {
    BeforeAll {
        . (Join-Path $repoRoot 'Installer\zip\VcRuntime.ps1')
    }

    It 'selects the latest non-OneCore MSVC redist directory for the requested architecture' {
        $caseId = ([Guid]::NewGuid().ToString('N')).Substring(0, 8)
        $tempRoot = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "vc-redist-$caseId"
        try {
            $older = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.50.10000\x64\Microsoft.VC145.CRT'
            $newer = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.51.20000\x64\Microsoft.VC145.CRT'
            $oneCore = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.99.99999\onecore\x64\Microsoft.VC145.CRT'
            New-Item -ItemType Directory -Path $older, $newer, $oneCore -Force | Out-Null

            Get-RSVcRuntimeRedistDirectory -Platform x64 -SearchRoots @($tempRoot) | Should Be $newer
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'copies app-local MSVC runtime DLLs and requires the core launch dependencies' {
        $caseId = ([Guid]::NewGuid().ToString('N')).Substring(0, 8)
        $tempRoot = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "vc-copy-$caseId"
        $dest = Join-Path $tempRoot 'package'
        try {
            $redist = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.51.20000\arm64\Microsoft.VC145.CRT'
            New-Item -ItemType Directory -Path $redist, $dest -Force | Out-Null
            foreach ($dllName in @('msvcp140.dll', 'msvcp140_atomic_wait.dll', 'vcruntime140.dll', 'vcruntime140_1.dll', 'vccorlib140.dll')) {
                Set-Content -Path (Join-Path $redist $dllName) -Value $dllName -Encoding ASCII
            }

            $copied = Copy-RSVcRuntimeDependencies -Platform ARM64 -DestinationDir $dest -SearchRoots @($tempRoot)

            ($copied -contains 'msvcp140.dll') | Should Be $true
            ($copied -contains 'msvcp140_atomic_wait.dll') | Should Be $true
            ($copied -contains 'vcruntime140.dll') | Should Be $true
            ($copied -contains 'vcruntime140_1.dll') | Should Be $true
            Test-Path (Join-Path $dest 'vccorlib140.dll') | Should Be $true
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'fails clearly when a required runtime DLL is missing' {
        $caseId = ([Guid]::NewGuid().ToString('N')).Substring(0, 8)
        $tempRoot = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "vc-missing-$caseId"
        $dest = Join-Path $tempRoot 'package'
        try {
            $redist = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.51.20000\x64\Microsoft.VC145.CRT'
            New-Item -ItemType Directory -Path $redist, $dest -Force | Out-Null
            foreach ($dllName in @('msvcp140.dll', 'vcruntime140.dll', 'vcruntime140_1.dll')) {
                Set-Content -Path (Join-Path $redist $dllName) -Value $dllName -Encoding ASCII
            }

            { Copy-RSVcRuntimeDependencies -Platform x64 -DestinationDir $dest -SearchRoots @($tempRoot) } |
                Should Throw "Required MSVC runtime DLL 'msvcp140_atomic_wait.dll'"
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
