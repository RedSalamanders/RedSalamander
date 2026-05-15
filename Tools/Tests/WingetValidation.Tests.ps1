Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Installer\winget\WingetValidation.ps1'

Describe 'Winget validation helper' {
    BeforeAll {
        . $helperScript
    }

    function New-RSTemporaryWingetManifestRoot {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "RSWingetValidationTest_$([Guid]::NewGuid())"
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        return $tempRoot
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

Describe 'Winget manifest template' {
    BeforeAll {
        $installerTemplatePath = Join-Path $repoRoot 'Installer\winget\templates\RedSalamanders.RedSalamander.installer.yaml'
        $installerTemplate = Get-Content -Path $installerTemplatePath -Raw
        $zipScriptPath = Join-Path $repoRoot 'Installer\zip\build-zip.ps1'
        $zipScript = Get-Content -Path $zipScriptPath -Raw
    }

    It 'declares the dependency-free launcher as the portable alias target' {
        [regex]::Matches($installerTemplate, '(?m)^NestedInstallerFiles:\r?$').Count | Should Be 1
        $installerTemplate | Should Match '(?ms)^NestedInstallerFiles:\s*\r?\n\s+- RelativeFilePath: RedLauncher\.exe\r?\n\s+PortableCommandAlias: RedSalamander'
    }

    It 'marks RedLauncher.exe as the launch file in installation metadata' {
        $installerTemplate | Should Match '(?ms)^InstallationMetadata:\s*\r?\n\s+Files:\s*\r?\n\s+- RelativeFilePath: RedLauncher\.exe\r?\n\s+FileType: launch\r?\n\s+DisplayName: RedSalamander'
    }

    It 'copies RedLauncher.exe into the portable ZIP root' {
        $zipScript | Should Match 'Copy-Item \(Join-Path \$BuildOutputDir "RedLauncher\.exe"\) \$TempDir'
    }
}

Describe 'RedLauncher project' {
    BeforeAll {
        $launcherProjectPath = Join-Path $repoRoot 'RedLauncher\RedLauncher.vcxproj'
        $launcherSourcePath = Join-Path $repoRoot 'RedLauncher\Main.cpp'
    }

    It 'is a first-party static-CRT executable project' {
        Test-Path $launcherProjectPath | Should Be $true

        $launcherProject = Get-Content -Path $launcherProjectPath -Raw
        $launcherProject | Should Match '<ClCompile Include="Main\.cpp" />'
        $launcherProject | Should Match '<SubSystem>Console</SubSystem>'
        $launcherProject | Should Match '<RuntimeLibrary>MultiThreaded</RuntimeLibrary>'
        $launcherProject | Should Match '<RuntimeLibrary>MultiThreadedDebug</RuntimeLibrary>'
        $launcherProject | Should Not Match '<ProjectReference'
        $launcherProject | Should Not Match 'Common\.vcxproj'
    }

    It 'resolves the WinGet alias symlink before launching RedSalamander.exe' {
        Test-Path $launcherSourcePath | Should Be $true

        $launcherSource = Get-Content -Path $launcherSourcePath -Raw
        $launcherSource | Should Match 'GetFinalPathNameByHandleW'
        $launcherSource | Should Match 'RedSalamander\.exe'
        $launcherSource | Should Match 'CreateProcessW'
    }

    It 'waits only for foreground self-test invocations' {
        Test-Path $launcherSourcePath | Should Be $true

        $launcherSource = Get-Content -Path $launcherSourcePath -Raw
        $launcherSource | Should Match 'ShouldWaitForTargetExit'
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
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "RSVcRuntimeTest_$([Guid]::NewGuid())"
        try {
            $older = Join-Path $tempRoot 'Microsoft Visual Studio\18\BuildTools\VC\Redist\MSVC\14.50.10000\x64\Microsoft.VC145.CRT'
            $newer = Join-Path $tempRoot 'Microsoft Visual Studio\18\BuildTools\VC\Redist\MSVC\14.51.20000\x64\Microsoft.VC145.CRT'
            $oneCore = Join-Path $tempRoot 'Microsoft Visual Studio\18\BuildTools\VC\Redist\MSVC\14.99.99999\onecore\x64\Microsoft.VC145.CRT'
            New-Item -ItemType Directory -Path $older, $newer, $oneCore -Force | Out-Null

            Get-RSVcRuntimeRedistDirectory -Platform x64 -SearchRoots @($tempRoot) | Should Be $newer
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'copies app-local MSVC runtime DLLs and requires the core launch dependencies' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "RSVcRuntimeTest_$([Guid]::NewGuid())"
        $dest = Join-Path $tempRoot 'package'
        try {
            $redist = Join-Path $tempRoot 'Microsoft Visual Studio\18\BuildTools\VC\Redist\MSVC\14.51.20000\arm64\Microsoft.VC145.CRT'
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
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "RSVcRuntimeTest_$([Guid]::NewGuid())"
        $dest = Join-Path $tempRoot 'package'
        try {
            $redist = Join-Path $tempRoot 'Microsoft Visual Studio\18\BuildTools\VC\Redist\MSVC\14.51.20000\x64\Microsoft.VC145.CRT'
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
