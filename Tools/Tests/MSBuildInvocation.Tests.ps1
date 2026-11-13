Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperModule = Join-Path $repoRoot 'Tools\Modules\Build\MSBuildInvocation.psm1'
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/ProcessStreaming.psm1') -ErrorAction Stop

Describe 'Workflow-selected MSBuild identity' {
    BeforeAll {
        Import-Module $helperModule -Force -ErrorAction Stop
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -ErrorAction Stop
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'build.ps1'),[ref]$tokens,[ref]$errors)
        $function=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Find-MSBuild'},$true)
        Invoke-Expression $function.Extent.Text
    }
    BeforeEach {
        $script:savedActions=$env:GITHUB_ACTIONS;$script:savedSelected=$env:MSBUILD_EXE_PATH
        $env:GITHUB_ACTIONS='true';$env:MSBUILD_EXE_PATH=$null
        $root=Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $global:rsMsbuildSelectionFixture=@{A=(Join-Path $root 'selected older instance');B=(Join-Path $root 'latest competing instance')}
        foreach($installation in @($global:rsMsbuildSelectionFixture.A,$global:rsMsbuildSelectionFixture.B)) {
            foreach($relative in @(Get-RSHostOrderedMSBuildRelativePaths)) {
                $path=Join-Path $installation $relative
                [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
                [IO.File]::WriteAllText($path,'synthetic discovery fixture; never executed')
            }
        }
        Mock Get-RSMSBuildFileVersion { [pscustomobject]@{FileMajorPart=18;OriginalFilename='MSBuild.exe'} } -ModuleName MSBuildInvocation
        Mock Get-Command { [pscustomobject]@{Source=(Join-Path $global:rsMsbuildSelectionFixture.A 'MSBuild/Current/Bin/MSBuild.exe')} } `
            -ModuleName MSBuildInvocation -ParameterFilter { $Name -eq 'msbuild.exe' }
    }
    AfterEach {
        $env:GITHUB_ACTIONS=$script:savedActions;$env:MSBUILD_EXE_PATH=$script:savedSelected
        Remove-Variable rsMsbuildSelectionFixture -Scope Global -ErrorAction SilentlyContinue
    }
    It 'keeps PATH-only build and dependency selection in the selected installation despite a competitor' {
        $expected=Join-Path $global:rsMsbuildSelectionFixture.A @(Get-RSHostOrderedMSBuildRelativePaths)[0]
        $selected=Get-RSWorkflowSelectedMSBuild
        $selected.Path | Should Be $expected
        $selected.Method | Should Be 'PATH'
        (Find-MSBuild).Path | Should Be $expected
        Get-RSVcpkgVisualStudioInstallation | Should Be $global:rsMsbuildSelectionFixture.A
    }
    It 'gives explicit workflow selection priority over PATH without silently changing its executable' {
        $env:MSBUILD_EXE_PATH=Join-Path $global:rsMsbuildSelectionFixture.B 'MSBuild/Current/Bin/MSBuild.exe'
        (Get-RSWorkflowSelectedMSBuild).Path | Should Be $env:MSBUILD_EXE_PATH
        (Find-MSBuild).Path | Should Be $env:MSBUILD_EXE_PATH
        Get-RSVcpkgVisualStudioInstallation | Should Be $global:rsMsbuildSelectionFixture.B
        Assert-MockCalled Get-Command -ModuleName MSBuildInvocation -Times 0 -Exactly -Scope It -ParameterFilter { $Name -eq 'msbuild.exe' }
    }
    It 'fails closed for missing, foreign and wrong-version explicit selections' {
        foreach($path in @((Join-Path $global:rsMsbuildSelectionFixture.B 'MSBuild/Current/Bin/missing/MSBuild.exe'),(Join-Path $TestDrive 'foreign/MSBuild.exe'))) {
            $env:MSBUILD_EXE_PATH=$path
            { Get-RSWorkflowSelectedMSBuild } | Should Throw 'not an existing Visual Studio 2026 executable'
        }
        $env:MSBUILD_EXE_PATH=Join-Path $global:rsMsbuildSelectionFixture.B 'MSBuild/Current/Bin/MSBuild.exe'
        Mock Get-RSMSBuildFileVersion { [pscustomobject]@{FileMajorPart=17;OriginalFilename='MSBuild.exe'} } -ModuleName MSBuildInvocation
        { Get-RSWorkflowSelectedMSBuild } | Should Throw 'not an existing Visual Studio 2026 executable'
        Assert-MockCalled Get-Command -ModuleName MSBuildInvocation -Times 0 -Exactly -Scope It -ParameterFilter { $Name -eq 'msbuild.exe' }
    }
    It 'leaves invalid PATH candidates to ordinary discovery and ignores workflow hints outside Actions' {
        Mock Get-RSMSBuildFileVersion { [pscustomobject]@{FileMajorPart=17;OriginalFilename='MSBuild.exe'} } -ModuleName MSBuildInvocation
        ($null -eq (Get-RSWorkflowSelectedMSBuild)) | Should Be $true
        $env:GITHUB_ACTIONS=$null
        $env:MSBUILD_EXE_PATH=Join-Path $TestDrive 'unavailable/MSBuild.exe'
        ($null -eq (Get-RSWorkflowSelectedMSBuild)) | Should Be $true
    }
}

Describe 'MSBuild invocation helper' {
    BeforeAll {
        Import-Module $helperModule -Force -ErrorAction Stop
    }

    It 'shares host-ordered discovery for installation and developer-prompt paths' {
        $installation = Join-Path $TestDrive 'Visual Studio'
        $relatives = @(Get-RSHostOrderedMSBuildRelativePaths)
        foreach ($relative in $relatives) {
            $candidate = Join-Path $installation $relative
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($candidate))
            [IO.File]::WriteAllText($candidate, 'synthetic path fixture; never executed')
        }
        $expected = Join-Path $installation $relatives[0]
        (Find-RSHostOrderedMSBuildInInstallation -InstallationPath $installation) | Should Be $expected
        (Resolve-RSHostOrderedMSBuildPath -Path (Join-Path $installation 'MSBuild/Current/Bin/MSBuild.exe')) | Should Be $expected
        Remove-Item -LiteralPath $expected
        (Find-RSHostOrderedMSBuildInInstallation -InstallationPath $installation) | Should Be (Join-Path $installation $relatives[1])
    }

    It 'does not invent an MSBuild executable or rewrite unrelated explicit paths' {
        $installation = Join-Path $TestDrive 'Missing Visual Studio'
        ($null -eq (Find-RSHostOrderedMSBuildInInstallation -InstallationPath $installation)) | Should Be $true
        $unrelated = Join-Path $TestDrive 'explicit/MSBuild.exe'
        (Resolve-RSHostOrderedMSBuildPath -Path $unrelated) | Should Be $unrelated
        $relatives = @(Get-RSHostOrderedMSBuildRelativePaths)
        $relatives[0] | Should Match 'Bin\\(amd64|arm64)\\MSBuild\.exe$'
        if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -cne 'Arm64') {
            @($relatives | Where-Object { $_ -match '\\arm64\\' }).Count | Should Be 0
        }
    }

    It 'keeps CI-style redirected shells on the streaming path' {
        $envMap = @{
            CI = 'true'
        }

        $interactive = Test-RSInteractiveTerminal `
            -IsOutputRedirected $true `
            -IsErrorRedirected $true `
            -HasRawUi $true `
            -CanReadWindowTitle $true `
            -Environment $envMap

        Assert-RSEqual -Actual $interactive -Expected $false -Message 'CI-style redirected shells should not be treated as interactive.'
    }

    It 'keeps Codex terminals on the streaming path even when the host is interactive' {
        $plan = Get-RSMSBuildInvocationPlan `
            -UseInteractiveTerminal $true `
            -LogPath (Join-Path $TestDrive 'msbuild.log') `
            -Environment @{
                CODEX_SHELL = '1'
            }

        Assert-RSEqual -Actual $plan.UseDirectConsole -Expected $false -Message 'Codex terminals should use replay streaming.'
        Assert-RSEqual -Actual (@($plan.AdditionalArguments).Count) -Expected 0 -Message 'Codex terminal plans should not add file logger arguments.'
    }

    It 'keeps Windows Terminal on the streaming path so MSBuild progress is replayed' {
        $plan = Get-RSMSBuildInvocationPlan `
            -UseInteractiveTerminal $true `
            -LogPath (Join-Path $TestDrive 'msbuild.log') `
            -Environment @{
                WT_SESSION = '1'
            }

        Assert-RSEqual -Actual $plan.UseDirectConsole -Expected $false -Message 'Windows Terminal should use replay streaming.'
        Assert-RSEqual -Actual (@($plan.AdditionalArguments).Count) -Expected 0 -Message 'Streaming terminal plans should not add file logger arguments.'
    }

    It 'uses direct console execution for a plain interactive console while still capturing a log file' {
        $plan = Get-RSMSBuildInvocationPlan `
            -UseInteractiveTerminal $true `
            -LogPath (Join-Path $TestDrive 'msbuild.log') `
            -Environment @{
                PLAIN_CONSOLE = '1'
            }

        Assert-RSEqual -Actual $plan.UseDirectConsole -Expected $true -Message 'Plain interactive terminals should use direct console execution.'
        Assert-RSSequenceEqual `
            -Actual @($plan.AdditionalArguments) `
            -Expected @('/fl', "/flp:Verbosity=minimal;LogFile=$([System.IO.Path]::GetFullPath((Join-Path $TestDrive 'msbuild.log')));Encoding=UTF-8") `
            -Message 'Direct console plans should add file logger arguments.'
    }

    It 'keeps replay streaming for non-interactive terminals' {
        $plan = Get-RSMSBuildInvocationPlan -UseInteractiveTerminal $false -LogPath (Join-Path $TestDrive 'msbuild.log')

        Assert-RSEqual -Actual $plan.UseDirectConsole -Expected $false -Message 'Non-interactive terminals should use replay streaming.'
        Assert-RSEqual -Actual (@($plan.AdditionalArguments).Count) -Expected 0 -Message 'Non-interactive plans should not add file logger arguments.'
    }

    It 'colors replayed MSBuild errors and warnings' {
        Assert-RSEqual `
            -Actual (Get-RSMSBuildLineForegroundColor -Line 'MSBUILD : error MSB1009: Project file does not exist.' -IsError $false) `
            -Expected 'Red' `
            -Message 'MSBuild errors should be red.'

        Assert-RSEqual `
            -Actual (Get-RSMSBuildLineForegroundColor -Line 'file.cpp(10,5): warning C4100: unreferenced formal parameter' -IsError $false) `
            -Expected 'Yellow' `
            -Message 'MSBuild warnings should be yellow.'

        Assert-RSEqual `
            -Actual (Get-RSMSBuildLineForegroundColor -Line 'tool wrote to stderr' -IsError $true) `
            -Expected 'Red' `
            -Message 'stderr lines should be red.'
    }

    It 'colors replayed MSBuild project completion and keeps ordinary lines default' {
        Assert-RSEqual `
            -Actual (Get-RSMSBuildLineForegroundColor -Line '  Common.vcxproj -> Z:\src\RedSalamander\.build\x64\Debug\Common.dll' -IsError $false) `
            -Expected 'Green' `
            -Message 'MSBuild project completion lines should be green.'

        Assert-RSEqual `
            -Actual (Get-RSMSBuildLineForegroundColor -Line '  SettingsStore.cpp' -IsError $false) `
            -Expected $null `
            -Message 'Ordinary compile progress lines should keep the terminal default color.'
    }

    It 'summarizes MSBuild diagnostics from captured logs' {
        $logPath = Join-Path $TestDrive 'msbuild.log'
        @'
  FolderView.ErrorOverlay.cpp
Z:\src\RedSalamander\RedSalamander\Preferences.FileActions.cpp(779,20): warning C5245: unreferenced function [Z:\src\RedSalamander\RedSalamander\RedSalamander.vcxproj]
MSBUILD : error MSB1009: Project file does not exist.
link : fatal error LNK1120: 1 unresolved externals
  0 Warning(s)
  0 Error(s)
'@ | Set-Content -Path $logPath -Encoding ASCII

        $summary = Get-RSMSBuildDiagnosticSummary -LogPath $logPath

        Assert-RSEqual -Actual $summary.WarningCount -Expected 1 -Message 'The log should report one warning diagnostic.'
        Assert-RSEqual -Actual $summary.ErrorCount -Expected 2 -Message 'The log should report two error diagnostics.'
    }

    It 'counts and colors the same diagnostic header: <Name>' -TestCases @(
        @{ Name = 'unnumbered vcpkg warning'; Line = 'C:\VS\vcpkg\vcpkg.targets(242,5): warning : [vcpkg] Failed to gather app local DLL dependencies.'; ExpectedWarnings = 1; ExpectedErrors = 0; Color = 'Yellow' },
        @{ Name = 'unnumbered build error'; Line = 'project.targets(10,5): error : Required runtime input is missing.'; ExpectedWarnings = 0; ExpectedErrors = 1; Color = 'Red' },
        @{ Name = 'unnumbered fatal linker error'; Line = 'link : fatal error : Link failed.'; ExpectedWarnings = 0; ExpectedErrors = 1; Color = 'Red' },
        @{ Name = 'blank warning origin and code'; Line = 'warning : Action required.'; ExpectedWarnings = 1; ExpectedErrors = 0; Color = 'Yellow' },
        @{ Name = 'blank error origin and code'; Line = 'error: Action required.'; ExpectedWarnings = 0; ExpectedErrors = 1; Color = 'Red' },
        @{ Name = 'blank origin with separator'; Line = ' : warning : Action required.'; ExpectedWarnings = 1; ExpectedErrors = 0; Color = 'Yellow' },
        @{ Name = 'code with no origin'; Line = 'error CS0006: Metadata file missing.'; ExpectedWarnings = 0; ExpectedErrors = 1; Color = 'Red' },
        @{ Name = 'command-line subcategory'; Line = 'cl : Command line warning D4024 : unrecognized source file type'; ExpectedWarnings = 1; ExpectedErrors = 0; Color = 'Yellow' },
        @{ Name = 'case-insensitive category'; Line = 'tool: WARNING w100: Message'; ExpectedWarnings = 1; ExpectedErrors = 0; Color = 'Yellow' },
        @{ Name = 'non-numeric code'; Line = 'tool: warning STAGING_FAILED: Message'; ExpectedWarnings = 1; ExpectedErrors = 0; Color = 'Yellow' },
        @{ Name = 'numeric code'; Line = 'tool: error 1234: Message'; ExpectedWarnings = 0; ExpectedErrors = 1; Color = 'Red' },
        @{ Name = 'parallel node prefix'; Line = '  12>C:\src\file.cpp(1): warning C4100: Message'; ExpectedWarnings = 1; ExpectedErrors = 0; Color = 'Yellow' },
        @{ Name = 'UNC source'; Line = '\\server\src\file.cpp(1): error C1000: Message'; ExpectedWarnings = 0; ExpectedErrors = 1; Color = 'Red' },
        @{ Name = 'warning mentioning an error'; Line = 'file.cpp(1): warning C4996: old API says: error C1000: no longer valid'; ExpectedWarnings = 1; ExpectedErrors = 0; Color = 'Yellow' },
        @{ Name = 'error mentioning a warning'; Line = 'file.cpp(1): error C1000: emitted: warning C4996: old API'; ExpectedWarnings = 0; ExpectedErrors = 1; Color = 'Red' },
        @{ Name = 'summary warning count'; Line = '  1 Warning(s)'; ExpectedWarnings = 0; ExpectedErrors = 0; Color = $null },
        @{ Name = 'summary error count'; Line = '  0 Error(s)'; ExpectedWarnings = 0; ExpectedErrors = 0; Color = $null },
        @{ Name = 'source filename'; Line = '  FolderView.ErrorOverlay.cpp'; ExpectedWarnings = 0; ExpectedErrors = 0; Color = $null },
        @{ Name = 'ordinary progress'; Line = '  Generating code'; ExpectedWarnings = 0; ExpectedErrors = 0; Color = $null },
        @{ Name = 'incomplete header'; Line = '  warning C1000 without a diagnostic separator'; ExpectedWarnings = 0; ExpectedErrors = 0; Color = $null },
        @{ Name = 'empty line'; Line = ''; ExpectedWarnings = 0; ExpectedErrors = 0; Color = $null }
    ) {
        param($Name, $Line, $ExpectedWarnings, $ExpectedErrors, $Color)
        $logPath = Join-Path $TestDrive 'diagnostic-header.txt'
        [IO.File]::WriteAllText($logPath, $Line, [Text.UTF8Encoding]::new($false))
        $summary = Get-RSMSBuildDiagnosticSummary -LogPath $logPath
        Assert-RSEqual $summary.WarningCount $ExpectedWarnings "Warning count for $Name."
        Assert-RSEqual $summary.ErrorCount $ExpectedErrors "Error count for $Name."
        Assert-RSEqual (Get-RSMSBuildLineForegroundColor -Line $Line) $Color "Replay color for $Name."
    }

    It 'surfaces unnumbered diagnostics in the summary and streaming presenters' {
        $logPath = Join-Path $TestDrive 'unnumbered-diagnostics.txt'
        $warningLine = 'tool.targets(1): warning : Missing staged dependency.'
        [IO.File]::WriteAllText($logPath, $warningLine, [Text.UTF8Encoding]::new($false))
        Mock Write-Host {} -ModuleName MSBuildInvocation

        Write-RSMSBuildDiagnosticSummary -LogPath $logPath
        Write-RSMSBuildStreamingLine -Line $warningLine
        Assert-MockCalled Write-Host -ModuleName MSBuildInvocation -Times 1 -Exactly -ParameterFilter {
            $Object -eq 'Diagnostics: 1 warning(s), 0 error(s)' -and $ForegroundColor -eq 'Yellow'
        }
        Assert-MockCalled Write-Host -ModuleName MSBuildInvocation -Times 1 -Exactly -ParameterFilter {
            $Object -eq 'tool.targets(1): warning : Missing staged dependency.' -and $ForegroundColor -eq 'Yellow'
        }
    }

    It 'classifies the locale-neutral category independently of the current culture' {
        # Use a cold process: PowerShell's regex cache can retain a pattern first
        # compiled under another culture and make a same-process probe pass.
        $scriptPath = Join-Path $TestDrive 'ColdDiagnosticCulture.ps1'
        $logPath = Join-Path $TestDrive 'locale-neutral-warning.txt'
        $scriptText = @'
param([string]$ModulePath, [string]$LogPath)
$ErrorActionPreference = 'Stop'
[Globalization.CultureInfo]::CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('tr-TR')
Import-Module $ModulePath
$line = 'tool: WARNING : Localized message text.'
[IO.File]::WriteAllText($LogPath, $line, [Text.UTF8Encoding]::new($false))
$summary = Get-RSMSBuildDiagnosticSummary -LogPath $LogPath
if ($summary.WarningCount -ne 1) { throw 'Cold Turkish locale did not count uppercase WARNING.' }
if ((Get-RSMSBuildLineForegroundColor -Line $line) -ne 'Yellow') { throw 'Cold Turkish locale did not color uppercase WARNING.' }
'Cold Turkish category classification passed.'
'@
        [IO.File]::WriteAllText($scriptPath, $scriptText, [Text.UTF8Encoding]::new($false))
        $exitCode = Invoke-RSStreamingProcess -FilePath 'pwsh.exe' -Arguments @('-NoProfile', '-File', $scriptPath, $helperModule, $logPath) `
            -LogPath (Join-Path $TestDrive 'locale-probe.txt') -TimeoutMilliseconds 10000
        Assert-RSEqual $exitCode 0 'The cold-process locale classification must pass.'
    }
}
