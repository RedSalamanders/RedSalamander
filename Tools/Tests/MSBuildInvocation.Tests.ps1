Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Tools\MSBuildInvocation.ps1'

function Assert-RSEqual {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Actual,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected' but got '$Actual'."
    }
}

function Assert-RSSequenceEqual {
    param(
        [object[]]$Actual,
        [object[]]$Expected,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($Actual.Count -ne $Expected.Count) {
        throw "$Message Expected $($Expected.Count) item(s) but got $($Actual.Count)."
    }

    for ($i = 0; $i -lt $Expected.Count; $i++) {
        if ($Actual[$i] -ne $Expected[$i]) {
            throw "$Message Item $i expected '$($Expected[$i])' but got '$($Actual[$i])'."
        }
    }
}

Describe 'MSBuild invocation helper' {
    BeforeAll {
        . $helperScript
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
}
