Set-StrictMode -Version Latest

Import-Module (Join-Path $PSScriptRoot 'TestRunPlan.Common.psm1') -Force -Scope Local
Import-Module (Join-Path $PSScriptRoot 'TestInvocation.psm1') -Force -Scope Local

function Get-RSCommandsSelfTestFamilies {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject][ordered]@{ name = 'settings'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp' }
        [pscustomobject][ordered]@{ name = 'theme-overlay'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.ThemeOverlay.cpp' }
        [pscustomobject][ordered]@{ name = 'batch-rename'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.BatchRename.cpp' }
        [pscustomobject][ordered]@{ name = 'plugin-config'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp' }
        [pscustomobject][ordered]@{ name = 'connections'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp' }
        [pscustomobject][ordered]@{ name = 'preferences'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp' }
        [pscustomobject][ordered]@{ name = 'search'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp' }
        [pscustomobject][ordered]@{ name = 'shell-commands'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp' }
        [pscustomobject][ordered]@{ name = 'shortcuts'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.Shortcuts.cpp' }
        [pscustomobject][ordered]@{ name = 'compare-options'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp' }
        [pscustomobject][ordered]@{ name = 'file-operations'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp' }
        [pscustomobject][ordered]@{ name = 'navigation'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp' }
        [pscustomobject][ordered]@{ name = 'dialogs'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp' }
        [pscustomobject][ordered]@{ name = 'view-commands'; source = 'RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp' }
    )
}

function Get-RSDxUiTestSuites {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject][ordered]@{ Name = 'Grid'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'Theme'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'Control'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'Menu'; AllowsActivation = $true; RequiresInteractiveDesktop = $true }
        [pscustomobject][ordered]@{ Name = 'NewControls'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'TextField'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'NativeTextInput'; AllowsActivation = $true; RequiresInteractiveDesktop = $true }
        [pscustomobject][ordered]@{ Name = 'ComboBox'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'WindowHost'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'Tree'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'MultilineText'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'ReadOnly'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'Tooltip'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'Rendering'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'Animation'; AllowsActivation = $false; RequiresInteractiveDesktop = $false }
        [pscustomobject][ordered]@{ Name = 'Accessibility'; AllowsActivation = $false; RequiresInteractiveDesktop = $true }
    )
}

function Get-RSViewerTestGroups {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject][ordered]@{ Id = 'viewer-pe.noninteractive'; Name = 'ViewerPETests.Noninteractive'; Executable = 'ViewerPETests.exe'; Group = 'noninteractive'; Interactive = $false }
        [pscustomobject][ordered]@{ Id = 'viewer-pe.interactive'; Name = 'ViewerPETests.Interactive'; Executable = 'ViewerPETests.exe'; Group = 'interactive'; Interactive = $true }
        [pscustomobject][ordered]@{ Id = 'viewer-sqlite.noninteractive'; Name = 'ViewerSqliteTests.Noninteractive'; Executable = 'ViewerSqliteTests.exe'; Group = 'noninteractive'; Interactive = $false }
        [pscustomobject][ordered]@{ Id = 'viewer-sqlite.interactive'; Name = 'ViewerSqliteTests.Interactive'; Executable = 'ViewerSqliteTests.exe'; Group = 'interactive'; Interactive = $true }
    )
}

function Get-RSTestRunPlan {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'CI', 'Full')]
        [string]$Suite,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$RedSalamanderExePath,

        [double]$TimeoutMultiplier = 1.0,

        [switch]$FailFast,

        [string]$CaseFilter = '',

        [string]$CommandsFamily = '',

        [uint32]$RepeatCount = 1,

        [string]$ShuffleSeed = '',

        [string]$PerfBudgetPath = '',

        [switch]$RequirePerfBudgets,

        [string]$SelfTestFlakyProofCase = '',

        [string]$SelfTestOrderProofCase = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($CommandsFamily)) {
        $knownFamilies = @(Get-RSCommandsSelfTestFamilies | ForEach-Object { $_.name })
        if ($Suite -ne 'Commands' -or $CommandsFamily -notin $knownFamilies) {
            throw "-CommandsFamily requires -Suite Commands and one of: $($knownFamilies -join ', ')."
        }
    }

    $buildOutputDir = Join-Path $RepoRoot (".build\{0}\{1}" -f $Platform, $Configuration)
    $plan = @()

    $selfTests = @()
    switch ($Suite) {
        'All' {
            $selfTests += @{ Id = 'selftest.compare-directories'; Name = 'CompareDirectories'; Flag = '--compare-selftest'; JsonName = 'compare' }
            $selfTests += @{ Id = 'selftest.commands'; Name = 'Commands'; Flag = '--commands-selftest'; JsonName = 'commands' }
            $selfTests += @{ Id = 'selftest.file-operations'; Name = 'FileOperations'; Flag = '--fileops-selftest'; JsonName = 'fileops' }
        }
        { $_ -in @('CI', 'Full') } {
            $selfTests += @{ Id = 'selftest.compare-directories'; Name = 'CompareDirectories'; Flag = '--compare-selftest'; JsonName = 'compare' }
            $selfTests += @{ Id = 'selftest.commands'; Name = 'Commands'; Flag = '--commands-selftest'; JsonName = 'commands' }
            $selfTests += @{ Id = 'selftest.file-operations'; Name = 'FileOperations'; Flag = '--fileops-selftest'; JsonName = 'fileops' }
        }
        'Compare' {
            $selfTests += @{ Id = 'selftest.compare-directories'; Name = 'CompareDirectories'; Flag = '--compare-selftest'; JsonName = 'compare' }
        }
        'Commands' {
            $selfTests += @{ Id = 'selftest.commands'; Name = 'Commands'; Flag = '--commands-selftest'; JsonName = 'commands' }
        }
        'FileOps' {
            $selfTests += @{ Id = 'selftest.file-operations'; Name = 'FileOperations'; Flag = '--fileops-selftest'; JsonName = 'fileops' }
        }
    }

    foreach ($selfTest in $selfTests) {
        $requiresInteractiveDesktop = $selfTest.Name -eq 'Commands' -or ($selfTest.Name -eq 'FileOperations' -and $CaseFilter -eq 'FileOps_VisualGalleryInteraction')
        $plan += New-RSTestRunPlanEntry `
            -Id $selfTest.Id `
            -Name $selfTest.Name `
            -Kind 'SelfTest' `
            -Path $RedSalamanderExePath `
            -Arguments (Get-RSSelfTestArguments -Flag $selfTest.Flag -TimeoutMultiplier $TimeoutMultiplier -FailFast:$FailFast -CaseFilter $CaseFilter -CommandsFamily $CommandsFamily -RepeatCount $RepeatCount -ShuffleSeed $ShuffleSeed -PerfBudgetPath $PerfBudgetPath -RequirePerfBudgets:$RequirePerfBudgets -NoActivate:(-not $requiresInteractiveDesktop) -SelfTestFlakyProofCase $SelfTestFlakyProofCase -SelfTestOrderProofCase $SelfTestOrderProofCase) `
            -WorkingDirectory $RepoRoot `
            -JsonName $selfTest.JsonName `
            -RequiresInteractiveDesktop:$requiresInteractiveDesktop
    }

    if ($Suite -eq 'CI') {
        $plan += New-RSTestRunPlanEntry `
            -Id 'standalone.product-ui' `
            -Name 'ProductUiTests' `
            -Kind 'Executable' `
            -Path (Join-Path $buildOutputDir 'ProductUiTests.exe') `
            -WorkingDirectory $buildOutputDir

        $plan += New-RSTestRunPlanEntry `
            -Id 'standalone.filesystem-curl' `
            -Name 'FileSystemCurlTests' `
            -Kind 'Executable' `
            -Path (Join-Path $buildOutputDir 'FileSystemCurlTests.exe') `
            -WorkingDirectory $buildOutputDir

        foreach ($viewerGroup in @(Get-RSViewerTestGroups)) {
            $plan += New-RSTestRunPlanEntry `
                -Id $viewerGroup.Id `
                -Name $viewerGroup.Name `
                -Kind 'Executable' `
                -Path (Join-Path $buildOutputDir $viewerGroup.Executable) `
                -Arguments @("--group=$($viewerGroup.Group)") `
                -WorkingDirectory $buildOutputDir `
                -RequiresInteractiveDesktop:$viewerGroup.Interactive
        }

        foreach ($viewerCase in @(
                'TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly',
                'TestViewerTextGotoPromptUsesDxUiHostAndClosesCleanly'
            )) {
            $viewerCaseId = if ($viewerCase -like '*FindPrompt*') { 'viewer-pe.find-prompt' } else { 'viewer-pe.goto-prompt' }
            $plan += New-RSTestRunPlanEntry `
                -Id $viewerCaseId `
                -Name "ViewerPETests.$viewerCase" `
                -Kind 'Executable' `
                -Path (Join-Path $buildOutputDir 'ViewerPETests.exe') `
                -Arguments @($viewerCase) `
                -WorkingDirectory $buildOutputDir `
                -RequiresInteractiveDesktop
        }

        $ciExecutableIds = [ordered]@{
            MonitorTest = 'standalone.monitor'
            LocalizationTests = 'standalone.localization'
            RedConfigureTests = 'standalone.red-configure'
            PluginContractTests = 'standalone.plugin-contract'
            SettingsSchemaTests = 'standalone.settings-schema'
            CrashHandlingTests = 'standalone.crash-handling'
        }
        foreach ($exeName in $ciExecutableIds.Keys) {
            $plan += New-RSTestRunPlanEntry `
                -Id $ciExecutableIds[$exeName] `
                -Name $exeName `
                -Kind 'Executable' `
                -Path (Join-Path $buildOutputDir "$exeName.exe") `
                -WorkingDirectory $buildOutputDir
        }

        $plan += New-RSTestRunPlanEntry `
            -Id 'perf.performance-tests2' `
            -Name 'PerformanceTests2' `
            -Kind 'CppUnitTest' `
            -Path (Join-Path $buildOutputDir 'PerformanceTests2.dll') `
            -WorkingDirectory $buildOutputDir

        $plan += New-RSTestRunPlanEntry `
            -Id 'tooling.pester' `
            -Name 'ToolsPesterTests' `
            -Kind 'Pester' `
            -Path (Join-Path $RepoRoot 'Tools\TerminalEngine\CurrentToolingProfile.Tests.ps1') `
            -Arguments @('-ExcludeTag', 'RequiresBuildToolchain') `
            -WorkingDirectory $RepoRoot

        $plan += New-RSTestRunPlanEntry `
            -Id 'tooling.vcpkg-merge-synthetic' `
            -Name 'VcpkgMergeSynthetic' `
            -Kind 'PowerShellScript' `
            -Path (Join-Path $RepoRoot 'Tests\vcpkg-merge-synthetic-test.ps1') `
            -WorkingDirectory $RepoRoot
    }

    # Suite CI is the GitHub Actions PR gate. Suite Full remains the broader local/closeout gate
    # and additionally includes diagnostics such as RedSalamanderMonitorEtwLatency.
    if ($Suite -eq 'Full') {
        # The artifact-mutating deployment proof is the first canonical Full entry.
        # Run-AllTests re-attests a byte-verified full receipt, opens a new artifact
        # epoch, and re-fingerprints every reader before it is launched.
        $plan = @((New-RSTestRunPlanEntry `
                -Id 'tooling.pester-build-toolchain' `
                -Name 'ToolsPesterBuildToolchain' `
                -Kind 'Pester' `
                -Path (Join-Path $RepoRoot 'Tools\TerminalEngine\CurrentToolingProfile.Tests.ps1') `
                -Arguments @('-Tag', 'RequiresBuildToolchain') `
                -WorkingDirectory $RepoRoot)) + @($plan)

        $plan += New-RSTestRunPlanEntry `
            -Id 'standalone.product-ui' `
            -Name 'ProductUiTests' `
            -Kind 'Executable' `
            -Path (Join-Path $buildOutputDir 'ProductUiTests.exe') `
            -WorkingDirectory $buildOutputDir

        foreach ($viewerGroup in @(Get-RSViewerTestGroups)) {
            $plan += New-RSTestRunPlanEntry `
                -Id $viewerGroup.Id `
                -Name $viewerGroup.Name `
                -Kind 'Executable' `
                -Path (Join-Path $buildOutputDir $viewerGroup.Executable) `
                -Arguments @("--group=$($viewerGroup.Group)") `
                -WorkingDirectory $buildOutputDir `
                -RequiresInteractiveDesktop:$viewerGroup.Interactive
        }

        $fullExecutableIds = [ordered]@{
            FileSystemCurlTests = 'standalone.filesystem-curl'
            MonitorTest = 'standalone.monitor'
            LocalizationTests = 'standalone.localization'
            RedConfigureTests = 'standalone.red-configure'
            PluginContractTests = 'standalone.plugin-contract'
            SettingsSchemaTests = 'standalone.settings-schema'
            CrashHandlingTests = 'standalone.crash-handling'
        }
        foreach ($exeName in $fullExecutableIds.Keys) {
            $plan += New-RSTestRunPlanEntry `
                -Id $fullExecutableIds[$exeName] `
                -Name $exeName `
                -Kind 'Executable' `
                -Path (Join-Path $buildOutputDir "$exeName.exe") `
                -WorkingDirectory $buildOutputDir
        }

        $plan += New-RSTestRunPlanEntry `
            -Id 'perf.monitor-etw-latency' `
            -Name 'RedSalamanderMonitorEtwLatency' `
            -Kind 'Executable' `
            -Path (Join-Path $buildOutputDir 'RedSalamanderMonitor.exe') `
            -Arguments @(
                '--chrome-selftest',
                '--wait-instance',
                '--perf',
                '--monitor-etw-burst-mode=latency',
                '--monitor-etw-burst-count=60',
                '--monitor-etw-burst-size=260'
            ) `
            -WorkingDirectory $RepoRoot `
            -RequiresInteractiveDesktop

        $plan += New-RSTestRunPlanEntry `
            -Id 'perf.performance-tests2' `
            -Name 'PerformanceTests2' `
            -Kind 'CppUnitTest' `
            -Path (Join-Path $buildOutputDir 'PerformanceTests2.dll') `
            -WorkingDirectory $buildOutputDir

        $plan += New-RSTestRunPlanEntry `
            -Id 'tooling.pester' `
            -Name 'ToolsPesterTests' `
            -Kind 'Pester' `
            -Path (Join-Path $RepoRoot 'Tools\TerminalEngine\CurrentToolingProfile.Tests.ps1') `
            -Arguments @('-ExcludeTag', 'RequiresBuildToolchain') `
            -WorkingDirectory $RepoRoot

        $plan += New-RSTestRunPlanEntry `
            -Id 'tooling.vcpkg-merge-synthetic' `
            -Name 'VcpkgMergeSynthetic' `
            -Kind 'PowerShellScript' `
            -Path (Join-Path $RepoRoot 'Tests\vcpkg-merge-synthetic-test.ps1') `
            -WorkingDirectory $RepoRoot
    }

    [void](Assert-RSTestRunPlanEntries -Entries $plan)
    return $plan
}

Export-ModuleMember -Function @(
    'Get-RSCommandsSelfTestFamilies',
    'Get-RSDxUiTestSuites',
    'Get-RSViewerTestGroups',
    'Get-RSTestRunPlan'
)
