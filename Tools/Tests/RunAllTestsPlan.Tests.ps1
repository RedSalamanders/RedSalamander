Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Tools\TestRunPlan.ps1'

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

Describe 'Run-AllTests plan helper' {
    BeforeAll {
        . $helperScript
    }

    It 'keeps Suite All scoped to the three in-product self-test suites' {
        $plan = Get-RSTestRunPlan `
            -Suite 'All' `
            -RepoRoot $repoRoot `
            -Platform 'x64' `
            -Configuration 'Debug' `
            -RedSalamanderExePath 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
            -TimeoutMultiplier 2.0 `
            -FailFast `
            -CaseFilter 'Smoke'

        Assert-RSSequenceEqual `
            -Actual @($plan | ForEach-Object { $_.Name }) `
            -Expected @('CompareDirectories', 'Commands', 'FileOperations') `
            -Message 'Suite All should remain the existing self-test-only local/CI contract.'

        foreach ($entry in $plan) {
            Assert-RSEqual -Actual $entry.Kind -Expected 'SelfTest' -Message "$($entry.Name) should be a self-test entry."
            Assert-RSEqual -Actual ($entry.Arguments -contains '--selftest-fail-fast') -Expected $true -Message "$($entry.Name) should receive fail-fast."
            Assert-RSEqual -Actual ($entry.Arguments -contains '--selftest-case=Smoke') -Expected $true -Message "$($entry.Name) should receive the case filter."
            Assert-RSEqual -Actual ($entry.Arguments -contains '--selftest-timeout-multiplier=2') -Expected $true -Message "$($entry.Name) should receive the timeout multiplier."
        }
    }

    It 'adds standalone, CppUnitTest, Pester, and fast script tests for Suite Full' {
        $plan = Get-RSTestRunPlan `
            -Suite 'Full' `
            -RepoRoot $repoRoot `
            -Platform 'x64' `
            -Configuration 'Debug' `
            -RedSalamanderExePath 'C:\repo\.build\x64\Debug\RedSalamander.exe'

        Assert-RSSequenceEqual `
            -Actual @($plan | ForEach-Object { $_.Name }) `
            -Expected @(
                'CompareDirectories',
                'Commands',
                'FileOperations',
                'DxUiTests',
                'FileSystemCurlTests',
                'ViewerPETests',
                'ViewerSqliteTests',
                'MonitorTest',
                'LocalizationTests',
                'RedConfigureTests',
                'RedSalamanderMonitorEtwLatency',
                'PerformanceTests2',
                'ToolsPesterTests',
                'VcpkgMergeSynthetic'
            ) `
            -Message 'Suite Full should enumerate every PR-safe local test surface.'

        $performance = $plan | Where-Object { $_.Name -eq 'PerformanceTests2' } | Select-Object -First 1
        Assert-RSEqual -Actual $performance.Kind -Expected 'CppUnitTest' -Message 'PerformanceTests2 should run through vstest.'
        Assert-RSEqual -Actual ($performance.Path -like '*\.build\x64\Debug\PerformanceTests2.dll') -Expected $true -Message 'PerformanceTests2 should target the built DLL.'

        $monitorLatency = $plan | Where-Object { $_.Name -eq 'RedSalamanderMonitorEtwLatency' } | Select-Object -First 1
        Assert-RSEqual -Actual $monitorLatency.Kind -Expected 'Executable' -Message 'Monitor ETW latency drill should run as a focused executable test.'
        Assert-RSEqual -Actual ($monitorLatency.Path -like '*\.build\x64\Debug\RedSalamanderMonitor.exe') -Expected $true -Message 'Monitor ETW latency drill should target the built monitor executable.'
        Assert-RSSequenceEqual `
            -Actual @($monitorLatency.Arguments) `
            -Expected @('--chrome-selftest', '--perf', '--monitor-etw-burst-mode=latency', '--monitor-etw-burst-count=60', '--monitor-etw-burst-size=260') `
            -Message 'Monitor ETW latency drill should replay the scenario-gated perf command.'

        $pester = $plan | Where-Object { $_.Name -eq 'ToolsPesterTests' } | Select-Object -First 1
        Assert-RSEqual -Actual $pester.Kind -Expected 'Pester' -Message 'Tools tests should run through Pester.'
        Assert-RSEqual -Actual ($pester.Path -like '*\Tools\Tests') -Expected $true -Message 'Pester path should target Tools\Tests.'

        $synthetic = $plan | Where-Object { $_.Name -eq 'VcpkgMergeSynthetic' } | Select-Object -First 1
        Assert-RSEqual -Actual $synthetic.Kind -Expected 'PowerShellScript' -Message 'The fast vcpkg merge test should run as a PowerShell script.'
        Assert-RSEqual -Actual ($synthetic.Path -like '*\Tests\vcpkg-merge-synthetic-test.ps1') -Expected $true -Message 'Suite Full should include only the fast vcpkg synthetic script by default.'

        $lockValidation = $plan | Where-Object { $_.Path -like '*vcpkg-merge-lock-validation.ps1' }
        Assert-RSEqual -Actual @($lockValidation).Count -Expected 0 -Message 'The intrusive vcpkg lock validation script should not run in Suite Full by default.'
    }

    It 'passes named build parameters to the build script' {
        $allBuild = Get-RSBuildScriptArguments -Suite 'All' -Configuration 'Debug' -Platform 'x64'

        Assert-RSEqual -Actual $allBuild['Configuration'] -Expected 'Debug' -Message 'Suite All should pass the requested configuration.'
        Assert-RSEqual -Actual $allBuild['Platform'] -Expected 'x64' -Message 'Suite All should pass the requested platform.'
        Assert-RSEqual -Actual $allBuild['ProjectName'] -Expected 'RedSalamander' -Message 'Suite All should build only RedSalamander.'

        $fullBuild = Get-RSBuildScriptArguments -Suite 'Full' -Configuration 'Debug' -Platform 'x64'

        Assert-RSEqual -Actual $fullBuild['Configuration'] -Expected 'Debug' -Message 'Suite Full should pass the requested configuration.'
        Assert-RSEqual -Actual $fullBuild['Platform'] -Expected 'x64' -Message 'Suite Full should pass the requested platform.'
        Assert-RSEqual -Actual $fullBuild.ContainsKey('ProjectName') -Expected $false -Message 'Suite Full should build the solution so standalone tests and CppUnitTest DLLs exist.'

        $allEnvironment = Get-RSBuildEnvironmentOverrides -Suite 'All'
        Assert-RSEqual -Actual $allEnvironment.ContainsKey('RSBuildEnableTests') -Expected $false -Message 'Suite All should not force monitor-specific test hooks.'

        $fullEnvironment = Get-RSBuildEnvironmentOverrides -Suite 'Full'
        Assert-RSEqual -Actual $fullEnvironment['RSBuildEnableTests'] -Expected 'true' -Message 'Suite Full should build Monitor with selftest hooks for executable monitor drills.'
    }

    It 'uses the native self-test root override when locating runner artifacts' {
        $defaultRoot = Get-RSSelfTestArtifactRoot -SelfTestRootOverride '' -LocalAppDataRoot 'C:\Users\eric\AppData\Local'
        Assert-RSEqual `
            -Actual $defaultRoot `
            -Expected 'C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run' `
            -Message 'Default artifacts should stay in the app-local self-test last_run folder.'

        $isolatedRoot = Get-RSSelfTestArtifactRoot -SelfTestRootOverride 'C:\Temp\rs-isolated-selftest' -LocalAppDataRoot 'C:\Users\eric\AppData\Local'
        Assert-RSEqual `
            -Actual $isolatedRoot `
            -Expected 'C:\Temp\rs-isolated-selftest\last_run' `
            -Message 'Runner artifacts should follow REDSALAMANDER_SELFTEST_ROOT so concurrent checkouts cannot overwrite evidence.'
    }

    It 'does not pass self-test-only options to standalone executables or scripts' {
        $plan = Get-RSTestRunPlan `
            -Suite 'Full' `
            -RepoRoot $repoRoot `
            -Platform 'x64' `
            -Configuration 'Debug' `
            -RedSalamanderExePath 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
            -TimeoutMultiplier 3.0 `
            -FailFast `
            -CaseFilter 'OneCase'

        foreach ($entry in @($plan | Where-Object { $_.Kind -ne 'SelfTest' -and $_.Name -ne 'RedSalamanderMonitorEtwLatency' })) {
            $argumentText = @($entry.Arguments) -join ' '
            if ($argumentText -match 'selftest') {
                throw "$($entry.Name) should not receive self-test-only arguments, but got '$argumentText'."
            }
        }
    }

    It 'builds runner-native case-list arguments from self-test entries only' {
        $arguments = Get-RSSelfTestListArguments -SelfTestArguments @(
            '--commands-selftest',
            '--selftest-fail-fast',
            '--selftest-timeout-multiplier=3',
            '--selftest-case=cmd_preferences_'
        )

        Assert-RSSequenceEqual `
            -Actual $arguments `
            -Expected @('--selftest-list-cases', '--commands-selftest', '--selftest-case=cmd_preferences_') `
            -Message 'Case-listing arguments should keep only suite selection and case filter.'
    }

    It 'detects self-test result coverage drift from runner-listed names' {
        $valid = Test-RSSelfTestResultCoverage `
            -ExpectedCaseNames @('alpha', 'beta') `
            -ActualCases @(
                [pscustomobject]@{ name = 'alpha'; status = 'passed' },
                [pscustomobject]@{ name = 'beta'; status = 'skipped' }
            )

        Assert-RSEqual -Actual $valid.IsValid -Expected $true -Message 'Matching expected and actual names should validate.'

        $invalid = Test-RSSelfTestResultCoverage `
            -ExpectedCaseNames @('alpha', 'beta', 'beta') `
            -ActualCases @(
                [pscustomobject]@{ name = 'alpha'; status = 'passed' },
                [pscustomobject]@{ name = 'gamma'; status = 'passed' },
                [pscustomobject]@{ name = 'gamma'; status = 'skipped' }
            )

        Assert-RSEqual -Actual $invalid.IsValid -Expected $false -Message 'Duplicate, missing, and extra names should fail validation.'
        Assert-RSSequenceEqual -Actual $invalid.DuplicateExpected -Expected @('beta') -Message 'Expected duplicates should be reported.'
        Assert-RSSequenceEqual -Actual $invalid.DuplicateActual -Expected @('gamma') -Message 'Actual duplicates should be reported.'
        Assert-RSSequenceEqual -Actual $invalid.Missing -Expected @('beta') -Message 'Missing expected names should be reported.'
        Assert-RSSequenceEqual -Actual $invalid.Extra -Expected @('gamma') -Message 'Extra actual names should be reported.'

        $emptyFiltered = Test-RSSelfTestResultCoverage `
            -ExpectedCaseNames @() `
            -ActualCases @() `
            -RequireExpectedCases

        Assert-RSEqual -Actual $emptyFiltered.IsValid -Expected $false -Message 'Filtered runs with no runner-listed cases should not validate.'
        Assert-RSEqual -Actual $emptyFiltered.NoExpectedCases -Expected $true -Message 'Zero expected cases should be reported explicitly.'
    }

    It 'builds an aggregate runner artifact that preserves every suite summary' {
        $summary = New-RSTestRunSummary `
            -Suite 'All' `
            -Platform 'x64' `
            -Configuration 'Debug' `
            -ExePath 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
            -ArtifactRoot 'C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run' `
            -RunStartedUtc ([datetime]'2026-05-05T17:00:00Z') `
            -RunEndedUtc ([datetime]'2026-05-05T17:00:30Z') `
            -DurationMs 30000 `
            -ExitCode 1 `
            -TimeoutMultiplier 2.0 `
            -CaseFilter 'Phase7_' `
            -Results @(
                @{
                    Name = 'CompareDirectories'
                    Kind = 'SelfTest'
                    ExitCode = 0
                    WallMs = 100
                    Passed = 2
                    Failed = 0
                    Skipped = 1
                    OutputLogPath = 'C:\runs\CompareDirectories.output.log'
                    Cases = @(
                        [pscustomobject]@{ name = 'compare_alpha'; status = 'passed'; durationMs = 10 },
                        [pscustomobject]@{ name = 'compare_remote'; status = 'skipped'; reason = 'missing credentials'; durationMs = 0 }
                    )
                    Failures = @()
                },
                @{
                    Name = 'Commands'
                    Kind = 'SelfTest'
                    ExitCode = 1
                    WallMs = 200
                    Passed = 1
                    Failed = 1
                    Skipped = 0
                    Cases = @(
                        [pscustomobject]@{ name = 'cmd_alpha'; status = 'passed'; durationMs = 20 },
                        [pscustomobject]@{ name = 'cmd_beta'; status = 'failed'; reason = 'state mismatch'; durationMs = 30 }
                    )
                    Failures = @(
                        [pscustomobject]@{ name = 'cmd_beta'; status = 'failed'; reason = 'state mismatch'; durationMs = 30 }
                    )
                }
            )

        Assert-RSEqual -Actual $summary.schema -Expected 'red-salamander.run-all-tests.v1' -Message 'Aggregate artifact schema should be versioned.'
        Assert-RSEqual -Actual $summary.total -Expected 5 -Message 'Aggregate artifact should total all suite cases.'
        Assert-RSEqual -Actual $summary.failed -Expected 1 -Message 'Aggregate artifact should count failures.'
        Assert-RSEqual -Actual @($summary.suites).Count -Expected 2 -Message 'Aggregate artifact should preserve both suites.'
        Assert-RSEqual -Actual $summary.suites[0].output_log_path -Expected 'C:\runs\CompareDirectories.output.log' -Message 'Aggregate artifact should preserve suite output logs.'
        Assert-RSEqual -Actual @($summary.suites[0].cases).Count -Expected 2 -Message 'Aggregate artifact should preserve Compare cases.'
        Assert-RSEqual -Actual $summary.suites[1].failures[0].reason -Expected 'state mismatch' -Message 'Aggregate artifact should preserve failure reasons.'
    }
}
