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

    It 'passes repeat and shuffle controls to in-product selftests' {
        $arguments = Get-RSSelfTestArguments `
            -Flag '--commands-selftest' `
            -TimeoutMultiplier 2.0 `
            -RepeatCount 5 `
            -ShuffleSeed '0xC0FFEE' `
            -CaseFilter 'cmd_focus_'

        Assert-RSSequenceEqual `
            -Actual @($arguments) `
            -Expected @(
                '--commands-selftest',
                '--selftest-timeout-multiplier=2',
                '--selftest-case=cmd_focus_',
                '--selftest-repeat=5',
                '--selftest-shuffle=0xC0FFEE'
            ) `
            -Message 'Self-test flake detector controls should be forwarded as native self-test arguments.'
    }

    It 'passes explicit injected classifier proof hooks to in-product selftests' {
        $arguments = Get-RSSelfTestArguments `
            -Flag '--commands-selftest' `
            -TimeoutMultiplier 2.0 `
            -SelfTestFlakyProofCase 'cmd_flaky_probe' `
            -SelfTestOrderProofCase 'cmd_order_probe'

        Assert-RSEqual `
            -Actual ($arguments -contains '--selftest-flaky-proof-case=cmd_flaky_probe') `
            -Expected $true `
            -Message 'The runner should forward the explicit flaky proof hook only when requested.'
        Assert-RSEqual `
            -Actual ($arguments -contains '--selftest-order-proof-case=cmd_order_probe') `
            -Expected $true `
            -Message 'The runner should forward the explicit order-dependent proof hook only when requested.'
    }

    It 'formats fractional timeout multipliers with invariant decimal separators' {
        $arguments = Get-RSSelfTestArguments `
            -Flag '--commands-selftest' `
            -TimeoutMultiplier 0.1

        Assert-RSEqual `
            -Actual ($arguments -contains '--selftest-timeout-multiplier=0.1') `
            -Expected $true `
            -Message 'Native self-test timeout multipliers require invariant dot decimal formatting.'
    }

    It 'forwards seeded shuffle to all explicit-order self-test suites' {
        $plan = Get-RSTestRunPlan `
            -Suite 'All' `
            -RepoRoot $repoRoot `
            -Platform 'x64' `
            -Configuration 'Debug' `
            -RedSalamanderExePath 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
            -RepeatCount 5 `
            -ShuffleSeed '12345'

        foreach ($entry in $plan) {
            Assert-RSEqual -Actual ($entry.Arguments -contains '--selftest-repeat=5') -Expected $true -Message "$($entry.Name) should receive repeat coverage."
        }

        $commands = $plan | Where-Object { $_.Name -eq 'Commands' } | Select-Object -First 1
        $compare = $plan | Where-Object { $_.Name -eq 'CompareDirectories' } | Select-Object -First 1
        $fileOps = $plan | Where-Object { $_.Name -eq 'FileOperations' } | Select-Object -First 1

        Assert-RSEqual -Actual ($commands.Arguments -contains '--selftest-shuffle=12345') -Expected $true -Message 'Commands should receive the seeded shuffle flag.'
        Assert-RSEqual -Actual ($compare.Arguments -contains '--selftest-shuffle=12345') -Expected $true -Message 'CompareDirectories should receive the seeded shuffle flag after explicit-order migration.'
        Assert-RSEqual -Actual ($fileOps.Arguments -contains '--selftest-shuffle=12345') -Expected $true -Message 'FileOperations should receive the seeded shuffle flag after explicit-order migration.'
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
                'PluginContractTests',
                'SettingsSchemaTests',
                'CrashHandlingTests',
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
            -Expected @('--chrome-selftest', '--wait-instance', '--perf', '--monitor-etw-burst-mode=latency', '--monitor-etw-burst-count=60', '--monitor-etw-burst-size=260') `
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

    It 'defines Suite CI as the GitHub Actions gate through the unified runner' {
        $plan = Get-RSTestRunPlan `
            -Suite 'CI' `
            -RepoRoot $repoRoot `
            -Platform 'x64' `
            -Configuration 'Debug' `
            -RedSalamanderExePath 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
            -TimeoutMultiplier 2.0

        Assert-RSSequenceEqual `
            -Actual @($plan | ForEach-Object { $_.Name }) `
            -Expected @(
                'CompareDirectories',
                'Commands',
                'FileOperations',
                'DxUiTests.Grid',
                'DxUiTests.Theme',
                'DxUiTests.Control',
                'DxUiTests.Menu',
                'DxUiTests.NewControls',
                'DxUiTests.TextField',
                'DxUiTests.NativeTextInput',
                'DxUiTests.ComboBox',
                'DxUiTests.WindowHost',
                'DxUiTests.Tree',
                'DxUiTests.MultilineText',
                'DxUiTests.ReadOnly',
                'DxUiTests.Tooltip',
                'DxUiTests.Rendering',
                'DxUiTests.Animation',
                'DxUiTests.Accessibility',
                'FileSystemCurlTests',
                'ViewerPETests',
                'ViewerPETests.TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly',
                'ViewerPETests.TestViewerTextGotoPromptUsesDxUiHostAndClosesCleanly',
                'ViewerSqliteTests',
                'MonitorTest',
                'LocalizationTests',
                'RedConfigureTests',
                'PluginContractTests',
                'SettingsSchemaTests',
                'CrashHandlingTests',
                'PerformanceTests2',
                'ToolsPesterTests',
                'VcpkgMergeSynthetic'
            ) `
            -Message 'Suite CI should preserve the current PR gate, split DxUiTests by suite, and include the deterministic contract/crash executables.'

        $commands = $plan | Where-Object { $_.Name -eq 'Commands' } | Select-Object -First 1
        Assert-RSEqual -Actual ($commands.Arguments -contains '--selftest-timeout-multiplier=2') -Expected $true -Message 'CI selftests should receive the timeout multiplier.'

        $dxUiMenu = $plan | Where-Object { $_.Name -eq 'DxUiTests.Menu' } | Select-Object -First 1
        Assert-RSSequenceEqual -Actual @($dxUiMenu.Arguments) -Expected @('--suite=Menu') -Message 'CI should run DxUi Menu in its own process.'

        $viewerFind = $plan | Where-Object { $_.Name -eq 'ViewerPETests.TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly' } | Select-Object -First 1
        Assert-RSSequenceEqual `
            -Actual @($viewerFind.Arguments) `
            -Expected @('TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly') `
            -Message 'CI should preserve the explicit ViewerPE find prompt case.'

        $pester = $plan | Where-Object { $_.Name -eq 'ToolsPesterTests' } | Select-Object -First 1
        Assert-RSSequenceEqual `
            -Actual @($pester.Arguments) `
            -Expected @('-ExcludeTag', 'RequiresBuildToolchain') `
            -Message 'CI should keep artifact-only Pester tests off the build-toolchain case.'

        Assert-RSEqual `
            -Actual @($plan | Where-Object { $_.Name -eq 'RedSalamanderMonitorEtwLatency' }).Count `
            -Expected 0 `
            -Message 'Monitor ETW latency should remain Full-only.'

        $workflow = Get-Content -Path (Join-Path $repoRoot '.github\workflows\ci.yml') -Raw
        $arm64BuildOnlyPattern = '(?ms)^  build-arm64:\r?\n    uses: \./\.github/workflows/build-reusable\.yml\r?\n    with:\r?\n      configuration: Debug\r?\n      platform: ARM64\r?\n      upload_build_output: false(?:\r?\n|$)'
        Assert-RSEqual `
            -Actual ([regex]::IsMatch($workflow, $arm64BuildOnlyPattern)) `
            -Expected $true `
            -Message 'The PR workflow should keep an explicit Debug ARM64 build-only lane through the reusable build workflow.'
    }

    It 'builds Invoke-Pester parameters for Pester 3 and newer parameter sets' {
        $pester3 = [pscustomobject]@{
            Parameters = @{
                Script = $true
                PassThru = $true
                ExcludeTag = $true
            }
        }
        $pester3Parameters = New-RSPesterInvokeParameters `
            -Path 'C:\repo\Tools\Tests' `
            -Arguments @('-ExcludeTag', 'RequiresBuildToolchain') `
            -InvokePesterCommand $pester3

        Assert-RSEqual -Actual $pester3Parameters.ContainsKey('Script') -Expected $true -Message 'Pester 3 should receive -Script.'
        Assert-RSEqual -Actual $pester3Parameters.ContainsKey('Path') -Expected $false -Message 'Pester 3 should not receive unsupported -Path.'
        Assert-RSEqual -Actual $pester3Parameters['Script'] -Expected 'C:\repo\Tools\Tests' -Message 'Pester 3 script path should be preserved.'
        Assert-RSEqual -Actual $pester3Parameters['PassThru'] -Expected $true -Message 'Pester should return result objects to the runner.'
        Assert-RSEqual -Actual $pester3Parameters['ExcludeTag'][0] -Expected 'RequiresBuildToolchain' -Message 'Pester excluded tags should be explicitly bound.'

        $pesterNewer = [pscustomobject]@{
            Parameters = @{
                Path = $true
                PassThru = $true
                Tag = $true
            }
        }
        $pesterNewerParameters = New-RSPesterInvokeParameters `
            -Path 'C:\repo\Tools\Tests' `
            -Arguments @('-Tag', 'Smoke') `
            -InvokePesterCommand $pesterNewer

        Assert-RSEqual -Actual $pesterNewerParameters.ContainsKey('Path') -Expected $true -Message 'Newer Pester should receive -Path.'
        Assert-RSEqual -Actual $pesterNewerParameters.ContainsKey('Script') -Expected $false -Message 'Newer Pester should not receive legacy -Script when -Path exists.'
        Assert-RSEqual -Actual $pesterNewerParameters['Path'] -Expected 'C:\repo\Tools\Tests' -Message 'Newer Pester path should be preserved.'
        Assert-RSEqual -Actual $pesterNewerParameters['Tag'][0] -Expected 'Smoke' -Message 'Pester tags should be explicitly bound.'
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

        $ciBuild = Get-RSBuildScriptArguments -Suite 'CI' -Configuration 'Debug' -Platform 'x64'

        Assert-RSEqual -Actual $ciBuild['Configuration'] -Expected 'Debug' -Message 'Suite CI should pass the requested configuration.'
        Assert-RSEqual -Actual $ciBuild['Platform'] -Expected 'x64' -Message 'Suite CI should pass the requested platform.'
        Assert-RSEqual -Actual $ciBuild.ContainsKey('ProjectName') -Expected $false -Message 'Suite CI should build the solution so standalone tests and CppUnitTest DLLs exist.'

        $allEnvironment = Get-RSBuildEnvironmentOverrides -Suite 'All'
        Assert-RSEqual -Actual $allEnvironment.ContainsKey('RSBuildEnableTests') -Expected $false -Message 'Suite All should not force monitor-specific test hooks.'

        $ciEnvironment = Get-RSBuildEnvironmentOverrides -Suite 'CI'
        Assert-RSEqual -Actual $ciEnvironment.ContainsKey('RSBuildEnableTests') -Expected $false -Message 'Suite CI should match the artifact-only GitHub gate and not force monitor-specific test hooks.'

        $fullEnvironment = Get-RSBuildEnvironmentOverrides -Suite 'Full'
        Assert-RSEqual -Actual $fullEnvironment['RSBuildEnableTests'] -Expected 'true' -Message 'Suite Full should build Monitor with selftest hooks for executable monitor drills.'
    }

    It 'uses the unified test sandbox root when locating runner artifacts' {
        $defaultContext = New-RSTestRunContext `
            -RepoRoot 'C:\repo' `
            -RunId '20260706T120000Z-42-abcdef' `
            -TestRootOverride '' `
            -SelfTestRootOverride '' `
            -LocalAppDataRoot 'C:\Users\eric\AppData\Local'
        Assert-RSEqual `
            -Actual $defaultContext.TestRoot `
            -Expected 'C:\repo\.build\TestSandbox' `
            -Message 'Default test sandbox should be repo-local.'
        Assert-RSEqual `
            -Actual $defaultContext.ArtifactRoot `
            -Expected 'C:\repo\.build\TestSandbox\runs\20260706T120000Z-42-abcdef\artifacts\selftest\last_run' `
            -Message 'Native last_run artifacts should be bridged under the per-run sandbox.'

        $sandboxRoot = Get-RSTestSandboxRoot `
            -RepoRoot 'C:\repo' `
            -TestRootOverride 'C:\Temp\RedSalamanderSandbox' `
            -SelfTestRootOverride '' `
            -LocalAppDataRoot 'C:\Users\eric\AppData\Local'
        Assert-RSEqual `
            -Actual $sandboxRoot `
            -Expected 'C:\Temp\RedSalamanderSandbox' `
            -Message 'REDSALAMANDER_TEST_ROOT should select the sandbox base.'

        $bridgedSandboxRoot = Get-RSTestSandboxRoot `
            -RepoRoot 'C:\repo' `
            -TestRootOverride 'C:\repo\.build\TestSandbox' `
            -SelfTestRootOverride 'C:\repo\.build\TestSandbox\runs\20260706T120000Z-42-abcdef\artifacts\selftest' `
            -LocalAppDataRoot 'C:\Users\eric\AppData\Local'
        Assert-RSEqual `
            -Actual $bridgedSandboxRoot `
            -Expected 'C:\repo\.build\TestSandbox' `
            -Message 'The runner-owned per-run legacy self-test bridge should be compatible with REDSALAMANDER_TEST_ROOT.'
    }

    It 'rejects conflicting legacy and unified test roots' {
        $threw = $false
        try {
            Get-RSTestSandboxRoot `
                -RepoRoot 'C:\repo' `
                -TestRootOverride 'C:\Temp\RedSalamanderSandbox' `
                -SelfTestRootOverride 'D:\OtherSelfTestRoot' `
                -LocalAppDataRoot 'C:\Users\eric\AppData\Local' | Out-Null
        } catch {
            $threw = $true
            $_.Exception.Message | Should Match 'REDSALAMANDER_TEST_ROOT'
            $_.Exception.Message | Should Match 'REDSALAMANDER_SELFTEST_ROOT'
        }

        Assert-RSEqual -Actual $threw -Expected $true -Message 'Conflicting root variables should fail fast.'
    }

    It 'creates tooling Pester scratch directories under the unified TestSandbox run root' {
        $oldTestRoot = $env:REDSALAMANDER_TEST_ROOT
        $oldRunId = $env:REDSALAMANDER_TEST_RUN_ID
        $testRoot = Join-Path $repoRoot '.build\TestSandboxUnitProof'
        $runId = 'pester-helper-unit-run'
        try {
            $env:REDSALAMANDER_TEST_ROOT = $testRoot
            $env:REDSALAMANDER_TEST_RUN_ID = $runId

            $scratch = New-RSTestSandboxScratchDirectory `
                -RepoRoot $repoRoot `
                -Harness 'tools-pester' `
                -Case 'helper-proof'

            $expected = Join-Path (Join-Path (Join-Path (Join-Path $testRoot 'runs') $runId) 'scratch') 'tools-pester\helper-proof'
            Assert-RSEqual `
                -Actual $scratch `
                -Expected ([System.IO.Path]::GetFullPath($expected)) `
                -Message 'Pester scratch roots should live below REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\<harness>\<case>.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath $scratch) -Expected $true -Message 'Scratch helper should create the case directory.'

            { New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness '..' -Case 'escape' } |
                Should Throw 'single filesystem path segment'
        } finally {
            if ([string]::IsNullOrWhiteSpace($oldTestRoot)) {
                Remove-Item Env:REDSALAMANDER_TEST_ROOT -ErrorAction SilentlyContinue
            } else {
                $env:REDSALAMANDER_TEST_ROOT = $oldTestRoot
            }

            if ([string]::IsNullOrWhiteSpace($oldRunId)) {
                Remove-Item Env:REDSALAMANDER_TEST_RUN_ID -ErrorAction SilentlyContinue
            } else {
                $env:REDSALAMANDER_TEST_RUN_ID = $oldRunId
            }

            Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'plans legacy test-sandbox cleanup targets without resolving wildcards' {
        $plan = @(Get-RSTestSandboxLegacyCleanupPlan `
                -LocalAppDataRoot 'C:\Users\tester\AppData\Local' `
                -TempRoot 'C:\Users\tester\AppData\Local\Temp' `
                -DriveRoots @('C:\', 'D:\'))

        Assert-RSEqual `
            -Actual ($plan | Where-Object { $_.Path -eq 'C:\Users\tester\AppData\Local\RedSalamander\SelfTest' -and $_.PathType -eq 'Literal' } | Measure-Object).Count `
            -Expected 1 `
            -Message 'Legacy %LOCALAPPDATA% self-test artifacts should be targeted as one literal cleanup root.'

        foreach ($pattern in @(
                'RedConfigure*',
                'CrashHandlingTests-*',
                'ViewerSqliteTests*',
                'RedSalamander.ViewerTextPerf.*',
                'RedSalamander.ViewerTextDiffPerf.*',
                'RedSalamander.ViewerImgRawClose.*',
                'RedSalamander.ViewerWebClose.*',
                'RedSalamander_FolderIconEnumerationPerf_*',
                'RedSalamander_FolderIconEnumerationDuplicatePathPerf_*',
                'RedSalamander_FolderViewRefreshDuplicatePathPerf_*',
                'RSWingetValidationTest_*',
                'RSVcRuntimeTest_*',
                'rs-vcpkg-install-root-test*',
                'rs-vcpkg-single-file-merge-*',
                'rs-show-perfruns-tests-*'
            )) {
            Assert-RSEqual `
                -Actual ($plan | Where-Object { $_.Path -eq "C:\Users\tester\AppData\Local\Temp\$pattern" -and $_.PathType -eq 'Wildcard' } | Measure-Object).Count `
                -Expected 1 `
                -Message "Legacy temp cleanup should include $pattern."
        }

        Assert-RSEqual `
            -Actual ($plan | Where-Object { $_.Path -eq 'C:\RedSalamanderCrossVolumeSelfTest_*' -and $_.PathType -eq 'Wildcard' } | Measure-Object).Count `
            -Expected 1 `
            -Message 'Cross-volume FileOps roots should be targeted on every supplied fixed drive root.'
        Assert-RSEqual `
            -Actual ($plan | Where-Object { $_.Path -eq 'D:\RedSalamanderCrossVolumeSelfTest_*' -and $_.PathType -eq 'Wildcard' } | Measure-Object).Count `
            -Expected 1 `
            -Message 'Cross-volume FileOps cleanup should not be hardcoded to C:.'

        $cleanupScript = Join-Path $repoRoot 'Tools\Clean-TestSandbox.ps1'
        $cleanupRoot = Join-Path $repoRoot '.build\TestSandboxCleanupFailureProof'
        $localAppDataRoot = Join-Path $cleanupRoot 'LocalAppData'
        $legacyRoot = Join-Path $localAppDataRoot 'RedSalamander\SelfTest'
        $lockedPath = Join-Path $legacyRoot 'last_run\commands\hang_dumps\locked.dmp'
        $lockedStream = $null
        try {
            New-Item -ItemType Directory -Path (Split-Path -Parent $lockedPath) -Force | Out-Null
            $lockedStream = [System.IO.File]::Open(
                $lockedPath,
                [System.IO.FileMode]::Create,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None)

            $cleanupResult = @(& $cleanupScript `
                    -Apply `
                    -Confirm:$false `
                    -LocalAppDataRoot $localAppDataRoot `
                    -TempRoot '' `
                    -DriveRoots @($cleanupRoot) `
                    -WarningAction SilentlyContinue)

            $legacyResult = $cleanupResult | Where-Object { $_.Path -eq ([System.IO.Path]::GetFullPath($legacyRoot)) } | Select-Object -First 1
            Assert-RSEqual -Actual ($null -ne $legacyResult) -Expected $true -Message 'Locked legacy cleanup target should still be reported.'
            Assert-RSEqual -Actual $legacyResult.Status -Expected 'Failed' -Message 'Locked legacy cleanup target should not abort the cleanup script.'
            Assert-RSEqual -Actual ([string]::IsNullOrWhiteSpace($legacyResult.Error)) -Expected $false -Message 'Failed cleanup rows should include the removal error.'
        } finally {
            if ($null -ne $lockedStream) {
                $lockedStream.Dispose()
            }
            Remove-Item -LiteralPath $cleanupRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'audits TestSandbox disk state for stale runs and legacy roots' {
        $auditRoot = Join-Path $repoRoot '.build\TestSandboxDiskAuditProof'
        $testRoot = Join-Path $auditRoot 'TestSandbox'
        $localAppDataRoot = Join-Path $auditRoot 'LocalAppData'
        $tempRoot = Join-Path $auditRoot 'Temp'
        $driveRoot = Join-Path $auditRoot 'DriveRoot'
        $runId = 'current-run'
        try {
            New-Item -ItemType Directory -Path (Join-Path (Join-Path $testRoot 'runs') $runId) -Force | Out-Null

            $cleanAudit = Get-RSTestSandboxDiskAudit `
                -TestRoot $testRoot `
                -RunId $runId `
                -LocalAppDataRoot $localAppDataRoot `
                -TempRoot $tempRoot `
                -DriveRoots @($driveRoot)

            Assert-RSEqual -Actual $cleanAudit.is_clean -Expected $true -Message 'Current run directory alone should be clean.'
            Assert-RSEqual -Actual $cleanAudit.issue_count -Expected 0 -Message 'Clean audit should report no disk issues.'
            Assert-RSEqual -Actual $cleanAudit.schema -Expected 'red-salamander.test-sandbox-disk-audit.v1' -Message 'Disk audit schema should be versioned.'

            New-Item -ItemType Directory -Path (Join-Path (Join-Path $testRoot 'runs') 'stale-run') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $testRoot 'rogue-child') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $localAppDataRoot 'RedSalamander\SelfTest') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $tempRoot 'RedConfigureAuditProof') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $driveRoot 'RedSalamanderCrossVolumeSelfTest_audit') -Force | Out-Null

            $dirtyAudit = Get-RSTestSandboxDiskAudit `
                -TestRoot $testRoot `
                -RunId $runId `
                -LocalAppDataRoot $localAppDataRoot `
                -TempRoot $tempRoot `
                -DriveRoots @($driveRoot)

            Assert-RSEqual -Actual $dirtyAudit.is_clean -Expected $false -Message 'Stale run and legacy roots should make the disk audit dirty.'
            foreach ($category in @(
                    'unexpected-test-run-dir',
                    'unexpected-test-root-child',
                    'legacy-selftest-root',
                    'legacy-temp-root',
                    'legacy-cross-volume-root'
                )) {
                Assert-RSEqual `
                    -Actual (@($dirtyAudit.issues | Where-Object { $_.category -eq $category }) | Measure-Object).Count `
                    -Expected 1 `
                    -Message "Disk audit should report $category."
            }
            Assert-RSEqual `
                -Actual (@($dirtyAudit.issues | Where-Object { $_.path -like "*$runId*" }) | Measure-Object).Count `
                -Expected 0 `
                -Message 'Current run directory should be allowed by the disk audit.'
        } finally {
            Remove-Item -LiteralPath $auditRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'sweeps stale TestSandbox run directories whose owner process is dead' {
        $cleanupRoot = Join-Path $repoRoot '.build\TestSandboxStaleRunCleanupProof'
        $testRoot = Join-Path $cleanupRoot 'TestSandbox'
        $runsRoot = Join-Path $testRoot 'runs'
        $currentRunId = '20260706T120000Z-4000-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
        $deadRunId = '20260706T110000Z-1234-bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb'
        $liveRunId = '20260706T111000Z-5678-cccccccccccccccccccccccccccccccc'
        $allowedRunId = '20260706T112000Z-9012-dddddddddddddddddddddddddddddddd'
        $deadFallbackRunId = 'viewer-pe-2468-12345678'
        $liveFallbackRunId = 'redconfigure-1357-87654321'
        $reusedPidFallbackRunId = 'viewer-sqlite-5678-11223344'
        $manualRunId = 'manual-run'
        try {
            foreach ($runId in @($currentRunId, $deadRunId, $liveRunId, $allowedRunId, $deadFallbackRunId, $liveFallbackRunId, $reusedPidFallbackRunId, $manualRunId)) {
                New-Item -ItemType Directory -Path (Join-Path $runsRoot $runId) -Force | Out-Null
            }
            Set-Content -LiteralPath (Join-Path (Join-Path $runsRoot $deadRunId) 'artifact.txt') -Value 'stale' -Encoding UTF8
            (Get-Item -LiteralPath (Join-Path $runsRoot $reusedPidFallbackRunId)).CreationTimeUtc = (Get-Date).ToUniversalTime().AddDays(-2)
            $liveProcessStartTimesUtc = @{
                1357 = (Get-Date).ToUniversalTime().AddHours(-1)
                5678 = (Get-Date).ToUniversalTime().AddDays(-1)
            }

            $targets = @(Resolve-RSTestSandboxStaleRunTargets `
                    -TestRoot $testRoot `
                    -RunId $currentRunId `
                    -AllowedRunIds @($allowedRunId) `
                    -LiveProcessIds @(1357, 4000, 5678, 9012) `
                    -LiveProcessStartTimesUtc $liveProcessStartTimesUtc)

            Assert-RSEqual -Actual @($targets).Count -Expected 3 -Message 'Dead-owner and PID-reused runner/harness-fallback sibling run dirs should be selected.'
            Assert-RSEqual -Actual @($targets | Where-Object { $_.RunId -eq $deadRunId }).Count -Expected 1 -Message 'The dead runner-owned process run should be selected for cleanup.'
            Assert-RSEqual -Actual @($targets | Where-Object { $_.RunId -eq $deadFallbackRunId }).Count -Expected 1 -Message 'The dead direct-harness fallback run should be selected for cleanup.'
            Assert-RSEqual -Actual @($targets | Where-Object { $_.RunId -eq $reusedPidFallbackRunId }).Count -Expected 1 -Message 'An older fallback run must not be protected by an unrelated process that reused its PID.'
            Assert-RSEqual -Actual ($targets | Where-Object { $_.RunId -eq $deadRunId } | Select-Object -First 1).ProcessId -Expected 1234 -Message 'The runner-id parser should expose the owning PID.'
            Assert-RSEqual -Actual ($targets | Where-Object { $_.RunId -eq $deadFallbackRunId } | Select-Object -First 1).ProcessId -Expected 2468 -Message 'The fallback-id parser should expose the owning PID.'
            Assert-RSEqual -Actual @($targets | Where-Object { $_.Category -ne 'stale-test-run-dir' }).Count -Expected 0 -Message 'Cleanup targets should be categorized for reporting.'

            $cleanupResults = @(Remove-RSTestSandboxStaleRunDirectories `
                    -TestRoot $testRoot `
                    -RunId $currentRunId `
                    -AllowedRunIds @($allowedRunId) `
                    -LiveProcessIds @(1357, 4000, 5678, 9012) `
                    -LiveProcessStartTimesUtc $liveProcessStartTimesUtc)

            Assert-RSEqual -Actual @($cleanupResults).Count -Expected 3 -Message 'Dead runner, dead direct-harness, and PID-reused siblings should be removed.'
            Assert-RSEqual -Actual @($cleanupResults | Where-Object { $_.Status -ne 'Removed' }).Count -Expected 0 -Message 'Successful stale run cleanup should report Removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $deadRunId)) -Expected $false -Message 'Dead-PID stale run directory should be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $deadFallbackRunId)) -Expected $false -Message 'Dead-PID direct-harness fallback run directory should be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $reusedPidFallbackRunId)) -Expected $false -Message 'PID-reused direct-harness fallback run directory should be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $currentRunId)) -Expected $true -Message 'Current run directory must never be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $liveRunId)) -Expected $true -Message 'Live-PID sibling run directory must not be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $liveFallbackRunId)) -Expected $true -Message 'Live-PID direct-harness fallback run directory must not be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $allowedRunId)) -Expected $true -Message 'Explicitly allowed sibling run directory must not be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $manualRunId)) -Expected $true -Message 'Unparseable manual run directory must not be removed by the dead-PID sweeper.'
        } finally {
            Remove-Item -LiteralPath $cleanupRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
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

    It 'threads optional perf budget gates into native self-test entries only' {
        $budgetPath = 'C:\repo\Specs\Testing\FolderViewPerfBudgets.json5'
        $plan = Get-RSTestRunPlan `
            -Suite 'Full' `
            -RepoRoot $repoRoot `
            -Platform 'x64' `
            -Configuration 'Debug' `
            -RedSalamanderExePath 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
            -PerfBudgetPath $budgetPath `
            -RequirePerfBudgets

        foreach ($entry in @($plan | Where-Object { $_.Kind -eq 'SelfTest' })) {
            Assert-RSEqual -Actual ($entry.Arguments -contains "--selftest-perf-budget=$budgetPath") -Expected $true -Message "$($entry.Name) should receive the perf budget path."
            Assert-RSEqual -Actual ($entry.Arguments -contains '--selftest-require-perf-budgets') -Expected $true -Message "$($entry.Name) should receive the strict perf budget flag."
        }

        foreach ($entry in @($plan | Where-Object { $_.Kind -ne 'SelfTest' })) {
            Assert-RSEqual -Actual ($entry.Arguments -contains "--selftest-perf-budget=$budgetPath") -Expected $false -Message "$($entry.Name) should not receive native self-test budget arguments."
            Assert-RSEqual -Actual ($entry.Arguments -contains '--selftest-require-perf-budgets') -Expected $false -Message "$($entry.Name) should not receive native self-test budget arguments."
        }
    }

    It 'classifies pass-on-rerun standalone failures as blocking flaky results' {
        $result = @{
            Name = 'DxUiTests.Menu'
            Kind = 'Executable'
            ExitCode = 1
            Passed = 0
            Failed = 1
            Skipped = 0
            Failures = @(
                [pscustomobject]@{ name = 'DxUiTests.Menu'; status = 'failed'; reason = 'Process exited with code 1.'; durationMs = 10 }
            )
        }

        $classified = Add-RSTestResultClassification `
            -Result $result `
            -RetryResults @(
                [pscustomobject]@{ name = 'DxUiTests.Menu'; mode = 'entry'; exit_code = 0; passed = $true }
            )

        Assert-RSEqual -Actual $classified.Classification -Expected 'FLAKY' -Message 'Standalone pass-on-rerun should be classified as flaky.'
        Assert-RSEqual -Actual $classified.ExitCode -Expected 1 -Message 'Flaky classifications must remain blocking.'
        Assert-RSEqual -Actual @($classified.RetryAttempts).Count -Expected 1 -Message 'Classification should preserve retry evidence.'
        Assert-RSEqual -Actual $classified.RetryAttempts[0].mode -Expected 'entry' -Message 'Standalone retry evidence should identify entry-level reruns.'
    }

    It 'classifies fail-again failures as blocking regressions' {
        $result = @{
            Name = 'ViewerPETests'
            Kind = 'Executable'
            ExitCode = 1
            Passed = 0
            Failed = 1
            Skipped = 0
            Failures = @(
                [pscustomobject]@{ name = 'ViewerPETests'; status = 'failed'; reason = 'Process exited with code 1.'; durationMs = 10 }
            )
        }

        $classified = Add-RSTestResultClassification `
            -Result $result `
            -RetryResults @(
                [pscustomobject]@{ name = 'ViewerPETests'; mode = 'entry'; exit_code = 1; passed = $false }
            )

        Assert-RSEqual -Actual $classified.Classification -Expected 'REGRESSION' -Message 'Fail-again retry evidence should be classified as regression.'
        Assert-RSEqual -Actual $classified.ExitCode -Expected 1 -Message 'Regression classifications must remain blocking.'
    }

    It 'routes broad self-test isolated-pass evidence to blocking isolation suspect instead of flaky' {
        $result = @{
            Name = 'Commands'
            Kind = 'SelfTest'
            ExitCode = 1
            Passed = 787
            Failed = 1
            Skipped = 0
            Failures = @(
                [pscustomobject]@{ name = 'cmd_focus_sensitive_case'; status = 'failed'; reason = 'focus mismatch'; durationMs = 10 }
            )
        }

        $classified = Add-RSTestResultClassification `
            -Result $result `
            -RetryResults @(
                [pscustomobject]@{ name = 'cmd_focus_sensitive_case'; mode = 'failed-case'; exit_code = 0; passed = $true }
            )

        Assert-RSEqual -Actual $classified.Classification -Expected 'ISOLATION_SUSPECT' -Message 'Broad self-test isolated-pass evidence needs shuffle triage, not a flaky label.'
        Assert-RSEqual -Actual $classified.ExitCode -Expected 1 -Message 'Isolation suspects must remain blocking.'
    }

    It 'classifies isolated-pass self-test failures as blocking flaky only after shuffle triage passes' {
        $result = @{
            Name = 'Commands'
            Kind = 'SelfTest'
            ExitCode = 1
            Passed = 787
            Failed = 1
            Skipped = 0
            Failures = @(
                [pscustomobject]@{ name = 'cmd_focus_sensitive_case'; status = 'failed'; reason = 'focus mismatch'; durationMs = 10 }
            )
        }

        $classified = Add-RSTestResultClassification `
            -Result $result `
            -RetryResults @(
                [pscustomobject]@{ name = 'cmd_focus_sensitive_case'; mode = 'failed-case'; exit_code = 0; passed = $true },
                [pscustomobject]@{ name = 'Commands'; mode = 'shuffle-triage'; exit_code = 0; passed = $true; shuffle_seed = '123' },
                [pscustomobject]@{ name = 'Commands'; mode = 'shuffle-triage'; exit_code = 0; passed = $true; shuffle_seed = '456' },
                [pscustomobject]@{ name = 'Commands'; mode = 'shuffle-triage'; exit_code = 0; passed = $true; shuffle_seed = '789' }
            )

        Assert-RSEqual -Actual $classified.Classification -Expected 'FLAKY' -Message 'Pass-all-shuffle triage should label the suite as a blocking flaky repair item.'
        Assert-RSEqual -Actual ($classified.ClassificationReason -match 'shuffle triage passed') -Expected $true -Message 'Flaky classification should cite shuffle triage evidence.'
        Assert-RSEqual -Actual $classified.ExitCode -Expected 1 -Message 'Flaky classifications must remain blocking.'
    }

    It 'classifies failed shuffle triage as a blocking regression and preserves the seed in summaries' {
        $result = @{
            Name = 'Commands'
            Kind = 'SelfTest'
            ExitCode = 1
            Passed = 787
            Failed = 1
            Skipped = 0
            Failures = @(
                [pscustomobject]@{ name = 'cmd_focus_sensitive_case'; status = 'failed'; reason = 'focus mismatch'; durationMs = 10 }
            )
        }

        $classified = Add-RSTestResultClassification `
            -Result $result `
            -RetryResults @(
                [pscustomobject]@{ name = 'cmd_focus_sensitive_case'; mode = 'failed-case'; exit_code = 0; passed = $true },
                [pscustomobject]@{ name = 'Commands'; mode = 'shuffle-triage'; exit_code = 1; passed = $false; shuffle_seed = '456'; reason = 'Process exited with code 1.' }
            )

        Assert-RSEqual -Actual $classified.Classification -Expected 'REGRESSION' -Message 'Fail-any-shuffle triage should remain a blocking regression.'
        Assert-RSEqual -Actual ($classified.ClassificationReason -match 'shuffle triage reproduced') -Expected $true -Message 'Regression classification should cite shuffle triage reproduction.'

        $summaryRetry = Convert-RSTestRetryAttemptForRunSummary -RetryAttempt $classified.RetryAttempts[1]
        Assert-RSEqual -Actual $summaryRetry.shuffle_seed -Expected '456' -Message 'Run summaries should preserve the shuffle seed that reproduced the failure.'
    }

    It 'keeps focused self-test filters on shuffle triage instead of falling back to entry retry' {
        $entry = New-RSTestRunPlanEntry `
            -Name 'Commands' `
            -Kind 'SelfTest' `
            -Path 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
            -Arguments @(
                '--commands-selftest',
                '--selftest-timeout-multiplier=2',
                '--selftest-case=cmd_alpha,cmd_beta',
                '--selftest-repeat=5',
                '--selftest-flaky-proof-case=cmd_beta'
            ) `
            -WorkingDirectory 'C:\repo' `
            -JsonName 'commands'

        $retryPlan = @(Get-RSSelfTestClassificationRetryPlan `
                -Entry $entry `
                -FailureNames @('cmd_beta') `
                -ShuffleTriageSeeds @('111', '222', '333'))

        Assert-RSEqual -Actual @($retryPlan).Count -Expected 4 -Message 'Focused self-test failures should still produce failed-case plus shuffle-triage retries.'
        Assert-RSEqual -Actual $retryPlan[0].mode -Expected 'failed-case' -Message 'First retry should isolate the failed case.'
        Assert-RSEqual -Actual ($retryPlan[0].arguments -contains '--selftest-case=cmd_beta') -Expected $true -Message 'Failed-case retry should run the exact failed case.'
        Assert-RSEqual -Actual ($retryPlan[0].arguments -contains '--selftest-case=cmd_alpha,cmd_beta') -Expected $false -Message 'Failed-case retry should replace the focused subset filter.'

        foreach ($triage in @($retryPlan | Select-Object -Skip 1)) {
            Assert-RSEqual -Actual $triage.mode -Expected 'shuffle-triage' -Message 'Subsequent retries should be shuffle triage.'
            Assert-RSEqual -Actual ($triage.arguments -contains '--selftest-case=cmd_alpha,cmd_beta') -Expected $true -Message 'Shuffle triage should preserve the original focused subset to keep runtime proof cheap.'
            Assert-RSEqual -Actual (@($triage.arguments | Where-Object { $_ -like '--selftest-repeat=*' }).Count) -Expected 0 -Message 'Shuffle triage should not inherit repeat expansion.'
            Assert-RSEqual -Actual ($triage.arguments -contains '--selftest-flaky-proof-case=cmd_beta') -Expected $true -Message 'Proof hooks should flow through retry invocations.'
        }
        Assert-RSSequenceEqual `
            -Actual @($retryPlan | Select-Object -Skip 1 | ForEach-Object { $_.shuffle_seed }) `
            -Expected @('111', '222', '333') `
            -Message 'Shuffle triage should preserve deterministic proof seeds.'
    }

    It 'accepts only reviewed unexpired quarantine entries and keeps them blocking' {
        $entries = @(
            [pscustomobject]@{
                harness = 'DxUiTests.Menu'
                name = 'DxUiTests.Menu'
                owner = 'ui-runtime-owner'
                opened = '2026-07-06'
                expires = '2026-07-20'
                issue = 'Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md#dxui-menu'
                root_cause_hypothesis = 'Foreground denial drops popup capture on hosted runners.'
                fix_or_replace_plan = 'Gate foreground-only assertions behind deterministic interactive-desktop probe.'
            }
        )

        $status = Get-RSTestQuarantineStatus -Entries $entries -Now ([datetime]'2026-07-10T00:00:00Z')

        Assert-RSEqual -Actual $status.IsValid -Expected $true -Message 'Reviewed quarantine metadata should validate before expiry.'
        Assert-RSEqual -Actual $status.HasBlockingEntries -Expected $true -Message 'Any active quarantine entry must keep the aggregate run red.'
        Assert-RSEqual -Actual @($status.ActiveEntries).Count -Expected 1 -Message 'Active quarantine entries should be reported.'
        Assert-RSEqual -Actual $status.ActiveEntries[0].owner -Expected 'ui-runtime-owner' -Message 'Active quarantine entry should preserve owner.'
    }

    It 'rejects anonymous expired or malformed quarantine entries' {
        $entries = @(
            [pscustomobject]@{
                harness = 'Commands'
                name = 'cmd_focus_sensitive_case'
                owner = ''
                opened = '2026-07-01'
                expires = '2026-07-05'
                issue = ''
                root_cause_hypothesis = ''
                fix_or_replace_plan = ''
            }
        )

        $status = Get-RSTestQuarantineStatus -Entries $entries -Now ([datetime]'2026-07-10T00:00:00Z')

        Assert-RSEqual -Actual $status.IsValid -Expected $false -Message 'Anonymous or expired quarantine entries should fail validation.'
        Assert-RSEqual -Actual $status.HasBlockingEntries -Expected $true -Message 'Invalid quarantine entries should keep the aggregate run red.'
        Assert-RSEqual -Actual (@($status.Errors) -join "`n" -match 'owner') -Expected $true -Message 'Missing owner should be reported.'
        Assert-RSEqual -Actual (@($status.Errors) -join "`n" -match 'expired') -Expected $true -Message 'Expired entry should be reported.'
    }

    It 'builds quarantine repair attempts only for matching harness adapters' {
        $plan = @(
            (New-RSTestRunPlanEntry `
                    -Name 'Commands' `
                    -Kind 'SelfTest' `
                    -Path 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
                    -Arguments @('--commands-selftest', '--selftest-timeout-multiplier=2') `
                    -WorkingDirectory 'C:\repo' `
                    -JsonName 'commands'),
            (New-RSTestRunPlanEntry `
                    -Name 'DxUiTests.Menu' `
                    -Kind 'Executable' `
                    -Path 'C:\repo\.build\x64\Debug\DxUiTests.exe' `
                    -Arguments @('--suite=Menu') `
                    -WorkingDirectory 'C:\repo')
        )
        $status = [pscustomobject]@{
            IsValid = $true
            HasBlockingEntries = $true
            ActiveEntries = @(
                [pscustomobject]@{
                    harness = 'Commands'
                    name = 'cmd_focus_sensitive_case'
                    owner = 'commands-owner'
                    opened = '2026-07-06'
                    expires = '2026-07-20'
                    issue = 'Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md#commands'
                    root_cause_hypothesis = 'Shared focus state leaks between in-suite cases.'
                    fix_or_replace_plan = 'Replace shared focus state with case fixture ownership.'
                },
                [pscustomobject]@{
                    harness = 'DxUiTests.Menu'
                    name = 'DxUiTests.Menu'
                    owner = 'dxui-owner'
                    opened = '2026-07-06'
                    expires = '2026-07-20'
                    issue = 'Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md#dxui'
                    root_cause_hypothesis = 'Foreground denial drops popup capture on hosted runners.'
                    fix_or_replace_plan = 'Gate foreground-only assertions behind deterministic interactive-desktop probe.'
                }
            )
            InvalidEntries = @()
            Errors = @()
        }

        $repairPlan = Get-RSTestQuarantineRepairPlan -QuarantineStatus $status -TestPlan $plan

        Assert-RSEqual -Actual $repairPlan.IsValid -Expected $true -Message 'Matching quarantine entries should produce a valid repair plan.'
        Assert-RSEqual -Actual @($repairPlan.Attempts).Count -Expected 2 -Message 'Both active entries should get repair attempts.'
        Assert-RSEqual -Actual $repairPlan.Attempts[0].harness -Expected 'Commands' -Message 'Self-test repair attempt should preserve harness.'
        Assert-RSEqual -Actual $repairPlan.Attempts[0].name -Expected 'cmd_focus_sensitive_case' -Message 'Self-test repair attempt should preserve case name.'
        Assert-RSEqual -Actual $repairPlan.Attempts[0].mode -Expected 'selftest-case' -Message 'Self-test repair attempt should run the single case.'
        Assert-RSEqual -Actual ($repairPlan.Attempts[0].arguments -contains '--commands-selftest') -Expected $true -Message 'Self-test repair attempt should preserve suite flag.'
        Assert-RSEqual -Actual ($repairPlan.Attempts[0].arguments -contains '--selftest-case=cmd_focus_sensitive_case') -Expected $true -Message 'Self-test repair attempt should add the quarantined case filter.'
        Assert-RSEqual -Actual $repairPlan.Attempts[0].owner -Expected 'commands-owner' -Message 'Repair attempt should preserve owner metadata.'
        Assert-RSEqual -Actual $repairPlan.Attempts[1].harness -Expected 'DxUiTests.Menu' -Message 'Standalone repair attempt should preserve harness.'
        Assert-RSEqual -Actual $repairPlan.Attempts[1].mode -Expected 'entry' -Message 'Standalone repair attempt should rerun the whole entry.'
        Assert-RSEqual -Actual ($repairPlan.Attempts[1].arguments -contains '--suite=Menu') -Expected $true -Message 'Standalone repair attempt should preserve entry arguments.'
    }

    It 'rejects quarantine entries that do not match a runnable harness adapter' {
        $plan = @(
            (New-RSTestRunPlanEntry `
                    -Name 'Commands' `
                    -Kind 'SelfTest' `
                    -Path 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
                    -Arguments @('--commands-selftest') `
                    -WorkingDirectory 'C:\repo' `
                    -JsonName 'commands')
        )
        $status = [pscustomobject]@{
            IsValid = $true
            HasBlockingEntries = $true
            ActiveEntries = @(
                [pscustomobject]@{
                    harness = 'MissingHarness'
                    name = 'missing_case'
                    owner = 'owner'
                    opened = '2026-07-06'
                    expires = '2026-07-20'
                    issue = 'Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md#missing'
                    root_cause_hypothesis = 'Unknown.'
                    fix_or_replace_plan = 'Fix the quarantine entry.'
                }
            )
            InvalidEntries = @()
            Errors = @()
        }

        $repairPlan = Get-RSTestQuarantineRepairPlan -QuarantineStatus $status -TestPlan $plan

        Assert-RSEqual -Actual $repairPlan.IsValid -Expected $false -Message 'Unknown harness quarantine entries should invalidate the repair plan.'
        Assert-RSEqual -Actual @($repairPlan.Attempts).Count -Expected 0 -Message 'Unknown harness entries should not produce attempts.'
        Assert-RSEqual -Actual @($repairPlan.InvalidEntries).Count -Expected 1 -Message 'Unknown harness entries should be reported as invalid.'
        Assert-RSEqual -Actual (@($repairPlan.Errors) -join "`n" -match 'no matching case') -Expected $true -Message 'Unknown harness errors should explain the missing adapter.'
    }

    It 'rejects self-test quarantine entries whose case name is absent from the adapter case list' {
        $plan = @(
            (New-RSTestRunPlanEntry `
                    -Name 'Commands' `
                    -Kind 'SelfTest' `
                    -Path 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
                    -Arguments @('--commands-selftest') `
                    -WorkingDirectory 'C:\repo' `
                    -JsonName 'commands')
        )
        $status = [pscustomobject]@{
            IsValid = $true
            HasBlockingEntries = $true
            ActiveEntries = @(
                [pscustomobject]@{
                    harness = 'Commands'
                    name = 'missing_case'
                    owner = 'owner'
                    opened = '2026-07-06'
                    expires = '2026-07-20'
                    issue = 'Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md#missing'
                    root_cause_hypothesis = 'Unknown.'
                    fix_or_replace_plan = 'Fix the quarantine entry.'
                }
            )
            InvalidEntries = @()
            Errors = @()
        }

        $repairPlan = Get-RSTestQuarantineRepairPlan `
            -QuarantineStatus $status `
            -TestPlan $plan `
            -HarnessCaseMap @{ Commands = @('known_case') }

        Assert-RSEqual -Actual $repairPlan.IsValid -Expected $false -Message 'Unknown self-test case names should invalidate the repair plan.'
        Assert-RSEqual -Actual @($repairPlan.Attempts).Count -Expected 0 -Message 'Unknown self-test cases should not produce attempts.'
        Assert-RSEqual -Actual @($repairPlan.InvalidEntries).Count -Expected 1 -Message 'Unknown self-test cases should be reported as invalid.'
        Assert-RSEqual -Actual (@($repairPlan.Errors) -join "`n" -match 'missing_case') -Expected $true -Message 'Unknown self-test case errors should name the missing case.'
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

        $repeated = Test-RSSelfTestResultCoverage `
            -ExpectedCaseNames @('alpha', 'beta') `
            -ExpectedRepeatCount 2 `
            -ActualCases @(
                [pscustomobject]@{ name = 'alpha'; status = 'passed'; repeat_index = 1 },
                [pscustomobject]@{ name = 'beta'; status = 'passed'; repeat_index = 1 },
                [pscustomobject]@{ name = 'alpha'; status = 'passed'; repeat_index = 2 },
                [pscustomobject]@{ name = 'beta'; status = 'passed'; repeat_index = 2 }
            )

        Assert-RSEqual -Actual $repeated.IsValid -Expected $true -Message 'Repeated self-test results should validate when every expected case has the requested repeat count.'
        Assert-RSSequenceEqual -Actual $repeated.DuplicateActual -Expected @() -Message 'Expected repeats should not be reported as duplicate drift.'

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
            -TestRoot 'C:\repo\.build\TestSandbox' `
            -RunId '20260706T120000Z-42-abcdef' `
            -QuarantineStatus ([pscustomobject]@{
                IsValid = $true
                HasBlockingEntries = $true
                ActiveEntries = @([pscustomobject]@{ harness = 'DxUiTests.Menu'; name = 'DxUiTests.Menu'; owner = 'ui-runtime-owner'; expires = '2026-07-20' })
                InvalidEntries = @()
                Errors = @()
            }) `
            -QuarantineRepairAttempts @(
                [pscustomobject]@{
                    harness = 'Commands'
                    name = 'cmd_beta'
                    mode = 'selftest-case'
                    owner = 'ui-runtime-owner'
                    expires = '2026-07-20'
                    exit_code = 1
                    passed = $false
                    duration_ms = 44
                    reason = 'Repair lane reproduced the quarantined failure.'
                }
            ) `
            -TestSandboxAudit ([pscustomobject]@{
                schema = 'red-salamander.test-sandbox-disk-audit.v1'
                is_clean = $false
                issue_count = 1
                test_root = 'C:\repo\.build\TestSandbox'
                run_id = '20260706T120000Z-42-abcdef'
                allowed_run_ids = @('20260706T120000Z-42-abcdef')
                issues = @([pscustomobject]@{
                        category = 'unexpected-test-run-dir'
                        path = 'C:\repo\.build\TestSandbox\runs\stale-run'
                        reason = 'Only the current runner-owned run directory should remain during post-run audit.'
                    })
            }) `
            -Results @(
                @{
                    Name = 'CompareDirectories'
                    Kind = 'SelfTest'
                    ExitCode = 0
                    WallMs = 100
                    Passed = 2
                    Failed = 0
                    Skipped = 1
                    Classification = 'PASSED'
                    RetryAttempts = @()
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
                    Classification = 'FLAKY'
                    RetryAttempts = @(
                        [pscustomobject]@{ name = 'cmd_beta'; mode = 'failed-case'; exit_code = 0; passed = $true }
                    )
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
        Assert-RSEqual -Actual $summary.test_root -Expected 'C:\repo\.build\TestSandbox' -Message 'Aggregate artifact should record the canonical test root.'
        Assert-RSEqual -Actual $summary.run_id -Expected '20260706T120000Z-42-abcdef' -Message 'Aggregate artifact should record the per-run sandbox id.'
        Assert-RSEqual -Actual $summary.classifications.flaky -Expected 1 -Message 'Aggregate artifact should count flaky classifications.'
        Assert-RSEqual -Actual $summary.classifications.regression -Expected 0 -Message 'Aggregate artifact should count regression classifications.'
        Assert-RSEqual -Actual $summary.quarantine.active_count -Expected 1 -Message 'Aggregate artifact should count active quarantine entries.'
        Assert-RSEqual -Actual $summary.quarantine.has_blocking_entries -Expected $true -Message 'Aggregate artifact should preserve quarantine blocking status.'
        Assert-RSEqual -Actual $summary.quarantine.repair_attempt_count -Expected 1 -Message 'Aggregate artifact should count quarantine repair attempts.'
        Assert-RSEqual -Actual $summary.quarantine.repair_attempts[0].name -Expected 'cmd_beta' -Message 'Aggregate artifact should preserve repair attempt case name.'
        Assert-RSEqual -Actual $summary.quarantine.repair_attempts[0].passed -Expected $false -Message 'Aggregate artifact should preserve repair attempt result.'
        Assert-RSEqual -Actual $summary.test_sandbox_audit.issue_count -Expected 1 -Message 'Aggregate artifact should preserve TestSandbox disk-audit issue count.'
        Assert-RSEqual -Actual $summary.test_sandbox_audit.issues[0].category -Expected 'unexpected-test-run-dir' -Message 'Aggregate artifact should preserve TestSandbox disk-audit issues.'
        Assert-RSEqual -Actual $summary.total -Expected 5 -Message 'Aggregate artifact should total all suite cases.'
        Assert-RSEqual -Actual $summary.failed -Expected 1 -Message 'Aggregate artifact should count failures.'
        Assert-RSEqual -Actual @($summary.suites).Count -Expected 2 -Message 'Aggregate artifact should preserve both suites.'
        Assert-RSEqual -Actual $summary.suites[0].output_log_path -Expected 'C:\runs\CompareDirectories.output.log' -Message 'Aggregate artifact should preserve suite output logs.'
        Assert-RSEqual -Actual $summary.suites[1].classification -Expected 'FLAKY' -Message 'Aggregate artifact should preserve suite classification.'
        Assert-RSEqual -Actual @($summary.suites[1].retry_attempts).Count -Expected 1 -Message 'Aggregate artifact should preserve retry evidence.'
        Assert-RSEqual -Actual @($summary.suites[0].cases).Count -Expected 2 -Message 'Aggregate artifact should preserve Compare cases.'
        Assert-RSEqual -Actual $summary.suites[1].failures[0].reason -Expected 'state mismatch' -Message 'Aggregate artifact should preserve failure reasons.'
    }

    It 'builds per-case history JSONL and a slow-case dashboard from aggregate summaries' {
        $summary = New-RSTestRunSummary `
            -Suite 'CI' `
            -Platform 'x64' `
            -Configuration 'Debug' `
            -ExePath 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
            -ArtifactRoot 'C:\repo\.build\TestSandbox\runs\run-1\artifacts\selftest\last_run' `
            -RunStartedUtc ([datetime]'2026-07-06T10:00:00Z') `
            -RunEndedUtc ([datetime]'2026-07-06T10:05:00Z') `
            -DurationMs 300000 `
            -ExitCode 1 `
            -TimeoutMultiplier 2.0 `
            -TestRoot 'C:\repo\.build\TestSandbox' `
            -RunId 'run-1' `
            -Results @(
                [pscustomobject]@{
                    Name = 'Commands'
                    Kind = 'SelfTest'
                    ExitCode = 1
                    WallMs = 1200
                    Passed = 1
                    Failed = 1
                    Skipped = 0
                    Classification = 'REGRESSION'
                    ClassificationReason = 'Shuffle triage reproduced a failure.'
                    ShuffleSeed = '456'
                    Cases = @(
                        [pscustomobject]@{ name = 'cmd_alpha'; status = 'passed'; reason = ''; durationMs = 25; repeat_index = 1 },
                        [pscustomobject]@{ name = 'cmd_beta'; status = 'failed'; reason = 'focus mismatch'; durationMs = 400; repeat_index = 1 }
                    )
                    RetryAttempts = @(
                        [pscustomobject]@{ name = 'Commands'; mode = 'shuffle-triage'; exit_code = 1; passed = $false; duration_ms = 900; shuffle_seed = '789'; reason = 'reproduced' }
                    )
                    Failures = @(
                        [pscustomobject]@{ name = 'cmd_beta'; status = 'failed'; reason = 'focus mismatch'; durationMs = 400 }
                    )
                }
            )

        $rows = @(Convert-RSTestRunSummaryToCaseHistoryRows -Summary $summary)
        Assert-RSEqual -Actual @($rows).Count -Expected 3 -Message 'History should include case rows and retry rows.'
        Assert-RSEqual -Actual $rows[1].harness -Expected 'Commands' -Message 'History row should preserve harness.'
        Assert-RSEqual -Actual $rows[1].case -Expected 'cmd_beta' -Message 'History row should preserve case name.'
        Assert-RSEqual -Actual $rows[1].duration_ms -Expected 400 -Message 'History row should preserve case duration.'
        Assert-RSEqual -Actual $rows[1].classification -Expected 'REGRESSION' -Message 'History row should preserve suite classification.'
        Assert-RSEqual -Actual $rows[1].seed -Expected '456' -Message 'History row should preserve suite shuffle seed.'
        Assert-RSEqual -Actual $rows[1].attempt -Expected 1 -Message 'History row should preserve repeat attempt.'
        Assert-RSEqual -Actual $rows[2].source -Expected 'retry' -Message 'Retry evidence should be distinguishable from case rows.'
        Assert-RSEqual -Actual $rows[2].seed -Expected '789' -Message 'Retry history should preserve the reproducing shuffle seed.'
        Assert-RSEqual -Actual $rows[2].reason -Expected 'reproduced' -Message 'Retry history should preserve the retry reason.'

        $jsonl = Convert-RSTestCaseHistoryRowsToJsonl -Rows $rows
        ($jsonl -split "`n")[0] | ConvertFrom-Json | ForEach-Object {
            Assert-RSEqual -Actual $_.schema -Expected 'red-salamander.case-history.v1' -Message 'JSONL rows should carry a schema.'
        }

        $dashboard = Convert-RSTestCaseHistoryRowsToDashboardMarkdown -Rows $rows -Summary $summary -TopCount 2
        Assert-RSEqual -Actual ($dashboard -match 'cmd_beta') -Expected $true -Message 'Dashboard should include the slow failing case.'
        Assert-RSEqual -Actual ($dashboard -match 'REGRESSION') -Expected $true -Message 'Dashboard should include classifications.'
        Assert-RSEqual -Actual ($dashboard -match '789') -Expected $true -Message 'Dashboard should include retry shuffle seeds.'
        Assert-RSEqual -Actual ($dashboard -match 'reproduced') -Expected $true -Message 'Dashboard should include retry reasons.'
    }

    It 'formats a GitHub step summary with blocking classification and quarantine repair evidence' {
        $summary = New-RSTestRunSummary `
            -Suite 'CI' `
            -Platform 'x64' `
            -Configuration 'Debug' `
            -ExePath 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
            -ArtifactRoot 'C:\repo\.build\TestSandbox\runs\20260706T120000Z-42-abcdef\artifacts\selftest\last_run' `
            -RunStartedUtc ([datetime]'2026-07-06T12:00:00Z') `
            -RunEndedUtc ([datetime]'2026-07-06T12:02:00Z') `
            -DurationMs 120000 `
            -ExitCode 1 `
            -TimeoutMultiplier 2.0 `
            -TestRoot 'C:\repo\.build\TestSandbox' `
            -RunId '20260706T120000Z-42-abcdef' `
            -QuarantineStatus ([pscustomobject]@{
                IsValid = $true
                HasBlockingEntries = $true
                ActiveEntries = @(
                    [pscustomobject]@{
                        harness = 'Commands'
                        name = 'cmd_beta'
                        owner = 'ui-runtime-owner'
                        expires = '2026-07-20'
                        issue = 'Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md#commands'
                    }
                )
                InvalidEntries = @()
                Errors = @()
            }) `
            -QuarantineRepairAttempts @(
                [pscustomobject]@{
                    harness = 'Commands'
                    name = 'cmd_beta'
                    mode = 'selftest-case'
                    owner = 'ui-runtime-owner'
                    expires = '2026-07-20'
                    issue = 'Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md#commands'
                    exit_code = 1
                    passed = $false
                    duration_ms = 44
                    reason = 'Repair lane reproduced the quarantined failure.'
                }
            ) `
            -Results @(
                @{
                    Name = 'Commands'
                    Kind = 'SelfTest'
                    ExitCode = 1
                    WallMs = 200
                    Passed = 1
                    Failed = 1
                    Skipped = 0
                    Classification = 'FLAKY'
                    ClassificationReason = 'Initial attempt failed, but retry evidence passed; this remains blocking.'
                    RetryAttempts = @(
                        [pscustomobject]@{ name = 'cmd_beta'; mode = 'failed-case'; exit_code = 0; passed = $true }
                    )
                    Cases = @()
                    Failures = @(
                        [pscustomobject]@{ name = 'cmd_beta'; status = 'failed'; reason = 'state mismatch'; durationMs = 30 }
                    )
                }
            )

        $markdown = Convert-RSTestRunSummaryToGitHubStepSummary -Summary $summary

        $markdown | Should Match 'RedSalamander Test Summary'
        $markdown | Should Match 'FLAKY'
        $markdown | Should Match 'flaky=1'
        $markdown | Should Match 'ui-runtime-owner'
        $markdown | Should Match '2026-07-20'
        $markdown | Should Match 'cmd_beta'
        $markdown | Should Match 'REPAIR FAIL'
        $markdown | Should Match 'Repair lane reproduced'
    }
}
