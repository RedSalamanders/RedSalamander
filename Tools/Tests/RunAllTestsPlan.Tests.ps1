Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
$testSupport = Join-Path $PSScriptRoot 'TestSupport.psm1'
Import-Module $testSupport -Force

Describe 'Run-AllTests runner source contracts' {
    It 'admits all six native profiles through the real runner and deployment preflights' {
        $runner = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $start = $runner.IndexOf('if ($ValidationMode -eq ''Resume'')', [StringComparison]::Ordinal)
        $end = $runner.IndexOf('# --- Helpers ---', $start, [StringComparison]::Ordinal)
        $runnerPreflight = [scriptblock]::Create(@'
param($Platform, $Configuration)
function Exit-RSInvalidRunnerInvocation { param([string]$Message) throw $Message }
$ValidationMode = 'Fresh'; $Suite = 'Full'; $ResumeFrom = ''; $ExplainPlan = $false
$ApplyLegacySandboxCleanup = $false
'@ + "`n" + $runner.Substring($start, $end - $start))
        $deployment = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1') -Raw
        $start = $deployment.IndexOf('$script:selectedPlatform = [Environment]::GetEnvironmentVariable', [StringComparison]::Ordinal)
        $end = $deployment.IndexOf('$script:outputDir =', $start, [StringComparison]::Ordinal)
        # BeforeAll's script variables become local to this isolated invocation;
        # otherwise unrelated Pester caller variables can shadow their values.
        $deploymentPreflight = [scriptblock]::Create(
            $deployment.Substring($start, $end - $start).Replace('$script:selected', '$local:selected'))
        $previousPlatform = $env:RS_VALIDATION_PLATFORM
        $previousConfiguration = $env:RS_VALIDATION_CONFIGURATION
        try {
            foreach ($Platform in @('x64', 'ARM64')) {
                foreach ($Configuration in @('Debug', 'Release', 'ASan Debug')) {
                    $env:RS_VALIDATION_PLATFORM = $Platform
                    $env:RS_VALIDATION_CONFIGURATION = $Configuration
                    & $runnerPreflight $Platform $Configuration
                    & $deploymentPreflight
                }
            }
            $rejected = $false
            try { & $runnerPreflight ARM64 Unknown } catch { $rejected = $true }
            $rejected | Should Be $true
            $env:RS_VALIDATION_CONFIGURATION = 'Unknown'
            $rejected = $false
            try { & $deploymentPreflight } catch { $rejected = $true }
            $rejected | Should Be $true
            $env:RS_VALIDATION_CONFIGURATION = 'Debug'; $env:RS_VALIDATION_PLATFORM = 'Unknown'
            $rejected = $false
            try { & $deploymentPreflight } catch { $rejected = $true }
            $rejected | Should Be $true
        } finally {
            $env:RS_VALIDATION_PLATFORM = $previousPlatform
            $env:RS_VALIDATION_CONFIGURATION = $previousConfiguration
        }
    }

    It 'rejects an executable outside the selected profile before acquiring its artifact lock' {
        $runnerSource = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $validationIndex = $runnerSource.IndexOf(
            'The -ExePath compatibility parameter must name the selected repository profile',
            [StringComparison]::Ordinal)
        $lockIndex = $runnerSource.IndexOf('$artifactOperationLock = $null', [StringComparison]::Ordinal)

        ($validationIndex -ge 0) | Should Be $true
        ($lockIndex -gt $validationIndex) | Should Be $true
        $runnerSource | Should Match '\$expectedExeFullPath\s*=\s*\[IO\.Path\]::GetFullPath'
        $runnerSource | Should Match '\[string\]::Equals\([\s\S]*?\[StringComparison\]::OrdinalIgnoreCase'
        $runnerSource | Should Match 'Refusing uncoordinated executable:'
    }

    It 'restores Startrail module commands after nested build and Pester scope boundaries' {
        $runner = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $runner | Should Match 'function Import-RSRunnerValidationModules'
        ([regex]::Matches($runner, 'Import-RSRunnerValidationModules')).Count | Should BeGreaterThan 2
        $afterReceipt = $runner.IndexOf('if ($null -eq $receipt)', [StringComparison]::Ordinal)
        $afterEntry = $runner.IndexOf('Invoke-RSTestPlanEntry', [StringComparison]::Ordinal)
        $runner.IndexOf('Import-RSRunnerValidationModules', $afterReceipt, [StringComparison]::Ordinal) | Should BeGreaterThan $afterReceipt
        $runner.IndexOf('Import-RSRunnerValidationModules', $afterEntry, [StringComparison]::Ordinal) | Should BeGreaterThan $afterEntry
    }

    It 'binds Resume fingerprints to the validated origin terminal artifact epoch' {
        $runner = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $originRead = $runner.IndexOf('$resolvedOrigin = Resolve-RSValidationRunDirectory', [StringComparison]::Ordinal)
        $stateBind = $runner.IndexOf('$resumeOriginState = $resolvedOrigin.State', [StringComparison]::Ordinal)
        $epochBind = $runner.IndexOf('$artifactEpoch = [int64]$resumeOriginState.artifact_epoch', [StringComparison]::Ordinal)
        $fingerprints = $runner.IndexOf('$entryFingerprintRecord = New-RSRunnerEntryFingerprintRecord', [StringComparison]::Ordinal)
        ($originRead -ge 0) | Should Be $true
        ($stateBind -gt $originRead) | Should Be $true
        ($epochBind -gt $stateBind) | Should Be $true
        ($fingerprints -gt $epochBind) | Should Be $true
    }

    It 'requires an exact suite match and routes every entry through a normalized terminal result' {
        $runner = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $runner | Should Match 'if \(\$suiteName -eq \$sc\.Name -and \$null -ne \$suiteCasesProperty\)'
        $runner | Should Not Match 'if \(\$suiteName -eq \$sc\.Name -or \$null -ne \$suiteCasesProperty\)'
        $runner | Should Match '\$totalCount\s*=\s*\[int64\]\(Get-JsonValue \$pesterResult'
        $runner | Should Match 'New-RSValidationTerminalResult -EntryContract \$portableEntry\[0\]'
        $runner | Should Match "exit \(Get-RSValidationExitCode -TerminalState NOT_EVALUATED\)"
        $runner | Should Match '\$resultArtifactPath\s*=\s*\$null'
        $runner | Should Match '(?s)\$resultArtifactPath\s*=\s*\$jsonPath.+?-CandidateArtifact\s+@\(\$resultArtifactPath\)'
        $runner | Should Not Match '-CandidateArtifact\s+\$jsonCandidates'
        $selfTestCommon = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp') -Raw
        $selfTestCommon | Should Match '"skip_class", CaseSkipClassName\(testCase\)'
        $selfTestCommon | Should Match 'return "declared-capability"'
        $selfTestCommon | Should Match 'return "unexpected"'
    }

    It 'suspends the selected receipt and isolates profile inputs before any artifact writer' {
        $runner = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $deployment = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1') -Raw
        $runner | Should Match '(?s)Suspend-RSBuildReceipt.+?Start-RSValidationArtifactMutation.+?Invoke-RSRunnerStandaloneValidationEntry -Entry \$artifactMutator'
        $runner | Should Match 'Complete-RSValidationArtifactMutation'
        $runner | Should Match 'function Assert-RSRunnerArtifactReadBarrier'
        $runner | Should Match "RS_VALIDATION_PLATFORM"
        $runner | Should Match "RS_VALIDATION_CONFIGURATION"
        $runner | Should Match '(?s)previousProfileEnvironment.+?finally.+?SetEnvironmentVariable'
        $deployment | Should Match "GetEnvironmentVariable\('RS_VALIDATION_PLATFORM'"
        $deployment | Should Match "GetEnvironmentVariable\('RS_VALIDATION_CONFIGURATION'"
        $deployment | Should Match '''\.build\\\{0\}\\\{1\}'' -f \$selectedPlatform, \$selectedConfiguration'
        $deployment | Should Not Match '\.build\\x64\\Debug'
        $deployment | Should Match '''-Configuration'', \$Configuration, ''-Platform'', \$Platform'
        $deployment | Should Match 'Invoke-RSTargetedBuildForDeploymentTest -Platform \$selectedPlatform -Configuration \$selectedConfiguration'
    }
}

Describe 'Run-AllTests plan helper' {
    BeforeAll {
        Import-Module $helperModule -Force -ErrorAction Stop
    }

    It 'leases the desktop only for the exact interactive documentation case' {
        foreach ($case in @('', 'FileOps_VisualGallery', 'FileOps_VisualGalleryInteraction', 'fileops_visualgalleryinteraction', 'FileOps_VisualGalleryInteraction*')) {
            $plan = @(Get-RSTestRunPlan -Suite FileOps -RepoRoot $repoRoot -Platform x64 -Configuration Debug `
                -RedSalamanderExePath (Join-Path $repoRoot '.build\x64\Debug\RedSalamander.exe') -CaseFilter $case)
            $interactive = $case -eq 'FileOps_VisualGalleryInteraction'
            $plan.Count | Should Be 1
            $plan[0].RequiresInteractiveDesktop | Should Be $interactive
            ($plan[0].Arguments -contains '--selftest-no-activate') | Should Be (-not $interactive)
        }
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
        $commands = $plan | Where-Object Name -eq 'Commands' | Select-Object -First 1
        Assert-RSEqual -Actual $commands.RequiresInteractiveDesktop -Expected $true -Message 'Commands requires the real foreground/focus isolation boundary.'
        foreach ($entry in @($plan | Where-Object Name -ne 'Commands')) {
            Assert-RSEqual -Actual $entry.RequiresInteractiveDesktop -Expected $false -Message "$($entry.Name) should run outside the interactive desktop mutex."
            Assert-RSEqual -Actual ($entry.Arguments -contains '--selftest-no-activate') -Expected $true -Message "$($entry.Name) should reject focus stealing."
        }
    }

    It 'assigns unique stable IDs and schema-complete metadata to every CI and Full entry' {
        $arguments = @{
            RepoRoot = $repoRoot
            Platform = 'x64'
            Configuration = 'Debug'
            RedSalamanderExePath = (Join-Path $repoRoot '.build\x64\Debug\RedSalamander.exe')
        }
        foreach ($suite in @('CI', 'Full')) {
            $entries = @(Get-RSTestRunPlan -Suite $suite @arguments)
            Assert-RSEqual -Actual @($entries.Id | Sort-Object -Unique).Count -Expected $entries.Count `
                -Message "$suite plan entry IDs must be unique."
            foreach ($entry in $entries) {
                Assert-RSEqual -Actual ([string]::IsNullOrWhiteSpace($entry.Id)) -Expected $false -Message "$($entry.Name) must have a stable ID."
                Assert-RSEqual -Actual (@($entry.ResourceClasses).Count -gt 0) -Expected $true -Message "$($entry.Id) must declare resources."
                Assert-RSEqual -Actual ([string]::IsNullOrWhiteSpace($entry.CachePolicy)) -Expected $false -Message "$($entry.Id) must declare cache policy."
            }
            $record = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite $suite -Entries $entries
            $json = $record | ConvertTo-Json -Depth 30
            Assert-RSEqual -Actual (Test-Json -Json $json -SchemaFile (Join-Path $repoRoot 'Specs\Testing\ValidationPlan.schema.json')) `
                -Expected $true -Message "$suite validation plan must satisfy its schema."
        }
    }

    It 'schedules the isolated artifact mutator first and re-fingerprints a verified epoch before artifact readers' {
        $full = @(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug `
            -RedSalamanderExePath (Join-Path $repoRoot '.build\x64\Debug\RedSalamander.exe') `
            -TimeoutMultiplier 1 -FailFast:$false -CaseFilter '' -RepeatCount 1 -ShuffleSeed '' `
            -PerfBudgetPath '' -RequirePerfBudgets:$false -SelfTestFlakyProofCase '' -SelfTestOrderProofCase '')
        $full[0].Id | Should Be 'tooling.pester-build-toolchain'
        @($full[0].ArtifactWriteIds).Count | Should BeGreaterThan 0
        @($full | Select-Object -Skip 1 | Where-Object { @($_.ArtifactWriteIds).Count -gt 0 }).Count | Should Be 0
        $full[0].CachePolicy | Should Be 'ExactArtifact'
        $runnerSource = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $runnerSource | Should Match 'Re-attesting profile after'
        $runnerSource | Should Match 'Start-RSValidationArtifactMutation'
        $runnerSource | Should Match 'Complete-RSValidationArtifactMutation'
        $runnerSource | Should Match 'New-RSRunnerEntryFingerprintRecord -ExecutionEntries \$testPlan'
        $runnerSource | Should Match 're-attested profile failed byte-for-byte receipt verification'
        $runnerSource | Should Match '(?s)Re-attesting profile after:.+Get-RSBuildEnvironmentOverrides -Suite \$Suite.+SetEnvironmentVariable\(\$key, \[string\]\$reattestEnvironmentOverrides\[\$key\], ''Process''\).+build\.ps1.+finally.+SetEnvironmentVariable\(\$key, \$previousReattestEnvironment\[\$key\], ''Process''\)'
        $runnerSource | Should Match '(?s)Re-attesting profile after:.+build\.ps1''\) -Configuration \$Configuration -Platform \$Platform.+-MaxCpuCount 4'
    }

    It 'keeps display labels out of plan identity while binding behavior coverage and policy' {
        $entries = @(Get-RSTestRunPlan -Suite Commands -RepoRoot $repoRoot -Platform x64 -Configuration Debug `
                -RedSalamanderExePath (Join-Path $repoRoot '.build\x64\Debug\RedSalamander.exe'))
        $baseline = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Commands -Entries $entries
        $entries[0].Name = 'Commands display renamed'
        $renamed = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Commands -Entries $entries
        Assert-RSEqual -Actual $renamed.plan_digest -Expected $baseline.plan_digest -Message 'Display rename must not alter semantic plan identity.'
        $entries[0].Arguments += '--selftest-repeat=2'
        $behaviorChanged = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Commands -Entries $entries
        Assert-RSEqual -Actual ($behaviorChanged.plan_digest -ne $baseline.plan_digest) -Expected $true -Message 'Argument changes must alter plan identity.'
        $entries[0].Arguments = @($entries[0].Arguments | Where-Object { $_ -ne '--selftest-repeat=2' })
        $entries[0].CachePolicy = 'Never'
        $policyChanged = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Commands -Entries $entries
        Assert-RSEqual -Actual ($policyChanged.plan_digest -ne $baseline.plan_digest) -Expected $true -Message 'Reuse policy changes must alter plan identity.'
        $entries[0].CachePolicy = 'SessionBound'
        $entries[0].CoverageContract.family = 'changed-family'
        $coverageChanged = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Commands -Entries $entries
        Assert-RSEqual -Actual ($coverageChanged.plan_digest -ne $baseline.plan_digest) -Expected $true -Message 'Coverage changes must alter plan identity.'
    }

    It 'constructs the artifact-writer contract for each supported Full profile without a Debug sibling path' {
        foreach ($profile in @(
                @{ Platform = 'x64'; Configuration = 'Release' },
                @{ Platform = 'x64'; Configuration = 'ASan Debug' },
                @{ Platform = 'ARM64'; Configuration = 'Debug' },
                @{ Platform = 'ARM64'; Configuration = 'Release' },
                @{ Platform = 'ARM64'; Configuration = 'ASan Debug' })) {
            $exe = Join-Path $repoRoot ('.build\{0}\{1}\RedSalamander.exe' -f $profile.Platform, $profile.Configuration)
            $plan = @(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot `
                -Platform $profile.Platform -Configuration $profile.Configuration `
                -RedSalamanderExePath $exe)
            $writer = @($plan | Where-Object Id -eq 'tooling.pester-build-toolchain')
            $writer.Count | Should Be 1
            (@($writer[0].EnvironmentInputs) -contains 'env.profile') | Should Be $true
            Assert-RSSequenceEqual -Actual @($writer[0].InvalidatesReceiptScopes) -Expected @('receipt.profile') `
                -Message 'The artifact writer must invalidate the selected profile receipt.'
            @($plan | Where-Object { $_.Path -like '*\.build\x64\Debug\*' }).Count | Should Be 0
        }
    }

    It 'separates stable coverage identity from truthful invocation mode' {
        $entries = @(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug `
                -RedSalamanderExePath (Join-Path $repoRoot '.build\x64\Debug\RedSalamander.exe'))
        $fresh = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries $entries -ValidationMode Fresh
        $resume = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries $entries -ValidationMode Resume
        $fresh.validation_mode | Should Be 'Fresh'
        $resume.validation_mode | Should Be 'Resume'
        $resume.coverage_plan_digest | Should Be $fresh.coverage_plan_digest
        $resume.plan_digest | Should Not Be $fresh.plan_digest

        $runner = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $runner | Should Match '(?s)\$validationPlan\s*=\s*New-RSValidationPlan.+?-ValidationMode\s+\$ValidationMode'
        $runner | Should Not Match '(?s)\$validationPlan\s*=\s*New-RSValidationPlan.+?-ValidationMode\s+Fresh'
        $runner | Should Match '-CoveragePlanDigest \$PortablePlan\.coverage_plan_digest'
        $runner | Should Not Match '-PlanDigest \$PortablePlan\.plan_digest'
    }

    It 'maps each terminal state to its documented distinct process exit in subprocesses' {
        $expected = [ordered]@{
            PASSED = 0
            FAILED = 1
            NOT_EVALUATED = 2
            INVALID_CLI_OR_PLAN = 3
            BLOCKED_ATTESTATION = 4
            CORRUPT_EVIDENCE = 5
        }
        foreach ($pair in $expected.GetEnumerator()) {
            & pwsh -NoLogo -NoProfile -Command `
                "Import-Module '$helperModule' -Force; exit (Get-RSValidationExitCode -TerminalState '$($pair.Key)')"
            $LASTEXITCODE | Should Be $pair.Value
        }
    }

    It 'keeps governed Commands family metadata equal to the native coordinator and emits a family-process plan' {
        $families = @(Get-RSCommandsSelfTestFamilies)
        @($families.name | Sort-Object -Unique).Count | Should Be $families.Count
        $families.Count | Should Be 14
        foreach ($family in $families) {
            Test-Path -LiteralPath (Join-Path $repoRoot $family.source) -PathType Leaf | Should Be $true
        }

        $commandsSource = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp') -Raw
        $tableMatch = [regex]::Match($commandsSource, 'kCommandsSelfTestFamilies\{(?<body>[\s\S]*?)\};')
        $tableMatch.Success | Should Be $true
        $nativeNames = @([regex]::Matches($tableMatch.Groups['body'].Value, 'L"(?<name>[a-z0-9-]+)"') | ForEach-Object { $_.Groups['name'].Value })
        Assert-RSSequenceEqual -Actual $nativeNames -Expected @($families.name) -Message 'PowerShell and native Commands families must have exact ordered equality.'
        foreach ($family in $families) {
            $commandsSource | Should Match ([regex]::Escape("runFamily(L`"$($family.name)`""))
        }

        $entry = @(Get-RSTestRunPlan -Suite Commands -RepoRoot $repoRoot -Platform x64 -Configuration Debug `
                -RedSalamanderExePath (Join-Path $repoRoot '.build\x64\Debug\RedSalamander.exe') `
                -CommandsFamily 'file-operations')
        $entry.Count | Should Be 1
        @($entry[0].Arguments) -contains '--selftest-family=file-operations' | Should Be $true
        $entry[0].CoverageContract.family | Should Be 'commands.file-operations'
        $entry[0].OrderContract | Should Be 'family-process'
        $entry[0].CachePolicy | Should Be 'Never'

        $wrongSuiteThrew = $false
        try {
            Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug `
                -RedSalamanderExePath (Join-Path $repoRoot '.build\x64\Debug\RedSalamander.exe') `
                -CommandsFamily 'file-operations' | Out-Null
        } catch { $wrongSuiteThrew = $true }
        $wrongSuiteThrew | Should Be $true
        $unknownThrew = $false
        try {
            Get-RSTestRunPlan -Suite Commands -RepoRoot $repoRoot -Platform x64 -Configuration Debug `
                -RedSalamanderExePath (Join-Path $repoRoot '.build\x64\Debug\RedSalamander.exe') `
                -CommandsFamily 'unknown' | Out-Null
        } catch { $unknownThrew = $true }
        $unknownThrew | Should Be $true
    }

    It 'rejects missing and duplicate stable IDs' {
        & pwsh -NoLogo -NoProfile -NonInteractive -Command `
            "Import-Module '$helperModule' -Force; try { New-RSTestRunPlanEntry -Name Missing -Kind Executable -Path 'C:\missing.exe' | Out-Null } catch { exit 17 }; exit 0"
        Assert-RSEqual -Actual $LASTEXITCODE -Expected 17 -Message 'Missing stable ID must fail without prompting.'
        $entry = New-RSTestRunPlanEntry -Id 'duplicate.entry' -Name Duplicate -Kind Executable `
            -Path (Join-Path $repoRoot '.build\x64\Debug\Duplicate.exe') -WorkingDirectory $repoRoot
        $duplicateThrew = $false
        try { New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries @($entry, $entry) | Out-Null } catch { $duplicateThrew = $true }
        Assert-RSEqual -Actual $duplicateThrew -Expected $true -Message 'Duplicate stable IDs must fail.'
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
                'ToolsPesterBuildToolchain',
                'CompareDirectories',
                'Commands',
                'FileOperations',
                'ProductUiTests',
                'ViewerPETests.Noninteractive',
                'ViewerPETests.Interactive',
                'ViewerSqliteTests.Noninteractive',
                'ViewerSqliteTests.Interactive',
                'FileSystemCurlTests',
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
        Assert-RSEqual -Actual $monitorLatency.RequiresInteractiveDesktop -Expected $true -Message 'Monitor ETW latency should own the interactive desktop only while it executes.'

        $interactiveNames = @($plan |
                Where-Object RequiresInteractiveDesktop |
                ForEach-Object Name)
        Assert-RSSequenceEqual `
            -Actual $interactiveNames `
            -Expected @('Commands', 'ViewerPETests.Interactive', 'ViewerSqliteTests.Interactive', 'RedSalamanderMonitorEtwLatency') `
            -Message 'Full should serialize only focus-sensitive or desktop-global entries.'

        $pester = $plan | Where-Object { $_.Name -eq 'ToolsPesterTests' } | Select-Object -First 1
        Assert-RSEqual -Actual $pester.Kind -Expected 'Pester' -Message 'Tools tests should run through Pester.'
        Assert-RSEqual `
            -Actual ($pester.Path -like '*\Tools\TerminalEngine\CurrentToolingProfile.Tests.ps1') `
            -Expected $true `
            -Message 'Full should run the current tooling profile and leave the sealed Terminal cohort explicit-only.'
        Assert-RSSequenceEqual -Actual @($pester.Arguments) -Expected @('-ExcludeTag', 'RequiresBuildToolchain') `
            -Message 'Full should isolate ordinary Pester coverage from artifact-mutating build-toolchain cases.'
        $buildPester = $plan | Where-Object { $_.Name -eq 'ToolsPesterBuildToolchain' } | Select-Object -First 1
        Assert-RSSequenceEqual -Actual @($buildPester.Arguments) -Expected @('-Tag', 'RequiresBuildToolchain') `
            -Message 'Full should expose build-toolchain Pester cases as a separate artifact-mutating entry.'
        Assert-RSEqual -Actual ($buildPester.ResourceClasses -contains 'artifact-mutating') -Expected $true `
            -Message 'The build-toolchain Pester entry must declare artifact mutation.'

        $synthetic = $plan | Where-Object { $_.Name -eq 'VcpkgMergeSynthetic' } | Select-Object -First 1
        Assert-RSEqual -Actual $synthetic.Kind -Expected 'PowerShellScript' -Message 'The fast vcpkg merge test should run as a PowerShell script.'
        Assert-RSEqual -Actual ($synthetic.Path -like '*\Tests\vcpkg-merge-synthetic-test.ps1') -Expected $true -Message 'Suite Full should include only the fast vcpkg synthetic script by default.'
        Assert-RSEqual -Actual $pester.RequiresInteractiveDesktop -Expected $false -Message 'Pester must not hold the interactive desktop mutex.'
        Assert-RSEqual -Actual $synthetic.RequiresInteractiveDesktop -Expected $false -Message 'Noninteractive script tests must not hold the interactive desktop mutex.'

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
                'ProductUiTests',
                'FileSystemCurlTests',
                'ViewerPETests.Noninteractive',
                'ViewerPETests.Interactive',
                'ViewerSqliteTests.Noninteractive',
                'ViewerSqliteTests.Interactive',
                'ViewerPETests.TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly',
                'ViewerPETests.TestViewerTextGotoPromptUsesDxUiHostAndClosesCleanly',
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
            -Message 'Suite CI should preserve the current PR gate, run pinned product UI regressions, and include the deterministic contract/crash executables.'

        $commands = $plan | Where-Object { $_.Name -eq 'Commands' } | Select-Object -First 1
        Assert-RSEqual -Actual ($commands.Arguments -contains '--selftest-timeout-multiplier=2') -Expected $true -Message 'CI selftests should receive the timeout multiplier.'

        $productUi = $plan | Where-Object { $_.Name -eq 'ProductUiTests' } | Select-Object -First 1
        Assert-RSEqual -Actual ([IO.Path]::GetFileName($productUi.Path)) -Expected 'ProductUiTests.exe' -Message 'CI executes the surviving product-owned regressions.'
        Assert-RSEqual -Actual $productUi.RequiresInteractiveDesktop -Expected $false -Message 'Product UI utility tests use only message windows and need no activation.'
        Assert-RSEqual -Actual @($plan | Where-Object Name -like 'DxUiTests.*').Count -Expected 0 -Message 'Legacy library tests cannot establish candidate coverage.'

        $viewerFind = $plan | Where-Object { $_.Name -eq 'ViewerPETests.TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly' } | Select-Object -First 1
        Assert-RSSequenceEqual `
            -Actual @($viewerFind.Arguments) `
            -Expected @('TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly') `
            -Message 'CI should preserve the explicit ViewerPE find prompt case.'
        Assert-RSEqual -Actual $viewerFind.RequiresInteractiveDesktop -Expected $true -Message 'Viewer prompt cases require the interactive desktop.'

        foreach ($name in @('CompareDirectories', 'FileOperations')) {
            $entry = $plan | Where-Object Name -eq $name | Select-Object -First 1
            Assert-RSEqual -Actual $entry.RequiresInteractiveDesktop -Expected $false -Message "$name should not own the interactive desktop."
            Assert-RSEqual -Actual ($entry.Arguments -contains '--selftest-no-activate') -Expected $true -Message "$name should reject focus stealing."
        }
        foreach ($name in @('ViewerPETests.Noninteractive', 'ViewerSqliteTests.Noninteractive')) {
            $entry = $plan | Where-Object Name -eq $name | Select-Object -First 1
            Assert-RSEqual -Actual $entry.RequiresInteractiveDesktop -Expected $false -Message "$name should run without desktop serialization."
            Assert-RSEqual -Actual ($entry.Arguments -contains '--group=noninteractive') -Expected $true -Message "$name should select the no-activation group."
        }

        $pester = $plan | Where-Object { $_.Name -eq 'ToolsPesterTests' } | Select-Object -First 1
        Assert-RSEqual `
            -Actual ($pester.Path -like '*\Tools\TerminalEngine\CurrentToolingProfile.Tests.ps1') `
            -Expected $true `
            -Message 'CI should run the current tooling profile and leave the sealed Terminal cohort explicit-only.'
        Assert-RSSequenceEqual `
            -Actual @($pester.Arguments) `
            -Expected @('-ExcludeTag', 'RequiresBuildToolchain') `
            -Message 'CI should keep artifact-only Pester tests off the build-toolchain case.'
        Assert-RSEqual -Actual $pester.RequiresInteractiveDesktop -Expected $false -Message 'CI Pester must run outside the interactive mutex.'

        Assert-RSEqual `
            -Actual @($plan | Where-Object { $_.Name -eq 'RedSalamanderMonitorEtwLatency' }).Count `
            -Expected 0 `
            -Message 'Monitor ETW latency should remain Full-only.'

        $workflow = Get-Content -Path (Join-Path $repoRoot '.github\workflows\ci.yml') -Raw
        # The native matrix is one fromJSON expression: the pull-request branch first, then the push branch.
        $matrixBranches = [System.Collections.Generic.List[object]]::new()
        foreach ($match in [regex]::Matches($workflow, "'(\[\{[^']*\}\])'")) {
            $matrixBranches.Add(@(ConvertFrom-Json -InputObject $match.Groups[1].Value | ForEach-Object { "$($_.platform)|$($_.configuration)|$($_.runner)" }))
        }
        Assert-RSEqual -Actual $matrixBranches.Count -Expected 2 -Message 'ci.yml selects the native matrix by event.'
        $pullRequestProfiles = @($matrixBranches[0])
        $pushProfiles = @($matrixBranches[1])
        Assert-RSSequenceEqual -Actual $pullRequestProfiles `
            -Expected @('x64|Debug|windows-2025-vs2026', 'ARM64|Debug|windows-11-vs2026-arm') `
            -Message 'Pull requests build and test one Debug profile per architecture on native runners.'
        Assert-RSSequenceEqual -Actual $pushProfiles `
            -Expected @('x64|Debug|windows-2025-vs2026', 'x64|Release|windows-2025-vs2026', 'ARM64|Debug|windows-11-vs2026-arm', 'ARM64|Release|windows-11-vs2026-arm') `
            -Message 'Pushes to main add both Release profiles on native runners.'
        Assert-RSEqual -Actual ($workflow -match 'run_full_tests: true') -Expected $true -Message 'Every profile runs the receipt-verified native suite.'
        Assert-RSEqual -Actual ($workflow -match 'test_suite: CI') -Expected $true -Message 'The quick gate runs Suite CI; Suite Full stays the local closeout gate.'

    }

    It 'isolates each FileOperations child profile without changing the caller or other suites' {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path $repoRoot 'Tools/Run-AllTests.ps1'), [ref]$tokens, [ref]$errors)
        @($errors).Count | Should Be 0
        $function = $ast.Find({ param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Invoke-RSSelfTestProcess'
        }, $true)

        & {
            . ([scriptblock]::Create($function.Extent.Text))
            $profiles = [Collections.Generic.List[string]]::new()
            $children = [Collections.Generic.List[object]]::new()
            $closed = [Collections.Generic.List[object]]::new()
            $callerLocalData = [Environment]::GetEnvironmentVariable('LOCALAPPDATA', 'Process')
            function New-RSTestSandboxScratchDirectory {
                param($RepoRoot, $Harness, $Case)
                $Harness | Should Be 'fileops-profile' | Out-Null
                $path = Join-Path $TestDrive $Case
                $profiles.Add($path)
                return $path
            }
            function Test-RSTestEntryRequiresInteractiveDesktop { param($Entry) return $false }
            function Invoke-RSInteractiveDesktopOperation {
                param($RequiresInteractiveDesktop, $Operation)
                & $Operation
            }
            function Start-RSContainedProcess {
                param($ProcessStartInfo)
                $children.Add($ProcessStartInfo)
                $process = [pscustomobject]@{ ExitCode = 17 }
                $process | Add-Member ScriptMethod WaitForExit { }
                return $process
            }
            function Close-RSContainedProcess { param($Process) $closed.Add($Process) }

            $entry = [pscustomobject]@{ Path = 'fixture.exe'; WorkingDirectory = $TestDrive }
            foreach ($attempt in 1..2) {
                Invoke-RSSelfTestProcess -Entry $entry -Arguments @(
                    '--fileops-selftest', '--selftest-case=R4A19_DiscoveryProviderControls') | Should Be 17
            }
            Invoke-RSSelfTestProcess -Entry $entry -Arguments @('--compare-selftest') | Should Be 17
            $profiles.Count | Should Be 2
            $children.Count | Should Be 3
            $closed.Count | Should Be 3
            $profiles[0] | Should Not Be $profiles[1]
            $children[0].EnvironmentVariables['LOCALAPPDATA'] | Should Be $profiles[0]
            $children[1].EnvironmentVariables['LOCALAPPDATA'] | Should Be $profiles[1]
            $children[2].EnvironmentVariables['LOCALAPPDATA'] | Should Be $callerLocalData
            [Environment]::GetEnvironmentVariable('LOCALAPPDATA', 'Process') | Should Be $callerLocalData
        }
    }

    It 'keeps the session-wide desktop mutex inside marked entry execution only' {
        $runnerPath = Join-Path $repoRoot 'Tools\Run-AllTests.ps1'
        $tokens = $null
        $parseErrors = $null
        $runnerAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $runnerPath,
            [ref]$tokens,
            [ref]$parseErrors)
        Assert-RSEqual -Actual @($parseErrors).Count -Expected 0 -Message 'Run-AllTests should parse before its mutex boundary is inspected.'

        $functions = @($runnerAst.FindAll({
                    param($node)
                    $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
                }, $true))
        $desktopFunction = $functions |
            Where-Object Name -eq 'Invoke-RSInteractiveDesktopOperation' |
            Select-Object -First 1
        Assert-RSEqual -Actual ($null -ne $desktopFunction) -Expected $true -Message 'The runner should centralize interactive desktop ownership.'

        $runnerText = Get-Content -LiteralPath $runnerPath -Raw
        Assert-RSEqual `
            -Actual ([regex]::Matches($runnerText, [regex]::Escape('Local\RedSalamander.InteractiveTestRun.v1')).Count) `
            -Expected 1 `
            -Message 'The session mutex must have one reviewed ownership site.'
        Assert-RSEqual -Actual ($desktopFunction.Extent.Text -match '\.WaitOne\(') -Expected $true -Message 'The entry boundary should acquire the mutex.'
        Assert-RSEqual -Actual ($desktopFunction.Extent.Text -match '\.ReleaseMutex\(\)') -Expected $true -Message 'The entry boundary should release the mutex in its finally block.'

        foreach ($functionName in @('Invoke-RSSelfTestProcess', 'Invoke-RSTestPlanEntry')) {
            $entryFunction = $functions | Where-Object Name -eq $functionName | Select-Object -First 1
            Assert-RSEqual `
                -Actual ($entryFunction.Extent.Text -match 'Invoke-RSInteractiveDesktopOperation') `
                -Expected $true `
                -Message "$functionName should route marked execution through the narrow desktop boundary."
        }

        $pesterCase = ($functions | Where-Object Name -eq 'Invoke-RSTestPlanEntry' | Select-Object -First 1).Extent.Text
        Assert-RSEqual `
            -Actual ($pesterCase -match "'Pester'[\s\S]*?Invoke-RSInteractiveDesktopOperation") `
            -Expected $false `
            -Message 'Pester execution must remain outside the interactive desktop boundary.'
    }

    It 'imports the contained-process API it calls directly' {
        $runnerText = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        Assert-RSEqual `
            -Actual ($runnerText -match "Tools\\Modules\\Build\\SanitizedEnvironment\.psm1") `
            -Expected $true `
            -Message 'Run-AllTests should import SanitizedEnvironment directly before calling Start/Close-RSContainedProcess.'
        Assert-RSEqual `
            -Actual ($runnerText -match 'Start-RSContainedProcess' -and $runnerText -match 'Close-RSContainedProcess') `
            -Expected $true `
            -Message 'The direct contained-process calls should stay source-visible.'
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
        Assert-RSEqual -Actual $fullBuild['MaxCpuCount'] -Expected 4 -Message 'Suite Full should bound complete-solution build fan-out.'

        $ciBuild = Get-RSBuildScriptArguments -Suite 'CI' -Configuration 'Debug' -Platform 'x64'

        Assert-RSEqual -Actual $ciBuild['Configuration'] -Expected 'Debug' -Message 'Suite CI should pass the requested configuration.'
        Assert-RSEqual -Actual $ciBuild['Platform'] -Expected 'x64' -Message 'Suite CI should pass the requested platform.'
        Assert-RSEqual -Actual $ciBuild.ContainsKey('ProjectName') -Expected $false -Message 'Suite CI should build the solution so standalone tests and CppUnitTest DLLs exist.'
        Assert-RSEqual -Actual $ciBuild.ContainsKey('MaxCpuCount') -Expected $false -Message 'Suite CI should retain the build host worker policy.'

        $allEnvironment = Get-RSBuildEnvironmentOverrides -Suite 'All'
        Assert-RSEqual -Actual $allEnvironment.ContainsKey('RSBuildEnableTests') -Expected $false -Message 'Suite All should not force monitor-specific test hooks.'

        $ciEnvironment = Get-RSBuildEnvironmentOverrides -Suite 'CI'
        Assert-RSEqual -Actual $ciEnvironment.ContainsKey('RSBuildEnableTests') -Expected $false -Message 'Suite CI should match the artifact-only GitHub gate and not force monitor-specific test hooks.'
        Assert-RSEqual -Actual $ciEnvironment.ContainsKey('RSBuildCompilerDebugInformation') -Expected $false -Message 'Suite CI should retain the normal compiler PDB policy.'

        $fullEnvironment = Get-RSBuildEnvironmentOverrides -Suite 'Full'
        Assert-RSEqual -Actual $fullEnvironment['RSBuildEnableTests'] -Expected 'true' -Message 'Suite Full should build Monitor with selftest hooks for executable monitor drills.'
        Assert-RSEqual -Actual $fullEnvironment['RSBuildCompilerDebugInformation'] -Expected 'Embedded' -Message 'Suite Full should avoid process-global compiler-PDB contention across worktrees.'
    }

    It 'uses the unified test sandbox root when locating runner artifacts' {
        $defaultContext = New-RSTestRunContext `
            -RepoRoot 'C:\repo' `
            -RunId '20260706T120000Z-42-abcdef' `
            -TestRootOverride '' `
            -SelfTestRootOverride '' `
            -LocalAppDataRoot 'C:\Users\eric\AppData\Local' `
            -AllowExternalInitialization
        Assert-RSEqual `
            -Actual $defaultContext.TestRoot `
            -Expected 'C:\RedSalamander.Perf' `
            -Message 'Default test sandbox should be the exact drive-owned root.'
        Assert-RSEqual `
            -Actual $defaultContext.ArtifactRoot `
            -Expected 'C:\RedSalamander.Perf\runs\20260706T120000Z-42-abcdef\artifacts\selftest\last_run' `
            -Message 'Native last_run artifacts should be bridged under the per-run sandbox.'

        foreach ($unauthorizedRoot in @(
                'C:\',
                'C:\Windows',
                'C:\Temp\RedSalamanderSandbox',
                'C:\repo\.build\TestSandbox',
                'C:\RedSalamander.Perf\child',
                'C:\RedSalamander.Perf:stream'
            )) {
            {
                Get-RSTestSandboxRoot `
                    -RepoRoot 'C:\repo' `
                    -TestRootOverride $unauthorizedRoot `
                    -SelfTestRootOverride ''
            } | Should Throw 'Unauthorized TestSandbox root'
        }

        $bridgedSandboxRoot = Get-RSTestSandboxRoot `
            -RepoRoot 'C:\repo' `
            -TestRootOverride 'C:\RedSalamander.Perf' `
            -SelfTestRootOverride 'C:\RedSalamander.Perf\runs\20260706T120000Z-42-abcdef\artifacts\selftest' `
            -LocalAppDataRoot 'C:\Users\eric\AppData\Local' `
            -AllowExternalInitialization
        Assert-RSEqual `
            -Actual $bridgedSandboxRoot `
            -Expected 'C:\RedSalamander.Perf' `
            -Message 'The runner-owned per-run legacy self-test bridge should be compatible with REDSALAMANDER_TEST_ROOT.'
    }

    It 'rejects conflicting legacy and unified test roots' {
        $threw = $false
        try {
            Get-RSTestSandboxRoot `
                -RepoRoot 'C:\repo' `
                -TestRootOverride 'C:\RedSalamander.Perf' `
                -SelfTestRootOverride 'C:\RedSalamander.Perf\other' `
                -LocalAppDataRoot 'C:\Users\eric\AppData\Local' `
                -AllowExternalInitialization | Out-Null
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
        $testRoot = $env:REDSALAMANDER_TEST_ROOT
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

            Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId $runId
        }
    }

    It 'keeps tooling Pester in the exact selected drive-owned sandbox' {
        $externalRoot = Get-RSToolingPesterSandboxRoot `
            -RepoRoot $repoRoot `
            -PrimaryTestRoot $env:REDSALAMANDER_TEST_ROOT
        Assert-RSEqual `
            -Actual $externalRoot `
            -Expected $env:REDSALAMANDER_TEST_ROOT `
            -Message 'Pester and native suites must share X:\RedSalamander.Perf.'
    }

    It 'removes only the exact auxiliary tooling Pester run directory' {
        $testRoot = $env:REDSALAMANDER_TEST_ROOT
        $runId = 'pester-exact-removal-proof'
        $runRoot = Join-Path (Join-Path $testRoot 'runs') $runId
        $neighbor = Join-Path (Join-Path $testRoot 'runs') 'manual-neighbor'
        try {
            New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
            New-Item -ItemType Directory -Path $neighbor -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $runRoot 'fixture.txt') -Value 'fixture'

            Remove-RSTestSandboxExactRunDirectory `
                -TestRoot $testRoot `
                -RunId $runId

            Assert-RSEqual `
                -Actual (Test-Path -LiteralPath $runRoot) `
                -Expected $false `
                -Message 'The exact auxiliary Pester run should be removed after its disk audit.'
            Assert-RSEqual `
                -Actual (Test-Path -LiteralPath $neighbor) `
                -Expected $true `
                -Message 'Exact-run cleanup must preserve neighboring TestSandbox state.'
        } finally {
            Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId $runId
            Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId 'manual-neighbor'
        }
    }

    It 'keeps the cleanup command inside one exact selected-root run' {
        $cleanupScript = Join-Path $repoRoot 'Tools\Clean-TestSandbox.ps1'
        $testRoot = $env:REDSALAMANDER_TEST_ROOT
        $suffix = [guid]::NewGuid().ToString('N')
        $runId = "cleanup-command-$PID-$suffix"
        $neighborRunId = "cleanup-neighbor-$PID-$suffix"
        $runRoot = Join-Path (Join-Path $testRoot 'runs') $runId
        $neighborRunRoot = Join-Path (Join-Path $testRoot 'runs') $neighborRunId
        try {
            New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
            New-Item -ItemType Directory -Path $neighborRunRoot -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $runRoot 'fixture.txt') -Value 'target'
            Set-Content -LiteralPath (Join-Path $neighborRunRoot 'fixture.txt') -Value 'neighbor'

            $dryRun = @(& $cleanupScript -RunId $runId -TestRoot $testRoot)
            Assert-RSEqual -Actual $dryRun.Count -Expected 1 -Message 'Dry run should report one exact run.'
            Assert-RSEqual -Actual $dryRun[0].Status -Expected 'Planned' -Message 'Dry run should not remove its target.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath $runRoot) -Expected $true -Message 'Dry run must preserve its target.'

            $cleanupResult = @(& $cleanupScript `
                    -Apply `
                    -Confirm:$false `
                    -RunId $runId `
                    -TestRoot $testRoot `
                    -WarningAction SilentlyContinue)

            Assert-RSEqual -Actual $cleanupResult.Count -Expected 1 -Message 'Apply should report one exact run.'
            Assert-RSEqual -Actual $cleanupResult[0].Status -Expected 'Removed' -Message 'Apply should remove its exact target.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath $runRoot) -Expected $false -Message 'Exact target should be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath $neighborRunRoot) -Expected $true -Message 'Neighboring runs must be preserved.'
        } finally {
            Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId $runId
            Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId $neighborRunId
        }
    }

    It 'audits TestSandbox disk state for stale runs and unexpected children' {
        $testRoot = $env:REDSALAMANDER_TEST_ROOT
        $fixtureSuffix = [guid]::NewGuid().ToString('N')
        $runId = 'audit-current-' + $fixtureSuffix
        $staleRunId = 'audit-stale-' + $fixtureSuffix
        $rogueChild = Join-Path $testRoot ('audit-rogue-' + $fixtureSuffix)
        $preExistingRunIds = @()
        $runsRoot = Join-Path $testRoot 'runs'
        if (Test-Path -LiteralPath $runsRoot) {
            $preExistingRunIds = @(Get-ChildItem -LiteralPath $runsRoot -Directory -Force | ForEach-Object Name)
        }
        try {
            New-Item -ItemType Directory -Path (Join-Path (Join-Path $testRoot 'runs') $runId) -Force | Out-Null

            $cleanAudit = Get-RSTestSandboxDiskAudit `
                -TestRoot $testRoot `
                -RunId $runId `
                -AllowedRunIds $preExistingRunIds

            Assert-RSEqual -Actual $cleanAudit.is_clean -Expected ($cleanAudit.issue_count -eq 0) -Message 'Overall cleanliness must agree with the issue count.'
            Assert-RSEqual -Actual @($cleanAudit.issues | Where-Object { $_.path -like "*$fixtureSuffix*" }).Count -Expected 0 -Message 'The current fixture run must be allowed even when unrelated root entries exist.'
            Assert-RSEqual -Actual $cleanAudit.schema -Expected 'red-salamander.test-sandbox-disk-audit.v1' -Message 'Disk audit schema should be versioned.'

            New-Item -ItemType Directory -Path (Join-Path $runsRoot $staleRunId) -Force | Out-Null
            New-Item -ItemType Directory -Path $rogueChild -Force | Out-Null

            $dirtyAudit = Get-RSTestSandboxDiskAudit `
                -TestRoot $testRoot `
                -RunId $runId `
                -AllowedRunIds $preExistingRunIds

            Assert-RSEqual -Actual $dirtyAudit.is_clean -Expected $false -Message 'Stale run and legacy roots should make the disk audit dirty.'
            foreach ($category in @(
                    'unexpected-test-run-dir',
                    'unexpected-test-root-child'
                )) {
                Assert-RSEqual `
                    -Actual (@($dirtyAudit.issues | Where-Object { $_.category -eq $category -and $_.path -like "*$fixtureSuffix*" }) | Measure-Object).Count `
                    -Expected 1 `
                    -Message "Disk audit should report $category."
            }
            Assert-RSEqual `
                -Actual (@($dirtyAudit.issues | Where-Object { $_.path -like "*$runId*" }) | Measure-Object).Count `
                -Expected 0 `
                -Message 'Current run directory should be allowed by the disk audit.'
        } finally {
            Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId $runId
            Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId $staleRunId
            Remove-Item -LiteralPath $rogueChild -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'sweeps stale TestSandbox run directories whose owner process is dead' {
        $testRoot = $env:REDSALAMANDER_TEST_ROOT
        $fixtureSuffix = [guid]::NewGuid().ToString('N')
        $runsRoot = Join-Path $testRoot 'runs'
        $currentRunId = '20260706T120000Z-4000-' + $fixtureSuffix
        $deadRunId = '20260706T110000Z-1234-' + $fixtureSuffix
        $liveRunId = '20260706T111000Z-5678-' + $fixtureSuffix
        $allowedRunId = '20260706T112000Z-9012-' + $fixtureSuffix
        $deadFallbackRunId = 'viewer-pe-2468-' + [Convert]::ToUInt32($fixtureSuffix.Substring(0,8),16).ToString()
        $liveFallbackRunId = 'redconfigure-1357-' + [Convert]::ToUInt32($fixtureSuffix.Substring(0,8),16).ToString()
        $reusedPidFallbackRunId = 'viewer-sqlite-5678-' + [Convert]::ToUInt32($fixtureSuffix.Substring(0,8),16).ToString()
        $manualRunId = 'manual-' + $fixtureSuffix
        $allowedRunIds = @($allowedRunId, $env:REDSALAMANDER_TEST_RUN_ID) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $neighborRunId = '20260706T100000Z-2468-' + [guid]::NewGuid().ToString('N')
        try {
            New-Item -ItemType Directory -Path (Join-Path $runsRoot $neighborRunId) -Force | Out-Null
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
                    -AllowedRunIds $allowedRunIds `
                    -LiveProcessIds @(1357, 4000, 5678, 9012) `
                    -LiveProcessStartTimesUtc $liveProcessStartTimesUtc)

            # Other dead-owner run directories may exist under the shared root (for example the finished
            # harness processes of the run hosting this test); assert on the fixture entries only.
            $fixtureRunIds = @($currentRunId, $deadRunId, $liveRunId, $allowedRunId, $deadFallbackRunId, $liveFallbackRunId, $reusedPidFallbackRunId, $manualRunId)
            $fixtureTargets = @($targets | Where-Object { $fixtureRunIds -contains $_.RunId })
            Assert-RSEqual -Actual @($fixtureTargets).Count -Expected 3 -Message 'Dead-owner and PID-reused runner/harness-fallback sibling run dirs should be selected.'
            Assert-RSEqual -Actual @($targets | Where-Object { $_.RunId -eq $deadRunId }).Count -Expected 1 -Message 'The dead runner-owned process run should be selected for cleanup.'
            Assert-RSEqual -Actual @($targets | Where-Object { $_.RunId -eq $deadFallbackRunId }).Count -Expected 1 -Message 'The dead direct-harness fallback run should be selected for cleanup.'
            Assert-RSEqual -Actual @($targets | Where-Object { $_.RunId -eq $reusedPidFallbackRunId }).Count -Expected 1 -Message 'An older fallback run must not be protected by an unrelated process that reused its PID.'
            Assert-RSEqual -Actual ($targets | Where-Object { $_.RunId -eq $deadRunId } | Select-Object -First 1).ProcessId -Expected 1234 -Message 'The runner-id parser should expose the owning PID.'
            Assert-RSEqual -Actual ($targets | Where-Object { $_.RunId -eq $deadFallbackRunId } | Select-Object -First 1).ProcessId -Expected 2468 -Message 'The fallback-id parser should expose the owning PID.'
            Assert-RSEqual -Actual @($fixtureTargets | Where-Object { $_.Category -ne 'stale-test-run-dir' }).Count -Expected 0 -Message 'Cleanup targets should be categorized for reporting.'

            $cleanupResults = @(Remove-RSTestSandboxStaleRunDirectories `
                    -TestRoot $testRoot `
                    -CandidateRunIds $fixtureRunIds `
                    -RunId $currentRunId `
                    -AllowedRunIds $allowedRunIds `
                    -LiveProcessIds @(1357, 4000, 5678, 9012) `
                    -LiveProcessStartTimesUtc $liveProcessStartTimesUtc)

            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $neighborRunId)) -Expected $true -Message 'An unrelated run must survive a fixture process snapshot.'
            $rejected = $false
            try { Remove-RSTestSandboxStaleRunDirectories -TestRoot $testRoot -LiveProcessIds @() | Out-Null } catch { $rejected = $true }
            Assert-RSEqual -Actual $rejected -Expected $true -Message 'Synthetic liveness without an exact cleanup scope must fail closed.'
            $fixtureCleanupResults = @($cleanupResults | Where-Object { $fixtureRunIds -contains $_.RunId })
            Assert-RSEqual -Actual @($fixtureCleanupResults).Count -Expected 3 -Message 'Dead runner, dead direct-harness, and PID-reused siblings should be removed.'
            Assert-RSEqual -Actual @($fixtureCleanupResults | Where-Object { $_.Status -ne 'Removed' }).Count -Expected 0 -Message 'Successful stale run cleanup should report Removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $deadRunId)) -Expected $false -Message 'Dead-PID stale run directory should be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $deadFallbackRunId)) -Expected $false -Message 'Dead-PID direct-harness fallback run directory should be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $reusedPidFallbackRunId)) -Expected $false -Message 'PID-reused direct-harness fallback run directory should be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $currentRunId)) -Expected $true -Message 'Current run directory must never be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $liveRunId)) -Expected $true -Message 'Live-PID sibling run directory must not be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $liveFallbackRunId)) -Expected $true -Message 'Live-PID direct-harness fallback run directory must not be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $allowedRunId)) -Expected $true -Message 'Explicitly allowed sibling run directory must not be removed.'
            Assert-RSEqual -Actual (Test-Path -LiteralPath (Join-Path $runsRoot $manualRunId)) -Expected $true -Message 'Unparseable manual run directory must not be removed by the dead-PID sweeper.'
        } finally {
            Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId $neighborRunId
            foreach ($cleanupRunId in @($currentRunId, $deadRunId, $liveRunId, $allowedRunId, $deadFallbackRunId, $liveFallbackRunId, $reusedPidFallbackRunId, $manualRunId)) {
                Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId $cleanupRunId
            }
        }
    }

    It 'audits runner-owned siblings with the same live-owner policy as stale cleanup' {
        $testRoot = $env:REDSALAMANDER_TEST_ROOT
        $runsRoot = Join-Path $testRoot 'runs'
        $currentRunId = '20260706T120000Z-4000-11111111111111111111111111111111'
        $sameOwnerRunId = '20260706T115900Z-4000-22222222222222222222222222222222'
        $otherLiveRunId = '20260706T115800Z-5678-33333333333333333333333333333333'
        $deadOwnerRunId = '20260706T115700Z-1234-44444444444444444444444444444444'
        $reusedPidRunId = '20260706T115600Z-2468-55555555555555555555555555555555'
        $manualRunId = 'audit-live-policy-manual-run'
        $fixtureRunIds = @($currentRunId, $sameOwnerRunId, $otherLiveRunId, $deadOwnerRunId, $reusedPidRunId, $manualRunId)
        $allowedRunIds = @()
        if (Test-Path -LiteralPath $runsRoot) {
            $allowedRunIds = @(Get-ChildItem -LiteralPath $runsRoot -Directory -Force | ForEach-Object Name)
        }

        try {
            foreach ($fixtureRunId in $fixtureRunIds) {
                New-Item -ItemType Directory -Path (Join-Path $runsRoot $fixtureRunId) -Force | Out-Null
            }
            (Get-Item -LiteralPath (Join-Path $runsRoot $reusedPidRunId)).CreationTimeUtc =
                (Get-Date).ToUniversalTime().AddDays(-2)
            $liveProcessStartTimesUtc = @{
                4000 = (Get-Date).ToUniversalTime().AddHours(-2)
                5678 = (Get-Date).ToUniversalTime().AddHours(-1)
                2468 = (Get-Date).ToUniversalTime().AddMinutes(-30)
            }

            $audit = Get-RSTestSandboxDiskAudit `
                -TestRoot $testRoot `
                -RunId $currentRunId `
                -AllowedRunIds $allowedRunIds `
                -LiveProcessIds @(2468, 4000, 5678) `
                -LiveProcessStartTimesUtc $liveProcessStartTimesUtc

            Assert-RSEqual -Actual $audit.is_clean -Expected $false -Message 'Dead-owner, PID-reused, and manual directories should remain visible audit issues.'
            Assert-RSEqual -Actual @($audit.issues | Where-Object { $_.category -eq 'stale-test-run-dir' }).Count -Expected 2 -Message 'Dead-owner and PID-reused runner directories should be stale audit issues.'
            Assert-RSEqual -Actual @($audit.issues | Where-Object { $_.category -eq 'unexpected-test-run-dir' }).Count -Expected 1 -Message 'The unparseable manual directory should remain an unexpected audit issue.'
            Assert-RSEqual -Actual @($audit.issues | Where-Object { $_.path -like "*$sameOwnerRunId*" }).Count -Expected 0 -Message 'A prior sequential run owned by the same live parent process should be protected.'
            Assert-RSEqual -Actual @($audit.issues | Where-Object { $_.path -like "*$otherLiveRunId*" }).Count -Expected 0 -Message 'A concurrently live sibling run should be protected.'
        } finally {
            foreach ($fixtureRunId in $fixtureRunIds) {
                Remove-RSTestSandboxExactRunDirectory -TestRoot $testRoot -RunId $fixtureRunId
            }
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
            -Id 'selftest.commands' `
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

    It 'treats an absent default quarantine ledger as the canonical empty state' {
        $defaultQuarantinePath = Join-Path $repoRoot 'Tools\test-quarantine.jsonl'

        Assert-RSEqual `
            -Actual (Test-Path -LiteralPath $defaultQuarantinePath) `
            -Expected $false `
            -Message 'The repository must not retain an empty default quarantine placeholder.'

        $status = Read-RSTestQuarantineFile `
            -Path $defaultQuarantinePath `
            -Now ([datetime]'2026-07-10T00:00:00Z')

        Assert-RSEqual -Actual $status.IsValid -Expected $true -Message 'A missing default ledger should be valid.'
        Assert-RSEqual -Actual $status.HasBlockingEntries -Expected $false -Message 'A missing default ledger should not block the run.'
        Assert-RSEqual -Actual $status.TotalCount -Expected 0 -Message 'A missing default ledger should resolve to zero entries.'
        Assert-RSEqual -Actual @($status.Errors).Count -Expected 0 -Message 'A missing default ledger should not synthesize errors.'
    }

    It 'keeps malformed and active entries in an explicit quarantine ledger blocking' {
        $quarantinePath = Join-Path $TestDrive 'explicit-quarantine.jsonl'
        Set-Content -LiteralPath $quarantinePath -Value '{not valid json' -Encoding utf8

        $malformedStatus = Read-RSTestQuarantineFile `
            -Path $quarantinePath `
            -Now ([datetime]'2026-07-10T00:00:00Z')

        Assert-RSEqual -Actual $malformedStatus.IsValid -Expected $false -Message 'Malformed explicit ledger content must fail validation.'
        Assert-RSEqual -Actual $malformedStatus.HasBlockingEntries -Expected $true -Message 'Malformed explicit ledger content must block the run.'
        Assert-RSEqual -Actual (@($malformedStatus.Errors) -join "`n" -match 'not valid JSON') -Expected $true -Message 'Malformed ledger errors should identify invalid JSON.'

        [pscustomobject]@{
            harness = 'DxUiTests.Menu'
            name = 'DxUiTests.Menu'
            owner = 'ui-runtime-owner'
            opened = '2026-07-06'
            expires = '2026-07-20'
            issue = 'Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md#dxui-menu'
            root_cause_hypothesis = 'Foreground denial drops popup capture on hosted runners.'
            fix_or_replace_plan = 'Gate foreground-only assertions behind deterministic interactive-desktop probe.'
        } | ConvertTo-Json -Compress | Set-Content -LiteralPath $quarantinePath -Encoding utf8

        $activeStatus = Read-RSTestQuarantineFile `
            -Path $quarantinePath `
            -Now ([datetime]'2026-07-10T00:00:00Z')

        Assert-RSEqual -Actual $activeStatus.IsValid -Expected $true -Message 'A reviewed explicit ledger entry should validate before expiry.'
        Assert-RSEqual -Actual $activeStatus.HasBlockingEntries -Expected $true -Message 'A reviewed active explicit ledger entry must block the run.'
        Assert-RSEqual -Actual @($activeStatus.ActiveEntries).Count -Expected 1 -Message 'The explicit ledger entry should remain available to the repair lane.'
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
                    -Id 'selftest.commands' `
                    -Name 'Commands' `
                    -Kind 'SelfTest' `
                    -Path 'C:\repo\.build\x64\Debug\RedSalamander.exe' `
                    -Arguments @('--commands-selftest', '--selftest-timeout-multiplier=2') `
                    -WorkingDirectory 'C:\repo' `
                    -JsonName 'commands' `
                    -RequiresInteractiveDesktop),
            (New-RSTestRunPlanEntry `
                    -Id 'dxui.menu' `
                    -Name 'DxUiTests.Menu' `
                    -Kind 'Executable' `
                    -Path 'C:\repo\.build\x64\Debug\DxUiTests.exe' `
                    -Arguments @('--suite=Menu') `
                    -WorkingDirectory 'C:\repo' `
                    -RequiresInteractiveDesktop)
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
        Assert-RSEqual -Actual $repairPlan.Attempts[0].requires_interactive_desktop -Expected $true -Message 'Self-test repair should preserve its desktop requirement.'
        Assert-RSEqual -Actual $repairPlan.Attempts[0].id -Expected 'selftest.commands' -Message 'Repair attempt should preserve stable entry ID.'
        Assert-RSEqual -Actual $repairPlan.Attempts[0].cache_policy -Expected 'ExactArtifact' -Message 'Repair attempt should preserve cache policy.'
        Assert-RSSequenceEqual -Actual @($repairPlan.Attempts[0].artifact_read_ids) -Expected @('artifact.redsalamander-runtime') `
            -Message 'Repair attempt should preserve artifact reads.'
        Assert-RSEqual -Actual $repairPlan.Attempts[0].coverage_contract.suite -Expected 'selftest.commands' `
            -Message 'Repair attempt should preserve coverage policy.'
        Assert-RSEqual -Actual ($repairPlan.Attempts[0].resource_classes -contains 'interactive-desktop') -Expected $true `
            -Message 'Repair attempt should preserve resource classes.'
        Assert-RSEqual -Actual $repairPlan.Attempts[0].owner -Expected 'commands-owner' -Message 'Repair attempt should preserve owner metadata.'
        Assert-RSEqual -Actual $repairPlan.Attempts[1].harness -Expected 'DxUiTests.Menu' -Message 'Standalone repair attempt should preserve harness.'
        Assert-RSEqual -Actual $repairPlan.Attempts[1].mode -Expected 'entry' -Message 'Standalone repair attempt should rerun the whole entry.'
        Assert-RSEqual -Actual ($repairPlan.Attempts[1].arguments -contains '--suite=Menu') -Expected $true -Message 'Standalone repair attempt should preserve entry arguments.'
        Assert-RSEqual -Actual $repairPlan.Attempts[1].requires_interactive_desktop -Expected $true -Message 'Standalone repair should preserve its desktop requirement.'
    }

    It 'rejects quarantine entries that do not match a runnable harness adapter' {
        $plan = @(
            (New-RSTestRunPlanEntry `
                    -Id 'selftest.commands' `
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
                    -Id 'selftest.commands' `
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
            -ArtifactRoot 'C:\RedSalamander.Perf\runs\20260706T120000Z-42-abcdef\artifacts\selftest\last_run' `
            -RunStartedUtc ([datetime]'2026-05-05T17:00:00Z') `
            -RunEndedUtc ([datetime]'2026-05-05T17:00:30Z') `
            -DurationMs 30000 `
            -ExitCode 1 `
            -TimeoutMultiplier 2.0 `
            -CaseFilter 'Phase7_' `
            -TestRoot 'C:\RedSalamander.Perf' `
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
                test_root = 'C:\RedSalamander.Perf'
                run_id = '20260706T120000Z-42-abcdef'
                allowed_run_ids = @('20260706T120000Z-42-abcdef')
                issues = @([pscustomobject]@{
                        category = 'unexpected-test-run-dir'
                        path = 'C:\RedSalamander.Perf\runs\stale-run'
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
                    OutputLogPath = 'C:\RedSalamander.Perf\runs\20260706T120000Z-42-abcdef\artifacts\logs\CompareDirectories.output.log'
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
        Assert-RSEqual -Actual $summary.test_root -Expected 'C:\RedSalamander.Perf' -Message 'Aggregate artifact should record the canonical test root.'
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
        Assert-RSEqual -Actual $summary.suites[0].output_log_path -Expected 'C:\RedSalamander.Perf\runs\20260706T120000Z-42-abcdef\artifacts\logs\CompareDirectories.output.log' -Message 'Aggregate artifact should preserve suite output logs.'
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
            -ArtifactRoot 'C:\RedSalamander.Perf\runs\run-1\artifacts\selftest\last_run' `
            -RunStartedUtc ([datetime]'2026-07-06T10:00:00Z') `
            -RunEndedUtc ([datetime]'2026-07-06T10:05:00Z') `
            -DurationMs 300000 `
            -ExitCode 1 `
            -TimeoutMultiplier 2.0 `
            -TestRoot 'C:\RedSalamander.Perf' `
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
            -ArtifactRoot 'C:\RedSalamander.Perf\runs\20260706T120000Z-42-abcdef\artifacts\selftest\last_run' `
            -RunStartedUtc ([datetime]'2026-07-06T12:00:00Z') `
            -RunEndedUtc ([datetime]'2026-07-06T12:02:00Z') `
            -DurationMs 120000 `
            -ExitCode 1 `
            -TimeoutMultiplier 2.0 `
            -TestRoot 'C:\RedSalamander.Perf' `
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

Describe 'Commands family runtime membership' -Tag RequiresBuildToolchain {
    BeforeAll {
        Import-Module $helperModule -Force -ErrorAction Stop
    }

    It 'partitions the broad native inventory into an exact non-overlapping family union' {
        $platform = $env:RS_VALIDATION_PLATFORM
        $configuration = $env:RS_VALIDATION_CONFIGURATION
        if ($platform -notin @('x64', 'ARM64') -or $configuration -notin @('Debug', 'Release', 'ASan Debug')) {
            throw 'Commands inventory requires an explicit supported validation profile.'
        }
        $exe = Join-Path $repoRoot ('.build\{0}\{1}\RedSalamander.exe' -f $platform, $configuration)
        Test-Path -LiteralPath $exe -PathType Leaf | Should Be $true
        function Get-RSNativeCommandsInventory {
            param([string[]]$Argument)
            $start = [Diagnostics.ProcessStartInfo]::new()
            $start.FileName = $exe
            $start.WorkingDirectory = $repoRoot
            $start.UseShellExecute = $false
            $start.RedirectStandardOutput = $true
            $start.RedirectStandardError = $true
            $start.CreateNoWindow = $true
            foreach ($value in $Argument) { [void]$start.ArgumentList.Add($value) }
            $process = [Diagnostics.Process]::new()
            $process.StartInfo = $start
            try {
                $null = ($process.Start() | Should Be $true)
                $stdoutTask = $process.StandardOutput.ReadToEndAsync()
                $stderrTask = $process.StandardError.ReadToEndAsync()
                $process.WaitForExit()
                $stderr = $stderrTask.GetAwaiter().GetResult()
                $null = ($process.ExitCode | Should Be 0)
                $null = ([string]::IsNullOrWhiteSpace($stderr) | Should Be $true)
                return ($stdoutTask.GetAwaiter().GetResult() | ConvertFrom-Json)
            } finally {
                $process.Dispose()
            }
        }
        $broad = Get-RSNativeCommandsInventory -Argument @('--commands-selftest', '--selftest-list-cases')
        $broadNames = @($broad.suites[0].cases | ForEach-Object { [string]$_.name })
        $familyNames = @()
        foreach ($family in @(Get-RSCommandsSelfTestFamilies)) {
            $inventory = Get-RSNativeCommandsInventory -Argument @('--commands-selftest', "--selftest-family=$($family.name)", '--selftest-list-cases')
            @($inventory.suites).Count | Should Be 1
            @($inventory.suites[0].cases).Count | Should BeGreaterThan 0
            $familyNames += @($inventory.suites[0].cases | ForEach-Object { [string]$_.name })
        }
        @($familyNames | Sort-Object -Unique).Count | Should Be $familyNames.Count
        Assert-RSSequenceEqual -Actual @($familyNames | Sort-Object) -Expected @($broadNames | Sort-Object) `
            -Message 'Commands family union must exactly equal broad native membership.'
    }

}
