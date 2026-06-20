Set-StrictMode -Version Latest

function New-RSTestRunPlanEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateSet('SelfTest', 'Executable', 'CppUnitTest', 'Pester', 'PowerShellScript')]
        [string]$Kind,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string[]]$Arguments = @(),

        [string]$WorkingDirectory = '',

        [string]$JsonName = ''
    )

    [pscustomobject]@{
        Name = $Name
        Kind = $Kind
        Path = $Path
        Arguments = @($Arguments)
        WorkingDirectory = $WorkingDirectory
        JsonName = $JsonName
    }
}

function Get-RSBuildOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    return (Join-Path $RepoRoot (".build\{0}\{1}\{2}" -f $Platform, $Configuration, $FileName))
}

function Get-RSSelfTestArtifactRoot {
    param(
        [string]$SelfTestRootOverride = $env:REDSALAMANDER_SELFTEST_ROOT,

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA
    )

    if (-not [string]::IsNullOrWhiteSpace($SelfTestRootOverride)) {
        return (Join-Path $SelfTestRootOverride 'last_run')
    }

    return (Join-Path (Join-Path $LocalAppDataRoot 'RedSalamander\SelfTest') 'last_run')
}

function Get-RSSelfTestArguments {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Flag,

        [double]$TimeoutMultiplier = 1.0,

        [switch]$FailFast,

        [string]$CaseFilter = ''
    )

    $arguments = @($Flag)
    if ($FailFast) {
        $arguments += '--selftest-fail-fast'
    }
    if ($TimeoutMultiplier -ne 1.0) {
        $arguments += ("--selftest-timeout-multiplier={0:g}" -f $TimeoutMultiplier)
    }
    if (-not [string]::IsNullOrWhiteSpace($CaseFilter)) {
        $arguments += "--selftest-case=$CaseFilter"
    }

    return $arguments
}

function Get-RSSelfTestListArguments {
    param(
        [string[]]$SelfTestArguments = @()
    )

    $arguments = @('--selftest-list-cases')
    foreach ($argument in @($SelfTestArguments)) {
        if ($argument -in @('--compare-selftest', '--commands-selftest', '--fileops-selftest') -or
            $argument -like '--selftest-case=*') {
            $arguments += $argument
        }
    }

    return $arguments
}

function Get-RSBuildScriptArguments {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'Full')]
        [string]$Suite,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform
    )

    $arguments = @{
        Configuration = $Configuration
        Platform = $Platform
    }

    if ($Suite -ne 'Full') {
        $arguments.ProjectName = 'RedSalamander'
    }

    return $arguments
}

function Get-RSBuildEnvironmentOverrides {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'Full')]
        [string]$Suite
    )

    if ($Suite -eq 'Full') {
        return @{
            RSBuildEnableTests = 'true'
        }
    }

    return @{}
}

function Test-RSSelfTestResultCoverage {
    param(
        [string[]]$ExpectedCaseNames = @(),

        [object[]]$ActualCases = @(),

        [string[]]$AllowedExtraCaseNames = @(),

        [switch]$RequireExpectedCases
    )

    $actualNames = @($ActualCases | ForEach-Object {
            $property = $_.PSObject.Properties['name']
            if ($null -ne $property) {
                [string]$property.Value
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    $expectedGroups = @($ExpectedCaseNames | Group-Object | Where-Object { $_.Count -gt 1 })
    $actualGroups = @($actualNames | Group-Object | Where-Object { $_.Count -gt 1 })

    $expectedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($name in @($ExpectedCaseNames)) {
        [void]$expectedSet.Add($name)
    }

    $actualSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($name in @($actualNames)) {
        [void]$actualSet.Add($name)
    }

    $allowedExtras = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($name in @($AllowedExtraCaseNames)) {
        [void]$allowedExtras.Add($name)
    }

    $missing = @($ExpectedCaseNames | Where-Object { -not $actualSet.Contains($_) } | Sort-Object -Unique)
    $extra = @($actualNames | Where-Object { (-not $expectedSet.Contains($_)) -and (-not $allowedExtras.Contains($_)) } | Sort-Object -Unique)
    $noExpectedCases = ($RequireExpectedCases -and @($ExpectedCaseNames).Count -eq 0)

    [pscustomobject]@{
        IsValid = ((-not $noExpectedCases) -and @($expectedGroups).Count -eq 0 -and @($actualGroups).Count -eq 0 -and @($missing).Count -eq 0 -and @($extra).Count -eq 0)
        NoExpectedCases = [bool]$noExpectedCases
        DuplicateExpected = @($expectedGroups | Select-Object -ExpandProperty Name)
        DuplicateActual = @($actualGroups | Select-Object -ExpandProperty Name)
        Missing = @($missing)
        Extra = @($extra)
    }
}

function Get-RSObjectValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Object,

        [Parameter(Mandatory = $true)]
        [string[]]$Names,

        [AllowNull()]
        [object]$DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }

    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($name in $Names) {
            if ($Object.Contains($name)) {
                return $Object[$name]
            }
        }
        return $DefaultValue
    }

    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties[$name]
        if ($null -ne $property) {
            return $property.Value
        }
    }

    return $DefaultValue
}

function Convert-RSTestCaseForRunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Case
    )

    [pscustomobject]@{
        name = [string](Get-RSObjectValue -Object $Case -Names @('name', 'Name', 'CaseName') -DefaultValue '(unnamed case)')
        status = [string](Get-RSObjectValue -Object $Case -Names @('status', 'Status') -DefaultValue '')
        reason = [string](Get-RSObjectValue -Object $Case -Names @('reason', 'Reason') -DefaultValue '')
        duration_ms = [uint64](Get-RSObjectValue -Object $Case -Names @('durationMs', 'duration_ms', 'Duration') -DefaultValue 0)
    }
}

function New-RSTestRunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'Full')]
        [string]$Suite,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$ExePath,

        [Parameter(Mandatory = $true)]
        [string]$ArtifactRoot,

        [Parameter(Mandatory = $true)]
        [datetime]$RunStartedUtc,

        [Parameter(Mandatory = $true)]
        [datetime]$RunEndedUtc,

        [uint64]$DurationMs = 0,

        [int]$ExitCode = 0,

        [double]$TimeoutMultiplier = 1.0,

        [bool]$FailFast = $false,

        [string]$CaseFilter = '',

        [object[]]$Results = @()
    )

    $totalPassed = 0
    $totalFailed = 0
    $totalSkipped = 0
    $summarySuites = @()

    foreach ($result in @($Results)) {
        $passed = [int](Get-RSObjectValue -Object $result -Names @('Passed', 'passed') -DefaultValue 0)
        $failed = [int](Get-RSObjectValue -Object $result -Names @('Failed', 'failed') -DefaultValue 0)
        $skipped = [int](Get-RSObjectValue -Object $result -Names @('Skipped', 'skipped') -DefaultValue 0)

        $totalPassed += $passed
        $totalFailed += $failed
        $totalSkipped += $skipped

        $cases = @(Get-RSObjectValue -Object $result -Names @('Cases', 'cases') -DefaultValue @() | ForEach-Object {
                Convert-RSTestCaseForRunSummary -Case $_
            })
        $failures = @(Get-RSObjectValue -Object $result -Names @('Failures', 'failures') -DefaultValue @() | ForEach-Object {
                Convert-RSTestCaseForRunSummary -Case $_
            })

        $summarySuites += [pscustomobject]@{
            suite = [string](Get-RSObjectValue -Object $result -Names @('Name', 'suite', 'name') -DefaultValue '(unnamed suite)')
            kind = [string](Get-RSObjectValue -Object $result -Names @('Kind', 'kind') -DefaultValue '')
            exit_code = [int](Get-RSObjectValue -Object $result -Names @('ExitCode', 'exit_code') -DefaultValue 0)
            duration_ms = [uint64](Get-RSObjectValue -Object $result -Names @('WallMs', 'DurationMs', 'duration_ms') -DefaultValue 0)
            passed = $passed
            failed = $failed
            skipped = $skipped
            output_log_path = [string](Get-RSObjectValue -Object $result -Names @('OutputLogPath', 'output_log_path') -DefaultValue '')
            cases = $cases
            failures = $failures
        }
    }

    [pscustomobject]@{
        schema = 'red-salamander.run-all-tests.v1'
        run_started_utc = $RunStartedUtc.ToString('o')
        run_ended_utc = $RunEndedUtc.ToString('o')
        duration_ms = $DurationMs
        suite = $Suite
        platform = $Platform
        configuration = $Configuration
        executable = $ExePath
        artifact_root = $ArtifactRoot
        fail_fast = $FailFast
        timeout_scale = $TimeoutMultiplier
        case_filter = $CaseFilter
        exit_code = $ExitCode
        passed = $totalPassed
        failed = $totalFailed
        skipped = $totalSkipped
        total = ($totalPassed + $totalFailed + $totalSkipped)
        suites = $summarySuites
    }
}

function Get-RSTestRunPlan {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'Full')]
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

        [string]$CaseFilter = ''
    )

    $buildOutputDir = Join-Path $RepoRoot (".build\{0}\{1}" -f $Platform, $Configuration)
    $plan = @()

    $selfTests = @()
    switch ($Suite) {
        'All' {
            $selfTests += @{ Name = 'CompareDirectories'; Flag = '--compare-selftest'; JsonName = 'compare' }
            $selfTests += @{ Name = 'Commands'; Flag = '--commands-selftest'; JsonName = 'commands' }
            $selfTests += @{ Name = 'FileOperations'; Flag = '--fileops-selftest'; JsonName = 'fileops' }
        }
        'Full' {
            $selfTests += @{ Name = 'CompareDirectories'; Flag = '--compare-selftest'; JsonName = 'compare' }
            $selfTests += @{ Name = 'Commands'; Flag = '--commands-selftest'; JsonName = 'commands' }
            $selfTests += @{ Name = 'FileOperations'; Flag = '--fileops-selftest'; JsonName = 'fileops' }
        }
        'Compare' {
            $selfTests += @{ Name = 'CompareDirectories'; Flag = '--compare-selftest'; JsonName = 'compare' }
        }
        'Commands' {
            $selfTests += @{ Name = 'Commands'; Flag = '--commands-selftest'; JsonName = 'commands' }
        }
        'FileOps' {
            $selfTests += @{ Name = 'FileOperations'; Flag = '--fileops-selftest'; JsonName = 'fileops' }
        }
    }

    foreach ($selfTest in $selfTests) {
        $plan += New-RSTestRunPlanEntry `
            -Name $selfTest.Name `
            -Kind 'SelfTest' `
            -Path $RedSalamanderExePath `
            -Arguments (Get-RSSelfTestArguments -Flag $selfTest.Flag -TimeoutMultiplier $TimeoutMultiplier -FailFast:$FailFast -CaseFilter $CaseFilter) `
            -WorkingDirectory $RepoRoot `
            -JsonName $selfTest.JsonName
    }

    # CI-parity note (for reference — do not edit the logic above to chase CI):
    # -Suite Full is a SUPERSET of what .github/workflows/ci.yml's selftest job runs.
    #   Full adds:  FileSystemCurlTests, RedConfigureTests, RedSalamanderMonitorEtwLatency (CI runs none of these; plan 011 owns adding them to CI).
    #   Known divergences vs CI:
    #     - ViewerPETests: CI also invokes two explicit named cases on top of the bare exe run; Full only runs the bare exe.
    #     - Pester: Full runs all Tools\Tests in pwsh with no tag filter; CI excludes -Tag RequiresBuildToolchain under Windows PowerShell 5.1.
    if ($Suite -eq 'Full') {
        foreach ($exeName in @('DxUiTests', 'FileSystemCurlTests', 'ViewerPETests', 'ViewerSqliteTests', 'MonitorTest', 'LocalizationTests', 'RedConfigureTests', 'PluginContractTests')) {
            $plan += New-RSTestRunPlanEntry `
                -Name $exeName `
                -Kind 'Executable' `
                -Path (Join-Path $buildOutputDir "$exeName.exe") `
                -WorkingDirectory $buildOutputDir
        }

        $plan += New-RSTestRunPlanEntry `
            -Name 'RedSalamanderMonitorEtwLatency' `
            -Kind 'Executable' `
            -Path (Join-Path $buildOutputDir 'RedSalamanderMonitor.exe') `
            -Arguments @(
                '--chrome-selftest',
                '--perf',
                '--monitor-etw-burst-mode=latency',
                '--monitor-etw-burst-count=60',
                '--monitor-etw-burst-size=260'
            ) `
            -WorkingDirectory $RepoRoot

        $plan += New-RSTestRunPlanEntry `
            -Name 'PerformanceTests2' `
            -Kind 'CppUnitTest' `
            -Path (Join-Path $buildOutputDir 'PerformanceTests2.dll') `
            -WorkingDirectory $buildOutputDir

        $plan += New-RSTestRunPlanEntry `
            -Name 'ToolsPesterTests' `
            -Kind 'Pester' `
            -Path (Join-Path $RepoRoot 'Tools\Tests') `
            -WorkingDirectory $RepoRoot

        $plan += New-RSTestRunPlanEntry `
            -Name 'VcpkgMergeSynthetic' `
            -Kind 'PowerShellScript' `
            -Path (Join-Path $RepoRoot 'Tests\vcpkg-merge-synthetic-test.ps1') `
            -WorkingDirectory $RepoRoot
    }

    return $plan
}
