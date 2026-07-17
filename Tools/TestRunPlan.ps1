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

function ConvertTo-RSFullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath($Path)
}

function New-RSTestRunId {
    $timestamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $guid = [guid]::NewGuid().ToString('N')
    return "$timestamp-$PID-$guid"
}

function ConvertFrom-RSTestRunId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunId
    )

    $match = [regex]::Match($RunId, '^(?<timestamp>\d{8}T\d{6}Z)-(?<pid>\d+)-(?<guid>[0-9a-fA-F]{32})$')
    $format = 'Runner'
    if (-not $match.Success) {
        # Standalone native harnesses use TestSupport's crash-tolerant fallback
        # <harness>-<pid>-<tick> form when no runner-owned id is supplied.
        $match = [regex]::Match($RunId, '^(?<prefix>[A-Za-z0-9][A-Za-z0-9._-]*?)-(?<pid>\d+)-(?<tick>\d+)$')
        $format = 'HarnessFallback'
    }
    if (-not $match.Success) {
        return $null
    }

    $processId = 0L
    if (-not [int64]::TryParse($match.Groups['pid'].Value, [ref]$processId)) {
        return $null
    }

    [pscustomobject]@{
        RunId = $RunId
        Format = $format
        Timestamp = $match.Groups['timestamp'].Value
        ProcessId = $processId
        Guid = $match.Groups['guid'].Value
        Prefix = $match.Groups['prefix'].Value
        TickCount = $match.Groups['tick'].Value
    }
}

function Get-RSLiveProcessSnapshots {
    $processes = @([System.Diagnostics.Process]::GetProcesses())
    $snapshots = @()
    try {
        $snapshots = @($processes | ForEach-Object {
                $startTimeUtc = $null
                try {
                    $startTimeUtc = $_.StartTime.ToUniversalTime()
                } catch {
                    # Access to protected process metadata can be denied. A null
                    # start time keeps the conservative live-PID behavior.
                }

                [pscustomobject]@{
                    ProcessId = [int64]$_.Id
                    StartTimeUtc = $startTimeUtc
                }
            })
    } finally {
        foreach ($process in $processes) {
            $process.Dispose()
        }
    }

    $missingStartIds = @($snapshots | Where-Object { $null -eq $_.StartTimeUtc } | ForEach-Object { [int64]$_.ProcessId })
    if ($missingStartIds.Count -gt 0) {
        try {
            foreach ($process in @(Get-CimInstance Win32_Process -Property ProcessId, CreationDate -ErrorAction Stop)) {
                if ($null -eq $process.CreationDate -or [int64]$process.ProcessId -notin $missingStartIds) {
                    continue
                }

                $snapshot = $snapshots | Where-Object { $_.ProcessId -eq [int64]$process.ProcessId } | Select-Object -First 1
                if ($null -ne $snapshot) {
                    $snapshot.StartTimeUtc = ([datetime]$process.CreationDate).ToUniversalTime()
                }
            }
        } catch {
            # Keep null start times conservative when process metadata is unavailable.
        }
    }

    return $snapshots
}

function Get-RSTestSandboxRoot {
    param(
        [string]$RepoRoot = '',

        [string]$TestRootOverride = $env:REDSALAMANDER_TEST_ROOT,

        [string]$SelfTestRootOverride = $env:REDSALAMANDER_SELFTEST_ROOT,

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA
    )

    if (-not [string]::IsNullOrWhiteSpace($TestRootOverride)) {
        $testRoot = ConvertTo-RSFullPath -Path $TestRootOverride
        if (-not [string]::IsNullOrWhiteSpace($SelfTestRootOverride)) {
            $legacyRoot = ConvertTo-RSFullPath -Path $SelfTestRootOverride
            $testRootNormalized = $testRoot.TrimEnd('\')
            $legacyRootNormalized = $legacyRoot.TrimEnd('\')
            $runsRootNormalized = (Join-Path $testRootNormalized 'runs').TrimEnd('\')
            $isSameRoot = [string]::Equals($testRootNormalized, $legacyRootNormalized, [System.StringComparison]::OrdinalIgnoreCase)
            $isRunnerBridge = $legacyRootNormalized.StartsWith("$runsRootNormalized\", [System.StringComparison]::OrdinalIgnoreCase)
            if (-not $isSameRoot -and -not $isRunnerBridge) {
                throw "REDSALAMANDER_TEST_ROOT ('$testRoot') and REDSALAMANDER_SELFTEST_ROOT ('$legacyRoot') conflict. Set only REDSALAMANDER_TEST_ROOT, clear REDSALAMANDER_SELFTEST_ROOT, or point both at the same sandbox base."
            }
        }

        return $testRoot
    }

    if (-not [string]::IsNullOrWhiteSpace($SelfTestRootOverride)) {
        return (ConvertTo-RSFullPath -Path $SelfTestRootOverride)
    }

    if (-not [string]::IsNullOrWhiteSpace($RepoRoot)) {
        return (ConvertTo-RSFullPath -Path (Join-Path $RepoRoot '.build\TestSandbox'))
    }

    return (ConvertTo-RSFullPath -Path (Join-Path $LocalAppDataRoot 'RedSalamander\TestSandbox'))
}

function New-RSTestRunContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [string]$RunId = '',

        [string]$TestRootOverride = $env:REDSALAMANDER_TEST_ROOT,

        [string]$SelfTestRootOverride = $env:REDSALAMANDER_SELFTEST_ROOT,

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA
    )

    $testRoot = Get-RSTestSandboxRoot `
        -RepoRoot $RepoRoot `
        -TestRootOverride $TestRootOverride `
        -SelfTestRootOverride $SelfTestRootOverride `
        -LocalAppDataRoot $LocalAppDataRoot
    if ([string]::IsNullOrWhiteSpace($RunId)) {
        $RunId = New-RSTestRunId
    }

    $runRoot = Join-Path (Join-Path $testRoot 'runs') $RunId
    $legacySelfTestRoot = Join-Path (Join-Path $runRoot 'artifacts') 'selftest'

    [pscustomobject]@{
        TestRoot = $testRoot
        RunId = $RunId
        RunRoot = $runRoot
        LegacySelfTestRoot = $legacySelfTestRoot
        ArtifactRoot = Join-Path $legacySelfTestRoot 'last_run'
    }
}

function Assert-RSTestSandboxPathSegment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "$Description must not be empty."
    }

    if ($Value -eq '.' -or $Value -eq '..' -or $Value -match '[\\/:\*\?"<>\|]') {
        throw "$Description '$Value' must be a single filesystem path segment."
    }

    return $Value
}

function New-RSTestSandboxScratchDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Harness,

        [Parameter(Mandatory = $true)]
        [string]$Case,

        [string]$RunId = $env:REDSALAMANDER_TEST_RUN_ID
    )

    $safeHarness = Assert-RSTestSandboxPathSegment -Value $Harness -Description 'TestSandbox harness segment'
    $safeCase = Assert-RSTestSandboxPathSegment -Value $Case -Description 'TestSandbox case segment'
    $context = New-RSTestRunContext -RepoRoot $RepoRoot -RunId $RunId
    $runRoot = $context.RunRoot
    $scratchRoot = Join-Path $runRoot 'scratch'
    $harnessRoot = Join-Path $scratchRoot $safeHarness
    $caseRoot = Join-Path $harnessRoot $safeCase

    New-Item -ItemType Directory -Path $caseRoot -Force | Out-Null
    return (ConvertTo-RSFullPath -Path $caseRoot)
}

function New-RSTestSandboxCleanupPlanEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Literal', 'Wildcard')]
        [string]$PathType,

        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    [pscustomobject]@{
        Path = $Path
        PathType = $PathType
        Category = $Category
        Reason = $Reason
    }
}

function Get-RSTestSandboxLegacyCleanupPlan {
    param(
        [string]$LocalAppDataRoot = $env:LOCALAPPDATA,

        [string]$TempRoot = [System.IO.Path]::GetTempPath(),

        [string[]]$DriveRoots = @()
    )

    $plan = @()
    if (-not [string]::IsNullOrWhiteSpace($LocalAppDataRoot)) {
        $plan += New-RSTestSandboxCleanupPlanEntry `
            -Path (Join-Path $LocalAppDataRoot 'RedSalamander\SelfTest') `
            -PathType 'Literal' `
            -Category 'legacy-selftest-root' `
            -Reason 'Fixed %LOCALAPPDATA%\RedSalamander\SelfTest artifacts predate REDSALAMANDER_TEST_ROOT and can poison later runs.'
    }

    if (-not [string]::IsNullOrWhiteSpace($TempRoot)) {
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
            $plan += New-RSTestSandboxCleanupPlanEntry `
                -Path (Join-Path $TempRoot $pattern) `
                -PathType 'Wildcard' `
                -Category 'legacy-temp-root' `
                -Reason 'Historical standalone/perf test temp roots must be purged before enforcing the unified TestSandbox wall.'
        }
    }

    foreach ($driveRoot in @($DriveRoots | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $plan += New-RSTestSandboxCleanupPlanEntry `
            -Path (Join-Path $driveRoot 'RedSalamanderCrossVolumeSelfTest_*') `
            -PathType 'Wildcard' `
            -Category 'legacy-cross-volume-root' `
            -Reason 'FileOps cross-volume selftests used drive-root scratch directories that survive crash/kill exits.'
    }

    return $plan
}

function Get-RSFixedDriveRoots {
    return @([System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq [System.IO.DriveType]::Fixed -and $_.IsReady } | ForEach-Object {
            $_.RootDirectory.FullName
        })
}

function Resolve-RSTestSandboxCleanupTargets {
    param(
        [object[]]$Plan = @()
    )

    $targets = @()
    foreach ($entry in @($Plan)) {
        if ($null -eq $entry -or [string]::IsNullOrWhiteSpace([string]$entry.Path)) {
            continue
        }

        $items = @()
        if ([string]$entry.PathType -eq 'Literal') {
            if (Test-Path -LiteralPath $entry.Path) {
                $items = @(Get-Item -LiteralPath $entry.Path -Force)
            }
        } else {
            $items = @(Get-ChildItem -Path $entry.Path -Directory -Force -ErrorAction SilentlyContinue)
        }

        foreach ($item in $items) {
            $targets += [pscustomobject]@{
                Path = $item.FullName
                SourcePath = $entry.Path
                PathType = $entry.PathType
                Category = $entry.Category
                Reason = $entry.Reason
            }
        }
    }

    return $targets
}

function Resolve-RSTestSandboxStaleRunTargets {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestRoot,

        [string]$RunId = '',

        [string[]]$AllowedRunIds = @(),

        [int64[]]$LiveProcessIds = @(),

        [hashtable]$LiveProcessStartTimesUtc = @{}
    )

    $normalizedTestRoot = ConvertTo-RSFullPath -Path $TestRoot
    $runsRoot = Join-Path $normalizedTestRoot 'runs'
    if (-not (Test-Path -LiteralPath $runsRoot)) {
        return @()
    }

    $allowedRunIds = @($AllowedRunIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (-not [string]::IsNullOrWhiteSpace($RunId)) {
        $allowedRunIds += $RunId
    }
    $allowedRunIds = @($allowedRunIds | Select-Object -Unique)

    $liveIds = @()
    $liveStartTimesUtc = @{}
    if ($PSBoundParameters.ContainsKey('LiveProcessIds')) {
        $liveIds = @($LiveProcessIds | ForEach-Object { [int64]$_ })
        foreach ($key in $LiveProcessStartTimesUtc.Keys) {
            $liveStartTimesUtc[[string]$key] = [datetime]$LiveProcessStartTimesUtc[$key]
        }
    } else {
        foreach ($process in @(Get-RSLiveProcessSnapshots)) {
            $liveIds += [int64]$process.ProcessId
            if ($null -ne $process.StartTimeUtc) {
                $liveStartTimesUtc[[string]$process.ProcessId] = [datetime]$process.StartTimeUtc
            }
        }
    }

    $targets = @()
    foreach ($runDir in @(Get-ChildItem -LiteralPath $runsRoot -Directory -Force -ErrorAction SilentlyContinue)) {
        $isAllowedRun = @($allowedRunIds | Where-Object {
                [string]::Equals($_, $runDir.Name, [System.StringComparison]::OrdinalIgnoreCase)
            }).Count -gt 0
        if ($isAllowedRun) {
            continue
        }

        $runInfo = ConvertFrom-RSTestRunId -RunId $runDir.Name
        if ($null -eq $runInfo) {
            continue
        }

        if (@($liveIds | Where-Object { $_ -eq $runInfo.ProcessId }).Count -gt 0) {
            $processKey = [string]$runInfo.ProcessId
            if (-not $liveStartTimesUtc.ContainsKey($processKey) -or
                $runDir.CreationTimeUtc -ge ([datetime]$liveStartTimesUtc[$processKey]).ToUniversalTime()) {
                continue
            }
        }

        $targets += [pscustomobject]@{
            Path = $runDir.FullName
            RunId = $runDir.Name
            ProcessId = $runInfo.ProcessId
            Category = 'stale-test-run-dir'
            Reason = "The owning process $($runInfo.ProcessId) is no longer live (or its PID has been reused by a newer process); the sibling TestSandbox run directory can be swept before this run starts."
        }
    }

    return $targets
}

function Remove-RSTestSandboxStaleRunDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestRoot,

        [string]$RunId = '',

        [string[]]$AllowedRunIds = @(),

        [int64[]]$LiveProcessIds = @(),

        [hashtable]$LiveProcessStartTimesUtc = @{}
    )

    $normalizedTestRoot = ConvertTo-RSFullPath -Path $TestRoot
    $runsRoot = (Join-Path $normalizedTestRoot 'runs').TrimEnd('\')
    $resolveParams = @{
        TestRoot = $normalizedTestRoot
        RunId = $RunId
        AllowedRunIds = @($AllowedRunIds)
    }
    if ($PSBoundParameters.ContainsKey('LiveProcessIds')) {
        $resolveParams['LiveProcessIds'] = @($LiveProcessIds)
        $resolveParams['LiveProcessStartTimesUtc'] = $LiveProcessStartTimesUtc
    }

    $results = @()
    foreach ($target in @(Resolve-RSTestSandboxStaleRunTargets @resolveParams)) {
        $targetPath = (ConvertTo-RSFullPath -Path $target.Path).TrimEnd('\')
        if (-not $targetPath.StartsWith("$runsRoot\", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to remove stale TestSandbox run outside '$runsRoot': $targetPath"
        }

        try {
            Remove-Item -LiteralPath $target.Path -Recurse -Force -ErrorAction Stop
            $target | Add-Member -NotePropertyName Status -NotePropertyValue 'Removed' -Force
            $target | Add-Member -NotePropertyName Error -NotePropertyValue $null -Force
        } catch {
            $message = $_.Exception.Message
            Write-Warning "Failed to remove stale TestSandbox run '$($target.Path)': $message"
            $target | Add-Member -NotePropertyName Status -NotePropertyValue 'Failed' -Force
            $target | Add-Member -NotePropertyName Error -NotePropertyValue $message -Force
        }

        $results += $target
    }

    return $results
}

function New-RSTestSandboxDiskAuditIssue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    [pscustomobject]@{
        category = $Category
        path = $Path
        reason = $Reason
    }
}

function Get-RSTestSandboxDiskAudit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestRoot,

        [string]$RunId = '',

        [string[]]$AllowedRunIds = @(),

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA,

        [string]$TempRoot = [System.IO.Path]::GetTempPath(),

        [string[]]$DriveRoots = @()
    )

    $normalizedTestRoot = ConvertTo-RSFullPath -Path $TestRoot
    $allowedRunIds = @($AllowedRunIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (-not [string]::IsNullOrWhiteSpace($RunId)) {
        $allowedRunIds += $RunId
    }
    $allowedRunIds = @($allowedRunIds | Select-Object -Unique)

    $issues = @()
    if (Test-Path -LiteralPath $normalizedTestRoot) {
        $runsRoot = Join-Path $normalizedTestRoot 'runs'
        foreach ($item in @(Get-ChildItem -LiteralPath $normalizedTestRoot -Force -ErrorAction SilentlyContinue)) {
            if ($item.Name -ne 'runs') {
                $issues += New-RSTestSandboxDiskAuditIssue `
                    -Category 'unexpected-test-root-child' `
                    -Path $item.FullName `
                    -Reason 'Unified TestSandbox root should contain only the runs directory.'
            }
        }

        if (Test-Path -LiteralPath $runsRoot) {
            foreach ($runDir in @(Get-ChildItem -LiteralPath $runsRoot -Directory -Force -ErrorAction SilentlyContinue)) {
                $isAllowedRun = @($allowedRunIds | Where-Object { $_ -eq $runDir.Name }).Count -gt 0
                if (-not $isAllowedRun) {
                    $issues += New-RSTestSandboxDiskAuditIssue `
                        -Category 'unexpected-test-run-dir' `
                        -Path $runDir.FullName `
                        -Reason 'Only the current runner-owned run directory should remain during post-run audit.'
                }
            }
        }
    }

    $legacyPlan = @(Get-RSTestSandboxLegacyCleanupPlan `
            -LocalAppDataRoot $LocalAppDataRoot `
            -TempRoot $TempRoot `
            -DriveRoots $DriveRoots)
    foreach ($target in @(Resolve-RSTestSandboxCleanupTargets -Plan $legacyPlan)) {
        $issues += New-RSTestSandboxDiskAuditIssue `
            -Category $target.Category `
            -Path $target.Path `
            -Reason $target.Reason
    }

    [pscustomobject]@{
        schema = 'red-salamander.test-sandbox-disk-audit.v1'
        is_clean = (@($issues).Count -eq 0)
        issue_count = @($issues).Count
        test_root = $normalizedTestRoot
        run_id = $RunId
        allowed_run_ids = @($allowedRunIds)
        issues = @($issues)
    }
}

function Get-RSSelfTestArtifactRoot {
    param(
        [string]$RepoRoot = '',

        [string]$RunId = '',

        [string]$SelfTestRootOverride = $env:REDSALAMANDER_SELFTEST_ROOT,

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA
    )

    if (-not [string]::IsNullOrWhiteSpace($RepoRoot) -or -not [string]::IsNullOrWhiteSpace($env:REDSALAMANDER_TEST_ROOT)) {
        return (New-RSTestRunContext -RepoRoot $RepoRoot -RunId $RunId).ArtifactRoot
    }

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

        [string]$CaseFilter = '',

        [uint32]$RepeatCount = 1,

        [string]$ShuffleSeed = '',

        [string]$PerfBudgetPath = '',

        [switch]$RequirePerfBudgets,

        [string]$SelfTestFlakyProofCase = '',

        [string]$SelfTestOrderProofCase = ''
    )

    $arguments = @($Flag)
    if ($FailFast) {
        $arguments += '--selftest-fail-fast'
    }
    if ($TimeoutMultiplier -ne 1.0) {
        $timeoutText = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:g}', $TimeoutMultiplier)
        $arguments += "--selftest-timeout-multiplier=$timeoutText"
    }
    if (-not [string]::IsNullOrWhiteSpace($CaseFilter)) {
        $arguments += "--selftest-case=$CaseFilter"
    }
    if ($RepeatCount -gt 1) {
        $arguments += "--selftest-repeat=$RepeatCount"
    }
    if (-not [string]::IsNullOrWhiteSpace($ShuffleSeed)) {
        $arguments += "--selftest-shuffle=$ShuffleSeed"
    }
    if (-not [string]::IsNullOrWhiteSpace($PerfBudgetPath)) {
        $arguments += "--selftest-perf-budget=$PerfBudgetPath"
    }
    if ($RequirePerfBudgets) {
        $arguments += '--selftest-require-perf-budgets'
    }
    if (-not [string]::IsNullOrWhiteSpace($SelfTestFlakyProofCase)) {
        $arguments += "--selftest-flaky-proof-case=$SelfTestFlakyProofCase"
    }
    if (-not [string]::IsNullOrWhiteSpace($SelfTestOrderProofCase)) {
        $arguments += "--selftest-order-proof-case=$SelfTestOrderProofCase"
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

function New-RSPesterInvokeParameters {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string[]]$Arguments = @(),

        [object]$InvokePesterCommand = (Get-Command Invoke-Pester -ErrorAction Stop)
    )

    $parameters = @{
        PassThru = $true
    }

    if ($InvokePesterCommand.Parameters.ContainsKey('Path')) {
        $parameters['Path'] = $Path
    } elseif ($InvokePesterCommand.Parameters.ContainsKey('Script')) {
        $parameters['Script'] = $Path
    } else {
        throw 'Invoke-Pester exposes neither -Path nor -Script; cannot run tooling Pester tests.'
    }

    for ($i = 0; $i -lt @($Arguments).Count; $i++) {
        $argument = [string]$Arguments[$i]
        switch ($argument) {
            '-ExcludeTag' {
                if ($i + 1 -ge @($Arguments).Count) {
                    throw 'Pester -ExcludeTag requires a value.'
                }
                $i++
                $existing = @()
                if ($parameters.ContainsKey('ExcludeTag')) {
                    $existing = @($parameters['ExcludeTag'])
                }
                $parameters['ExcludeTag'] = $existing + [string]$Arguments[$i]
            }
            '-Tag' {
                if ($i + 1 -ge @($Arguments).Count) {
                    throw 'Pester -Tag requires a value.'
                }
                $i++
                $existing = @()
                if ($parameters.ContainsKey('Tag')) {
                    $existing = @($parameters['Tag'])
                }
                $parameters['Tag'] = $existing + [string]$Arguments[$i]
            }
            default {
                throw "Unsupported Pester runner argument '$argument'. Add explicit binding support before forwarding it."
            }
        }
    }

    return $parameters
}

function Get-RSBuildScriptArguments {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'CI', 'Full')]
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

    if ($Suite -notin @('CI', 'Full')) {
        $arguments.ProjectName = 'RedSalamander'
    }

    return $arguments
}

function Get-RSBuildEnvironmentOverrides {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'CI', 'Full')]
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

        [uint32]$ExpectedRepeatCount = 1,

        [switch]$RequireExpectedCases
    )

    $actualNames = @($ActualCases | ForEach-Object {
            $property = $_.PSObject.Properties['name']
            if ($null -ne $property) {
                [string]$property.Value
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    $expectedGroups = @($ExpectedCaseNames | Group-Object | Where-Object { $_.Count -gt 1 })
    $expectedRepeat = [Math]::Max(1, [int]$ExpectedRepeatCount)

    $expectedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($name in @($ExpectedCaseNames)) {
        [void]$expectedSet.Add($name)
    }

    $actualGroups = @($actualNames | Group-Object | Where-Object {
            if ($expectedRepeat -le 1) {
                return $_.Count -gt 1
            }

            return ($expectedSet.Contains($_.Name) -and $_.Count -ne $expectedRepeat) -or
                ((-not $expectedSet.Contains($_.Name)) -and $_.Count -gt 1)
        })

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

function Set-RSObjectValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Object,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        [object]$Value = $null
    )

    if ($Object -is [System.Collections.IDictionary]) {
        $Object[$Name] = $Value
        return $Object
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) {
        $property.Value = $Value
    } else {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
    }

    return $Object
}

function Test-RSTestResultPassed {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Result
    )

    $exitCode = [int](Get-RSObjectValue -Object $Result -Names @('ExitCode', 'exit_code') -DefaultValue 1)
    $failed = [int](Get-RSObjectValue -Object $Result -Names @('Failed', 'failed') -DefaultValue 0)
    return ($exitCode -eq 0 -and $failed -eq 0)
}

function Add-RSTestResultClassification {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Result,

        [object[]]$RetryResults = @()
    )

    $retryAttempts = @($RetryResults)
    $classification = ''
    $classificationReason = ''

    if (Test-RSTestResultPassed -Result $Result) {
        $classification = 'PASSED'
        $classificationReason = 'Initial attempt passed.'
    } elseif (@($retryAttempts).Count -eq 0) {
        $classification = 'UNCLASSIFIED_FAILURE'
        $classificationReason = 'Initial attempt failed and no retry classification evidence was collected.'
    } else {
        $shuffleTriageAttempts = @($retryAttempts | Where-Object {
                [string](Get-RSObjectValue -Object $_ -Names @('mode', 'Mode') -DefaultValue '') -eq 'shuffle-triage'
            })
        $hasShuffleTriage = @($shuffleTriageAttempts).Count -gt 0
        $hasFailedShuffleTriage = @($shuffleTriageAttempts | Where-Object {
                -not [bool](Get-RSObjectValue -Object $_ -Names @('passed', 'Passed') -DefaultValue $false)
            }).Count -gt 0

        $allRetriesPassed = $true
        foreach ($retryAttempt in $retryAttempts) {
            $retryPassed = [bool](Get-RSObjectValue -Object $retryAttempt -Names @('passed', 'Passed') -DefaultValue $false)
            if (-not $retryPassed) {
                $allRetriesPassed = $false
                break
            }
        }

        if ($hasFailedShuffleTriage) {
            $classification = 'REGRESSION'
            $classificationReason = 'Initial broad self-test attempt failed, failed cases passed when rerun in isolation, but shuffle triage reproduced a failure; this is a blocking isolation/order regression.'
        } elseif ($allRetriesPassed) {
            $kind = [string](Get-RSObjectValue -Object $Result -Names @('Kind', 'kind') -DefaultValue '')
            $nonShuffleAttempts = @($retryAttempts | Where-Object {
                    [string](Get-RSObjectValue -Object $_ -Names @('mode', 'Mode') -DefaultValue '') -ne 'shuffle-triage'
                })
            $allFailedCaseRetries = @($nonShuffleAttempts).Count -gt 0
            foreach ($retryAttempt in $nonShuffleAttempts) {
                $mode = [string](Get-RSObjectValue -Object $retryAttempt -Names @('mode', 'Mode') -DefaultValue '')
                if ($mode -ne 'failed-case') {
                    $allFailedCaseRetries = $false
                    break
                }
            }

            if ($kind -eq 'SelfTest' -and $allFailedCaseRetries -and $hasShuffleTriage) {
                $classification = 'FLAKY'
                $classificationReason = 'Initial broad self-test attempt failed, failed cases passed when rerun in isolation, and shuffle triage passed; this remains a blocking flaky-test repair item.'
            } elseif ($kind -eq 'SelfTest' -and $allFailedCaseRetries) {
                $classification = 'ISOLATION_SUSPECT'
                $classificationReason = 'Initial broad self-test attempt failed, but failed cases passed when rerun in isolation; shuffle triage is still required.'
            } else {
                $classification = 'FLAKY'
                $classificationReason = 'Initial attempt failed, but retry evidence passed; this remains a blocking flaky-test repair item.'
            }
        } else {
            $classification = 'REGRESSION'
            $classificationReason = 'Initial attempt failed and retry evidence failed again.'
        }
    }

    $Result = Set-RSObjectValue -Object $Result -Name 'Classification' -Value $classification
    $Result = Set-RSObjectValue -Object $Result -Name 'ClassificationReason' -Value $classificationReason
    $Result = Set-RSObjectValue -Object $Result -Name 'RetryAttempts' -Value $retryAttempts

    return $Result
}

function Get-RSSelfTestClassificationRetryPlan {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [string[]]$FailureNames = @(),

        [string[]]$ShuffleTriageSeeds = @()
    )

    $entryArguments = @($Entry.Arguments)
    $baseArguments = @($entryArguments | Where-Object { $_ -notlike '--selftest-case=*' })
    $originalCaseFilter = ''
    foreach ($argument in $entryArguments) {
        if ($argument -like '--selftest-case=*') {
            $originalCaseFilter = $argument.Substring('--selftest-case='.Length)
            break
        }
    }

    $failedCaseNames = @($FailureNames | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_) -and $_ -notin @('selftest_result_coverage', 'setup')
        } | Sort-Object -Unique)

    $plans = @()
    foreach ($caseName in $failedCaseNames) {
        $plans += [pscustomobject]@{
            name = $caseName
            mode = 'failed-case'
            arguments = @($baseArguments) + @("--selftest-case=$caseName")
            shuffle_seed = ''
        }
    }

    if (@($failedCaseNames).Count -eq 0) {
        return $plans
    }

    $shuffleBaseArguments = @($baseArguments | Where-Object {
            $_ -notlike '--selftest-shuffle=*' -and $_ -notlike '--selftest-repeat=*'
        })
    if (-not [string]::IsNullOrWhiteSpace($originalCaseFilter)) {
        $shuffleBaseArguments += "--selftest-case=$originalCaseFilter"
    }

    foreach ($seed in @($ShuffleTriageSeeds)) {
        if ([string]::IsNullOrWhiteSpace($seed)) {
            continue
        }

        $plans += [pscustomobject]@{
            name = [string](Get-RSObjectValue -Object $Entry -Names @('Name', 'name') -DefaultValue '')
            mode = 'shuffle-triage'
            arguments = @($shuffleBaseArguments) + @("--selftest-shuffle=$seed")
            shuffle_seed = $seed
        }
    }

    return $plans
}

function Get-RSTestQuarantineDate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    try {
        return [datetime]::ParseExact($Value, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture).Date
    } catch {
        return $null
    }
}

function Get-RSTestQuarantineStatus {
    param(
        [object[]]$Entries = @(),

        [datetime]$Now = (Get-Date),

        [int]$MaxDays = 30
    )

    $errors = @()
    $activeEntries = @()
    $invalidEntries = @()
    $requiredFields = @(
        'harness',
        'name',
        'owner',
        'opened',
        'expires',
        'issue',
        'root_cause_hypothesis',
        'fix_or_replace_plan'
    )
    $today = $Now.ToUniversalTime().Date
    $index = 0

    foreach ($entry in @($Entries)) {
        ++$index
        $entryErrors = @()
        foreach ($field in $requiredFields) {
            $value = [string](Get-RSObjectValue -Object $entry -Names @($field) -DefaultValue '')
            if ([string]::IsNullOrWhiteSpace($value)) {
                $entryErrors += "entry $index missing $field"
            }
        }

        $openedText = [string](Get-RSObjectValue -Object $entry -Names @('opened') -DefaultValue '')
        $expiresText = [string](Get-RSObjectValue -Object $entry -Names @('expires') -DefaultValue '')
        $openedDate = Get-RSTestQuarantineDate -Value $openedText
        $expiresDate = Get-RSTestQuarantineDate -Value $expiresText

        if ($null -eq $openedDate) {
            $entryErrors += "entry $index opened must use YYYY-MM-DD"
        }
        if ($null -eq $expiresDate) {
            $entryErrors += "entry $index expires must use YYYY-MM-DD"
        }
        if ($null -ne $openedDate -and $openedDate -gt $today) {
            $entryErrors += "entry $index opened is in the future"
        }
        if ($null -ne $expiresDate -and $expiresDate -lt $today) {
            $entryErrors += "entry $index expired on $($expiresDate.ToString('yyyy-MM-dd'))"
        }
        if ($null -ne $openedDate -and $null -ne $expiresDate) {
            if ($expiresDate -lt $openedDate) {
                $entryErrors += "entry $index expires before opened"
            }
            if (($expiresDate - $openedDate).TotalDays -gt $MaxDays) {
                $entryErrors += "entry $index expires more than $MaxDays days after opened"
            }
        }

        if (@($entryErrors).Count -gt 0) {
            $errors += $entryErrors
            $invalidEntries += $entry
        } else {
            $activeEntries += $entry
        }
    }

    [pscustomobject]@{
        IsValid = (@($errors).Count -eq 0)
        HasBlockingEntries = (@($errors).Count -gt 0 -or @($activeEntries).Count -gt 0)
        TotalCount = @($Entries).Count
        ActiveEntries = @($activeEntries)
        InvalidEntries = @($invalidEntries)
        Errors = @($errors)
    }
}

function Read-RSTestQuarantineFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [datetime]$Now = (Get-Date)
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return (Get-RSTestQuarantineStatus -Entries @() -Now $Now)
    }

    $entries = @()
    $errors = @()
    $lineNumber = 0
    foreach ($line in @(Get-Content -LiteralPath $Path)) {
        ++$lineNumber
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $entries += ($line | ConvertFrom-Json)
        } catch {
            $errors += "line $lineNumber is not valid JSON: $($_.Exception.Message)"
        }
    }

    $status = Get-RSTestQuarantineStatus -Entries $entries -Now $Now
    if (@($errors).Count -gt 0) {
        $allErrors = @($status.Errors) + $errors
        $status = [pscustomobject]@{
            IsValid = $false
            HasBlockingEntries = $true
            TotalCount = $status.TotalCount
            ActiveEntries = @($status.ActiveEntries)
            InvalidEntries = @($status.InvalidEntries)
            Errors = @($allErrors)
        }
    }

    return $status
}

function Get-RSTestQuarantineRepairPlan {
    param(
        [AllowNull()]
        [object]$QuarantineStatus = $null,

        [object[]]$TestPlan = @(),

        [hashtable]$HarnessCaseMap = @{}
    )

    $activeEntries = @()
    $invalidEntries = @()
    $errors = @()
    $statusIsValid = $true
    $statusHasBlockingEntries = $false

    if ($null -ne $QuarantineStatus) {
        $activeEntries = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('ActiveEntries', 'active_entries') -DefaultValue @())
        $invalidEntries = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('InvalidEntries', 'invalid_entries') -DefaultValue @())
        $errors = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('Errors', 'errors') -DefaultValue @())
        $statusIsValid = [bool](Get-RSObjectValue -Object $QuarantineStatus -Names @('IsValid', 'is_valid') -DefaultValue $true)
        $statusHasBlockingEntries = [bool](Get-RSObjectValue -Object $QuarantineStatus -Names @('HasBlockingEntries', 'has_blocking_entries') -DefaultValue $false)
    }

    $attempts = @()
    $resolvedInvalidEntries = @($invalidEntries)
    $resolvedErrors = @($errors)

    foreach ($entry in @($activeEntries)) {
        $harness = [string](Get-RSObjectValue -Object $entry -Names @('harness') -DefaultValue '')
        $name = [string](Get-RSObjectValue -Object $entry -Names @('name') -DefaultValue '')
        $matchingPlanEntry = @($TestPlan | Where-Object {
                [string]::Equals(
                    [string](Get-RSObjectValue -Object $_ -Names @('Name', 'name') -DefaultValue ''),
                    $harness,
                    [System.StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1)

        if (@($matchingPlanEntry).Count -eq 0) {
            $resolvedInvalidEntries += $entry
            $resolvedErrors += "quarantine entry [$harness] $name has no matching case in the harness adapter"
            continue
        }

        $planEntry = $matchingPlanEntry[0]
        $kind = [string](Get-RSObjectValue -Object $planEntry -Names @('Kind', 'kind') -DefaultValue '')
        $mode = if ($kind -eq 'SelfTest') { 'selftest-case' } else { 'entry' }
        if ($mode -eq 'selftest-case' -and $null -ne $HarnessCaseMap -and $HarnessCaseMap.ContainsKey($harness)) {
            $knownCaseNames = @($HarnessCaseMap[$harness])
            $caseMatches = @($knownCaseNames | Where-Object {
                    [string]::Equals([string]$_, $name, [System.StringComparison]::OrdinalIgnoreCase)
                })
            if (@($caseMatches).Count -eq 0) {
                $resolvedInvalidEntries += $entry
                $resolvedErrors += "quarantine entry [$harness] $name has no matching case in the harness adapter"
                continue
            }
        }

        $arguments = @($planEntry.Arguments)
        if ($mode -eq 'selftest-case') {
            $arguments = @($arguments | Where-Object { $_ -notlike '--selftest-case=*' })
            $arguments += "--selftest-case=$name"
        }

        $attempts += [pscustomobject]@{
            harness = $harness
            name = $name
            mode = $mode
            kind = $kind
            path = [string](Get-RSObjectValue -Object $planEntry -Names @('Path', 'path') -DefaultValue '')
            arguments = @($arguments)
            working_directory = [string](Get-RSObjectValue -Object $planEntry -Names @('WorkingDirectory', 'working_directory') -DefaultValue '')
            json_name = [string](Get-RSObjectValue -Object $planEntry -Names @('JsonName', 'json_name') -DefaultValue '')
            owner = [string](Get-RSObjectValue -Object $entry -Names @('owner') -DefaultValue '')
            opened = [string](Get-RSObjectValue -Object $entry -Names @('opened') -DefaultValue '')
            expires = [string](Get-RSObjectValue -Object $entry -Names @('expires') -DefaultValue '')
            issue = [string](Get-RSObjectValue -Object $entry -Names @('issue') -DefaultValue '')
            root_cause_hypothesis = [string](Get-RSObjectValue -Object $entry -Names @('root_cause_hypothesis') -DefaultValue '')
            fix_or_replace_plan = [string](Get-RSObjectValue -Object $entry -Names @('fix_or_replace_plan') -DefaultValue '')
        }
    }

    [pscustomobject]@{
        IsValid = ($statusIsValid -and @($resolvedInvalidEntries).Count -eq 0 -and @($resolvedErrors).Count -eq 0)
        HasBlockingEntries = ($statusHasBlockingEntries -or @($attempts).Count -gt 0 -or @($resolvedInvalidEntries).Count -gt 0)
        Attempts = @($attempts)
        InvalidEntries = @($resolvedInvalidEntries)
        Errors = @($resolvedErrors)
    }
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
        attempt = [int](Get-RSObjectValue -Object $Case -Names @('attempt', 'repeat_index', 'repeatIndex') -DefaultValue 1)
    }
}

function Convert-RSTestRetryAttemptForRunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$RetryAttempt
    )

    [pscustomobject]@{
        name = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('name', 'Name') -DefaultValue '(unnamed retry)')
        mode = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('mode', 'Mode') -DefaultValue '')
        exit_code = [int](Get-RSObjectValue -Object $RetryAttempt -Names @('exit_code', 'ExitCode') -DefaultValue 0)
        passed = [bool](Get-RSObjectValue -Object $RetryAttempt -Names @('passed', 'Passed') -DefaultValue $false)
        duration_ms = [uint64](Get-RSObjectValue -Object $RetryAttempt -Names @('duration_ms', 'DurationMs', 'WallMs') -DefaultValue 0)
        output_log_path = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('output_log_path', 'OutputLogPath') -DefaultValue '')
        reason = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('reason', 'Reason') -DefaultValue '')
        shuffle_seed = [string](Get-RSObjectValue -Object $RetryAttempt -Names @('shuffle_seed', 'ShuffleSeed') -DefaultValue '')
    }
}

function Convert-RSTestQuarantineRepairAttemptForRunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$RepairAttempt
    )

    [pscustomobject]@{
        harness = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('harness', 'Harness') -DefaultValue '')
        name = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('name', 'Name') -DefaultValue '')
        mode = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('mode', 'Mode') -DefaultValue '')
        owner = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('owner', 'Owner') -DefaultValue '')
        expires = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('expires', 'Expires') -DefaultValue '')
        issue = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('issue', 'Issue') -DefaultValue '')
        exit_code = [int](Get-RSObjectValue -Object $RepairAttempt -Names @('exit_code', 'ExitCode') -DefaultValue -1)
        passed = [bool](Get-RSObjectValue -Object $RepairAttempt -Names @('passed', 'Passed') -DefaultValue $false)
        duration_ms = [uint64](Get-RSObjectValue -Object $RepairAttempt -Names @('duration_ms', 'DurationMs', 'WallMs') -DefaultValue 0)
        output_log_path = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('output_log_path', 'OutputLogPath') -DefaultValue '')
        reason = [string](Get-RSObjectValue -Object $RepairAttempt -Names @('reason', 'Reason') -DefaultValue '')
    }
}

function New-RSTestRunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'CI', 'Full')]
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

        [string]$TestRoot = '',

        [string]$RunId = '',

        [object]$QuarantineStatus = $null,

        [object[]]$QuarantineRepairAttempts = @(),

        [object]$TestSandboxAudit = $null,

        [object[]]$Results = @()
    )

    $totalPassed = 0
    $totalFailed = 0
    $totalSkipped = 0
    $classificationCounts = [ordered]@{
        passed = 0
        flaky = 0
        regression = 0
        isolation_suspect = 0
        unclassified_failure = 0
    }
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
        $retryAttempts = @(Get-RSObjectValue -Object $result -Names @('RetryAttempts', 'retry_attempts') -DefaultValue @() | ForEach-Object {
                Convert-RSTestRetryAttemptForRunSummary -RetryAttempt $_
            })
        $classification = [string](Get-RSObjectValue -Object $result -Names @('Classification', 'classification') -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($classification)) {
            $classification = if (Test-RSTestResultPassed -Result $result) { 'PASSED' } else { 'UNCLASSIFIED_FAILURE' }
        }
        $classificationReason = [string](Get-RSObjectValue -Object $result -Names @('ClassificationReason', 'classification_reason') -DefaultValue '')

        switch ($classification) {
            'PASSED' { ++$classificationCounts.passed }
            'FLAKY' { ++$classificationCounts.flaky }
            'REGRESSION' { ++$classificationCounts.regression }
            'ISOLATION_SUSPECT' { ++$classificationCounts.isolation_suspect }
            default { ++$classificationCounts.unclassified_failure }
        }

        $summarySuites += [pscustomobject]@{
            suite = [string](Get-RSObjectValue -Object $result -Names @('Name', 'suite', 'name') -DefaultValue '(unnamed suite)')
            kind = [string](Get-RSObjectValue -Object $result -Names @('Kind', 'kind') -DefaultValue '')
            exit_code = [int](Get-RSObjectValue -Object $result -Names @('ExitCode', 'exit_code') -DefaultValue 0)
            duration_ms = [uint64](Get-RSObjectValue -Object $result -Names @('WallMs', 'DurationMs', 'duration_ms') -DefaultValue 0)
            classification = $classification
            classification_reason = $classificationReason
            shuffle_seed = [string](Get-RSObjectValue -Object $result -Names @('ShuffleSeed', 'shuffle_seed') -DefaultValue '')
            repeat_count = [int](Get-RSObjectValue -Object $result -Names @('RepeatCount', 'repeat_count') -DefaultValue 1)
            passed = $passed
            failed = $failed
            skipped = $skipped
            output_log_path = [string](Get-RSObjectValue -Object $result -Names @('OutputLogPath', 'output_log_path') -DefaultValue '')
            retry_attempts = $retryAttempts
            cases = $cases
            failures = $failures
        }
    }

    $quarantineActiveEntries = @()
    $quarantineInvalidEntries = @()
    $quarantineErrors = @()
    $quarantineIsValid = $true
    $quarantineHasBlockingEntries = $false
    if ($null -ne $QuarantineStatus) {
        $quarantineActiveEntries = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('ActiveEntries', 'active_entries') -DefaultValue @())
        $quarantineInvalidEntries = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('InvalidEntries', 'invalid_entries') -DefaultValue @())
        $quarantineErrors = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('Errors', 'errors') -DefaultValue @())
        $quarantineIsValid = [bool](Get-RSObjectValue -Object $QuarantineStatus -Names @('IsValid', 'is_valid') -DefaultValue $true)
        $quarantineHasBlockingEntries = [bool](Get-RSObjectValue -Object $QuarantineStatus -Names @('HasBlockingEntries', 'has_blocking_entries') -DefaultValue $false)
    }
    $quarantineRepairAttempts = @($QuarantineRepairAttempts | ForEach-Object {
            Convert-RSTestQuarantineRepairAttemptForRunSummary -RepairAttempt $_
        })
    $quarantineRepairReproducedCount = @($quarantineRepairAttempts | Where-Object { -not [bool]$_.passed }).Count

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
        test_root = $TestRoot
        run_id = $RunId
        fail_fast = $FailFast
        timeout_scale = $TimeoutMultiplier
        case_filter = $CaseFilter
        exit_code = $ExitCode
        passed = $totalPassed
        failed = $totalFailed
        skipped = $totalSkipped
        total = ($totalPassed + $totalFailed + $totalSkipped)
        classifications = [pscustomobject]$classificationCounts
        quarantine = [pscustomobject]@{
            is_valid = $quarantineIsValid
            has_blocking_entries = $quarantineHasBlockingEntries
            active_count = @($quarantineActiveEntries).Count
            invalid_count = @($quarantineInvalidEntries).Count
            active_entries = $quarantineActiveEntries
            invalid_entries = $quarantineInvalidEntries
            errors = $quarantineErrors
            repair_attempt_count = @($quarantineRepairAttempts).Count
            repair_reproduced_count = $quarantineRepairReproducedCount
            repair_attempts = $quarantineRepairAttempts
        }
        test_sandbox_audit = $TestSandboxAudit
        suites = $summarySuites
    }
}

function Convert-RSTestRunSummaryToCaseHistoryRows {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Summary
    )

    $runId = [string](Get-RSObjectValue -Object $Summary -Names @('run_id', 'RunId') -DefaultValue '')
    $runStartedUtc = [string](Get-RSObjectValue -Object $Summary -Names @('run_started_utc', 'RunStartedUtc') -DefaultValue '')
    $suiteSelection = [string](Get-RSObjectValue -Object $Summary -Names @('suite', 'Suite') -DefaultValue '')
    $rows = @()

    foreach ($suite in @(Get-RSObjectValue -Object $Summary -Names @('suites', 'Suites') -DefaultValue @())) {
        $harness = [string](Get-RSObjectValue -Object $suite -Names @('suite', 'name', 'Name') -DefaultValue '')
        $classification = [string](Get-RSObjectValue -Object $suite -Names @('classification', 'Classification') -DefaultValue '')
        $suiteSeed = [string](Get-RSObjectValue -Object $suite -Names @('shuffle_seed', 'ShuffleSeed', 'seed') -DefaultValue '')

        foreach ($case in @(Get-RSObjectValue -Object $suite -Names @('cases', 'Cases') -DefaultValue @())) {
            $rows += [pscustomobject][ordered]@{
                schema = 'red-salamander.case-history.v1'
                run_id = $runId
                run_started_utc = $runStartedUtc
                suite = $suiteSelection
                harness = $harness
                case = [string](Get-RSObjectValue -Object $case -Names @('name', 'Name') -DefaultValue '')
                duration_ms = [uint64](Get-RSObjectValue -Object $case -Names @('duration_ms', 'durationMs', 'Duration') -DefaultValue 0)
                status = [string](Get-RSObjectValue -Object $case -Names @('status', 'Status') -DefaultValue '')
                reason = [string](Get-RSObjectValue -Object $case -Names @('reason', 'Reason') -DefaultValue '')
                classification = $classification
                seed = $suiteSeed
                attempt = [int](Get-RSObjectValue -Object $case -Names @('attempt', 'repeat_index', 'repeatIndex') -DefaultValue 1)
                source = 'case'
                mode = ''
            }
        }

        $retryAttemptIndex = 0
        foreach ($retryAttempt in @(Get-RSObjectValue -Object $suite -Names @('retry_attempts', 'RetryAttempts') -DefaultValue @())) {
            ++$retryAttemptIndex
            $retryPassed = [bool](Get-RSObjectValue -Object $retryAttempt -Names @('passed', 'Passed') -DefaultValue $false)
            $retrySeed = [string](Get-RSObjectValue -Object $retryAttempt -Names @('shuffle_seed', 'ShuffleSeed') -DefaultValue '')
            if ([string]::IsNullOrWhiteSpace($retrySeed)) {
                $retrySeed = $suiteSeed
            }

            $rows += [pscustomobject][ordered]@{
                schema = 'red-salamander.case-history.v1'
                run_id = $runId
                run_started_utc = $runStartedUtc
                suite = $suiteSelection
                harness = $harness
                case = [string](Get-RSObjectValue -Object $retryAttempt -Names @('name', 'Name') -DefaultValue '')
                duration_ms = [uint64](Get-RSObjectValue -Object $retryAttempt -Names @('duration_ms', 'DurationMs', 'WallMs') -DefaultValue 0)
                status = if ($retryPassed) { 'passed' } else { 'failed' }
                reason = [string](Get-RSObjectValue -Object $retryAttempt -Names @('reason', 'Reason') -DefaultValue '')
                classification = $classification
                seed = $retrySeed
                attempt = $retryAttemptIndex
                source = 'retry'
                mode = [string](Get-RSObjectValue -Object $retryAttempt -Names @('mode', 'Mode') -DefaultValue '')
            }
        }
    }

    return $rows
}

function Convert-RSTestCaseHistoryRowsToJsonl {
    param(
        [object[]]$Rows = @()
    )

    return (@($Rows) | ForEach-Object {
            $_ | ConvertTo-Json -Compress -Depth 8
        }) -join "`n"
}

function Convert-RSTestCaseHistoryRowsToDashboardMarkdown {
    param(
        [object[]]$Rows = @(),

        [Parameter(Mandatory = $true)]
        [object]$Summary,

        [int]$TopCount = 20
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $runId = [string](Get-RSObjectValue -Object $Summary -Names @('run_id', 'RunId') -DefaultValue '')
    $suite = [string](Get-RSObjectValue -Object $Summary -Names @('suite', 'Suite') -DefaultValue '')
    $lines.Add('# Run-AllTests Case Dashboard')
    $lines.Add('')
    $lines.Add(('- Run: `{0}`' -f $runId))
    $lines.Add(('- Suite: `{0}`' -f $suite))
    $lines.Add('')

    $lines.Add('## Top Slowest Cases')
    $lines.Add('')
    $lines.Add('| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |')
    $lines.Add('|---|---|---|---|---|---:|---|---:|')
    $topRows = @($Rows | Sort-Object -Property duration_ms -Descending | Select-Object -First $TopCount)
    foreach ($row in $topRows) {
        $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} |' -f `
                    (ConvertTo-RSMarkdownTableCell $row.harness), `
                    (ConvertTo-RSMarkdownTableCell $row.case), `
                    (ConvertTo-RSMarkdownTableCell $row.source), `
                    (ConvertTo-RSMarkdownTableCell $row.status), `
                    (ConvertTo-RSMarkdownTableCell $row.classification), `
                    (ConvertTo-RSMarkdownTableCell $row.duration_ms), `
                    (ConvertTo-RSMarkdownTableCell $row.seed), `
                    (ConvertTo-RSMarkdownTableCell $row.attempt)))
    }
    if (@($topRows).Count -eq 0) {
        $lines.Add('| (none) |  |  |  |  | 0 |  | 0 |')
    }

    $lines.Add('')
    $lines.Add('## Failure And Retry Evidence')
    $lines.Add('')
    $lines.Add('| Harness | Case | Source | Status | Classification | Seed | Reason |')
    $lines.Add('|---|---|---|---|---|---|---|')
    $evidenceRows = @($Rows | Where-Object { $_.status -ne 'passed' -or $_.source -eq 'retry' })
    foreach ($row in $evidenceRows) {
        $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} |' -f `
                    (ConvertTo-RSMarkdownTableCell $row.harness), `
                    (ConvertTo-RSMarkdownTableCell $row.case), `
                    (ConvertTo-RSMarkdownTableCell $row.source), `
                    (ConvertTo-RSMarkdownTableCell $row.status), `
                    (ConvertTo-RSMarkdownTableCell $row.classification), `
                    (ConvertTo-RSMarkdownTableCell $row.seed), `
                    (ConvertTo-RSMarkdownTableCell $row.reason)))
    }
    if (@($evidenceRows).Count -eq 0) {
        $lines.Add('| (none) |  |  |  |  |  |  |')
    }

    return ($lines -join "`n")
}

function ConvertTo-RSMarkdownTableCell {
    param(
        [AllowNull()]
        [object]$Value = $null
    )

    if ($null -eq $Value) {
        return ''
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    return $text.Replace('|', '\|').Replace("`r`n", '<br>').Replace("`n", '<br>').Replace("`r", '<br>')
}

function Convert-RSTestRunSummaryToGitHubStepSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Summary
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $exitCode = [int](Get-RSObjectValue -Object $Summary -Names @('exit_code', 'ExitCode') -DefaultValue 1)
    $status = if ($exitCode -eq 0) { 'PASSED' } else { 'FAILED' }
    $suite = [string](Get-RSObjectValue -Object $Summary -Names @('suite', 'Suite') -DefaultValue '')
    $runId = [string](Get-RSObjectValue -Object $Summary -Names @('run_id', 'RunId') -DefaultValue '')
    $markdownTick = [char]96

    $lines.Add('## RedSalamander Test Summary')
    $lines.Add('')
    $lines.Add("- Status: **$status**")
    $lines.Add("- Suite: $markdownTick$suite$markdownTick")
    if (-not [string]::IsNullOrWhiteSpace($runId)) {
        $lines.Add("- Run id: $markdownTick$runId$markdownTick")
    }
    $lines.Add(("- Cases: passed={0} failed={1} skipped={2} total={3}" -f `
                (Get-RSObjectValue -Object $Summary -Names @('passed') -DefaultValue 0), `
                (Get-RSObjectValue -Object $Summary -Names @('failed') -DefaultValue 0), `
                (Get-RSObjectValue -Object $Summary -Names @('skipped') -DefaultValue 0), `
                (Get-RSObjectValue -Object $Summary -Names @('total') -DefaultValue 0)))

    $classifications = Get-RSObjectValue -Object $Summary -Names @('classifications') -DefaultValue $null
    if ($null -ne $classifications) {
        $lines.Add(("- Classifications: passed={0} flaky={1} regression={2} isolation_suspect={3} unclassified_failure={4}" -f `
                    (Get-RSObjectValue -Object $classifications -Names @('passed') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $classifications -Names @('flaky') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $classifications -Names @('regression') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $classifications -Names @('isolation_suspect') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $classifications -Names @('unclassified_failure') -DefaultValue 0)))
    }

    $lines.Add('')
    $lines.Add('### Suites')
    $lines.Add('')
    $lines.Add('| Suite | Classification | Passed | Failed | Skipped | Retry attempts |')
    $lines.Add('|---|---:|---:|---:|---:|---:|')
    foreach ($summarySuite in @(Get-RSObjectValue -Object $Summary -Names @('suites') -DefaultValue @())) {
        $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} |' -f `
                    (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $summarySuite -Names @('suite') -DefaultValue '')), `
                    (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $summarySuite -Names @('classification') -DefaultValue '')), `
                    (Get-RSObjectValue -Object $summarySuite -Names @('passed') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $summarySuite -Names @('failed') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $summarySuite -Names @('skipped') -DefaultValue 0), `
                    @((Get-RSObjectValue -Object $summarySuite -Names @('retry_attempts') -DefaultValue @())).Count))
    }

    $quarantine = Get-RSObjectValue -Object $Summary -Names @('quarantine') -DefaultValue $null
    if ($null -ne $quarantine) {
        $lines.Add('')
        $lines.Add('### Quarantine')
        $lines.Add('')
        $lines.Add(("- Active: {0}; invalid: {1}; repair_attempts: {2}; repair_reproduced: {3}" -f `
                    (Get-RSObjectValue -Object $quarantine -Names @('active_count') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $quarantine -Names @('invalid_count') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $quarantine -Names @('repair_attempt_count') -DefaultValue 0), `
                    (Get-RSObjectValue -Object $quarantine -Names @('repair_reproduced_count') -DefaultValue 0)))

        $activeEntries = @(Get-RSObjectValue -Object $quarantine -Names @('active_entries') -DefaultValue @())
        if (@($activeEntries).Count -gt 0) {
            $lines.Add('')
            $lines.Add('| Active harness | Case | Owner | Expires | Issue |')
            $lines.Add('|---|---|---|---|---|')
            foreach ($entry in $activeEntries) {
                $lines.Add(('| {0} | {1} | {2} | {3} | {4} |' -f `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('harness') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('name') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('owner') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('expires') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $entry -Names @('issue') -DefaultValue ''))))
            }
        }

        $repairAttempts = @(Get-RSObjectValue -Object $quarantine -Names @('repair_attempts') -DefaultValue @())
        if (@($repairAttempts).Count -gt 0) {
            $lines.Add('')
            $lines.Add('| Repair status | Harness | Case | Owner | Expires | Reason |')
            $lines.Add('|---|---|---|---|---|---|')
            foreach ($attempt in $repairAttempts) {
                $repairStatus = if ([bool](Get-RSObjectValue -Object $attempt -Names @('passed') -DefaultValue $false)) { 'REPAIR PASS' } else { 'REPAIR FAIL' }
                $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} |' -f `
                            $repairStatus, `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('harness') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('name') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('owner') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('expires') -DefaultValue '')), `
                            (ConvertTo-RSMarkdownTableCell (Get-RSObjectValue -Object $attempt -Names @('reason') -DefaultValue ''))))
            }
        }

        $errors = @(Get-RSObjectValue -Object $quarantine -Names @('errors') -DefaultValue @())
        if (@($errors).Count -gt 0) {
            $lines.Add('')
            $lines.Add('Invalid quarantine entries:')
            foreach ($errorMessage in $errors) {
                $lines.Add("- $(ConvertTo-RSMarkdownTableCell $errorMessage)")
            }
        }
    }

    return (($lines.ToArray()) -join [Environment]::NewLine) + [Environment]::NewLine
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

        [uint32]$RepeatCount = 1,

        [string]$ShuffleSeed = '',

        [string]$PerfBudgetPath = '',

        [switch]$RequirePerfBudgets,

        [string]$SelfTestFlakyProofCase = '',

        [string]$SelfTestOrderProofCase = ''
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
        { $_ -in @('CI', 'Full') } {
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
            -Arguments (Get-RSSelfTestArguments -Flag $selfTest.Flag -TimeoutMultiplier $TimeoutMultiplier -FailFast:$FailFast -CaseFilter $CaseFilter -RepeatCount $RepeatCount -ShuffleSeed $ShuffleSeed -PerfBudgetPath $PerfBudgetPath -RequirePerfBudgets:$RequirePerfBudgets -SelfTestFlakyProofCase $SelfTestFlakyProofCase -SelfTestOrderProofCase $SelfTestOrderProofCase) `
            -WorkingDirectory $RepoRoot `
            -JsonName $selfTest.JsonName
    }

    if ($Suite -eq 'CI') {
        foreach ($dxUiSuite in @(
                'Grid',
                'Theme',
                'Control',
                'Menu',
                'NewControls',
                'TextField',
                'NativeTextInput',
                'ComboBox',
                'WindowHost',
                'Tree',
                'MultilineText',
                'ReadOnly',
                'Tooltip',
                'Rendering',
                'Animation',
                'Accessibility'
            )) {
            $plan += New-RSTestRunPlanEntry `
                -Name "DxUiTests.$dxUiSuite" `
                -Kind 'Executable' `
                -Path (Join-Path $buildOutputDir 'DxUiTests.exe') `
                -Arguments @("--suite=$dxUiSuite") `
                -WorkingDirectory $buildOutputDir
        }

        $plan += New-RSTestRunPlanEntry `
            -Name 'FileSystemCurlTests' `
            -Kind 'Executable' `
            -Path (Join-Path $buildOutputDir 'FileSystemCurlTests.exe') `
            -WorkingDirectory $buildOutputDir

        $plan += New-RSTestRunPlanEntry `
            -Name 'ViewerPETests' `
            -Kind 'Executable' `
            -Path (Join-Path $buildOutputDir 'ViewerPETests.exe') `
            -WorkingDirectory $buildOutputDir

        foreach ($viewerCase in @(
                'TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly',
                'TestViewerTextGotoPromptUsesDxUiHostAndClosesCleanly'
            )) {
            $plan += New-RSTestRunPlanEntry `
                -Name "ViewerPETests.$viewerCase" `
                -Kind 'Executable' `
                -Path (Join-Path $buildOutputDir 'ViewerPETests.exe') `
                -Arguments @($viewerCase) `
                -WorkingDirectory $buildOutputDir
        }

        foreach ($exeName in @('ViewerSqliteTests', 'MonitorTest', 'LocalizationTests', 'RedConfigureTests', 'PluginContractTests', 'SettingsSchemaTests', 'CrashHandlingTests')) {
            $plan += New-RSTestRunPlanEntry `
                -Name $exeName `
                -Kind 'Executable' `
                -Path (Join-Path $buildOutputDir "$exeName.exe") `
                -WorkingDirectory $buildOutputDir
        }

        $plan += New-RSTestRunPlanEntry `
            -Name 'PerformanceTests2' `
            -Kind 'CppUnitTest' `
            -Path (Join-Path $buildOutputDir 'PerformanceTests2.dll') `
            -WorkingDirectory $buildOutputDir

        $plan += New-RSTestRunPlanEntry `
            -Name 'ToolsPesterTests' `
            -Kind 'Pester' `
            -Path (Join-Path $RepoRoot 'Tools\Tests') `
            -Arguments @('-ExcludeTag', 'RequiresBuildToolchain') `
            -WorkingDirectory $RepoRoot

        $plan += New-RSTestRunPlanEntry `
            -Name 'VcpkgMergeSynthetic' `
            -Kind 'PowerShellScript' `
            -Path (Join-Path $RepoRoot 'Tests\vcpkg-merge-synthetic-test.ps1') `
            -WorkingDirectory $RepoRoot
    }

    # Suite CI is the GitHub Actions PR gate. Suite Full remains the broader local/closeout gate
    # and additionally includes diagnostics such as RedSalamanderMonitorEtwLatency.
    if ($Suite -eq 'Full') {
        foreach ($exeName in @('DxUiTests', 'FileSystemCurlTests', 'ViewerPETests', 'ViewerSqliteTests', 'MonitorTest', 'LocalizationTests', 'RedConfigureTests', 'PluginContractTests', 'SettingsSchemaTests', 'CrashHandlingTests')) {
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
                '--wait-instance',
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
