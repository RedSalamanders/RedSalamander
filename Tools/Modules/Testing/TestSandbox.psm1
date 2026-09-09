Set-StrictMode -Version Latest

Import-Module (Join-Path $PSScriptRoot 'TestRunPlan.Common.psm1') -Force -Scope Local

$script:RSTestSandboxMarkerFileName = '.red-salamander-test-root'
$script:RSTestSandboxMarkerContents = 'red-salamander.test-sandbox-root.v1'
$script:RSTestSandboxDirectoryName = 'RedSalamander.Perf'

function Get-RSTestSandboxRepositoryRoot {
    param([string]$RepoRoot = '')

    if (-not [string]::IsNullOrWhiteSpace($RepoRoot)) {
        return (ConvertTo-RSFullPath -Path $RepoRoot)
    }

    return (ConvertTo-RSFullPath -Path (Join-Path $PSScriptRoot '..\..\..'))
}

function Test-RSPathSameOrDescendant {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Parent
    )

    $candidate = (ConvertTo-RSFullPath -Path $Path).TrimEnd('\')
    $container = (ConvertTo-RSFullPath -Path $Parent).TrimEnd('\')
    return [string]::Equals($candidate, $container, [System.StringComparison]::OrdinalIgnoreCase) -or
        $candidate.StartsWith("$container\", [System.StringComparison]::OrdinalIgnoreCase)
}

function Assert-RSTestSandboxPathIsReparseFree {
    param([Parameter(Mandatory = $true)][string]$Path)

    $candidate = ConvertTo-RSFullPath -Path $Path
    $current = $candidate
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Refusing TestSandbox path with reparse-point component '$($item.FullName)': $candidate"
            }
        }

        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or
            [string]::Equals($parent, $current, [System.StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        $current = $parent
    }
}

function Test-RSTestSandboxMarker {
    param([Parameter(Mandatory = $true)][string]$TestRoot)

    $markerPath = Join-Path (ConvertTo-RSFullPath -Path $TestRoot) $script:RSTestSandboxMarkerFileName
    if (-not (Test-Path -LiteralPath $markerPath -PathType Leaf)) {
        return $false
    }
    $marker = Get-Item -LiteralPath $markerPath -Force -ErrorAction Stop
    if (($marker.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        return $false
    }
    return [string]::Equals(
        (Get-Content -LiteralPath $markerPath -Raw -ErrorAction Stop).TrimEnd("`r", "`n"),
        $script:RSTestSandboxMarkerContents,
        [System.StringComparison]::Ordinal)
}

function Get-RSDedicatedExternalTestSandboxRoot {
    param([Parameter(Mandatory = $true)][string]$Path)

    $candidate = ConvertTo-RSFullPath -Path $Path
    $root = [System.IO.Path]::GetPathRoot($candidate)
    if ([string]::IsNullOrWhiteSpace($root) -or $root.StartsWith('\\')) {
        return $null
    }

    try {
        $drive = [System.IO.DriveInfo]::new($root)
        if (-not $drive.IsReady -or
            $drive.DriveType -ne [System.IO.DriveType]::Fixed) {
            return $null
        }
    } catch {
        return $null
    }

    $externalRoot = ConvertTo-RSFullPath -Path (Join-Path $root $script:RSTestSandboxDirectoryName)
    if (-not [string]::Equals($candidate.TrimEnd('\'), $externalRoot.TrimEnd('\'), [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }
    return $externalRoot
}

function Assert-RSTestSandboxRootAuthorized {
    param(
        [Parameter(Mandatory = $true)][string]$TestRoot,
        [string]$RepoRoot = '',
        [switch]$AllowExternalInitialization,
        [switch]$RequireMarker
    )

    [void](Get-RSTestSandboxRepositoryRoot -RepoRoot $RepoRoot)
    $candidate = ConvertTo-RSFullPath -Path $TestRoot
    $candidateRoot = [System.IO.Path]::GetPathRoot($candidate)
    if ($candidate.Substring($candidateRoot.Length).Contains(':')) {
        throw "Unauthorized TestSandbox root with alternate-data-stream syntax: $candidate"
    }
    Assert-RSTestSandboxPathIsReparseFree -Path $candidate

    $externalRoot = Get-RSDedicatedExternalTestSandboxRoot -Path $candidate
    if ($null -eq $externalRoot) {
        throw "Unauthorized TestSandbox root '$candidate'. Use exact X:\$script:RSTestSandboxDirectoryName on a fixed local drive."
    }
    if (-not (Test-RSTestSandboxMarker -TestRoot $externalRoot) -and -not $AllowExternalInitialization) {
        throw "TestSandbox root '$externalRoot' is not authorized. Initialize it explicitly with -AllowExternalTestRoot."
    }
    if ($RequireMarker -and -not (Test-RSTestSandboxMarker -TestRoot $externalRoot)) {
        throw "Refusing to mutate unowned TestSandbox root without '$script:RSTestSandboxMarkerFileName': $externalRoot"
    }
    return $externalRoot
}

function Initialize-RSTestSandboxRoot {
    param(
        [Parameter(Mandatory = $true)][string]$TestRoot,
        [string]$RepoRoot = '',
        [switch]$AllowExternalInitialization
    )

    $root = Assert-RSTestSandboxRootAuthorized `
        -TestRoot $TestRoot `
        -RepoRoot $RepoRoot `
        -AllowExternalInitialization:$AllowExternalInitialization
    $markerPath = Join-Path $root $script:RSTestSandboxMarkerFileName
    if (-not (Test-RSTestSandboxMarker -TestRoot $root) -and (Test-Path -LiteralPath $root)) {
        $existingChildren = @(Get-ChildItem -LiteralPath $root -Force -ErrorAction Stop)
        if ($existingChildren.Count -ne 0) {
            throw "Refusing to adopt non-empty TestSandbox root without a valid ownership marker: $root"
        }
    }

    New-Item -ItemType Directory -Path $root -Force -ErrorAction Stop | Out-Null
    Assert-RSTestSandboxPathIsReparseFree -Path $root
    if (Test-Path -LiteralPath $markerPath) {
        if (-not (Test-RSTestSandboxMarker -TestRoot $root)) {
            throw "TestSandbox ownership marker is invalid: $markerPath"
        }
    } else {
        Set-Content -LiteralPath $markerPath -Value $script:RSTestSandboxMarkerContents -NoNewline -Encoding ascii -ErrorAction Stop
    }
    return $root
}

function Assert-RSTestSandboxContainedPath {
    param(
        [Parameter(Mandatory = $true)][string]$TestRoot,
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$RequireDescendant
    )

    $root = Assert-RSTestSandboxRootAuthorized -TestRoot $TestRoot -RequireMarker
    $candidate = ConvertTo-RSFullPath -Path $Path
    $candidateRoot = [System.IO.Path]::GetPathRoot($candidate)
    if ($candidate.Substring($candidateRoot.Length).Contains(':')) {
        throw "Unauthorized TestSandbox path with alternate-data-stream syntax: $candidate"
    }
    if (-not (Test-RSPathSameOrDescendant -Path $candidate -Parent $root) -or
        ($RequireDescendant -and [string]::Equals(
                $candidate.TrimEnd('\'),
                $root.TrimEnd('\'),
                [System.StringComparison]::OrdinalIgnoreCase))) {
        throw "TestSandbox path must remain below '$root': $candidate"
    }
    Assert-RSTestSandboxPathIsReparseFree -Path $candidate
    return $candidate
}

function Get-RSTestSandboxRoot {
    param(
        [string]$RepoRoot = '',

        [string]$TestRootOverride = $env:REDSALAMANDER_TEST_ROOT,

        [string]$SelfTestRootOverride = $env:REDSALAMANDER_SELFTEST_ROOT,

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA,

        [switch]$AllowExternalInitialization
    )

    $repository = Get-RSTestSandboxRepositoryRoot -RepoRoot $RepoRoot

    if (-not [string]::IsNullOrWhiteSpace($TestRootOverride)) {
        $testRoot = Assert-RSTestSandboxRootAuthorized `
            -TestRoot $TestRootOverride `
            -RepoRoot $repository `
            -AllowExternalInitialization:$AllowExternalInitialization
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
        return (Assert-RSTestSandboxRootAuthorized `
                -TestRoot $SelfTestRootOverride `
                -RepoRoot $repository)
    }

    $repositoryDriveRoot = [System.IO.Path]::GetPathRoot($repository)
    if ([string]::IsNullOrWhiteSpace($repositoryDriveRoot) -or $repositoryDriveRoot.StartsWith('\\')) {
        throw "Cannot derive X:\$script:RSTestSandboxDirectoryName from repository root '$repository'. Set REDSALAMANDER_TEST_ROOT explicitly."
    }
    return (Assert-RSTestSandboxRootAuthorized `
            -TestRoot (Join-Path $repositoryDriveRoot $script:RSTestSandboxDirectoryName) `
            -RepoRoot $repository `
            -AllowExternalInitialization:$AllowExternalInitialization)
}

function New-RSTestRunContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [string]$RunId = '',

        [string]$TestRootOverride = $env:REDSALAMANDER_TEST_ROOT,

        [string]$SelfTestRootOverride = $env:REDSALAMANDER_SELFTEST_ROOT,

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA,

        [switch]$AllowExternalInitialization
    )

    $testRoot = Get-RSTestSandboxRoot `
        -RepoRoot $RepoRoot `
        -TestRootOverride $TestRootOverride `
        -SelfTestRootOverride $SelfTestRootOverride `
        -LocalAppDataRoot $LocalAppDataRoot `
        -AllowExternalInitialization:$AllowExternalInitialization
    if ([string]::IsNullOrWhiteSpace($RunId)) {
        $RunId = New-RSTestRunId
    }
    $RunId = Assert-RSTestSandboxPathSegment -Value $RunId -Description 'TestSandbox run id'

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

    if ($Value -eq '.' -or $Value -eq '..' -or $Value.EndsWith('.') -or $Value -match '[\\/:\*\?"<>\|]') {
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
    [void](Initialize-RSTestSandboxRoot -TestRoot $context.TestRoot -RepoRoot $RepoRoot)
    $runRoot = $context.RunRoot
    $scratchRoot = Join-Path $runRoot 'scratch'
    $harnessRoot = Join-Path $scratchRoot $safeHarness
    $caseRoot = Join-Path $harnessRoot $safeCase

    New-Item -ItemType Directory -Path $caseRoot -Force | Out-Null
    return (ConvertTo-RSFullPath -Path $caseRoot)
}

function Get-RSToolingPesterSandboxRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$PrimaryTestRoot
    )

    $primary = (ConvertTo-RSFullPath -Path $PrimaryTestRoot).TrimEnd('\')
    [void](Assert-RSTestSandboxRootAuthorized -TestRoot $primary -RepoRoot $RepoRoot -RequireMarker)
    return $primary
}

function Remove-RSTestSandboxExactRunDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestRoot,

        [Parameter(Mandatory = $true)]
        [string]$RunId
    )

    $safeRunId = Assert-RSTestSandboxPathSegment `
        -Value $RunId `
        -Description 'TestSandbox run id'
    $root = ConvertTo-RSFullPath -Path $TestRoot
    [void](Assert-RSTestSandboxRootAuthorized -TestRoot $root -RequireMarker)
    $runsRoot = ConvertTo-RSFullPath -Path (Join-Path $root 'runs')
    $runRoot = ConvertTo-RSFullPath -Path (Join-Path $runsRoot $safeRunId)
    if (-not $runRoot.StartsWith(
            "$runsRoot\",
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove TestSandbox run outside '$runsRoot': $runRoot"
    }
    if (-not (Test-Path -LiteralPath $runRoot)) {
        return
    }
    if (((Get-Item -LiteralPath $runRoot -Force).Attributes -band
            [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Refusing to remove reparse-point TestSandbox run root: $runRoot"
    }

    Remove-Item -LiteralPath $runRoot -Recurse -Force -ErrorAction Stop
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
    [void](Assert-RSTestSandboxRootAuthorized -TestRoot $normalizedTestRoot -RequireMarker)
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
    [void](Assert-RSTestSandboxRootAuthorized -TestRoot $normalizedTestRoot -RequireMarker)
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

        [int64[]]$LiveProcessIds = @(),

        [hashtable]$LiveProcessStartTimesUtc = @{}
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
        $resolveParams = @{
            TestRoot = $normalizedTestRoot
            RunId = $RunId
            AllowedRunIds = @($AllowedRunIds)
        }
        if ($PSBoundParameters.ContainsKey('LiveProcessIds')) {
            $resolveParams['LiveProcessIds'] = @($LiveProcessIds)
            $resolveParams['LiveProcessStartTimesUtc'] = $LiveProcessStartTimesUtc
        }
        $staleRunTargets = @{}
        foreach ($target in @(Resolve-RSTestSandboxStaleRunTargets @resolveParams)) {
            $staleRunTargets[$target.RunId] = $target
        }

        foreach ($item in @(Get-ChildItem -LiteralPath $normalizedTestRoot -Force -ErrorAction SilentlyContinue)) {
            if ($item.Name -notin @('runs', 'config', 'evidence', $script:RSTestSandboxMarkerFileName)) {
                $issues += New-RSTestSandboxDiskAuditIssue `
                    -Category 'unexpected-test-root-child' `
                    -Path $item.FullName `
                    -Reason 'Unified TestSandbox root should contain only runner-owned runs, config, and evidence.'
            }
        }

        if (Test-Path -LiteralPath $runsRoot) {
            foreach ($runDir in @(Get-ChildItem -LiteralPath $runsRoot -Directory -Force -ErrorAction SilentlyContinue)) {
                $isAllowedRun = @($allowedRunIds | Where-Object {
                        [string]::Equals($_, $runDir.Name, [System.StringComparison]::OrdinalIgnoreCase)
                    }).Count -gt 0
                if ($isAllowedRun) {
                    continue
                }

                if ($staleRunTargets.ContainsKey($runDir.Name)) {
                    $target = $staleRunTargets[$runDir.Name]
                    $issues += New-RSTestSandboxDiskAuditIssue `
                        -Category $target.Category `
                        -Path $runDir.FullName `
                        -Reason $target.Reason
                    continue
                }

                if ($null -eq (ConvertFrom-RSTestRunId -RunId $runDir.Name)) {
                    $issues += New-RSTestSandboxDiskAuditIssue `
                        -Category 'unexpected-test-run-dir' `
                        -Path $runDir.FullName `
                        -Reason 'The run directory is neither explicitly allowed nor attributable to a protected live owner.'
                }
            }
        }
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

    # Retain the legacy parameters only for source compatibility. They are never
    # authorization to write outside the exact marked X:\RedSalamander.Perf root.
    [void]$SelfTestRootOverride
    [void]$LocalAppDataRoot
    $repository = Get-RSTestSandboxRepositoryRoot -RepoRoot $RepoRoot
    return (New-RSTestRunContext -RepoRoot $repository -RunId $RunId).ArtifactRoot
}

Export-ModuleMember -Function @(
    'Initialize-RSTestSandboxRoot',
    'Assert-RSTestSandboxContainedPath',
    'Get-RSTestSandboxRoot',
    'New-RSTestRunContext',
    'New-RSTestSandboxScratchDirectory',
    'Get-RSToolingPesterSandboxRoot',
    'Remove-RSTestSandboxExactRunDirectory',
    'Resolve-RSTestSandboxStaleRunTargets',
    'Remove-RSTestSandboxStaleRunDirectories',
    'Get-RSTestSandboxDiskAudit',
    'Get-RSSelfTestArtifactRoot'
)
