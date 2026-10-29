Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRunPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
$archiveModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunArchive.psm1'
$archiveValidator = Join-Path $repoRoot 'Tools\Test-TestRunArchive.ps1'
Import-Module $testRunPlanModule -Force -ErrorAction Stop
Import-Module $archiveModule -Force -ErrorAction Stop

function New-RSArchiveContractRoot {
    return (New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case ("test-run-archive-" + [guid]::NewGuid().ToString('N')))
}

function New-RSArchiveMigrationFixture {
    $root = New-RSArchiveContractRoot
    & git -C $root init --quiet --initial-branch=main
    & git -C $root config core.longpaths true
    & git -C $root config user.email 'archive-contract@example.invalid'
    & git -C $root config user.name 'Archive Contract'
    $legacy = 'Specs/TestRuns/7d3a1247382a/FileOps/2026-07-01_120000/legacy.raw'
    $new = 'Specs/TestRuns/7d3a1247382a/FileOps/2026-07-20_120000/new.raw'
    foreach ($path in @($legacy, $new)) {
        New-Item -ItemType Directory -Path (Split-Path -Parent (Join-Path $root $path)) -Force | Out-Null
    }
    [IO.File]::WriteAllBytes((Join-Path $root $legacy), [byte[]]::new(2049))
    & git -C $root add -- $legacy
    & git -C $root commit --quiet -m 'immutable pre-contract fixture'
    if ($LASTEXITCODE -ne 0) { throw 'Failed to create the legacy archive baseline fixture.' }
    $baseline = (& git -C $root rev-parse HEAD).Trim()
    & git -C $root update-ref refs/remotes/origin/main $baseline
    & git -C $root switch --quiet -c ej/archive-contract
    [IO.File]::WriteAllBytes((Join-Path $root $new), [byte[]]::new(2049))
    & git -C $root add -- $new
    & git -C $root commit --quiet -m 'new governed invalid fixture'
    if ($LASTEXITCODE -ne 0) { throw 'Failed to create the post-contract archive fixture.' }
    return [pscustomobject]@{ Root = $root; Baseline = $baseline; Legacy = $legacy; New = $new }
}

function Get-RSArchiveMigrationFixtureInventory {
    param([string]$Root, [string]$Baseline)
    return @(& (Get-Module TestRunArchive) {
        param($FixtureRoot, $FixtureBaseline)
        $previous = $script:RSTestRunArchiveImportedBaselineCommit
        try {
            $script:RSTestRunArchiveImportedBaselineCommit = $FixtureBaseline
            Get-RSTestRunArchiveInventoryPaths -RepoRoot $FixtureRoot
        } finally { $script:RSTestRunArchiveImportedBaselineCommit = $previous }
    } $Root $Baseline)
}

function Write-RSMetricFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [string]$Profile,

        [Parameter(Mandatory = $true)]
        [string]$MachineHash
    )

    $relative = "Specs\TestRuns\$Profile\Commands\2026-07-11_120000\perf\perf_metrics.jsonl"
    $path = Join-Path $Root $relative
    New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
    ([pscustomobject]@{ metric = 'FileOps.Test'; machineHash = $MachineHash; value = 1; unit = 'count' } | ConvertTo-Json -Compress) |
        Set-Content -LiteralPath $path -Encoding UTF8
    return $relative
}

function New-RSTerminalVtArchiveFixture {
    $root = New-RSArchiveContractRoot
    $schemaDestination = Join-Path $root 'Specs\Terminal\TerminalVtUpgradeEvidence.schema.json'
    New-Item -ItemType Directory -Path (Split-Path -Parent $schemaDestination) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $repoRoot 'Specs\Terminal\TerminalVtUpgradeEvidence.schema.json') `
        -Destination $schemaDestination

    # Synthetic evidence keeps contract tests executable in fresh/shallow clones;
    # historical local benchmark archives are never prerequisites for unit tests.
    $baselineRelative = 'Specs/TestRuns/0123456789ab/Terminal/20261028_120000_baseline'
    $candidateRelative = 'Specs/TestRuns/0123456789ab/Terminal/20261028_120001_candidate'
    foreach ($relative in @($baselineRelative, $candidateRelative)) {
        $destination = Join-Path $root $relative
        New-Item -ItemType Directory -Path (Join-Path $destination 'perf') -Force | Out-Null
        $isBaseline = $relative -eq $baselineRelative
        $runId = Split-Path $relative -Leaf
        $results = [ordered]@{
            schema = 'red-salamander.terminal-vt-upgrade-evidence.v1'
            runId = $runId; createdUtc = '2026-10-28T12:00:00Z'
            repositoryCommit = 'a' * 40; repositoryBranch = 'ej/synthetic-vt-fixture'
            machineHash = '0123456789ab'; configuration = 'Release'
            buildReceiptId = $(if ($isBaseline) { 'b' * 64 } else { 'c' * 64 })
            runtimeIdentitySha256 = $(if ($isBaseline) { 'd' * 64 } else { 'e' * 64 })
            environment = [ordered]@{
                windowsBuild='26100'; cpu='Synthetic CPU'; architecture='AMD64'
                compiler='MSVC synthetic'; windowsSdk='10.0.26100.0'
                zig='0.16.0'; vcpkgCommit='f' * 40
            }
            corpus = [ordered]@{
                formatId='red-salamander-terminal-vt-upgrade-corpus'; version=1
                sha256='1' * 64; inputByteCount=1234
            }
            behavior = [ordered]@{
                ptyOutputSha256='2' * 64; snapshotSha256='3' * 64
                withoutExpectedDeltas = @{ ptyOutputSha256='2' * 64; snapshotSha256='3' * 64 }
                expectedDeltas = @(@{ id='dcs-high-byte-preservation'; disposition='candidate-only-if-reviewed' })
            }
            measurements = [ordered]@{ sampleCount=200; percentileRule='nearest-rank' }
            kitty = @{ placementCount=1 }; memory = @{ peakPrivateBytes=1024 }
            assertionCount=1; outcome='pass'
        }
        $rows = [Collections.Generic.List[string]]::new()
        foreach ($entry in @(
                @{ name='vtWrite'; metric='vt_write_us'; offset=$(if($isBaseline){100}else{50}) },
                @{ name='snapshot'; metric='snapshot_us'; offset=$(if($isBaseline){2000}else{1980}) },
                @{ name='hostedRender'; metric='hosted_render_us'; offset=$(if($isBaseline){7194}else{7000}) })) {
            $samples = @(1..200 | ForEach-Object { [long]($entry.offset + $_) })
            $results.measurements[$entry.name] = @{
                unit='us'; p50=$entry.offset + 100; p95=$entry.offset + 190
                candidateMaximumP95=[long][Math]::Ceiling([decimal]($entry.offset + 190) * [decimal]'1.10')
                samples=$samples
            }
            foreach ($sample in $samples) {
                $rows.Add((@{
                            metric="terminal.vt_upgrade.$($entry.metric)"; value=$sample; unit='us'
                            machineHash=$results.machineHash; runId=$runId; build='Release'
                            commit=$results.repositoryCommit; scenario='terminal-vt-upgrade-corpus-v1'
                        } | ConvertTo-Json -Compress))
            }
        }
        $results | ConvertTo-Json -Depth 15 | Set-Content (Join-Path $destination 'results.json') -Encoding UTF8
        [IO.File]::WriteAllLines((Join-Path $destination 'perf/perf_metrics.jsonl'), $rows)
        'Synthetic Terminal VT archive fixture' | Set-Content (Join-Path $destination 'trace.txt') -Encoding UTF8
    }

    return [pscustomobject]@{
        Root = $root
        Baseline = $baselineRelative
        Candidate = $candidateRelative
    }
}

function Invoke-RSArchiveValidatorProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [string]$RunPath,

        [switch]$Inventory,

        [long]$MaxFileBytes = 1KB,

        [long]$MaxRunBytes = 2KB
    )

    $stdoutPath = Join-Path $Root 'validator.stdout.txt'
    $stderrPath = Join-Path $Root 'validator.stderr.txt'
    $arguments = [System.Collections.Generic.List[string]]::new()
    $arguments.Add('-NoProfile')
    $arguments.Add('-ExecutionPolicy')
    $arguments.Add('Bypass')
    $arguments.Add('-File')
    $arguments.Add($archiveValidator)
    $arguments.Add('-RepoRoot')
    $arguments.Add($Root)
    if ($Inventory) {
        $arguments.Add('-Inventory')
    } else {
        $arguments.Add('-RunPath')
        $arguments.Add($RunPath)
    }
    $arguments.Add('-MaxFileBytes')
    $arguments.Add([string]$MaxFileBytes)
    $arguments.Add('-MaxRunBytes')
    $arguments.Add([string]$MaxRunBytes)

    $process = Start-Process `
        -FilePath (Get-Process -Id $PID).Path `
        -WindowStyle Hidden `
        -ArgumentList $arguments `
        -Wait `
        -PassThru `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath

    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        Output = ((Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue) +
            (Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue))
    }
}

Describe 'TestRuns archive contract' {
    # Specs/TestRuns is gitignored evidence: the archives these contracts inspect exist only on the
    # machines that produced them. The inventory is governed from the archive contract commit, which a
    # squashed or shallow clone without either immutable ancestor does not carry; there
    # Get-RSTestRunArchiveInventoryPaths returns every tracked file, grandfathered evidence included.
    # Skip those cases there instead of failing every clone (Testing_SelfTests.md).
    $fileOpsSummaryArchivePresent = Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\TestRuns\7d3a1247382a\FileOps\2026-07-20_205550') -PathType Container
    $archiveContractCommit = & (Get-Module TestRunArchive) { $script:RSTestRunArchiveContractCommit }
    & git -C $repoRoot merge-base --is-ancestor $archiveContractCommit HEAD *> $null
    $archiveContractInHistory = ($LASTEXITCODE -eq 0)
    $importedBaseline = & (Get-Module TestRunArchive) { $script:RSTestRunArchiveImportedBaselineCommit }
    & git -C $repoRoot merge-base --is-ancestor $importedBaseline HEAD *> $null
    $importedBaselineInHistory = ($LASTEXITCODE -eq 0)
    $hasInventoryBaseline = $archiveContractInHistory -or $importedBaselineInHistory
    $checkedInInventoryPresent = $hasInventoryBaseline -and @(Get-RSTestRunArchiveInventoryPaths -RepoRoot $repoRoot).Count -gt 0
    if (-not $fileOpsSummaryArchivePresent) {
        Write-Host 'Skipping the FileOps summary archive contract: Specs\TestRuns\7d3a1247382a\FileOps\2026-07-20_205550 is not present (gitignored archive).'
    }
    if (-not $hasInventoryBaseline) {
        Write-Host 'Skipping the real archive inventory contract: neither immutable baseline is in this clone; explicit Inventory still validates all files.'
    } elseif (-not $checkedInInventoryPresent) {
        Write-Host 'Skipping the archive inventory contract: no Specs\TestRuns archive is present (gitignored archive).'
    }

    It 'accepts a bounded artifact under its embedded machine profile' {
        $root = New-RSArchiveContractRoot
        $relative = Write-RSMetricFile -Root $root -Profile '7d3a1247382a' -MachineHash '7d3a1247382a'

        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $root -Paths @($relative) -MaxFileBytes 1KB -MaxRunBytes 2KB)
        $violations.Count | Should Be 0
    }

    It 'rejects a commit hash used as the machine profile' {
        $root = New-RSArchiveContractRoot
        $relative = Write-RSMetricFile -Root $root -Profile '00013ba62' -MachineHash '7d3a1247382a'

        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $root -Paths @($relative) -MaxFileBytes 1KB -MaxRunBytes 2KB)
        @($violations | Where-Object Kind -eq 'MachineProfile').Count | Should Be 1
    }

    It 'rejects an oversized individual artifact' {
        $root = New-RSArchiveContractRoot
        $relative = 'Specs\TestRuns\7d3a1247382a\Commands\2026-07-11_120000\trace.txt'
        $path = Join-Path $root $relative
        New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
        [IO.File]::WriteAllBytes($path, [byte[]]::new(2049))

        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $root -Paths @($relative) -MaxFileBytes 2KB -MaxRunBytes 4KB)
        @($violations | Where-Object Kind -eq 'FileSize').Count | Should Be 1
    }

    It 'rejects an oversized aggregate run including unchanged sibling files' {
        $root = New-RSArchiveContractRoot
        $paths = @(
            'Specs\TestRuns\7d3a1247382a\Commands\2026-07-11_120000\results.json',
            'Specs\TestRuns\7d3a1247382a\Commands\2026-07-11_120000\trace.txt'
        )
        foreach ($relative in $paths) {
            $path = Join-Path $root $relative
            New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
            [IO.File]::WriteAllBytes($path, [byte[]]::new(1536))
        }

        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $root -Paths @($paths[0]) -MaxFileBytes 2KB -MaxRunBytes 2KB)
        @($violations | Where-Object Kind -eq 'RunSize').Count | Should Be 1
    }

    It 'sizes a real worktree run from tracked files and excludes ignored raw artifacts' {
        $root = New-RSArchiveContractRoot
        & git -C $root init --quiet
        if ($LASTEXITCODE -ne 0) {
            throw 'Failed to initialize the synthetic archive worktree.'
        }
        & git -C $root config core.longpaths true
        if ($LASTEXITCODE -ne 0) {
            throw 'Failed to enable long-path support for the synthetic archive worktree.'
        }

        '.raw' | Set-Content -LiteralPath (Join-Path $root '.gitignore') -Encoding UTF8
        $relative = 'Specs\TestRuns\7d3a1247382a\Continuation\2026-08-03_contract\README.md'
        $trackedPath = Join-Path $root $relative
        $ignoredPath = Join-Path (Split-Path -Parent $trackedPath) 'trace.raw'
        New-Item -ItemType Directory -Path (Split-Path -Parent $trackedPath) -Force | Out-Null
        [IO.File]::WriteAllBytes($trackedPath, [byte[]]::new(512))
        [IO.File]::WriteAllBytes($ignoredPath, [byte[]]::new(4096))
        & git -C $root add -- .gitignore ($relative -replace '\\', '/')
        if ($LASTEXITCODE -ne 0) {
            throw 'Failed to stage the synthetic checked-in archive artifact.'
        }

        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $root -Paths @($relative) -MaxFileBytes 1KB -MaxRunBytes 1KB)
        $violations.Count | Should Be 0
    }

    It 'accepts an empty path set' {
        $root = New-RSArchiveContractRoot
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $root -Paths @())
        $violations.Count | Should Be 0
    }

    It 'validates every Terminal VT JSONL row and requires its schema-owned artifacts' {
        $fixture = New-RSTerminalVtArchiveFixture
        $candidateRoot = Join-Path $fixture.Root $fixture.Candidate
        Add-Content -LiteralPath (Join-Path $candidateRoot 'perf\perf_metrics.jsonl') -Value '{not-json'
        Remove-Item -LiteralPath (Join-Path $candidateRoot 'trace.txt') -Force

        $paths = @(Get-RSTestRunArchivePathsForRuns -RepoRoot $fixture.Root -RunPath @($fixture.Candidate))
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $paths)
        @($violations | Where-Object Kind -eq 'PerfJson').Count | Should BeGreaterThan 0
        @($violations | Where-Object Kind -eq 'RequiredArtifact').Count | Should Be 1
    }

    It 'rejects Terminal VT results that do not satisfy the area schema' {
        $fixture = New-RSTerminalVtArchiveFixture
        $resultsPath = Join-Path (Join-Path $fixture.Root $fixture.Candidate) 'results.json'
        $results = Get-Content -LiteralPath $resultsPath -Raw | ConvertFrom-Json
        $results.schema = 'red-salamander.terminal-vt-upgrade-evidence.invalid'
        $results | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $resultsPath -Encoding UTF8

        $paths = @(Get-RSTestRunArchivePathsForRuns -RepoRoot $fixture.Root -RunPath @($fixture.Candidate))
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $paths)
        @($violations | Where-Object Kind -in @('EvidenceSchema', 'EvidenceJson')).Count | Should BeGreaterThan 0
    }

    It 'recomputes a Terminal VT pair and rejects a candidate beyond baseline x1.10' {
        $fixture = New-RSTerminalVtArchiveFixture
        $passing = Compare-RSTerminalVtUpgradeEvidence `
            -RepoRoot $fixture.Root `
            -BaselineRunPath $fixture.Baseline `
            -CandidateRunPath $fixture.Candidate
        $passing.Outcome | Should Be 'pass'
        @($passing.Measurements).Count | Should Be 3

        $resultsPath = Join-Path (Join-Path $fixture.Root $fixture.Candidate) 'results.json'
        $results = Get-Content -LiteralPath $resultsPath -Raw | ConvertFrom-Json
        $results.measurements.hostedRender.samples = @(9000) * 200
        $results.measurements.hostedRender.p50 = 9000
        $results.measurements.hostedRender.p95 = 9000
        $results.measurements.hostedRender.candidateMaximumP95 = 9900
        $results | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $resultsPath -Encoding UTF8

        $failing = Compare-RSTerminalVtUpgradeEvidence `
            -RepoRoot $fixture.Root `
            -BaselineRunPath $fixture.Baseline `
            -CandidateRunPath $fixture.Candidate
        $failing.Outcome | Should Be 'fail'
        ($failing.Violations.Message -join "`n") | Should Match 'candidate p95 9000 exceeds baseline ceiling 8123'
    }

    It 'accepts only the exact reviewed Zig transition and otherwise requires equal versions' {
        $fixture = New-RSTerminalVtArchiveFixture
        $paths = @(Get-RSTestRunArchivePathsForRuns -RepoRoot $fixture.Root -RunPath @($fixture.Baseline, $fixture.Candidate))
        @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $paths).Count | Should Be 0
        $resultsPath = Join-Path (Join-Path $fixture.Root $fixture.Candidate) 'results.json'
        $results = Get-Content $resultsPath -Raw | ConvertFrom-Json
        $results.environment.zig = '0.17.0'
        $results | ConvertTo-Json -Depth 15 | Set-Content $resultsPath -Encoding UTF8
        $arguments = @{ RepoRoot=$fixture.Root; BaselineRunPath=$fixture.Baseline; CandidateRunPath=$fixture.Candidate }
        $strict = Compare-RSTerminalVtUpgradeEvidence @arguments
        $strict.Outcome | Should Be 'fail'
        ($strict.Violations.Message -join " ") | Should Match "Environment identity 'zig'"
        $reviewed = Compare-RSTerminalVtUpgradeEvidence @arguments -ZigToolchainTransition @('0.16.0','0.17.0')
        $reviewed.Outcome | Should Be 'pass'
        ($reviewed.ZigToolchainTransition -join ',') | Should Be '0.16.0,0.17.0'
        $mismatch = Compare-RSTerminalVtUpgradeEvidence @arguments -ZigToolchainTransition @('0.16.0','0.18.0')
        $mismatch.Outcome | Should Be 'fail'
        ($mismatch.Violations.Message -join " ") | Should Match 'exact reviewed'
        { Compare-RSTerminalVtUpgradeEvidence @arguments -ZigToolchainTransition @('0.17.0') } | Should Throw 'exactly two'
        { Compare-RSTerminalVtUpgradeEvidence @arguments -ZigToolchainTransition @('0.17.0','0.17.0') } | Should Throw 'distinct'
    }

    It 'retains environment, behavior, and p95 rejection during a reviewed Zig transition' {
        $fixture = New-RSTerminalVtArchiveFixture
        $resultsPath = Join-Path (Join-Path $fixture.Root $fixture.Candidate) 'results.json'
        $results = Get-Content $resultsPath -Raw | ConvertFrom-Json
        $results.environment.zig = '0.17.0'
        $results.environment.cpu = 'Different synthetic CPU'
        $results.behavior.withoutExpectedDeltas.snapshotSha256 = '9' * 64
        $results.measurements.hostedRender.samples = @(9000) * 200
        $results.measurements.hostedRender.p50 = 9000
        $results.measurements.hostedRender.p95 = 9000
        $results.measurements.hostedRender.candidateMaximumP95 = 9900
        $results | ConvertTo-Json -Depth 15 | Set-Content $resultsPath -Encoding UTF8
        $comparison = Compare-RSTerminalVtUpgradeEvidence -RepoRoot $fixture.Root -BaselineRunPath $fixture.Baseline `
            -CandidateRunPath $fixture.Candidate -ZigToolchainTransition @('0.16.0','0.17.0')
        $comparison.Outcome | Should Be 'fail'
        $messages = $comparison.Violations.Message -join " "
        $messages | Should Match "Environment identity 'cpu'"
        $messages | Should Match 'Control behavior digests differ'
        $messages | Should Match 'candidate p95 9000 exceeds baseline ceiling 8123'
    }

    It 'finds an unchanged historical violation only in whole-inventory mode' {
        $root = New-RSArchiveContractRoot
        & git -c core.longpaths=true -C $root init --quiet
        if ($LASTEXITCODE -ne 0) {
            throw 'Failed to initialize the historical archive fixture.'
        }
        & git -C $root config core.longpaths true
        & git -C $root config user.email 'archive-contract@example.invalid'
        & git -C $root config user.name 'Archive Contract'
        $relative = 'Specs\TestRuns\7d3a1247382a\FileOps\2026-07-20_205550\perf\perf_metrics.jsonl'
        $path = Join-Path $root $relative
        New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
        [IO.File]::WriteAllBytes($path, [byte[]]::new(2049))
        & git -C $root add -- ($relative -replace '\\', '/')
        & git -C $root commit --quiet -m 'historical oversized archive fixture'
        if ($LASTEXITCODE -ne 0) {
            throw 'Failed to commit the historical archive fixture.'
        }

        @(Get-RSChangedTestRunArchivePaths -RepoRoot $root).Count | Should Be 0
        $result = Invoke-RSArchiveValidatorProcess -Root $root -Inventory -MaxFileBytes 2KB -MaxRunBytes 4KB
        $result.ExitCode | Should Be 1
        $result.Output | Should Match 'FileSize:.*perf_metrics\.jsonl:.*maximum is 2048 bytes'
        $result.Output | Should Not Match 'archive contract passed'
    }

    It 'retains the imported pre-contract boundary while rejecting new committed archives' {
        $fixture = New-RSArchiveMigrationFixture
        $paths = @(Get-RSArchiveMigrationFixtureInventory -Root $fixture.Root -Baseline $fixture.Baseline)
        $paths.Count | Should Be 1
        $paths[0] | Should Be $fixture.New
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $paths -MaxFileBytes 2KB -MaxRunBytes 4KB)
        @($violations | Where-Object Kind -eq 'FileSize').Count | Should Be 1
        $violations.Path | Should Match 'new\.raw'
        $explicit = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths @($fixture.Legacy) -MaxFileBytes 2KB -MaxRunBytes 4KB)
        @($explicit | Where-Object Kind -eq 'FileSize').Count | Should Be 1
    }

    It 'rejects committed and uncommitted edits to imported legacy evidence' {
        $fixture = New-RSArchiveMigrationFixture
        [IO.File]::WriteAllBytes((Join-Path $fixture.Root $fixture.Legacy), [byte[]]::new(2050))
        $dirty = @(Get-RSArchiveMigrationFixtureInventory -Root $fixture.Root -Baseline $fixture.Baseline)
        $dirty.Count | Should Be 2
        & git -C $fixture.Root add -- $fixture.Legacy
        $staged = @(Get-RSArchiveMigrationFixtureInventory -Root $fixture.Root -Baseline $fixture.Baseline)
        $staged.Count | Should Be 2
        & git -C $fixture.Root commit --quiet -m 'edited legacy fixture'
        if ($LASTEXITCODE -ne 0) { throw 'Failed to commit the edited legacy fixture.' }
        $committed = @(Get-RSArchiveMigrationFixtureInventory -Root $fixture.Root -Baseline $fixture.Baseline)
        $committed.Count | Should Be 2
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $committed -MaxFileBytes 2KB -MaxRunBytes 4KB)
        @($violations | Where-Object Kind -eq 'FileSize').Count | Should Be 2
    }

    It 'validates staged and committed moves of grandfathered blobs at their new paths' {
        $fixture = New-RSArchiveMigrationFixture
        & git -C $fixture.Root config diff.renames true
        $moved = 'Specs/TestRuns/7d3a1247382a/FileOps/2026-07-21_120000/moved.raw'
        New-Item -ItemType Directory -Path (Split-Path -Parent (Join-Path $fixture.Root $moved)) -Force | Out-Null
        & git -C $fixture.Root mv -- $fixture.Legacy $moved
        if ($LASTEXITCODE -ne 0) { throw 'Failed to stage the legacy evidence move.' }
        foreach ($phase in @('staged', 'committed')) {
            if ($phase -eq 'committed') {
                & git -C $fixture.Root commit --quiet -m 'moved legacy evidence fixture'
                if ($LASTEXITCODE -ne 0) { throw 'Failed to commit the legacy evidence move.' }
            }
            $paths = @(Get-RSArchiveMigrationFixtureInventory -Root $fixture.Root -Baseline $fixture.Baseline)
            $paths.Count | Should Be 2
            ($paths -contains $moved) | Should Be $true
            $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $paths -MaxFileBytes 2KB -MaxRunBytes 4KB)
            @($violations | Where-Object Kind -eq 'FileSize').Count | Should Be 2
            $changed = @(Get-RSChangedTestRunArchivePaths -RepoRoot $fixture.Root)
            $changed.Count | Should Be 2
            ($changed -contains $moved) | Should Be $true
        }
    }

    It 'fails closed when an imported baseline object is present but is not an ancestor' {
        $fixture = New-RSArchiveMigrationFixture
        $tree = (& git -C $fixture.Root rev-parse 'HEAD^{tree}').Trim()
        $unrelated = (& git -C $fixture.Root commit-tree $tree -m 'unrelated parentless fixture').Trim()
        if ($LASTEXITCODE -ne 0) { throw 'Failed to create the unrelated baseline object.' }
        $paths = @(Get-RSArchiveMigrationFixtureInventory -Root $fixture.Root -Baseline $unrelated)
        $paths.Count | Should Be 2
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $paths -MaxFileBytes 2KB -MaxRunBytes 4KB)
        @($violations | Where-Object Kind -eq 'FileSize').Count | Should Be 2
    }

    It 'checks committed branch evidence against origin main in a history-free import' {
        $fixture = New-RSArchiveMigrationFixture
        $paths = @(Get-RSChangedTestRunArchivePaths -RepoRoot $fixture.Root)
        $paths.Count | Should Be 1
        $paths[0] | Should Be $fixture.New
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $paths -MaxFileBytes 2KB -MaxRunBytes 4KB)
        @($violations | Where-Object Kind -eq 'FileSize').Count | Should Be 1
    }

    It 'rejects a newly produced oversized run before staging and never prints a green closeout' {
        $root = New-RSArchiveContractRoot
        $relativeRun = 'Specs\TestRuns\7d3a1247382a\FileOps\2026-08-09_120000'
        $run = Join-Path $root $relativeRun
        New-Item -ItemType Directory -Path $run -Force | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $run 'new.raw'), [byte[]]::new(2049))

        $result = Invoke-RSArchiveValidatorProcess -Root $root -RunPath $run -MaxFileBytes 2KB -MaxRunBytes 4KB
        $result.ExitCode | Should Be 1
        $result.Output | Should Match 'FileSize:.*new\.raw:.*maximum is 2048 bytes'
        $result.Output | Should Not Match 'archive contract passed'
    }

    It 'retains a digest-bound compact summary and the behavioral results supporting the archived claim' -Skip:(-not $fileOpsSummaryArchivePresent) {
        $relativeRun = 'Specs\TestRuns\7d3a1247382a\FileOps\2026-07-20_205550'
        $run = Join-Path $repoRoot $relativeRun
        $summaryPath = Join-Path $run 'perf\perf_metrics_summary.json'
        $readmePath = Join-Path $run 'README.md'
        $rawPath = Join-Path $run 'perf\perf_metrics.jsonl'

        (Test-Path -LiteralPath $rawPath) | Should Be $false
        (Test-Path -LiteralPath $summaryPath -PathType Leaf) | Should Be $true
        (Test-Path -LiteralPath $readmePath -PathType Leaf) | Should Be $true
        $summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
        $summary.compaction.originalSourceBytes | Should Be 7009000
        $summary.compaction.originalRunBytes | Should Be 7067986
        $summary.compaction.originalSourceSha256 | Should Be '6ce75a44c2d7cae22d1e5ab15fefb6aec334585e54856ce7ad33217bbcee446f'
        $summary.compaction.originalRowCount | Should Be 17190
        $summary.behavioralResult.passed | Should Be 47
        $summary.behavioralResult.failed | Should Be 0
        @($summary.representativeMetrics).Count | Should Be 6

        $readme = Get-Content -LiteralPath $readmePath -Raw
        $readme | Should Match '47 passed / 0 failed / 0 skipped'
        $readme | Should Match 'Microsoft Drive'
        ($readme -replace '\*', '') | Should Match 'does not\s+claim live Microsoft Graph coverage'

        $paths = @(Get-RSTestRunArchivePathsForRuns -RepoRoot $repoRoot -RunPath @($relativeRun))
        @(Get-RSTestRunArchiveViolations -RepoRoot $repoRoot -Paths $paths).Count | Should Be 0
    }

    It 'keeps the current changed archive set within the contract' {
        $paths = @(Get-RSChangedTestRunArchivePaths -RepoRoot $repoRoot)
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $repoRoot -Paths $paths)
        if ($violations.Count -ne 0) {
            throw (($violations | ForEach-Object { "$($_.Kind): $($_.Path): $($_.Message)" }) -join [Environment]::NewLine)
        }
    }

    It 'keeps the complete checked-in archive inventory within the contract' -Skip:(-not $checkedInInventoryPresent) {
        $paths = @(Get-RSTestRunArchiveInventoryPaths -RepoRoot $repoRoot)
        $paths.Count | Should BeGreaterThan 0
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $repoRoot -Paths $paths)
        if ($violations.Count -ne 0) {
            throw (($violations | ForEach-Object { "$($_.Kind): $($_.Path): $($_.Message)" }) -join [Environment]::NewLine)
        }
    }
}
