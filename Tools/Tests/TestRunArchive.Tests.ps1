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

    $baselineRelative = 'Specs\TestRuns\4cb089111a23\Terminal\20260826_071812_pair2-product'
    $candidateRelative = 'Specs\TestRuns\4cb089111a23\Terminal\20260826_071820_pair2-candidate'
    foreach ($relative in @($baselineRelative, $candidateRelative)) {
        $destination = Join-Path $root $relative
        New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $repoRoot $relative) -Destination $destination -Recurse
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
    # machines that produced them, and the governed inventory is empty on a clone whose history
    # lacks the archive contract commit. Skip those cases there instead of failing every clone.
    $terminalVtArchivePresent = (Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\TestRuns\4cb089111a23\Terminal\20260826_071812_pair2-product') -PathType Container) -and
        (Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\TestRuns\4cb089111a23\Terminal\20260826_071820_pair2-candidate') -PathType Container)
    $fileOpsSummaryArchivePresent = Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\TestRuns\7d3a1247382a\FileOps\2026-07-20_205550') -PathType Container
    $checkedInInventoryPresent = @(Get-RSTestRunArchiveInventoryPaths -RepoRoot $repoRoot).Count -gt 0
    if (-not $terminalVtArchivePresent) {
        Write-Host 'Skipping the Terminal VT archive contracts: Specs\TestRuns\4cb089111a23\Terminal pair2 evidence is not present (gitignored archive).'
    }
    if (-not $fileOpsSummaryArchivePresent) {
        Write-Host 'Skipping the FileOps summary archive contract: Specs\TestRuns\7d3a1247382a\FileOps\2026-07-20_205550 is not present (gitignored archive).'
    }
    if (-not $checkedInInventoryPresent) {
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

    It 'validates every Terminal VT JSONL row and requires its schema-owned artifacts' -Skip:(-not $terminalVtArchivePresent) {
        $fixture = New-RSTerminalVtArchiveFixture
        $candidateRoot = Join-Path $fixture.Root $fixture.Candidate
        Add-Content -LiteralPath (Join-Path $candidateRoot 'perf\perf_metrics.jsonl') -Value '{not-json'
        Remove-Item -LiteralPath (Join-Path $candidateRoot 'trace.txt') -Force

        $paths = @(Get-RSTestRunArchivePathsForRuns -RepoRoot $fixture.Root -RunPath @($fixture.Candidate))
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $paths)
        @($violations | Where-Object Kind -eq 'PerfJson').Count | Should BeGreaterThan 0
        @($violations | Where-Object Kind -eq 'RequiredArtifact').Count | Should Be 1
    }

    It 'rejects Terminal VT results that do not satisfy the area schema' -Skip:(-not $terminalVtArchivePresent) {
        $fixture = New-RSTerminalVtArchiveFixture
        $resultsPath = Join-Path (Join-Path $fixture.Root $fixture.Candidate) 'results.json'
        $results = Get-Content -LiteralPath $resultsPath -Raw | ConvertFrom-Json
        $results.schema = 'red-salamander.terminal-vt-upgrade-evidence.invalid'
        $results | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $resultsPath -Encoding UTF8

        $paths = @(Get-RSTestRunArchivePathsForRuns -RepoRoot $fixture.Root -RunPath @($fixture.Candidate))
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $fixture.Root -Paths $paths)
        @($violations | Where-Object Kind -in @('EvidenceSchema', 'EvidenceJson')).Count | Should BeGreaterThan 0
    }

    It 'recomputes the admitted Terminal VT pair and rejects a candidate beyond baseline x1.10' -Skip:(-not $terminalVtArchivePresent) {
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
