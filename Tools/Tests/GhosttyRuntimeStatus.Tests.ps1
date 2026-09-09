Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyRuntimeStatus.psm1'
$commandPath = Join-Path $repoRoot 'Tools\Get-GhosttyTerminalRuntimeStatus.ps1'
$statusSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeStatus.schema.json'
$productLock = Get-Content -LiteralPath (Join-Path $repoRoot 'External\TerminalEngine\GhosttyRuntimeLock.v1.json') -Raw |
    ConvertFrom-Json -Depth 100 -NoEnumerate -DateKind String
$lockedCommit = [string]$productLock.upstream.commit
$lockedTree = [string]$productLock.upstream.tree
$candidateCommit = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
$candidateTree = 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb'
$fixedUtc = [DateTime]::Parse('2026-08-25T12:34:56Z').ToUniversalTime()

Import-Module $modulePath -Force -ErrorAction Stop

function New-RSGhosttyStatusObservation {
    param(
        [string]$Commit = $candidateCommit,
        [string]$Tree = $candidateTree,
        [string]$Version = '1.3.2-dev',
        [AllowNull()][object]$AheadCount = 717,
        [AllowNull()][object]$BehindCount = 0,
        [string]$ComparisonStatus = 'ahead'
    )

    return [ordered]@{
        commit = $Commit
        tree = $Tree
        version = $Version
        aheadCount = $AheadCount
        behindCount = $BehindCount
        comparisonStatus = $ComparisonStatus
    }
}

function New-RSGhosttyStatusFixture {
    param(
        [ValidateSet('success', 'malformed', 'rate-limit', 'timeout', 'refused')]
        [string]$Outcome = 'success',
        [AllowNull()][object]$Observation = (New-RSGhosttyStatusObservation),
        [AllowNull()][object]$Candidate = (New-RSGhosttyStatusObservation),
        [object[]]$Advisories = @()
    )

    $path = Join-Path $TestDrive ("status-fixture-$([Guid]::NewGuid().ToString('N')).json")
    [ordered]@{
        schemaVersion = 1
        outcome = $Outcome
        observation = $Observation
        candidate = $Candidate
        advisories = $Advisories
    } | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $path -Encoding utf8NoBOM
    return $path
}

function New-RSGhosttyStatusTestOutput {
    $root = Join-Path $repoRoot ".build\TerminalEngineStatus\Tests\$([Guid]::NewGuid().ToString('N'))"
    [void](New-Item -ItemType Directory -Path $root -Force)
    return $root
}

function Assert-RSGhosttyStatusRejected {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)

    $message = ''
    try {
        [void](& $Action)
    }
    catch {
        $message = $_.Exception.Message
    }
    [string]::IsNullOrWhiteSpace($message) | Should Be $false
}

Describe 'Ghostty runtime read-only observation status' {
    It 'publishes a strict canonical network-free Locked report' {
        $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Locked -GeneratedUtc $fixedUtc
        $report.result | Should Be 'current'
        $report.nextAction | Should Be 'none'
        @($report.networkAttempts).Count | Should Be 0
        $report.canonicalLock.commit | Should Be $lockedCommit
        $report.observation.commit | Should Be $lockedCommit
        $report.sourceBinding.repositoryCommit | Should Match '^[0-9a-f]{40}$'
        $json = ConvertTo-Json $report -Depth 30
        Test-Json -Json $json -SchemaFile $statusSchemaPath | Should Be $true
        $json | Should Not Match '(?i)authorization|responseHeaders|credential|token'
    }

    It 'runs the documented Locked command in a fresh PowerShell process' {
        $root = New-RSGhosttyStatusTestOutput
        $outputPath = Join-Path $root 'fresh-process.json'
        try {
            $output = & pwsh.exe -NoProfile -File $commandPath -Mode Locked `
                -GeneratedUtc $fixedUtc.ToString('o') -OutputPath $outputPath 2>&1
            $exitCode = $LASTEXITCODE
            $exitCode | Should Be 0
            ($output -join "`n") | Should Match 'mode=Locked result=current'
            Test-Path -LiteralPath $outputPath -PathType Leaf | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force
        }
    }

    It 'produces identical canonical bytes for repeated fixed-time reports' {
        $root = New-RSGhosttyStatusTestOutput
        try {
            $first = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Locked -GeneratedUtc $fixedUtc
            $second = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Locked -GeneratedUtc $fixedUtc
            $firstPath = Write-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Report $first -OutputPath (Join-Path $root 'first.json')
            $secondPath = Write-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Report $second -OutputPath (Join-Path $root 'second.json')
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($firstPath)) |
                Should Be ([Convert]::ToBase64String([IO.File]::ReadAllBytes($secondPath)))
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force
        }
    }

    It 'classifies a default-branch descendant as an observed update without recommendation' {
        $fixture = New-RSGhosttyStatusFixture
        $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Online -NetworkFixturePath $fixture -GeneratedUtc $fixedUtc
        $report.result | Should Be 'update-observed'
        $report.nextAction | Should Be 'review-update'
        $report.observation.commit | Should Be $candidateCommit
        (@($report.reasonCodes) -contains 'descendant-update-observed') | Should Be $true
    }

    It 'requires manual review for unrelated history' {
        $fixture = New-RSGhosttyStatusFixture -Observation (New-RSGhosttyStatusObservation -Commit ('a' * 40) -Tree ('b' * 40) -AheadCount 3 -BehindCount 2 -ComparisonStatus diverged)
        $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Online -NetworkFixturePath $fixture -GeneratedUtc $fixedUtc
        $report.result | Should Be 'required-manual'
        (@($report.reasonCodes) -contains 'unrelated-history') | Should Be $true
    }

    It 'records exactly the requested candidate and refuses branch-head substitution' {
        $exactFixture = New-RSGhosttyStatusFixture
        $exact = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Candidate -CandidateCommit $candidateCommit -NetworkFixturePath $exactFixture -GeneratedUtc $fixedUtc
        $exact.result | Should Be 'update-observed'
        $exact.observation.commit | Should Be $candidateCommit
        $exact.observation.tree | Should Be $candidateTree
        $exact.observation.candidateRequestedCommit | Should Be $candidateCommit
        $exact.nextAction | Should Be 'freeze-candidate'

        $substitutionFixture = New-RSGhosttyStatusFixture -Candidate (New-RSGhosttyStatusObservation -Commit ('c' * 40) -Tree ('d' * 40))
        $refused = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Candidate -CandidateCommit $candidateCommit -NetworkFixturePath $substitutionFixture -GeneratedUtc $fixedUtc
        $refused.result | Should Be 'required-manual'
        $refused.observation.commit | Should Be $null
        (@($refused.reasonCodes) -contains 'candidate-substitution-refused') | Should Be $true
    }

    It 'reports a deleted or refused candidate as unreachable without substitution' {
        $fixture = New-RSGhosttyStatusFixture -Outcome refused
        $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Candidate -CandidateCommit $candidateCommit -NetworkFixturePath $fixture -GeneratedUtc $fixedUtc
        $report.result | Should Be 'unreachable'
        $report.observation.commit | Should Be $null
        $report.observation.candidateRequestedCommit | Should Be $candidateCommit
        $report.nextAction | Should Be 'retry-observation'
    }

    It 'classifies affected and unresolved advisories for security review' {
        foreach ($disposition in @('affected', 'unknown')) {
            $fixture = New-RSGhosttyStatusFixture -Advisories @(
                [ordered]@{
                    identifier = "GHSA-test-$disposition"
                    url = "https://osv.dev/vulnerability/GHSA-test-$disposition"
                    disposition = $disposition
                    retrievedUtc = '2026-08-25T12:34:56.000Z'
                })
            $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Online -NetworkFixturePath $fixture -GeneratedUtc $fixedUtc
            $report.result | Should Be 'security-review'
            $report.nextAction | Should Be 'review-security'
        }
    }

    It 'keeps advisory-only observation bound to the locked runtime' {
        $fixture = New-RSGhosttyStatusFixture -Observation (New-RSGhosttyStatusObservation) -Advisories @(
            [ordered]@{
                identifier = 'GHSA-daily-advisory'
                url = 'https://osv.dev/vulnerability/GHSA-daily-advisory'
                disposition = 'affected'
                retrievedUtc = '2026-08-25T12:34:56.000Z'
            })
        $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Online -AdvisoryOnly -NetworkFixturePath $fixture -GeneratedUtc $fixedUtc
        $report.result | Should Be 'security-review'
        $report.observation.source | Should Be 'canonical-lock'
        $report.observation.commit | Should Be $lockedCommit
        $report.observation.aheadCount | Should Be 0
        @($report.networkAttempts).Count | Should Be 1
    }

    It 'does not escalate an explicitly not-affected advisory' {
        $fixture = New-RSGhosttyStatusFixture -Advisories @(
            [ordered]@{
                identifier = 'GHSA-test-not-affected'
                url = 'https://osv.dev/vulnerability/GHSA-test-not-affected'
                disposition = 'not-affected'
                retrievedUtc = '2026-08-25T12:34:56.000Z'
            })
        $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Online -NetworkFixturePath $fixture -GeneratedUtc $fixedUtc
        $report.result | Should Be 'update-observed'
        @($report.advisories).Count | Should Be 1
        $report.advisories[0].disposition | Should Be 'not-affected'
    }

    It 'maps malformed responses to manual review and rate limits or timeouts to unreachable' {
        $malformed = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Online -NetworkFixturePath (New-RSGhosttyStatusFixture -Outcome malformed) -GeneratedUtc $fixedUtc
        $malformed.result | Should Be 'required-manual'
        (@($malformed.reasonCodes) -contains 'observation-malformed') | Should Be $true
        foreach ($outcome in @('rate-limit', 'timeout')) {
            $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Online -NetworkFixturePath (New-RSGhosttyStatusFixture -Outcome $outcome) -GeneratedUtc $fixedUtc
            $report.result | Should Be 'unreachable'
            (@($report.reasonCodes) -contains "network-$outcome") | Should Be $true
        }
    }

    It 'rejects incomplete mode arguments and output escape paths' {
        Assert-RSGhosttyStatusRejected {
            New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Candidate -CandidateCommit 'main'
        }
        Assert-RSGhosttyStatusRejected {
            New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Locked -NetworkFixturePath 'unused.json'
        }
        $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Locked -GeneratedUtc $fixedUtc
        Assert-RSGhosttyStatusRejected {
            Write-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Report $report -OutputPath (Join-Path $TestDrive 'escape.json')
        }
    }

    It 'binds release evidence to the exact checked-out source commit' {
        $actual = (& git.exe -C $repoRoot rev-parse HEAD).Trim()
        $bound = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Locked -GeneratedUtc $fixedUtc -ExpectedSourceCommit $actual
        $bound.sourceBinding.repositoryCommit | Should Be $actual
        Assert-RSGhosttyStatusRejected {
            New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Locked -GeneratedUtc $fixedUtc -ExpectedSourceCommit ('f' * 40)
        }
    }

    It 'writes only ignored build evidence and leaves tracked status unchanged' {
        $root = New-RSGhosttyStatusTestOutput
        try {
            $before = (& git.exe -C $repoRoot status --porcelain=v1 --untracked-files=no) -join "`n"
            & $commandPath -Mode Online -NetworkFixturePath (New-RSGhosttyStatusFixture) -GeneratedUtc $fixedUtc -OutputPath (Join-Path $root 'status.json')
            $after = (& git.exe -C $repoRoot status --porcelain=v1 --untracked-files=no) -join "`n"
            $after | Should Be $before
            Test-Path -LiteralPath (Join-Path $root 'status.json') -PathType Leaf | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force
        }
    }

    It 'keeps the status schema closed' {
        $report = New-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Mode Locked -GeneratedUtc $fixedUtc
        $report.unexpected = 'rejected'
        Test-Json -Json (ConvertTo-Json $report -Depth 30) -SchemaFile $statusSchemaPath -ErrorAction SilentlyContinue |
            Should Be $false
    }
}
