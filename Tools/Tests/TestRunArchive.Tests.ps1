Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRunPlanScript = Join-Path $repoRoot 'Tools\TestRunPlan.ps1'
$archiveModule = Join-Path $repoRoot 'Tools\TestRunArchive.psm1'
. $testRunPlanScript
Import-Module $archiveModule -Force

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

Describe 'TestRuns archive contract' {
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

    It 'accepts an empty path set' {
        $root = New-RSArchiveContractRoot
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $root -Paths @())
        $violations.Count | Should Be 0
    }

    It 'keeps the current changed archive set within the contract' {
        $paths = @(Get-RSChangedTestRunArchivePaths -RepoRoot $repoRoot)
        $violations = @(Get-RSTestRunArchiveViolations -RepoRoot $repoRoot -Paths $paths)
        if ($violations.Count -ne 0) {
            throw (($violations | ForEach-Object { "$($_.Kind): $($_.Path): $($_.Message)" }) -join [Environment]::NewLine)
        }
    }
}
