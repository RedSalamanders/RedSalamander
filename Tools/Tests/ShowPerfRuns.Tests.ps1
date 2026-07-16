Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRunPlanScript = Join-Path $repoRoot 'Tools\TestRunPlan.ps1'
$showPerfRunsScript = Join-Path $repoRoot 'Tools\Show-PerfRuns.ps1'
. $testRunPlanScript

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

function New-RSPerfRunsRoot {
    $root = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case ("show-perfruns-" + [guid]::NewGuid().ToString('N'))
    [pscustomobject]@{
        budgets = @(
            [pscustomobject]@{ metric = 'folder.frame.total_us'; stat = 'p95'; minimumSamples = 200 }
            [pscustomobject]@{ metric = 'folder.frame.input_to_paint_us'; stat = 'p95'; minimumSamples = 40 }
        )
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $root 'FolderViewPerfBudgets.json5') -Encoding UTF8
    return $root
}

function New-RSPerfMetric {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [int]$Count,

        [double]$StartValue = 1000.0,

        [string]$Unit = 'us'
    )

    return [pscustomobject]@{
        Name = $Name
        Count = $Count
        StartValue = $StartValue
        Unit = $Unit
    }
}

function New-RSPerfRun {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,

        [Parameter(Mandatory = $true)]
        [string]$RunName,

        [Parameter(Mandatory = $true)]
        [object[]]$Metrics
    )

    $runRoot = Join-Path $RunsRoot "test-machine\Commands\$RunName"
    $perfRoot = Join-Path $runRoot 'perf'
    New-Item -ItemType Directory -Path $perfRoot -Force | Out-Null
    @(
        'git_branch: synthetic'
        'git_commit: synthetic'
    ) | Set-Content -LiteralPath (Join-Path $runRoot 'env.txt') -Encoding UTF8

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($metric in $Metrics) {
        for ($i = 0; $i -lt $metric.Count; $i++) {
            $record = [pscustomobject]@{
                metric = $metric.Name
                scenario = 'gr8-show-perfruns'
                value = ([double]$metric.StartValue + [double]$i)
                unit = $metric.Unit
                build = 'Release'
            }
            $lines.Add(($record | ConvertTo-Json -Compress))
        }
    }

    $lines | Set-Content -LiteralPath (Join-Path $perfRoot 'perf_metrics.jsonl') -Encoding UTF8
    return $runRoot
}

function Invoke-RSPwsh {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $pwshCommand = Get-Command pwsh -ErrorAction SilentlyContinue
    if (-not $pwshCommand -or [string]::IsNullOrWhiteSpace($pwshCommand.Source)) {
        throw 'Show-PerfRuns.ps1 requires PowerShell 7+; unable to locate pwsh.'
    }
    $pwsh = $pwshCommand.Source

    $outputRoot = Join-Path $RunsRoot 'process-output'
    New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
    $stdout = Join-Path $outputRoot ([guid]::NewGuid().ToString('N') + '.stdout.txt')
    $stderr = Join-Path $outputRoot ([guid]::NewGuid().ToString('N') + '.stderr.txt')

    $process = Start-Process -FilePath $pwsh `
        -ArgumentList $Arguments `
        -WorkingDirectory $repoRoot `
        -Wait `
        -PassThru `
        -NoNewWindow `
        -RedirectStandardOutput $stdout `
        -RedirectStandardError $stderr

    return [pscustomobject]@{
        ExitCode = [int]$process.ExitCode
        Stdout = if (Test-Path -LiteralPath $stdout) { Get-Content -LiteralPath $stdout -Raw } else { '' }
        Stderr = if (Test-Path -LiteralPath $stderr) { Get-Content -LiteralPath $stderr -Raw } else { '' }
    }
}

function Invoke-RSShowPerfRuns {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $processArguments = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        $showPerfRunsScript,
        '-RunsRoot',
        $RunsRoot,
        '-BudgetPath',
        (Join-Path $RunsRoot 'FolderViewPerfBudgets.json5')
    ) + $Arguments

    return Invoke-RSPwsh -RunsRoot $RunsRoot -Arguments $processArguments
}

function ConvertTo-RSPowerShellSingleQuotedString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return "'" + $Value.Replace("'", "''") + "'"
}

function Invoke-RSShowPerfRunsCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,

        [Parameter(Mandatory = $true)]
        [string]$CommandText
    )

    return Invoke-RSPwsh -RunsRoot $RunsRoot -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $CommandText)
}

Describe 'Show-PerfRuns quality gates' {
    It 'does not fail FolderViewPreset quality for one-shot gauges when distributions have enough samples' {
        $root = New-RSPerfRunsRoot
        try {
            $run = New-RSPerfRun -RunsRoot $root -RunName 'candidate' -Metrics @(
                (New-RSPerfMetric -Name 'folder.frame.total_us' -Count 200)
                (New-RSPerfMetric -Name 'folder.scale.working_set_bytes' -Count 1 -StartValue 123456789 -Unit 'bytes')
            )

            $result = Invoke-RSShowPerfRuns -RunsRoot $root -Arguments @('-Run', $run, '-FolderViewPreset', '-FailOnQuality')

            Assert-RSEqual -Actual $result.ExitCode -Expected 0 -Message "Gauge rows should not latch -FailOnQuality. Stdout=$($result.Stdout) Stderr=$($result.Stderr)"
        } finally {
            if (Test-Path -LiteralPath $root) {
                Remove-Item -LiteralPath $root -Recurse -Force
            }
        }
    }

    It 'uses FolderView budget minimumSamples instead of the global default when available' {
        $root = New-RSPerfRunsRoot
        try {
            $run = New-RSPerfRun -RunsRoot $root -RunName 'candidate' -Metrics @(
                (New-RSPerfMetric -Name 'folder.frame.input_to_paint_us' -Count 40)
            )

            $result = Invoke-RSShowPerfRuns -RunsRoot $root -Arguments @('-Run', $run, '-FolderViewPreset', '-FailOnQuality')

            Assert-RSEqual -Actual $result.ExitCode -Expected 0 -Message "folder.frame.input_to_paint_us has a budget minimum of 40 samples. Stdout=$($result.Stdout) Stderr=$($result.Stderr)"
        } finally {
            if (Test-Path -LiteralPath $root) {
                Remove-Item -LiteralPath $root -Recurse -Force
            }
        }
    }

    It 'treats FolderViewPreset metrics without p95 budgets as informational for quality gating' {
        $root = New-RSPerfRunsRoot
        try {
            $run = New-RSPerfRun -RunsRoot $root -RunName 'candidate' -Metrics @(
                (New-RSPerfMetric -Name 'icons.extract_us' -Count 1)
            )

            $result = Invoke-RSShowPerfRuns -RunsRoot $root -Arguments @('-Run', $run, '-FolderViewPreset', '-FailOnQuality')

            Assert-RSEqual -Actual $result.ExitCode -Expected 0 -Message "Preset metrics without p95 budget rows should not latch -FailOnQuality. Stdout=$($result.Stdout) Stderr=$($result.Stderr)"
        } finally {
            if (Test-Path -LiteralPath $root) {
                Remove-Item -LiteralPath $root -Recurse -Force
            }
        }
    }

    It 'still fails FolderViewPreset quality for low-sample distribution metrics' {
        $root = New-RSPerfRunsRoot
        try {
            $run = New-RSPerfRun -RunsRoot $root -RunName 'candidate' -Metrics @(
                (New-RSPerfMetric -Name 'folder.frame.total_us' -Count 199)
            )

            $result = Invoke-RSShowPerfRuns -RunsRoot $root -Arguments @('-Run', $run, '-FolderViewPreset', '-FailOnQuality')

            Assert-RSEqual -Actual $result.ExitCode -Expected 2 -Message "Sparse distribution rows should latch -FailOnQuality. Stdout=$($result.Stdout) Stderr=$($result.Stderr)"
        } finally {
            if (Test-Path -LiteralPath $root) {
                Remove-Item -LiteralPath $root -Recurse -Force
            }
        }
    }

    It 'honors an explicitly supplied MinimumSamplesForP95 ahead of budget defaults' {
        $root = New-RSPerfRunsRoot
        try {
            $run = New-RSPerfRun -RunsRoot $root -RunName 'candidate' -Metrics @(
                (New-RSPerfMetric -Name 'folder.frame.total_us' -Count 5)
            )

            $result = Invoke-RSShowPerfRuns -RunsRoot $root -Arguments @('-Run', $run, '-FolderViewPreset', '-MinimumSamplesForP95', '5', '-FailOnQuality')

            Assert-RSEqual -Actual $result.ExitCode -Expected 0 -Message "Explicit sample minimum should override the budget default. Stdout=$($result.Stdout) Stderr=$($result.Stderr)"
        } finally {
            if (Test-Path -LiteralPath $root) {
                Remove-Item -LiteralPath $root -Recurse -Force
            }
        }
    }

    It 'warns and fails quality when an existing budget file cannot parse' {
        $root = New-RSPerfRunsRoot
        try {
            $run = New-RSPerfRun -RunsRoot $root -RunName 'candidate' -Metrics @(
                (New-RSPerfMetric -Name 'folder.frame.total_us' -Count 200)
            )
            Set-Content -LiteralPath (Join-Path $root 'FolderViewPerfBudgets.json5') -Value '{ malformed' -Encoding UTF8

            $result = Invoke-RSShowPerfRuns -RunsRoot $root -Arguments @('-Run', $run, '-FolderViewPreset', '-FailOnQuality')

            Assert-RSEqual -Actual $result.ExitCode -Expected 2 -Message "Malformed existing budgets should latch quality failure. Stdout=$($result.Stdout) Stderr=$($result.Stderr)"
            if (("$($result.Stdout)`n$($result.Stderr)") -notmatch 'Unable to parse existing perf budget file') {
                throw "Malformed existing budgets should emit a warning. Stdout=$($result.Stdout) Stderr=$($result.Stderr)"
            }
        } finally {
            if (Test-Path -LiteralPath $root) {
                Remove-Item -LiteralPath $root -Recurse -Force
            }
        }
    }

    It 'does not fail compare-mode quality because the baseline run is sparse' {
        $root = New-RSPerfRunsRoot
        try {
            $baseline = New-RSPerfRun -RunsRoot $root -RunName 'baseline' -Metrics @(
                (New-RSPerfMetric -Name 'folder.frame.total_us' -Count 1)
            )
            $candidate = New-RSPerfRun -RunsRoot $root -RunName 'candidate' -Metrics @(
                (New-RSPerfMetric -Name 'folder.frame.total_us' -Count 200)
            )

            $commandText = '& {0} -RunsRoot {1} -BudgetPath {2} -CompareRun @({3}, {4}) -FolderViewPreset -FailOnQuality' -f `
                (ConvertTo-RSPowerShellSingleQuotedString $showPerfRunsScript),
                (ConvertTo-RSPowerShellSingleQuotedString $root),
                (ConvertTo-RSPowerShellSingleQuotedString (Join-Path $root 'FolderViewPerfBudgets.json5')),
                (ConvertTo-RSPowerShellSingleQuotedString $baseline),
                (ConvertTo-RSPowerShellSingleQuotedString $candidate)
            $result = Invoke-RSShowPerfRunsCommand -RunsRoot $root -CommandText $commandText

            Assert-RSEqual -Actual $result.ExitCode -Expected 0 -Message "Compare mode should gate candidate quality only. Stdout=$($result.Stdout) Stderr=$($result.Stderr)"
        } finally {
            if (Test-Path -LiteralPath $root) {
                Remove-Item -LiteralPath $root -Recurse -Force
            }
        }
    }
}
