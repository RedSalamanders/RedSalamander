Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$manifestPath = Join-Path $repoRoot 'Tools\TerminalEngine\HistoricalTerminalProfiles.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

Describe 'Historical Terminal profile integrity' {
    It 'keeps the sealed Gate0 and Round4 test identities complete and unchanged' {
        $definitions = @($manifest.historicalTests)
        $definitions.Count | Should Be 20
        @($definitions.path | Sort-Object -Unique).Count | Should Be 20

        foreach ($definition in $definitions) {
            $testPath = Join-Path $repoRoot ([string]$definition.path).Replace('/', '\')
            (Test-Path -LiteralPath $testPath -PathType Leaf) | Should Be $true
            (Get-FileHash -Algorithm SHA256 -LiteralPath $testPath).Hash.ToLowerInvariant() |
                Should Be ([string]$definition.sha256)
        }

        @($definitions | Where-Object { @($_.profiles) -contains 'Gate0' }).Count |
            Should Be ([int]$manifest.profiles.Gate0.expectedFileCount)
        @($definitions | Where-Object { @($_.profiles) -contains 'Round4' }).Count |
            Should Be ([int]$manifest.profiles.Round4.expectedFileCount)
    }

    It 'keeps current Terminal runtime tests outside the sealed cohort' {
        $historicalPaths = @($manifest.historicalTests.path)
        foreach ($relativePath in @($manifest.currentRuntimeTests)) {
            ($historicalPaths -contains $relativePath) | Should Be $false
            (Test-Path -LiteralPath (Join-Path $repoRoot $relativePath.Replace('/', '\')) -PathType Leaf) |
                Should Be $true
        }

        $runPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
        Import-Module $runPlanModule -Force -ErrorAction Stop
        $arguments = @{
            RepoRoot = $repoRoot
            Platform = 'x64'
            Configuration = 'Debug'
            RedSalamanderExePath = Join-Path $repoRoot '.build\x64\Debug\RedSalamander.exe'
        }
        foreach ($suite in @('CI', 'Full')) {
            $entry = @(Get-RSTestRunPlan -Suite $suite @arguments |
                Where-Object { $_.Name -eq 'ToolsPesterTests' })
            $entry.Count | Should Be 1
            [IO.Path]::GetFileName([string]$entry[0].Path) |
                Should Be 'CurrentToolingProfile.Tests.ps1'
        }

        $profileSource = Get-Content -LiteralPath (
            Join-Path $repoRoot 'Tools\TerminalEngine\CurrentToolingProfile.Tests.ps1') -Raw
        $historicalRunnerSource = Get-Content -LiteralPath (
            Join-Path $repoRoot 'Tools\TerminalEngine\Invoke-HistoricalTerminalTests.ps1') -Raw
        $profileSource | Should Match 'HistoricalTerminalProfiles\.json'
        $profileSource | Should Match '-not \$historicalPaths\.Contains\(\$relativePath\)'
        $historicalRunnerSource | Should Match "ValidateSet\('Gate0', 'Round4'\)"
        $historicalRunnerSource | Should Match "Properties\['expectedPassedCount'\]"
        foreach ($relativePath in @($manifest.currentRuntimeTests)) {
            $testName = [IO.Path]::GetFileName([string]$relativePath)
            $testName | Should Match '\.Tests\.ps1$'
        }
    }

    It 'keeps historical workflows frozen behind their explicit archival triggers' {
        $gate0 = Get-Content -LiteralPath (
            Join-Path $repoRoot '.github\workflows\terminal-engine-gate0.yml') -Raw
        $round4 = Get-Content -LiteralPath (
            Join-Path $repoRoot '.github\workflows\terminal-engine-round4-closeout.yml') -Raw

        (Get-FileHash -Algorithm SHA256 -LiteralPath (
                Join-Path $repoRoot '.github\workflows\terminal-engine-gate0.yml')).Hash.ToLowerInvariant() |
            Should Be 'e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c'
        $gate0 | Should Match '(?m)^  workflow_dispatch:'
        $gate0 | Should Match '(?ms)^  push:\s+branches:\s+- codex/terminal-engine-gate0'
        $gate0 | Should Match "github\.event_name == 'workflow_dispatch' \|\| contains\(github\.event\.head_commit\.message, '\[run-gate0\]'\)"
        $gate0 | Should Not Match '(?m)^  pull_request:'
        $round4 | Should Match '(?m)^  workflow_dispatch:'
        $round4 | Should Not Match '(?m)^  (push|pull_request):'
        $round4 | Should Match '\$tests\.Count -ne 20'
        $round4 | Should Match '\$passed -ne 239'

        $getWorkflowTests = {
            param([Parameter(Mandatory)][string]$WorkflowText)

            return @([regex]::Matches(
                    $WorkflowText,
                    "'\.\\Tools\\Tests\\(?<name>[^']+\.Tests\.ps1)'") |
                ForEach-Object { "Tools/Tests/$($_.Groups['name'].Value)" })
        }
        $gate0Tests = @(& $getWorkflowTests $gate0 | Sort-Object)
        $round4Tests = @(& $getWorkflowTests $round4 | Sort-Object)
        $expectedGate0 = @($manifest.historicalTests |
            Where-Object { @($_.profiles) -contains 'Gate0' } |
            ForEach-Object { [string]$_.path } |
            Sort-Object)
        $expectedRound4 = @($manifest.historicalTests |
            Where-Object { @($_.profiles) -contains 'Round4' } |
            ForEach-Object { [string]$_.path } |
            Sort-Object)
        ($gate0Tests -join "`0") | Should Be ($expectedGate0 -join "`0")
        ($round4Tests -join "`0") | Should Be ($expectedRound4 -join "`0")
    }

    It 'classifies the unused generic lifecycle facade without deleting its stable paths' {
        $facade = $manifest.genericLifecycleFacade
        [string]$facade.classification | Should Be 'internal-historical-future-upgrade'
        foreach ($relativePath in @($facade.entryPoints) + @($facade.compatibilityProject)) {
            (Test-Path -LiteralPath (Join-Path $repoRoot $relativePath.Replace('/', '\')) -PathType Leaf) |
                Should Be $true
        }
        (Test-Path -LiteralPath (Join-Path $repoRoot $facade.defaultLock.Replace('/', '\'))) |
            Should Be $false
        [string]$facade.currentProductBuilder | Should Be 'Tools/Build-GhosttyTerminalRuntime.ps1'
    }
}
