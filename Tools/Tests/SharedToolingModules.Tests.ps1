Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools\Modules\Auditing\RepositorySourceScanner.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Reporting\TestRunReporting.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Reporting\SourceLineMeasurement.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force

Describe 'Repository source scanner module' {
    It 'normalizes paths and applies extension, tree, and file exclusions once' {
        $root = Join-Path $TestDrive 'scanner'
        New-Item -ItemType Directory -Path (Join-Path $root 'Source\Nested') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $root 'Source\Generated') -Force | Out-Null
        'TOKEN alpha' | Set-Content -LiteralPath (Join-Path $root 'Source\One.cpp') -Encoding UTF8
        'TOKEN beta' | Set-Content -LiteralPath (Join-Path $root 'Source\Nested\Two.h') -Encoding UTF8
        'TOKEN ignored extension' | Set-Content -LiteralPath (Join-Path $root 'Source\Nested\Three.txt') -Encoding UTF8
        'TOKEN ignored tree' | Set-Content -LiteralPath (Join-Path $root 'Source\Generated\Four.cpp') -Encoding UTF8

        $matches = @(Find-RSRepositorySourceMatches `
            -RepositoryRoot $root `
            -SourceRoots @('Source') `
            -Extensions @('.cpp', '*.h') `
            -Patterns @(@{ Name = 'Token'; Pattern = 'TOKEN' }) `
            -ExcludedPathFragments @('\Generated\') `
            -ExcludedRelativePaths @('Source\Nested\Two.h') `
            -PathSeparator Slash)

        $matches.Count | Should Be 1
        $matches[0].PatternName | Should Be 'Token'
        $matches[0].Path | Should Be 'Source/One.cpp'
        $matches[0].LineNumber | Should Be 1
    }
}

Describe 'Test-run reporting module' {
    It 'uses stable candidate order and makes missing-artifact behavior explicit' {
        $run = Join-Path $TestDrive 'run'
        New-Item -ItemType Directory -Path $run -Force | Out-Null
        '{"suite":"selftest"}' | Set-Content -LiteralPath (Join-Path $run 'results.json') -Encoding UTF8
        '{"suite":"commands"}' | Set-Content -LiteralPath (Join-Path $run 'commands_results.json') -Encoding UTF8

        (Split-Path -Leaf (Find-RSTestRunResultsJson -RunRoot $run -Suite Auto -MissingPolicy Throw)) |
            Should Be 'commands_results.json'
        (Split-Path -Leaf (Find-RSTestRunResultsJson -RunRoot $run -Suite SelfTest -MissingPolicy Throw)) |
            Should Be 'results.json'

        $missing = Find-RSTestRunTraceFile -RunRoot $run -Suite Commands -MissingPolicy ReturnNull
        ($null -eq $missing) | Should Be $true
        $expectedMessage = "Could not find trace text in run folder: $run"
        { Find-RSTestRunTraceFile -RunRoot $run -Suite Commands -MissingPolicy Throw } |
            Should Throw $expectedMessage
    }

    It 'preserves shared totals, percentage formatting, and summary colors' {
        $resultsWithCases = [pscustomobject]@{
            passed = 1
            failed = 0
            skipped = 0
            cases = @([pscustomobject]@{ name = 'a' }, [pscustomobject]@{ name = 'b' })
        }
        (Get-RSTestRunTotalCases -Results $resultsWithCases) | Should Be 2
        (Get-RSTestRunTotalCases -Passed 0 -Failed 0 -Skipped 0 -FallbackCaseCount 7) | Should Be 7
        (Format-RSPercent 1 4) | Should Be '25.0%'
        (Format-RSDeltaPercent 100 125) | Should Be '+25.0%'
        (Format-RSDeltaPercent 100 100) | Should Be '0%'
        (Get-RSTestRunSummaryColor 1 0) | Should Be 'Red'
        (Get-RSTestRunSummaryColor 0 1) | Should Be 'Yellow'
        (Get-RSTestRunSummaryColor 0 0) | Should Be 'Green'
    }
}

Describe 'Source-line measurement ownership' {
    It 'derives a minimal solution-backed directory set' {
        $root = Join-Path $TestDrive 'solution'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        $solutionPath = Join-Path $root 'Fixture.sln'
        @'
Project("{11111111-1111-1111-1111-111111111111}") = "App", "src\App\App.vcxproj", "{22222222-2222-2222-2222-222222222222}"
EndProject
Project("{11111111-1111-1111-1111-111111111111}") = "Nested", "src\Nested\Nested.vcxproj", "{33333333-3333-3333-3333-333333333333}"
EndProject
Project("{44444444-4444-4444-4444-444444444444}") = "Solution Items", "Solution Items", "{55555555-5555-5555-5555-555555555555}"
    ProjectSection(SolutionItems) = preProject
        Common\Helper.h = Common\Helper.h
        src\Shared.h = src\Shared.h
    EndProjectSection
EndProject
'@ | Set-Content -LiteralPath $solutionPath -Encoding UTF8

        $directories = @(Get-RSSolutionSourceDirectories -SolutionPath $solutionPath -RepositoryRoot $root)
        $directories.Count | Should Be 2
        $directories[0] | Should Be 'src'
        $directories[1] | Should Be 'Common'
    }

    It 'keeps one source-line implementation behind both supported command paths' {
        $source = Get-RSText -Path 'Tools\Measure-SourceLines.ps1'
        $source | Should Match "ValidateSet\('Repository', 'Solution'\)"
        $source | Should Match '\[string\]\$Scope = ''Repository'''

        $manualSource = Get-RSText -Path 'Tools\InvokeSolutionCloc.ps1'
        $manualSource | Should Match 'Join-Path \$PSScriptRoot ''Measure-SourceLines\.ps1'''
        $manualSource | Should Match '(?s)& \$canonicalCommand.*-Scope Solution.*-ClocArgs \$ClocArgs'
        $manualSource | Should Not Match '(?m)^function\s'
    }

    It 'preserves trailing raw cloc arguments through the manual command' {
        $fakeCloc = Join-Path $TestDrive 'fake-cloc.ps1'
        @'
[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
)

if ($Arguments -contains '--version') {
    'fake-cloc 1.0'
    return
}

"FAKE_CLOC_ARGS=$($Arguments -join '|')"
'@ | Set-Content -LiteralPath $fakeCloc -Encoding UTF8

        $commandPath = Join-Path $repoRoot 'Tools\InvokeSolutionCloc.ps1'
        $output = @(& $commandPath -ClocPath $fakeCloc --by-file --xml 6>&1)
        ($output -join "`n") | Should Match 'FAKE_CLOC_ARGS=--by-file\|--xml\|'
    }

    It 'preserves a nonzero cloc exit status through the manual command' {
        $fakeCloc = Join-Path $TestDrive 'failing-cloc.ps1'
        @'
[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
)

if ($Arguments -contains '--version') {
    'failing-cloc 1.0'
    return
}

exit 7
'@ | Set-Content -LiteralPath $fakeCloc -Encoding UTF8

        $pwsh = (Get-Command pwsh -ErrorAction Stop).Source
        $commandPath = Join-Path $repoRoot 'Tools\InvokeSolutionCloc.ps1'
        & $pwsh -NoLogo -NoProfile -File $commandPath -ClocPath $fakeCloc *> $null
        $LASTEXITCODE | Should Be 7
    }
}
