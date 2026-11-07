Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools\Modules\Auditing\RepositorySourceScanner.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Reporting\TestRunReporting.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Reporting\SourceLineMeasurement.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\TestSandbox.psm1') -Force

Describe 'Repository source scanner module' {
    It 'normalizes paths and applies extension, tree, and file exclusions once' {
        # A checkout may itself live beneath .build; only paths inside it are excluded.
        $root = Join-Path $TestDrive '.build\scanner'
        New-Item -ItemType Directory -Path (Join-Path $root 'Source\Nested') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $root 'Source\Generated') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $root 'Source\.build') -Force | Out-Null
        '    TOKEN alpha   ' | Set-Content -LiteralPath (Join-Path $root 'Source\One.cpp') -Encoding UTF8
        'TOKEN ignored build output' | Set-Content -LiteralPath (Join-Path $root 'Source\.build\Generated.cpp') -Encoding UTF8
        'TOKEN beta' | Set-Content -LiteralPath (Join-Path $root 'Source\Nested\Two.h') -Encoding UTF8
        'TOKEN ignored extension' | Set-Content -LiteralPath (Join-Path $root 'Source\Nested\Three.txt') -Encoding UTF8
        'TOKEN ignored tree' | Set-Content -LiteralPath (Join-Path $root 'Source\Generated\Four.cpp') -Encoding UTF8

        $matches = @(Find-RSRepositorySourceMatches `
            -RepositoryRoot $root `
            -SourceRoots @('Source') `
            -Extensions @('.cpp', '*.h') `
            -Patterns @(@{ Name = 'Token'; Pattern = 'TOKEN' }) `
            -ExcludedPathFragments @('\Generated\', '\.build\') `
            -ExcludedRelativePaths @('Source\Nested\Two.h') `
            -PathSeparator Slash)

        $matches.Count | Should Be 1
        $matches[0].PatternName | Should Be 'Token'
        $matches[0].Path | Should Be 'Source/One.cpp'
        $matches[0].LineNumber | Should Be 1
        $matches[0].Line | Should Be 'TOKEN alpha'
        $escapeMessage = ''
        try {
            [void](Get-RSRepositorySourceFiles -RepositoryRoot $root -SourceRoots @('..\outside') -Extensions @('.cpp'))
        } catch {
            $escapeMessage = $_.Exception.Message
        }
        $escapeMessage | Should Match 'Source root escapes the repository'
    }
    It 'rejects a junction at <Name> before scanning it' -TestCases @(
        @{ Name = 'the repository root'; Link = ''; SourceRoot = 'Source'; TargetSubdirectory = 'Source' },
        @{ Name = 'the selected source root'; Link = 'Source'; SourceRoot = 'Source'; TargetSubdirectory = '' },
        @{ Name = 'a selected root ancestor'; Link = 'Parent'; SourceRoot = 'Parent\Source'; TargetSubdirectory = 'Source' }
    ) {
        param($Name, $Link, $SourceRoot, $TargetSubdirectory)
        $fixture = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness tools-pester -Case ('scanner-link-' + [guid]::NewGuid().ToString('N'))
        $repository = Join-Path $fixture 'repository'
        $target = Join-Path $fixture 'target'
        $targetSource = if ($TargetSubdirectory) { Join-Path $target $TargetSubdirectory } else { $target }
        [void][IO.Directory]::CreateDirectory($targetSource)
        [IO.File]::WriteAllText((Join-Path $targetSource 'witness.cpp'), 'TOKEN')
        $junction = if ($Link) { Join-Path $repository $Link } else { $repository }
        [void][IO.Directory]::CreateDirectory((Split-Path $junction))
        [void](New-RSTestSandboxLink -ItemType Junction -Path $junction -Target $target)
        try {
            $message = ''
            try { [void](Get-RSRepositorySourceFiles -RepositoryRoot $repository -SourceRoots @($SourceRoot) -Extensions @('.cpp')) }
            catch { $message = $_.Exception.Message }
            $message | Should Match 'reparse point'
        } finally {
            # This fixture owns the exact link. Unlink it without traversing the target
            # so the completed run remains relocatable under the archive contract.
            [IO.Directory]::Delete($junction, $false)
        }
        (Test-Path -LiteralPath $junction) | Should Be $false
        [IO.File]::ReadAllText((Join-Path $targetSource 'witness.cpp')) | Should Be 'TOKEN'
    }

    It 'does not enumerate a descendant junction' {
        $fixture = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness tools-pester -Case ('scanner-descendant-' + [guid]::NewGuid().ToString('N'))
        $source = Join-Path $fixture 'repository\Source'
        $target = Join-Path $fixture 'target'
        [void][IO.Directory]::CreateDirectory($source)
        [void][IO.Directory]::CreateDirectory($target)
        [IO.File]::WriteAllText((Join-Path $source 'owned.cpp'), 'owned')
        [IO.File]::WriteAllText((Join-Path $target 'must-not-read.cpp'), 'redirected')
        $junction = Join-Path $source 'redirected'
        [void](New-RSTestSandboxLink -ItemType Junction -Path $junction -Target $target)
        try {
            $files = @(Get-RSRepositorySourceFiles -RepositoryRoot (Join-Path $fixture 'repository') -SourceRoots @('Source') -Extensions @('.cpp'))
            $files.Count | Should Be 1
            $files[0].Name | Should Be 'owned.cpp'
        } finally { [IO.Directory]::Delete($junction, $false) }
        (Test-Path -LiteralPath $junction) | Should Be $false
        [IO.File]::ReadAllText((Join-Path $target 'must-not-read.cpp')) | Should Be 'redirected'
    }

    It 'rejects a source file symbolic link when the host permits its creation' {
        $fixture = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness tools-pester -Case ('scanner-file-link-' + [guid]::NewGuid().ToString('N'))
        $source = Join-Path $fixture 'repository\Source'
        [void][IO.Directory]::CreateDirectory($source)
        $target = Join-Path $fixture 'target.cpp'
        [IO.File]::WriteAllText($target, 'redirected')
        try {
            [void](New-RSTestSandboxLink -ItemType SymbolicLink -Path (Join-Path $source 'redirected.cpp') -Target $target)
        } catch {
            if ($_.Exception.Message -notmatch 'privilege|1314') { throw }
            Write-Warning 'File symbolic link creation is unavailable on this host; the file-link rejection was not exercised.'
            return
        }
        try {
            $message = ''
            try { [void](Get-RSRepositorySourceFiles -RepositoryRoot (Join-Path $fixture 'repository') -SourceRoots @('Source') -Extensions @('.cpp')) }
            catch { $message = $_.Exception.Message }
            $message | Should Match 'Source file is a reparse point'
        } finally { [IO.File]::Delete((Join-Path $source 'redirected.cpp')) }
        (Test-Path -LiteralPath (Join-Path $source 'redirected.cpp')) | Should Be $false
        [IO.File]::ReadAllText($target) | Should Be 'redirected'
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

Describe 'Positive test-block success assertion' {
    It 'preserves successful output and makes a deliberate exception fail the success result' {
        $success = Invoke-RSTestBlock { 'success-value' }
        $success.Succeeded | Should Be $true
        @($success.Output).Count | Should Be 1
        $success.Output[0] | Should Be 'success-value'
        $null -eq $success.Error | Should Be $true

        $failure = Invoke-RSTestBlock { throw 'intentional positive-assertion negative control' }
        $failure.Succeeded | Should Be $false
        $failure.Error.Exception.Message | Should Be 'intentional positive-assertion negative control'
    }

    It 'rejects unreliable positive Throw assertions in executable test syntax' {
        $findUnsafeAssertions = {
            param([System.Management.Automation.Language.Ast]$Ast)
            @($Ast.FindAll({
                param($node)
                if ($node -isnot [System.Management.Automation.Language.CommandAst] -or
                    $node.GetCommandName() -ine 'Should') { return $false }

                $parameters = @($node.CommandElements | Where-Object {
                    $_ -is [System.Management.Automation.Language.CommandParameterAst]
                } | ForEach-Object { $_.ParameterName })
                if ($parameters -contains 'Not' -and $parameters -contains 'Throw') { return $true }

                $operators = @($node.CommandElements | Select-Object -Skip 1 -First 2 | ForEach-Object {
                    if ($_ -is [System.Management.Automation.Language.CommandParameterAst]) { $_.ParameterName }
                    elseif ($_ -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $_.Value }
                })
                return $operators.Count -eq 2 -and $operators -contains 'Not' -and $operators -contains 'Throw'
            }, $true))
        }
        $fixture = @'
{ throw 'seed' } | Should Not Throw
{ throw 'seed' } | Should -Not -Throw
{ throw 'seed' } | Should -Throw -Not
{ throw 'seed' } | Should -Not Throw
'Should Not Throw' | Should Match 'literal'
# Should Not Throw is permitted in explanatory comments.
'@
        $tokens = $null
        $parseErrors = $null
        $fixtureAst = [System.Management.Automation.Language.Parser]::ParseInput($fixture, [ref]$tokens, [ref]$parseErrors)
        $parseErrors.Count | Should Be 0
        @(& $findUnsafeAssertions $fixtureAst).Count | Should Be 4

        # These reproduction inputs are byte-sealed, excluded from the current tooling
        # profile, and checked by HistoricalTerminalProfile.Tests.ps1. Apply current
        # assertion policy to the active cohort without rewriting historical evidence.
        $historicalProfiles = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/TerminalEngine/HistoricalTerminalProfiles.json') -Raw | ConvertFrom-Json
        $historicalPaths = @($historicalProfiles.historicalTests.path)
        $historicalPaths.Count | Should Be 20
        $violations = @()
        foreach ($file in @(Get-RSRepositorySourceFiles -RepositoryRoot $repoRoot `
                -SourceRoots @('Tools/Tests') -Extensions @('.ps1') -ExcludedRelativePaths $historicalPaths |
                Where-Object Name -Like '*.Tests.ps1')) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
            $parseErrors.Count | Should Be 0
            foreach ($command in @(& $findUnsafeAssertions $ast)) {
                $violations += '{0}:{1}: {2}' -f $file.FullName, $command.Extent.StartLineNumber, $command.Extent.Text
            }
        }
        ($violations -join "`n") | Should BeNullOrEmpty
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
