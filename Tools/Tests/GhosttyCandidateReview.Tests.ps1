Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyCandidateReview.psm1'
$commandPath = Join-Path $repoRoot 'Tools\New-GhosttyTerminalRuntimeCandidateReview.ps1'
$reviewSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeCandidateReview.schema.json'
$descriptorSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeCandidateDescriptor.schema.json'

Import-Module $modulePath -Force -ErrorAction Stop

function Invoke-RSGhosttyCandidateTestGit {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    $output = @(& git.exe -C $Root @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "git failed: $($output -join "`n")" }
    return [string[]]$output
}

function New-RSGhosttyCandidateReviewRepository {
    $root = Join-Path $TestDrive 'candidate-review-repository'
    [void](New-Item -ItemType Directory -Path (Join-Path $root 'include\ghostty\vt') -Force)
    [void](New-Item -ItemType Directory -Path (Join-Path $root 'src\build') -Force)
    [void](New-Item -ItemType Directory -Path (Join-Path $root 'src\terminal') -Force)
    Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('init', '--initial-branch=main') | Out-Null
    Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('config', 'user.name', 'Candidate Review Test') | Out-Null
    Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('config', 'user.email', 'candidate-review@example.invalid') | Out-Null
    Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('config', 'core.autocrlf', 'false') | Out-Null

    @'
GHOSTTY_API GhosttyResult ghostty_terminal_mode_get(void);
GHOSTTY_API GhosttyResult ghostty_terminal_keep(void);
typedef enum { GHOSTTY_TERMINAL_OPT_A = 1 } GhosttyTerminalOption;
'@ | Set-Content -LiteralPath (Join-Path $root 'include\ghostty\vt\terminal.h') -Encoding utf8NoBOM
    '.version = "1.0-test",' | Set-Content -LiteralPath (Join-Path $root 'build.zig.zon') -Encoding utf8NoBOM
    'base build' | Set-Content -LiteralPath (Join-Path $root 'build.zig') -Encoding utf8NoBOM
    'base lib build' | Set-Content -LiteralPath (Join-Path $root 'src\build\GhosttyLibVt.zig') -Encoding utf8NoBOM
    'base lib' | Set-Content -LiteralPath (Join-Path $root 'src\lib_vt.zig') -Encoding utf8NoBOM
    'base parser' | Set-Content -LiteralPath (Join-Path $root 'src\terminal\Parser.zig') -Encoding utf8NoBOM
    Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('add', '.') | Out-Null
    Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('commit', '-m', 'base') | Out-Null
    $base = (Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('rev-parse', 'HEAD') | Select-Object -First 1)

    @'
GHOSTTY_API GhosttyResult ghostty_terminal_get(void);
GHOSTTY_API GhosttyResult ghostty_terminal_keep(void);
typedef enum { GHOSTTY_TERMINAL_OPT_A = 1, GHOSTTY_TERMINAL_OPT_B = 2 } GhosttyTerminalOption;
'@ | Set-Content -LiteralPath (Join-Path $root 'include\ghostty\vt\terminal.h') -Encoding utf8NoBOM
    'candidate parser with bounded OSC state' | Set-Content -LiteralPath (Join-Path $root 'src\terminal\Parser.zig') -Encoding utf8NoBOM
    [void](New-Item -ItemType Directory -Path (Join-Path $root 'src\terminal\kitty') -Force)
    'pending image allocation' | Set-Content -LiteralPath (Join-Path $root 'src\terminal\kitty\graphics.zig') -Encoding utf8NoBOM
    Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('add', '.') | Out-Null
    Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('commit', '-m', 'candidate bounds and kitty image') | Out-Null
    $candidate = (Invoke-RSGhosttyCandidateTestGit -Root $root -Arguments @('rev-parse', 'HEAD') | Select-Object -First 1)
    return [pscustomobject]@{ Root = $root; Base = $base; Candidate = $candidate }
}

Describe 'Ghostty exact candidate review' {
    It 'classifies every review category deterministically' {
        Get-RSGhosttyCandidatePathCategory 'include/ghostty/vt/terminal.h' | Should Be 'public-abi'
        Get-RSGhosttyCandidatePathCategory 'build.zig' | Should Be 'build-toolchain'
        Get-RSGhosttyCandidatePathCategory 'src/terminal/kitty/graphics.zig' | Should Be 'kitty-graphics'
        Get-RSGhosttyCandidatePathCategory 'src/terminal/snapshot/reader.zig' | Should Be 'snapshot-serialization'
        Get-RSGhosttyCandidatePathCategory 'src/terminal/osc.zig' | Should Be 'clipboard-security'
        Get-RSGhosttyCandidatePathCategory 'src/terminal/page_list.zig' | Should Be 'allocation-memory'
        Get-RSGhosttyCandidatePathCategory 'src/terminal/Parser.zig' | Should Be 'parser-state'
    }

    It 'derives complete inventory and public ABI deltas from exact Git objects' {
        $fixture = New-RSGhosttyCandidateReviewRepository
        $inventory = Get-RSGhosttyCandidateChangeInventory -SourcePath $fixture.Root -BaseCommit $fixture.Base -CandidateCommit $fixture.Candidate
        $inventory.fileCount | Should Be 3
        (($inventory.entries | ForEach-Object path) -contains 'include/ghostty/vt/terminal.h') | Should Be $true
        (($inventory.entries | ForEach-Object category) -contains 'kitty-graphics') | Should Be $true

        $abi = Get-RSGhosttyCandidateAbiReview -SourcePath $fixture.Root -BaseCommit $fixture.Base -CandidateCommit $fixture.Candidate `
            -LoaderRequiredExports @('ghostty_terminal_keep', 'ghostty_terminal_mode_get')
        $abi.declarations.baseCount | Should Be 2
        $abi.declarations.candidateCount | Should Be 2
        @($abi.declarations.added).Count | Should Be 1
        @($abi.declarations.removed).Count | Should Be 1
        $removedLoader = @($abi.loaderRequiredExports.entries | Where-Object { -not $_.candidateDeclared })
        $removedLoader.Count | Should Be 1
        $removedLoader[0].name | Should Be 'ghostty_terminal_mode_get'
        $removedLoader[0].replacement | Should Be 'ghostty_terminal_get'
    }

    It 'keeps the review and descriptor schemas closed' {
        foreach ($path in @($reviewSchemaPath, $descriptorSchemaPath)) {
            $schema = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -Depth 100
            $schema.additionalProperties | Should Be $false
            @($schema.required).Count | Should BeGreaterThan 5
        }
    }

    It 'keeps the public command thin and rejects non-full candidate identities' {
        (Get-Content -LiteralPath $commandPath).Count | Should BeLessThan 100
        (Get-Content -LiteralPath $commandPath -Raw) | Should Match 'GhosttyCandidateReview\.psm1'
        $message = ''
        try {
            New-RSGhosttyRuntimeCandidateReview -RepoRoot $repoRoot -CandidateCommit main -SourcePath $TestDrive `
                -ArchivePath (Join-Path $TestDrive 'one.tar.gz') -RepeatArchivePath (Join-Path $TestDrive 'two.tar.gz') `
                -ObservationReportPath (Join-Path $TestDrive 'status.json')
        }
        catch {
            $message = $_.Exception.Message
        }
        $message | Should Match 'lowercase full SHA'
    }
}
