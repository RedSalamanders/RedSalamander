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
typedef enum {
  GHOSTTY_TERMINAL_OPT_A = 1,
} GhosttyTerminalOption;
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
typedef enum {
  GHOSTTY_TERMINAL_OPT_A = 1,
  GHOSTTY_TERMINAL_OPT_B = 2,
} GhosttyTerminalOption;
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

    It 'does not repeat historical findings outside the reviewed commit range' {
        $fixture = New-RSGhosttyCandidateReviewRepository
        $security = Get-RSGhosttyCandidateSecurityReview -SourcePath $fixture.Root -BaseCommit $fixture.Base -CandidateCommit $fixture.Candidate
        $security.reviewedCommitCount | Should Be 1
        @($security.findings).Count | Should Be 0
        $schema = Get-Content -LiteralPath $reviewSchemaPath -Raw | ConvertFrom-Json -Depth 100
        $schema.properties.security.properties.findings.minItems | Should Be 0
    }

    It 'freezes all dependency hashes separately and binds a changed dependency into the descriptor digest' {
        $fixture = New-RSGhosttyCandidateReviewRepository
        $owner = Join-Path $TestDrive 'review-owner'
        & git.exe clone --quiet --no-hardlinks $fixture.Root $owner
        if ($LASTEXITCODE -ne 0) { throw 'Fixture owner clone failed.' }
        $candidateRoot = Join-Path $owner ".build/TerminalEngineUpgrade/$($fixture.Candidate)"
        [void](New-Item -ItemType Directory -Path $candidateRoot -Force)
        $source = Join-Path $candidateRoot 'source-review'
        Move-Item -LiteralPath $fixture.Root -Destination $source
        Invoke-RSGhosttyCandidateTestGit -Root $source -Arguments @('remote', 'add', 'origin', 'https://github.com/ghostty-org/ghostty.git') | Out-Null
        Invoke-RSGhosttyCandidateTestGit -Root $source -Arguments @('checkout', '--detach', $fixture.Candidate) | Out-Null
        $tree = Invoke-RSGhosttyCandidateTestGit -Root $source -Arguments @('rev-parse', 'HEAD^{tree}') | Select-Object -First 1

        $schemas = Join-Path $owner 'Specs/Terminal'
        [void](New-Item -ItemType Directory -Path $schemas -Force)
        foreach ($schema in @($reviewSchemaPath, $descriptorSchemaPath)) {
            Copy-Item -LiteralPath $schema -Destination $schemas
        }
        $lockDirectory = Join-Path $owner 'External/TerminalEngine'
        [void](New-Item -ItemType Directory -Path $lockDirectory -Force)
        $lock = Get-Content (Join-Path $repoRoot 'External/TerminalEngine/GhosttyRuntimeLock.v1.json') -Raw | ConvertFrom-Json -Depth 100
        $lock.upstream.commit = $fixture.Base
        $lock.upstream.tree = Invoke-RSGhosttyCandidateTestGit -Root $source -Arguments @('rev-parse', "$($fixture.Base)^{tree}") | Select-Object -First 1
        $lock | ConvertTo-Json -Depth 100 | Set-Content (Join-Path $lockDirectory 'GhosttyRuntimeLock.v1.json')
        $observation = [pscustomobject]@{ mode = 'Candidate'; observation = [pscustomobject]@{
                candidateRequestedCommit = $fixture.Candidate; commit = $fixture.Candidate; tree = $tree
            } }
        & (Get-Module GhosttyCandidateReview) {
            param($fixtureLock, $fixtureObservation)
            $script:descriptorTestLock = $fixtureLock
            $script:descriptorTestObservation = $fixtureObservation
        } $lock $observation
        Mock -ModuleName GhosttyCandidateReview Read-RSGhosttyRuntimeLock { return $script:descriptorTestLock }
        Mock -ModuleName GhosttyCandidateReview Read-RSValidationJsonFile { return $script:descriptorTestObservation }
        $archive = Join-Path $candidateRoot 'one.tar.gz'
        $repeat = Join-Path $candidateRoot 'two.tar.gz'
        'same archive fixture' | Set-Content $archive
        Copy-Item -LiteralPath $archive -Destination $repeat
        $observationPath = Join-Path $candidateRoot 'observation.json'
        '{}' | Set-Content $observationPath
        $verificationPath = Join-Path $candidateRoot 'verification.json'
        @{
            sha = $fixture.Candidate
            html_url = "https://github.com/ghostty-org/ghostty/commit/$($fixture.Candidate)"
            commit = @{
                tree = @{ sha = $tree }
                verification = @{ verified = $true; reason = 'valid' }
                author = @{ date = '2026-10-28T00:00:00Z' }
                committer = @{ date = '2026-10-28T00:00:00Z' }
            }
        } | ConvertTo-Json -Depth 8 | Set-Content $verificationPath
        $arguments = @{
            RepoRoot = $owner; CandidateCommit = $fixture.Candidate; SourcePath = $source
            ArchivePath = $archive; RepeatArchivePath = $repeat; ObservationReportPath = $observationPath
            VerificationFixturePath = $verificationPath; GeneratedUtc = [DateTime]'2026-10-28T00:00:00Z'
        }
        $result = New-RSGhosttyRuntimeCandidateReview @arguments
        $first = Get-Content $result.DescriptorPath -Raw | ConvertFrom-Json -Depth 100
        $first.schemaVersion | Should Be 2
        @($first.dependencies).Count | Should Be $lock.dependencies.Count
        $lock.dependencies.Count | Should BeGreaterThan 1
        for ($i = 0; $i -lt $lock.dependencies.Count; $i++) {
            $first.dependencies[$i].id | Should Be $lock.dependencies[$i].id
            $first.dependencies[$i].sourceArchiveSha256 | Should Be $lock.dependencies[$i].sourceArchive.sha256
            $first.dependencies[$i].contentArchiveSha256 | Should Be $lock.dependencies[$i].contentArchive.sha256
            $first.dependencies[$i].licenseSha256 | Should Be $lock.dependencies[$i].licenseSha256
        }
        $lock.dependencies[1].sourceArchive.sha256 = 'e' * 64
        $result = New-RSGhosttyRuntimeCandidateReview @arguments
        $second = Get-Content $result.DescriptorPath -Raw | ConvertFrom-Json -Depth 100
        $second.dependencies[1].sourceArchiveSha256 | Should Be ('e' * 64)
        $second.inputDigest | Should Not Be $first.inputDigest
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
