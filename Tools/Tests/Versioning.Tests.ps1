Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperModule = Join-Path $repoRoot 'Tools\Modules\Build\Versioning.psm1'

function New-RSTestVersionRepo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    $commonDir = Join-Path $Root 'Common'
    New-Item -ItemType Directory -Path $commonDir -Force | Out-Null

    @'
// WARNING: minimal test header
#define VERSINFO_MAJOR 7
#define VERSINFO_MINOR 0
'@ | Set-Content -Path (Join-Path $commonDir 'Version.h') -Encoding ASCII
}

Describe 'Versioning helper' {
    BeforeAll {
        Import-Module $helperModule -Force -ErrorAction Stop
        # Resolve-RSVersionContext prefers GITHUB_RUN_NUMBER over the local per-worktree counter.
        # These contracts describe the local counter, so a hosted runner's run number must not
        # leak into them (the CI lane saw 'Expected 64 but was 63', the run number).
        $script:savedRunNumber = [Environment]::GetEnvironmentVariable('GITHUB_RUN_NUMBER', 'Process')
        [Environment]::SetEnvironmentVariable('GITHUB_RUN_NUMBER', $null, 'Process')
    }

    AfterAll {
        [Environment]::SetEnvironmentVariable('GITHUB_RUN_NUMBER', $script:savedRunNumber, 'Process')
    }

    It 'reuses the saved local build number across ordinary local builds' {
        $testRepo = Join-Path $TestDrive 'ReuseSavedLocalBuildNumber'
        New-RSTestVersionRepo -Root $testRepo

        $firstContext = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64

        $statePath = Get-RSVersionStatePath -RepoRoot $testRepo
        (Get-Item $statePath).LastWriteTimeUtc = [DateTime]::UtcNow.AddDays(-1)

        $secondContext = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Release -Platform ARM64

        $secondContext.BuildNumber | Should Be $firstContext.BuildNumber
        (Read-RSVersionContext -RepoRoot $testRepo).Platform | Should Be 'ARM64'
    }

    It 'allocates a new local build number when the saved context is missing' {
        $testRepo = Join-Path $TestDrive 'MissingSavedContextGetsNextNumber'
        New-RSTestVersionRepo -Root $testRepo

        $firstContext = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64

        Remove-Item -LiteralPath (Get-RSVersionStatePath -RepoRoot $testRepo) -Force

        $secondContext = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64

        $secondContext.BuildNumber | Should Be ($firstContext.BuildNumber + 1)
    }

    It 'constructs explicit package metadata without creating version state' {
        $testRepo = Join-Path $TestDrive 'MajorMinorBuildPackageVersion'
        New-RSTestVersionRepo -Root $testRepo

        $context = New-RSVersionContext -RepoRoot $testRepo -Configuration Release -Platform x64 -BuildNumber 184 -OfficialRelease

        $context.DisplayBaseVersion | Should Be '7.0'
        $context.DisplayVersion | Should Be '7.0.184'
        $context.PackagingVersion | Should Be '7.0.184'
        (Test-Path -LiteralPath (Join-Path $testRepo '.build\version')) | Should Be $false
    }

    It 'uses GITHUB_RUN_NUMBER as the CI build number' {
        $testRepo = Join-Path $TestDrive 'GithubRunNumberBuild'
        New-RSTestVersionRepo -Root $testRepo

        $oldRunNumber = $env:GITHUB_RUN_NUMBER
        try {
            $env:GITHUB_RUN_NUMBER = '185'
            $context = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Release -Platform x64 -OfficialRelease
        }
        finally {
            $env:GITHUB_RUN_NUMBER = $oldRunNumber
        }

        $context.BuildNumber | Should Be 185
        $context.PackagingVersion | Should Be '7.0.185'
    }

    It 'publishes valid JSON atomically without leaving temporary files' {
        $testRepo = Join-Path $TestDrive 'AtomicStatePublication'
        New-RSTestVersionRepo -Root $testRepo

        $context = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64
        $statePath = Get-RSVersionStatePath -RepoRoot $testRepo
        $saved = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json

        $saved.BuildNumber | Should Be $context.BuildNumber
        @((Get-ChildItem -LiteralPath (Split-Path -Parent $statePath) -Filter '*.tmp')).Count | Should Be 0
    }

    It 'replaces the previous state atomically and removes scratch and backup files' {
        $testRepo = Join-Path $TestDrive 'AtomicReplacement'
        New-RSTestVersionRepo -Root $testRepo
        $first = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64
        $statePath = Get-RSVersionStatePath -RepoRoot $testRepo
        $before = [IO.File]::ReadAllBytes($statePath)
        $second = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Release -Platform x64 -BuildNumber 19

        $second.BuildNumber | Should Be 19
        ([Convert]::ToBase64String([IO.File]::ReadAllBytes($statePath)) -ne [Convert]::ToBase64String($before)) | Should Be $true
        @((Get-ChildItem -LiteralPath (Split-Path -Parent $statePath) -Filter '*.tmp*')).Count | Should Be 0
    }

    It 'preserves the previous state and removes scratch files when atomic replacement fails' {
        $testRepo = Join-Path $TestDrive 'AtomicReplacementFailure'
        New-RSTestVersionRepo -Root $testRepo
        Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64 | Out-Null
        $statePath = Get-RSVersionStatePath -RepoRoot $testRepo
        $before = [Convert]::ToBase64String([IO.File]::ReadAllBytes($statePath))
        Mock -CommandName Invoke-RSVersionFileReplace -ModuleName Versioning -MockWith {
            throw 'Injected atomic replacement failure.'
        }

        $message = $null
        try { $null = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Release -Platform x64 -BuildNumber 19 }
        catch { $message = $_.Exception.Message }

        $message | Should Match 'Injected atomic replacement failure'
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($statePath)) | Should Be $before
        @((Get-ChildItem -LiteralPath (Split-Path -Parent $statePath) -Filter '*.tmp*')).Count | Should Be 0
    }

    It 'rejects malformed persisted state and invalid local counters instead of resetting them' {
        $testRepo = Join-Path $TestDrive 'InvalidState'
        New-RSTestVersionRepo -Root $testRepo
        $statePath = Get-RSVersionStatePath -RepoRoot $testRepo
        New-Item -ItemType Directory -Path (Split-Path -Parent $statePath) -Force | Out-Null
        [IO.File]::WriteAllText($statePath, '{not-json', [Text.UTF8Encoding]::new($false))

        $readMessage = $null
        try { $null = Read-RSVersionContext -RepoRoot $testRepo }
        catch { $readMessage = $_.Exception.Message }
        ($null -ne $readMessage) | Should Be $true
        $readMessage | Should Match 'Invalid version state'
        Remove-Item -LiteralPath $statePath -Force

        $counterPath = Join-Path (Split-Path -Parent $statePath) 'local-build-counter.txt'
        foreach ($invalidCounter in @('', '0', '-1', 'not-a-number', '65535')) {
            [IO.File]::WriteAllText($counterPath, $invalidCounter, [Text.Encoding]::ASCII)
            $counterMessage = $null
            try { $null = Resolve-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64 }
            catch { $counterMessage = $_.Exception.Message }
            ($null -ne $counterMessage) | Should Be $true
            $counterMessage | Should Match 'Invalid local build counter'
            [IO.File]::ReadAllText($counterPath) | Should Be $invalidCounter
        }
    }

    It 'rejects valid JSON with wrong primitive types or unsupported profile values' {
        $testRepo = Join-Path $TestDrive 'InvalidTypedState'
        New-RSTestVersionRepo -Root $testRepo
        $statePath = Get-RSVersionStatePath -RepoRoot $testRepo
        New-Item -ItemType Directory -Path (Split-Path -Parent $statePath) -Force | Out-Null
        $valid = [ordered]@{
            Major = 7
            Minor = 0
            BuildNumber = 1
            Configuration = 'Debug'
            Platform = 'x64'
            OfficialRelease = $false
        }
        $invalidStates = @(
            @{ Major = '7' },
            @{ Minor = -1 },
            @{ BuildNumber = '1' },
            @{ Configuration = 'Profile' },
            @{ Platform = 'arm64' },
            @{ OfficialRelease = 'false' }
        )

        foreach ($change in $invalidStates) {
            $candidate = [ordered]@{}
            foreach ($key in $valid.Keys) { $candidate[$key] = $valid[$key] }
            foreach ($key in $change.Keys) { $candidate[$key] = $change[$key] }
            [IO.File]::WriteAllText(
                $statePath,
                ($candidate | ConvertTo-Json -Compress),
                [Text.UTF8Encoding]::new($false))
            $message = $null
            try { $null = Read-RSVersionContext -RepoRoot $testRepo }
            catch { $message = $_.Exception.Message }
            $message | Should Match 'invalid type or value'
        }
    }

    It 'serializes concurrent resolvers to one valid persisted build identity' {
        $testRepo = Join-Path $TestDrive 'ConcurrentResolvers'
        New-RSTestVersionRepo -Root $testRepo
        $childScript = Join-Path $TestDrive 'resolve-version-child.ps1'
        @'
param([string]$ModulePath, [string]$RepoRoot, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
Import-Module $ModulePath -Force
$context = Resolve-RSVersionContext -RepoRoot $RepoRoot -Configuration Debug -Platform x64
[IO.File]::WriteAllText($ResultPath, [string]$context.BuildNumber, [Text.Encoding]::ASCII)
'@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $oldRunNumber = $env:GITHUB_RUN_NUMBER
        $env:GITHUB_RUN_NUMBER = $null
        $processes = @()
        try {
            foreach ($index in 1..4) {
                $resultPath = Join-Path $TestDrive "concurrent-result-$index.txt"
                $processes += [pscustomobject]@{
                    ResultPath = $resultPath
                    Process = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList @(
                        '-NoLogo', '-NoProfile', '-File', $childScript, $helperModule, $testRepo, $resultPath
                    ) -PassThru
                }
            }
            foreach ($entry in $processes) {
                $entry.Process.WaitForExit(30000) | Should Be $true
                $entry.Process.ExitCode | Should Be 0
            }
        }
        finally {
            $env:GITHUB_RUN_NUMBER = $oldRunNumber
            foreach ($entry in $processes) {
                if (-not $entry.Process.HasExited) { $entry.Process.Kill() }
                $entry.Process.Dispose()
            }
        }

        $numbers = @($processes | ForEach-Object { [int](Get-Content -LiteralPath $_.ResultPath -Raw) })
        @($numbers | Sort-Object -Unique).Count | Should Be 1
        (Read-RSVersionContext -RepoRoot $testRepo).BuildNumber | Should Be $numbers[0]
    }

    It 'waits for a repository file lease and recovers automatically after its owner crashes' {
        $testRepo = Join-Path $TestDrive 'CrashedFileLease'
        New-RSTestVersionRepo -Root $testRepo
        $stateDirectory = Join-Path $testRepo '.build\version'
        New-Item -ItemType Directory -Path $stateDirectory -Force | Out-Null
        $lockPath = Join-Path $stateDirectory 'version-state.lock'
        $readyPath = Join-Path $TestDrive 'version-lock-ready.txt'
        $releasePath = Join-Path $TestDrive 'version-lock-release.txt'
        $holderScript = Join-Path $TestDrive 'hold-version-file-lock.ps1'
        @'
param([string]$LockPath, [string]$ReadyPath, [string]$ReleasePath)
$stream = [IO.File]::Open($LockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
[IO.File]::WriteAllText($ReadyPath, 'ready')
while (-not [IO.File]::Exists($ReleasePath)) { Start-Sleep -Milliseconds 20 }
[Environment]::Exit(0)
'@ | Set-Content -LiteralPath $holderScript -Encoding UTF8
        $resolverScript = Join-Path $TestDrive 'wait-for-version-file-lock.ps1'
        $resultPath = Join-Path $TestDrive 'version-lock-result.txt'
        @'
param([string]$ModulePath, [string]$RepoRoot, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
Import-Module $ModulePath -Force
$context = Resolve-RSVersionContext -RepoRoot $RepoRoot -Configuration Debug -Platform x64
[IO.File]::WriteAllText($ResultPath, [string]$context.BuildNumber, [Text.Encoding]::ASCII)
'@ | Set-Content -LiteralPath $resolverScript -Encoding UTF8
        $holder = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList @(
            '-NoLogo', '-NoProfile', '-File', $holderScript, $lockPath, $readyPath, $releasePath
        ) -PassThru
        $resolver = $null
        try {
            $deadline = [DateTime]::UtcNow.AddSeconds(10)
            while (-not (Test-Path -LiteralPath $readyPath) -and [DateTime]::UtcNow -lt $deadline) {
                Start-Sleep -Milliseconds 20
            }
            (Test-Path -LiteralPath $readyPath) | Should Be $true

            $resolver = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList @(
                '-NoLogo', '-NoProfile', '-File', $resolverScript, $helperModule, $testRepo, $resultPath
            ) -PassThru
            Start-Sleep -Milliseconds 300
            $resolver.HasExited | Should Be $false

            [IO.File]::WriteAllText($releasePath, 'crash')
            $holder.WaitForExit(10000) | Should Be $true
            $resolver.WaitForExit(30000) | Should Be $true
            $holder.ExitCode | Should Be 0
            $resolver.ExitCode | Should Be 0
            [int](Get-Content -LiteralPath $resultPath -Raw) | Should Be 1
        }
        finally {
            if (-not $holder.HasExited) { $holder.Kill() }
            $holder.Dispose()
            if ($null -ne $resolver) {
                if (-not $resolver.HasExited) { $resolver.Kill() }
                $resolver.Dispose()
            }
        }
    }

    It 'exports only semantic state operations' {
        $commands = @(Get-Command -Module Versioning | Select-Object -ExpandProperty Name)
        ($commands -contains 'New-RSVersionContext') | Should Be $true
        ($commands -contains 'Resolve-RSVersionContext') | Should Be $true
        ($commands -contains 'Read-RSVersionContext') | Should Be $true
        ($commands -contains 'Use-RSVersionStateLock') | Should Be $false
        ($commands -contains 'Save-RSVersionContext') | Should Be $false
        $moduleSource = Get-Content -LiteralPath $helperModule -Raw
        $moduleSource | Should Match 'version-state\.lock'
        $moduleSource | Should Match '\[IO\.FileShare\]::None'
        $moduleSource | Should Not Match 'Local\\RedSalamander\.Versioning'
    }
}
