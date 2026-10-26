Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1')
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestRunSummary.psm1')
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1')
$sandboxModule = Get-Module TestSandbox
$testRoot = Get-RSTestSandboxRoot -RepoRoot $repoRoot

Describe 'Verified same-volume run evidence archival' {
    BeforeEach { $script:fixture = New-RSTestRunArchiveFixture -TestRoot $testRoot }
    AfterEach {
        # Only this test's synthetic source fixture; never clean archived evidence or other runs.
        if ([IO.Directory]::Exists($fixture.Source)) {
            Remove-RSTestSandboxExactRunDirectory -TestRoot $fixture.Root -RunId $fixture.Id
        }
    }

    It 'preserves hidden readonly content, named streams, empty directories and security' {
        [void][IO.Directory]::CreateDirectory((Join-Path $fixture.Source 'empty'))
        Set-Content -LiteralPath $fixture.Payload -Stream witness -Value 'unique stream' -NoNewline
        [IO.File]::SetAttributes($fixture.Payload, [IO.FileAttributes]'Hidden, ReadOnly')
        $hash = (Get-FileHash -LiteralPath $fixture.Payload).Hash
        $acl = (Get-Acl -LiteralPath $fixture.Payload).Sddl
        $result = Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id
        $result.Status | Should Be 'Archived'
        (Test-Path -LiteralPath $fixture.Source) | Should Be $false
        $payload = Join-Path $result.RunRoot 'unique.txt'
        (Get-FileHash -LiteralPath $payload).Hash | Should Be $hash
        (Get-Content -LiteralPath $payload -Stream witness -Raw) | Should Be 'unique stream'
        (Get-Acl -LiteralPath $payload).Sddl | Should Be $acl
        ((Get-Item -LiteralPath $payload -Force).Attributes -band [IO.FileAttributes]::ReadOnly) | Should Be 1
        (Test-Path -LiteralPath (Join-Path $result.RunRoot 'empty')) | Should Be $true
        (Test-Path -LiteralPath $result.RecordPath) | Should Be $true
    }

    It 'never merges with or overwrites an existing destination' {
        [void][IO.Directory]::CreateDirectory($fixture.Destination)
        [IO.File]::WriteAllText((Join-Path $fixture.Destination 'existing.txt'), 'keep')
        { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id } | Should Throw 'destination already exists'
        (Get-Content -LiteralPath $fixture.Payload -Raw) | Should Be 'synthetic archive regression evidence'
        (Get-Content -LiteralPath (Join-Path $fixture.Destination 'existing.txt') -Raw) | Should Be 'keep'
    }

    It 'preserves self-contained relative symlinks without traversing them' {
        $link = Join-Path $fixture.Source 'relative.txt'
        New-RSTestSandboxLink -ItemType SymbolicLink -Path $link -Target 'unique.txt' | Out-Null
        $result = Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id
        (Get-Item -LiteralPath (Join-Path $result.RunRoot 'relative.txt')).LinkTarget | Should Be 'unique.txt'
        (Get-Content -LiteralPath (Join-Path $result.RunRoot 'relative.txt') -Raw) | Should Be 'synthetic archive regression evidence'
    }

    It 'keeps read-only shared hardlink payloads read-only at both paths' {
        [void][IO.Directory]::CreateDirectory($fixture.Records)
        $shared = Join-Path $fixture.Records 'shared-payload.txt'
        [IO.File]::WriteAllText($shared, 'synthetic shared immutable payload')
        New-RSTestSandboxLink -ItemType HardLink -Path (Join-Path $fixture.Source 'shared.txt') -Target $shared | Out-Null
        [IO.File]::SetAttributes($shared, [IO.FileAttributes]::ReadOnly)
        $result = Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id
        foreach ($path in @($shared, (Join-Path $result.RunRoot 'shared.txt'))) {
            ((Get-Item -LiteralPath $path -Force).Attributes -band [IO.FileAttributes]::ReadOnly) | Should Be 1
            (Get-Content -LiteralPath $path -Raw) | Should Be 'synthetic shared immutable payload'
        }
    }

    It 'archives Unicode and long paths with the same bytes and stream metadata' {
        $directory = $fixture.Source
        foreach ($index in 1..7) { $directory = Join-Path $directory ('évidence-' + ('x' * 32) + $index) }
        [void][IO.Directory]::CreateDirectory($directory)
        $payload = Join-Path $directory 'witness.txt'
        [IO.File]::WriteAllText($payload, 'long path evidence')
        $relative = [IO.Path]::GetRelativePath($fixture.Source, $payload)
        $result = Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id
        (Get-Content -LiteralPath (Join-Path $result.RunRoot $relative) -Raw) | Should Be 'long path evidence'
    }

    It 'retains a locked tree then recovers the flushed intent without a payload copy' {
        $lease = [IO.FileStream]::new($fixture.Payload, 'Open', 'Read', 'Read')
        try {
            { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id } | Should Throw 'Exception calling'
            (Test-Path -LiteralPath $fixture.Payload) | Should Be $true
            (Test-Path -LiteralPath (Join-Path $fixture.Records 'intent.json')) | Should Be $true
        } finally { $lease.Dispose() }
        $result = Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id -Disposition interrupted
        (Test-Path -LiteralPath $result.RecordPath) | Should Be $true
    }

    It 'recovers a crash after rename but before completion publication' {
        $inventory = & $sandboxModule { param($Path) @(Get-RSTestRunArchiveInventory -RunRoot $Path) } $fixture.Source
        [void][IO.Directory]::CreateDirectory($fixture.Records)
        $intent = [ordered]@{
            schema = 'red-salamander.run-archive.v1'; run_id = $fixture.Id
            source = $fixture.Source; destination = $fixture.Destination; disposition = 'interrupted'
            inventory = @($inventory)
        }
        $intent | ConvertTo-Json -Depth 16 -Compress | Set-Content -LiteralPath (Join-Path $fixture.Records 'intent.json')
        [void][IO.Directory]::CreateDirectory((Split-Path $fixture.Destination))
        [IO.Directory]::Move($fixture.Source, $fixture.Destination)
        $result = Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id
        (Test-Path -LiteralPath $result.RecordPath) | Should Be $true
    }

    It 'retains changed data after interrupted publication rather than silently rebaselining' {
        $lease = [IO.FileStream]::new($fixture.Payload, 'Open', 'Read', 'Read')
        try { try { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id } catch { } }
        finally { $lease.Dispose() }
        [IO.File]::WriteAllText($fixture.Payload, 'changed unique witness')
        { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id } | Should Throw 'source changed'
        (Get-Content -LiteralPath $fixture.Payload -Raw) | Should Be 'changed unique witness'
    }

    It 'protects live owners and unparseable manual folders during abandoned archival' {
        { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id -Abandoned } | Should Throw 'live, allowed, or unparseable'
        (Test-Path -LiteralPath $fixture.Payload) | Should Be $true
        (& $sandboxModule { param($Id) (ConvertFrom-RSTestRunId -RunId $Id).ProcessId } "adhoc-$PID") | Should Be $PID
    }

    It 'preserves an abandoned native fallback run after its real owner exits' {
        $process = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList @('-NoProfile', '-Command', 'exit 0') -WindowStyle Hidden -PassThru -Wait
        try {
            $abandonedId = "adhoc-$($process.Id)"
            $abandonedSource = Join-Path $fixture.Root "runs\$abandonedId"
            (Test-Path -LiteralPath $abandonedSource) | Should Be $false
            [IO.Directory]::Move($fixture.Source, $abandonedSource)
            $fixture.Id = $abandonedId; $fixture.Source = $abandonedSource
            $result = Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id -Abandoned -Disposition interrupted
            (Get-Content -LiteralPath (Join-Path $result.RunRoot 'unique.txt') -Raw) | Should Be 'synthetic archive regression evidence'
        } finally { $process.Dispose() }
    }

    It 'retains a manual unparseable owner instead of treating it as abandoned' {
        $manualId = 'manual-archive-' + [guid]::NewGuid().ToString('N')
        $manualSource = Join-Path $fixture.Root "runs\$manualId"
        [IO.Directory]::Move($fixture.Source, $manualSource)
        $fixture.Id = $manualId; $fixture.Source = $manualSource
        { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id -Abandoned } | Should Throw 'unparseable'
    }

    It 'refuses unsafe junctions without reading or changing the target' {
        $target = Join-Path $fixture.Root 'evidence\archive-test-outside-target'
        [void](Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $target -RequireDescendant)
        [void][IO.Directory]::CreateDirectory($target)
        $link = Join-Path $fixture.Source 'outside'
        New-RSTestSandboxLink -ItemType Junction -Path $link -Target $target | Out-Null
        try {
            { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id } | Should Throw 'reparse point'
            (Test-Path -LiteralPath $target) | Should Be $true
        } finally { [IO.Directory]::Delete($link) }
    }

    It 'refuses reparse ancestors of the archive destination' {
        $link = Join-Path $fixture.Root "evidence\completed-runs\$($fixture.Id)"
        New-RSTestSandboxLink -ItemType Junction -Path $link -Target $fixture.Source | Out-Null
        try { { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id } | Should Throw 'reparse-point component' }
        finally { [IO.Directory]::Delete($link) }
    }

    It 'retains malformed disk receipts without invoking destructive disk cleanup' {
        $receiptDirectory = Join-Path $fixture.Source 'artifacts\test-owned-vhdx'
        [void][IO.Directory]::CreateDirectory($receiptDirectory)
        [IO.File]::WriteAllText((Join-Path $receiptDirectory 'test.receipt.json'), '{"schema":"unreviewed"}')
        try {
            { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id } | Should Throw 'Malformed test-owned'
            (Test-Path -LiteralPath $fixture.Payload) | Should Be $true
        } finally { [IO.File]::Delete((Join-Path $receiptDirectory 'test.receipt.json')) }
    }

    It 'retains a valid unresolved mounted-disk receipt without invoking storage mutation' {
        $receiptPath = & $sandboxModule {
            param($RunRoot)
            $layout = Get-RSTestOwnedVirtualDiskLayout -RunRoot $RunRoot -Name 'archive-fixture'
            Write-RSTestOwnedVirtualDiskReceipt -Layout $layout -Stage mounted -VolumeGuidPath ''
            $layout.ReceiptPath
        } $fixture.Source
        try {
            { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id } | Should Throw 'unresolved test-owned VHDX'
            (Test-Path -LiteralPath $receiptPath) | Should Be $true
        } finally { [IO.File]::Delete($receiptPath) }
    }

    It 'fails closed on absent sources and unauthorized paths' {
        { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId '..' } | Should Throw 'single filesystem path segment'
        { Save-RSTestSandboxRunEvidence -TestRoot (Join-Path $fixture.Root 'wrong') -RunId $fixture.Id } | Should Throw 'Unauthorized TestSandbox root'
        { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId ([guid]::NewGuid().ToString('N')) } | Should Throw 'source is absent'
    }

    It 'serializes the same archive transaction without excluding other runs' {
        [void][IO.Directory]::CreateDirectory($fixture.Records)
        $lease = [IO.FileStream]::new((Join-Path $fixture.Records 'archive.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
        try {
            { Save-RSTestSandboxRunEvidence -TestRoot $fixture.Root -RunId $fixture.Id } | Should Throw 'being used by another process'
            $other = New-RSTestRunArchiveFixture -TestRoot $testRoot
            (Save-RSTestSandboxRunEvidence -TestRoot $other.Root -RunId $other.Id).Status | Should Be 'Archived'
        } finally { $lease.Dispose() }
    }
}

Describe 'Archived runner navigation' {
    It 'initializes an explicitly authorized root before allocating a fresh run and refuses run reuse' {
        $source = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Run-AllTests.ps1') -Raw
        $start = $source.IndexOf('$testRunContext = New-RSTestRunContext', [StringComparison]::Ordinal)
        $end = $source.IndexOf('$env:REDSALAMANDER_TEST_ROOT =', $start, [StringComparison]::Ordinal)
        $body = [scriptblock]::Create($source.Substring($start, $end - $start))
        $directory = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness archive-startup -Case ([guid]::NewGuid().ToString('N'))
        & {
            param($Body, $Directory)
            $state = [pscustomobject]@{ Initialized = $false }
            $fixtureContext = [pscustomobject]@{ TestRoot = $testRoot; RunRoot = Join-Path $Directory 'owned-run' }
            function New-RSTestRunContext { return $fixtureContext }
            function Initialize-RSTestSandboxRoot { $state.Initialized = $true }
            function Assert-RSTestSandboxContainedPath {
                if (-not $state.Initialized) { throw 'Containment requested before explicit initialization' }
            }
            $RunId = ''; $selectedTestRoot = $testRoot; $AllowExternalTestRoot = $true
            $runnerOwnsRunDirectory = $false
            & $Body
            # Ownership assignment is local to the invoked scriptblock; actual creation proves the sequence.
            [IO.Directory]::Exists($fixtureContext.RunRoot) | Should Be $true
            { & $Body } | Should Throw 'Refusing to reuse an existing run directory'
        } $body $directory
    }

    It 'rebases suite and retry log links while retaining original failure provenance' {
        $summary = New-RSTestRunSummary -Suite Compare -Platform x64 -Configuration Debug -ExePath 'Z:\src\app.exe' `
            -ArtifactRoot 'Z:\RedSalamander.Perf\evidence\completed-runs\id\artifacts' `
            -OriginalArtifactRoot 'Z:\RedSalamander.Perf\runs\id\artifacts' -RunStartedUtc ([datetime]::UtcNow) -RunEndedUtc ([datetime]::UtcNow) `
            -Results @([pscustomobject]@{ Name = 'Compare'; OutputLogPath = 'Z:\RedSalamander.Perf\runs\id\artifacts\Compare.log'; RetryAttempts = @([pscustomobject]@{ OutputLogPath = 'Z:\RedSalamander.Perf\runs\id\artifacts\retry.log' }); Cases = @() })
        $summary.suites[0].output_log_path | Should Match '\\evidence\\completed-runs\\'
        $summary.suites[0].retry_attempts[0].output_log_path | Should Match '\\evidence\\completed-runs\\'
    }
}
