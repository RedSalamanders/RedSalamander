$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyRuntimeLifecycle.psm1'
$lockPath = Join-Path $repoRoot 'External\TerminalEngine\GhosttyRuntimeLock.v1.json'
$schemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeLock.schema.json'

Import-Module $modulePath -Force

function Copy-RSGhosttyRuntimeLockFixture {
    return Get-Content -LiteralPath $lockPath -Raw |
        ConvertFrom-Json -Depth 100 -NoEnumerate -DateKind String
}

function New-RSGhosttyLockedPackageFixture {
    # A nested canonical root covers transports whose wrapper directory differs
    # from the final Zig hash-named package directory.
    $owner = Join-Path $TestDrive ('packages-' + [guid]::NewGuid().ToString('N'))
    $source = Join-Path $owner '.build/source/fixture-pkg/wrapper'
    $downloads = Join-Path $owner '.build/downloads'
    [void](New-Item -ItemType Directory -Path $source, $downloads -Force)
    'Synthetic package license' | Set-Content (Join-Path $source 'LICENSE') -Encoding utf8NoBOM
    '.{ .name = .fixture, .version = "0.0.0" }' | Set-Content (Join-Path $source 'build.zig.zon') -Encoding utf8NoBOM
    $archive = Join-Path $downloads 'fixture-pkg.tar.gz'
    & tar.exe -czf $archive -C (Join-Path $owner '.build/source') fixture-pkg/wrapper/LICENSE fixture-pkg/wrapper/build.zig.zon
    if ($LASTEXITCODE -ne 0) { throw 'Canonical package fixture creation failed.' }
    $lock = [pscustomobject]@{ dependencies = @([pscustomobject]@{
                id='fixture'; packageName='fixture-pkg'; licenseArchiveRelativePath='LICENSE'
                contentArchive=[pscustomobject]@{ fileName='fixture-pkg.tar.gz'; internalRoot='fixture-pkg/wrapper'
                    sha256=(Get-FileHash $archive -Algorithm SHA256).Hash.ToLowerInvariant() }
            }) }
    return [pscustomobject]@{
        Owner=$owner; Source=$source; Downloads=$downloads; Lock=$lock
        Destination=(Join-Path $owner '.build/closed-packages')
    }
}

Describe 'Ghostty closed package directory' {
    It 'stages verified nested archive contents directly beneath their Zig package name' {
        $fixture = New-RSGhosttyLockedPackageFixture
        $root = Expand-RSGhosttyRuntimeLockedPackages -RepoRoot $fixture.Owner -RuntimeLock $fixture.Lock `
            -DownloadsRoot $fixture.Downloads -DestinationRoot $fixture.Destination
        (Get-Content (Join-Path $root 'fixture-pkg/LICENSE') -Raw) |
            Should Be (Get-Content (Join-Path $fixture.Source 'LICENSE') -Raw)
        Test-Path (Join-Path $root 'fixture-pkg/build.zig.zon') | Should Be $true
        Test-Path (Join-Path $root 'fixture-pkg/wrapper') | Should Be $false
    }

    It 'rejects changed archives before package publication' {
        $fixture = New-RSGhosttyLockedPackageFixture
        Add-Content (Join-Path $fixture.Downloads 'fixture-pkg.tar.gz') -Value 'changed'
        { Expand-RSGhosttyRuntimeLockedPackages -RepoRoot $fixture.Owner -RuntimeLock $fixture.Lock `
                -DownloadsRoot $fixture.Downloads -DestinationRoot $fixture.Destination } | Should Throw 'SHA-256 mismatch'
        Test-Path (Join-Path $fixture.Destination 'fixture-pkg') | Should Be $false
    }

    It 'preserves occupied destinations and rejects an escaped package name' {
        $fixture = New-RSGhosttyLockedPackageFixture
        [void](New-Item -ItemType Directory -Path $fixture.Destination)
        'preserve' | Set-Content (Join-Path $fixture.Destination 'owner.txt')
        { Expand-RSGhosttyRuntimeLockedPackages -RepoRoot $fixture.Owner -RuntimeLock $fixture.Lock `
                -DownloadsRoot $fixture.Downloads -DestinationRoot $fixture.Destination } | Should Throw 'absent or empty'
        Get-Content (Join-Path $fixture.Destination 'owner.txt') | Should Be 'preserve'
        $fixture.Lock.dependencies[0].packageName = '../escape'
        { Expand-RSGhosttyRuntimeLockedPackages -RepoRoot $fixture.Owner -RuntimeLock $fixture.Lock `
                -DownloadsRoot $fixture.Downloads -DestinationRoot (Join-Path $fixture.Owner '.build/other-packages') } |
            Should Throw 'package name escapes'
    }
}

Describe 'Ghostty native command output contract' {
    It 'keeps both captured streams separate from process progress' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1')
        $tokens = $null
        $parseErrors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($modulePath, [ref]$tokens, [ref]$parseErrors)
        @($parseErrors).Count | Should Be 0
        $commandFunction = $ast.Find({ param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -ceq 'Invoke-CheckedProcess'
        }, $true)
        $invoke = [scriptblock]::Create(
            'param($FilePath, $Arguments, $WorkingDirectory)' + "`n" +
            $commandFunction.Extent.Text + "`n" +
            'Invoke-CheckedProcess -FilePath $FilePath -ArgumentList $Arguments -WorkingDirectory $WorkingDirectory')
        $progressLog = Join-Path $TestDrive 'native-progress.txt'
        $child = Join-Path $PSHOME 'pwsh.exe'
        $childCode = "[Console]::Out.Write('exact-package-hash'); [Console]::Error.Write('captured-stderr'); exit 0"
        $result = & $invoke $child @('-NoProfile', '-Command', $childCode) ([string]$TestDrive) 6> $progressLog
        $result.Stdout | Should Be 'exact-package-hash'
        $result.Stderr | Should Be 'captured-stderr'
        $progress = Get-Content -LiteralPath $progressLog -Raw
        $progress | Should Match '\[GhosttyRuntime\] started pid='
        $progress | Should Match 'code=0; draining output'
        $progress | Should Match '\[GhosttyRuntime\] completed'

        $failure = ''
        try {
            $null = & $invoke $child @('-NoProfile', '-Command', "[Console]::Error.Write('expected-native-failure'); exit 7") ([string]$TestDrive) 6>> $progressLog
        }
        catch { $failure = $_.Exception.Message }
        $failure | Should Match 'failed with exit code 7'
        $failure | Should Match 'expected-native-failure'
    }
}

Describe 'Ghostty native command bounds' {
    Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1')
    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($modulePath, [ref]$tokens, [ref]$parseErrors)
    $functions = @($ast.FindAll({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -in @('Write-RSGhosttyContainedProcessSnapshot', 'Invoke-CheckedProcess')
    }, $true) | ForEach-Object { $_.Extent.Text })
    $script:boundedInvoke = [scriptblock]::Create(
        'param($FilePath, $Arguments, $WorkingDirectory, $TimeoutSeconds, $DrainSeconds)' + "`n" +
        ($functions -join "`n") + "`n" +
        'Invoke-CheckedProcess -FilePath $FilePath -ArgumentList $Arguments -WorkingDirectory $WorkingDirectory ' +
        '-TimeoutSeconds $TimeoutSeconds -OutputDrainTimeoutSeconds $DrainSeconds')
    $script:pwsh = Join-Path $PSHOME 'pwsh.exe'

    function Get-RSGhosttyBoundProbeProcesses([string]$Marker) {
        return @(Get-CimInstance Win32_Process -Filter "Name='pwsh.exe'" |
            Where-Object { [string]$_.CommandLine -like "*$Marker*" })
    }

    It 'terminates a stalled contained tree at its bound and reports its snapshot' {
        $marker = 'ghostty-stall-' + [guid]::NewGuid().ToString('N')
        $grandchild = "Start-Process -NoNewWindow -FilePath '$script:pwsh' -ArgumentList '-NoProfile','-Command','Start-Sleep 300 # $marker'; Start-Sleep 300"
        $progressLog = Join-Path $TestDrive 'stall-progress.txt'
        $failure = ''
        $clock = [Diagnostics.Stopwatch]::StartNew()
        try {
            $null = & $script:boundedInvoke $script:pwsh @('-NoProfile', '-Command', $grandchild) ([string]$TestDrive) 2 30 6> $progressLog
        }
        catch { $failure = $_.Exception.Message }
        $clock.Elapsed.TotalSeconds | Should BeLessThan 60
        $failure | Should Match 'did not exit within its 2-second bound'
        $failure | Should Match 'contained process tree, which is now terminated'
        $progress = Get-Content -LiteralPath $progressLog -Raw
        $progress | Should Match '\[GhosttyRuntime\] timeout snapshot: pid=\d+ parent=\d+ cpuSeconds=[0-9.]+ cpuSecondsLast5s=[0-9.]+ workingSetMB=\d+ threads=\d+ \['
        $progress | Should Not Match ([regex]::Escape($marker))
        Start-Sleep -Seconds 1
        @(Get-RSGhosttyBoundProbeProcesses $marker).Count | Should Be 0
    }

    It 'terminates a descendant that keeps output open after the child exits' {
        $marker = 'ghostty-drain-' + [guid]::NewGuid().ToString('N')
        $grandchild = "Start-Process -NoNewWindow -FilePath '$script:pwsh' -ArgumentList '-NoProfile','-Command','Start-Sleep 300 # $marker'; exit 0"
        $progressLog = Join-Path $TestDrive 'drain-progress.txt'
        $failure = ''
        try {
            $null = & $script:boundedInvoke $script:pwsh @('-NoProfile', '-Command', $grandchild) ([string]$TestDrive) 120 2 6> $progressLog
        }
        catch { $failure = $_.Exception.Message }
        $failure | Should Match 'exited with code 0, but a contained descendant kept its output open for 2 seconds'
        (Get-Content -LiteralPath $progressLog -Raw) | Should Match '\[GhosttyRuntime\] output-drain snapshot: pid=\d+'
        Start-Sleep -Seconds 1
        @(Get-RSGhosttyBoundProbeProcesses $marker).Count | Should Be 0
    }
}

Describe 'Ghostty host toolchain selection' {
    It 'selects the Zig build for the machine architecture and rejects other hosts' {
        Get-RSGhosttyRuntimeHostPlatform -OSArchitecture ([Runtime.InteropServices.Architecture]::Arm64) | Should BeExactly 'ARM64'
        Get-RSGhosttyRuntimeHostPlatform -OSArchitecture ([Runtime.InteropServices.Architecture]::X64) | Should BeExactly 'x64'
        { Get-RSGhosttyRuntimeHostPlatform -OSArchitecture ([Runtime.InteropServices.Architecture]::X86) } |
            Should Throw 'require an x64 or ARM64 Windows host'
    }

    It 'locks one native Zig archive per supported host for the same release' {
        $lock = Read-RSGhosttyRuntimeLock -LiteralPath $lockPath -SchemaPath $schemaPath -RepoRoot $repoRoot
        $version = [string]$lock.toolchain.zigVersion
        (@($lock.toolchain.archives.hostPlatform) -join ',') | Should BeExactly 'ARM64,x64'
        foreach ($archive in @($lock.toolchain.archives)) {
            $zigArchitecture = if ([string]$archive.hostPlatform -ceq 'ARM64') { 'aarch64' } else { 'x86_64' }
            $stem = "zig-$zigArchitecture-windows-$version"
            [string]$archive.fileName | Should BeExactly "$stem.zip"
            [string]$archive.url | Should BeExactly "https://ziglang.org/download/$version/$stem.zip"
            [string]$archive.executableRelativePath | Should BeExactly "$stem/zig.exe"
        }
    }

    It 'never pins a builder to the emulated x86-64 toolchain on ARM64' {
        $lifecycle = Get-Content -LiteralPath $modulePath -Raw
        $nativeWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/ghostty-runtime-native-qualification.yml') -Raw
        $lifecycle | Should Match '\$hostPlatform = Get-RSGhosttyRuntimeHostPlatform'
        $nativeWorkflow | Should Match '\$hostPlatform = Get-RSGhosttyRuntimeHostPlatform'
        foreach ($source in @($lifecycle, $nativeWorkflow)) {
            $source | Should Not Match "hostPlatform\s+-ceq\s+'x64'"
        }
    }
}

Describe 'Ghostty owned cache cleanup' {
    Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/ArtifactOperationLock.psm1')
    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($modulePath, [ref]$tokens, [ref]$parseErrors)
    $functions = @($ast.FindAll({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -in @('Test-PathInsideTerminalBuildRoot', 'Remove-TerminalBuildDirectory')
    }, $true) | ForEach-Object { $_.Extent.Text })
    $script:ghosttyCleanupInvocation = [scriptblock]::Create(
        'param($repoRoot, $terminalBuildRoot, $Path, [switch]$PreservePrimaryError)' + "`n" +
        ($functions -join "`n") + "`n" +
        'Remove-TerminalBuildDirectory -Path $Path -PreservePrimaryError:$PreservePrimaryError')

    BeforeEach {
        $cleanupRepo = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $terminalRoot = Join-Path $cleanupRepo '.build/TerminalEngine'
        $cache = Join-Path $terminalRoot 'Cache/fetch'
        [void](New-Item -ItemType Directory -Path $cache -Force)
        $marker = Join-Path $cache 'maker.exe'
        'owned harmless fixture' | Set-Content -LiteralPath $marker
    }

    It 'waits for an owned open file to close before removing its cache' {
        if (-not ('RSGhosttyCleanupProbe' -as [type])) {
            Add-Type -TypeDefinition @'
public static class RSGhosttyCleanupProbe {
    public static System.Threading.Tasks.Task ReleaseAfter(System.IDisposable owned) {
        return System.Threading.Tasks.Task.Run(() => {
            System.Threading.Thread.Sleep(1000);
            owned.Dispose();
        });
    }
}
'@
        }
        $stream = [IO.File]::Open($marker, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $release = [RSGhosttyCleanupProbe]::ReleaseAfter($stream)
        try {
            & $script:ghosttyCleanupInvocation $cleanupRepo $terminalRoot $cache
            $release.Wait(5000) | Should Be $true
            Test-Path -LiteralPath $cache | Should Be $false
        } finally {
            $stream.Dispose()
        }
    }

    It 'bounds permanent lock retries and preserves the remaining locked input' {
        Mock Remove-Item { throw [UnauthorizedAccessException]::new('still mapped') }
        Mock Start-Sleep {}
        { & $script:ghosttyCleanupInvocation $cleanupRepo $terminalRoot $cache } | Should Throw 'still mapped'
        Assert-MockCalled Remove-Item -Times 20 -Exactly -Scope It
        Assert-MockCalled Start-Sleep -Times 19 -Exactly -Scope It
        Test-Path -LiteralPath $marker | Should Be $true
    }

    It 'preserves the original fetch error when bounded cleanup also fails' {
        Mock Remove-Item { throw [IO.IOException]::new('cleanup locked') }
        Mock Start-Sleep {}
        Mock Write-Warning {}
        $message = ''
        try {
            try { throw [InvalidOperationException]::new('primary fetch failure') }
            finally { & $script:ghosttyCleanupInvocation $cleanupRepo $terminalRoot $cache -PreservePrimaryError }
        } catch { $message = $_.Exception.Message }
        $message | Should Be 'primary fetch failure'
        Assert-MockCalled Remove-Item -Times 20 -Exactly -Scope It
        Assert-MockCalled Write-Warning -Times 1 -Exactly -Scope It
        Test-Path -LiteralPath $marker | Should Be $true
    }

    It 'rejects an escaped path before any removal attempt' {
        Mock Remove-Item {}
        { & $script:ghosttyCleanupInvocation $cleanupRepo $terminalRoot (Join-Path $cleanupRepo 'outside') } | Should Throw 'outside the terminal build root'
        Assert-MockCalled Remove-Item -Times 0 -Exactly -Scope It
        Test-Path -LiteralPath $marker | Should Be $true
    }
}

Describe 'Ghostty fetch session cache ownership' {
    Import-Module (Join-Path $repoRoot 'Tools/TerminalEngine/TerminalEngineArchive.psm1') -Force
    Import-Module (Join-Path $repoRoot 'Tools/TerminalEngine/TerminalEngineTextIdentity.psm1') -Force
    $source = Get-Content -LiteralPath $modulePath -Raw
    $start = $source.IndexOf('$fetchCacheRoot = Join-Path $cacheRoot ("dependency-fetch-{0}"', [StringComparison]::Ordinal)
    $end = $source.IndexOf('$privateRuntimeOverlayCanApply =', $start, [StringComparison]::Ordinal)
    $session = $source.Substring($start, $end - $start)
    # Run the real bootstrap control flow with a synthetic compiler. Its maker
    # marker is an independent witness for reuse; real archive/hash/license and
    # guarded cleanup helpers still run.
    $script:ghosttyFetchSessionInvocation = [scriptblock]::Create(@'
param($context, $dependencyStates)
$repoRoot = $context.Fixture.Owner
$terminalBuildRoot = Join-Path $repoRoot '.build/TerminalEngine'
$cacheRoot = Join-Path $terminalBuildRoot 'Cache'
$zigVersion = 'fixture'
$zigExecutable = 'fixture-zig'
$sourceRoot = $context.Fixture.Source
$tarCommand = [pscustomobject]@{ Source='fixture-tar' }
function Remove-TerminalBuildDirectory {
    param($Path, [switch]$PreservePrimaryError)
    $context.CleanupFlags.Add([bool]$PreservePrimaryError)
    & $script:ghosttyCleanupInvocation $repoRoot $terminalBuildRoot $Path -PreservePrimaryError:$PreservePrimaryError
}
function Invoke-CheckedProcess {
    param($FilePath, $ArgumentList, $WorkingDirectory, $Environment)
    if ($FilePath -ceq 'fixture-tar') {
        return [pscustomobject]@{ Stdout=[IO.File]::ReadAllText((Join-Path $sourceRoot 'LICENSE')) }
    }
    if ($FilePath -cne 'fixture-zig') { throw 'Unexpected compiler fixture command.' }
    $context.FetchCount++
    $maker = Join-Path $Environment.ZIG_GLOBAL_CACHE_DIR 'maker-witness'
    if (-not [IO.File]::Exists($maker)) {
        $context.BootstrapCount++
        [IO.File]::WriteAllText($maker, 'one owned compiler bootstrap')
    }
    if ($dependency.id -ceq $context.FailPackage) { throw 'Synthetic second fetch failure.' }
    $packages = Join-Path $Environment.ZIG_GLOBAL_CACHE_DIR 'p'
    [void][IO.Directory]::CreateDirectory($packages)
    Copy-Item -LiteralPath $context.TemplateArchive -Destination (Join-Path $packages "$($dependency.packageName).tar.gz")
    return [pscustomobject]@{ Stdout=$dependency.packageName }
}
'@ + "`n" + $session)

    foreach ($scenario in @('cold', 'warm', 'second-fetch-failure')) {
        It "retains one owned compiler cache across packages and cleans the $scenario session" {
            $fixture = New-RSGhosttyLockedPackageFixture
            $template = Join-Path $fixture.Downloads 'fixture-pkg.tar.gz'
            $archiveHash = (Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash.ToLowerInvariant()
            $licenseHash = (Get-RSTerminalLfStringIdentity -Text ([IO.File]::ReadAllText((Join-Path $fixture.Source 'LICENSE')))).Sha256
            $states = @('one','two' | ForEach-Object {
                $target = Join-Path $fixture.Downloads "$_-canonical.tar.gz"
                if ($scenario -ceq 'warm') { Copy-Item -LiteralPath $template -Destination $target }
                [pscustomobject]@{ Id=$_; SourceArchive=$template; ContentArchive=$target; Lock=[pscustomobject]@{
                    id=$_; packageName="$_-pkg"; licenseArchiveRelativePath='LICENSE'; licenseSha256=$licenseHash
                    sourceArchive=[pscustomobject]@{ sha256=$archiveHash; internalRoot='fixture-pkg/wrapper' }
                    contentArchive=[pscustomobject]@{ sha256=$archiveHash; internalRoot='fixture-pkg/wrapper' }
                } }
            })
            $context = [pscustomobject]@{
                Fixture=$fixture; TemplateArchive=$template; FetchCount=0; BootstrapCount=0
                FailPackage=$(if ($scenario -ceq 'second-fetch-failure') { 'two' } else { '' })
                CleanupFlags=[Collections.Generic.List[bool]]::new()
            }
            if ($scenario -ceq 'second-fetch-failure') {
                { & $script:ghosttyFetchSessionInvocation $context $states } | Should Throw 'Synthetic second fetch failure'
                $context.CleanupFlags[$context.CleanupFlags.Count-1] | Should Be $true
                Test-Path -LiteralPath $states[0].ContentArchive | Should Be $true
                Test-Path -LiteralPath $states[1].ContentArchive | Should Be $false
            } else {
                & $script:ghosttyFetchSessionInvocation $context $states
                @($states | Where-Object { -not [IO.File]::Exists($_.ContentArchive) }).Count | Should Be 0
            }
            $context.FetchCount | Should Be $(if ($scenario -ceq 'warm') { 0 } else { 2 })
            $context.BootstrapCount | Should Be $(if ($scenario -ceq 'warm') { 0 } else { 1 })
            Test-Path -LiteralPath (Join-Path $fixture.Owner '.build/TerminalEngine/Cache/dependency-fetch-fixture') | Should Be $false
        }
    }
}

function Assert-RSInvalidGhosttyRuntimeLock {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Mutate,
        [Parameter(Mandatory = $true)][string]$ExpectedMessage
    )

    $fixture = Copy-RSGhosttyRuntimeLockFixture
    & $Mutate $fixture
    $fixturePath = Join-Path $TestDrive ('ghostty-runtime-lock-' + [Guid]::NewGuid().ToString('N') + '.json')
    $fixture | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $fixturePath -Encoding utf8NoBOM

    $message = ''
    try {
        [void](Read-RSGhosttyRuntimeLock `
                -LiteralPath $fixturePath `
                -SchemaPath $schemaPath `
                -RepoRoot $repoRoot)
    }
    catch {
        $message = $_.Exception.Message
    }
    [string]::IsNullOrWhiteSpace($message) | Should Be $false
    $message | Should Match $ExpectedMessage
}

function Assert-RSGhosttyBuildPlanRejected {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [Parameter(Mandatory = $true)][string]$ExpectedMessage
    )

    $message = ''
    try {
        [void](& $Action)
    }
    catch {
        $message = $_.Exception.Message
    }
    [string]::IsNullOrWhiteSpace($message) | Should Be $false
    $message | Should Match $ExpectedMessage
}

Describe 'Ghostty nested schema diagnostics' {
    It 'preserves field-specific rejection when Test-Json emits no nested errors' {
        [void](Microsoft.PowerShell.Utility\Test-Json -Json '{}' -Schema '{"type":"object"}')
        $globalFetchBefore = [Json.Schema.SchemaRegistry]::Global.Fetch
        Mock Test-Json { $false } -ModuleName GhosttyRuntimeLifecycle
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'upstream\.archive\.sha256' -Mutate {
            param($lock)
            $lock.upstream.archive.sha256 = $lock.upstream.archive.sha256.ToUpperInvariant()
        }
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'abi\.exports' -Mutate {
            param($lock)
            $lock.abi.exports = @($lock.abi.exports) + $lock.abi.exports[-1]
            $lock.abi.exportCount = @($lock.abi.exports).Count
        }
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'overlay\.path' -Mutate {
            param($lock)
            $lock.overlay.path = '../escape.patch'
        }
        Assert-MockCalled Test-Json -ModuleName GhosttyRuntimeLifecycle -Times 3 -Exactly -Scope It
        [object]::ReferenceEquals($globalFetchBefore, [Json.Schema.SchemaRegistry]::Global.Fetch) | Should Be $true
    }

    It 'never admits a declined document when diagnostic reevaluation finds no errors' {
        [void](Microsoft.PowerShell.Utility\Test-Json -Json '{}' -Schema '{"type":"object"}')
        Mock Test-Json { $false } -ModuleName GhosttyRuntimeLifecycle
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'schema validation failed' -Mutate { param($lock) }
        Assert-MockCalled Test-Json -ModuleName GhosttyRuntimeLifecycle -Times 1 -Exactly -Scope It
    }

    It 'refuses external reference fetching during diagnostic reevaluation' {
        [void](Microsoft.PowerShell.Utility\Test-Json -Json '{}' -Schema '{"type":"object"}')
        Mock Test-Json { $false } -ModuleName GhosttyRuntimeLifecycle
        $externalSchema = Join-Path $TestDrive 'external-schema.json'
        '{"$ref":"https://example.invalid/blocked-schema"}' | Set-Content -LiteralPath $externalSchema -Encoding utf8NoBOM
        $message = ''
        try {
            [void](Read-RSGhosttyRuntimeLock -LiteralPath $lockPath -SchemaPath $externalSchema -RepoRoot $repoRoot)
        }
        catch {
            $message = $_.Exception.Message
        }
        $message | Should Match 'reference fetching is disabled'
        Assert-MockCalled Test-Json -ModuleName GhosttyRuntimeLifecycle -Times 1 -Exactly -Scope It
    }
}

Describe 'Ghostty active product runtime lock' {
    It 'keeps native ARM64 qualification evidence closed and machine-verifiable' {
        $nativeSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeNativeQualification.schema.json'
        $record = [ordered]@{
            formatId = 'red-salamander-ghostty-native-qualification'
            schemaVersion = 1
            generatedUtc = '2026-08-26T00:00:00Z'
            sourceCommit = '1' * 40
            workflowRunId = '12345'
            workflowRunAttempt = 1
            testSeed = '0x97d35bc7'
            candidate = [ordered]@{
                commit = '2' * 40
                tree = '3' * 40
                lockFileSha256 = '4' * 64
                lockCanonicalDigest = '5' * 64
                overlaySha256 = '6' * 64
            }
            host = [ordered]@{
                osArchitecture = 'Arm64'
                processArchitecture = 'Arm64'
                pythonArchitecture = 'ARM64'
                pythonVersion = '3.13.14'
                osVersion = 'Microsoft Windows NT 10.0.26200.0'
            }
            testHarness = [ordered]@{
                runnerMode = 'simple'
                runnerSource = 'pinned-zig-default-test-runner'
                runnerSourceSha256 = 'b' * 64
                transportOnlyOverride = $true
                cAbiDiagnosticModuleClone = $true
                cAbiStrip = $false
                cAbiOmitFramePointer = $false
                cAbiUnwindTables = 'sync'
                codeFiveRetryUsed = $false
                codeFiveRetryExecutableSha256 = $null
                codeFiveRetryStatus = 'not-needed'
            }
            runtime = [ordered]@{
                identitySha256 = '7' * 64
                dllSha256 = '8' * 64
                dllNormalizedSha256 = '9' * 64
                dllSizeBytes = 1881600
                machine = 'ARM64'
                exportCount = 191
                exportSetSha256 = 'a' * 64
                importModules = @('kernel32.dll', 'ntdll.dll')
                dependencyIds = @('translate-c', 'uucode', 'wuffs')
                noticeIds = @('ghostty', 'unicode', 'uucode', 'uucode-bjoern-hoehrmann', 'wuffs', 'zig')
            }
            tests = @(
                [ordered]@{ target = 'test-lib-vt'; status = 'passed' }
                [ordered]@{ target = 'test-lib-vt-build'; status = 'passed' }
                [ordered]@{ target = 'test-lib-vt-schema'; status = 'passed' }
            )
        }
        $json = $record | ConvertTo-Json -Depth 100
        (Test-Json -Json $json -SchemaFile $nativeSchemaPath) | Should Be $true
        $record.testHarness.codeFiveRetryUsed = $true
        $record.testHarness.codeFiveRetryExecutableSha256 = 'c' * 64
        $record.testHarness.codeFiveRetryStatus = 'passed'
        (Test-Json -Json ($record | ConvertTo-Json -Depth 100) -SchemaFile $nativeSchemaPath -ErrorAction SilentlyContinue) |
            Should Be $false
        $record.testHarness.codeFiveRetryUsed = $false
        (Test-Json -Json ($record | ConvertTo-Json -Depth 100) -SchemaFile $nativeSchemaPath -ErrorAction SilentlyContinue) |
            Should Be $false
        $record.testHarness.codeFiveRetryExecutableSha256 = $null
        $record.testHarness.codeFiveRetryStatus = 'not-needed'
        $record.unexpected = $true
        (Test-Json -Json ($record | ConvertTo-Json -Depth 100) -SchemaFile $nativeSchemaPath -ErrorAction SilentlyContinue) |
            Should Be $false
    }

    It 'validates the canonical admitted lock and produces a stable normalized digest' {
        $first = Read-RSGhosttyRuntimeLock `
            -LiteralPath $lockPath `
            -SchemaPath $schemaPath `
            -RepoRoot $repoRoot
        $second = Read-RSGhosttyRuntimeLock `
            -LiteralPath $lockPath `
            -SchemaPath $schemaPath `
            -RepoRoot $repoRoot

        $first.upstream.commit | Should Be '83fb95624b0439ee660c2679b43b38c8b63769cd'
        $first.upstream.tree | Should Be '4dbd485a3b72c6ced65ad7ea447644a5967e382e'
        @($first.abi.exports).Count | Should Be 207
        (Get-RSGhosttyRuntimeLockDigest -Lock $first) | Should Be (Get-RSGhosttyRuntimeLockDigest -Lock $second)
    }

    It 'locks exactly the Ghostty exports consumed by TerminalRuntimeLoader' {
        $lock = Read-RSGhosttyRuntimeLock `
            -LiteralPath $lockPath `
            -SchemaPath $schemaPath `
            -RepoRoot $repoRoot
        $loaderSource = @(
            (Get-Content -LiteralPath (Join-Path $repoRoot 'Plugins\Terminal\TerminalRuntimeLoader.cpp') -Raw),
            (Get-Content -LiteralPath (Join-Path $repoRoot 'Plugins\Terminal\TerminalRuntimeLoader.h') -Raw)
        ) -join "`n"
        $actual = [string[]]@([regex]::Matches($loaderSource, '\bghostty_[A-Za-z0-9_]+\b') |
            ForEach-Object { $_.Value } |
            Sort-Object -Unique -CaseSensitive)
        $expected = [string[]]@($lock.abi.loaderRequiredExports)
        ($actual -join "`0") | Should Be ($expected -join "`0")
    }

    It 'keeps changed canonical representations separate from retained pre-upgrade downloads' {
        $legacyPath = Join-Path $repoRoot 'Specs/TestRuns/4cb089111a23/Terminal/20261028_174540_ghostty140-product/provenance/runtime.lock.json'
        $legacy = Get-Content -LiteralPath $legacyPath -Raw | ConvertFrom-Json -Depth 100
        $current = Copy-RSGhosttyRuntimeLockFixture
        $oldDownloads = @{}
        foreach ($dependency in $legacy.dependencies) {
            foreach ($archive in @($dependency.sourceArchive, $dependency.contentArchive)) {
                $oldDownloads[$archive.fileName] = $archive.sha256
            }
        }
        foreach ($dependency in $current.dependencies) {
            foreach ($archive in @($dependency.sourceArchive, $dependency.contentArchive)) {
                if ($oldDownloads.ContainsKey($archive.fileName)) {
                    $archive.sha256 | Should Be $oldDownloads[$archive.fileName]
                }
            }
        }
    }

    It 'keeps the product builder lock-driven without duplicated active identities' {
        $builder = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Build-GhosttyTerminalRuntime.ps1') -Raw
        $lifecycle = Get-Content -LiteralPath $modulePath -Raw
        @($builder -split "`r?`n").Count | Should BeLessThan 100
        $builder | Should Match 'Invoke-RSGhosttyRuntimeBuild'
        $builder | Should Match 'CandidateLockPath'
        $lifecycle | Should Match 'External\\TerminalEngine\\GhosttyRuntimeLock\.v1\.json'
        $lifecycle | Should Match 'Read-RSGhosttyRuntimeLock'
        $lifecycle | Should Match 'Start-RSContainedProcess\s+-ProcessStartInfo'
        $lifecycle | Should Not Match '\bStop-Process\b|\btaskkill(?:\.exe)?\b'
        $lifecycle | Should Match 'Assert-RSGhosttyRuntimeOutputReplaceable\s+-OutputRoot\s+\$productRoot\s*\r?\n\s*Remove-TerminalBuildDirectory\s+-Path\s+\$productRoot\s*\r?\n\s*Move-Item\s+-LiteralPath\s+\$partialProduct'
        $builder | Should Not Match 'TerminalEngine\.lock\.json'
        foreach ($duplicatedIdentity in @(
                'b2d44625908eb56aee299d8511d803bc11ab79fc',
                '68659eb5f1e4eb1437a722f1dd889c5a322c9954607f5edcf337bc3684a75a7e',
                '596894726504f7d398ee2a87b8a53116c8d727e5390b1e066293e1770e40fced',
                'f2aa0f100e6d03a03fab433ee7d5ec72d58fef82663326732eddd6547781999b')) {
            $builder | Should Not Match $duplicatedIdentity
        }
    }

    It 'resolves Product mode only to canonical Product state with pairing enabled' {
        $plan = Resolve-RSGhosttyRuntimeBuildPlan -RepoRoot $repoRoot -Platform x64
        $plan.Mode | Should Be 'Product'
        $plan.LockPath | Should Be $lockPath
        $plan.StateRoot | Should Be (Join-Path $repoRoot '.build\TerminalEngine')
        $plan.OutputRoot | Should Be (Join-Path $repoRoot '.build\TerminalEngine\Product\x64')
        $plan.WritePairingHeader | Should Be $true
    }

    It 'confines explicit candidate locks and outputs and disables pairing publication' {
        $caseRoot = Join-Path $repoRoot ".build\TerminalEngineUpgrade\Tests\$([Guid]::NewGuid().ToString('N'))"
        $candidateLock = Join-Path $caseRoot 'candidate.lock.json'
        $candidateOutput = Join-Path $caseRoot 'output\x64'
        try {
            [void](New-Item -ItemType Directory -Path $caseRoot -Force)
            Copy-Item -LiteralPath $lockPath -Destination $candidateLock
            $plan = Resolve-RSGhosttyRuntimeBuildPlan `
                -RepoRoot $repoRoot `
                -Platform x64 `
                -CandidateLockPath $candidateLock `
                -CandidateOutputRoot $candidateOutput
            $plan.Mode | Should Be 'Candidate'
            $plan.LockPath | Should Be $candidateLock
            $plan.StateRoot | Should Be (Join-Path $repoRoot '.build\TerminalEngineUpgrade')
            $plan.OutputRoot | Should Be $candidateOutput
            $plan.WritePairingHeader | Should Be $false
            $plan.OutputRoot | Should Not Match '[\\/]TerminalEngine[\\/]Product[\\/]'
        }
        finally {
            if (Test-Path -LiteralPath $caseRoot) {
                Remove-Item -LiteralPath $caseRoot -Recurse -Force
            }
        }
    }

    It 'rejects incomplete missing escaping root and Product candidate destinations without fallback' {
        $caseRoot = Join-Path $repoRoot ".build\TerminalEngineUpgrade\Tests\$([Guid]::NewGuid().ToString('N'))"
        $candidateLock = Join-Path $caseRoot 'candidate.lock.json'
        try {
            [void](New-Item -ItemType Directory -Path $caseRoot -Force)
            Copy-Item -LiteralPath $lockPath -Destination $candidateLock
            Assert-RSGhosttyBuildPlanRejected -ExpectedMessage 'requires both' -Action {
                Resolve-RSGhosttyRuntimeBuildPlan -RepoRoot $repoRoot -Platform x64 -CandidateLockPath $candidateLock
            }
            Assert-RSGhosttyBuildPlanRejected -ExpectedMessage 'requires both' -Action {
                Resolve-RSGhosttyRuntimeBuildPlan -RepoRoot $repoRoot -Platform x64 -CandidateOutputRoot (Join-Path $caseRoot 'out')
            }
            Assert-RSGhosttyBuildPlanRejected -ExpectedMessage 'no canonical-lock fallback' -Action {
                Resolve-RSGhosttyRuntimeBuildPlan -RepoRoot $repoRoot -Platform x64 -CandidateLockPath (Join-Path $caseRoot 'missing.json') -CandidateOutputRoot (Join-Path $caseRoot 'out')
            }
            Assert-RSGhosttyBuildPlanRejected -ExpectedMessage 'contained child' -Action {
                Resolve-RSGhosttyRuntimeBuildPlan -RepoRoot $repoRoot -Platform x64 -CandidateLockPath $candidateLock -CandidateOutputRoot $repoRoot
            }
            Assert-RSGhosttyBuildPlanRejected -ExpectedMessage 'contained child' -Action {
                Resolve-RSGhosttyRuntimeBuildPlan -RepoRoot $repoRoot -Platform x64 -CandidateLockPath $candidateLock -CandidateOutputRoot (Join-Path $repoRoot '.build\TerminalEngineUpgrade')
            }
            Assert-RSGhosttyBuildPlanRejected -ExpectedMessage 'Candidate lock must be' -Action {
                Resolve-RSGhosttyRuntimeBuildPlan -RepoRoot $repoRoot -Platform x64 -CandidateLockPath $lockPath -CandidateOutputRoot (Join-Path $caseRoot 'out')
            }
        }
        finally {
            if (Test-Path -LiteralPath $caseRoot) {
                Remove-Item -LiteralPath $caseRoot -Recurse -Force
            }
        }
    }

    It 'rejects a missing candidate descriptor before mutating Product identity or pairing bytes' {
        $productRoot = Join-Path $repoRoot '.build\TerminalEngine\Product\x64'
        $protectedPaths = @(
            (Join-Path $productRoot 'terminal-runtime-identity.json'),
            (Join-Path $productRoot 'include\red_salamander\terminal_runtime_identity.generated.h'),
            (Join-Path $productRoot 'bin\ghostty-vt.dll'))
        $before = @($protectedPaths | ForEach-Object {
                if (Test-Path -LiteralPath $_ -PathType Leaf) {
                    (Get-FileHash -Algorithm SHA256 -LiteralPath $_).Hash
                }
                else { '<missing>' }
            })
        $caseRoot = Join-Path $repoRoot ".build\TerminalEngineUpgrade\Tests\$([Guid]::NewGuid().ToString('N'))"
        try {
            [void](New-Item -ItemType Directory -Path $caseRoot -Force)
            $message = ''
            try {
                & (Join-Path $repoRoot 'Tools\Build-GhosttyTerminalRuntime.ps1') `
                    -Platform x64 `
                    -CandidateLockPath (Join-Path $caseRoot 'missing.lock.json') `
                    -CandidateOutputRoot (Join-Path $caseRoot 'output')
            }
            catch {
                $message = $_.Exception.Message
            }
            $message | Should Match 'no canonical-lock fallback'
            $after = @($protectedPaths | ForEach-Object {
                    if (Test-Path -LiteralPath $_ -PathType Leaf) {
                        (Get-FileHash -Algorithm SHA256 -LiteralPath $_).Hash
                    }
                    else { '<missing>' }
                })
            ($after -join "`0") | Should Be ($before -join "`0")
        }
        finally {
            if (Test-Path -LiteralPath $caseRoot) {
                Remove-Item -LiteralPath $caseRoot -Recurse -Force
            }
        }
    }

    It 'fails closed before publication when an exact runtime output is independently owned' {
        $caseRoot = Join-Path $repoRoot ".build\TerminalEngineUpgrade\Tests\$([Guid]::NewGuid().ToString('N'))"
        $runtimePath = Join-Path $caseRoot 'output\bin\ghostty-vt.dll'
        $probe = $null
        try {
            [void](New-Item -ItemType Directory -Path (Split-Path -Parent $runtimePath) -Force)
            [IO.File]::WriteAllBytes($runtimePath, [byte[]](0x4D, 0x5A, 0x01, 0x02))
            $before = (Get-FileHash -Algorithm SHA256 -LiteralPath $runtimePath).Hash
            $probe = [IO.File]::Open($runtimePath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
            $message = ''
            try {
                Assert-RSGhosttyRuntimeOutputReplaceable -OutputRoot (Split-Path -Parent (Split-Path -Parent $runtimePath))
            }
            catch {
                $message = $_.Exception.Message
            }
            $message | Should Match 'independently owned or protected file'
            $message | Should Match 'No process was terminated'
            $probe.Dispose()
            $probe = $null
            (Get-FileHash -Algorithm SHA256 -LiteralPath $runtimePath).Hash | Should Be $before
        }
        finally {
            if ($null -ne $probe) {
                $probe.Dispose()
            }
            if (Test-Path -LiteralPath $caseRoot) {
                Remove-Item -LiteralPath $caseRoot -Recurse -Force
            }
        }
    }

    It 'rejects a missing required property with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage "missing required property 'upstream'" -Mutate {
            param($lock)
            $lock.PSObject.Properties.Remove('upstream')
        }
    }

    It 'rejects an additional property with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage "additional property 'unexpected'" -Mutate {
            param($lock)
            $lock | Add-Member -NotePropertyName unexpected -NotePropertyValue 'rejected'
        }
    }

    It 'rejects uppercase hashes with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'upstream\.archive\.sha256' -Mutate {
            param($lock)
            $lock.upstream.archive.sha256 = $lock.upstream.archive.sha256.ToUpperInvariant()
        }
    }

    It 'rejects an unsorted export set with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'abi\.exports.*sorted' -Mutate {
            param($lock)
            $exports = @($lock.abi.exports)
            $first = $exports[0]
            $exports[0] = $exports[1]
            $exports[1] = $first
            $lock.abi.exports = $exports
        }
    }

    It 'rejects a duplicate export with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'abi\.exports' -Mutate {
            param($lock)
            $lock.abi.exports = @($lock.abi.exports) + $lock.abi.exports[-1]
            $lock.abi.exportCount = @($lock.abi.exports).Count
        }
    }

    It 'rejects a derived export hash mismatch with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'abi\.exportSetSha256' -Mutate {
            param($lock)
            $lock.abi.exportSetSha256 = 'a' * 64
        }
    }

    It 'rejects a loader-required export outside the complete export set' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'abi\.loaderRequiredExports.*ghostty_z_missing' -Mutate {
            param($lock)
            $lock.abi.loaderRequiredExports = @($lock.abi.loaderRequiredExports) + 'ghostty_z_missing'
        }
    }

    It 'rejects a bad platform machine with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'build\.platforms\[x64\]\.machine' -Mutate {
            param($lock)
            @($lock.build.platforms | Where-Object { $_.platform -ceq 'x64' })[0].machine = 'ARM64'
        }
    }

    It 'rejects a path escape with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'overlay\.path' -Mutate {
            param($lock)
            $lock.overlay.path = '../escape.patch'
        }
    }

    It 'rejects an unsupported schema version with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'schemaVersion.*exactly 1' -Mutate {
            param($lock)
            $lock.schemaVersion = 2
        }
    }

    It 'rejects an overlay digest mismatch with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'overlay\.normalizedTextSha256.*does not match' -Mutate {
            param($lock)
            $lock.overlay.normalizedTextSha256 = 'a' * 64
        }
    }

    It 'rejects a missing repository license input with a field-specific message' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'notices\[unicode\]\.source\.path.*missing license input' -Mutate {
            param($lock)
            @($lock.notices | Where-Object { $_.id -ceq 'unicode' })[0].source.path =
                'External/TerminalEngine/Notices/missing-unicode-license.txt'
        }
    }

    It 'validates an ordinally sorted multi-dependency lock and dependency notice closure' {
        $fixture = Copy-RSGhosttyRuntimeLockFixture
        $synthetic = $fixture.dependencies[0] |
            ConvertTo-Json -Depth 100 |
            ConvertFrom-Json -Depth 100 -NoEnumerate -DateKind String
        $synthetic.id = 'a-test'
        $synthetic.packageName = 'a-test-package'
        $synthetic.sourceArchive.fileName = 'a-test-source.tar.gz'
        $synthetic.sourceArchive.url = 'https://deps.files.ghostty.org/a-test-source.tar.gz'
        $synthetic.sourceArchive.internalRoot = 'a-test-source'
        $synthetic.contentArchive.fileName = 'a-test-package.tar.gz'
        $synthetic.contentArchive.internalRoot = 'a-test-package'
        $fixture.dependencies = @($synthetic) + @($fixture.dependencies)

        $syntheticNotice = [pscustomobject]@{
            id = 'a-test'
            source = [pscustomobject]@{
                kind = 'dependency'
                path = [string]$synthetic.licenseArchiveRelativePath
                dependencyId = 'a-test'
            }
            normalizedTextSha256 = [string]$synthetic.licenseSha256
            outputRelativePath = 'share/licenses/terminal-runtime/LICENSE-a-test.txt'
        }
        $fixture.notices = @($syntheticNotice) + @($fixture.notices)
        $fixturePath = Join-Path $TestDrive 'ghostty-runtime-lock-multi-dependency.json'
        $fixture | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $fixturePath -Encoding utf8NoBOM

        $lock = Read-RSGhosttyRuntimeLock `
            -LiteralPath $fixturePath `
            -SchemaPath $schemaPath `
            -RepoRoot $repoRoot

        (@($lock.dependencies.id) -join ',') | Should Be 'a-test,aro,pixels,translate-c,uucode,wuffs'
        (@($lock.notices.id) -contains 'a-test') | Should Be $true
    }

    It 'rejects a dependency notice that names an undeclared dependency' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage "names unknown dependency 'missing-dependency'" -Mutate {
            param($lock)
            @($lock.notices | Where-Object { $_.id -ceq 'uucode' })[0].source.dependencyId =
                'missing-dependency'
        }
    }

    It 'rejects case-insensitive download and notice path collisions' {
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'download fileName.*duplicate' -Mutate {
            param($lock)
            $lock.dependencies[0].sourceArchive.fileName =
                $lock.toolchain.archives[0].fileName.ToUpperInvariant()
        }
        Assert-RSInvalidGhosttyRuntimeLock -ExpectedMessage 'notices\.outputRelativePath.*duplicate' -Mutate {
            param($lock)
            $lock.notices[1].outputRelativePath = $lock.notices[0].outputRelativePath.ToUpperInvariant()
        }
    }
}
