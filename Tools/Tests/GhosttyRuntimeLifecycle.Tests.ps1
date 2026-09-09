$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyRuntimeLifecycle.psm1'
$lockPath = Join-Path $repoRoot 'External\TerminalEngine\GhosttyRuntimeLock.v1.json'
$schemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeLock.schema.json'

Import-Module $modulePath -Force

function Copy-RSGhosttyRuntimeLockFixture {
    return Get-Content -LiteralPath $lockPath -Raw |
        ConvertFrom-Json -Depth 100 -NoEnumerate -DateKind String
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

        $first.upstream.commit | Should Be '9c3ec931d64561a8407dde7ac984ce156ae91539'
        $first.upstream.tree | Should Be '359872ab74fe1ecc96428133c0546be3cb969fdd'
        @($first.abi.exports).Count | Should Be 191
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

        @($lock.dependencies.id) | Should Be @('a-test', 'translate-c', 'uucode', 'wuffs')
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
