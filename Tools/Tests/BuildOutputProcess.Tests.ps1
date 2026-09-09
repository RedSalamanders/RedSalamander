Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$buildScript = Join-Path $repoRoot 'build.ps1'

Describe 'Build output process preflight' {
    BeforeAll {
        $tokens = $null
        $parseErrors = $null
        $buildAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $buildScript,
            [ref]$tokens,
            [ref]$parseErrors)
        if ($parseErrors.Count -ne 0) {
            throw "build.ps1 did not parse: $($parseErrors[0].Message)"
        }

        foreach ($functionName in @('Assert-BuildOutputProcessNotRunning')) {
            $functionAst = $buildAst.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq $functionName
            }, $true)
            if ($null -eq $functionAst) {
                throw "build.ps1 function '$functionName' was not found."
            }
            Invoke-Expression $functionAst.Extent.Text
        }
    }

    BeforeEach {
        $script:buildOutputProcesses = @()
        Mock Get-CimInstance { return $script:buildOutputProcesses }
        Mock Stop-Process {}
    }

    It 'fails closed with diagnostics and kills nothing when an exact-path process is a self-test' {
        $expectedPath = Join-Path $TestDrive 'Debug\RedSalamander.exe'
        $script:buildOutputProcesses = @(
            [pscustomobject]@{
                ProcessId = 4100
                ExecutablePath = $expectedPath
                CommandLine = "`"$expectedPath`""
            },
            [pscustomobject]@{
                ProcessId = 4101
                ExecutablePath = $expectedPath
                CommandLine = "`"$expectedPath`" --commands-selftest --selftest-repeat=100"
            }
        )

        $errorMessage = $null
        try {
            Assert-BuildOutputProcessNotRunning -ProcessName 'RedSalamander.exe' -ExpectedExePath $expectedPath
        }
        catch {
            $errorMessage = $_.Exception.Message
        }

        $errorMessage | Should Match 'independently launched process is using an exact target output'
        $errorMessage | Should Match 'was not terminated because it is not proven to belong to this operation'
        $errorMessage | Should Match 'PID=4101'
        $errorMessage | Should Match ([regex]::Escape([System.IO.Path]::GetFullPath($expectedPath)))
        $errorMessage | Should Match ([regex]::Escape('--commands-selftest --selftest-repeat=100'))
        Assert-MockCalled Stop-Process -Times 0
    }

    It 'refuses to kill an exact-path process when its command line is unavailable' {
        $expectedPath = Join-Path $TestDrive 'Debug\RedSalamander.exe'
        $script:buildOutputProcesses = @(
            [pscustomobject]@{
                ProcessId = 4201
                ExecutablePath = $expectedPath
                CommandLine = $null
            }
        )

        $errorMessage = $null
        try {
            Assert-BuildOutputProcessNotRunning -ProcessName 'RedSalamander.exe' -ExpectedExePath $expectedPath
        }
        catch {
            $errorMessage = $_.Exception.Message
        }

        $errorMessage | Should Match 'PID=4201'
        $errorMessage | Should Match "CommandLine='<unavailable>'"
        Assert-MockCalled Stop-Process -Times 0
    }

    It 'preserves both exact-path and foreign-checkout interactive processes while failing only for the exact target' {
        $expectedPath = Join-Path $TestDrive 'Debug\RedSalamander.exe'
        $otherPath = Join-Path $TestDrive 'Other\RedSalamander.exe'
        $script:buildOutputProcesses = @(
            [pscustomobject]@{
                ProcessId = 4301
                ExecutablePath = $otherPath
                CommandLine = "`"$otherPath`" --interactive"
            },
            [pscustomobject]@{
                ProcessId = 4302
                ExecutablePath = $expectedPath
                CommandLine = "`"$expectedPath`""
            }
        )

        $errorMessage = $null
        try {
            Assert-BuildOutputProcessNotRunning -ProcessName 'RedSalamander.exe' -ExpectedExePath $expectedPath
        }
        catch {
            $errorMessage = $_.Exception.Message
        }

        $errorMessage | Should Match 'PID=4302'
        $errorMessage | Should Not Match 'PID=4301'
        Assert-MockCalled Stop-Process -Times 0
    }

    It 'allows a foreign-checkout process to survive without blocking this output profile' {
        $expectedPath = Join-Path $TestDrive 'CurrentCheckout\.build\x64\Debug\RedSalamander.exe'
        $foreignPath = Join-Path $TestDrive 'ForeignCheckout\.build\x64\Debug\RedSalamander.exe'
        $script:buildOutputProcesses = @(
            [pscustomobject]@{
                ProcessId = 4401
                ExecutablePath = $foreignPath
                CommandLine = "`"$foreignPath`" --interactive"
            }
        )

        { Assert-BuildOutputProcessNotRunning -ProcessName 'RedSalamander.exe' -ExpectedExePath $expectedPath } |
            Should Not Throw
        Assert-MockCalled Stop-Process -Times 0
    }

    It 'fails closed and kills nothing when target process metadata cannot be enumerated' {
        $expectedPath = Join-Path $TestDrive 'CurrentCheckout\.build\x64\Debug\RedSalamander.exe'
        Mock Get-CimInstance { throw 'CIM metadata unavailable' }

        $errorMessage = $null
        try {
            Assert-BuildOutputProcessNotRunning `
                -ProcessName 'RedSalamander.exe' `
                -ExpectedExePath $expectedPath
        }
        catch {
            $errorMessage = $_.Exception.Message
        }

        $errorMessage | Should Match 'target-output process safety could not be verified'
        $errorMessage | Should Match 'unable to enumerate'
        $errorMessage | Should Match 'No process was terminated'
        $errorMessage | Should Match 'CIM metadata unavailable'
        Assert-MockCalled Stop-Process -Times 0
    }

    It 'creates collision-resistant build log paths for concurrent profiles' {
        $tokens = $null
        $parseErrors = $null
        $buildAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $buildScript,
            [ref]$tokens,
            [ref]$parseErrors)
        @($parseErrors).Count | Should Be 0
        $logFunction = $buildAst.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq 'New-ProcessLogPath'
            }, $true)
        $scriptRootForLogTest = $TestDrive
        Invoke-Expression ($logFunction.Extent.Text.Replace('$PSScriptRoot', '$scriptRootForLogTest'))

        $first = New-ProcessLogPath -Prefix 'msbuild'
        $second = New-ProcessLogPath -Prefix 'msbuild'
        $first | Should Not Be $second
        (Split-Path -Leaf $first) | Should Match "^msbuild-\d{8}_\d{6}_\d{3}-pid$PID-[0-9a-f]{8}\.log$"
    }
}

Describe 'Artifact operation lock' {
    BeforeAll {
        $sanitizedEnvironmentModule = Join-Path $repoRoot 'Tools\Modules\Build\SanitizedEnvironment.psm1'
        $artifactLockModule = Join-Path $repoRoot 'Tools\Modules\Build\ArtifactOperationLock.psm1'
        Import-Module $sanitizedEnvironmentModule -Force -ErrorAction Stop
        # Contained-process delegation discovers the optional lock hooks from
        # the global command table, matching production callers.
        Import-Module $artifactLockModule -Force -Global -ErrorAction Stop

        if (-not ('RedSalamander.ToolingTests.InheritableEvent' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace RedSalamander.ToolingTests
{
    public sealed class InheritableEvent : IDisposable
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct SecurityAttributes
        {
            public int Length;
            public IntPtr SecurityDescriptor;
            [MarshalAs(UnmanagedType.Bool)]
            public bool InheritHandle;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateEventW(
            ref SecurityAttributes attributes,
            [MarshalAs(UnmanagedType.Bool)] bool manualReset,
            [MarshalAs(UnmanagedType.Bool)] bool initialState,
            string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);

        private IntPtr handle;

        public InheritableEvent()
        {
            SecurityAttributes attributes = new SecurityAttributes();
            attributes.Length = Marshal.SizeOf(typeof(SecurityAttributes));
            attributes.InheritHandle = true;
            handle = CreateEventW(ref attributes, true, false, null);
            if (handle == IntPtr.Zero || handle == new IntPtr(-1))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "CreateEventW failed.");
            }
        }

        public long HandleValue { get { return handle.ToInt64(); } }

        public bool IsSignaled
        {
            get
            {
                uint result = WaitForSingleObject(handle, 0);
                if (result == 0u) return true;
                if (result == 258u) return false;
                throw new Win32Exception(Marshal.GetLastWin32Error(), "WaitForSingleObject failed.");
            }
        }

        public void Dispose()
        {
            if (handle != IntPtr.Zero)
            {
                CloseHandle(handle);
                handle = IntPtr.Zero;
            }
            GC.SuppressFinalize(this);
        }
    }
}
'@
        }
    }

    It 'creates only reparse-free destinations beneath an exact guarded repository root' {
        $testRepo = Join-Path $TestDrive 'GuardedPathRepo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $destination = Resolve-RSGuardedRepositoryPath `
            -RepoRoot $testRepo `
            -GuardedRelativeRoot '.build\AppPackages' `
            -Path '.build\AppPackages\nested\output' `
            -CreateDirectory
        (Test-Path -LiteralPath $destination -PathType Container) | Should Be $true

        {
            Resolve-RSGuardedRepositoryPath `
                -RepoRoot $testRepo `
                -GuardedRelativeRoot '.build\AppPackages' `
                -Path '.build\escaped' `
                -CreateDirectory
        } | Should Throw 'Destination must remain beneath the guarded repository root'
        (Test-Path -LiteralPath (Join-Path $testRepo '.build\escaped')) | Should Be $false
    }

    It 'rejects a junction escape before creating anything in the junction target' {
        $testRepo = Join-Path $TestDrive 'GuardedJunctionRepo'
        $outside = Join-Path $TestDrive 'GuardedJunctionOutside'
        $buildRoot = Join-Path $testRepo '.build'
        $junction = Join-Path $buildRoot 'AppPackages'
        $sentinel = Join-Path $outside 'sentinel.txt'
        New-Item -ItemType Directory -Path $buildRoot,$outside -Force | Out-Null
        [IO.File]::WriteAllText($sentinel, 'preserve-app-packages-target')
        New-Item -ItemType Junction -Path $junction -Target $outside | Out-Null

        try {
            {
                Resolve-RSGuardedRepositoryPath `
                    -RepoRoot $testRepo `
                    -GuardedRelativeRoot '.build\AppPackages' `
                    -Path '.build\AppPackages\escaped' `
                    -CreateDirectory
            } | Should Throw 'contains a reparse point'
            (Test-Path -LiteralPath (Join-Path $outside 'escaped')) | Should Be $false
            (Get-Content -LiteralPath $sentinel -Raw) | Should Be 'preserve-app-packages-target'
        }
        finally {
            Remove-Item -LiteralPath $junction -Force
        }
    }

    It 'rejects a .build junction before touching its external sentinel' {
        $testRepo = Join-Path $TestDrive 'GuardedBuildJunctionRepo'
        $outside = Join-Path $TestDrive 'GuardedBuildJunctionOutside'
        $junction = Join-Path $testRepo '.build'
        $sentinel = Join-Path $outside 'sentinel.txt'
        New-Item -ItemType Directory -Path $testRepo,$outside -Force | Out-Null
        [IO.File]::WriteAllText($sentinel, 'preserve-build-target')
        New-Item -ItemType Junction -Path $junction -Target $outside | Out-Null

        try {
            {
                Resolve-RSGuardedRepositoryPath `
                    -RepoRoot $testRepo `
                    -GuardedRelativeRoot '.build\AppPackages' `
                    -Path '.build\AppPackages\package.zip' `
                    -CreateDirectory
            } | Should Throw 'contains a reparse point'
            (Test-Path -LiteralPath (Join-Path $outside 'AppPackages')) | Should Be $false
            (Get-Content -LiteralPath $sentinel -Raw) | Should Be 'preserve-build-target'
        }
        finally {
            Remove-Item -LiteralPath $junction -Force
        }
    }

    It 'atomically publishes a same-directory guarded file and preserves prior bytes on rejected input' {
        $testRepo = Join-Path $TestDrive 'GuardedPublicationRepo'
        $output = Join-Path $testRepo '.build\AppPackages'
        [void](New-Item -ItemType Directory -Path $testRepo)
        [void](Resolve-RSGuardedRepositoryPath `
                -RepoRoot $testRepo `
                -GuardedRelativeRoot '.build\AppPackages' `
                -Path $output `
                -CreateDirectory)
        $destination = Join-Path $output 'package.zip'
        $staged = Join-Path $output '.package.new.tmp.zip'
        [IO.File]::WriteAllText($destination, 'previous-complete-zip')
        [IO.File]::WriteAllText($staged, 'new-complete-zip')

        $published = Publish-RSGuardedRepositoryFile `
            -RepoRoot $testRepo `
            -GuardedRelativeRoot '.build\AppPackages' `
            -StagedPath $staged `
            -DestinationPath $destination
        $published | Should Be $destination
        (Get-Content -LiteralPath $destination -Raw) | Should Be 'new-complete-zip'
        (Test-Path -LiteralPath $staged) | Should Be $false

        $outsideStaged = Join-Path $TestDrive 'outside-staged.zip'
        [IO.File]::WriteAllText($outsideStaged, 'untrusted-new-bytes')
        {
            Publish-RSGuardedRepositoryFile `
                -RepoRoot $testRepo `
                -GuardedRelativeRoot '.build\AppPackages' `
                -StagedPath $outsideStaged `
                -DestinationPath $destination
        } | Should Throw 'Destination must remain beneath the guarded repository root'
        (Get-Content -LiteralPath $destination -Raw) | Should Be 'new-complete-zip'
    }

    It 'rejects a final-file reparse destination without changing its external target' {
        $testRepo = Join-Path $TestDrive 'GuardedFinalReparseRepo'
        $output = Join-Path $testRepo '.build\AppPackages'
        $outside = Join-Path $TestDrive 'GuardedFinalReparseOutside'
        [void](New-Item -ItemType Directory -Path $testRepo)
        [void](Resolve-RSGuardedRepositoryPath `
                -RepoRoot $testRepo `
                -GuardedRelativeRoot '.build\AppPackages' `
                -Path $output `
                -CreateDirectory)
        [void](New-Item -ItemType Directory -Path $outside -Force)
        $sentinel = Join-Path $outside 'sentinel.zip'
        $destination = Join-Path $output 'package.zip'
        $staged = Join-Path $output '.package.new.tmp.zip'
        [IO.File]::WriteAllText($sentinel, 'external-final-sentinel')
        [IO.File]::WriteAllText($staged, 'new-complete-zip')
        [void](New-Item -ItemType SymbolicLink -Path $destination -Target $sentinel)

        try {
            {
                Publish-RSGuardedRepositoryFile `
                    -RepoRoot $testRepo `
                    -GuardedRelativeRoot '.build\AppPackages' `
                    -StagedPath $staged `
                    -DestinationPath $destination
            } | Should Throw 'contains a reparse point'
            (Get-Content -LiteralPath $sentinel -Raw) | Should Be 'external-final-sentinel'
        }
        finally {
            Remove-Item -LiteralPath $destination -Force
        }
    }

    It 'uses a canonical repository identity and supports balanced same-thread reentrancy' {
        $testRepo = Join-Path $TestDrive 'Repo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null
        $debugScope = @{ configuration = 'Debug'; platform = 'x64'; kind = 'test' }
        $releaseScope = @{ configuration = 'Release'; platform = 'x64'; kind = 'build' }
        $arm64Scope = @{ configuration = 'Debug'; platform = 'ARM64'; kind = 'build' }
        $debugDependencyScope = @{ configuration = 'Debug'; platform = 'x64'; kind = 'build'; coordination = 'shared-dependency' }
        $releaseDependencyScope = @{ configuration = 'Release'; platform = 'x64'; kind = 'build'; coordination = 'shared-dependency' }

        $lockPath = Get-RSArtifactOperationLockPath -RepoRoot $testRepo -Scope $debugScope
        $caseVariantLockPath = Get-RSArtifactOperationLockPath `
            -RepoRoot $testRepo.ToUpperInvariant() `
            -Scope @{ configuration = 'debug'; platform = 'X64' }
        $lockPath.ToUpperInvariant() | Should Be $caseVariantLockPath.ToUpperInvariant()
        $lockPath | Should Not Be (Get-RSArtifactOperationLockPath -RepoRoot $testRepo -Scope $releaseScope)
        $lockPath | Should Not Be (Get-RSArtifactOperationLockPath -RepoRoot $testRepo -Scope $arm64Scope)
        (Get-RSArtifactOperationLockPath -RepoRoot $testRepo -Scope $debugDependencyScope) |
            Should Be (Get-RSArtifactOperationLockPath -RepoRoot $testRepo -Scope $releaseDependencyScope)
        $debugPackagingScope = @{ configuration = 'Debug'; platform = 'x64'; coordination = 'packaging' }
        $arm64PackagingScope = @{ configuration = 'Release'; platform = 'ARM64'; coordination = 'packaging' }
        (Get-RSArtifactOperationLockPath -RepoRoot $testRepo -Scope $debugPackagingScope) |
            Should Be (Get-RSArtifactOperationLockPath -RepoRoot $testRepo -Scope $arm64PackagingScope)

        $outer = $null
        $inner = $null
        $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $testRepo -Scope $debugScope
        try {
            $outer = Enter-RSArtifactOperationLock `
                -RepoRoot $testRepo `
                -Operation 'outer-test' `
                -Scope $debugScope
            (Test-Path -LiteralPath $metadataPath) | Should Be $true
            $metadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $testRepo -Scope $debugScope
            $metadata.owner_pid | Should Be $PID
            $metadata.operation | Should Be 'outer-test'
            $metadata.schema | Should Be 'red-salamander.artifact-operation-owner.v4'
            $metadata.scope_key | Should Be 'profile|platform=X64|configuration=DEBUG'
            [System.IO.Path]::GetFullPath([string]$metadata.owner_process_path) |
                Should Be ([System.IO.Path]::GetFullPath((Get-Process -Id $PID).Path))
            [string]::IsNullOrWhiteSpace([string]$metadata.started_utc) | Should Be $false
            (Get-Content -LiteralPath $metadataPath -Raw) |
                Should Match '"started_utc"\s*:\s*"[^"]+Z"'
            (Remove-RSArtifactOperationOwnerMetadata -RepoRoot $testRepo -Token 'not-the-owner' -Scope $debugScope) |
                Should Be $false
            (Test-Path -LiteralPath $metadataPath) | Should Be $true

            $inner = Enter-RSArtifactOperationLock `
                -RepoRoot $testRepo `
                -Operation 'inner-test' `
                -Scope $debugScope
            $outer.WasAbandoned | Should Be $false
            $inner.WasAbandoned | Should Be $false
            $outer.IsRootOwner | Should Be $true
            $inner.IsRootOwner | Should Be $false

            Exit-RSArtifactOperationLock -Lock $inner
            $inner = $null
            (Test-Path -LiteralPath $metadataPath) | Should Be $true
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $inner
            Exit-RSArtifactOperationLock -Lock $outer
        }
        (Test-Path -LiteralPath $metadataPath) | Should Be $false

        $afterRelease = $null
        try {
            $afterRelease = Enter-RSArtifactOperationLock `
                -RepoRoot $testRepo `
                -Operation 'after-release-test' `
                -Scope $debugScope
            $afterRelease.WasAbandoned | Should Be $false
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $afterRelease
        }
    }

    It 'requires an explicit production artifact profile and keeps repository fallback opt-in' {
        $testRepo = Join-Path $TestDrive 'ExplicitScopeRepo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        foreach ($invalidScope in @(@{}, @{ platform = 'x64' })) {
            $message = $null
            try {
                $unexpected = Enter-RSArtifactOperationLock `
                    -RepoRoot $testRepo `
                    -Operation 'invalid-scope' `
                    -Scope $invalidScope
                Exit-RSArtifactOperationLock -Lock $unexpected
            }
            catch {
                $message = $_.Exception.Message
            }
            $message | Should Match 'must provide non-empty, path-safe Scope\.configuration and Scope\.platform'
        }

        $compatibilityLock = $null
        try {
            $compatibilityLock = Enter-RSArtifactOperationLock `
                -RepoRoot $testRepo `
                -Operation 'compatibility-test' `
                -AllowRepositoryScope
            $compatibilityLock.IsRootOwner | Should Be $true
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $compatibilityLock
        }
    }

    It 'rejects unknown coordination kinds instead of silently selecting a profile lock' {
        $testRepo = Join-Path $TestDrive 'InvalidCoordinationRepo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null
        $message = $null
        try {
            $unexpected = Enter-RSArtifactOperationLock `
                -RepoRoot $testRepo `
                -Operation 'typo-scope' `
                -Scope @{
                    configuration = 'Debug'
                    platform = 'x64'
                    coordination = 'shared-dependecy'
                }
            Exit-RSArtifactOperationLock -Lock $unexpected
        }
        catch {
            $message = $_.Exception.Message
        }

        $message | Should Match 'Scope\.coordination must be artifact-profile, shared-dependency, or packaging'
    }

    It 'matches command-line paths only at complete token boundaries' {
        $testRepo = Join-Path $TestDrive 'PathBoundaryRepo'
        $quotedOutput = Join-Path $testRepo '.build\x64\Debug\sample.obj'
        $nearPrefix = "$testRepo-copy\.build\x64\Debug\sample.obj"
        $concatenatedPrefix = "C:\Other$testRepo\.build\x64\Debug\sample.obj"
        $attachedFakePrefix = "C:\Other/Fo$testRepo\.build\x64\Debug\sample.obj"
        $global:RSTestBoundaryRepo = $testRepo
        $global:RSTestQuotedOutput = $quotedOutput
        $global:RSTestNearPrefix = $nearPrefix
        $global:RSTestConcatenatedPrefix = $concatenatedPrefix
        $global:RSTestAttachedFakePrefix = $attachedFakePrefix
        try {
            InModuleScope ArtifactOperationLock {
                (Test-RSArtifactCommandLineContainsPath `
                        -CommandLine ('cl.exe /Fo"' + $global:RSTestQuotedOutput + '"') `
                        -Path $global:RSTestBoundaryRepo) |
                    Should Be $true
                (Test-RSArtifactCommandLineContainsPath `
                        -CommandLine ('cl.exe /Fo' + $global:RSTestBoundaryRepo + '\.build\x64\Debug\sample.obj') `
                        -Path $global:RSTestBoundaryRepo) |
                    Should Be $true
                (Test-RSArtifactCommandLineContainsPath `
                        -CommandLine ('link.exe /OUT:' + $global:RSTestBoundaryRepo + '\.build\x64\Debug\sample.exe') `
                        -Path $global:RSTestBoundaryRepo) |
                    Should Be $true
                (Test-RSArtifactCommandLineContainsPath `
                        -CommandLine ('cl.exe /Fo"' + $global:RSTestNearPrefix + '"') `
                        -Path $global:RSTestBoundaryRepo) |
                    Should Be $false
                (Test-RSArtifactCommandLineContainsPath `
                        -CommandLine ('cl.exe /Fo"' + $global:RSTestConcatenatedPrefix + '"') `
                        -Path $global:RSTestBoundaryRepo) |
                    Should Be $false
                (Test-RSArtifactCommandLineContainsPath `
                        -CommandLine ('cl.exe ' + $global:RSTestAttachedFakePrefix) `
                        -Path $global:RSTestBoundaryRepo) |
                    Should Be $false
            }
        }
        finally {
            foreach ($name in @(
                    'RSTestBoundaryRepo',
                    'RSTestQuotedOutput',
                    'RSTestNearPrefix',
                    'RSTestConcatenatedPrefix',
                    'RSTestAttachedFakePrefix')) {
                Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
            }
        }
    }

    It 'matches combined MSBuild property lists without crossing configuration or platform scope' {
        $testRepo = Join-Path $TestDrive 'CombinedPropertyRepo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null
        $global:RSTestCombinedPropertyRepo = $testRepo
        try {
            InModuleScope ArtifactOperationLock {
                $debugX64 = Get-RSArtifactOperationScopeInfo `
                    -RepoRoot $global:RSTestCombinedPropertyRepo `
                    -Scope @{ configuration = 'Debug'; platform = 'x64' }
                $releaseX64 = Get-RSArtifactOperationScopeInfo `
                    -RepoRoot $global:RSTestCombinedPropertyRepo `
                    -Scope @{ configuration = 'Release'; platform = 'x64' }
                $debugArm64 = Get-RSArtifactOperationScopeInfo `
                    -RepoRoot $global:RSTestCombinedPropertyRepo `
                    -Scope @{ configuration = 'Debug'; platform = 'ARM64' }
                $solution = Join-Path $global:RSTestCombinedPropertyRepo 'RedSalamander.sln'

                (Test-RSArtifactCommandLineMatchesScope `
                        -CommandLine ('MSBuild.exe "' + $solution + '" /p:Configuration=Debug;Platform=x64') `
                        -ScopeInfo $debugX64) |
                    Should Be $true
                (Test-RSArtifactCommandLineMatchesScope `
                        -CommandLine ('MSBuild.exe "' + $solution + '" -p:Configuration=Debug,Platform=x64') `
                        -ScopeInfo $debugX64) |
                    Should Be $true
                (Test-RSArtifactCommandLineMatchesScope `
                        -CommandLine ('MSBuild.exe "' + $solution + '" /p:Configuration=Release;Platform=x64') `
                        -ScopeInfo $debugX64) |
                    Should Be $false
                (Test-RSArtifactCommandLineMatchesScope `
                        -CommandLine ('MSBuild.exe "' + $solution + '" /p:Configuration=Debug;Platform=ARM64') `
                        -ScopeInfo $debugArm64) |
                    Should Be $true
                (Test-RSArtifactCommandLineMatchesScope `
                        -CommandLine ('MSBuild.exe "' + $solution + '" /p:Configuration=Debug;Platform=ARM64') `
                        -ScopeInfo $releaseX64) |
                    Should Be $false
                (Get-RSMSBuildPropertyValue `
                        -Arguments @('/p:Configuration=Release;Platform=x64', '-p:Configuration=Debug') `
                        -Name 'Configuration') |
                    Should Be 'Debug'
                (Get-RSMSBuildPropertyValue `
                        -Arguments @('-p:Configuration=Debug,Platform=x64') `
                        -Name 'Platform') |
                    Should Be 'x64'
            }
        }
        finally {
            Remove-Variable -Name RSTestCombinedPropertyRepo -Scope Global -ErrorAction SilentlyContinue
        }
    }

    It 'requires distinct nested coordination locks to release in same-thread LIFO order' {
        $testRepo = Join-Path $TestDrive 'LifoRepo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null
        $profileScope = @{ configuration = 'Debug'; platform = 'x64' }
        $dependencyScope = @{
            configuration = 'Debug'
            platform = 'x64'
            coordination = 'shared-dependency'
        }
        $profileLock = $null
        $dependencyLock = $null
        try {
            $profileLock = Enter-RSArtifactOperationLock `
                -RepoRoot $testRepo `
                -Operation 'profile' `
                -Scope $profileScope
            $dependencyLock = Enter-RSArtifactOperationLock `
                -RepoRoot $testRepo `
                -Operation 'dependency' `
                -Scope $dependencyScope

            $message = $null
            try {
                Exit-RSArtifactOperationLock -Lock $profileLock
            }
            catch {
                $message = $_.Exception.Message
            }
            $message | Should Match 'same-thread LIFO order'

            Exit-RSArtifactOperationLock -Lock $dependencyLock
            $dependencyLock = $null
            Exit-RSArtifactOperationLock -Lock $profileLock
            $profileLock = $null
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $dependencyLock
            Exit-RSArtifactOperationLock -Lock $profileLock
        }
    }

    It 'rejects stale owner PID reuse when the executable path no longer matches metadata' {
        $currentProcess = Get-Process -Id $PID
        try {
            $global:RSTestOwnerPid = $PID
            $global:RSTestOwnerStartTicks = $currentProcess.StartTime.ToUniversalTime().Ticks
            $global:RSTestOwnerPath = [System.IO.Path]::GetFullPath($currentProcess.Path)
            $global:RSTestWrongOwnerPath = Join-Path $TestDrive 'Foreign\pwsh.exe'
        }
        finally {
            $currentProcess.Dispose()
        }

        try {
            InModuleScope ArtifactOperationLock {
                (Test-RSArtifactOperationOwnerProcess `
                        -OwnerProcessId $global:RSTestOwnerPid `
                        -OwnerProcessStartUtcTicks $global:RSTestOwnerStartTicks `
                        -OwnerProcessPath $global:RSTestOwnerPath) |
                    Should Be $true
                (Test-RSArtifactOperationOwnerProcess `
                        -OwnerProcessId $global:RSTestOwnerPid `
                        -OwnerProcessStartUtcTicks $global:RSTestOwnerStartTicks `
                        -OwnerProcessPath $global:RSTestWrongOwnerPath) |
                    Should Be $false
            }
        }
        finally {
            Remove-Variable -Name RSTestOwnerPid -Scope Global -ErrorAction SilentlyContinue
            Remove-Variable -Name RSTestOwnerStartTicks -Scope Global -ErrorAction SilentlyContinue
            Remove-Variable -Name RSTestOwnerPath -Scope Global -ErrorAction SilentlyContinue
            Remove-Variable -Name RSTestWrongOwnerPath -Scope Global -ErrorAction SilentlyContinue
        }
    }

    It 'creates the process inside the exact containment job and preserves Unicode arguments and redirected streams' {
        $childScript = Join-Path $TestDrive 'ContainedUnicodeChild.ps1'
        $workingDirectory = Join-Path $TestDrive 'work café 漢'
        New-Item -ItemType Directory -Path $workingDirectory -Force | Out-Null
        @'
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Values)
$utf8 = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8
[Console]::Error.WriteLine('stderr-café-漢')
[pscustomobject]@{
    values = @($Values)
    workingDirectory = (Get-Location).Path
} | ConvertTo-Json -Compress
exit 37
'@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $values = @('', 'two words', 'quote"value', 'trailing\', 'café', '漢')
        $pwsh = (Get-Process -Id $PID).Path
        $startInfo = New-RSProcessStartInfo `
            -FilePath $pwsh `
            -Arguments (@('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $childScript) + $values) `
            -WorkingDirectory $workingDirectory
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.CreateNoWindow = $true
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        $startInfo.StandardOutputEncoding = $utf8
        $startInfo.StandardErrorEncoding = $utf8

        $process = $null
        try {
            $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
            $process.IsInContainmentJob | Should Be $true
            $stdoutTask = $process.StandardOutput.ReadToEndAsync()
            $stderrTask = $process.StandardError.ReadToEndAsync()
            $process.WaitForExit(10000) | Should Be $true
            $process.ExitCode | Should Be 37
            $payload = $stdoutTask.GetAwaiter().GetResult().Trim() | ConvertFrom-Json
            $stderrTask.GetAwaiter().GetResult().Trim() | Should Be 'stderr-café-漢'
            (@($payload.values) | ConvertTo-Json -Compress) | Should Be ($values | ConvertTo-Json -Compress)
            [System.IO.Path]::GetFullPath([string]$payload.workingDirectory) | Should Be ([System.IO.Path]::GetFullPath($workingDirectory))
        }
        finally {
            Close-RSContainedProcess -Process $process
        }
    }

    It 'allows only the explicit standard handles to cross process creation' {
        $childScript = Join-Path $TestDrive 'ContainedHandleProbe.ps1'
        @'
param([long]$HandleValue)
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ChildHandleProbe
{
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetEvent(IntPtr handle);
    public static void Probe(long value)
    {
        SetEvent(new IntPtr(value));
    }
}
"@
[ChildHandleProbe]::Probe($HandleValue)
'probe-complete'
'@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $sentinel = [RedSalamander.ToolingTests.InheritableEvent]::new()
        $process = $null
        try {
            $pwsh = (Get-Process -Id $PID).Path
            $startInfo = New-RSProcessStartInfo `
                -FilePath $pwsh `
                -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $childScript, '-HandleValue', [string]$sentinel.HandleValue) `
                -WorkingDirectory $TestDrive
            $startInfo.RedirectStandardOutput = $true
            $startInfo.RedirectStandardError = $true
            $startInfo.CreateNoWindow = $true
            $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
            $stdoutTask = $process.StandardOutput.ReadToEndAsync()
            $stderrTask = $process.StandardError.ReadToEndAsync()
            $process.WaitForExit(10000) | Should Be $true
            $process.ExitCode | Should Be 0
            $stderrTask.GetAwaiter().GetResult() | Should Be ''
            $stdoutTask.GetAwaiter().GetResult().Trim() | Should Be 'probe-complete'
            $sentinel.IsSignaled | Should Be $false
        }
        finally {
            Close-RSContainedProcess -Process $process
            $sentinel.Dispose()
        }
    }

    It 'kills an immediately spawned grandchild when the containment job closes' {
        $childScript = Join-Path $TestDrive 'ContainedTreeChild.ps1'
        $grandchildScript = Join-Path $TestDrive 'ContainedTreeGrandchild.ps1'
        $pidPath = Join-Path $TestDrive 'contained-grandchild.pid'
        'Start-Sleep -Seconds 60' | Set-Content -LiteralPath $grandchildScript -Encoding UTF8
        @'
param([string]$PowerShellPath, [string]$GrandchildScript, [string]$PidPath)
$grandchild = Start-Process -FilePath $PowerShellPath -ArgumentList @('-NoProfile', '-File', ('"' + $GrandchildScript + '"')) -PassThru -WindowStyle Hidden
[System.IO.File]::WriteAllText($PidPath, [string]$grandchild.Id)
while ($true) { Start-Sleep -Milliseconds 100 }
'@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $pwsh = (Get-Process -Id $PID).Path
        $startInfo = New-RSProcessStartInfo `
            -FilePath $pwsh `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $childScript, '-PowerShellPath', $pwsh, '-GrandchildScript', $grandchildScript, '-PidPath', $pidPath) `
            -WorkingDirectory $TestDrive
        $startInfo.CreateNoWindow = $true

        $process = $null
        $rootPid = 0
        $grandchildPid = 0
        try {
            $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
            $rootPid = $process.Id
            $process.IsInContainmentJob | Should Be $true
            for ($attempt = 0; $attempt -lt 200 -and -not (Test-Path -LiteralPath $pidPath); $attempt++) {
                Start-Sleep -Milliseconds 25
            }
            (Test-Path -LiteralPath $pidPath) | Should Be $true
            $grandchildPid = [int](Get-Content -LiteralPath $pidPath -Raw)
            $null -ne (Get-Process -Id $grandchildPid -ErrorAction SilentlyContinue) | Should Be $true
            Close-RSContainedProcess -Process $process
            $process = $null

            for ($attempt = 0; $attempt -lt 200; $attempt++) {
                if ($null -eq (Get-Process -Id $rootPid -ErrorAction SilentlyContinue) -and
                    $null -eq (Get-Process -Id $grandchildPid -ErrorAction SilentlyContinue)) {
                    break
                }
                Start-Sleep -Milliseconds 25
            }
            $null -eq (Get-Process -Id $rootPid -ErrorAction SilentlyContinue) | Should Be $true
            $null -eq (Get-Process -Id $grandchildPid -ErrorAction SilentlyContinue) | Should Be $true
        }
        finally {
            Close-RSContainedProcess -Process $process
            if ($grandchildPid -gt 0) {
                Stop-Process -Id $grandchildPid -Force -ErrorAction SilentlyContinue
            }
            if ($rootPid -gt 0) {
                Stop-Process -Id $rootPid -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'cleans native launch resources after repeated creation failures' {
        $beforeHandles = (Get-Process -Id $PID).HandleCount
        for ($attempt = 0; $attempt -lt 10; $attempt++) {
            $startInfo = New-RSProcessStartInfo `
                -FilePath (Join-Path $TestDrive 'missing-contained-process.exe') `
                -WorkingDirectory $TestDrive
            $launchMessage = $null
            try {
                $unexpected = Start-RSContainedProcess -ProcessStartInfo $startInfo
                Close-RSContainedProcess -Process $unexpected
            }
            catch {
                $launchMessage = $_.Exception.Message
            }
            $launchMessage | Should Match 'CreateProcessW failed'
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        $afterHandles = (Get-Process -Id $PID).HandleCount
        ($afterHandles - $beforeHandles) | Should BeLessThan 24
    }

    It 'rejects a same-profile process while allowing a disjoint configuration concurrently' {
        $testRepo = Join-Path $TestDrive 'ContendedRepo'
        $readyPath = Join-Path $TestDrive 'owner.ready'
        $releasePath = Join-Path $TestDrive 'owner.release'
        $childScript = Join-Path $TestDrive 'HoldArtifactLock.ps1'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $escapedHelper = $artifactLockModule.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")
        $escapedReady = $readyPath.Replace("'", "''")
        $escapedRelease = $releasePath.Replace("'", "''")
        @"
`$ErrorActionPreference = 'Stop'
Import-Module '$escapedHelper' -Force -ErrorAction Stop
`$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'child-owner' -Scope @{ configuration = 'Debug'; platform = 'x64' }
`$dependencyLock = `$null
try {
    `$dependencyLock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'child-vcpkg-owner' -Scope @{ configuration = 'Debug'; platform = 'x64'; coordination = 'shared-dependency' }
    [System.IO.File]::WriteAllText('$escapedReady', 'ready')
    while (-not (Test-Path -LiteralPath '$escapedRelease')) {
        Start-Sleep -Milliseconds 25
    }
}
finally {
    Exit-RSArtifactOperationLock -Lock `$dependencyLock
    Exit-RSArtifactOperationLock -Lock `$lock
}
"@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $pwsh = (Get-Process -Id $PID).Path
        $child = Start-Process -FilePath $pwsh `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$childScript`"") `
            -PassThru `
            -WindowStyle Hidden
        try {
            $ready = $false
            for ($attempt = 0; $attempt -lt 200; $attempt++) {
                if (Test-Path -LiteralPath $readyPath) {
                    $ready = $true
                    break
                }
                Start-Sleep -Milliseconds 25
            }
            $ready | Should Be $true

            $disjoint = $null
            try {
                $disjoint = Enter-RSArtifactOperationLock `
                    -RepoRoot $testRepo `
                    -Operation 'release-owner' `
                    -Scope @{ configuration = 'Release'; platform = 'x64' }
                $disjoint.IsRootOwner | Should Be $true
            }
            finally {
                Exit-RSArtifactOperationLock -Lock $disjoint
            }

            $dependencyMessage = $null
            try {
                $unexpectedDependency = Enter-RSArtifactOperationLock `
                    -RepoRoot $testRepo `
                    -Operation 'release-vcpkg-owner' `
                    -Scope @{ configuration = 'Release'; platform = 'x64'; coordination = 'shared-dependency' }
                Exit-RSArtifactOperationLock -Lock $unexpectedDependency
            }
            catch {
                $dependencyMessage = $_.Exception.Message
            }
            $dependencyMessage | Should Match 'shared-dependency\|platform=X64'

            $message = $null
            try {
                $unexpected = Enter-RSArtifactOperationLock `
                    -RepoRoot $testRepo `
                    -Operation 'second-owner' `
                    -Scope @{ configuration = 'Debug'; platform = 'x64' }
                Exit-RSArtifactOperationLock -Lock $unexpected
            }
            catch {
                $message = $_.Exception.Message
            }
            $message | Should Match 'another RedSalamander build or test operation owns'
            $message | Should Match 'profile\|platform=X64\|configuration=DEBUG'
            $message | Should Match "Owner PID=$($child.Id)"
            $message | Should Match "Operation='child-owner'"
            $message | Should Match "StartedUtc='[^']+'"
        }
        finally {
            New-Item -ItemType File -Path $releasePath -Force | Out-Null
            if (-not $child.WaitForExit(5000)) {
                Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
            }
            $child.Dispose()
        }
    }

    It 'rejects a parallel runspace in the owner process instead of treating it as local nesting' {
        $testRepo = Join-Path $TestDrive 'ParallelRunspaceRepo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null
        $escapedHelper = $artifactLockModule.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")

        $outer = $null
        $parallelPowerShell = $null
        try {
            $outer = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'owning-runspace' -AllowRepositoryScope
            $parallelPowerShell = [PowerShell]::Create()
            [void]$parallelPowerShell.AddScript(@"
`$ErrorActionPreference = 'Stop'
Import-Module '$escapedHelper' -Force -ErrorAction Stop
try {
    `$unexpected = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'parallel-runspace' -AllowRepositoryScope
    Exit-RSArtifactOperationLock -Lock `$unexpected
    Write-Output 'ACQUIRED'
}
catch {
    Write-Output ('ERROR: ' + `$_.Exception.Message)
}
"@)

            $parallelResult = @($parallelPowerShell.Invoke()) -join "`n"
            $parallelResult | Should Match 'another RedSalamander build or test operation owns'
            $parallelResult | Should Not Match 'ACQUIRED'
        }
        finally {
            if ($null -ne $parallelPowerShell) {
                $parallelPowerShell.Dispose()
            }
            Exit-RSArtifactOperationLock -Lock $outer
        }
    }

    It 'delegates only to a direct child after kill-on-close containment is active' {
        $testRepo = Join-Path $TestDrive 'InheritedRepo'
        $childScript = Join-Path $TestDrive 'InheritArtifactLock.ps1'
        $resultPath = Join-Path $TestDrive 'inherited-result.txt'
        $releasePath = Join-Path $TestDrive 'inherited-release.txt'
        $errorPath = Join-Path $TestDrive 'inherited-error.txt'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $escapedEnvironment = $sanitizedEnvironmentModule.Replace("'", "''")
        $escapedHelper = $artifactLockModule.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")
        $escapedResult = $resultPath.Replace("'", "''")
        $escapedRelease = $releasePath.Replace("'", "''")
        $escapedError = $errorPath.Replace("'", "''")
        @"
`$ErrorActionPreference = 'Stop'
trap {
    [System.IO.File]::AppendAllText('$escapedError', "`n" + `$_.Exception.ToString())
    exit 1
}
Import-Module '$escapedEnvironment' -Force -ErrorAction Stop
Import-Module '$escapedHelper' -Force -ErrorAction Stop
`$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'descendant-build' -AllowRepositoryScope
`$nestedLock = `$null
`$secondHop = `$null
try {
    `$nestedLock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'nested-descendant-build' -AllowRepositoryScope
    `$secondHop = New-RSArtifactOperationChildDelegation
    `$secondHopBlocked = `$null -eq `$secondHop
    [System.IO.File]::WriteAllText(
        '$escapedResult',
        ('{0}|{1}|{2}' -f `$lock.IsDelegated, `$nestedLock.IsDelegated, `$secondHopBlocked))
    while (-not (Test-Path -LiteralPath '$escapedRelease')) {
        Start-Sleep -Milliseconds 25
    }
}
finally {
    Complete-RSArtifactOperationChildDelegation -Delegation `$secondHop
    Exit-RSArtifactOperationLock -Lock `$nestedLock
    Exit-RSArtifactOperationLock -Lock `$lock
}
"@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $testRepo
        $rootLock = $null
        try {
            $rootLock = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'root-runner' -AllowRepositoryScope
            $rootMetadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $testRepo

            $delegationProbe = New-RSArtifactOperationChildDelegation
            if ($null -eq $delegationProbe) {
                throw 'The owning thread could not create a contained-child delegation probe.'
            }
            Complete-RSArtifactOperationChildDelegation -Delegation $delegationProbe

            $pwsh = (Get-Process -Id $PID).Path
            $startInfo = New-RSProcessStartInfo `
                -FilePath $pwsh `
                -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $childScript) `
                -WorkingDirectory $testRepo
            $startInfo.CreateNoWindow = $true
            $child = $null
            try {
                $child = Start-RSContainedProcess `
                    -ProcessStartInfo $startInfo `
                    -DelegateArtifactOperation

                $resultReady = $false
                for ($attempt = 0; $attempt -lt 600; $attempt++) {
                    if (Test-Path -LiteralPath $resultPath) {
                        $resultReady = $true
                        break
                    }
                    if ($child.HasExited) { break }
                    Start-Sleep -Milliseconds 25
                }
                if (-not $resultReady) {
                    $childError = if (Test-Path -LiteralPath $errorPath) {
                        Get-Content -LiteralPath $errorPath -Raw
                    } else {
                        '<no child error record>'
                    }
                    throw "Delegated child did not publish its result. Exited=$($child.HasExited) ExitCode=$(if ($child.HasExited) { $child.ExitCode } else { '<running>' }) Error=$childError"
                }
                $releaseMessage = $null
                try {
                    Exit-RSArtifactOperationLock -Lock $rootLock
                }
                catch {
                    $releaseMessage = $_.Exception.Message
                }
                $releaseMessage | Should Match 'contained child delegation.*remain active'
                (Test-Path -LiteralPath $metadataPath) | Should Be $true

                New-Item -ItemType File -Path $releasePath -Force | Out-Null
                $child.WaitForExit(5000) | Should Be $true
                $child.ExitCode | Should Be 0
            }
            finally {
                New-Item -ItemType File -Path $releasePath -Force | Out-Null
                Close-RSContainedProcess -Process $child
            }

            (Get-Content -LiteralPath $resultPath -Raw) | Should Be 'True|True|True'
            (Test-Path -LiteralPath $metadataPath) | Should Be $true
            $metadataAfterChild = Read-RSArtifactOperationOwnerMetadata -RepoRoot $testRepo
            $metadataAfterChild.token | Should Be $rootMetadata.token
            $metadataAfterChild.operation | Should Be 'root-runner'
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $rootLock
        }

        (Test-Path -LiteralPath $metadataPath) | Should Be $false
    }

    It 'makes a surviving stale descendant acquire abandoned ownership instead of bypassing it' {
        $testRepo = Join-Path $TestDrive 'StaleDescendantRepo'
        $ownerScript = Join-Path $TestDrive 'LaunchStaleDescendant.ps1'
        $survivorScript = Join-Path $TestDrive 'StaleDescendant.ps1'
        $survivorReadyPath = Join-Path $TestDrive 'stale-descendant.ready'
        $survivorResultPath = Join-Path $TestDrive 'stale-descendant.result'
        $survivorReleasePath = Join-Path $TestDrive 'stale-descendant.release'
        $survivorPidPath = Join-Path $TestDrive 'stale-descendant.pid'
        $ownerStdoutPath = Join-Path $TestDrive 'stale-owner.stdout.log'
        $ownerStderrPath = Join-Path $TestDrive 'stale-owner.stderr.log'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $escapedHelper = $artifactLockModule.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")
        $escapedSurvivor = $survivorScript.Replace("'", "''")
        $escapedReady = $survivorReadyPath.Replace("'", "''")
        $escapedResult = $survivorResultPath.Replace("'", "''")
        $escapedRelease = $survivorReleasePath.Replace("'", "''")
        $escapedPid = $survivorPidPath.Replace("'", "''")
        @"
param([int]`$OwnerPid)
`$ErrorActionPreference = 'Stop'
Import-Module '$escapedHelper' -Force -ErrorAction Stop
[System.IO.File]::WriteAllText('$escapedReady', 'ready')
while (`$null -ne (Get-Process -Id `$OwnerPid -ErrorAction SilentlyContinue)) {
    Start-Sleep -Milliseconds 25
}
`$lock = `$null
try {
    `$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'surviving-descendant' -AllowRepositoryScope -TimeoutMilliseconds 5000
    [System.IO.File]::WriteAllText(
        '$escapedResult',
        "`$(`$lock.IsDelegated)|`$(`$lock.WasAbandoned)")
    while (-not (Test-Path -LiteralPath '$escapedRelease')) {
        Start-Sleep -Milliseconds 25
    }
}
catch {
    [System.IO.File]::WriteAllText('$escapedResult', 'ERROR: ' + `$_.Exception.Message)
}
finally {
    Exit-RSArtifactOperationLock -Lock `$lock
}
"@ | Set-Content -LiteralPath $survivorScript -Encoding UTF8

        @"
`$ErrorActionPreference = 'Stop'
Import-Module '$escapedHelper' -Force -ErrorAction Stop
`$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'doomed-owner' -AllowRepositoryScope
`$pwsh = (Get-Process -Id `$PID).Path
`$survivor = Start-Process -FilePath `$pwsh -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', '"$escapedSurvivor"', '-OwnerPid', `$PID) -PassThru -WindowStyle Hidden
[System.IO.File]::WriteAllText('$escapedPid', [string]`$survivor.Id)
while (-not (Test-Path -LiteralPath '$escapedReady')) {
    Start-Sleep -Milliseconds 25
}
[Environment]::Exit(0)
"@ | Set-Content -LiteralPath $ownerScript -Encoding UTF8

        $pwsh = (Get-Process -Id $PID).Path
        $owner = Start-Process -FilePath $pwsh `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$ownerScript`"") `
            -PassThru `
            -RedirectStandardOutput $ownerStdoutPath `
            -RedirectStandardError $ownerStderrPath `
            -WindowStyle Hidden
        $survivorProcess = $null
        try {
            $owner.WaitForExit(5000) | Should Be $true
            $owner.WaitForExit()
            $owner.Refresh()
            if ($null -ne $owner.ExitCode -and $owner.ExitCode -ne 0) {
                $ownerStdout = if (Test-Path -LiteralPath $ownerStdoutPath) {
                    Get-Content -LiteralPath $ownerStdoutPath -Raw
                } else { '' }
                $ownerStderr = if (Test-Path -LiteralPath $ownerStderrPath) {
                    Get-Content -LiteralPath $ownerStderrPath -Raw
                } else { '' }
                throw "Stale owner fixture exited $($owner.ExitCode). stdout='$ownerStdout' stderr='$ownerStderr'"
            }

            $resultReady = $false
            for ($attempt = 0; $attempt -lt 200; $attempt++) {
                if (Test-Path -LiteralPath $survivorResultPath) {
                    $resultReady = $true
                    break
                }
                Start-Sleep -Milliseconds 25
            }
            $resultReady | Should Be $true
            (Get-Content -LiteralPath $survivorResultPath -Raw) | Should Be 'False|True'

            $survivorPid = [int](Get-Content -LiteralPath $survivorPidPath -Raw)
            $survivorProcess = Get-Process -Id $survivorPid -ErrorAction Stop
            $message = $null
            try {
                $unexpected = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'observer' -AllowRepositoryScope
                Exit-RSArtifactOperationLock -Lock $unexpected
            }
            catch {
                $message = $_.Exception.Message
            }
            $message | Should Match 'another RedSalamander build or test operation owns'
            $message | Should Match "Owner PID=$survivorPid"
        }
        finally {
            New-Item -ItemType File -Path $survivorReleasePath -Force | Out-Null
            if ($null -ne $survivorProcess) {
                if (-not $survivorProcess.WaitForExit(5000)) {
                    Stop-Process -Id $survivorProcess.Id -Force -ErrorAction SilentlyContinue
                }
                $survivorProcess.Dispose()
            }
            if (-not $owner.HasExited) {
                Stop-Process -Id $owner.Id -Force -ErrorAction SilentlyContinue
            }
            $owner.Dispose()
        }
    }

    It 'reports abandoned ownership and persists contamination until explicitly cleared' {
        $testRepo = Join-Path $TestDrive 'AbandonedRepo'
        $childScript = Join-Path $TestDrive 'AbandonArtifactLock.ps1'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $escapedHelper = $artifactLockModule.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")
        @"
`$ErrorActionPreference = 'Stop'
Import-Module '$escapedHelper' -Force -ErrorAction Stop
`$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'abandoned-owner' -AllowRepositoryScope
[Environment]::Exit(0)
"@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $pwsh = (Get-Process -Id $PID).Path
        $child = Start-Process -FilePath $pwsh `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$childScript`"") `
            -PassThru `
            -WindowStyle Hidden
        $lock = $null
        try {
            $child.WaitForExit(5000) | Should Be $true
            $child.ExitCode | Should Be 0

            $lock = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'abandonment-observer' -AllowRepositoryScope
            $lock.WasAbandoned | Should Be $true
            $markerPath = Set-RSArtifactOperationContaminated -RepoRoot $testRepo -Reason 'test abandonment'
            (Test-Path -LiteralPath $markerPath) | Should Be $true
            (Test-RSArtifactOperationContaminated -RepoRoot $testRepo) | Should Be $true
            Clear-RSArtifactOperationContaminated -RepoRoot $testRepo
            (Test-RSArtifactOperationContaminated -RepoRoot $testRepo) | Should Be $false
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $lock
            if (-not $child.HasExited) {
                Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
            }
            $child.Dispose()
        }

        $legacyRepo = Join-Path $TestDrive 'LegacyOwnerRepo'
        $legacyMetadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $legacyRepo
        New-Item -ItemType Directory -Path (Split-Path -Parent $legacyMetadataPath) -Force | Out-Null
        [System.IO.File]::WriteAllText(
            $legacyMetadataPath,
            '{"schema":"red-salamander.artifact-operation-owner.v1","owner_pid":82648}',
            [System.Text.UTF8Encoding]::new($false))

        $legacyLock = $null
        try {
            $legacyLock = Enter-RSArtifactOperationLock -RepoRoot $legacyRepo -Operation 'legacy-sidecar-observer' -AllowRepositoryScope
            $legacyLock.WasAbandoned | Should Be $true
            (Read-RSArtifactOperationOwnerMetadata -RepoRoot $legacyRepo).schema |
                Should Be 'red-salamander.artifact-operation-owner.v4'
            [void](Set-RSArtifactOperationContaminated -RepoRoot $legacyRepo -Reason 'stale v1 owner')
            (Test-RSArtifactOperationContaminated -RepoRoot $legacyRepo) | Should Be $true
            Clear-RSArtifactOperationContaminated -RepoRoot $legacyRepo
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $legacyLock
        }
    }

    It 'rejects only residual tools proven to target the same repository and output profile' {
        $testRepo = Join-Path $TestDrive 'ResidualRepo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null
        Mock Get-CimInstance -ModuleName ArtifactOperationLock {
            $testRepo = $global:RSTestResidualArtifactRepo
            return @(
                [pscustomobject]@{
                    ProcessId = 7711
                    ParentProcessId = 7700
                    Name = 'cl.exe'
                    CommandLine = 'cl.exe /Fo"' + $testRepo + '\.build\Intermediate\x64\Debug\Project\sample.obj" "' + $testRepo + '\sample.cpp"'
                },
                [pscustomobject]@{
                    ProcessId = 7722
                    ParentProcessId = 7700
                    Name = 'link.exe'
                    CommandLine = 'link.exe /OUT:"' + $testRepo + '\.build\x64\Release\RedSalamander.exe"'
                },
                [pscustomobject]@{
                    ProcessId = 8811
                    ParentProcessId = 8800
                    Name = 'MSBuild.exe'
                    CommandLine = 'MSBuild.exe C:\OtherRepo\Other.sln /p:Configuration=Debug /p:Platform=x64'
                },
                [pscustomobject]@{
                    ProcessId = 9911
                    ParentProcessId = 9900
                    Name = 'MSBuild.exe'
                    CommandLine = 'MSBuild.exe "' + $testRepo + '\RedSalamander.sln" /p:Configuration=Release /p:Platform=x64'
                },
                [pscustomobject]@{
                    ProcessId = 9922
                    ParentProcessId = 9900
                    Name = 'MSBuild.exe'
                    CommandLine = 'MSBuild.exe "' + $testRepo + '\RedSalamander.sln" /p:Configuration=Release /p:Platform=ARM64'
                },
                [pscustomobject]@{
                    ProcessId = 9933
                    ParentProcessId = 9900
                    Name = 'vcpkg.exe'
                    CommandLine = 'vcpkg.exe install --triplet x64-windows --x-manifest-root "' + $testRepo + '" --x-install-root "' + $testRepo + '\.build\vcpkg_install_staging\x64-windows"'
                },
                [pscustomobject]@{
                    ProcessId = 9934
                    ParentProcessId = 9900
                    Name = 'vcpkg.exe'
                    CommandLine = 'vcpkg.exe install --triplet x64-windows --x-manifest-root "' + $testRepo + '" --x-install-root "' + $testRepo + '\.build\vcpkg_installed\x64"'
                },
                [pscustomobject]@{
                    ProcessId = 9944
                    ParentProcessId = 9900
                    Name = 'vcpkg.exe'
                    CommandLine = 'vcpkg.exe install --triplet x64-windows --x-manifest-root "C:\OtherRepo" --x-install-root "C:\OtherRepo\.build\vcpkg_installed\x64-windows"'
                },
                [pscustomobject]@{
                    ProcessId = 9955
                    ParentProcessId = 9900
                    Name = 'vcpkg.exe'
                    CommandLine = 'vcpkg.exe install --triplet=x64-windows-static-md --x-manifest-root "' + $testRepo + '" --x-install-root="' + $testRepo + '\.build\vcpkg_install_staging\x64-windows-static-md"'
                },
                [pscustomobject]@{
                    ProcessId = 9966
                    ParentProcessId = 9900
                    Name = 'vcpkg.exe'
                    CommandLine = 'vcpkg.exe install --triplet arm64-windows-static-md --x-manifest-root "' + $testRepo + '" --x-install-root "' + $testRepo + '\.build\vcpkg_install_staging\arm64-windows-static-md"'
                },
                [pscustomobject]@{
                    ProcessId = 9967
                    ParentProcessId = 9900
                    Name = 'vcpkg.exe'
                    CommandLine = 'vcpkg.exe install --triplet arm64-windows --x-manifest-root "' + $testRepo + '" --x-install-root "' + $testRepo + '\.build\vcpkg_installed\arm64"'
                },
                [pscustomobject]@{
                    ProcessId = 9977
                    ParentProcessId = 9900
                    Name = 'vcpkg.exe'
                    CommandLine = 'vcpkg.exe install --triplet x64-windows-static-md --x-manifest-root "C:\ForeignRepo" --x-install-root "C:\ForeignRepo\.build\vcpkg_install_staging\x64-windows-static-md"'
                },
                [pscustomobject]@{
                    ProcessId = 9988
                    ParentProcessId = 9900
                    Name = 'cl.exe'
                    CommandLine = 'cl.exe /I"' + $testRepo + '\.build\vcpkg_installed\x64\x64-windows-static-md\include" "' + $testRepo + '\sample.cpp"'
                },
                [pscustomobject]@{
                    ProcessId = 9999
                    ParentProcessId = 9900
                    Name = 'link.exe'
                    CommandLine = 'link.exe /LIBPATH:"' + $testRepo + '\.build\vcpkg_install_staging\arm64-windows-static-md\lib"'
                },
                [pscustomobject]@{
                    ProcessId = 10011
                    ParentProcessId = 9900
                    Name = 'cl.exe'
                    CommandLine = 'cl.exe /I"C:\ForeignRepo\.build\vcpkg_installed\x64-windows-static-md\include"'
                }
            )
        }

        $global:RSTestResidualArtifactRepo = $testRepo
        $message = $null
        try {
            Assert-RSNoResidualArtifactToolProcesses `
                -RepoRoot $testRepo `
                -Scope @{ configuration = 'Debug'; platform = 'x64' }
        }
        catch {
            $message = $_.Exception.Message
        }

        $message | Should Match 'build tool is still touching artifact scope'
        $message | Should Match 'profile\|platform=X64\|configuration=DEBUG'
        $message | Should Match 'PID=7711'
        $message | Should Not Match 'PID=7722'
        $message | Should Not Match 'PID=8811'

        $sharedDependencyMessage = $null
        try {
            Assert-RSNoResidualArtifactToolProcesses `
                -RepoRoot $testRepo `
                -Scope @{ configuration = 'Debug'; platform = 'x64'; coordination = 'shared-dependency' }
        }
        catch {
            $sharedDependencyMessage = $_.Exception.Message
        }
        $sharedDependencyMessage | Should Match 'PID=9911'
        $sharedDependencyMessage | Should Match 'PID=9933'
        $sharedDependencyMessage | Should Match 'PID=9934'
        $sharedDependencyMessage | Should Match 'PID=9955'
        $sharedDependencyMessage | Should Match 'PID=9988'
        $sharedDependencyMessage | Should Not Match 'PID=9922'
        $sharedDependencyMessage | Should Not Match 'PID=9944'
        $sharedDependencyMessage | Should Not Match 'PID=9966'
        $sharedDependencyMessage | Should Not Match 'PID=9967'
        $sharedDependencyMessage | Should Not Match 'PID=9977'
        $sharedDependencyMessage | Should Not Match 'PID=9999'
        $sharedDependencyMessage | Should Not Match 'PID=10011'

        $arm64SharedDependencyMessage = $null
        try {
            Assert-RSNoResidualArtifactToolProcesses `
                -RepoRoot $testRepo `
                -Scope @{ configuration = 'Debug'; platform = 'ARM64'; coordination = 'shared-dependency' }
        }
        catch {
            $arm64SharedDependencyMessage = $_.Exception.Message
        }
        $arm64SharedDependencyMessage | Should Match 'PID=9922'
        $arm64SharedDependencyMessage | Should Match 'PID=9966'
        $arm64SharedDependencyMessage | Should Match 'PID=9967'
        $arm64SharedDependencyMessage | Should Match 'PID=9999'
        $arm64SharedDependencyMessage | Should Not Match 'PID=9911'
        $arm64SharedDependencyMessage | Should Not Match 'PID=9955'
        $arm64SharedDependencyMessage | Should Not Match 'PID=9977'
        $arm64SharedDependencyMessage | Should Not Match 'PID=9988'
        $arm64SharedDependencyMessage | Should Not Match 'PID=10011'
        Remove-Variable -Name RSTestResidualArtifactRepo -Scope Global -ErrorAction SilentlyContinue
    }
}
