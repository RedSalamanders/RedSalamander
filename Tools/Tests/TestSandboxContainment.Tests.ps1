# Fresh containment regression coverage: run with the marked test sandbox.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$sandboxModulePath = Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1'
Import-Module $sandboxModulePath -ErrorAction Stop
$containmentModule = Get-Module | Where-Object { $_.Path -eq $sandboxModulePath }
if (@($containmentModule).Count -ne 1) { throw 'Expected one canonical TestSandbox module.' }
$script:RSContainmentOwnedLinks = [Collections.Generic.List[object]]::new()

function New-RSContainmentRegressionFixture {
    $root = Get-RSTestSandboxRoot -RepoRoot $repoRoot
    # Require a pre-existing valid marker before the allocator can initialize anything.
    [void](Assert-RSTestSandboxContainedPath -TestRoot $root -Path $root)
    $directory = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot `
        -Harness sandbox-containment -Case ([Guid]::NewGuid().ToString('N'))
    [void](Assert-RSTestSandboxContainedPath -TestRoot $root -Path $directory -RequireDescendant)
    return [pscustomobject]@{ Root = $root; Directory = $directory }
}

function New-RSContainmentOrdinaryDirectory {
    param($Fixture, [string]$Path)
    [void](Assert-RSTestSandboxContainedPath -TestRoot $Fixture.Root -Path $Path -RequireDescendant)
    [void][IO.Directory]::CreateDirectory($Path)
    [void](Assert-RSTestSandboxContainedPath -TestRoot $Fixture.Root -Path $Path -RequireDescendant)
}

function New-RSContainmentReparseFixture {
    param($Fixture, [string]$Link, [string]$Target, [string]$ItemType)
    # The link is intentionally unsafe after creation. Authorize its ordinary parent first.
    [void](Assert-RSTestSandboxContainedPath -TestRoot $Fixture.Root `
        -Path ([IO.Path]::GetDirectoryName($Link)) -RequireDescendant)
    [void](Assert-RSTestSandboxLinkTarget -Path $Link -Target $Target)
    New-RSTestSandboxLink -ItemType $ItemType -Path $Link -Target $Target | Out-Null
    (([IO.File]::GetAttributes($Link) -band [IO.FileAttributes]::ReparsePoint) -ne 0) | Should Be $true
    $item = Get-Item -LiteralPath $Link -Force
    $script:RSContainmentOwnedLinks.Add([pscustomobject]@{
        Root = $Fixture.Root; Directory = $Fixture.Directory; Path = $Link
        Target = $item.LinkTarget; Created = $item.CreationTimeUtc.ToFileTimeUtc(); Attributes = [int]$item.Attributes
    })
}

function Complete-RSContainmentFixtureLinks {
    foreach ($link in $script:RSContainmentOwnedLinks) {
        [void](Assert-RSTestSandboxContainedPath -TestRoot $link.Root -Path ([IO.Path]::GetDirectoryName($link.Path)) -RequireDescendant)
        try { $attributes = [IO.File]::GetAttributes($link.Path) }
        catch [IO.FileNotFoundException] { continue }
        catch [IO.DirectoryNotFoundException] { continue }
        $item = Get-Item -LiteralPath $link.Path -Force
        if (($attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 -or
            $item.LinkTarget -cne $link.Target -or $item.CreationTimeUtc.ToFileTimeUtc() -ne $link.Created) {
            throw "Synthetic containment link changed ownership; retaining it: $($link.Path)"
        }
        # Keep the observed negative fixture and unlink only the exact link this test created.
        $receipt = Join-Path $link.Directory ('retired-test-link-' + [guid]::NewGuid().ToString('N') + '.json')
        [void](Assert-RSTestSandboxContainedPath -TestRoot $link.Root -Path $receipt -RequireDescendant)
        [IO.File]::WriteAllText($receipt, ($link | ConvertTo-Json -Compress))
        if (($attributes -band [IO.FileAttributes]::Directory) -ne 0) { [IO.Directory]::Delete($link.Path, $false) }
        else { [IO.File]::Delete($link.Path) }
    }
    $script:RSContainmentOwnedLinks.Clear()
}

function Invoke-RSContainmentNamedAttributeFault {
    param($Fixture, [string]$ExceptionType)
    $faultPath = Join-Path $Fixture.Directory 'injected-attribute-error'
    $original = & $containmentModule {
        (Get-Item Function:Assert-RSTestSandboxPathIsReparseFree -ErrorAction Stop).ScriptBlock
    }
    $sourceHash = (Get-FileHash -LiteralPath $sandboxModulePath -Algorithm SHA256).Hash
    $statement = '$attributes = [IO.File]::GetAttributes($current)'
    $body = $original.ToString()
    ([regex]::Matches($body, [regex]::Escape($statement)).Count) | Should Be 1
    $fault = @'
if ([string]::Equals($current, $script:RSContainmentAttributeFaultPath, [StringComparison]::OrdinalIgnoreCase)) {
    throw __EXCEPTION__::new('RSContainment injected non-not-found attribute failure')
}
$attributes = [IO.File]::GetAttributes($current)
'@
    $fault = $fault.Replace('__EXCEPTION__', $ExceptionType)
    $replacement = $body.Replace($statement, $fault)
    $previousVariable = & $containmentModule {
        $value = Get-Variable -Name RSContainmentAttributeFaultPath -Scope Script -ErrorAction SilentlyContinue
        if ($null -eq $value) { return [pscustomobject]@{ Present = $false; Value = $null } }
        return [pscustomobject]@{ Present = $true; Value = $value.Value }
    }
    try {
        & $containmentModule {
            param($Text, $FaultPath)
            $script:RSContainmentAttributeFaultPath = $FaultPath
            Set-Item Function:script:Assert-RSTestSandboxPathIsReparseFree -Value ([scriptblock]::Create($Text))
        } $replacement $faultPath
        # Full authority/marker validation still runs; only this unique leaf read is injected.
        { Assert-RSTestSandboxContainedPath -TestRoot $Fixture.Root -Path $faultPath -RequireDescendant } |
            Should Throw 'RSContainment injected non-not-found attribute failure'
    } finally {
        & $containmentModule {
            param($SavedFunction, $SavedVariable)
            Set-Item Function:script:Assert-RSTestSandboxPathIsReparseFree -Value $SavedFunction
            if ($SavedVariable.Present) {
                $script:RSContainmentAttributeFaultPath = $SavedVariable.Value
            } else {
                Remove-Variable -Name RSContainmentAttributeFaultPath -Scope Script -ErrorAction SilentlyContinue
            }
        } $original $previousVariable
        (& $containmentModule {
            (Get-Item Function:Assert-RSTestSandboxPathIsReparseFree -ErrorAction Stop).ScriptBlock.ToString()
        }) | Should Be $body
        (Get-FileHash -LiteralPath $sandboxModulePath -Algorithm SHA256).Hash | Should Be $sourceHash
    }
}

Describe 'TestSandbox fresh ordinary-path containment' {
    AfterEach { Complete-RSContainmentFixtureLinks }
    It 'preserves exact marked root, lexical descendant and ADS authority' {
        $fixture = New-RSContainmentRegressionFixture
        $marker = Join-Path $fixture.Root '.red-salamander-test-root'
        $markerHash = (Get-FileHash -LiteralPath $marker -Algorithm SHA256).Hash
        try {
            $leaf = Join-Path $fixture.Directory 'ordinary-output.txt'
            (Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $leaf -RequireDescendant) |
                Should Be ([IO.Path]::GetFullPath($leaf))
            { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $fixture.Root -RequireDescendant } |
                Should Throw 'must remain below'
            $sibling = $fixture.Root + '-sibling\output.txt'
            { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $sibling } |
                Should Throw 'must remain below'
            $escape = Join-Path $fixture.Root '..\outside-output.txt'
            { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $escape } |
                Should Throw 'must remain below'
            { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path ($leaf + ':stream') } |
                Should Throw 'alternate-data-stream'
            { Assert-RSTestSandboxContainedPath -TestRoot ($fixture.Root + ':stream') -Path $leaf } |
                Should Throw 'alternate-data-stream'
            { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Directory -Path $leaf } |
                Should Throw 'Use exact'
            { Assert-RSTestSandboxContainedPath -TestRoot ([IO.Path]::GetPathRoot($fixture.Root)) -Path $leaf } |
                Should Throw 'Use exact'
        } finally {
            (Get-FileHash -LiteralPath $marker -Algorithm SHA256).Hash | Should Be $markerHash
        }
    }

    It 'rejects an injected missing marker without changing the real marker' {
        $fixture = New-RSContainmentRegressionFixture
        $marker = Join-Path $fixture.Root '.red-salamander-test-root'
        $markerHash = (Get-FileHash -LiteralPath $marker -Algorithm SHA256).Hash
        $markerFunction = & $containmentModule {
            (Get-Item Function:Test-RSTestSandboxMarker -ErrorAction Stop).ScriptBlock
        }
        try {
            # Pester 3 module mocks can remain active for later It cases. Restore
            # this explicitly labeled private injection before returning.
            & $containmentModule {
                Set-Item Function:script:Test-RSTestSandboxMarker -Value { param([string]$TestRoot) return $false }
            }
            { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $fixture.Directory } |
                Should Throw 'not authorized'
        } finally {
            & $containmentModule {
                param($SavedFunction)
                Set-Item Function:script:Test-RSTestSandboxMarker -Value $SavedFunction
            } $markerFunction
            (& $containmentModule {
                (Get-Item Function:Test-RSTestSandboxMarker -ErrorAction Stop).ScriptBlock.ToString()
            }) | Should Be $markerFunction.ToString()
            (Get-FileHash -LiteralPath $marker -Algorithm SHA256).Hash | Should Be $markerHash
        }
    }

    It 'permits multiple missing descendants while checking their existing ancestors' {
        $fixture = New-RSContainmentRegressionFixture
        $missing = Join-Path $fixture.Directory 'missing-a\missing-b\new-output.txt'
        (Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $missing -RequireDescendant) |
            Should Be ([IO.Path]::GetFullPath($missing))
        [IO.Directory]::Exists((Join-Path $fixture.Directory 'missing-a')) | Should Be $false
        [IO.File]::Exists($missing) | Should Be $false
    }

    It 'preserves existing Unicode and deep ordinary paths and permits their missing leaves' {
        $fixture = New-RSContainmentRegressionFixture
        $directory = $fixture.Directory
        foreach ($index in 1..16) {
            $directory = Join-Path $directory ('archives-familiales-été-{0:D2}' -f $index)
            New-RSContainmentOrdinaryDirectory -Fixture $fixture -Path $directory
        }
        $file = Join-Path $directory 'résumé-français.txt'
        $file.Length | Should BeGreaterThan 260
        [void](Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $file -RequireDescendant)
        [IO.File]::WriteAllText($file, 'octets français: été, œuvre, à bientôt', [Text.UTF8Encoding]::new($false))
        (Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $file -RequireDescendant) |
            Should Be ([IO.Path]::GetFullPath($file))
        $missing = Join-Path $directory 'absent\nouveau.txt'
        (Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $missing -RequireDescendant) |
            Should Be ([IO.Path]::GetFullPath($missing))
    }

    It 'rejects an existing junction leaf and ancestor even when the descendant is absent' {
        $fixture = New-RSContainmentRegressionFixture
        $target = Join-Path $fixture.Directory 'target'
        New-RSContainmentOrdinaryDirectory -Fixture $fixture -Path $target
        $link = Join-Path $fixture.Directory 'junction'
        New-RSContainmentReparseFixture -Fixture $fixture -Link $link -Target $target -ItemType Junction
        { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $link } | Should Throw 'reparse-point'
        { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path (Join-Path $link 'missing\leaf') } |
            Should Throw 'reparse-point'
    }

    It 'rejects a real directory symbolic-link ancestor' {
        $fixture = New-RSContainmentRegressionFixture
        $target = Join-Path $fixture.Directory 'directory-target'
        New-RSContainmentOrdinaryDirectory -Fixture $fixture -Path $target
        $link = Join-Path $fixture.Directory 'directory-link'
        New-RSContainmentReparseFixture -Fixture $fixture -Link $link -Target $target -ItemType SymbolicLink
        { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path (Join-Path $link 'missing.txt') } |
            Should Throw 'reparse-point'
    }

    It 'rejects a real file symbolic link and a dangling symbolic-link component' {
        $fixture = New-RSContainmentRegressionFixture
        $target = Join-Path $fixture.Directory 'target.txt'
        [void](Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $target -RequireDescendant)
        [IO.File]::WriteAllText($target, 'real link target')
        $link = Join-Path $fixture.Directory 'file-link.txt'
        New-RSContainmentReparseFixture -Fixture $fixture -Link $link -Target $target -ItemType SymbolicLink
        { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $link } | Should Throw 'reparse-point'
        $dangling = Join-Path $fixture.Directory 'dangling-link'
        New-RSContainmentReparseFixture -Fixture $fixture -Link $dangling `
            -Target (Join-Path $fixture.Directory 'absent-target') -ItemType SymbolicLink
        { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $dangling } | Should Throw 'reparse-point'
        { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path (Join-Path $dangling 'missing.txt') } |
            Should Throw 'reparse-point'
    }

    It 'refuses a test link that would live or point outside its marked sandbox, before creating it' {
        # Q12: New-RSTestSandboxLink is the only governed PowerShell creator of test links.
        $fixture = New-RSContainmentRegressionFixture
        $target = Join-Path $fixture.Directory 'inside-target'
        New-RSContainmentOrdinaryDirectory -Fixture $fixture -Path $target
        $link = Join-Path $fixture.Directory 'refused-link'
        $refusals = [ordered]@{
            'system-directory' = @{ ItemType = 'Junction'; Path = $link; Target = $env:SystemRoot }
            'sandbox-base' = @{ ItemType = 'Junction'; Path = $link; Target = $fixture.Root }
            'volume-root' = @{ ItemType = 'Junction'; Path = $link; Target = [IO.Path]::GetPathRoot($fixture.Root) }
            'relative-climb' = @{ ItemType = 'SymbolicLink'; Path = $link; Target = ('..\' * 24) + 'Windows' }
            'drive-relative' = @{ ItemType = 'SymbolicLink'; Path = $link; Target = 'C:Windows' }
            'rooted-without-drive' = @{ ItemType = 'SymbolicLink'; Path = $link; Target = '\Windows' }
            'unc' = @{ ItemType = 'SymbolicLink'; Path = $link; Target = '\\localhost\C$\Windows' }
            'stream' = @{ ItemType = 'SymbolicLink'; Path = $link; Target = "$($target):stream" }
            'relative-junction' = @{ ItemType = 'Junction'; Path = $link; Target = 'inside-target' }
            'link-outside' = @{ ItemType = 'Junction'; Path = (Join-Path $env:SystemRoot 'rs-q12-refused-link'); Target = $target }
        }
        foreach ($name in $refusals.Keys) {
            $case = $refusals[$name]
            $message = $null
            try { [void](New-RSTestSandboxLink @case) } catch { $message = $_.Exception.Message }
            "$name -> $message" | Should Match '^[a-z-]+ -> Refusing test '
            (Test-Path -LiteralPath $case.Path) | Should Be $false
        }

        New-RSContainmentReparseFixture -Fixture $fixture -Link $link -Target $target -ItemType Junction
        (([IO.File]::GetAttributes($link) -band [IO.FileAttributes]::ReparsePoint) -ne 0) | Should Be $true
        $relative = Join-Path $fixture.Directory 'relative-link'
        New-RSContainmentReparseFixture -Fixture $fixture -Link $relative -Target 'inside-target' -ItemType SymbolicLink
        (Assert-RSTestSandboxLinkTarget -Path $relative -Target 'inside-target') | Should Be $target
    }

    It 'observes an ordinary ancestor replaced by a junction between successive calls' {
        $fixture = New-RSContainmentRegressionFixture
        $ancestor = Join-Path $fixture.Directory 'replaceable'
        $retained = Join-Path $fixture.Directory 'retained-ordinary'
        $target = Join-Path $fixture.Directory 'replacement-target'
        New-RSContainmentOrdinaryDirectory -Fixture $fixture -Path $ancestor
        New-RSContainmentOrdinaryDirectory -Fixture $fixture -Path $target
        $leaf = Join-Path $ancestor 'missing.txt'
        (Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $leaf) | Should Be $leaf
        [void](Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $ancestor -RequireDescendant)
        [void](Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $retained -RequireDescendant)
        [IO.Directory]::Move($ancestor, $retained)
        New-RSContainmentReparseFixture -Fixture $fixture -Link $ancestor -Target $target -ItemType Junction
        { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $leaf } | Should Throw 'reparse-point'
        [IO.Directory]::Exists($retained) | Should Be $true
    }

    It 'propagates a named injected IOException instead of treating it as a missing path' {
        $fixture = New-RSContainmentRegressionFixture
        Invoke-RSContainmentNamedAttributeFault -Fixture $fixture -ExceptionType '[IO.IOException]'
    }

    It 'propagates a named injected UnauthorizedAccessException instead of treating it as a missing path' {
        $fixture = New-RSContainmentRegressionFixture
        Invoke-RSContainmentNamedAttributeFault -Fixture $fixture -ExceptionType '[UnauthorizedAccessException]'
    }
}

# A mount-point reparse buffer naming a volume GUID needs no administrator when
# set directly, unlike SetVolumeMountPointW. The GUID below names no volume and
# no mount-manager entry is registered, so the fixture can never reach real data.
if (-not ('RSSandboxCleanupMountPointFixture' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
public static class RSSandboxCleanupMountPointFixture {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern SafeFileHandle CreateFileW(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool DeviceIoControl(SafeFileHandle device, uint code, byte[] input, int inputSize, IntPtr output, int outputSize, out int returned, IntPtr overlapped);
    const string SyntheticVolume = @"\??\Volume{00000000-5a4d-4e2b-9c11-7e57c1ea0f00}\";
    public static void SetMountPoint(string emptyDirectory, string substituteName) {
        // Q12 runtime guard: the only raw mount point a test creates names this reserved GUID,
        // which resolves to no volume, and lives inside a marked X:\RedSalamander.Perf.
        if (!string.Equals(substituteName, SyntheticVolume, StringComparison.OrdinalIgnoreCase) ||
            System.IO.Directory.Exists(@"\\?\Volume{00000000-5a4d-4e2b-9c11-7e57c1ea0f00}\"))
            throw new InvalidOperationException("Refusing a mount point that could name a real volume: " + substituteName);
        string full = System.IO.Path.GetFullPath(emptyDirectory);
        string sandbox = System.IO.Path.Combine(System.IO.Path.GetPathRoot(full), "RedSalamander.Perf") + "\\";
        string marker = System.IO.Path.Combine(sandbox, ".red-salamander-test-root");
        if (!full.StartsWith(sandbox, StringComparison.OrdinalIgnoreCase) || !System.IO.File.Exists(marker) ||
            System.IO.File.ReadAllText(marker).TrimEnd('\r', '\n') != "red-salamander.test-sandbox-root.v1")
            throw new InvalidOperationException("Refusing a mount point outside a marked test sandbox: " + full);
        byte[] substitute = Encoding.Unicode.GetBytes(substituteName);
        int dataLength = 8 + substitute.Length + 2 + 2;
        byte[] buffer = new byte[8 + dataLength];
        BitConverter.GetBytes(0xA0000003u).CopyTo(buffer, 0);           // IO_REPARSE_TAG_MOUNT_POINT
        BitConverter.GetBytes((ushort)dataLength).CopyTo(buffer, 4);
        BitConverter.GetBytes((ushort)substitute.Length).CopyTo(buffer, 10);
        BitConverter.GetBytes((ushort)(substitute.Length + 2)).CopyTo(buffer, 12); // empty print name
        substitute.CopyTo(buffer, 16);
        // GENERIC_WRITE, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT
        using (SafeFileHandle handle = CreateFileW(emptyDirectory, 0x40000000, 0, IntPtr.Zero, 3, 0x02200000, IntPtr.Zero)) {
            if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
            int returned;
            if (!DeviceIoControl(handle, 0x000900A4, buffer, buffer.Length, IntPtr.Zero, 0, out returned, IntPtr.Zero)) // FSCTL_SET_REPARSE_POINT
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
    }
}
'@
}

function New-RSSandboxCleanupFixture {
    $fixture = New-RSContainmentRegressionFixture
    # A dead owner PID; cleanup is scoped by CandidateRunIds to this fixture only.
    $runId = '20260706T110000Z-1234-' + [Guid]::NewGuid().ToString('N')
    $runRoot = Join-Path (Join-Path $fixture.Root 'runs') $runId
    $nested = Join-Path $runRoot 'artifacts\selftest\last_run\fileops\work'
    [void](Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $nested -RequireDescendant)
    [void][IO.Directory]::CreateDirectory($nested)
    [IO.File]::WriteAllText((Join-Path $nested 'evidence.txt'), 'stale run evidence')
    $readOnly = Join-Path $nested 'read-only.txt'
    [IO.File]::WriteAllText($readOnly, 'read-only')
    [IO.File]::SetAttributes($readOnly, [IO.FileAttributes]::ReadOnly -bor [IO.FileAttributes]::Hidden)
    # The victim lives outside the stale run, in the current run's scratch.
    $victim = Join-Path $fixture.Directory 'victim'
    New-RSContainmentOrdinaryDirectory -Fixture $fixture -Path (Join-Path $victim 'nested')
    [IO.File]::WriteAllText((Join-Path $victim 'keep.txt'), 'victim')
    [IO.File]::WriteAllText((Join-Path $victim 'nested\keep-nested.txt'), 'victim nested')
    $junction = Join-Path $nested 'victim-junction'
    New-RSTestSandboxLink -ItemType Junction -Path $junction -Target $victim | Out-Null
    return [pscustomobject]@{
        Fixture = $fixture
        RunId = $runId
        RunRoot = $runRoot
        Nested = $nested
        Victim = $victim
        Junction = $junction
    }
}

function Add-RSSandboxCleanupVolumeMountPoint {
    param($Cleanup)
    $mount = Join-Path $Cleanup.Nested 'f7-volume-mount'
    [void](Assert-RSTestSandboxContainedPath -TestRoot $Cleanup.Fixture.Root -Path $mount -RequireDescendant)
    [void][IO.Directory]::CreateDirectory($mount)
    [RSSandboxCleanupMountPointFixture]::SetMountPoint(
        $mount, '\??\Volume{00000000-5a4d-4e2b-9c11-7e57c1ea0f00}\')
    if (-not (Test-RSSandboxCleanupLinkPresent -Path $mount)) {
        throw "The fixture volume mount point was not created: $mount"
    }
    return $mount
}

function Test-RSSandboxCleanupLinkPresent {
    param([string]$Path)
    # Directory.Exists follows the link; a dangling mount point reads as absent.
    try {
        return (([IO.File]::GetAttributes($Path) -band [IO.FileAttributes]::ReparsePoint) -ne 0)
    } catch [IO.FileNotFoundException] {
        return $false
    } catch [IO.DirectoryNotFoundException] {
        return $false
    }
}

function Assert-RSSandboxCleanupVictimIntact {
    param($Cleanup)
    [IO.File]::ReadAllText((Join-Path $Cleanup.Victim 'keep.txt')) | Should Be 'victim'
    [IO.File]::ReadAllText((Join-Path $Cleanup.Victim 'nested\keep-nested.txt')) | Should Be 'victim nested'
}

function Remove-RSSandboxCleanupFixture {
    param($Cleanup, [string]$Mount = '')
    # The fixture mount point has no mount-manager registration, so removing the
    # directory itself is the complete cleanup.
    if (-not [string]::IsNullOrEmpty($Mount) -and (Test-RSSandboxCleanupLinkPresent -Path $Mount)) {
        [IO.Directory]::Delete($Mount, $false)
    }
    Remove-RSTestSandboxExactRunDirectory -TestRoot $Cleanup.Fixture.Root -RunId $Cleanup.RunId
}

Describe 'TestSandbox run cleanup never follows reparse points below a run root' {
    It 'unlinks a junction inside a stale run without touching the outside victim' {
        $cleanup = New-RSSandboxCleanupFixture
        try {
            $results = @(Remove-RSTestSandboxStaleRunDirectories -TestRoot $cleanup.Fixture.Root `
                -CandidateRunIds @($cleanup.RunId) -LiveProcessIds @() -LiveProcessStartTimesUtc @{})
            $results.Count | Should Be 1
            $results[0].Status | Should Be 'Removed'
            [IO.Directory]::Exists($cleanup.RunRoot) | Should Be $false
            Assert-RSSandboxCleanupVictimIntact -Cleanup $cleanup
        } finally {
            Remove-RSSandboxCleanupFixture -Cleanup $cleanup
        }
    }

    It 'retains a whole stale run that contains a volume mount point and deletes nothing' {
        $cleanup = New-RSSandboxCleanupFixture
        $mount = $null
        try {
            $mount = Add-RSSandboxCleanupVolumeMountPoint -Cleanup $cleanup
            $results = @(Remove-RSTestSandboxStaleRunDirectories -TestRoot $cleanup.Fixture.Root `
                -CandidateRunIds @($cleanup.RunId) -LiveProcessIds @() -LiveProcessStartTimesUtc @{} `
                -WarningAction SilentlyContinue)
            $results.Count | Should Be 1
            $results[0].Status | Should Be 'Retained'
            $results[0].Error | Should Match 'volume mount point'
            $results[0].Error.Contains($mount) | Should Be $true
            Test-RSSandboxCleanupLinkPresent -Path $mount | Should Be $true
            [IO.File]::Exists((Join-Path $cleanup.Nested 'evidence.txt')) | Should Be $true
            [IO.File]::Exists((Join-Path $cleanup.Nested 'read-only.txt')) | Should Be $true
            Test-RSSandboxCleanupLinkPresent -Path $cleanup.Junction | Should Be $true
            Assert-RSSandboxCleanupVictimIntact -Cleanup $cleanup
        } finally {
            Remove-RSSandboxCleanupFixture -Cleanup $cleanup -Mount $mount
        }
    }

    It 'removes an exact run through a junction as a link and refuses one with a volume mount point' {
        $cleanup = New-RSSandboxCleanupFixture
        $mount = $null
        try {
            $mount = Add-RSSandboxCleanupVolumeMountPoint -Cleanup $cleanup
            { Remove-RSTestSandboxExactRunDirectory -TestRoot $cleanup.Fixture.Root -RunId $cleanup.RunId } |
                Should Throw 'volume mount point'
            [IO.File]::Exists((Join-Path $cleanup.Nested 'evidence.txt')) | Should Be $true
            Test-RSSandboxCleanupLinkPresent -Path $cleanup.Junction | Should Be $true

            [IO.Directory]::Delete($mount, $false)
            Remove-RSTestSandboxExactRunDirectory -TestRoot $cleanup.Fixture.Root -RunId $cleanup.RunId
            [IO.Directory]::Exists($cleanup.RunRoot) | Should Be $false
            Assert-RSSandboxCleanupVictimIntact -Cleanup $cleanup
        } finally {
            Remove-RSSandboxCleanupFixture -Cleanup $cleanup -Mount $mount
        }
    }
}

# Test-owned VHDX coverage runs unelevated: every privileged storage operation is replaced by
# a recording fake, and the "mounted" volume is the unregistered mount-point fixture above.
$rsOwnedVolumeGuid = '\\?\Volume{00000000-5a4d-4e2b-9c11-7e57c1ea0f00}\'

function Invoke-RSOwnedVirtualDiskModule {
    param([scriptblock]$Body, [object[]]$Arguments = @())
    return (& $containmentModule $Body @Arguments)
}

function New-RSOwnedVirtualDiskFakeOperations {
    param(
        [bool]$Elevated,
        [System.Collections.Generic.List[string]]$Log,
        [hashtable]$State,
        [int[]]$VolumeDiskNumbers = @(7),
        [hashtable]$DiskFacts = @{},
        [switch]$FailUnmount
    )
    $guid = $rsOwnedVolumeGuid
    $failUnmountNow = [bool]$FailUnmount
    return @{
        IsElevated = { $Elevated }.GetNewClosure()
        CreateFile = { param($VhdxPath, $SizeMB) $Log.Add('CreateFile'); [IO.File]::WriteAllText($VhdxPath, 'fake vhdx') }.GetNewClosure()
        Attach = { param($VhdxPath) $Log.Add('Attach'); $State.Attached = $true; $State.VhdxPath = $VhdxPath; return 7 }.GetNewClosure()
        GetDiskFacts = {
            param($DiskNumber)
            $Log.Add('GetDiskFacts')
            # The blank 128 MB file-backed disk the attach created, unless a case overrides a fact.
            $facts = [ordered]@{
                DiskNumber = [int]$DiskNumber; Location = [string]$State.VhdxPath; BusType = 'File Backed Virtual'
                PartitionStyle = 'RAW'; Size = [uint64]128MB
            }
            foreach ($key in $DiskFacts.Keys) { $facts[$key] = $DiskFacts[$key] }
            return [pscustomobject]$facts
        }.GetNewClosure()
        Format = { param($DiskNumber) $Log.Add('Format'); return [pscustomobject]@{ PartitionNumber = 2; VolumeGuidPath = $guid } }.GetNewClosure()
        GetVolumeFacts = {
            param($VhdxPath, $VolumeGuidPath)
            $Log.Add('GetVolumeFacts')
            return [pscustomobject]@{
                ImageDiskNumber = 7
                DiskLocation = $VhdxPath
                DiskBusType = '15'
                VolumeDiskNumbers = $VolumeDiskNumbers
                VolumeAccessPaths = @($VolumeGuidPath)
            }
        }.GetNewClosure()
        AddMount = {
            param($DiskNumber, $PartitionNumber, $MountPath)
            $receipt = Get-Content -LiteralPath $State.ReceiptPath -Raw | ConvertFrom-Json
            $Log.Add("AddMount:$($receipt.stage):$($receipt.volume_guid_path)")
            [RSSandboxCleanupMountPointFixture]::SetMountPoint($MountPath, '\??\' + $guid.Substring(4))
        }.GetNewClosure()
        Unmount = {
            param($MountPath)
            $Log.Add('Unmount')
            if ($failUnmountNow) { throw 'injected unmount failure' }
            [IO.Directory]::Delete($MountPath, $false)
            [void][IO.Directory]::CreateDirectory($MountPath)
        }.GetNewClosure()
        IsAttached = { param($VhdxPath) return [bool]$State.Attached }.GetNewClosure()
        Detach = { param($VhdxPath) $Log.Add('Detach'); $State.Attached = $false }.GetNewClosure()
    }
}

function Invoke-RSWithOwnedVirtualDiskOperations {
    param([hashtable]$Operations, [scriptblock]$Body)
    $saved = Invoke-RSOwnedVirtualDiskModule { $script:RSTestOwnedVirtualDiskOperations }
    try {
        Invoke-RSOwnedVirtualDiskModule { param($Value) $script:RSTestOwnedVirtualDiskOperations = $Value } @($Operations)
        & $Body
    } finally {
        Invoke-RSOwnedVirtualDiskModule { param($Value) $script:RSTestOwnedVirtualDiskOperations = $Value } @($saved)
    }
}

function New-RSOwnedVirtualDiskRecordedRun {
    param($Cleanup)
    # A stale run as a killed elevated runner leaves it: receipt, disk file and live mount point.
    $layout = Invoke-RSOwnedVirtualDiskModule {
        param($RunRoot) Get-RSTestOwnedVirtualDiskLayout -RunRoot $RunRoot -Name 'f7-admission-baseline'
    } @($Cleanup.RunRoot)
    [void][IO.Directory]::CreateDirectory($layout.MountPath)
    [IO.File]::WriteAllText($layout.VhdxPath, 'fake vhdx')
    Invoke-RSOwnedVirtualDiskModule {
        param($Layout, $Guid) Write-RSTestOwnedVirtualDiskReceipt -Layout $Layout -Stage mounted -VolumeGuidPath $Guid
    } @($layout, $rsOwnedVolumeGuid)
    [RSSandboxCleanupMountPointFixture]::SetMountPoint($layout.MountPath, '\??\' + $rsOwnedVolumeGuid.Substring(4))
    return $layout
}

Describe 'Test-owned virtual disk lifecycle and receipt-aware cleanup' {
    It 'writes, reads and validates only run-owned receipts' {
        $cleanup = New-RSSandboxCleanupFixture
        try {
            $layout = Invoke-RSOwnedVirtualDiskModule {
                param($RunRoot) Get-RSTestOwnedVirtualDiskLayout -RunRoot $RunRoot -Name 'f7-admission-baseline'
            } @($cleanup.RunRoot)
            $layout.ReceiptPath | Should Be (Join-Path $cleanup.RunRoot 'artifacts\test-owned-vhdx\f7-admission-baseline.receipt.json')
            $layout.MountPath | Should Be (Join-Path $cleanup.RunRoot 'scratch\test-owned-vhdx\f7-admission-baseline\mount')
            Invoke-RSOwnedVirtualDiskModule { param($Layout) Write-RSTestOwnedVirtualDiskReceipt -Layout $Layout -Stage creating } @($layout)
            Invoke-RSOwnedVirtualDiskModule {
                param($Layout, $Guid) Write-RSTestOwnedVirtualDiskReceipt -Layout $Layout -Stage mounting -VolumeGuidPath $Guid
            } @($layout, $rsOwnedVolumeGuid)
            @(Get-RSTestOwnedVirtualDiskReceiptPaths -RunRoot $cleanup.RunRoot) | Should Be @($layout.ReceiptPath)
            $read = Invoke-RSOwnedVirtualDiskModule {
                param($Path, $RunRoot) Read-RSTestOwnedVirtualDiskReceipt -ReceiptPath $Path -RunRoot $RunRoot
            } @($layout.ReceiptPath, $cleanup.RunRoot)
            $read.Stage | Should Be 'mounting'
            $read.VolumeGuidPath | Should Be $rsOwnedVolumeGuid
            $read.VhdxPath | Should Be $layout.VhdxPath
            $read.MountPath | Should Be $layout.MountPath

            $original = [IO.File]::ReadAllText($layout.ReceiptPath)
            foreach ($tamper in @(
                    @{ Field = 'vhdx_path'; Value = 'D:\disk.vhdx'; Message = 'is not the run-owned' },
                    @{ Field = 'mount_path'; Value = 'D:\'; Message = 'is not the run-owned' },
                    @{ Field = 'volume_guid_path'; Value = 'D:\'; Message = 'malformed volume GUID path' },
                    @{ Field = 'schema'; Value = 'other'; Message = 'unexpected schema' })) {
                $value = $original | ConvertFrom-Json
                $value.($tamper.Field) = $tamper.Value
                [IO.File]::WriteAllText($layout.ReceiptPath, ($value | ConvertTo-Json))
                {
                    Invoke-RSOwnedVirtualDiskModule {
                        param($Path, $RunRoot) Read-RSTestOwnedVirtualDiskReceipt -ReceiptPath $Path -RunRoot $RunRoot
                    } @($layout.ReceiptPath, $cleanup.RunRoot)
                } | Should Throw $tamper.Message
            }
            [IO.File]::Delete($layout.ReceiptPath)
        } finally {
            Remove-RSSandboxCleanupFixture -Cleanup $cleanup
        }
    }

    It 'proves volume ownership and refuses a foreign volume GUID' {
        $vhdx = 'C:\RedSalamander.Perf\runs\r\scratch\test-owned-vhdx\f7\disk.vhdx'
        $facts = [pscustomobject]@{
            ImageDiskNumber = 7; DiskLocation = $vhdx; DiskBusType = 'File Backed Virtual'
            VolumeDiskNumbers = @(7); VolumeAccessPaths = @($rsOwnedVolumeGuid)
        }
        $assert = { param($Facts, $Guid, $Vhdx) Assert-RSTestOwnedVirtualDiskVolume -VhdxPath $Vhdx -VolumeGuidPath $Guid -Facts $Facts }
        Invoke-RSOwnedVirtualDiskModule $assert @($facts, $rsOwnedVolumeGuid, $vhdx)
        foreach ($case in @(
                @{ Change = { param($f) $f.VolumeDiskNumbers = @(0) }; Message = 'Refusing foreign volume' },
                @{ Change = { param($f) $f.VolumeDiskNumbers = @(0, 7) }; Message = 'Refusing foreign volume' },
                @{ Change = { param($f) $f.VolumeAccessPaths = @($rsOwnedVolumeGuid, 'D:\') }; Message = 'already reachable' },
                @{ Change = { param($f) $f.DiskLocation = 'D:\other.vhdx' }; Message = 'not the test-owned VHDX' },
                @{ Change = { param($f) $f.DiskBusType = '17' }; Message = 'not a file-backed virtual disk' })) {
            $changed = $facts.PSObject.Copy()
            & $case.Change $changed
            { Invoke-RSOwnedVirtualDiskModule $assert @($changed, $rsOwnedVolumeGuid, $vhdx) } | Should Throw $case.Message
        }
        { Invoke-RSOwnedVirtualDiskModule $assert @($facts, 'D:\', $vhdx) } | Should Throw 'malformed volume GUID'
    }

    It 'skips provisioning without elevation and performs no storage operation' {
        $cleanup = New-RSSandboxCleanupFixture
        $log = [System.Collections.Generic.List[string]]::new()
        try {
            $operations = New-RSOwnedVirtualDiskFakeOperations -Elevated $false -Log $log -State @{}
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $operations -Body {
                $result = New-RSTestOwnedVirtualDiskMount -TestRoot $cleanup.Fixture.Root -RunId $cleanup.RunId -Name 'f7-admission-baseline'
                $result.Status | Should Be 'Skipped'
                $result.Reason | Should Match '^not-elevated'
            }
            $log.Count | Should Be 0
            [IO.Directory]::Exists((Join-Path $cleanup.RunRoot 'artifacts\test-owned-vhdx')) | Should Be $false
            [IO.Directory]::Exists((Join-Path $cleanup.RunRoot 'scratch\test-owned-vhdx')) | Should Be $false
        } finally {
            Remove-RSSandboxCleanupFixture -Cleanup $cleanup
        }
    }

    It 'records the receipt before mounting and tears down in reverse order' {
        $cleanup = New-RSSandboxCleanupFixture
        $log = [System.Collections.Generic.List[string]]::new()
        $state = @{ Attached = $false; ReceiptPath = (Join-Path $cleanup.RunRoot 'artifacts\test-owned-vhdx\f7-admission-baseline.receipt.json') }
        try {
            $operations = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $operations -Body {
                $result = New-RSTestOwnedVirtualDiskMount -TestRoot $cleanup.Fixture.Root -RunId $cleanup.RunId -Name 'f7-admission-baseline'
                $result.Status | Should Be 'Mounted'
                $result.VolumeGuidPath | Should Be $rsOwnedVolumeGuid
                $result.ReceiptPath | Should Be $state.ReceiptPath
                ($log -join ',') | Should Be "CreateFile,Attach,GetDiskFacts,Format,GetVolumeFacts,AddMount:mounting:$($rsOwnedVolumeGuid)"
                (Get-Content -LiteralPath $state.ReceiptPath -Raw | ConvertFrom-Json).stage | Should Be 'mounted'
                Test-RSSandboxCleanupLinkPresent -Path $result.MountPath | Should Be $true

                $log.Clear()
                Remove-RSTestOwnedVirtualDiskMount -ReceiptPath $result.ReceiptPath -RunRoot $result.RunRoot
                ($log -join ',') | Should Be 'Unmount,Detach'
                Test-RSSandboxCleanupLinkPresent -Path $result.MountPath | Should Be $false
                [IO.File]::Exists($result.VhdxPath) | Should Be $false
                [IO.File]::Exists($result.ReceiptPath) | Should Be $false
            }
        } finally {
            Remove-RSSandboxCleanupFixture -Cleanup $cleanup
        }
    }

    It 'refuses a foreign volume before any mount and tears the disk down' {
        $cleanup = New-RSSandboxCleanupFixture
        $log = [System.Collections.Generic.List[string]]::new()
        $state = @{ Attached = $false; ReceiptPath = (Join-Path $cleanup.RunRoot 'artifacts\test-owned-vhdx\f7-admission-baseline.receipt.json') }
        try {
            $operations = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state -VolumeDiskNumbers @(0)
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $operations -Body {
                { New-RSTestOwnedVirtualDiskMount -TestRoot $cleanup.Fixture.Root -RunId $cleanup.RunId -Name 'f7-admission-baseline' } |
                    Should Throw 'Refusing foreign volume'
            }
            ($log -join ',') | Should Be 'CreateFile,Attach,GetDiskFacts,Format,GetVolumeFacts,Detach'
            [IO.File]::Exists((Join-Path $cleanup.RunRoot 'scratch\test-owned-vhdx\f7-admission-baseline\disk.vhdx')) | Should Be $false
            [IO.File]::Exists($state.ReceiptPath) | Should Be $false
        } finally {
            Remove-RSSandboxCleanupFixture -Cleanup $cleanup
        }
    }

    It 'stops teardown at the first failure and keeps the disk and receipt' {
        $cleanup = New-RSSandboxCleanupFixture
        $log = [System.Collections.Generic.List[string]]::new()
        $state = @{ Attached = $true }
        $layout = $null
        try {
            $layout = New-RSOwnedVirtualDiskRecordedRun -Cleanup $cleanup
            $failing = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state -FailUnmount
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $failing -Body {
                { Remove-RSTestOwnedVirtualDiskMount -ReceiptPath $layout.ReceiptPath -RunRoot $cleanup.RunRoot } |
                    Should Throw 'injected unmount failure'
            }
            ($log -join ',') | Should Be 'Unmount'
            $state.Attached | Should Be $true
            [IO.File]::Exists($layout.VhdxPath) | Should Be $true
            [IO.File]::Exists($layout.ReceiptPath) | Should Be $true
            Test-RSSandboxCleanupLinkPresent -Path $layout.MountPath | Should Be $true
        } finally {
            $working = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $working -Body {
                Remove-RSSandboxCleanupFixture -Cleanup $cleanup
            }
        }
    }

    It 'retains a stale or exact run with a receipt when not elevated and reports the recovery commands' {
        $cleanup = New-RSSandboxCleanupFixture
        $log = [System.Collections.Generic.List[string]]::new()
        $state = @{ Attached = $true }
        try {
            $layout = New-RSOwnedVirtualDiskRecordedRun -Cleanup $cleanup
            $operations = New-RSOwnedVirtualDiskFakeOperations -Elevated $false -Log $log -State $state
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $operations -Body {
                $results = @(Remove-RSTestSandboxStaleRunDirectories -TestRoot $cleanup.Fixture.Root `
                    -CandidateRunIds @($cleanup.RunId) -LiveProcessIds @() -LiveProcessStartTimesUtc @{} `
                    -WarningAction SilentlyContinue)
                $results.Count | Should Be 1
                $results[0].Status | Should Be 'Retained'
                $results[0].Error.Contains($layout.ReceiptPath) | Should Be $true
                $results[0].Error | Should Match 'Remove-RSTestOwnedVirtualDiskMount -ReceiptPath'
                $results[0].Error | Should Match 'mountvol "[^"]+\\mount\\" /D; Dismount-DiskImage -ImagePath'
                { Remove-RSTestSandboxExactRunDirectory -TestRoot $cleanup.Fixture.Root -RunId $cleanup.RunId } |
                    Should Throw $layout.ReceiptPath
            }
            $log.Count | Should Be 0
            Test-RSSandboxCleanupLinkPresent -Path $layout.MountPath | Should Be $true
            [IO.File]::Exists($layout.VhdxPath) | Should Be $true
            [IO.File]::Exists($layout.ReceiptPath) | Should Be $true
            [IO.File]::Exists((Join-Path $cleanup.Nested 'evidence.txt')) | Should Be $true
            Test-RSSandboxCleanupLinkPresent -Path $cleanup.Junction | Should Be $true
        } finally {
            $working = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $working -Body {
                Remove-RSSandboxCleanupFixture -Cleanup $cleanup
            }
        }
    }

    It 'tears recorded disks down before sweeping a stale run when elevated' {
        $cleanup = New-RSSandboxCleanupFixture
        $log = [System.Collections.Generic.List[string]]::new()
        $state = @{ Attached = $true }
        try {
            $null = New-RSOwnedVirtualDiskRecordedRun -Cleanup $cleanup
            $operations = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $operations -Body {
                $results = @(Remove-RSTestSandboxStaleRunDirectories -TestRoot $cleanup.Fixture.Root `
                    -CandidateRunIds @($cleanup.RunId) -LiveProcessIds @() -LiveProcessStartTimesUtc @{})
                $results.Count | Should Be 1
                $results[0].Status | Should Be 'Removed'
            }
            ($log -join ',') | Should Be 'Unmount,Detach'
            [IO.Directory]::Exists($cleanup.RunRoot) | Should Be $false
            Assert-RSSandboxCleanupVictimIntact -Cleanup $cleanup
        } finally {
            Remove-RSSandboxCleanupFixture -Cleanup $cleanup
        }
    }

    It 'formats only the blank file-backed disk of the run VHDX and tears down any other without touching it' {
        $vhdx = 'C:\RedSalamander.Perf\runs\r\scratch\test-owned-vhdx\f7\disk.vhdx'
        $blank = [pscustomobject]@{
            DiskNumber = 7; Location = $vhdx; BusType = 'File Backed Virtual'; PartitionStyle = 'RAW'; Size = [uint64]128MB
        }
        $assert = { param($Facts, $Vhdx, $Number) Assert-RSTestOwnedVirtualDiskBlankDisk -VhdxPath $Vhdx -DiskNumber $Number -Facts $Facts }
        Invoke-RSOwnedVirtualDiskModule $assert @($blank, $vhdx, 7)
        $numeric = $blank.PSObject.Copy()
        $numeric.BusType = '15'
        $numeric.PartitionStyle = '0'
        Invoke-RSOwnedVirtualDiskModule $assert @($numeric, $vhdx, 7)
        $refusals = @(
            @{ Facts = @{ DiskNumber = 0 }; Message = 'describe disk 0' },
            @{ Facts = @{ Location = 'PCIROOT(0)#PCI(1D00)#NVME(P00T00L00)' }; Message = 'not the test-owned VHDX' },
            @{ Facts = @{ Location = 'D:\other.vhdx' }; Message = 'not the test-owned VHDX' },
            @{ Facts = @{ BusType = 'NVMe' }; Message = 'not a file-backed virtual disk' },
            @{ Facts = @{ PartitionStyle = 'GPT' }; Message = 'not a blank RAW disk' },
            @{ Facts = @{ Size = [uint64]4000787030016 }; Message = 'not the test-owned VHDX size' })
        foreach ($case in $refusals) {
            $changed = $blank.PSObject.Copy()
            foreach ($key in $case.Facts.Keys) { $changed.$key = $case.Facts[$key] }
            $failure = $null
            try { Invoke-RSOwnedVirtualDiskModule $assert @($changed, $vhdx, 7) } catch { $failure = $_.Exception.Message }
            $failure | Should Match ([regex]::Escape($case.Message))
            $failure | Should Match '^Refusing to format disk 7'
        }

        # Through provisioning: a refused disk is never initialized, partitioned or formatted.
        foreach ($case in $refusals) {
            $cleanup = New-RSSandboxCleanupFixture
            $log = [System.Collections.Generic.List[string]]::new()
            $state = @{ Attached = $false; ReceiptPath = (Join-Path $cleanup.RunRoot 'artifacts\test-owned-vhdx\f7-admission-baseline.receipt.json') }
            try {
                $operations = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state -DiskFacts $case.Facts
                $outcome = @{ Failure = $null }
                Invoke-RSWithOwnedVirtualDiskOperations -Operations $operations -Body {
                    try {
                        $null = New-RSTestOwnedVirtualDiskMount -TestRoot $cleanup.Fixture.Root -RunId $cleanup.RunId -Name 'f7-admission-baseline'
                    } catch {
                        $outcome.Failure = $_.Exception.Message
                    }
                }
                $outcome.Failure | Should Match 'provisioning failed and was torn down: Refusing to format disk 7'
                ($log -join ',') | Should Be 'CreateFile,Attach,GetDiskFacts,Detach'
                $state.Attached | Should Be $false
                [IO.File]::Exists((Join-Path $cleanup.RunRoot 'scratch\test-owned-vhdx\f7-admission-baseline\disk.vhdx')) | Should Be $false
                [IO.File]::Exists($state.ReceiptPath) | Should Be $false
            } finally {
                Remove-RSSandboxCleanupFixture -Cleanup $cleanup
            }
        }
    }

    It 'keeps a detached VHDX as run evidence and reclaims it with the run without elevation' {
        $cleanup = New-RSSandboxCleanupFixture
        $log = [System.Collections.Generic.List[string]]::new()
        $state = @{ Attached = $true }
        try {
            $layout = New-RSOwnedVirtualDiskRecordedRun -Cleanup $cleanup
            $elevated = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $elevated -Body {
                Remove-RSTestOwnedVirtualDiskMount -ReceiptPath $layout.ReceiptPath -RunRoot $cleanup.RunRoot -RetainImage
            }
            ($log -join ',') | Should Be 'Unmount,Detach'
            $state.Attached | Should Be $false
            Test-RSSandboxCleanupLinkPresent -Path $layout.MountPath | Should Be $false
            [IO.File]::ReadAllText($layout.VhdxPath) | Should Be 'fake vhdx'
            $receipt = Get-Content -LiteralPath $layout.ReceiptPath -Raw | ConvertFrom-Json
            $receipt.stage | Should Be 'detached'
            $receipt.volume_guid_path | Should Be $rsOwnedVolumeGuid
            $receipt.vhdx_path | Should Be $layout.VhdxPath

            # The detached disk goes with its run: no privilege and no storage operation.
            $log.Clear()
            $unelevated = New-RSOwnedVirtualDiskFakeOperations -Elevated $false -Log $log -State $state
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $unelevated -Body {
                $results = @(Remove-RSTestSandboxStaleRunDirectories -TestRoot $cleanup.Fixture.Root `
                    -CandidateRunIds @($cleanup.RunId) -LiveProcessIds @() -LiveProcessStartTimesUtc @{})
                $results.Count | Should Be 1
                $results[0].Status | Should Be 'Removed'
            }
            $log.Count | Should Be 0
            [IO.Directory]::Exists($cleanup.RunRoot) | Should Be $false
            Assert-RSSandboxCleanupVictimIntact -Cleanup $cleanup
        } finally {
            $working = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $working -Body {
                Remove-RSSandboxCleanupFixture -Cleanup $cleanup
            }
        }
    }

    It 'retains a detached image without elevation once it is mounted or attached again' {
        $cleanup = New-RSSandboxCleanupFixture
        $log = [System.Collections.Generic.List[string]]::new()
        $state = @{ Attached = $true }
        try {
            # A detached receipt whose disk something attached and mounted at the run path again.
            $layout = New-RSOwnedVirtualDiskRecordedRun -Cleanup $cleanup
            Invoke-RSOwnedVirtualDiskModule {
                param($Layout, $Guid) Write-RSTestOwnedVirtualDiskReceipt -Layout $Layout -Stage detached -VolumeGuidPath $Guid
            } @($layout, $rsOwnedVolumeGuid)
            $operations = New-RSOwnedVirtualDiskFakeOperations -Elevated $false -Log $log -State $state
            foreach ($mounted in @($true, $false)) {
                if (-not $mounted) {
                    [IO.Directory]::Delete($layout.MountPath, $false)
                    [void][IO.Directory]::CreateDirectory($layout.MountPath)
                }
                Invoke-RSWithOwnedVirtualDiskOperations -Operations $operations -Body {
                    $results = @(Remove-RSTestSandboxStaleRunDirectories -TestRoot $cleanup.Fixture.Root `
                        -CandidateRunIds @($cleanup.RunId) -LiveProcessIds @() -LiveProcessStartTimesUtc @{} `
                        -WarningAction SilentlyContinue)
                    $results.Count | Should Be 1
                    $results[0].Status | Should Be 'Retained'
                    $results[0].Error.Contains($layout.ReceiptPath) | Should Be $true
                    $results[0].Error | Should Match 'requires elevation'
                }
                $log.Count | Should Be 0
                [IO.File]::Exists($layout.VhdxPath) | Should Be $true
                [IO.File]::Exists($layout.ReceiptPath) | Should Be $true
                [IO.File]::Exists((Join-Path $cleanup.Nested 'evidence.txt')) | Should Be $true
            }
        } finally {
            $working = New-RSOwnedVirtualDiskFakeOperations -Elevated $true -Log $log -State $state
            Invoke-RSWithOwnedVirtualDiskOperations -Operations $working -Body {
                Remove-RSSandboxCleanupFixture -Cleanup $cleanup
            }
        }
    }
}

Describe 'Sandbox audit root authority' {
    It 'rejects a missing unauthorized root rather than reporting a clean disk' {
        $fixture = New-RSContainmentRegressionFixture
        $missing = Join-Path $fixture.Directory 'missing-audit-root'
        [IO.Directory]::Exists($missing) | Should Be $false
        $auditCommand = Get-Command Get-RSTestSandboxDiskAudit -ErrorAction Stop
        $rejection = ''
        try { & $auditCommand -TestRoot $missing | Out-Null } catch { $rejection = $_.Exception.Message }
        $rejection | Should Match 'Unauthorized TestSandbox root'
        [IO.Directory]::Exists($missing) | Should Be $false
    }

    It 'requires the existing ownership marker without modifying that marker' {
        $fixture = New-RSContainmentRegressionFixture
        $marker = Join-Path $fixture.Root '.red-salamander-test-root'
        $markerHash = (Get-FileHash -LiteralPath $marker -Algorithm SHA256).Hash
        $auditCommand = Get-Command Get-RSTestSandboxDiskAudit -ErrorAction Stop
        $auditModule = $auditCommand.Module
        $original = & $auditModule { (Get-Item Function:Test-RSTestSandboxMarker).ScriptBlock }
        try {
            # In-process failed lookup only; the shared on-disk marker is never changed.
            & $auditModule { Set-Item Function:script:Test-RSTestSandboxMarker -Value { return $false } }
            $rejection = ''
            try { & $auditCommand -TestRoot $fixture.Root | Out-Null } catch { $rejection = $_.Exception.Message }
            $rejection | Should Match 'not authorized|unowned'
        } finally {
            & $auditModule { param($Body) Set-Item Function:script:Test-RSTestSandboxMarker -Value $Body } $original
        }
        (Get-FileHash -LiteralPath $marker -Algorithm SHA256).Hash | Should Be $markerHash
        { Assert-RSTestSandboxContainedPath -TestRoot $fixture.Root -Path $fixture.Directory } | Should Not Throw
    }
}
