Set-StrictMode -Version Latest

Import-Module (Join-Path $PSScriptRoot 'TestRunPlan.Common.psm1') -Force -Scope Local

$script:RSTestSandboxMarkerFileName = '.red-salamander-test-root'
$script:RSTestSandboxMarkerContents = 'red-salamander.test-sandbox-root.v1'
$script:RSTestSandboxDirectoryName = 'RedSalamander.Perf'

function Get-RSTestSandboxRepositoryRoot {
    param([string]$RepoRoot = '')

    if (-not [string]::IsNullOrWhiteSpace($RepoRoot)) {
        return (ConvertTo-RSFullPath -Path $RepoRoot)
    }

    return (ConvertTo-RSFullPath -Path (Join-Path $PSScriptRoot '..\..\..'))
}

function Test-RSPathSameOrDescendant {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Parent
    )

    $candidate = (ConvertTo-RSFullPath -Path $Path).TrimEnd('\')
    $container = (ConvertTo-RSFullPath -Path $Parent).TrimEnd('\')
    return [string]::Equals($candidate, $container, [System.StringComparison]::OrdinalIgnoreCase) -or
        $candidate.StartsWith("$container\", [System.StringComparison]::OrdinalIgnoreCase)
}

function Assert-RSTestSandboxPathIsReparseFree {
    param([Parameter(Mandatory = $true)][string]$Path)

    $candidate = ConvertTo-RSFullPath -Path $Path
    $current = $candidate
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        # Query every component afresh. Never cache a prior reparse-free result:
        # FILE_SHARE_* does not exclude FILE_WRITE_ATTRIBUTES reparse mutation.
        # Missing create-only descendants are permitted; every existing ancestor
        # must still be checked. Other IO/access errors fail closed.
        $attributes = $null
        try {
            $attributes = [IO.File]::GetAttributes($current)
        } catch [IO.FileNotFoundException] {
            # The canonical policy permits an absent output leaf/component.
        } catch [IO.DirectoryNotFoundException] {
            # Continue upward so a missing descendant cannot hide a reparse.
        }
        if ($null -ne $attributes -and
            ($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Refusing TestSandbox path with reparse-point component '$current': $candidate"
        }

        $parent = [IO.Path]::GetDirectoryName($current)
        if ([string]::IsNullOrWhiteSpace($parent) -or
            [string]::Equals($parent, $current, [System.StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        $current = $parent
    }
}

function Test-RSTestSandboxMarker {
    param([Parameter(Mandatory = $true)][string]$TestRoot)

    $markerPath = Join-Path (ConvertTo-RSFullPath -Path $TestRoot) $script:RSTestSandboxMarkerFileName
    if (-not (Test-Path -LiteralPath $markerPath -PathType Leaf)) {
        return $false
    }
    $marker = Get-Item -LiteralPath $markerPath -Force -ErrorAction Stop
    if (($marker.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        return $false
    }
    return [string]::Equals(
        (Get-Content -LiteralPath $markerPath -Raw -ErrorAction Stop).TrimEnd("`r", "`n"),
        $script:RSTestSandboxMarkerContents,
        [System.StringComparison]::Ordinal)
}

function Get-RSDedicatedExternalTestSandboxRoot {
    param([Parameter(Mandatory = $true)][string]$Path)

    $candidate = ConvertTo-RSFullPath -Path $Path
    $root = [System.IO.Path]::GetPathRoot($candidate)
    if ([string]::IsNullOrWhiteSpace($root) -or $root.StartsWith('\\')) {
        return $null
    }

    try {
        $drive = [System.IO.DriveInfo]::new($root)
        if (-not $drive.IsReady -or
            $drive.DriveType -ne [System.IO.DriveType]::Fixed) {
            return $null
        }
    } catch {
        return $null
    }

    $externalRoot = ConvertTo-RSFullPath -Path (Join-Path $root $script:RSTestSandboxDirectoryName)
    if (-not [string]::Equals($candidate.TrimEnd('\'), $externalRoot.TrimEnd('\'), [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }
    return $externalRoot
}

function Assert-RSTestSandboxRootAuthorized {
    param(
        [Parameter(Mandatory = $true)][string]$TestRoot,
        [string]$RepoRoot = '',
        [switch]$AllowExternalInitialization,
        [switch]$RequireMarker
    )

    [void](Get-RSTestSandboxRepositoryRoot -RepoRoot $RepoRoot)
    $candidate = ConvertTo-RSFullPath -Path $TestRoot
    $candidateRoot = [System.IO.Path]::GetPathRoot($candidate)
    if ($candidate.Substring($candidateRoot.Length).Contains(':')) {
        throw "Unauthorized TestSandbox root with alternate-data-stream syntax: $candidate"
    }
    Assert-RSTestSandboxPathIsReparseFree -Path $candidate

    $externalRoot = Get-RSDedicatedExternalTestSandboxRoot -Path $candidate
    if ($null -eq $externalRoot) {
        throw "Unauthorized TestSandbox root '$candidate'. Use exact X:\$script:RSTestSandboxDirectoryName on a fixed local drive."
    }
    if (-not (Test-RSTestSandboxMarker -TestRoot $externalRoot) -and -not $AllowExternalInitialization) {
        throw "TestSandbox root '$externalRoot' is not authorized. Initialize it explicitly with -AllowExternalTestRoot."
    }
    if ($RequireMarker -and -not (Test-RSTestSandboxMarker -TestRoot $externalRoot)) {
        throw "Refusing to mutate unowned TestSandbox root without '$script:RSTestSandboxMarkerFileName': $externalRoot"
    }
    return $externalRoot
}

function Initialize-RSTestSandboxRoot {
    param(
        [Parameter(Mandatory = $true)][string]$TestRoot,
        [string]$RepoRoot = '',
        [switch]$AllowExternalInitialization
    )

    $root = Assert-RSTestSandboxRootAuthorized `
        -TestRoot $TestRoot `
        -RepoRoot $RepoRoot `
        -AllowExternalInitialization:$AllowExternalInitialization
    $markerPath = Join-Path $root $script:RSTestSandboxMarkerFileName
    if (-not (Test-RSTestSandboxMarker -TestRoot $root) -and (Test-Path -LiteralPath $root)) {
        $existingChildren = @(Get-ChildItem -LiteralPath $root -Force -ErrorAction Stop)
        if ($existingChildren.Count -ne 0) {
            throw "Refusing to adopt non-empty TestSandbox root without a valid ownership marker: $root"
        }
    }

    New-Item -ItemType Directory -Path $root -Force -ErrorAction Stop | Out-Null
    Assert-RSTestSandboxPathIsReparseFree -Path $root
    if (Test-Path -LiteralPath $markerPath) {
        if (-not (Test-RSTestSandboxMarker -TestRoot $root)) {
            throw "TestSandbox ownership marker is invalid: $markerPath"
        }
    } else {
        Set-Content -LiteralPath $markerPath -Value $script:RSTestSandboxMarkerContents -NoNewline -Encoding ascii -ErrorAction Stop
    }
    return $root
}

function Assert-RSTestSandboxContainedPath {
    param(
        [Parameter(Mandatory = $true)][string]$TestRoot,
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$RequireDescendant
    )

    $root = Assert-RSTestSandboxRootAuthorized -TestRoot $TestRoot -RequireMarker
    $candidate = ConvertTo-RSFullPath -Path $Path
    $candidateRoot = [System.IO.Path]::GetPathRoot($candidate)
    if ($candidate.Substring($candidateRoot.Length).Contains(':')) {
        throw "Unauthorized TestSandbox path with alternate-data-stream syntax: $candidate"
    }
    if (-not (Test-RSPathSameOrDescendant -Path $candidate -Parent $root) -or
        ($RequireDescendant -and [string]::Equals(
                $candidate.TrimEnd('\'),
                $root.TrimEnd('\'),
                [System.StringComparison]::OrdinalIgnoreCase))) {
        throw "TestSandbox path must remain below '$root': $candidate"
    }
    Assert-RSTestSandboxPathIsReparseFree -Path $candidate
    return $candidate
}

function Get-RSTestSandboxRoot {
    param(
        [string]$RepoRoot = '',

        [string]$TestRootOverride = $env:REDSALAMANDER_TEST_ROOT,

        [string]$SelfTestRootOverride = $env:REDSALAMANDER_SELFTEST_ROOT,

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA,

        [switch]$AllowExternalInitialization
    )

    $repository = Get-RSTestSandboxRepositoryRoot -RepoRoot $RepoRoot

    if (-not [string]::IsNullOrWhiteSpace($TestRootOverride)) {
        $testRoot = Assert-RSTestSandboxRootAuthorized `
            -TestRoot $TestRootOverride `
            -RepoRoot $repository `
            -AllowExternalInitialization:$AllowExternalInitialization
        if (-not [string]::IsNullOrWhiteSpace($SelfTestRootOverride)) {
            $legacyRoot = ConvertTo-RSFullPath -Path $SelfTestRootOverride
            $testRootNormalized = $testRoot.TrimEnd('\')
            $legacyRootNormalized = $legacyRoot.TrimEnd('\')
            $runsRootNormalized = (Join-Path $testRootNormalized 'runs').TrimEnd('\')
            $isSameRoot = [string]::Equals($testRootNormalized, $legacyRootNormalized, [System.StringComparison]::OrdinalIgnoreCase)
            $isRunnerBridge = $legacyRootNormalized.StartsWith("$runsRootNormalized\", [System.StringComparison]::OrdinalIgnoreCase)
            if (-not $isSameRoot -and -not $isRunnerBridge) {
                throw "REDSALAMANDER_TEST_ROOT ('$testRoot') and REDSALAMANDER_SELFTEST_ROOT ('$legacyRoot') conflict. Set only REDSALAMANDER_TEST_ROOT, clear REDSALAMANDER_SELFTEST_ROOT, or point both at the same sandbox base."
            }
        }

        return $testRoot
    }

    if (-not [string]::IsNullOrWhiteSpace($SelfTestRootOverride)) {
        return (Assert-RSTestSandboxRootAuthorized `
                -TestRoot $SelfTestRootOverride `
                -RepoRoot $repository)
    }

    $repositoryDriveRoot = [System.IO.Path]::GetPathRoot($repository)
    if ([string]::IsNullOrWhiteSpace($repositoryDriveRoot) -or $repositoryDriveRoot.StartsWith('\\')) {
        throw "Cannot derive X:\$script:RSTestSandboxDirectoryName from repository root '$repository'. Set REDSALAMANDER_TEST_ROOT explicitly."
    }
    return (Assert-RSTestSandboxRootAuthorized `
            -TestRoot (Join-Path $repositoryDriveRoot $script:RSTestSandboxDirectoryName) `
            -RepoRoot $repository `
            -AllowExternalInitialization:$AllowExternalInitialization)
}

function New-RSTestRunContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [string]$RunId = '',

        [string]$TestRootOverride = $env:REDSALAMANDER_TEST_ROOT,

        [string]$SelfTestRootOverride = $env:REDSALAMANDER_SELFTEST_ROOT,

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA,

        [switch]$AllowExternalInitialization
    )

    $testRoot = Get-RSTestSandboxRoot `
        -RepoRoot $RepoRoot `
        -TestRootOverride $TestRootOverride `
        -SelfTestRootOverride $SelfTestRootOverride `
        -LocalAppDataRoot $LocalAppDataRoot `
        -AllowExternalInitialization:$AllowExternalInitialization
    if ([string]::IsNullOrWhiteSpace($RunId)) {
        $RunId = New-RSTestRunId
    }
    $RunId = Assert-RSTestSandboxPathSegment -Value $RunId -Description 'TestSandbox run id'

    $runRoot = Join-Path (Join-Path $testRoot 'runs') $RunId
    $legacySelfTestRoot = Join-Path (Join-Path $runRoot 'artifacts') 'selftest'

    [pscustomobject]@{
        TestRoot = $testRoot
        RunId = $RunId
        RunRoot = $runRoot
        LegacySelfTestRoot = $legacySelfTestRoot
        ArtifactRoot = Join-Path $legacySelfTestRoot 'last_run'
    }
}

function Assert-RSTestSandboxPathSegment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "$Description must not be empty."
    }

    if ($Value -eq '.' -or $Value -eq '..' -or $Value.EndsWith('.') -or $Value -match '[\\/:\*\?"<>\|]') {
        throw "$Description '$Value' must be a single filesystem path segment."
    }

    return $Value
}

function New-RSTestSandboxScratchDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Harness,

        [Parameter(Mandatory = $true)]
        [string]$Case,

        [string]$RunId = $env:REDSALAMANDER_TEST_RUN_ID
    )

    $safeHarness = Assert-RSTestSandboxPathSegment -Value $Harness -Description 'TestSandbox harness segment'
    $safeCase = Assert-RSTestSandboxPathSegment -Value $Case -Description 'TestSandbox case segment'
    $context = New-RSTestRunContext -RepoRoot $RepoRoot -RunId $RunId
    [void](Initialize-RSTestSandboxRoot -TestRoot $context.TestRoot -RepoRoot $RepoRoot)
    $runRoot = $context.RunRoot
    $scratchRoot = Join-Path $runRoot 'scratch'
    $harnessRoot = Join-Path $scratchRoot $safeHarness
    $caseRoot = Join-Path $harnessRoot $safeCase

    New-Item -ItemType Directory -Path $caseRoot -Force | Out-Null
    return (ConvertTo-RSFullPath -Path $caseRoot)
}

function Get-RSToolingPesterSandboxRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$PrimaryTestRoot
    )

    $primary = (ConvertTo-RSFullPath -Path $PrimaryTestRoot).TrimEnd('\')
    [void](Assert-RSTestSandboxRootAuthorized -TestRoot $primary -RepoRoot $RepoRoot -RequireMarker)
    return $primary
}

function Assert-RSTestSandboxLinkTarget {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Target
    )

    # Mirrors Common::Testing::RequireSandboxedLinkTarget: a test link lives in, and points into,
    # one marked X:\RedSalamander.Perf. Both are strict descendants of that base with reparse-free
    # parents; the target leaf may itself be a guarded link. A relative target resolves against the
    # link's parent; a drive-relative or rooted target without a drive is refused.
    $link = ConvertTo-RSFullPath -Path $Path
    if ($Target -match '^[A-Za-z]:(?![\\/])' -or $Target -match '^[\\/](?![\\/])') {
        throw "Refusing test link '$link': target '$Target' depends on the current drive."
    }
    $targetPath = if ([IO.Path]::IsPathRooted($Target)) {
        ConvertTo-RSFullPath -Path $Target
    } else {
        ConvertTo-RSFullPath -Path (Join-Path ([IO.Path]::GetDirectoryName($link)) $Target)
    }
    $bases = foreach ($candidate in @($link, $targetPath)) {
        $driveRoot = [IO.Path]::GetPathRoot($candidate)
        $base = if ([string]::IsNullOrWhiteSpace($driveRoot) -or $driveRoot.StartsWith('\\')) { $null } else {
            Get-RSDedicatedExternalTestSandboxRoot -Path (Join-Path $driveRoot $script:RSTestSandboxDirectoryName)
        }
        if ($null -eq $base -or
            -not (Test-RSPathSameOrDescendant -Path $candidate -Parent $base) -or
            [string]::Equals($candidate.TrimEnd('\'), $base.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) -or
            $candidate.Substring($driveRoot.Length).Contains(':')) {
            throw ("Refusing test link '$link' -> '$targetPath': '$candidate' is not inside a marked X:\$script:RSTestSandboxDirectoryName. " +
                'Tooling Pester tests create links only under Tools\Run-AllTests.ps1, which places TEMP and $TestDrive in the run sandbox.')
        }
        [void](Assert-RSTestSandboxContainedPath -TestRoot $base -Path ([IO.Path]::GetDirectoryName($candidate)))
        $base
    }
    if (-not [string]::Equals($bases[0], $bases[1], [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing test link '$link' -> '$targetPath': the link and its target are in different sandboxes."
    }
    return $targetPath
}

function New-RSTestSandboxLink {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Target,
        [Parameter(Mandatory = $true)][ValidateSet('Junction', 'SymbolicLink', 'HardLink')][string]$ItemType
    )

    # The one governed creator of test links in PowerShell (Specs/Testing/Testing_SelfTests.md).
    # New-Item resolves a relative junction or hard-link target against the current location, not
    # the link's parent, so only a symbolic link may keep a relative target.
    if ($ItemType -ne 'SymbolicLink' -and -not [IO.Path]::IsPathFullyQualified($Target)) {
        throw "Refusing test $ItemType '$Path': target '$Target' must be a fully qualified path."
    }
    [void](Assert-RSTestSandboxLinkTarget -Path $Path -Target $Target)
    return (New-Item -ItemType $ItemType -Path $Path -Target $Target -ErrorAction Stop)
}

function Get-RSTestSandboxTreeEntries {
    param([Parameter(Mandatory = $true)][string]$Directory)

    # Default EnumerationOptions skip Hidden/System entries; a cleanup walk must see them.
    $options = [IO.EnumerationOptions]::new()
    $options.AttributesToSkip = [IO.FileAttributes]0
    $options.RecurseSubdirectories = $false
    $options.IgnoreInaccessible = $false
    return @([IO.DirectoryInfo]::new($Directory).GetFileSystemInfos('*', $options))
}

function Get-RSTestSandboxUnlinkableReparseRefusal {
    param([Parameter(Mandatory = $true)][IO.FileSystemInfo]$Entry)

    # Junctions and directory symlinks are removed as links only. A volume mount
    # point needs DeleteVolumeMountPointW (an administrator) or the mount-manager
    # registration is orphaned; any other directory reparse point has no known
    # link-only removal. Both refuse the whole tree.
    $target = $Entry.LinkTarget
    if ([string]::IsNullOrEmpty($target)) {
        return "unrecognized directory reparse point '$($Entry.FullName)'"
    }
    if ($target -match '^(\\\\\?\\|\\\?\?\\)?Volume\{') {
        return "volume mount point '$($Entry.FullName)' -> '$target' (an administrator must remove it with DeleteVolumeMountPointW/mountvol /D)"
    }
    return $null
}

function Get-RSTestSandboxTreeReparseRefusals {
    param([Parameter(Mandatory = $true)][string]$Path)

    # Walk without following links: never enumerate into a reparse directory.
    $refusals = [Collections.Generic.List[string]]::new()
    $pending = [Collections.Generic.Stack[string]]::new()
    $pending.Push($Path)
    while ($pending.Count -gt 0) {
        foreach ($entry in (Get-RSTestSandboxTreeEntries -Directory $pending.Pop())) {
            if (($entry.Attributes -band [IO.FileAttributes]::Directory) -eq 0) {
                continue
            }
            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) {
                $pending.Push($entry.FullName)
                continue
            }
            $refusal = Get-RSTestSandboxUnlinkableReparseRefusal -Entry $entry
            if ($null -ne $refusal) {
                $refusals.Add($refusal)
            }
        }
    }
    return @($refusals)
}

function Remove-RSTestSandboxTreeWithoutFollowingLinks {
    param([Parameter(Mandatory = $true)][string]$Path)

    # Returns refusal reasons and deletes nothing when any exist. Otherwise deletes
    # the tree post-order: reparse directories are unlinked with a non-recursive
    # delete and never enumerated, so no link can extend the delete outside the run.
    $root = (ConvertTo-RSFullPath -Path $Path).TrimEnd('\')
    if (([IO.File]::GetAttributes($root) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Refusing to remove reparse-point TestSandbox run root: $root"
    }
    $refusals = @(Get-RSTestSandboxTreeReparseRefusals -Path $root)
    if ($refusals.Count -gt 0) {
        return $refusals
    }
    $rootInfo = [IO.DirectoryInfo]::new($root)
    if (($rootInfo.Attributes -band [IO.FileAttributes]::ReadOnly) -ne 0) {
        $rootInfo.Attributes = $rootInfo.Attributes -band (-bnot [IO.FileAttributes]::ReadOnly)
    }

    $directories = [Collections.Generic.List[string]]::new()
    $pending = [Collections.Generic.Stack[string]]::new()
    $pending.Push($root)
    while ($pending.Count -gt 0) {
        $directory = $pending.Pop()
        # Re-read the attributes before enumerating: a directory replaced by a link
        # since the parent listing fails the removal instead of being entered.
        if (([IO.File]::GetAttributes($directory) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Refusing to enumerate TestSandbox directory that became a reparse point: $directory"
        }
        $directories.Add($directory)
        foreach ($entry in (Get-RSTestSandboxTreeEntries -Directory $directory)) {
            $attributes = $entry.Attributes
            $isDirectory = ($attributes -band [IO.FileAttributes]::Directory) -ne 0
            if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                if ($isDirectory) {
                    $refusal = Get-RSTestSandboxUnlinkableReparseRefusal -Entry $entry
                    if ($null -ne $refusal) {
                        throw "Refusing to unlink $refusal"
                    }
                    [IO.Directory]::Delete($entry.FullName, $false)
                } else {
                    [IO.File]::Delete($entry.FullName)
                }
                continue
            }
            if (($attributes -band [IO.FileAttributes]::ReadOnly) -ne 0) {
                $entry.Attributes = $attributes -band (-bnot [IO.FileAttributes]::ReadOnly)
            }
            if ($isDirectory) {
                $pending.Push($entry.FullName)
            } else {
                [IO.File]::Delete($entry.FullName)
            }
        }
    }
    for ($index = $directories.Count - 1; $index -ge 0; --$index) {
        [IO.Directory]::Delete($directories[$index], $false)
    }
    return @()
}

$script:RSTestOwnedVirtualDiskReceiptSchema = 'red-salamander.test-owned-vhdx-receipt.v1'
$script:RSTestOwnedVirtualDiskDirectoryName = 'test-owned-vhdx'
$script:RSTestOwnedVirtualDiskSizeMB = 128
$script:RSTestOwnedVirtualDiskVolumePattern = '^\\\\\?\\Volume\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}\\$'

# The only privileged storage operations test tooling performs. Each acts on the exact
# test-owned VHDX or its proven volume; Pester replaces this table with recording fakes.
$script:RSTestOwnedVirtualDiskOperations = @{
    IsElevated = {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        try {
            return ([Security.Principal.WindowsPrincipal]::new($identity)).IsInRole(
                [Security.Principal.WindowsBuiltInRole]::Administrator)
        } finally {
            $identity.Dispose()
        }
    }
    CreateFile = {
        param([string]$VhdxPath, [int]$SizeMB)
        $scriptPath = "$VhdxPath.create.diskpart.txt"
        [IO.File]::WriteAllText($scriptPath, "create vdisk file=`"$VhdxPath`" maximum=$SizeMB type=fixed`r`n", [Text.Encoding]::ASCII)
        try {
            $output = & (Join-Path $env:SystemRoot 'System32\diskpart.exe') /s $scriptPath 2>&1 | Out-String
            $exitCode = $LASTEXITCODE
        } finally {
            [IO.File]::Delete($scriptPath)
        }
        if ($exitCode -ne 0 -or -not [IO.File]::Exists($VhdxPath)) {
            throw "diskpart could not create the test-owned VHDX '$VhdxPath' (exit $exitCode): $output"
        }
    }
    Attach = {
        param([string]$VhdxPath)
        [void](Mount-DiskImage -ImagePath $VhdxPath -StorageType VHDX -Access ReadWrite -NoDriveLetter -ErrorAction Stop)
        $image = Get-DiskImage -ImagePath $VhdxPath -ErrorAction Stop
        if (-not $image.Attached -or $null -eq $image.Number) {
            throw "The test-owned VHDX '$VhdxPath' did not attach as a disk."
        }
        return [int]$image.Number
    }
    GetDiskFacts = {
        param([int]$DiskNumber)
        $disk = Get-Disk -Number $DiskNumber -ErrorAction Stop
        return [pscustomobject]@{
            DiskNumber = [int]$disk.Number
            Location = [string]$disk.Location
            BusType = [string]$disk.BusType
            PartitionStyle = [string]$disk.PartitionStyle
            Size = [uint64]$disk.Size
        }
    }
    Format = {
        param([int]$DiskNumber)
        Initialize-Disk -Number $DiskNumber -PartitionStyle GPT -ErrorAction Stop
        $partition = New-Partition -DiskNumber $DiskNumber -UseMaximumSize `
            -GptType '{ebd0a0a2-b9e5-4433-87c0-68b6b72699c7}' -ErrorAction Stop
        Set-Partition -DiskNumber $DiskNumber -PartitionNumber $partition.PartitionNumber `
            -NoDefaultDriveLetter $true -ErrorAction Stop
        [void](Format-Volume -Partition $partition -FileSystem NTFS -NewFileSystemLabel 'RSTestOwned' `
                -Force -Confirm:$false -ErrorAction Stop)
        $partition = Get-Partition -DiskNumber $DiskNumber -PartitionNumber $partition.PartitionNumber -ErrorAction Stop
        $volumes = @($partition.AccessPaths | Where-Object { $_ -match $script:RSTestOwnedVirtualDiskVolumePattern })
        if ($volumes.Count -ne 1) {
            throw "The test-owned partition on disk $DiskNumber has no single volume GUID path: $($partition.AccessPaths -join ', ')"
        }
        return [pscustomobject]@{ PartitionNumber = [int]$partition.PartitionNumber; VolumeGuidPath = [string]$volumes[0] }
    }
    GetVolumeFacts = {
        param([string]$VhdxPath, [string]$VolumeGuidPath)
        $image = Get-DiskImage -ImagePath $VhdxPath -ErrorAction Stop
        $disk = Get-Disk -Number $image.Number -ErrorAction Stop
        # Resolve the volume back to its disk independently of the partition that named it.
        $partitions = @(Get-Partition -ErrorAction Stop | Where-Object { @($_.AccessPaths) -contains $VolumeGuidPath })
        return [pscustomobject]@{
            ImageDiskNumber = [int]$image.Number
            DiskLocation = [string]$disk.Location
            DiskBusType = [string]$disk.BusType
            VolumeDiskNumbers = @($partitions | ForEach-Object { [int]$_.DiskNumber })
            VolumeAccessPaths = @($partitions | ForEach-Object { @($_.AccessPaths) })
        }
    }
    AddMount = {
        param([int]$DiskNumber, [int]$PartitionNumber, [string]$MountPath)
        Add-PartitionAccessPath -DiskNumber $DiskNumber -PartitionNumber $PartitionNumber `
            -AccessPath ($MountPath.TrimEnd('\') + '\') -ErrorAction Stop
    }
    Unmount = {
        param([string]$MountPath)
        $output = & (Join-Path $env:SystemRoot 'System32\mountvol.exe') ($MountPath.TrimEnd('\') + '\') /D 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
            throw "mountvol /D failed for '$MountPath' (exit $LASTEXITCODE): $output"
        }
    }
    IsAttached = {
        param([string]$VhdxPath)
        if (-not [IO.File]::Exists($VhdxPath)) {
            return $false
        }
        try {
            return [bool](Get-DiskImage -ImagePath $VhdxPath -ErrorAction Stop).Attached
        } catch {
            # A partial file that never attached; deleting an attached image fails instead.
            return $false
        }
    }
    Detach = {
        param([string]$VhdxPath)
        [void](Dismount-DiskImage -ImagePath $VhdxPath -ErrorAction Stop)
    }
}

function Get-RSTestOwnedVirtualDiskLayout {
    param(
        [Parameter(Mandatory = $true)][string]$RunRoot,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $safeName = Assert-RSTestSandboxPathSegment -Value $Name -Description 'Test-owned virtual disk name'
    $run = (ConvertTo-RSFullPath -Path $RunRoot).TrimEnd('\')
    $directory = Join-Path (Join-Path (Join-Path $run 'scratch') $script:RSTestOwnedVirtualDiskDirectoryName) $safeName
    $receiptDirectory = Join-Path (Join-Path $run 'artifacts') $script:RSTestOwnedVirtualDiskDirectoryName
    [pscustomobject]@{
        Name = $safeName
        RunRoot = $run
        Directory = $directory
        VhdxPath = Join-Path $directory 'disk.vhdx'
        MountPath = Join-Path $directory 'mount'
        ReceiptPath = Join-Path $receiptDirectory "$safeName.receipt.json"
    }
}

function Write-RSTestOwnedVirtualDiskReceipt {
    param(
        [Parameter(Mandatory = $true)]$Layout,
        [Parameter(Mandatory = $true)][ValidateSet('creating', 'mounting', 'mounted', 'detached')][string]$Stage,
        [string]$VolumeGuidPath = ''
    )

    $receiptDirectory = [IO.Path]::GetDirectoryName($Layout.ReceiptPath)
    Assert-RSTestSandboxPathIsReparseFree -Path $receiptDirectory
    [void][IO.Directory]::CreateDirectory($receiptDirectory)
    $json = [ordered]@{
        schema = $script:RSTestOwnedVirtualDiskReceiptSchema
        name = $Layout.Name
        run_root = $Layout.RunRoot
        vhdx_path = $Layout.VhdxPath
        mount_path = $Layout.MountPath
        volume_guid_path = $VolumeGuidPath
        stage = $Stage
        owner_pid = $PID
        written_utc = [DateTime]::UtcNow.ToString('o')
    } | ConvertTo-Json
    # Durable before the step it records: flushed to disk, then renamed over the previous stage.
    $temporary = "$($Layout.ReceiptPath).$PID.tmp"
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($json)
    $stream = [IO.FileStream]::new($temporary, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    } finally {
        $stream.Dispose()
    }
    [IO.File]::Move($temporary, $Layout.ReceiptPath, $true)
}

function Read-RSTestOwnedVirtualDiskReceipt {
    param(
        [Parameter(Mandatory = $true)][string]$ReceiptPath,
        [Parameter(Mandatory = $true)][string]$RunRoot
    )

    $path = ConvertTo-RSFullPath -Path $ReceiptPath
    try {
        Assert-RSTestSandboxPathIsReparseFree -Path $path
        $value = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ([string]$value.schema -ne $script:RSTestOwnedVirtualDiskReceiptSchema) {
            throw "unexpected schema '$($value.schema)'"
        }
        # A receipt may name only the exact paths derived from its own run and name;
        # nothing it records outside that layout is ever unmounted, detached or deleted.
        $layout = Get-RSTestOwnedVirtualDiskLayout -RunRoot $RunRoot -Name ([string]$value.name)
        foreach ($pair in @(
                @([string]$value.run_root, $layout.RunRoot),
                @([string]$value.vhdx_path, $layout.VhdxPath),
                @([string]$value.mount_path, $layout.MountPath),
                @($path, $layout.ReceiptPath))) {
            if (-not [string]::Equals($pair[0].TrimEnd('\'), $pair[1].TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
                throw "path '$($pair[0])' is not the run-owned '$($pair[1])'"
            }
        }
        $volumeGuidPath = [string]$value.volume_guid_path
        if (-not [string]::IsNullOrEmpty($volumeGuidPath) -and $volumeGuidPath -notmatch $script:RSTestOwnedVirtualDiskVolumePattern) {
            throw "malformed volume GUID path '$volumeGuidPath'"
        }
        $layout | Add-Member -NotePropertyName VolumeGuidPath -NotePropertyValue $volumeGuidPath
        $layout | Add-Member -NotePropertyName Stage -NotePropertyValue ([string]$value.stage)
        return $layout
    } catch {
        throw "Malformed test-owned virtual disk receipt '$path': $($_.Exception.Message)"
    }
}

function Get-RSTestOwnedVirtualDiskReceiptPaths {
    param([Parameter(Mandatory = $true)][string]$RunRoot)

    $receiptDirectory = Join-Path (Join-Path (ConvertTo-RSFullPath -Path $RunRoot).TrimEnd('\') 'artifacts') $script:RSTestOwnedVirtualDiskDirectoryName
    try {
        $attributes = [IO.File]::GetAttributes($receiptDirectory)
    } catch [IO.FileNotFoundException] {
        return @()
    } catch [IO.DirectoryNotFoundException] {
        return @()
    }
    if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Refusing reparse-point test-owned virtual disk receipt directory: $receiptDirectory"
    }
    return @(Get-RSTestSandboxTreeEntries -Directory $receiptDirectory |
            Where-Object { ($_.Attributes -band [IO.FileAttributes]::Directory) -eq 0 -and $_.Name -like '*.receipt.json' } |
            ForEach-Object { $_.FullName } |
            Sort-Object)
}

function Assert-RSTestOwnedVirtualDiskBlankDisk {
    param(
        [Parameter(Mandatory = $true)][string]$VhdxPath,
        [Parameter(Mandatory = $true)][int]$DiskNumber,
        [Parameter(Mandatory = $true)]$Facts
    )

    # Pure pre-format proof: the disk number the attach returned must name exactly the blank,
    # file-backed disk this run's VHDX just created. Nothing initializes, partitions or formats
    # a disk that fails it.
    $expectedSize = [uint64]$script:RSTestOwnedVirtualDiskSizeMB * 1MB
    if ([int]$Facts.DiskNumber -ne $DiskNumber) {
        throw "Refusing to format disk ${DiskNumber}: its facts describe disk $($Facts.DiskNumber)."
    }
    if (-not [string]::Equals(([string]$Facts.Location).TrimEnd('\'), $VhdxPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to format disk ${DiskNumber}: it is backed by '$($Facts.Location)', not the test-owned VHDX '$VhdxPath'."
    }
    if ([string]$Facts.BusType -notin @('15', 'File Backed Virtual')) {
        throw "Refusing to format disk ${DiskNumber}: bus type '$($Facts.BusType)' is not a file-backed virtual disk."
    }
    if ([string]$Facts.PartitionStyle -notin @('RAW', '0')) {
        throw "Refusing to format disk ${DiskNumber}: partition style '$($Facts.PartitionStyle)' is not a blank RAW disk."
    }
    if ([uint64]$Facts.Size -ne $expectedSize) {
        throw "Refusing to format disk ${DiskNumber}: size $($Facts.Size) is not the test-owned VHDX size $expectedSize."
    }
}

function Assert-RSTestOwnedVirtualDiskVolume {
    param(
        [Parameter(Mandatory = $true)][string]$VhdxPath,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$VolumeGuidPath,
        [Parameter(Mandatory = $true)]$Facts
    )

    # Pure ownership proof, required before any mount: the volume must sit on exactly the
    # disk that this run's VHDX file backs, and nothing else may already reach it.
    if ($VolumeGuidPath -notmatch $script:RSTestOwnedVirtualDiskVolumePattern) {
        throw "Refusing malformed volume GUID path '$VolumeGuidPath'."
    }
    if (-not [string]::Equals(([string]$Facts.DiskLocation).TrimEnd('\'), $VhdxPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing volume '$VolumeGuidPath': disk $($Facts.ImageDiskNumber) is backed by '$($Facts.DiskLocation)', not the test-owned VHDX '$VhdxPath'."
    }
    if ([string]$Facts.DiskBusType -notin @('15', 'File Backed Virtual')) {
        throw "Refusing volume '$VolumeGuidPath': disk $($Facts.ImageDiskNumber) has bus type '$($Facts.DiskBusType)', not a file-backed virtual disk."
    }
    $diskNumbers = @($Facts.VolumeDiskNumbers | Select-Object -Unique)
    if ($diskNumbers.Count -ne 1 -or [int]$diskNumbers[0] -ne [int]$Facts.ImageDiskNumber) {
        throw "Refusing foreign volume '$VolumeGuidPath': it is on disk(s) '$($diskNumbers -join ', ')', not the test-owned virtual disk $($Facts.ImageDiskNumber)."
    }
    $otherPaths = @($Facts.VolumeAccessPaths | Where-Object {
            -not [string]::Equals([string]$_, $VolumeGuidPath, [StringComparison]::OrdinalIgnoreCase)
        })
    if ($otherPaths.Count -gt 0) {
        throw "Refusing volume '$VolumeGuidPath': it is already reachable through '$($otherPaths -join ', ')'."
    }
}

function Get-RSTestOwnedVirtualDiskMountTarget {
    param([Parameter(Mandatory = $true)][string]$MountPath)

    # Reads the mount directory's own reparse data: no privilege, and it still names the
    # volume after the disk detached. $null means a plain or absent directory.
    try {
        $attributes = [IO.File]::GetAttributes($MountPath)
    } catch [IO.FileNotFoundException] {
        return $null
    } catch [IO.DirectoryNotFoundException] {
        return $null
    }
    if (($attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) {
        return $null
    }
    $target = [string][IO.DirectoryInfo]::new($MountPath).LinkTarget
    if ($target -match '^(?:\\\\\?\\|\\\?\?\\)?(Volume\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\})\\?$') {
        return "\\?\$($Matches[1])\"
    }
    return "<non-volume reparse point '$target'>"
}

function Get-RSTestOwnedVirtualDiskRecoveryText {
    param([Parameter(Mandatory = $true)]$Receipt)

    $module = Join-Path (Get-RSTestSandboxRepositoryRoot) 'Tools\Modules\Testing\TestRunPlan.psm1'
    return ("From an elevated PowerShell 7 prompt run: Import-Module '$module'; " +
        "Remove-RSTestOwnedVirtualDiskMount -ReceiptPath '$($Receipt.ReceiptPath)' -RunRoot '$($Receipt.RunRoot)'. " +
        "Manual equivalent, in this order: mountvol `"$($Receipt.MountPath)\`" /D; " +
        "Dismount-DiskImage -ImagePath '$($Receipt.VhdxPath)'; " +
        "Remove-Item -LiteralPath '$($Receipt.VhdxPath)', '$($Receipt.ReceiptPath)'. Then rerun the cleanup.")
}

function Remove-RSTestOwnedVirtualDiskMount {
    param(
        [Parameter(Mandatory = $true)][string]$ReceiptPath,
        [Parameter(Mandatory = $true)][string]$RunRoot,
        # Unmount and detach, but keep the VHDX file in the run with a 'detached' receipt so
        # evidence written to the volume survives; it is reclaimed when the run is removed.
        [switch]$RetainImage
    )

    $receipt = Read-RSTestOwnedVirtualDiskReceipt -ReceiptPath $ReceiptPath -RunRoot $RunRoot
    $operations = $script:RSTestOwnedVirtualDiskOperations
    # A detached image needs no privilege to reclaim: it must be neither mounted nor attached,
    # and deleting a file that something attached again fails instead of succeeding.
    $elevated = [bool](& $operations.IsElevated)
    $requiresElevation = "Tearing down the test-owned virtual disk '$($receipt.VhdxPath)' requires elevation. $(Get-RSTestOwnedVirtualDiskRecoveryText -Receipt $receipt)"
    if (-not $elevated -and $receipt.Stage -ne 'detached') {
        throw $requiresElevation
    }
    Assert-RSTestSandboxPathIsReparseFree -Path $receipt.Directory

    # Reverse of creation; each step must succeed before the next. The receipt goes last.
    $target = Get-RSTestOwnedVirtualDiskMountTarget -MountPath $receipt.MountPath
    if ($null -ne $target) {
        if (-not $elevated) {
            throw $requiresElevation
        }
        if ([string]::IsNullOrEmpty($receipt.VolumeGuidPath) -or
            -not [string]::Equals($target, $receipt.VolumeGuidPath, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to unmount '$($receipt.MountPath)': it names '$target', not the receipt's test-owned volume '$($receipt.VolumeGuidPath)'."
        }
        & $operations.Unmount $receipt.MountPath
        if ($null -ne (Get-RSTestOwnedVirtualDiskMountTarget -MountPath $receipt.MountPath)) {
            throw "The test-owned mount '$($receipt.MountPath)' is still present after unmount."
        }
    }
    if ([bool](& $operations.IsAttached $receipt.VhdxPath)) {
        if (-not $elevated) {
            throw $requiresElevation
        }
        & $operations.Detach $receipt.VhdxPath
        if ([bool](& $operations.IsAttached $receipt.VhdxPath)) {
            throw "The test-owned VHDX '$($receipt.VhdxPath)' is still attached after detach."
        }
    }
    if ($RetainImage) {
        Write-RSTestOwnedVirtualDiskReceipt -Layout $receipt -Stage detached -VolumeGuidPath $receipt.VolumeGuidPath
        return
    }
    if ([IO.File]::Exists($receipt.VhdxPath)) {
        [IO.File]::Delete($receipt.VhdxPath)
    }
    [IO.File]::Delete($receipt.ReceiptPath)
}

function New-RSTestOwnedVirtualDiskMount {
    param(
        [Parameter(Mandatory = $true)][string]$TestRoot,
        [Parameter(Mandatory = $true)][string]$RunId,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $operations = $script:RSTestOwnedVirtualDiskOperations
    if (-not [bool](& $operations.IsElevated)) {
        return [pscustomobject]@{
            Status = 'Skipped'
            Reason = 'not-elevated: creating, attaching and mounting a test-owned VHDX needs an elevated Run-AllTests'
        }
    }

    $root = Assert-RSTestSandboxRootAuthorized -TestRoot $TestRoot -RequireMarker
    $safeRunId = Assert-RSTestSandboxPathSegment -Value $RunId -Description 'TestSandbox run id'
    $layout = Get-RSTestOwnedVirtualDiskLayout -RunRoot (Join-Path (Join-Path $root 'runs') $safeRunId) -Name $Name
    [void](Assert-RSTestSandboxContainedPath -TestRoot $root -Path $layout.MountPath -RequireDescendant)
    [void](Assert-RSTestSandboxContainedPath -TestRoot $root -Path $layout.ReceiptPath -RequireDescendant)
    if ([IO.File]::Exists($layout.ReceiptPath) -or [IO.File]::Exists($layout.VhdxPath) -or [IO.Directory]::Exists($layout.MountPath)) {
        throw "A test-owned virtual disk already exists for '$($layout.Name)' in this run; tear it down first: $($layout.ReceiptPath)"
    }
    [void][IO.Directory]::CreateDirectory($layout.MountPath)
    # The receipt exists before the disk file, so a crash at any later point leaves recovery evidence.
    Write-RSTestOwnedVirtualDiskReceipt -Layout $layout -Stage creating
    try {
        & $operations.CreateFile $layout.VhdxPath $script:RSTestOwnedVirtualDiskSizeMB
        $diskNumber = [int](& $operations.Attach $layout.VhdxPath)
        Assert-RSTestOwnedVirtualDiskBlankDisk -VhdxPath $layout.VhdxPath -DiskNumber $diskNumber `
            -Facts (& $operations.GetDiskFacts $diskNumber)
        $partition = & $operations.Format $diskNumber
        $volumeGuidPath = [string]$partition.VolumeGuidPath
        Assert-RSTestOwnedVirtualDiskVolume -VhdxPath $layout.VhdxPath -VolumeGuidPath $volumeGuidPath `
            -Facts (& $operations.GetVolumeFacts $layout.VhdxPath $volumeGuidPath)
        [void](Assert-RSTestSandboxContainedPath -TestRoot $root -Path $layout.MountPath -RequireDescendant)
        if (@(Get-RSTestSandboxTreeEntries -Directory $layout.MountPath).Count -ne 0) {
            throw "The owned mount directory is not empty: $($layout.MountPath)"
        }
        Write-RSTestOwnedVirtualDiskReceipt -Layout $layout -Stage mounting -VolumeGuidPath $volumeGuidPath
        & $operations.AddMount $diskNumber ([int]$partition.PartitionNumber) $layout.MountPath
        $target = Get-RSTestOwnedVirtualDiskMountTarget -MountPath $layout.MountPath
        if (-not [string]::Equals($target, $volumeGuidPath, [StringComparison]::OrdinalIgnoreCase)) {
            throw "The owned mount '$($layout.MountPath)' names '$target', not the proven volume '$volumeGuidPath'."
        }
        Write-RSTestOwnedVirtualDiskReceipt -Layout $layout -Stage mounted -VolumeGuidPath $volumeGuidPath
        return [pscustomobject]@{
            Status = 'Mounted'
            Reason = $null
            RunRoot = $layout.RunRoot
            MountPath = $layout.MountPath
            VolumeGuidPath = $volumeGuidPath
            VhdxPath = $layout.VhdxPath
            ReceiptPath = $layout.ReceiptPath
        }
    } catch {
        $failure = $_.Exception.Message
        try {
            Remove-RSTestOwnedVirtualDiskMount -ReceiptPath $layout.ReceiptPath -RunRoot $layout.RunRoot
        } catch {
            throw "Test-owned virtual disk provisioning failed: $failure Teardown also failed; the receipt is retained: $($_.Exception.Message)"
        }
        throw "Test-owned virtual disk provisioning failed and was torn down: $failure"
    }
}

function Get-RSTestOwnedVirtualDiskRunRefusals {
    param([Parameter(Mandatory = $true)][string]$RunRoot)

    # Runs before any sweep of a run tree. Without elevation a receipt retains the whole run
    # and reports how to recover, except a 'detached' image, which is reclaimed with the run;
    # elevated, each recorded disk is torn down first. Empty output means the run holds no
    # recorded disk any more.
    $run = (ConvertTo-RSFullPath -Path $RunRoot).TrimEnd('\')
    $receiptPaths = @()
    try {
        $receiptPaths = @(Get-RSTestOwnedVirtualDiskReceiptPaths -RunRoot $run)
    } catch {
        return @("unreadable test-owned virtual disk receipts: $($_.Exception.Message)")
    }
    if ($receiptPaths.Count -eq 0) {
        return @()
    }
    $elevated = [bool](& $script:RSTestOwnedVirtualDiskOperations.IsElevated)
    $refusals = [Collections.Generic.List[string]]::new()
    foreach ($receiptPath in $receiptPaths) {
        $receipt = $null
        try {
            $receipt = Read-RSTestOwnedVirtualDiskReceipt -ReceiptPath $receiptPath -RunRoot $run
        } catch {
            $refusals.Add("$($_.Exception.Message); inspect it and tear its disk down by hand from an elevated prompt")
            continue
        }
        if (-not $elevated -and $receipt.Stage -ne 'detached') {
            $refusals.Add("test-owned virtual disk receipt '$receiptPath' (the disk may still be attached or mounted; cleanup is not elevated). $(Get-RSTestOwnedVirtualDiskRecoveryText -Receipt $receipt)")
            continue
        }
        try {
            Remove-RSTestOwnedVirtualDiskMount -ReceiptPath $receiptPath -RunRoot $run
        } catch {
            $refusals.Add("test-owned virtual disk receipt '$receiptPath' could not be torn down: $($_.Exception.Message) $(Get-RSTestOwnedVirtualDiskRecoveryText -Receipt $receipt)")
        }
    }
    return @($refusals)
}

function Remove-RSTestSandboxExactRunDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestRoot,

        [Parameter(Mandatory = $true)]
        [string]$RunId
    )

    $safeRunId = Assert-RSTestSandboxPathSegment `
        -Value $RunId `
        -Description 'TestSandbox run id'
    $root = ConvertTo-RSFullPath -Path $TestRoot
    [void](Assert-RSTestSandboxRootAuthorized -TestRoot $root -RequireMarker)
    $runsRoot = ConvertTo-RSFullPath -Path (Join-Path $root 'runs')
    $runRoot = ConvertTo-RSFullPath -Path (Join-Path $runsRoot $safeRunId)
    if (-not $runRoot.StartsWith(
            "$runsRoot\",
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove TestSandbox run outside '$runsRoot': $runRoot"
    }
    if (-not (Test-Path -LiteralPath $runRoot)) {
        return
    }
    if (((Get-Item -LiteralPath $runRoot -Force).Attributes -band
            [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Refusing to remove reparse-point TestSandbox run root: $runRoot"
    }

    $virtualDiskRefusals = @(Get-RSTestOwnedVirtualDiskRunRefusals -RunRoot $runRoot)
    if ($virtualDiskRefusals.Count -gt 0) {
        throw "Refusing to remove TestSandbox run '$runRoot'; the whole tree is retained because it contains: $($virtualDiskRefusals -join '; ')"
    }
    $refusals = @(Remove-RSTestSandboxTreeWithoutFollowingLinks -Path $runRoot)
    if ($refusals.Count -gt 0) {
        throw "Refusing to remove TestSandbox run '$runRoot'; the whole tree is retained because it contains: $($refusals -join '; ')"
    }
}

function Resolve-RSTestSandboxStaleRunTargets {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestRoot,

        [string]$RunId = '',

        [string[]]$AllowedRunIds = @(),

        [int64[]]$LiveProcessIds = @(),

        [hashtable]$LiveProcessStartTimesUtc = @{}
    )

    $normalizedTestRoot = ConvertTo-RSFullPath -Path $TestRoot
    [void](Assert-RSTestSandboxRootAuthorized -TestRoot $normalizedTestRoot -RequireMarker)
    $runsRoot = Join-Path $normalizedTestRoot 'runs'
    if (-not (Test-Path -LiteralPath $runsRoot)) {
        return @()
    }

    $allowedRunIds = @($AllowedRunIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (-not [string]::IsNullOrWhiteSpace($RunId)) {
        $allowedRunIds += $RunId
    }
    $allowedRunIds = @($allowedRunIds | Select-Object -Unique)

    $liveIds = @()
    $liveStartTimesUtc = @{}
    if ($PSBoundParameters.ContainsKey('LiveProcessIds')) {
        $liveIds = @($LiveProcessIds | ForEach-Object { [int64]$_ })
        foreach ($key in $LiveProcessStartTimesUtc.Keys) {
            $liveStartTimesUtc[[string]$key] = [datetime]$LiveProcessStartTimesUtc[$key]
        }
    } else {
        foreach ($process in @(Get-RSLiveProcessSnapshots)) {
            $liveIds += [int64]$process.ProcessId
            if ($null -ne $process.StartTimeUtc) {
                $liveStartTimesUtc[[string]$process.ProcessId] = [datetime]$process.StartTimeUtc
            }
        }
    }

    $targets = @()
    foreach ($runDir in @(Get-ChildItem -LiteralPath $runsRoot -Directory -Force -ErrorAction SilentlyContinue)) {
        $isAllowedRun = @($allowedRunIds | Where-Object {
                [string]::Equals($_, $runDir.Name, [System.StringComparison]::OrdinalIgnoreCase)
            }).Count -gt 0
        if ($isAllowedRun) {
            continue
        }

        $runInfo = ConvertFrom-RSTestRunId -RunId $runDir.Name
        if ($null -eq $runInfo) {
            continue
        }

        if (@($liveIds | Where-Object { $_ -eq $runInfo.ProcessId }).Count -gt 0) {
            $processKey = [string]$runInfo.ProcessId
            if (-not $liveStartTimesUtc.ContainsKey($processKey) -or
                $runDir.CreationTimeUtc -ge ([datetime]$liveStartTimesUtc[$processKey]).ToUniversalTime()) {
                continue
            }
        }

        $targets += [pscustomobject]@{
            Path = $runDir.FullName
            RunId = $runDir.Name
            ProcessId = $runInfo.ProcessId
            Category = 'stale-test-run-dir'
            Reason = "The owning process $($runInfo.ProcessId) is no longer live (or its PID has been reused by a newer process); the sibling TestSandbox run directory can be swept before this run starts."
        }
    }

    return $targets
}

function Get-RSTestRunArchiveStreams {
    param([string]$Path)
    # The PowerShell -Stream provider rejects some otherwise valid long Windows paths.
    # Initialize only when invoked; module import remains read-only.
    if (-not ('RedSalamander.Tooling.RunArchiveStreams' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
namespace RedSalamander.Tooling {
    public sealed class ArchiveStream {
        public string Name { get; set; }
        public long Length { get; set; }
    }
    public static class RunArchiveStreams {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct StreamData {
            public long Size;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 296)] public string Name;
        }
        sealed class FindHandle : SafeHandleZeroOrMinusOneIsInvalid {
            public FindHandle() : base(true) { }
            protected override bool ReleaseHandle() { return FindClose(handle); }
        }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern FindHandle FindFirstStreamW(string path, int level, out StreamData data, uint flags);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern bool FindNextStreamW(FindHandle handle, out StreamData data);
        [DllImport("kernel32.dll", SetLastError = true)] static extern bool FindClose(IntPtr handle);
        public static ArchiveStream[] Read(string path) {
            var result = new List<ArchiveStream>();
            string extended = path.StartsWith(@"\\?\", StringComparison.Ordinal) ? path : @"\\?\" + path;
            using (var iterator = FindFirstStreamW(extended, 0, out var data, 0)) {
                if (iterator.IsInvalid) {
                    int openError = Marshal.GetLastWin32Error();
                    if (openError == 38) return result.ToArray();
                    throw new Win32Exception(openError, "Cannot inventory streams: " + path);
                }
                do {
                    if (data.Name != "::$DATA") result.Add(new ArchiveStream { Name = data.Name, Length = data.Size });
                } while (FindNextStreamW(iterator, out data));
                int error = Marshal.GetLastWin32Error();
                if (error != 38) throw new Win32Exception(error, "Cannot finish stream inventory: " + path);
            }
            return result.ToArray();
        }
    }
}
'@
    }
    return [RedSalamander.Tooling.RunArchiveStreams]::Read($Path)
}

function Get-RSTestRunArchiveFileHash {
    param([string]$Path)
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try { return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream)) }
    finally { $stream.Dispose() }
}

function Get-RSTestRunArchiveInventory {
    param([string]$RunRoot)

    # Enumerate one directory at a time, including hidden/system entries. Never follow a link.
    $rows = [Collections.Generic.List[object]]::new()
    $pending = [Collections.Generic.Stack[string]]::new()
    $pending.Push($RunRoot)
    while ($pending.Count -gt 0) {
        $directory = $pending.Pop()
        Assert-RSTestSandboxPathIsReparseFree -Path $directory
        $entries = @([IO.DirectoryInfo]::new($directory))
        $entries += @(Get-RSTestSandboxTreeEntries -Directory $directory)
        foreach ($entry in $entries) {
            $isDirectory = ($entry.Attributes -band [IO.FileAttributes]::Directory) -ne 0
            $isLink = ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
            $relative = [IO.Path]::GetRelativePath($RunRoot, $entry.FullName)
            # The current directory was recorded by its parent (except the root).
            if ($entry.FullName -eq $directory -and $directory -ne $RunRoot) { continue }
            $target = $null
            if ($isLink) {
                $target = $entry.LinkTarget
                # Only self-contained relative symlinks survive relocation unchanged.
                if ([string]::IsNullOrWhiteSpace($target) -or [IO.Path]::IsPathRooted($target) -or
                    $target.Contains(':') -or -not (Test-RSPathSameOrDescendant `
                        -Path ([IO.Path]::GetFullPath($target, [IO.Path]::GetDirectoryName($entry.FullName))) -Parent $RunRoot)) {
                    throw "Run archive retains unsafe or mount-like reparse point: $($entry.FullName)"
                }
            }
            $streams = @()
            if (-not $isLink) {
                foreach ($stream in @(Get-RSTestRunArchiveStreams -Path $entry.FullName | Sort-Object Name)) {
                    $streamPath = $entry.FullName + $stream.Name
                    $streams += [ordered]@{ name = $stream.Name; length = $stream.Length; sha256 = Get-RSTestRunArchiveFileHash -Path $streamPath }
                }
            }
            $rows.Add([ordered]@{
                path = $relative; directory = $isDirectory; attributes = [int]$entry.Attributes
                created_filetime_utc = $entry.CreationTimeUtc.ToFileTimeUtc(); written_filetime_utc = $entry.LastWriteTimeUtc.ToFileTimeUtc()
                security = $(if (-not $isLink) { (Get-Acl -LiteralPath $entry.FullName -ErrorAction Stop).Sddl } else { $null })
                target = $target; streams = $streams
                length = $(if (-not $isDirectory -and -not $isLink) { $entry.Length } else { 0 })
                sha256 = $(if (-not $isDirectory -and -not $isLink) { Get-RSTestRunArchiveFileHash -Path $entry.FullName } else { $null })
            })
            if ($isDirectory -and -not $isLink -and $entry.FullName -ne $directory) { $pending.Push($entry.FullName) }
        }
    }
    return @($rows | Sort-Object { $_.path })
}

function Write-RSTestRunArchiveRecord {
    param([string]$Path, [string]$Json)
    # Create-only, flushed records. A crash leaves a recoverable intent, never an overwritten receipt.
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($Json)
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    $stream = [IO.FileStream]::new($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
    [IO.File]::Move($temporary, $Path)
}

function Get-RSTestRunArchiveStableInventory {
    param([string]$RunRoot)
    # Newly closed Windows handles can finish publishing directory metadata after enumeration.
    # Require two equal acquisitions before recording an immutable boundary; never rebaseline an intent.
    $previous = $null
    foreach ($attempt in 1..3) {
        $rows = @(Get-RSTestRunArchiveInventory -RunRoot $RunRoot)
        $json = ConvertTo-Json -InputObject $rows -Depth 16 -Compress
        if ($null -ne $previous -and $json -ceq $previous) { return $rows }
        $previous = $json
    }
    throw "Run archive source has no stable inventory; retaining it: $RunRoot"
}

function Save-RSTestSandboxRunEvidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TestRoot,
        [Parameter(Mandatory = $true)][string]$RunId,
        [ValidateSet('completed', 'interrupted')][string]$Disposition = 'completed',
        [switch]$Abandoned
    )

    $root = Assert-RSTestSandboxRootAuthorized -TestRoot $TestRoot -RequireMarker
    $id = Assert-RSTestSandboxPathSegment -Value $RunId -Description 'Run archive id'
    $source = Assert-RSTestSandboxContainedPath -TestRoot $root -Path (Join-Path $root "runs\$id") -RequireDescendant
    $destination = Assert-RSTestSandboxContainedPath -TestRoot $root -Path (Join-Path $root "evidence\completed-runs\$id") -RequireDescendant
    $records = Assert-RSTestSandboxContainedPath -TestRoot $root -Path (Join-Path $root "evidence\run-archive-records\$id") -RequireDescendant
    if ($Abandoned -and [IO.Directory]::Exists($source)) {
        # Recheck real liveness immediately before the transaction; never accept synthetic PID lists here.
        if (@(Resolve-RSTestSandboxStaleRunTargets -TestRoot $root | Where-Object { $_.RunId -eq $id }).Count -ne 1) {
            throw "Run archive refuses a live, allowed, or unparseable owner: $source"
        }
    } elseif (-not $Abandoned) {
        $owner = ConvertFrom-RSTestRunId -RunId $id
        if ($null -ne $owner -and $owner.ProcessId -ne $PID) { throw "Run archive refuses another live runner's id: $id" }
    }
    [void][IO.Directory]::CreateDirectory($records)
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
    [void](Assert-RSTestSandboxContainedPath -TestRoot $root -Path $records -RequireDescendant)
    # Same-run transactions serialize across configurations and processes. No global test/build lock.
    $lockPath = Join-Path $records 'archive.lock'
    [void](Assert-RSTestSandboxContainedPath -TestRoot $root -Path $lockPath -RequireDescendant)
    $lease = [IO.FileStream]::new($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try {
        $intentPath = Join-Path $records 'intent.json'
        $completePath = Join-Path $records 'complete.json'
        foreach ($path in @($source, $destination, $intentPath, $completePath)) {
            [void](Assert-RSTestSandboxContainedPath -TestRoot $root -Path $path -RequireDescendant)
        }
        if ([IO.File]::Exists($completePath)) { throw "Run archive already completed; refusing overwrite: $completePath" }
        if ([IO.Directory]::Exists($source) -and (Test-Path -LiteralPath $destination)) {
            throw "Run archive destination already exists; both trees retained: $destination"
        }
        if ([IO.Directory]::Exists($source)) {
            # Repeat on recovery too: a prior detached image may have been attached since the intent.
            foreach ($receiptPath in @(Get-RSTestOwnedVirtualDiskReceiptPaths -RunRoot $source)) {
                $receipt = Read-RSTestOwnedVirtualDiskReceipt -ReceiptPath $receiptPath -RunRoot $source
                if ($receipt.Stage -ne 'detached' -or (Get-DiskImage -ImagePath $receipt.VhdxPath -ErrorAction Stop).Attached) {
                    throw "Run archive retains unresolved test-owned VHDX receipt: $receiptPath"
                }
            }
        }
        if ([IO.File]::Exists($intentPath)) {
            $intent = Get-Content -LiteralPath $intentPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($intent.schema -ne 'red-salamander.run-archive.v1' -or $intent.source -cne $source -or $intent.destination -cne $destination) {
                throw "Run archive intent does not match the selected sandbox/run: $intentPath"
            }
        } else {
            if (-not [IO.Directory]::Exists($source)) { throw "Run archive source is absent: $source" }
            # Never invoke the cleanup helper that can detach/delete disks to make an archive eligible.
            $intent = [ordered]@{
                schema = 'red-salamander.run-archive.v1'; run_id = $id; source = $source; destination = $destination
                disposition = $Disposition; written_utc = [DateTime]::UtcNow.ToString('o')
                inventory = @(Get-RSTestRunArchiveStableInventory -RunRoot $source)
            }
            Write-RSTestRunArchiveRecord -Path $intentPath -Json ($intent | ConvertTo-Json -Depth 16 -Compress)
        }
        $expected = ConvertTo-Json -InputObject @($intent.inventory) -Depth 16 -Compress
        if ([IO.Directory]::Exists($source)) {
            $actual = ConvertTo-Json -InputObject @(Get-RSTestRunArchiveInventory -RunRoot $source) -Depth 16 -Compress
            if ($actual -cne $expected) { throw "Run archive source changed after intent; retaining it: $source" }
            # Fresh checks, no recursive move/copy or overwrite. Both endpoints are on the same marked volume.
            foreach ($path in @($source, $destination)) {
                [void](Assert-RSTestSandboxContainedPath -TestRoot $root -Path $path -RequireDescendant)
            }
            [IO.Directory]::Move($source, $destination)
        }
        $actual = ConvertTo-Json -InputObject @(Get-RSTestRunArchiveInventory -RunRoot $destination) -Depth 16 -Compress
        if ($actual -cne $expected) { throw "Run archive verification failed; retained evidence needs review: $destination" }
        Write-RSTestRunArchiveRecord -Path $completePath -Json (([ordered]@{
            schema = 'red-salamander.run-archive-complete.v1'; source = $source; destination = $destination
            intent_sha256 = (Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash; verified_utc = [DateTime]::UtcNow.ToString('o')
        }) | ConvertTo-Json -Compress)
        return [pscustomobject]@{
            Status = 'Archived'; RunId = $id; SourceRoot = $source; RunRoot = $destination
            ArtifactRoot = Join-Path $destination 'artifacts\selftest\last_run'; RecordPath = $completePath
        }
    } finally { $lease.Dispose() }
}

function Save-RSTestSandboxStaleRunEvidence {
    param([Parameter(Mandatory = $true)][string]$TestRoot, [string]$RunId = '')
    foreach ($target in @(Resolve-RSTestSandboxStaleRunTargets -TestRoot $TestRoot -RunId $RunId)) {
        try { Save-RSTestSandboxRunEvidence -TestRoot $TestRoot -RunId $target.RunId -Abandoned -Disposition interrupted }
        catch { [pscustomobject]@{ Status = 'Retained'; RunId = $target.RunId; Reason = $_.Exception.Message } }
    }
}

function Remove-RSTestSandboxStaleRunDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestRoot,

        [string]$RunId = '',

        [string[]]$AllowedRunIds = @(),

        [string[]]$CandidateRunIds = @(),

        [int64[]]$LiveProcessIds = @(),

        [hashtable]$LiveProcessStartTimesUtc = @{}
    )

    # Synthetic liveness is used by fixtures. Never let that view authorize a root-wide sweep.
    if ($PSBoundParameters.ContainsKey('LiveProcessIds') -and -not $PSBoundParameters.ContainsKey('CandidateRunIds')) {
        throw 'A synthetic process snapshot requires an explicit CandidateRunIds cleanup scope.'
    }
    foreach ($candidate in $CandidateRunIds) {
        [void](Assert-RSTestSandboxPathSegment -Value $candidate -Description 'Cleanup candidate run id')
    }
    $normalizedTestRoot = ConvertTo-RSFullPath -Path $TestRoot
    [void](Assert-RSTestSandboxRootAuthorized -TestRoot $normalizedTestRoot -RequireMarker)
    $runsRoot = (Join-Path $normalizedTestRoot 'runs').TrimEnd('\')
    $resolveParams = @{
        TestRoot = $normalizedTestRoot
        RunId = $RunId
        AllowedRunIds = @($AllowedRunIds)
    }
    if ($PSBoundParameters.ContainsKey('LiveProcessIds')) {
        $resolveParams['LiveProcessIds'] = @($LiveProcessIds)
        $resolveParams['LiveProcessStartTimesUtc'] = $LiveProcessStartTimesUtc
    }

    $results = @()
    foreach ($target in @(Resolve-RSTestSandboxStaleRunTargets @resolveParams)) {
        if ($PSBoundParameters.ContainsKey('CandidateRunIds') -and $target.RunId -notin $CandidateRunIds) { continue }
        $targetPath = (ConvertTo-RSFullPath -Path $target.Path).TrimEnd('\')
        if (-not $targetPath.StartsWith("$runsRoot\", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to remove stale TestSandbox run outside '$runsRoot': $targetPath"
        }

        try {
            # Recorded test-owned disks are torn down (elevated) or retain the run before any walk.
            $refusals = @(Get-RSTestOwnedVirtualDiskRunRefusals -RunRoot $target.Path)
            if ($refusals.Count -eq 0) {
                $refusals = @(Remove-RSTestSandboxTreeWithoutFollowingLinks -Path $target.Path)
            }
            if ($refusals.Count -gt 0) {
                $message = "Retained the whole stale TestSandbox run '$($target.Path)' (nothing deleted) because it contains: $($refusals -join '; ')"
                Write-Warning $message
                $target | Add-Member -NotePropertyName Status -NotePropertyValue 'Retained' -Force
                $target | Add-Member -NotePropertyName Error -NotePropertyValue $message -Force
            } else {
                $target | Add-Member -NotePropertyName Status -NotePropertyValue 'Removed' -Force
                $target | Add-Member -NotePropertyName Error -NotePropertyValue $null -Force
            }
        } catch {
            $message = $_.Exception.Message
            Write-Warning "Failed to remove stale TestSandbox run '$($target.Path)': $message"
            $target | Add-Member -NotePropertyName Status -NotePropertyValue 'Failed' -Force
            $target | Add-Member -NotePropertyName Error -NotePropertyValue $message -Force
        }

        $results += $target
    }

    return $results
}

function New-RSTestSandboxDiskAuditIssue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    [pscustomobject]@{
        category = $Category
        path = $Path
        reason = $Reason
    }
}

function Get-RSTestSandboxDiskAudit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TestRoot,

        [string]$RunId = '',

        [string[]]$AllowedRunIds = @(),

        [int64[]]$LiveProcessIds = @(),

        [hashtable]$LiveProcessStartTimesUtc = @{}
    )

    # Absence is not cleanliness: validate authority even when the root/runs do not exist.
    $normalizedTestRoot = Assert-RSTestSandboxRootAuthorized -TestRoot $TestRoot -RequireMarker
    $allowedRunIds = @($AllowedRunIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (-not [string]::IsNullOrWhiteSpace($RunId)) {
        $allowedRunIds += $RunId
    }
    $allowedRunIds = @($allowedRunIds | Select-Object -Unique)

    $issues = @()
    if (Test-Path -LiteralPath $normalizedTestRoot) {
        $runsRoot = Join-Path $normalizedTestRoot 'runs'
        $resolveParams = @{
            TestRoot = $normalizedTestRoot
            RunId = $RunId
            AllowedRunIds = @($AllowedRunIds)
        }
        if ($PSBoundParameters.ContainsKey('LiveProcessIds')) {
            $resolveParams['LiveProcessIds'] = @($LiveProcessIds)
            $resolveParams['LiveProcessStartTimesUtc'] = $LiveProcessStartTimesUtc
        }
        $staleRunTargets = @{}
        foreach ($target in @(Resolve-RSTestSandboxStaleRunTargets @resolveParams)) {
            $staleRunTargets[$target.RunId] = $target
        }

        foreach ($item in @(Get-ChildItem -LiteralPath $normalizedTestRoot -Force -ErrorAction SilentlyContinue)) {
            if ($item.Name -notin @('runs', 'config', 'evidence', $script:RSTestSandboxMarkerFileName)) {
                $issues += New-RSTestSandboxDiskAuditIssue `
                    -Category 'unexpected-test-root-child' `
                    -Path $item.FullName `
                    -Reason 'Unified TestSandbox root should contain only runner-owned runs, config, and evidence.'
            }
        }

        if (Test-Path -LiteralPath $runsRoot) {
            foreach ($runDir in @(Get-ChildItem -LiteralPath $runsRoot -Directory -Force -ErrorAction SilentlyContinue)) {
                $isAllowedRun = @($allowedRunIds | Where-Object {
                        [string]::Equals($_, $runDir.Name, [System.StringComparison]::OrdinalIgnoreCase)
                    }).Count -gt 0
                if ($isAllowedRun) {
                    continue
                }

                if ($staleRunTargets.ContainsKey($runDir.Name)) {
                    $target = $staleRunTargets[$runDir.Name]
                    $issues += New-RSTestSandboxDiskAuditIssue `
                        -Category $target.Category `
                        -Path $runDir.FullName `
                        -Reason $target.Reason
                    continue
                }

                if ($null -eq (ConvertFrom-RSTestRunId -RunId $runDir.Name)) {
                    $issues += New-RSTestSandboxDiskAuditIssue `
                        -Category 'unexpected-test-run-dir' `
                        -Path $runDir.FullName `
                        -Reason 'The run directory is neither explicitly allowed nor attributable to a protected live owner.'
                }
            }
        }
    }

    [pscustomobject]@{
        schema = 'red-salamander.test-sandbox-disk-audit.v1'
        is_clean = (@($issues).Count -eq 0)
        issue_count = @($issues).Count
        test_root = $normalizedTestRoot
        run_id = $RunId
        allowed_run_ids = @($allowedRunIds)
        issues = @($issues)
    }
}

function Get-RSSelfTestArtifactRoot {
    param(
        [string]$RepoRoot = '',

        [string]$RunId = '',

        [string]$SelfTestRootOverride = $env:REDSALAMANDER_SELFTEST_ROOT,

        [string]$LocalAppDataRoot = $env:LOCALAPPDATA
    )

    # Retain the legacy parameters only for source compatibility. They are never
    # authorization to write outside the exact marked X:\RedSalamander.Perf root.
    [void]$SelfTestRootOverride
    [void]$LocalAppDataRoot
    $repository = Get-RSTestSandboxRepositoryRoot -RepoRoot $RepoRoot
    return (New-RSTestRunContext -RepoRoot $repository -RunId $RunId).ArtifactRoot
}

Export-ModuleMember -Function @(
    'Initialize-RSTestSandboxRoot',
    'Assert-RSTestSandboxContainedPath',
    'Get-RSTestSandboxRoot',
    'New-RSTestRunContext',
    'New-RSTestSandboxScratchDirectory',
    'Get-RSToolingPesterSandboxRoot',
    'Assert-RSTestSandboxLinkTarget',
    'New-RSTestSandboxLink',
    'Remove-RSTestSandboxExactRunDirectory',
    'Resolve-RSTestSandboxStaleRunTargets',
    'Remove-RSTestSandboxStaleRunDirectories',
    'Save-RSTestSandboxRunEvidence',
    'Save-RSTestSandboxStaleRunEvidence',
    'Get-RSTestSandboxDiskAudit',
    'Get-RSSelfTestArtifactRoot',
    'New-RSTestOwnedVirtualDiskMount',
    'Remove-RSTestOwnedVirtualDiskMount',
    'Get-RSTestOwnedVirtualDiskReceiptPaths'
)
