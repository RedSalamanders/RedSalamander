# Private packaging module; import through its reviewed callers.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-RSReleaseVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReleaseVersion
    )

    if ($ReleaseVersion -notmatch '^\d+\.\d+\.\d+$') {
        throw "Release version must be a three-part numeric version: $ReleaseVersion"
    }
}

function Get-RSReleasePlatforms {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$BuildArm64
    )

    if ($BuildArm64) {
        return @('x64', 'ARM64')
    }
    return @('x64')
}

function Get-RSPortableReleaseFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReleaseVersion,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform
    )

    Assert-RSReleaseVersion -ReleaseVersion $ReleaseVersion
    return "RedSalamander-$ReleaseVersion-$Platform-Portable.zip"
}

function Get-RSMsixReleaseFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReleaseVersion,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform
    )

    Assert-RSReleaseVersion -ReleaseVersion $ReleaseVersion
    return "RedSalamander-$ReleaseVersion-$Platform.msix"
}

function Read-RSExactBytes {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Stream]$Stream,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 1048576)]
        [int]$Count
    )

    $bytes = [byte[]]::new($Count)
    $offset = 0
    while ($offset -lt $Count) {
        $read = $Stream.Read($bytes, $offset, $Count - $offset)
        if ($read -le 0) {
            throw "Unexpected end of stream while reading $Count bytes."
        }
        $offset += $read
    }

    return ,$bytes
}

function Get-RSPeMachineFromStream {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Stream]$Stream
    )

    $dosHeader = Read-RSExactBytes -Stream $Stream -Count 64
    if ($dosHeader[0] -ne 0x4D -or $dosHeader[1] -ne 0x5A) {
        throw 'Portable package RedSalamander.exe does not start with an MZ header.'
    }

    $peOffset = [BitConverter]::ToInt32($dosHeader, 0x3C)
    if ($peOffset -lt 64 -or $peOffset -gt 16MB) {
        throw "Portable package RedSalamander.exe has an invalid PE header offset: $peOffset"
    }

    $remaining = $peOffset - 64
    $skipBuffer = [byte[]]::new(4096)
    while ($remaining -gt 0) {
        $requested = [Math]::Min($remaining, $skipBuffer.Length)
        $read = $Stream.Read($skipBuffer, 0, $requested)
        if ($read -le 0) {
            throw 'Portable package RedSalamander.exe ended before its PE header.'
        }
        $remaining -= $read
    }

    $peHeader = Read-RSExactBytes -Stream $Stream -Count 6
    if ($peHeader[0] -ne 0x50 -or $peHeader[1] -ne 0x45 -or $peHeader[2] -ne 0 -or $peHeader[3] -ne 0) {
        throw 'Portable package RedSalamander.exe has an invalid PE signature.'
    }

    return [BitConverter]::ToUInt16($peHeader, 4)
}

function Get-RSPortableArchiveMachine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        $entries = @($archive.Entries | Where-Object {
                [string]::Equals($_.FullName, 'RedSalamander.exe', [StringComparison]::OrdinalIgnoreCase)
            })
        if ($entries.Count -ne 1) {
            throw "Portable archive must contain exactly one root RedSalamander.exe: $ArchivePath"
        }

        $stream = $entries[0].Open()
        try {
            return Get-RSPeMachineFromStream -Stream $stream
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $archive.Dispose()
    }
}

# Distributable binaries must carry no ENABLE_TESTS surface (Specs/Build/Build_Toolchain.md,
# "Distributable test surface"). Test-only native exports use these reserved, case-sensitive
# name forms, and Common/VersionInfo.rc stamps every test-enabled first-party module with the
# marker below in its version resource.
$script:RSTestSurfaceExportPattern = 'SelfTest|ForTest|^RedSalamander\w*(Test|Debug)'
$script:RSTestSurfaceVersionMarker = 'RedSalamander.TestSurface=ENABLE_TESTS'

function Read-RSPeUInt16 {
    param([byte[]]$Bytes, [long]$Offset, [string]$Name)

    if ($Offset -lt 0 -or $Offset + 2 -gt $Bytes.LongLength) {
        throw "PE image '$Name' is truncated at offset $Offset."
    }
    return [BitConverter]::ToUInt16($Bytes, [int]$Offset)
}

function Read-RSPeUInt32 {
    param([byte[]]$Bytes, [long]$Offset, [string]$Name)

    if ($Offset -lt 0 -or $Offset + 4 -gt $Bytes.LongLength) {
        throw "PE image '$Name' is truncated at offset $Offset."
    }
    return [BitConverter]::ToUInt32($Bytes, [int]$Offset)
}

function ConvertTo-RSPeFileOffset {
    param([object[]]$Sections, [long]$Rva, [string]$Name)

    foreach ($section in $Sections) {
        if ($Rva -ge $section.VirtualAddress -and $Rva -lt $section.VirtualAddress + $section.Span) {
            return $section.PointerToRawData + ($Rva - $section.VirtualAddress)
        }
    }
    throw "PE image '$Name' references RVA 0x$('{0:X}' -f $Rva) outside every section."
}

function Get-RSPeTestSurface {
    <#
    .SYNOPSIS
        Reports the ENABLE_TESTS surface of one PE image held in memory.
    .DESCRIPTION
        Parses the PE32/PE32+ export directory and the resource section without loading the
        image. Returns the version-resource test marker state and every export whose name uses a
        reserved test-only form. A malformed image throws so callers fail closed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($Bytes.LongLength -lt 64 -or $Bytes[0] -ne 0x4D -or $Bytes[1] -ne 0x5A) {
        throw "PE image '$Name' does not start with an MZ header."
    }
    $peOffset = [long](Read-RSPeUInt32 -Bytes $Bytes -Offset 0x3C -Name $Name)
    if ((Read-RSPeUInt32 -Bytes $Bytes -Offset $peOffset -Name $Name) -ne 0x00004550) {
        throw "PE image '$Name' has an invalid PE signature."
    }
    $sectionCount = Read-RSPeUInt16 -Bytes $Bytes -Offset ($peOffset + 6) -Name $Name
    $optionalHeaderSize = Read-RSPeUInt16 -Bytes $Bytes -Offset ($peOffset + 20) -Name $Name
    $optionalHeader = $peOffset + 24
    $magic = Read-RSPeUInt16 -Bytes $Bytes -Offset $optionalHeader -Name $Name
    $directoryCountOffset = switch ($magic) {
        0x10B { $optionalHeader + 92 }
        0x20B { $optionalHeader + 108 }
        default { throw "PE image '$Name' has unsupported optional-header magic 0x$('{0:X4}' -f $magic)." }
    }
    $directoryCount = Read-RSPeUInt32 -Bytes $Bytes -Offset $directoryCountOffset -Name $Name
    $directories = $directoryCountOffset + 4

    $sections = [System.Collections.Generic.List[object]]::new()
    $sectionTable = $optionalHeader + $optionalHeaderSize
    for ($index = 0; $index -lt $sectionCount; ++$index) {
        $header = $sectionTable + (40 * $index)
        $virtualSize = Read-RSPeUInt32 -Bytes $Bytes -Offset ($header + 8) -Name $Name
        $rawSize = Read-RSPeUInt32 -Bytes $Bytes -Offset ($header + 16) -Name $Name
        $sections.Add([pscustomobject]@{
                VirtualAddress = [long](Read-RSPeUInt32 -Bytes $Bytes -Offset ($header + 12) -Name $Name)
                Span = [long][Math]::Max($virtualSize, $rawSize)
                PointerToRawData = [long](Read-RSPeUInt32 -Bytes $Bytes -Offset ($header + 20) -Name $Name)
                SizeOfRawData = [long]$rawSize
            })
    }

    $testExports = [System.Collections.Generic.List[string]]::new()
    $exportRva = if ($directoryCount -gt 0) { Read-RSPeUInt32 -Bytes $Bytes -Offset $directories -Name $Name } else { 0 }
    if ($exportRva -ne 0) {
        $exportDirectory = ConvertTo-RSPeFileOffset -Sections $sections -Rva $exportRva -Name $Name
        $nameCount = Read-RSPeUInt32 -Bytes $Bytes -Offset ($exportDirectory + 24) -Name $Name
        $nameTableRva = Read-RSPeUInt32 -Bytes $Bytes -Offset ($exportDirectory + 32) -Name $Name
        if ($nameCount -gt 0) {
            $nameTable = ConvertTo-RSPeFileOffset -Sections $sections -Rva $nameTableRva -Name $Name
            for ($index = 0; $index -lt $nameCount; ++$index) {
                $nameRva = Read-RSPeUInt32 -Bytes $Bytes -Offset ($nameTable + (4 * $index)) -Name $Name
                $nameOffset = ConvertTo-RSPeFileOffset -Sections $sections -Rva $nameRva -Name $Name
                if ($nameOffset -ge $Bytes.LongLength) {
                    throw "PE image '$Name' has an export name outside the file."
                }
                $terminator = [Array]::IndexOf($Bytes, [byte]0, [int]$nameOffset)
                if ($terminator -lt 0) {
                    throw "PE image '$Name' has an unterminated export name."
                }
                $exportName = [Text.Encoding]::ASCII.GetString($Bytes, [int]$nameOffset, $terminator - [int]$nameOffset)
                if ($exportName -cmatch $script:RSTestSurfaceExportPattern) {
                    $testExports.Add($exportName)
                }
            }
        }
    }

    # Version-resource strings are UTF-16 and 2-byte aligned within the section, whose raw data
    # starts on a FileAlignment boundary, so decoding the whole section finds the marker.
    $testBuildMarker = $false
    $resourceRva = if ($directoryCount -gt 2) { Read-RSPeUInt32 -Bytes $Bytes -Offset ($directories + 16) -Name $Name } else { 0 }
    if ($resourceRva -ne 0) {
        $resourceSection = @($sections | Where-Object {
                $resourceRva -ge $_.VirtualAddress -and $resourceRva -lt $_.VirtualAddress + $_.Span
            } | Select-Object -First 1)
        if ($resourceSection.Count -ne 1) {
            throw "PE image '$Name' references a resource directory outside every section."
        }
        $start = $resourceSection[0].PointerToRawData
        $length = [Math]::Min($resourceSection[0].SizeOfRawData, $Bytes.LongLength - $start)
        if ($start -lt 0 -or $length -lt 0) {
            throw "PE image '$Name' has a resource section outside the file."
        }
        $resourceText = [Text.Encoding]::Unicode.GetString($Bytes, [int]$start, [int]($length - ($length % 2)))
        $testBuildMarker = $resourceText.Contains($script:RSTestSurfaceVersionMarker, [StringComparison]::Ordinal)
    }

    return [pscustomobject]@{
        Name = $Name
        TestBuildMarker = $testBuildMarker
        TestExports = @($testExports)
    }
}

function Test-RSShippingBinaryName {
    param([Parameter(Mandatory = $true)][string]$Name)

    return $Name.EndsWith('.exe', [StringComparison]::OrdinalIgnoreCase) -or
        $Name.EndsWith('.dll', [StringComparison]::OrdinalIgnoreCase)
}

function Find-RSShippingTestSurface {
    <#
    .SYNOPSIS
        Lists the ENABLE_TESTS surface in a staged package directory, portable ZIP, or MSIX.
    .DESCRIPTION
        Inspects every .exe and .dll beneath a directory or inside a ZIP-format package and
        returns one finding per binary that carries the test-enabled version marker or a
        reserved test-only export. Unreadable or malformed binaries throw.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath
    )

    $findings = [System.Collections.Generic.List[object]]::new()
    $inspected = 0
    $addFinding = {
        param($surface)
        if ($surface.TestBuildMarker -or $surface.TestExports.Count -gt 0) {
            $findings.Add($surface)
        }
    }

    if (Test-Path -LiteralPath $LiteralPath -PathType Container) {
        $root = [IO.Path]::GetFullPath($LiteralPath).TrimEnd('\') + '\'
        foreach ($file in @(Get-ChildItem -LiteralPath $LiteralPath -File -Recurse | Sort-Object FullName)) {
            if (-not (Test-RSShippingBinaryName -Name $file.Name)) { continue }
            $relative = $file.FullName.Substring($root.Length)
            & $addFinding (Get-RSPeTestSurface -Bytes ([IO.File]::ReadAllBytes($file.FullName)) -Name $relative)
            ++$inspected
        }
    }
    elseif (Test-Path -LiteralPath $LiteralPath -PathType Leaf) {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $archive = [System.IO.Compression.ZipFile]::OpenRead($LiteralPath)
        try {
            foreach ($entry in @($archive.Entries | Sort-Object FullName)) {
                if (-not (Test-RSShippingBinaryName -Name $entry.FullName)) { continue }
                $stream = $entry.Open()
                $buffer = [IO.MemoryStream]::new()
                try {
                    $stream.CopyTo($buffer)
                }
                finally {
                    $stream.Dispose()
                }
                & $addFinding (Get-RSPeTestSurface -Bytes $buffer.ToArray() -Name $entry.FullName)
                $buffer.Dispose()
                ++$inspected
            }
        }
        finally {
            $archive.Dispose()
        }
    }
    else {
        throw "Shipping test-surface input not found: $LiteralPath"
    }

    if ($inspected -eq 0) {
        throw "Shipping test-surface input contains no .exe or .dll to inspect: $LiteralPath"
    }
    return @($findings)
}

function Assert-RSShippingPackageHasNoTestSurface {
    <#
    .SYNOPSIS
        Fails when a distributable package or its staged contents carries an ENABLE_TESTS surface.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath
    )

    $findings = @(Find-RSShippingTestSurface -LiteralPath $LiteralPath)
    if ($findings.Count -eq 0) {
        return
    }

    $details = @($findings | ForEach-Object {
            $reasons = @()
            if ($_.TestBuildMarker) { $reasons += 'built with ENABLE_TESTS' }
            if ($_.TestExports.Count -gt 0) { $reasons += "test exports [$($_.TestExports -join ', ')]" }
            "$($_.Name): $($reasons -join '; ')"
        })
    throw ("Distributable package carries a test surface and must be rebuilt with RSBuildEnableTests=false: {0}`n  {1}" -f
        $LiteralPath, ($details -join "`n  "))
}

function Assert-RSProducedMsixHasNoTestSurface {
    <#
    .SYNOPSIS
        Checks an MSIX that packaging just wrote beneath AppPackages and removes it when rejected.
    .DESCRIPTION
        MSBuild publishes the package at its normal output path before it can be inspected, so a
        rejected package is deleted together with its package folder (the packaging project writes
        each package into its own RedSalamanderInstaller_<version>_<platform>_Test folder) before
        the rejection is rethrown. A package directly in AppPackages is removed alone. The removal
        never follows a link: a reparse point at the AppPackages root, on the way to the package or
        anywhere beneath its folder refuses it, and the rejection then says the package must not be
        distributed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$AppPackagesRoot
    )

    $packagePath = [IO.Path]::GetFullPath($LiteralPath)
    $root = [IO.Path]::GetFullPath($AppPackagesRoot).TrimEnd([char[]]@('\', '/')) + '\'
    if (-not $packagePath.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Produced MSIX is not beneath the AppPackages root '$root': $packagePath"
    }

    try {
        Assert-RSShippingPackageHasNoTestSurface -LiteralPath $packagePath
    }
    catch {
        $rejection = $_.Exception.Message
        $packageFolder = [IO.Path]::GetDirectoryName($packagePath).TrimEnd([char[]]@('\', '/')) + '\'
        $removalTarget = if ($packageFolder.Length -gt $root.Length) { $packageFolder.TrimEnd('\') } else { $packagePath }
        try {
            if (Test-Path -LiteralPath $removalTarget) {
                Remove-RSRejectedMsixWithoutFollowingLinks -AppPackagesRoot $root -Target $removalTarget
            }
        }
        catch {
            throw "$rejection`nThe rejected package could not be removed and must not be distributed: $removalTarget ($($_.Exception.Message))"
        }
        throw "$rejection`nThe rejected package was removed: $removalTarget"
    }
}

function Get-RSRejectedMsixTreeEntries {
    param([Parameter(Mandatory = $true)][string]$Directory)

    # Default EnumerationOptions skip Hidden/System entries; a removal walk must see them.
    $options = [IO.EnumerationOptions]::new()
    $options.AttributesToSkip = [IO.FileAttributes]0
    $options.RecurseSubdirectories = $false
    $options.IgnoreInaccessible = $false
    return @([IO.DirectoryInfo]::new($Directory).GetFileSystemInfos('*', $options))
}

function Remove-RSRejectedMsixWithoutFollowingLinks {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppPackagesRoot,

        [Parameter(Mandatory = $true)]
        [string]$Target
    )

    # Mirrors TestSandbox.psm1's Remove-RSTestSandboxTreeWithoutFollowingLinks; kept local because
    # this production packaging module must not import test tooling.
    # Never delete through a link: the AppPackages root, every component down to the target and
    # every entry beneath it must be a plain file or directory. Any reparse point refuses the
    # whole removal before anything is deleted.
    $root = $AppPackagesRoot.TrimEnd([char[]]@('\', '/'))
    $components = [Collections.Generic.List[string]]::new()
    $components.Add($root)
    $current = $root
    foreach ($segment in @($Target.Substring($root.Length).Split([char[]]@('\', '/'), [StringSplitOptions]::RemoveEmptyEntries))) {
        $current = Join-Path $current $segment
        $components.Add($current)
    }
    foreach ($component in $components) {
        if (([IO.File]::GetAttributes($component) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "refusing to delete through the reparse point '$component'"
        }
    }

    $targetAttributes = [IO.File]::GetAttributes($Target)
    if (($targetAttributes -band [IO.FileAttributes]::Directory) -eq 0) {
        if (($targetAttributes -band [IO.FileAttributes]::ReadOnly) -ne 0) {
            [IO.File]::SetAttributes($Target, $targetAttributes -band (-bnot [IO.FileAttributes]::ReadOnly))
        }
        [IO.File]::Delete($Target)
        return
    }

    $pending = [Collections.Generic.Stack[string]]::new()
    $pending.Push($Target)
    while ($pending.Count -gt 0) {
        foreach ($entry in (Get-RSRejectedMsixTreeEntries -Directory $pending.Pop())) {
            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "refusing to delete a package folder that contains the reparse point '$($entry.FullName)'"
            }
            if (($entry.Attributes -band [IO.FileAttributes]::Directory) -ne 0) {
                $pending.Push($entry.FullName)
            }
        }
    }

    # Delete post-order with non-recursive deletes. Each directory is re-read before it is
    # enumerated, so one replaced by a link since the scan fails the removal instead of being entered.
    $directories = [Collections.Generic.List[string]]::new()
    $pending.Push($Target)
    while ($pending.Count -gt 0) {
        $directory = $pending.Pop()
        if (([IO.File]::GetAttributes($directory) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "refusing to enumerate a package directory that became a reparse point: '$directory'"
        }
        $directories.Add($directory)
        foreach ($entry in (Get-RSRejectedMsixTreeEntries -Directory $directory)) {
            $attributes = $entry.Attributes
            if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "refusing to delete a package folder that contains the reparse point '$($entry.FullName)'"
            }
            if (($attributes -band [IO.FileAttributes]::ReadOnly) -ne 0) {
                $entry.Attributes = $attributes -band (-bnot [IO.FileAttributes]::ReadOnly)
            }
            if (($attributes -band [IO.FileAttributes]::Directory) -ne 0) {
                $pending.Push($entry.FullName)
            }
            else {
                [IO.File]::Delete($entry.FullName)
            }
        }
    }
    for ($index = $directories.Count - 1; $index -ge 0; --$index) {
        [IO.Directory]::Delete($directories[$index], $false)
    }
}

function Get-RSMsixOutputSnapshot {
    <#
    .SYNOPSIS
        Records each MSIX beneath AppPackages as Length|LastWriteTimeUtc before packaging runs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppPackagesRoot
    )

    $snapshot = @{}
    if (Test-Path -LiteralPath $AppPackagesRoot -PathType Container) {
        foreach ($existing in @(Get-ChildItem -LiteralPath $AppPackagesRoot -Filter *.msix -File -Recurse)) {
            $snapshot[$existing.FullName] = '{0}|{1}' -f $existing.Length, $existing.LastWriteTimeUtc.Ticks
        }
    }
    return $snapshot
}

function Assert-RSMsixPackagingOutput {
    <#
    .SYNOPSIS
        Checks the MSIX one packaging run produced, or the up-to-date package MSBuild left in place.
    .DESCRIPTION
        Every MSIX that is new or changed since the snapshot is checked with
        Assert-RSProducedMsixHasNoTestSurface (a rejected one is removed). When MSBuild found the
        package up to date and wrote nothing (a re-run with the same version), the existing package
        whose identity matches this version and platform is checked without being removed, and the
        run succeeds when it is clean, as build.ps1 does for an unchanged package. Only when no such
        package exists does the run fail for producing no MSIX.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppPackagesRoot,

        [Parameter(Mandatory = $true)]
        [hashtable]$Before,

        [Parameter(Mandatory = $true)]
        [string]$ReleaseVersion,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform
    )

    Assert-RSReleaseVersion -ReleaseVersion $ReleaseVersion
    $current = Get-RSMsixOutputSnapshot -AppPackagesRoot $AppPackagesRoot
    $produced = @($current.Keys | Where-Object {
            -not $Before.ContainsKey($_) -or $Before[$_] -ne $current[$_]
        } | Sort-Object)
    foreach ($msix in $produced) {
        Assert-RSProducedMsixHasNoTestSurface -LiteralPath $msix -AppPackagesRoot $AppPackagesRoot
    }
    if ($produced.Count -gt 0) {
        return
    }

    $expectedVersion = "$ReleaseVersion.0"
    $expectedArchitecture = if ($Platform -eq 'ARM64') { 'arm64' } else { 'x64' }
    $upToDate = @($current.Keys | Sort-Object | Where-Object {
            $identity = $null
            try {
                $identity = Get-RSMsixIdentity -PackagePath $_
            }
            catch {
                # An unreadable package is not this run's package; it is never selected.
            }
            $null -ne $identity -and
            $identity.Name -ceq 'RedSalamander' -and
            $identity.Version -ceq $expectedVersion -and
            $identity.ProcessorArchitecture -ieq $expectedArchitecture
        })
    if ($upToDate.Count -eq 0) {
        throw "MSIX build reported success but produced no new MSIX, and no up-to-date $expectedVersion $Platform MSIX exists beneath $AppPackagesRoot"
    }
    foreach ($msix in $upToDate) {
        Assert-RSShippingPackageHasNoTestSurface -LiteralPath $msix
    }
}

function Get-RSMsixIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackagePath
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
    try {
        $entries = @($archive.Entries | Where-Object {
                [string]::Equals($_.FullName, 'AppxManifest.xml', [StringComparison]::OrdinalIgnoreCase)
            })
        if ($entries.Count -ne 1) {
            throw "MSIX must contain exactly one root AppxManifest.xml: $PackagePath"
        }

        $stream = $entries[0].Open()
        $reader = [System.IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true)
        try {
            $manifest = [xml]$reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
            $stream.Dispose()
        }

        $identity = $manifest.Package.Identity
        if (-not $identity) {
            throw "MSIX AppxManifest.xml has no package Identity: $PackagePath"
        }

        return [pscustomobject]@{
            Name = [string]$identity.Name
            Publisher = [string]$identity.Publisher
            Version = [string]$identity.Version
            ProcessorArchitecture = [string]$identity.ProcessorArchitecture
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Assert-RSReleaseArtifacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArtifactsRoot,

        [Parameter(Mandatory = $true)]
        [string]$ReleaseVersion,

        [Parameter(Mandatory = $true)]
        [bool]$BuildArm64,

        [Parameter(Mandatory = $true)]
        [bool]$BuildMsix
    )

    Assert-RSReleaseVersion -ReleaseVersion $ReleaseVersion
    if (-not (Test-Path -LiteralPath $ArtifactsRoot -PathType Container)) {
        throw "Release artifacts directory not found: $ArtifactsRoot"
    }

    $expected = [System.Collections.Generic.List[object]]::new()
    foreach ($platform in Get-RSReleasePlatforms -BuildArm64 $BuildArm64) {
        $expected.Add([pscustomobject]@{
                Kind = 'Portable'
                Platform = $platform
                Name = Get-RSPortableReleaseFileName -ReleaseVersion $ReleaseVersion -Platform $platform
            })
        if ($BuildMsix) {
            $expected.Add([pscustomobject]@{
                    Kind = 'MSIX'
                    Platform = $platform
                    Name = Get-RSMsixReleaseFileName -ReleaseVersion $ReleaseVersion -Platform $platform
                })
        }
    }

    $actual = @(Get-ChildItem -LiteralPath $ArtifactsRoot -File -Recurse | Where-Object {
            $_.Extension -ieq '.zip' -or $_.Extension -ieq '.msix'
        })
    if ($actual.Count -ne $expected.Count) {
        $actualNames = @($actual | ForEach-Object { $_.Name } | Sort-Object) -join ', '
        $expectedNames = @($expected | ForEach-Object { $_.Name } | Sort-Object) -join ', '
        throw "Release artifact count mismatch. Expected [$expectedNames]; found [$actualNames]."
    }

    $validated = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $expected) {
        $matches = @($actual | Where-Object { $_.Name -ceq $item.Name })
        if ($matches.Count -ne 1) {
            throw "Expected exactly one release artifact named '$($item.Name)'; found $($matches.Count)."
        }

        $file = $matches[0]
        if ($file.Length -le 0) {
            throw "Release artifact is empty: $($item.Name)"
        }

        if ($item.Kind -eq 'Portable') {
            $expectedMachine = if ($item.Platform -eq 'ARM64') { 0xAA64 } else { 0x8664 }
            $actualMachine = Get-RSPortableArchiveMachine -ArchivePath $file.FullName
            if ($actualMachine -ne $expectedMachine) {
                throw ("Portable artifact '{0}' has PE machine 0x{1:X4}; expected 0x{2:X4} for {3}." -f
                    $item.Name, $actualMachine, $expectedMachine, $item.Platform)
            }
        }
        else {
            $identity = Get-RSMsixIdentity -PackagePath $file.FullName
            $expectedArchitecture = if ($item.Platform -eq 'ARM64') { 'arm64' } else { 'x64' }
            if ($identity.Name -cne 'RedSalamander') {
                throw "MSIX '$($item.Name)' has unexpected Identity Name '$($identity.Name)'."
            }
            if ($identity.Publisher -cne 'CN=RedSalmanders') {
                throw "MSIX '$($item.Name)' has unexpected Identity Publisher '$($identity.Publisher)'."
            }
            if ($identity.Version -cne "$ReleaseVersion.0") {
                throw "MSIX '$($item.Name)' has Identity Version '$($identity.Version)'; expected '$ReleaseVersion.0'."
            }
            if ($identity.ProcessorArchitecture -cne $expectedArchitecture) {
                throw "MSIX '$($item.Name)' has architecture '$($identity.ProcessorArchitecture)'; expected '$expectedArchitecture'."
            }
        }

        Assert-RSShippingPackageHasNoTestSurface -LiteralPath $file.FullName

        $validated.Add([pscustomobject]@{
                Name = $item.Name
                FullName = $file.FullName
                Length = [long]$file.Length
                Kind = $item.Kind
                Platform = $item.Platform
            })
    }

    return @($validated)
}

function New-RSReleaseChecksumFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArtifactsRoot,

        [Parameter(Mandatory = $true)]
        [string]$ReleaseVersion,

        [Parameter(Mandatory = $true)]
        [bool]$BuildArm64,

        [Parameter(Mandatory = $true)]
        [bool]$BuildMsix,

        [string]$OutputPath = (Join-Path $ArtifactsRoot 'checksums.sha256')
    )

    $artifacts = @(Assert-RSReleaseArtifacts `
            -ArtifactsRoot $ArtifactsRoot `
            -ReleaseVersion $ReleaseVersion `
            -BuildArm64 $BuildArm64 `
            -BuildMsix $BuildMsix)
    $lines = @($artifacts | Sort-Object Name | ForEach-Object {
            $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            "$hash  $($_.Name)"
        })

    $outputDirectory = Split-Path -Parent ([IO.Path]::GetFullPath($OutputPath))
    [IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
    [IO.File]::WriteAllLines($OutputPath, $lines, [Text.UTF8Encoding]::new($false))
    return Get-Item -LiteralPath $OutputPath
}

function Assert-RSReleaseChecksumFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArtifactsRoot,

        [Parameter(Mandatory = $true)]
        [string]$ReleaseVersion,

        [Parameter(Mandatory = $true)]
        [bool]$BuildArm64,

        [Parameter(Mandatory = $true)]
        [bool]$BuildMsix,

        [string]$ChecksumPath = (Join-Path $ArtifactsRoot 'checksums.sha256')
    )

    if (-not (Test-Path -LiteralPath $ChecksumPath -PathType Leaf)) {
        throw "Release checksum file not found: $ChecksumPath"
    }

    $artifacts = @(Assert-RSReleaseArtifacts `
            -ArtifactsRoot $ArtifactsRoot `
            -ReleaseVersion $ReleaseVersion `
            -BuildArm64 $BuildArm64 `
            -BuildMsix $BuildMsix)
    $expectedLines = @($artifacts | Sort-Object Name | ForEach-Object {
            $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            "$hash  $($_.Name)"
        })
    $actualLines = @(Get-Content -LiteralPath $ChecksumPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if ($actualLines.Count -ne $expectedLines.Count) {
        throw "Release checksum entry count mismatch. Expected $($expectedLines.Count); found $($actualLines.Count)."
    }
    for ($index = 0; $index -lt $expectedLines.Count; ++$index) {
        if ($actualLines[$index] -cne $expectedLines[$index]) {
            throw "Release checksum mismatch at line $($index + 1). Expected '$($expectedLines[$index])'; found '$($actualLines[$index])'."
        }
    }

    return $true
}

Export-ModuleMember -Function @(
    'Get-RSPortableReleaseFileName',
    'Get-RSMsixReleaseFileName',
    'Assert-RSReleaseArtifacts',
    'Get-RSPeTestSurface',
    'Find-RSShippingTestSurface',
    'Assert-RSShippingPackageHasNoTestSurface',
    'Assert-RSProducedMsixHasNoTestSurface',
    'Get-RSMsixOutputSnapshot',
    'Assert-RSMsixPackagingOutput',
    'New-RSReleaseChecksumFile',
    'Assert-RSReleaseChecksumFile'
)
