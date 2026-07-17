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
