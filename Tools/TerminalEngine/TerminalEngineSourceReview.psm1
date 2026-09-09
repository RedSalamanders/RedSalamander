Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-RSSha256Bytes {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [ValidateRange(0, [int]::MaxValue)]
        [int]$Offset = 0,

        [int]$Count = -1
    )

    if ($Count -lt 0) {
        $Count = $Bytes.Length - $Offset
    }
    if ($Offset -gt $Bytes.Length -or
        $Count -lt 0 -or
        $Count -gt ($Bytes.Length - $Offset)) {
        throw [IO.InvalidDataException]::new('SHA-256 byte range is invalid.')
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToHexString(
            $sha.ComputeHash($Bytes, $Offset, $Count)).ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-RSSourceLineRangeIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [ValidateRange(1, 10000000)]
        [int]$LineStart,

        [Parameter(Mandatory)]
        [ValidateRange(1, 10000000)]
        [int]$LineEnd
    )

    if ($LineEnd -lt $LineStart) {
        throw [IO.InvalidDataException]::new('Source line range is reversed.')
    }
    if ($Bytes.Length -eq 0) {
        throw [IO.InvalidDataException]::new('An empty source member has no reviewable line.')
    }

    $line = 1
    $startOffset = if ($LineStart -eq 1) { 0 } else { -1 }
    $endOffset = -1
    for ($index = 0; $index -lt $Bytes.Length; ++$index) {
        if ($Bytes[$index] -ne 0x0a) {
            continue
        }
        if ($line -eq $LineEnd) {
            $endOffset = $index + 1
            break
        }
        $line++
        if ($line -eq $LineStart) {
            $startOffset = $index + 1
        }
    }
    if ($startOffset -lt 0) {
        throw [IO.InvalidDataException]::new(
            "Source lineStart $LineStart is beyond the archive member.")
    }
    if ($endOffset -lt 0) {
        $lastLine = $line
        if ($Bytes[$Bytes.Length - 1] -eq 0x0a) {
            $lastLine++
        }
        if ($LineEnd -gt $lastLine) {
            throw [IO.InvalidDataException]::new(
                "Source lineEnd $LineEnd is beyond the archive member.")
        }
        $endOffset = $Bytes.Length
    }
    if ($endOffset -lt $startOffset) {
        throw [IO.InvalidDataException]::new('Source line range offsets are invalid.')
    }

    return [pscustomobject]@{
        offset = $startOffset
        sizeBytes = ($endOffset - $startOffset)
        sha256 = Get-RSSha256Bytes `
            -Bytes $Bytes `
            -Offset $startOffset `
            -Count ($endOffset - $startOffset)
    }
}

function Read-RSSourceArchiveMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ArchivePath,

        [Parameter(Mandatory)]
        [string]$RelativePath
    )

    if (-not [IO.File]::Exists($ArchivePath)) {
        throw [IO.FileNotFoundException]::new(
            'Source archive is missing.',
            $ArchivePath)
    }
    $normalizedRelative = $RelativePath.Replace('\', '/')
    if ([string]::IsNullOrWhiteSpace($normalizedRelative) -or
        $normalizedRelative.StartsWith('/', [StringComparison]::Ordinal) -or
        $normalizedRelative -match '^[A-Za-z]:' -or
        @($normalizedRelative.Split('/') | Where-Object {
                $_ -eq '..' -or $_ -eq '.'
            }).Count -ne 0) {
        throw [IO.InvalidDataException]::new(
            "Unsafe source-review relative path '$RelativePath'.")
    }

    $archiveStream = [IO.File]::Open(
        $ArchivePath,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::Read)
    try {
        $gzip = [IO.Compression.GZipStream]::new(
            $archiveStream,
            [IO.Compression.CompressionMode]::Decompress,
            $true)
        try {
            $reader = [System.Formats.Tar.TarReader]::new($gzip, $true)
            try {
                $found = $null
                while ($null -ne ($entry = $reader.GetNextEntry())) {
                    $entryName = ([string]$entry.Name).Replace('\', '/')
                    if ($entryName.StartsWith('/', [StringComparison]::Ordinal) -or
                        $entryName -match '^[A-Za-z]:' -or
                        @($entryName.Split('/') | Where-Object {
                                $_ -eq '..'
                            }).Count -ne 0) {
                        throw [IO.InvalidDataException]::new(
                            "Source archive contains unsafe member '$entryName'.")
                    }
                    $matches = $entryName -ceq $normalizedRelative -or
                        $entryName.EndsWith(
                            "/$normalizedRelative",
                            [StringComparison]::Ordinal)
                    if (-not $matches) {
                        continue
                    }
                    if ($null -ne $found) {
                        throw [IO.InvalidDataException]::new(
                            "Source archive contains multiple matches for '$normalizedRelative'.")
                    }
                    if ($entry.EntryType -cne
                        [System.Formats.Tar.TarEntryType]::RegularFile -and
                        $entry.EntryType -cne
                        [System.Formats.Tar.TarEntryType]::V7RegularFile) {
                        throw [IO.InvalidDataException]::new(
                            "Source-review member '$entryName' is not a regular file.")
                    }
                    $memory = [IO.MemoryStream]::new()
                    try {
                        $entry.DataStream.CopyTo($memory)
                        $found = $memory.ToArray()
                    }
                    finally {
                        $memory.Dispose()
                    }
                }
                if ($null -eq $found) {
                    throw [IO.InvalidDataException]::new(
                        "Source archive lacks '$normalizedRelative'.")
                }
                return [byte[]]$found
            }
            finally {
                $reader.Dispose()
            }
        }
        finally {
            $gzip.Dispose()
        }
    }
    finally {
        $archiveStream.Dispose()
    }
}

function Test-RSSourceReviewBlocker {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ArchivePath,

        [Parameter(Mandatory)]
        [string]$ExpectedArchiveSha256,

        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$Blocker
    )

    $archiveSha256 = (Get-FileHash `
            -LiteralPath $ArchivePath `
            -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($archiveSha256 -cne $ExpectedArchiveSha256) {
        throw [IO.InvalidDataException]::new(
            'Source-review archive SHA-256 differs from the frozen input.')
    }
    $source = $Blocker.source
    $bytes = Read-RSSourceArchiveMember `
        -ArchivePath $ArchivePath `
        -RelativePath ([string]$source.relativePath)
    $fileSha256 = Get-RSSha256Bytes -Bytes $bytes
    if ($fileSha256 -cne [string]$source.fileSha256) {
        throw [IO.InvalidDataException]::new(
            "Source-review file SHA-256 differs for '$($source.relativePath)'.")
    }
    $range = Get-RSSourceLineRangeIdentity `
        -Bytes $bytes `
        -LineStart ([int]$source.lineStart) `
        -LineEnd ([int]$source.lineEnd)
    if ([string]$range.sha256 -cne [string]$source.lineRangeSha256) {
        throw [IO.InvalidDataException]::new(
            "Source-review line-range SHA-256 differs for '$($source.relativePath)'.")
    }
    return [ordered]@{
        archiveSha256 = $archiveSha256
        relativePath = ([string]$source.relativePath).Replace('\', '/')
        fileSha256 = $fileSha256
        lineStart = [int]$source.lineStart
        lineEnd = [int]$source.lineEnd
        lineRangeSha256 = [string]$range.sha256
        lineRangeSizeBytes = ([uint64]$range.sizeBytes).ToString(
            [Globalization.CultureInfo]::InvariantCulture)
    }
}

Export-ModuleMember -Function `
    Get-RSSourceLineRangeIdentity, `
    Read-RSSourceArchiveMember, `
    Test-RSSourceReviewBlocker
