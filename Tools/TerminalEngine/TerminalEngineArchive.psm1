Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RSArchiveMaxRawBytes = [uint64]536870912
$script:RSArchiveMaxEntries = 100000
$script:RSArchiveMaxEntryBytes = [uint64]268435456
$script:RSArchiveMaxExpandedBytes = [uint64]2147483648
$script:RSArchiveMaxExpansionRatio = [uint64]200
$script:RSArchiveMaxDepth = 32
$script:RSArchiveMaxPathLength = 1024

function Get-RSArchiveSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath
    )

    return (Get-FileHash -LiteralPath $LiteralPath -Algorithm SHA256).
        Hash.ToLowerInvariant()
}

function Get-RSArchiveUInt16 {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes,

        [Parameter(Mandatory = $true)]
        [int]$Offset
    )

    if ($Offset -lt 0 -or $Offset + 2 -gt $Bytes.Length) {
        throw [IO.InvalidDataException]::new(
            'ZIP metadata extends outside the archive.')
    }
    return [uint16](
        [uint16]$Bytes[$Offset] -bor
        ([uint16]$Bytes[$Offset + 1] -shl 8))
}

function Get-RSArchiveUInt32 {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes,

        [Parameter(Mandatory = $true)]
        [int]$Offset
    )

    if ($Offset -lt 0 -or $Offset + 4 -gt $Bytes.Length) {
        throw [IO.InvalidDataException]::new(
            'ZIP metadata extends outside the archive.')
    }
    return [uint32](
        [uint32]$Bytes[$Offset] -bor
        ([uint32]$Bytes[$Offset + 1] -shl 8) -bor
        ([uint32]$Bytes[$Offset + 2] -shl 16) -bor
        ([uint32]$Bytes[$Offset + 3] -shl 24))
}

function Assert-RSArchivePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Value) -or
        $Value.Length -gt $script:RSArchiveMaxPathLength -or
        $Value.Contains('\') -or
        $Value.StartsWith('/', [StringComparison]::Ordinal) -or
        $Value -match '^[A-Za-z]:' -or
        $Value.IndexOf([char]0) -ge 0) {
        throw [IO.InvalidDataException]::new(
            "$Label has an unsafe archive path.")
    }
    $trimmed = $Value.TrimEnd('/')
    $segments = [string[]]@($trimmed.Split('/'))
    if ($segments.Count -eq 0 -or
        $segments.Count -gt $script:RSArchiveMaxDepth) {
        throw [IO.InvalidDataException]::new(
            "$Label exceeds the archive path-depth bound.")
    }
    foreach ($segment in $segments) {
        $baseName = $segment.Split('.')[0]
        if ([string]::IsNullOrWhiteSpace($segment) -or
            $segment -ceq '.' -or
            $segment -ceq '..' -or
            $segment.Contains(':') -or
            $segment.EndsWith('.', [StringComparison]::Ordinal) -or
            $segment.EndsWith(' ', [StringComparison]::Ordinal) -or
            $baseName -match
                '^(?i:CON|PRN|AUX|NUL|CLOCK\$|COM[1-9]|LPT[1-9])$') {
            throw [IO.InvalidDataException]::new(
                "$Label contains an unsafe Windows path segment.")
        }
    }
    return $trimmed
}

function Assert-RSArchiveBudgets {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Entry,

        [Parameter(Mandatory = $true)]
        [uint64]$RawSizeBytes
    )

    if ($Entry.Count -eq 0 -or
        $Entry.Count -gt $script:RSArchiveMaxEntries) {
        throw [IO.InvalidDataException]::new(
            'Archive entry count violates the fixed bound.')
    }
    [uint64]$expanded = 0
    foreach ($record in $Entry) {
        if ([string]$record.Type -ceq 'directory') {
            continue
        }
        [uint64]$size = $record.SizeBytes
        if ($size -gt $script:RSArchiveMaxEntryBytes -or
            $expanded -gt
                $script:RSArchiveMaxExpandedBytes - $size) {
            throw [IO.InvalidDataException]::new(
                'Archive expanded bytes violate the fixed bound.')
        }
        $expanded += $size
        if ($record.PSObject.Properties.Name -contains
                'CompressedSizeBytes') {
            [uint64]$compressed = $record.CompressedSizeBytes
            if ($size -gt 0 -and $compressed -eq 0) {
                throw [IO.InvalidDataException]::new(
                    'Archive entry has an impossible zero-byte compressed size.')
            }
            if ($compressed -gt 0 -and
                $size -gt
                    $compressed * $script:RSArchiveMaxExpansionRatio) {
                throw [IO.InvalidDataException]::new(
                    'Archive entry exceeds the fixed expansion-ratio bound.')
            }
        }
    }
    if ($RawSizeBytes -eq 0 -or
        $expanded -gt
            $RawSizeBytes * $script:RSArchiveMaxExpansionRatio) {
        throw [IO.InvalidDataException]::new(
            'Archive exceeds the fixed total expansion-ratio bound.')
    }
    return $expanded
}

function Invoke-RSArchiveProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList,

        [Parameter(Mandatory = $true)]
        [string]$WorkingDirectory,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $FilePath
    $startInfo.WorkingDirectory = $WorkingDirectory
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in $ArgumentList) {
        [void]$startInfo.ArgumentList.Add($argument)
    }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) {
            throw [InvalidOperationException]::new(
                "Unable to start $Label.")
        }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(120000)) {
            $process.Kill($true)
            $process.WaitForExit()
            throw [TimeoutException]::new("$Label timed out.")
        }
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        if ($process.ExitCode -ne 0) {
            throw [IO.InvalidDataException]::new(
                "$Label failed with exit $($process.ExitCode). stderr=$stderr")
        }
        return [pscustomobject]@{
            Stdout = $stdout
            Stderr = $stderr
        }
    }
    finally {
        $process.Dispose()
    }
}

function Get-RSZipArchivePlan {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [uint64]$RawSizeBytes,

        [Parameter(Mandatory = $true)]
        [ValidateSet('zip', 'nuget')]
        [string]$ArchiveKind
    )

    $bytes = [IO.File]::ReadAllBytes($LiteralPath)
    $minimumEocd = 22
    if ($bytes.Length -lt $minimumEocd) {
        throw [IO.InvalidDataException]::new('ZIP archive is truncated.')
    }
    $firstEocd = [Math]::Max(0, $bytes.Length - 65557)
    $eocd = -1
    for ($offset = $bytes.Length - $minimumEocd;
        $offset -ge $firstEocd;
        --$offset) {
        if ((Get-RSArchiveUInt32 -Bytes $bytes -Offset $offset) -eq
            [uint32]0x06054b50) {
            $eocd = $offset
            break
        }
    }
    if ($eocd -lt 0) {
        throw [IO.InvalidDataException]::new(
            'ZIP archive has no bounded end-of-central-directory record.')
    }
    $disk = Get-RSArchiveUInt16 -Bytes $bytes -Offset ($eocd + 4)
    $centralDisk = Get-RSArchiveUInt16 -Bytes $bytes -Offset ($eocd + 6)
    $diskEntries = Get-RSArchiveUInt16 -Bytes $bytes -Offset ($eocd + 8)
    $totalEntries = Get-RSArchiveUInt16 -Bytes $bytes -Offset ($eocd + 10)
    $centralSize = Get-RSArchiveUInt32 -Bytes $bytes -Offset ($eocd + 12)
    $centralOffset = Get-RSArchiveUInt32 -Bytes $bytes -Offset ($eocd + 16)
    $commentLength = Get-RSArchiveUInt16 -Bytes $bytes -Offset ($eocd + 20)
    if ($disk -ne 0 -or $centralDisk -ne 0 -or
        $diskEntries -ne $totalEntries -or
        $totalEntries -eq [uint16]::MaxValue -or
        $centralSize -eq [uint32]::MaxValue -or
        $centralOffset -eq [uint32]::MaxValue -or
        $eocd + 22 + $commentLength -ne $bytes.Length -or
        [uint64]$centralOffset + [uint64]$centralSize -gt
            [uint64]$eocd) {
        throw [IO.InvalidDataException]::new(
            'ZIP64, split, appended, or inconsistent ZIP archives are unsupported.')
    }

    [Text.Encoding]::RegisterProvider(
        [Text.CodePagesEncodingProvider]::Instance)
    $utf8 = [Text.UTF8Encoding]::new($false, $true)
    $legacy = [Text.Encoding]::GetEncoding(437)
    $caseFolded =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    $entries = [Collections.Generic.List[object]]::new()
    [uint64]$cursor = $centralOffset
    for ($index = 0; $index -lt $totalEntries; ++$index) {
        if ($cursor + 46 -gt [uint64]$bytes.Length -or
            (Get-RSArchiveUInt32 -Bytes $bytes -Offset ([int]$cursor)) -ne
                [uint32]0x02014b50) {
            throw [IO.InvalidDataException]::new(
                'ZIP central-directory entry is truncated or invalid.')
        }
        $flags = Get-RSArchiveUInt16 -Bytes $bytes -Offset ([int]$cursor + 8)
        $method = Get-RSArchiveUInt16 -Bytes $bytes -Offset ([int]$cursor + 10)
        $compressed = Get-RSArchiveUInt32 `
            -Bytes $bytes `
            -Offset ([int]$cursor + 20)
        $expanded = Get-RSArchiveUInt32 `
            -Bytes $bytes `
            -Offset ([int]$cursor + 24)
        $nameLength = Get-RSArchiveUInt16 `
            -Bytes $bytes `
            -Offset ([int]$cursor + 28)
        $extraLength = Get-RSArchiveUInt16 `
            -Bytes $bytes `
            -Offset ([int]$cursor + 30)
        $entryCommentLength = Get-RSArchiveUInt16 `
            -Bytes $bytes `
            -Offset ([int]$cursor + 32)
        $externalAttributes = Get-RSArchiveUInt32 `
            -Bytes $bytes `
            -Offset ([int]$cursor + 38)
        $entryLength =
            46 + $nameLength + $extraLength + $entryCommentLength
        if ($nameLength -eq 0 -or
            $cursor + [uint64]$entryLength -gt [uint64]$bytes.Length) {
            throw [IO.InvalidDataException]::new(
                'ZIP central-directory name is empty or truncated.')
        }
        if (($flags -band 1) -ne 0 -or
            $method -cnotin @([uint16]0, [uint16]8)) {
            throw [IO.InvalidDataException]::new(
                'Encrypted or unsupported ZIP entries are forbidden.')
        }
        $nameBytes = [byte[]]::new($nameLength)
        [Array]::Copy(
            $bytes,
            [int]$cursor + 46,
            $nameBytes,
            0,
            $nameLength)
        $encoding = if (($flags -band 0x0800) -ne 0) {
            $utf8
        }
        else {
            $legacy
        }
        try {
            $name = $encoding.GetString($nameBytes)
        }
        catch [Text.DecoderFallbackException] {
            throw [IO.InvalidDataException]::new(
                'ZIP entry name is not valid text.')
        }
        $normalized = Assert-RSArchivePath `
            -Value $name `
            -Label "ZIP entry $index"
        if (-not $caseFolded.Add($normalized)) {
            throw [IO.InvalidDataException]::new(
                'ZIP archive has a case-folding duplicate path.')
        }
        $unixMode = [uint32]($externalAttributes -shr 16)
        $fileType = $unixMode -band [uint32]0xF000
        $isDirectory = $name.EndsWith('/', [StringComparison]::Ordinal)
        if ($fileType -eq [uint32]0xA000 -or
            ($fileType -ne 0 -and
                $fileType -ne [uint32]0x8000 -and
                $fileType -ne [uint32]0x4000)) {
            throw [IO.InvalidDataException]::new(
                'ZIP links and special file types are forbidden.')
        }
        if (($fileType -eq [uint32]0x4000 -and -not $isDirectory) -or
            ($fileType -eq [uint32]0x8000 -and $isDirectory) -or
            ($isDirectory -and ($expanded -ne 0 -or $compressed -ne 0))) {
            throw [IO.InvalidDataException]::new(
                'ZIP directory metadata is inconsistent.')
        }
        $entries.Add([pscustomobject]@{
                Path = $normalized
                Type = if ($isDirectory) { 'directory' } else { 'file' }
                SizeBytes = [uint64]$expanded
                CompressedSizeBytes = [uint64]$compressed
                CompressionMethod = [uint16]$method
            })
        $cursor += [uint64]$entryLength
    }
    if ($cursor -ne [uint64]$centralOffset + [uint64]$centralSize) {
        throw [IO.InvalidDataException]::new(
            'ZIP central-directory byte closure is invalid.')
    }
    $totalExpanded = Assert-RSArchiveBudgets `
        -Entry $entries.ToArray() `
        -RawSizeBytes $RawSizeBytes
    return [pscustomobject]@{
        ArchiveKind = $ArchiveKind
        Sha256 = Get-RSArchiveSha256 -LiteralPath $LiteralPath
        RawSizeBytes = $RawSizeBytes
        ExpandedSizeBytes = $totalExpanded
        Entries = [object[]]$entries.ToArray()
    }
}

function Get-RSTarArchivePlan {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [uint64]$RawSizeBytes
    )

    $tar = @(Get-Command `
            tar.exe `
            -CommandType Application `
            -ErrorAction Stop)[0].Source
    $working = [IO.Path]::GetDirectoryName(
        [IO.Path]::GetFullPath($LiteralPath))
    $namesResult = Invoke-RSArchiveProcess `
        -FilePath $tar `
        -ArgumentList @('-tf', $LiteralPath) `
        -WorkingDirectory $working `
        -Label 'tar archive path preflight'
    $typesResult = Invoke-RSArchiveProcess `
        -FilePath $tar `
        -ArgumentList @('-tvf', $LiteralPath, '--numeric-owner') `
        -WorkingDirectory $working `
        -Label 'tar archive type and size preflight'
    $names = [string[]]@(
        $namesResult.Stdout -split "\r?\n" |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $details = [string[]]@(
        $typesResult.Stdout -split "\r?\n" |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($names.Count -ne $details.Count) {
        throw [IO.InvalidDataException]::new(
            'Tar path and metadata listings disagree.')
    }
    $caseFolded =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    $entries = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $names.Count; ++$index) {
        $normalized = Assert-RSArchivePath `
            -Value ($names[$index] -replace '\\', '/') `
            -Label "tar entry $index"
        if (-not $caseFolded.Add($normalized)) {
            throw [IO.InvalidDataException]::new(
                'Tar archive has a case-folding duplicate path.')
        }
        $detail = $details[$index]
        $type = $detail[0]
        if ($type -cne '-' -and $type -cne 'd') {
            throw [IO.InvalidDataException]::new(
                'Tar links and special file types are forbidden.')
        }
        # Numeric owners keep every field through size locale-invariant. The
        # date and time that follow can remain localized because they are not
        # security inputs and are deliberately excluded from this parse.
        $metadata = [Text.RegularExpressions.Regex]::Match(
            $detail,
            '^(?<mode>\S+)\s+(?<links>[0-9]+)\s+(?<uid>[0-9]+)\s+(?<gid>[0-9]+)\s+(?<size>[0-9]+)\s+',
            [Text.RegularExpressions.RegexOptions]::CultureInvariant)
        if (-not $metadata.Success -or
            $metadata.Groups['mode'].Value[0] -cne $type) {
            throw [IO.InvalidDataException]::new(
                'Tar verbose listing has an unsupported metadata format.')
        }
        [uint64]$size = [uint64]::Parse(
            $metadata.Groups['size'].Value,
            [Globalization.CultureInfo]::InvariantCulture)
        if ($type -ceq 'd' -and $size -ne 0) {
            throw [IO.InvalidDataException]::new(
                'Tar directory metadata is inconsistent.')
        }
        $entries.Add([pscustomobject]@{
                Path = $normalized
                Type = if ($type -ceq 'd') { 'directory' } else { 'file' }
                SizeBytes = $size
            })
    }
    $totalExpanded = Assert-RSArchiveBudgets `
        -Entry $entries.ToArray() `
        -RawSizeBytes $RawSizeBytes
    return [pscustomobject]@{
        ArchiveKind = 'tar'
        Sha256 = Get-RSArchiveSha256 -LiteralPath $LiteralPath
        RawSizeBytes = $RawSizeBytes
        ExpandedSizeBytes = $totalExpanded
        Entries = [object[]]$entries.ToArray()
        TarPath = $tar
    }
}

function Assert-RSTerminalEngineArchiveLayout {
    param(
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][string]$ExpectedSha256,
        [Parameter(Mandatory = $true)][string]$InternalRoot,
        [Parameter(Mandatory = $true)][string]$RequiredFileRelativePath,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $plan = Get-RSTerminalEngineArchivePlan `
        -LiteralPath $LiteralPath `
        -ArchiveKind tar
    if ($plan.Sha256 -cne $ExpectedSha256) {
        throw [IO.InvalidDataException]::new(
            "$Label SHA-256 mismatch: $($plan.Sha256)")
    }
    $rootPrefix = "$InternalRoot/"
    $outsideRoot = @($plan.Entries | Where-Object {
            $_.Path -cne $InternalRoot -and
            -not $_.Path.StartsWith(
                $rootPrefix,
                [StringComparison]::Ordinal)
        })
    if ($outsideRoot.Count -ne 0) {
        throw [IO.InvalidDataException]::new(
            "$Label contains entries outside the pinned internal root '$InternalRoot'.")
    }
    $requiredPath = "$InternalRoot/$RequiredFileRelativePath"
    $requiredEntries = @($plan.Entries | Where-Object {
            $_.Path -ceq $requiredPath -and $_.Type -ceq 'file'
        })
    if ($requiredEntries.Count -ne 1) {
        throw [IO.InvalidDataException]::new(
            "$Label must contain exactly one regular '$requiredPath' entry.")
    }
    return $plan
}

function Get-RSTerminalEngineArchivePlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [ValidateSet('zip', 'nuget', 'tar')]
        [string]$ArchiveKind
    )

    $absolute = [IO.Path]::GetFullPath((
            Resolve-Path -LiteralPath $LiteralPath -ErrorAction Stop).Path)
    $item = Get-Item -LiteralPath $absolute -Force
    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
        [uint64]$item.Length -eq 0 -or
        [uint64]$item.Length -gt $script:RSArchiveMaxRawBytes) {
        throw [IO.InvalidDataException]::new(
            'Archive raw file violates the fixed identity and byte bound.')
    }
    if ($ArchiveKind -cin @('zip', 'nuget')) {
        return Get-RSZipArchivePlan `
            -LiteralPath $absolute `
            -RawSizeBytes ([uint64]$item.Length) `
            -ArchiveKind $ArchiveKind
    }
    return Get-RSTarArchivePlan `
        -LiteralPath $absolute `
        -RawSizeBytes ([uint64]$item.Length)
}

function Assert-RSArchiveDestination {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DestinationRoot
    )

    $root = [IO.Path]::GetFullPath($DestinationRoot)
    if ([IO.File]::Exists($root)) {
        throw [IO.InvalidDataException]::new(
            'Archive destination aliases an existing file.')
    }
    if ([IO.Directory]::Exists($root) -and
        @(Get-ChildItem -LiteralPath $root -Force).Count -ne 0) {
        throw [IO.InvalidDataException]::new(
            'Archive destination must be absent or empty.')
    }
    $cursor = [IO.Path]::GetDirectoryName($root)
    while (-not [string]::IsNullOrEmpty($cursor) -and
        [IO.Directory]::Exists($cursor)) {
        if (((Get-Item -LiteralPath $cursor -Force).Attributes -band
                [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw [IO.InvalidDataException]::new(
                'Archive destination has a reparse-point ancestor.')
        }
        $parent = [IO.Path]::GetDirectoryName($cursor)
        if ($parent -ceq $cursor) {
            break
        }
        $cursor = $parent
    }
    [void][IO.Directory]::CreateDirectory($root)
    return $root
}

function Assert-RSArchiveExpandedClosure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [object]$Plan
    )

    $expected =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in [object[]]$Plan.Entries) {
        if ([string]$entry.Type -ceq 'file') {
            $expected.Add([string]$entry.Path, $entry)
        }
    }
    $actualFiles = @(
        Get-ChildItem -LiteralPath $Root -File -Force -Recurse)
    if ($actualFiles.Count -ne $expected.Count) {
        throw [IO.InvalidDataException]::new(
            'Archive extraction file closure differs from preflight.')
    }
    foreach ($item in (
            Get-ChildItem -LiteralPath $Root -Force -Recurse)) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw [IO.InvalidDataException]::new(
                'Archive extraction produced a reparse point.')
        }
        if (-not $item.PSIsContainer) {
            $relative = [IO.Path]::GetRelativePath(
                $Root,
                $item.FullName).Replace('\', '/')
            $record = $null
            if (-not $expected.TryGetValue($relative, [ref]$record) -or
                [uint64]$item.Length -ne [uint64]$record.SizeBytes) {
                throw [IO.InvalidDataException]::new(
                    'Archive extraction identity differs from preflight.')
            }
        }
    }
}

function Expand-RSTerminalEngineArchive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('zip', 'nuget', 'tar')]
        [string]$ArchiveKind,

        [AllowNull()]
        [object]$PreflightPlan = $null
    )

    $absolute = [IO.Path]::GetFullPath((
            Resolve-Path -LiteralPath $LiteralPath -ErrorAction Stop).Path)
    $plan = if ($null -eq $PreflightPlan) {
        Get-RSTerminalEngineArchivePlan `
            -LiteralPath $absolute `
            -ArchiveKind $ArchiveKind
    }
    else {
        $PreflightPlan
    }
    if ([string]$plan.ArchiveKind -cne $ArchiveKind -or
        [string]$plan.Sha256 -cne
            (Get-RSArchiveSha256 -LiteralPath $absolute)) {
        throw [IO.InvalidDataException]::new(
            'Archive changed after preflight or uses another kind.')
    }
    $destination = Assert-RSArchiveDestination `
        -DestinationRoot $DestinationRoot
    try {
        if ($ArchiveKind -cin @('zip', 'nuget')) {
            $stream = [IO.FileStream]::new(
                $absolute,
                [IO.FileMode]::Open,
                [IO.FileAccess]::Read,
                [IO.FileShare]::Read)
            try {
                $archive = [IO.Compression.ZipArchive]::new(
                    $stream,
                    [IO.Compression.ZipArchiveMode]::Read,
                    $false)
                try {
                    $entryByPath =
                        [Collections.Generic.Dictionary[string, object]]::new(
                            [StringComparer]::Ordinal)
                    foreach ($entry in $archive.Entries) {
                        $normalized = Assert-RSArchivePath `
                            -Value $entry.FullName `
                            -Label 'ZIP extraction entry'
                        $entryByPath.Add($normalized, $entry)
                    }
                    foreach ($record in [object[]]$plan.Entries) {
                        $target = [IO.Path]::GetFullPath((
                                Join-Path $destination (
                                    ([string]$record.Path).Replace('/', '\'))))
                        if (-not $target.StartsWith(
                                $destination +
                                    [IO.Path]::DirectorySeparatorChar,
                                [StringComparison]::OrdinalIgnoreCase)) {
                            throw [IO.InvalidDataException]::new(
                                'ZIP extraction path escaped its destination.')
                        }
                        if ([string]$record.Type -ceq 'directory') {
                            [void][IO.Directory]::CreateDirectory($target)
                            continue
                        }
                        [void][IO.Directory]::CreateDirectory(
                            [IO.Path]::GetDirectoryName($target))
                        $input = $entryByPath[[string]$record.Path].Open()
                        try {
                            $output = [IO.FileStream]::new(
                                $target,
                                [IO.FileMode]::CreateNew,
                                [IO.FileAccess]::Write,
                                [IO.FileShare]::None)
                            try {
                                $input.CopyTo($output)
                                $output.Flush($true)
                            }
                            finally {
                                $output.Dispose()
                            }
                        }
                        finally {
                            $input.Dispose()
                        }
                    }
                }
                finally {
                    $archive.Dispose()
                }
            }
            finally {
                $stream.Dispose()
            }
        }
        else {
            [void](Invoke-RSArchiveProcess `
                    -FilePath ([string]$plan.TarPath) `
                    -ArgumentList @('-xf', $absolute, '-C', $destination) `
                    -WorkingDirectory $destination `
                    -Label 'tar archive extraction')
        }
        Assert-RSArchiveExpandedClosure `
            -Root $destination `
            -Plan $plan
        return $plan
    }
    catch [IO.InvalidDataException] {
        throw
    }
    catch [IOException] {
        throw [IO.InvalidDataException]::new(
            'Archive extraction failed.',
            $_.Exception)
    }
}

Export-ModuleMember -Function @(
    'Assert-RSTerminalEngineArchiveLayout',
    'Get-RSTerminalEngineArchivePlan',
    'Expand-RSTerminalEngineArchive')
