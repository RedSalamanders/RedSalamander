Set-StrictMode -Version Latest

function Read-RSPeUInt16 {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [int]$Offset
    )

    if ($Offset -lt 0 -or $Offset + 2 -gt $Bytes.Length) {
        throw "PE read outside file at offset $Offset (UInt16)."
    }
    return [BitConverter]::ToUInt16($Bytes, $Offset)
}

function Read-RSPeUInt32 {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [int]$Offset
    )

    if ($Offset -lt 0 -or $Offset + 4 -gt $Bytes.Length) {
        throw "PE read outside file at offset $Offset (UInt32)."
    }
    return [BitConverter]::ToUInt32($Bytes, $Offset)
}

function Read-RSPeUInt64 {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [int]$Offset
    )

    if ($Offset -lt 0 -or $Offset + 8 -gt $Bytes.Length) {
        throw "PE read outside file at offset $Offset (UInt64)."
    }
    return [BitConverter]::ToUInt64($Bytes, $Offset)
}

function Read-RSPeAscii {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [int]$Offset,

        [int]$MaximumLength = 4096
    )

    if ($Offset -lt 0 -or $Offset -ge $Bytes.Length) {
        throw "PE string starts outside file at offset $Offset."
    }
    $limit = [Math]::Min($Bytes.Length, $Offset + $MaximumLength)
    $end = $Offset
    while ($end -lt $limit -and $Bytes[$end] -ne 0) {
        $end++
    }
    if ($end -eq $limit) {
        throw "PE string at offset $Offset is unterminated or exceeds $MaximumLength bytes."
    }
    return [Text.Encoding]::ASCII.GetString($Bytes, $Offset, $end - $Offset)
}

function ConvertFrom-RSPeRva {
    param(
        [Parameter(Mandatory)]
        [uint32]$Rva,

        [Parameter(Mandatory)]
        [object[]]$Sections,

        [Parameter(Mandatory)]
        [int]$FileLength,

        [Parameter(Mandatory)]
        [uint32]$SizeOfHeaders
    )

    if ($Rva -lt $SizeOfHeaders) {
        if ($Rva -ge $FileLength) {
            throw "PE header RVA 0x$($Rva.ToString('x')) is outside the file."
        }
        return [int]$Rva
    }

    foreach ($section in $Sections) {
        $span = [Math]::Max([uint64]$section.virtualSize, [uint64]$section.rawSize)
        $start = [uint64]$section.virtualAddress
        $candidate = [uint64]$Rva
        if ($candidate -ge $start -and $candidate -lt ($start + $span)) {
            $relative = $candidate - $start
            if ($relative -ge [uint64]$section.rawSize) {
                throw "PE RVA 0x$($Rva.ToString('x')) resolves to zero-filled section data."
            }
            $offset = [uint64]$section.rawPointer + $relative
            if ($offset -ge [uint64]$FileLength) {
                throw "PE RVA 0x$($Rva.ToString('x')) resolves outside the file."
            }
            return [int]$offset
        }
    }

    throw "PE RVA 0x$($Rva.ToString('x')) is not mapped by any section."
}

function Get-RSPeDataDirectory {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [int]$DirectoryOffset,

        [Parameter(Mandatory)]
        [uint32]$DirectoryCount,

        [Parameter(Mandatory)]
        [int]$Index
    )

    if ($Index -ge $DirectoryCount) {
        return [pscustomobject]@{ rva = [uint32]0; size = [uint32]0 }
    }
    $offset = $DirectoryOffset + 8 * $Index
    return [pscustomobject]@{
        rva = Read-RSPeUInt32 -Bytes $Bytes -Offset $offset
        size = Read-RSPeUInt32 -Bytes $Bytes -Offset ($offset + 4)
    }
}

function Read-RSPeImports {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [object[]]$Sections,

        [Parameter(Mandatory)]
        [uint32]$SizeOfHeaders,

        [Parameter(Mandatory)]
        [object]$Directory,

        [Parameter(Mandatory)]
        [bool]$IsPe32Plus
    )

    if ($Directory.rva -eq 0 -or $Directory.size -eq 0) {
        return @()
    }

    $descriptorOffset = ConvertFrom-RSPeRva `
        -Rva $Directory.rva `
        -Sections $Sections `
        -FileLength $Bytes.Length `
        -SizeOfHeaders $SizeOfHeaders
    $records = [Collections.Generic.List[object]]::new()
    $descriptorCount = 0
    while ($true) {
        if (++$descriptorCount -gt 4096) {
            throw 'PE import descriptor count exceeds the fail-closed limit.'
        }
        if ($descriptorOffset + 20 -gt $Bytes.Length) {
            throw 'PE import descriptor extends outside the file.'
        }

        $originalFirstThunk = Read-RSPeUInt32 -Bytes $Bytes -Offset $descriptorOffset
        $timeDateStamp = Read-RSPeUInt32 -Bytes $Bytes -Offset ($descriptorOffset + 4)
        $forwarderChain = Read-RSPeUInt32 -Bytes $Bytes -Offset ($descriptorOffset + 8)
        $nameRva = Read-RSPeUInt32 -Bytes $Bytes -Offset ($descriptorOffset + 12)
        $firstThunk = Read-RSPeUInt32 -Bytes $Bytes -Offset ($descriptorOffset + 16)
        if (($originalFirstThunk -bor $timeDateStamp -bor $forwarderChain -bor $nameRva -bor $firstThunk) -eq 0) {
            break
        }

        if ($nameRva -eq 0) {
            throw 'PE import descriptor has no module name.'
        }
        $nameOffset = ConvertFrom-RSPeRva `
            -Rva $nameRva `
            -Sections $Sections `
            -FileLength $Bytes.Length `
            -SizeOfHeaders $SizeOfHeaders
        $module = (Read-RSPeAscii -Bytes $Bytes -Offset $nameOffset -MaximumLength 512).ToLowerInvariant()
        if ($module -cnotmatch '^[a-z0-9_.-]+\.dll$') {
            throw "PE import module name '$module' is invalid."
        }

        $lookupRva = if ($originalFirstThunk -ne 0) { $originalFirstThunk } else { $firstThunk }
        if ($lookupRva -eq 0) {
            throw "PE import module '$module' has no thunk table."
        }
        $lookupOffset = ConvertFrom-RSPeRva `
            -Rva $lookupRva `
            -Sections $Sections `
            -FileLength $Bytes.Length `
            -SizeOfHeaders $SizeOfHeaders
        $thunkSize = if ($IsPe32Plus) { 8 } else { 4 }
        $symbolCount = 0
        while ($true) {
            if (++$symbolCount -gt 131072) {
                throw "PE import symbol count for '$module' exceeds the fail-closed limit."
            }
            $value = if ($IsPe32Plus) {
                Read-RSPeUInt64 -Bytes $Bytes -Offset $lookupOffset
            } else {
                [uint64](Read-RSPeUInt32 -Bytes $Bytes -Offset $lookupOffset)
            }
            if ($value -eq 0) {
                break
            }

            $ordinalMask = if ($IsPe32Plus) {
                [uint64]::Parse('8000000000000000', [Globalization.NumberStyles]::HexNumber)
            } else {
                [uint64]0x80000000
            }
            if (($value -band $ordinalMask) -ne 0) {
                $symbol = "#$([uint16]($value -band 0xffff))"
            } else {
                if ($value -gt [uint32]::MaxValue) {
                    throw "PE import name RVA for '$module' exceeds 32 bits."
                }
                $hintNameOffset = ConvertFrom-RSPeRva `
                    -Rva ([uint32]$value) `
                    -Sections $Sections `
                    -FileLength $Bytes.Length `
                    -SizeOfHeaders $SizeOfHeaders
                [void](Read-RSPeUInt16 -Bytes $Bytes -Offset $hintNameOffset)
                $symbol = Read-RSPeAscii -Bytes $Bytes -Offset ($hintNameOffset + 2)
            }

            $records.Add([pscustomobject]@{
                module = $module
                symbol = $symbol
            })
            $lookupOffset += $thunkSize
        }

        $descriptorOffset += 20
    }

    return @($records | Sort-Object -Stable -Property `
        @{ Expression = 'module'; Ascending = $true }, `
        @{ Expression = 'symbol'; Ascending = $true })
}

function Read-RSPeExports {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [object[]]$Sections,

        [Parameter(Mandatory)]
        [uint32]$SizeOfHeaders,

        [Parameter(Mandatory)]
        [object]$Directory
    )

    if ($Directory.rva -eq 0 -or $Directory.size -eq 0) {
        return @()
    }
    $offset = ConvertFrom-RSPeRva `
        -Rva $Directory.rva `
        -Sections $Sections `
        -FileLength $Bytes.Length `
        -SizeOfHeaders $SizeOfHeaders
    if ($offset + 40 -gt $Bytes.Length) {
        throw 'PE export directory extends outside the file.'
    }

    $ordinalBase = Read-RSPeUInt32 -Bytes $Bytes -Offset ($offset + 16)
    $functionCount = Read-RSPeUInt32 -Bytes $Bytes -Offset ($offset + 20)
    $nameCount = Read-RSPeUInt32 -Bytes $Bytes -Offset ($offset + 24)
    $functionsRva = Read-RSPeUInt32 -Bytes $Bytes -Offset ($offset + 28)
    $namesRva = Read-RSPeUInt32 -Bytes $Bytes -Offset ($offset + 32)
    $ordinalsRva = Read-RSPeUInt32 -Bytes $Bytes -Offset ($offset + 36)
    if ($functionCount -gt 131072 -or $nameCount -gt 131072 -or $nameCount -gt $functionCount) {
        throw 'PE export counts exceed the fail-closed limit.'
    }
    if ($functionCount -eq 0) {
        return @()
    }

    $functionsOffset = ConvertFrom-RSPeRva `
        -Rva $functionsRva `
        -Sections $Sections `
        -FileLength $Bytes.Length `
        -SizeOfHeaders $SizeOfHeaders
    $nameByIndex = @{}
    if ($nameCount -gt 0) {
        $namesOffset = ConvertFrom-RSPeRva `
            -Rva $namesRva `
            -Sections $Sections `
            -FileLength $Bytes.Length `
            -SizeOfHeaders $SizeOfHeaders
        $ordinalsOffset = ConvertFrom-RSPeRva `
            -Rva $ordinalsRva `
            -Sections $Sections `
            -FileLength $Bytes.Length `
            -SizeOfHeaders $SizeOfHeaders
        for ($index = 0; $index -lt $nameCount; $index++) {
            $nameRva = Read-RSPeUInt32 -Bytes $Bytes -Offset ($namesOffset + 4 * $index)
            $functionIndex = Read-RSPeUInt16 -Bytes $Bytes -Offset ($ordinalsOffset + 2 * $index)
            if ($functionIndex -ge $functionCount) {
                throw 'PE export name references an invalid function index.'
            }
            $nameOffset = ConvertFrom-RSPeRva `
                -Rva $nameRva `
                -Sections $Sections `
                -FileLength $Bytes.Length `
                -SizeOfHeaders $SizeOfHeaders
            $name = Read-RSPeAscii -Bytes $Bytes -Offset $nameOffset
            if ($nameByIndex.ContainsKey([int]$functionIndex)) {
                throw 'PE export function has duplicate public names.'
            }
            $nameByIndex[[int]$functionIndex] = $name
        }
    }

    $records = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $functionCount; $index++) {
        $functionRva = Read-RSPeUInt32 -Bytes $Bytes -Offset ($functionsOffset + 4 * $index)
        if ($functionRva -eq 0) {
            continue
        }
        $forwarder = $null
        $directoryEnd = [uint64]$Directory.rva + [uint64]$Directory.size
        if ([uint64]$functionRva -ge [uint64]$Directory.rva -and [uint64]$functionRva -lt $directoryEnd) {
            $forwarderOffset = ConvertFrom-RSPeRva `
                -Rva $functionRva `
                -Sections $Sections `
                -FileLength $Bytes.Length `
                -SizeOfHeaders $SizeOfHeaders
            $forwarder = Read-RSPeAscii -Bytes $Bytes -Offset $forwarderOffset
        }
        $records.Add([pscustomobject]@{
            name = if ($nameByIndex.ContainsKey($index)) { [string]$nameByIndex[$index] } else { $null }
            ordinal = [uint64]$ordinalBase + [uint64]$index
            forwarder = $forwarder
        })
    }

    return @($records | Sort-Object -Stable -Property `
        @{ Expression = { if ($null -eq $_.name) { '' } else { $_.name } }; Ascending = $true }, `
        @{ Expression = 'ordinal'; Ascending = $true })
}

function Get-RSPeIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    $bytes = [IO.File]::ReadAllBytes($resolved)
    if ($bytes.Length -lt 64 -or $bytes[0] -ne 0x4d -or $bytes[1] -ne 0x5a) {
        throw "File is not a PE image: $resolved"
    }

    $peOffset = [int](Read-RSPeUInt32 -Bytes $bytes -Offset 0x3c)
    if ($peOffset -lt 64 -or $peOffset + 24 -gt $bytes.Length) {
        throw 'PE signature offset is invalid.'
    }
    if ($bytes[$peOffset] -ne 0x50 -or $bytes[$peOffset + 1] -ne 0x45 -or
        $bytes[$peOffset + 2] -ne 0 -or $bytes[$peOffset + 3] -ne 0) {
        throw 'PE signature is invalid.'
    }

    $coffOffset = $peOffset + 4
    $machineValue = Read-RSPeUInt16 -Bytes $bytes -Offset $coffOffset
    $machine = switch ($machineValue) {
        0x8664 { 'AMD64' }
        0xaa64 { 'ARM64' }
        default { "UNKNOWN-0x$($machineValue.ToString('x4'))" }
    }
    $sectionCount = Read-RSPeUInt16 -Bytes $bytes -Offset ($coffOffset + 2)
    if ($sectionCount -eq 0 -or $sectionCount -gt 96) {
        throw "PE section count $sectionCount is invalid."
    }
    $coffCharacteristics = Read-RSPeUInt16 -Bytes $bytes -Offset ($coffOffset + 18)
    $optionalSize = Read-RSPeUInt16 -Bytes $bytes -Offset ($coffOffset + 16)
    $optionalOffset = $coffOffset + 20
    if ($optionalSize -lt 96 -or $optionalOffset + $optionalSize -gt $bytes.Length) {
        throw 'PE optional header size is invalid.'
    }

    $magic = Read-RSPeUInt16 -Bytes $bytes -Offset $optionalOffset
    $isPe32Plus = switch ($magic) {
        0x20b { $true }
        0x10b { $false }
        default { throw "Unsupported PE optional-header magic 0x$($magic.ToString('x4'))." }
    }
    $directoryCountOffset = $optionalOffset + $(if ($isPe32Plus) { 108 } else { 92 })
    $directoryOffset = $optionalOffset + $(if ($isPe32Plus) { 112 } else { 96 })
    $directoryCount = Read-RSPeUInt32 -Bytes $bytes -Offset $directoryCountOffset
    $maximumDirectoryCount = [Math]::Floor(($optionalOffset + $optionalSize - $directoryOffset) / 8)
    if ($directoryCount -gt $maximumDirectoryCount) {
        throw 'PE data-directory count exceeds the optional header.'
    }

    $sizeOfImage = Read-RSPeUInt32 -Bytes $bytes -Offset ($optionalOffset + 56)
    $sizeOfHeaders = Read-RSPeUInt32 -Bytes $bytes -Offset ($optionalOffset + 60)
    $subsystem = Read-RSPeUInt16 -Bytes $bytes -Offset ($optionalOffset + 68)
    $dllCharacteristics = Read-RSPeUInt16 -Bytes $bytes -Offset ($optionalOffset + 70)

    $sectionOffset = $optionalOffset + $optionalSize
    if ($sectionOffset + 40 * $sectionCount -gt $bytes.Length) {
        throw 'PE section table extends outside the file.'
    }
    $sections = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $sectionCount; $index++) {
        $offset = $sectionOffset + 40 * $index
        $nameLength = 0
        while ($nameLength -lt 8 -and $bytes[$offset + $nameLength] -ne 0) {
            $nameLength++
        }
        $name = [Text.Encoding]::ASCII.GetString($bytes, $offset, $nameLength)
        $sections.Add([pscustomobject]@{
            name = $name
            virtualSize = Read-RSPeUInt32 -Bytes $bytes -Offset ($offset + 8)
            virtualAddress = Read-RSPeUInt32 -Bytes $bytes -Offset ($offset + 12)
            rawSize = Read-RSPeUInt32 -Bytes $bytes -Offset ($offset + 16)
            rawPointer = Read-RSPeUInt32 -Bytes $bytes -Offset ($offset + 20)
            characteristics = Read-RSPeUInt32 -Bytes $bytes -Offset ($offset + 36)
        })
    }

    $exportDirectory = Get-RSPeDataDirectory `
        -Bytes $bytes `
        -DirectoryOffset $directoryOffset `
        -DirectoryCount $directoryCount `
        -Index 0
    $importDirectory = Get-RSPeDataDirectory `
        -Bytes $bytes `
        -DirectoryOffset $directoryOffset `
        -DirectoryCount $directoryCount `
        -Index 1
    $debugDirectory = Get-RSPeDataDirectory `
        -Bytes $bytes `
        -DirectoryOffset $directoryOffset `
        -DirectoryCount $directoryCount `
        -Index 6

    $imports = @(Read-RSPeImports `
        -Bytes $bytes `
        -Sections $sections.ToArray() `
        -SizeOfHeaders $sizeOfHeaders `
        -Directory $importDirectory `
        -IsPe32Plus $isPe32Plus)
    $exports = @(Read-RSPeExports `
        -Bytes $bytes `
        -Sections $sections.ToArray() `
        -SizeOfHeaders $sizeOfHeaders `
        -Directory $exportDirectory)

    # Zig/LLD writes per-link volatile COFF/debug timestamps and an RSDS PDB GUID.
    # They do not affect loaded code or data. Normalize only those structurally
    # identified PE fields so the content lock remains stable across rebuilds.
    $normalizedBytes = [byte[]]$bytes.Clone()
    [Array]::Clear($normalizedBytes, $coffOffset + 4, 4)
    if ($debugDirectory.rva -ne 0 -or $debugDirectory.size -ne 0) {
        if ($debugDirectory.rva -eq 0 -or $debugDirectory.size -eq 0 -or
            $debugDirectory.size % 28 -ne 0 -or $debugDirectory.size / 28 -gt 64) {
            throw 'PE debug directory shape is invalid.'
        }
        $debugOffset = ConvertFrom-RSPeRva `
            -Rva $debugDirectory.rva `
            -Sections $sections.ToArray() `
            -FileLength $bytes.Length `
            -SizeOfHeaders $sizeOfHeaders
        if ([uint64]$debugOffset + [uint64]$debugDirectory.size -gt [uint64]$bytes.Length) {
            throw 'PE debug directory extends outside the file.'
        }

        for ($entryOffset = $debugOffset;
             $entryOffset -lt $debugOffset + $debugDirectory.size;
             $entryOffset += 28) {
            [Array]::Clear($normalizedBytes, $entryOffset + 4, 4)
            $debugType = Read-RSPeUInt32 -Bytes $bytes -Offset ($entryOffset + 12)
            $dataSize = Read-RSPeUInt32 -Bytes $bytes -Offset ($entryOffset + 16)
            $dataOffset = Read-RSPeUInt32 -Bytes $bytes -Offset ($entryOffset + 24)
            if ($dataSize -gt 0 -and
                ([uint64]$dataOffset + [uint64]$dataSize -gt [uint64]$bytes.Length)) {
                throw 'PE debug data extends outside the file.'
            }

            # IMAGE_DEBUG_TYPE_CODEVIEW with an RSDS record: preserve the age
            # and PDB leaf, but normalize the volatile 16-byte link GUID.
            if ($debugType -eq 2 -and $dataSize -ge 24 -and
                $bytes[$dataOffset] -eq 0x52 -and $bytes[$dataOffset + 1] -eq 0x53 -and
                $bytes[$dataOffset + 2] -eq 0x44 -and $bytes[$dataOffset + 3] -eq 0x53) {
                [Array]::Clear($normalizedBytes, $dataOffset + 4, 16)
            }
        }
    }

    $sha256 = (Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash.ToLowerInvariant()
    $normalizedSha256 = [Convert]::ToHexString(
        [Security.Cryptography.SHA256]::HashData($normalizedBytes)).ToLowerInvariant()
    $semantic = [ordered]@{
        formatId = 'red-salamander-pe-semantic-identity'
        version = 1
        machine = $machine
        machineValue = [int]$machineValue
        peKind = if ($isPe32Plus) { 'PE32+' } else { 'PE32' }
        coffCharacteristics = [uint32]$coffCharacteristics
        subsystem = [uint32]$subsystem
        dllCharacteristics = [uint32]$dllCharacteristics
        sizeOfImage = [uint64]$sizeOfImage
        sections = @($sections | ForEach-Object {
            [ordered]@{
                name = $_.name
                virtualSize = [uint64]$_.virtualSize
                rawSize = [uint64]$_.rawSize
                characteristics = [uint64]$_.characteristics
            }
        })
        imports = $imports
        exports = $exports
    }

    return [pscustomobject]@{
        path = $resolved
        sizeBytes = [uint64]$bytes.LongLength
        sha256 = $sha256
        normalizedSha256 = $normalizedSha256
        machine = $machine
        semantic = $semantic
    }
}

function Get-RSPeSemanticSha256 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [string]$TerminalEvidenceModule = (Join-Path $PSScriptRoot '..\TerminalEvidence.psm1')
    )

    Import-Module $TerminalEvidenceModule -Force -ErrorAction Stop
    $identity = Get-RSPeIdentity -Path $Path
    return Get-RSJcsSha256 -InputObject $identity.semantic
}

Export-ModuleMember -Function Get-RSPeIdentity, Get-RSPeSemanticSha256
