Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$script:RSResourceRoots = @(
    'Common',
    'Plugins',
    'PoC',
    'RedConfigure',
    'RedSalamander',
    'RedSalamanderMonitor',
    'Tests'
)

function Read-RSResourceText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xff -and $bytes[1] -eq 0xfe) {
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

function Get-RSResourceStringEntries {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $text = Read-RSResourceText -Path $File.FullName
    $entries = @()
    $lineNo = 0
    $insideStringTable = $false
    $pendingId = $null
    $pendingLine = 0
    foreach ($line in ($text -split "`r?`n")) {
        $lineNo++

        if (-not $insideStringTable) {
            if ($line -match '^\s*STRINGTABLE\b') {
                $insideStringTable = $true
            }
            continue
        }

        if ($line -match '^\s*END\b') {
            $insideStringTable = $false
            $pendingId = $null
            $pendingLine = 0
            continue
        }

        if ($line -match '^\s*BEGIN\b') {
            continue
        }

        $match = [regex]::Match($line, '^\s*(?<id>[A-Za-z_][A-Za-z0-9_]*)\s+(?<literal>"(?:[^"]|"")*")')
        if ($match.Success) {
            $literal = $match.Groups['literal'].Value
            $entries += [pscustomobject]@{
                Path = $File.FullName
                Line = $lineNo
                Id = $match.Groups['id'].Value
                Text = $literal.Substring(1, $literal.Length - 2).Replace('""', '"')
                IsSatellite = $File.FullName -match '\\Lang\\'
            }
            $pendingId = $null
            $pendingLine = 0
            continue
        }

        if ($null -ne $pendingId) {
            $literalMatch = [regex]::Match($line, '^\s*(?<literal>"(?:[^"]|"")*")')
            if ($literalMatch.Success) {
                $literal = $literalMatch.Groups['literal'].Value
                $entries += [pscustomobject]@{
                    Path = $File.FullName
                    Line = $pendingLine
                    Id = $pendingId
                    Text = $literal.Substring(1, $literal.Length - 2).Replace('""', '"')
                    IsSatellite = $File.FullName -match '\\Lang\\'
                }
                $pendingId = $null
                $pendingLine = 0
                continue
            }
        }

        $idMatch = [regex]::Match($line, '^\s*(?<id>[A-Za-z_][A-Za-z0-9_]*)\s*$')
        if ($idMatch.Success) {
            $pendingId = $idMatch.Groups['id'].Value
            $pendingLine = $lineNo
            continue
        }

        $pendingId = $null
        $pendingLine = 0
    }

    return $entries
}

function Get-RSResourceFiles {
    $resourceFiles = @()
    $resourceFiles += @(Get-ChildItem -LiteralPath $repoRoot -Filter '*.rc')
    foreach ($resourceRoot in $script:RSResourceRoots) {
        $rootPath = Join-Path $repoRoot $resourceRoot
        if (-not (Test-Path -LiteralPath $rootPath)) {
            continue
        }
        $resourceFiles += @(Get-ChildItem -LiteralPath $rootPath -Recurse -Filter '*.rc')
    }
    return $resourceFiles
}

function Get-RSResourceOwnerRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$SatelliteFile
    )

    return $SatelliteFile.Directory.Parent.Parent.FullName.Substring($repoRoot.Length + 1)
}

function Get-RSLanguageNeutralStringIdsByOwner {
    return [ordered]@{
        'RedSalamander' = @(
            'IDS_APP_TITLE',
            'IDS_CONNECTIONS_SECTION_S3',
            'IDS_CONNECTIONS_SECTION_SSH',
            'IDS_FIND_ACTION_HELP',
            'IDS_FMT_FILEOPS_OP_COUNTS',
            'IDS_FMT_FILEOPS_OP_COUNTS_UNKNOWN_TOTAL',
            'IDS_FMT_FILEOPS_OP_STATUS',
            'IDS_FMT_FILEOPS_SIZE_PROGRESS',
            'IDS_FMT_HRESULT_DETAILS',
            'IDS_FMT_STATUS_SELECTED_SINGLE_DIR_ATTRS',
            'IDS_FMT_STATUS_SELECTED_SINGLE_DIR_TIME_ATTRS',
            'IDS_FMT_STATUS_SELECTED_SINGLE_FILE_SIZE_ATTRS',
            'IDS_FMT_STATUS_SELECTED_SINGLE_FILE_SIZE_TIME_ATTRS',
            'IDS_ITEM_PROPERTIES_FIELD_CTAG',
            'IDS_ITEM_PROPERTIES_FIELD_ETAG',
            'IDS_ITEM_PROPERTIES_FIELD_UID',
            'IDS_ITEM_PROPERTIES_FIELD_URL',
            'IDS_ITEM_PROPERTIES_SECTION_IMAP',
            'IDS_ITEM_PROPERTIES_SECTION_S3',
            'IDS_MAKE_FILE_LIST_FORMAT_CSV',
            'IDS_MAKE_FILE_LIST_FORMAT_JSON',
            'IDS_MENU_NAV_ONEDRIVE',
            'IDS_MOD_CTRL',
            'IDS_OVERLAY_DEBUG_SAMPLE_FOLDER_PATH',
            'IDS_PREFS_FILE_ACTION_TEST_FILE_DEFAULT',
            'IDS_PREFS_GENERAL_OPTION_LANGUAGE_CZECH',
            'IDS_PREFS_GENERAL_OPTION_LANGUAGE_ENGLISH',
            'IDS_PREFS_GENERAL_OPTION_LANGUAGE_FRENCH',
            'IDS_PREFS_GENERAL_OPTION_LANGUAGE_JAPANESE',
            'IDS_PREFS_GENERAL_OPTION_LANGUAGE_SLOVAK',
            'IDS_PREFS_GENERAL_SECTION_DXUI',
            'IDS_PREFS_HOT_PATHS_SLOT_HEADER_FMT',
            'IDS_SHORTCUTS_COL_CTRL',
            'IDS_STATUS_SIZE_UNKNOWN',
            'IDS_STATUS_SORT_INDICATOR'
        )
        'Plugins\FileSystemCurl' = @(
            'IDS_FILESYSTEMCURL_FTP_NAME',
            'IDS_FILESYSTEMCURL_IMAP_NAME',
            'IDS_FILESYSTEMCURL_SCP_NAME',
            'IDS_FILESYSTEMCURL_SFTP_NAME'
        )
        'Plugins\FileSystemGoogleDrive' = @(
            'IDS_FILESYSTEMGOOGLEDRIVE_NAME'
        )
        'Plugins\FileSystemMicrosoftDrive' = @(
            'IDS_FILESYSTEMMICROSOFTDRIVE_NAME',
            'IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_APP_TITLE',
            'IDS_FILESYSTEMMICROSOFTDRIVE_SHAREPOINT_NAME'
        )
        'Plugins\FileSystemMtp' = @(
            'IDS_FILESYSTEMMTP_NAME',
            'IDS_FILESYSTEMMTP_FSNAME'
        )
        'Plugins\FileSystemS3' = @(
            'IDS_FILESYSTEMS3_NAME'
        )
        'Plugins\ViewerSpace' = @(
            'IDS_VIEWERSPACE_HEADER_FORMAT',
            'IDS_VIEWERSPACE_TOOLTIP_SHARE_UNKNOWN'
        )
        'Plugins\ViewerSqlite' = @(
            'IDS_VIEWERSQLITE_TITLE_FORMAT'
        )
        'Plugins\ViewerText' = @(
            'IDS_VIEWERTEXT_CODEPAGE_FORMAT',
            'IDS_VIEWERTEXT_ENCODING_UTF16BE',
            'IDS_VIEWERTEXT_ENCODING_UTF8',
            'IDS_VIEWERTEXT_MODE_RAW',
            'IDS_VIEWERTEXT_OFFSET_COL_DEC_FORMAT',
            'IDS_VIEWERTEXT_OFFSET_COL_FORMAT_32',
            'IDS_VIEWERTEXT_OFFSET_COL_FORMAT_64',
            'IDS_VIEWERTEXT_OFFSET_STATUS_FORMAT_32',
            'IDS_VIEWERTEXT_OFFSET_STATUS_FORMAT_64'
        )
        'Plugins\ViewerVLC' = @(
            'IDS_VIEWERVLC_LABEL_TIME_UNKNOWN',
            'IDS_VIEWERVLC_NAME'
        )
    }
}

function Get-RSLanguageNeutralStringIdsForOwner {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OwnerPath
    )

    $byOwner = Get-RSLanguageNeutralStringIdsByOwner
    if ($byOwner.Contains($OwnerPath)) {
        return @($byOwner[$OwnerPath])
    }
    return @()
}

function Get-RSSatelliteMissingStringIdAllowList {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$SatelliteFile
    )

    $allowList = @(Get-RSLanguageNeutralStringIdsForOwner -OwnerPath (Get-RSResourceOwnerRelativePath -SatelliteFile $SatelliteFile))
    if ($SatelliteFile.FullName -match '\\Tests\\LocalizationTests\\Lang\\') {
        $allowList += 'IDS_LOCALIZATION_TEST_EMBEDDED_ONLY'
    }
    return $allowList
}

function Get-RSResourceFormatPlaceholders {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text
    )

    $placeholders = @()
    foreach ($match in [regex]::Matches($Text, '(?<!\{)\{([^{}]*)\}(?!\})')) {
        $body = $match.Groups[1].Value
        if ($body -eq '') {
            $placeholders += [pscustomobject]@{ Kind = 'bare'; Token = $match.Value; Index = $null }
            continue
        }
        if ($body.StartsWith(':')) {
            $placeholders += [pscustomobject]@{ Kind = 'unindexed-format'; Token = $match.Value; Index = $null }
            continue
        }

        $indexed = [regex]::Match($body, '^(?<index>\d+)(?<format>:[^}]*)?$')
        if ($indexed.Success) {
            $placeholders += [pscustomobject]@{ Kind = 'indexed'; Token = $match.Value; Index = [int]$indexed.Groups['index'].Value }
        }
    }

    return $placeholders
}

function Get-RSResourcePlaceholderSignature {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Placeholders
    )

    return (@($Placeholders | Where-Object { $_.Kind -eq 'indexed' } | ForEach-Object { $_.Token } | Sort-Object) -join '|')
}

function Get-RSSatelliteBaseEntries {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$SatelliteFile
    )

    $ownerRoot = $SatelliteFile.Directory.Parent.Parent.FullName
    $baseEntries = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $ownerRoot -Filter '*.rc')) {
        if ($file.FullName -match '\\Lang\\') {
            continue
        }
        foreach ($entry in @(Get-RSResourceStringEntries -File $file)) {
            if (-not $baseEntries.ContainsKey($entry.Id)) {
                $baseEntries[$entry.Id] = $entry
            }
        }
    }

    return $baseEntries
}

Describe 'Resource localization contracts' {
    It 'requires every top-level RC resource root to be explicitly allowlisted' {
        $excludedRoots = @('.build', 'packages', '.claude', '.git')
        $unexpectedRoots = @(
            Get-ChildItem -LiteralPath $repoRoot -Directory |
                Where-Object { $_.Name -notin $excludedRoots -and $_.Name -notin $script:RSResourceRoots } |
                Where-Object { @(Get-ChildItem -LiteralPath $_.FullName -Recurse -Filter '*.rc' -File -ErrorAction SilentlyContinue).Count -gt 0 } |
                Select-Object -ExpandProperty Name
        )

        if ($unexpectedRoots.Count -ne 0) {
            throw "Top-level RC resource roots must be reviewed and added to the localization allowlist: $($unexpectedRoots -join ', ')"
        }
    }

    It 'keeps resource format placeholders positional and translation-safe' {
        $resourceFiles = @(Get-RSResourceFiles)
        $entries = @($resourceFiles | ForEach-Object { Get-RSResourceStringEntries -File $_ })
        $violations = @()

        foreach ($entry in $entries) {
            $placeholders = @(Get-RSResourceFormatPlaceholders -Text $entry.Text)
            foreach ($invalid in @($placeholders | Where-Object { $_.Kind -ne 'indexed' })) {
                $violations += "$($entry.Path):$($entry.Line) $($entry.Id) uses $($invalid.Kind) placeholder $($invalid.Token)"
            }

            $indexes = @($placeholders | Where-Object { $_.Kind -eq 'indexed' } | ForEach-Object { $_.Index })
            if ($indexes.Count -gt 0) {
                $maxIndex = ($indexes | Measure-Object -Maximum).Maximum
                $missing = @(0..$maxIndex | Where-Object { $_ -notin $indexes })
                if ($missing.Count -gt 0) {
                    $violations += "$($entry.Path):$($entry.Line) $($entry.Id) skips placeholder index(es) $($missing -join ',')"
                }

                $firstUseOrder = @()
                foreach ($placeholder in $placeholders) {
                    if ($placeholder.Kind -eq 'indexed' -and $placeholder.Index -notin $firstUseOrder) {
                        $firstUseOrder += $placeholder.Index
                    }
                }
                $expectedOrder = if ($firstUseOrder.Count -eq 1) { @(0) } else { @(0..($firstUseOrder.Count - 1)) }
                if (-not $entry.IsSatellite -and (@($firstUseOrder) -join ',') -ne (@($expectedOrder) -join ',')) {
                    $violations += "$($entry.Path):$($entry.Line) $($entry.Id) source placeholder first-use order is $($firstUseOrder -join ',')"
                }
            }

            $printf = [regex]::Match($entry.Text, '%(?:[-+#0]*(?:\d+|\*)?(?:\.(?:\d+|\*))?(?:hh|h|ll|l|I64|I32|I)?[sSdiuxXfcC])', [System.Text.RegularExpressions.RegexOptions]::None)
            if ($printf.Success) {
                $violations += "$($entry.Path):$($entry.Line) $($entry.Id) uses printf-style placeholder $($printf.Value)"
            }
        }

        foreach ($satelliteFile in @($resourceFiles | Where-Object { $_.FullName -match '\\Lang\\[^\\]+\\.*-[a-z][a-z](?:-[A-Z][A-Z])?\.rc$' })) {
            $sourceEntries = Get-RSSatelliteBaseEntries -SatelliteFile $satelliteFile
            foreach ($targetEntry in @(Get-RSResourceStringEntries -File $satelliteFile)) {
                if (-not $sourceEntries.ContainsKey($targetEntry.Id)) {
                    continue
                }

                $sourceEntry = $sourceEntries[$targetEntry.Id]
                $sourceSignature = Get-RSResourcePlaceholderSignature -Placeholders @(Get-RSResourceFormatPlaceholders -Text $sourceEntry.Text)
                $targetSignature = Get-RSResourcePlaceholderSignature -Placeholders @(Get-RSResourceFormatPlaceholders -Text $targetEntry.Text)
                if ($sourceSignature -ne $targetSignature) {
                    $violations += "$($targetEntry.Path):$($targetEntry.Line) $($targetEntry.Id) placeholder mismatch source=[$sourceSignature] target=[$targetSignature]"
                }
            }
        }

        if (@($violations).Count -ne 0) {
            throw "Resource placeholder contract violations:`r`n$($violations -join "`r`n")"
        }
    }

    It 'keeps documented language-neutral string ids embedded-only in their owning resource module' {
        $violations = @()
        $byOwner = Get-RSLanguageNeutralStringIdsByOwner
        foreach ($ownerPath in $byOwner.Keys) {
            $ownerRoot = Join-Path $repoRoot $ownerPath
            $baseEntries = @{}
            foreach ($baseFile in @(Get-ChildItem -LiteralPath $ownerRoot -Filter '*.rc')) {
                if ($baseFile.FullName -match '\\Lang\\') {
                    continue
                }
                foreach ($entry in @(Get-RSResourceStringEntries -File $baseFile)) {
                    if (-not $baseEntries.ContainsKey($entry.Id)) {
                        $baseEntries[$entry.Id] = $entry
                    }
                }
            }

            $neutralIds = @($byOwner[$ownerPath])
            foreach ($id in $neutralIds) {
                if (-not $baseEntries.ContainsKey($id)) {
                    $violations += "$ownerRoot is missing embedded-only resource $id"
                }
            }

            $satelliteRoot = Join-Path $ownerRoot 'Lang'
            if (-not (Test-Path -LiteralPath $satelliteRoot)) {
                continue
            }

            foreach ($satelliteFile in @(Get-ChildItem -LiteralPath $satelliteRoot -Recurse -Filter '*.rc')) {
                foreach ($entry in @(Get-RSResourceStringEntries -File $satelliteFile)) {
                    if ($entry.Id -in $neutralIds) {
                        $violations += "$($entry.Path):$($entry.Line) $($entry.Id) is language-neutral and must stay in embedded resources only"
                    }
                }
            }
        }

        if (@($violations).Count -ne 0) {
            throw "Language-neutral resource contract violations:`r`n$($violations -join "`r`n")"
        }
    }

    It 'loads documented language-neutral string ids through embedded helpers only' {
        $violations = @()
        $byOwner = Get-RSLanguageNeutralStringIdsByOwner
        $localizedHelperPattern = '\b(?:LoadStringResource|FormatStringResource|MessageBoxResource)\s*\('
        $embeddedNullHelperPattern = '\b(?:LoadEmbeddedStringResource|FormatEmbeddedStringResource)\s*\(\s*nullptr\b'

        foreach ($ownerPath in $byOwner.Keys) {
            $ownerRoot = Join-Path $repoRoot $ownerPath
            $neutralIds = @($byOwner[$ownerPath])
            $neutralIdPattern = '(?<![A-Za-z0-9_])(?:' + (($neutralIds | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')(?![A-Za-z0-9_])'
            $sourceFiles = @(Get-ChildItem -LiteralPath $ownerRoot -Recurse -File |
                Where-Object { $_.Extension -in @('.cpp', '.h', '.hpp', '.inl') -and $_.FullName -notmatch '\\Lang\\' })

            foreach ($sourceFile in $sourceFiles) {
                $text = [System.IO.File]::ReadAllText($sourceFile.FullName)
                $lines = [regex]::Split($text, '\r?\n')
                for ($lineIndex = 0; $lineIndex -lt $lines.Count; $lineIndex++) {
                    $line = $lines[$lineIndex]
                    if ($line -notmatch $localizedHelperPattern -and $line -notmatch $embeddedNullHelperPattern) {
                        continue
                    }

                    $statement = $line
                    $endLineIndex = $lineIndex
                    while ($statement -notmatch ';' -and $endLineIndex + 1 -lt $lines.Count -and ($endLineIndex - $lineIndex) -lt 20) {
                        $endLineIndex++
                        $statement += "`n$($lines[$endLineIndex])"
                    }
                    $statement = ($statement -split ';', 2)[0]

                    if ($statement -match $localizedHelperPattern -and $statement -match $neutralIdPattern) {
                        $id = [regex]::Match($statement, $neutralIdPattern).Value
                        $violations += "$($sourceFile.FullName):$($lineIndex + 1) loads language-neutral $id through a localized resource helper"
                    }

                    if ($ownerPath.StartsWith('Plugins\') -and $statement -match $embeddedNullHelperPattern -and $statement -match $neutralIdPattern) {
                        $id = [regex]::Match($statement, $neutralIdPattern).Value
                        $violations += "$($sourceFile.FullName):$($lineIndex + 1) loads plugin-owned language-neutral $id without the plugin HINSTANCE"
                    }
                }
            }
        }

        if (@($violations).Count -ne 0) {
            throw "Language-neutral lookup contract violations:`r`n$($violations -join "`r`n")"
        }
    }

    It 'keeps satellite string ids complete except documented embedded-only ids' {
        $resourceFiles = @(Get-RSResourceFiles)
        $satelliteFiles = @($resourceFiles | Where-Object { $_.FullName -match '\\Lang\\[^\\]+\\.*-[a-z][a-z](?:-[A-Z][A-Z])?\.rc$' })
        $violations = @()

        foreach ($satelliteFile in $satelliteFiles) {
            $sourceEntries = Get-RSSatelliteBaseEntries -SatelliteFile $satelliteFile
            $targetEntries = @{}
            foreach ($entry in @(Get-RSResourceStringEntries -File $satelliteFile)) {
                if (-not $targetEntries.ContainsKey($entry.Id)) {
                    $targetEntries[$entry.Id] = $entry
                }
            }

            $allowedMissingIds = @(Get-RSSatelliteMissingStringIdAllowList -SatelliteFile $satelliteFile)
            foreach ($sourceId in @($sourceEntries.Keys | Sort-Object)) {
                if ($sourceId -in $allowedMissingIds) {
                    continue
                }
                if (-not $targetEntries.ContainsKey($sourceId)) {
                    $violations += "$($satelliteFile.FullName) is missing localized string id $sourceId"
                }
            }

            foreach ($targetId in @($targetEntries.Keys | Sort-Object)) {
                if (-not $sourceEntries.ContainsKey($targetId)) {
                    $entry = $targetEntries[$targetId]
                    $violations += "$($entry.Path):$($entry.Line) has satellite-only string id $targetId"
                }
            }
        }

        if (@($violations).Count -ne 0) {
            throw "Resource satellite id parity violations:`r`n$($violations -join "`r`n")"
        }
    }
}
