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

function Get-RSMenuOwnerContracts {
    return @(
        [pscustomobject]@{ Owner = 'RedSalamander'; Base = 'RedSalamander\RedSalamander.rc'; SatelliteDirectory = 'RedSalamander\Lang' },
        [pscustomobject]@{ Owner = 'RedSalamanderMonitor'; Base = 'RedSalamanderMonitor\RedSalamanderMonitor.rc'; SatelliteDirectory = 'RedSalamanderMonitor\Lang' },
        [pscustomobject]@{ Owner = 'RedConfigure'; Base = 'RedConfigure\RedConfigure.rc'; SatelliteDirectory = 'RedConfigure\Lang' },
        [pscustomobject]@{ Owner = 'ViewerText'; Base = 'Plugins\ViewerText\ViewerTextResources.rc'; SatelliteDirectory = 'Plugins\ViewerText\Lang' },
        [pscustomobject]@{ Owner = 'ViewerImgRaw'; Base = 'Plugins\ViewerImgRaw\ViewerImgRawResources.rc'; SatelliteDirectory = 'Plugins\ViewerImgRaw\Lang' },
        [pscustomobject]@{ Owner = 'ViewerPE'; Base = 'Plugins\ViewerPE\ViewerPEResources.rc'; SatelliteDirectory = 'Plugins\ViewerPE\Lang' },
        [pscustomobject]@{ Owner = 'ViewerSpace'; Base = 'Plugins\ViewerSpace\ViewerSpaceResources.rc'; SatelliteDirectory = 'Plugins\ViewerSpace\Lang' },
        [pscustomobject]@{ Owner = 'ViewerWeb'; Base = 'Plugins\ViewerWeb\ViewerWebResources.rc'; SatelliteDirectory = 'Plugins\ViewerWeb\Lang' }
    )
}

function Get-RSQuotedRcLiteral {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $match = [regex]::Match($Text, '^\s*"(?<value>(?:[^"]|"")*)"(?<tail>.*)$')
    if (-not $match.Success) {
        return $null
    }

    return [pscustomobject]@{
        Text = $match.Groups['value'].Value.Replace('""', '"')
        Tail = $match.Groups['tail'].Value.Trim()
    }
}

function Get-RSMenuStructures {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $lines = (Read-RSResourceText -Path $File.FullName) -split "`r?`n"
    $menus = [ordered]@{}
    for ($index = 0; $index -lt $lines.Count; $index++) {
        $header = [regex]::Match($lines[$index], '^\s*(?<id>[A-Za-z_][A-Za-z0-9_]*)\s+(?<kind>MENU(?:EX)?)\b')
        if (-not $header.Success) {
            continue
        }

        $menuId = $header.Groups['id'].Value
        $kind = $header.Groups['kind'].Value.ToUpperInvariant()
        $tokens = [System.Collections.Generic.List[object]]::new()
        $depth = 0
        $started = $false
        for ($menuIndex = $index + 1; $menuIndex -lt $lines.Count; $menuIndex++) {
            $trimmed = $lines[$menuIndex].Trim()
            if ($trimmed -eq '') {
                continue
            }

            if ($trimmed -match '^BEGIN\b') {
                $depth++
                $started = $true
                $tokens.Add([pscustomobject]@{ Text = 'BEGIN'; Line = $menuIndex + 1 })
                continue
            }
            if (-not $started) {
                continue
            }
            if ($trimmed -match '^END\b') {
                $tokens.Add([pscustomobject]@{ Text = 'END'; Line = $menuIndex + 1 })
                $depth--
                if ($depth -eq 0) {
                    $index = $menuIndex
                    break
                }
                continue
            }
            if ($trimmed -match '^#(?:if|ifdef|ifndef|else|elif|endif)\b') {
                $tokens.Add([pscustomobject]@{ Text = ($trimmed -replace '\s+', ' '); Line = $menuIndex + 1 })
                continue
            }

            $popup = [regex]::Match($trimmed, '^POPUP\s+(?<body>.+)$')
            if ($popup.Success) {
                $literal = Get-RSQuotedRcLiteral -Text $popup.Groups['body'].Value
                $tail = if ($null -ne $literal) { $literal.Tail } else { $popup.Groups['body'].Value }
                $tokens.Add([pscustomobject]@{ Text = "POPUP|$($tail -replace '\s+', '')"; Line = $menuIndex + 1 })
                continue
            }

            $menuItem = [regex]::Match($trimmed, '^MENUITEM\s+(?<body>.+)$')
            if ($menuItem.Success) {
                $body = $menuItem.Groups['body'].Value.Trim()
                if ($body -match '^(?:SEPARATOR|MFT_SEPARATOR)\b') {
                    $tokens.Add([pscustomobject]@{ Text = 'MENUITEM|SEPARATOR'; Line = $menuIndex + 1 })
                    continue
                }
                $literal = Get-RSQuotedRcLiteral -Text $body
                $tail = if ($null -ne $literal) { $literal.Tail } else { $body }
                $tokens.Add([pscustomobject]@{ Text = "MENUITEM|$($tail -replace '\s+', '')"; Line = $menuIndex + 1 })
            }
        }

        $menus[$menuId] = [pscustomobject]@{
            Id = $menuId
            Kind = $kind
            File = $File.FullName
            HeaderLine = $index + 1
            Tokens = @($tokens)
        }
    }

    return $menus
}

function Compare-RSMenuStructures {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Owner,
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$BaseFile,
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$SatelliteFile
    )

    $baseMenus = Get-RSMenuStructures -File $BaseFile
    $satelliteMenus = Get-RSMenuStructures -File $SatelliteFile
    $findings = [System.Collections.Generic.List[string]]::new()
    foreach ($menuId in $baseMenus.Keys) {
        if (-not $satelliteMenus.Contains($menuId)) {
            $findings.Add("$Owner $($SatelliteFile.FullName): missing $menuId from $($BaseFile.FullName):$($baseMenus[$menuId].HeaderLine)")
            continue
        }

        $expected = $baseMenus[$menuId]
        $actual = $satelliteMenus[$menuId]
        if ($expected.Kind -ne $actual.Kind) {
            $findings.Add("$Owner $menuId kind mismatch: $($BaseFile.FullName):$($expected.HeaderLine)=$($expected.Kind), $($SatelliteFile.FullName):$($actual.HeaderLine)=$($actual.Kind)")
        }

        $tokenCount = [Math]::Max($expected.Tokens.Count, $actual.Tokens.Count)
        for ($tokenIndex = 0; $tokenIndex -lt $tokenCount; $tokenIndex++) {
            if ($tokenIndex -ge $expected.Tokens.Count) {
                $extra = $actual.Tokens[$tokenIndex]
                $findings.Add("$Owner $menuId extra token at $($SatelliteFile.FullName):$($extra.Line): $($extra.Text); token counts base=$($expected.Tokens.Count), satellite=$($actual.Tokens.Count)")
                break
            }
            if ($tokenIndex -ge $actual.Tokens.Count) {
                $missing = $expected.Tokens[$tokenIndex]
                $findings.Add("$Owner $menuId missing token from $($BaseFile.FullName):$($missing.Line): $($missing.Text); token counts base=$($expected.Tokens.Count), satellite=$($actual.Tokens.Count)")
                break
            }

            $expectedToken = $expected.Tokens[$tokenIndex]
            $actualToken = $actual.Tokens[$tokenIndex]
            if ($expectedToken.Text -ne $actualToken.Text) {
                $findings.Add("$Owner $menuId token $tokenIndex mismatch: $($BaseFile.FullName):$($expectedToken.Line) '$($expectedToken.Text)' versus $($SatelliteFile.FullName):$($actualToken.Line) '$($actualToken.Text)'; token counts base=$($expected.Tokens.Count), satellite=$($actual.Tokens.Count)")
                break
            }
        }
    }

    foreach ($menuId in $satelliteMenus.Keys) {
        if (-not $baseMenus.Contains($menuId)) {
            $findings.Add("$Owner $($SatelliteFile.FullName):$($satelliteMenus[$menuId].HeaderLine) has satellite-only menu $menuId")
        }
    }
    return @($findings)
}

function Get-RSMenuAccessKeyEntries {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $lines = (Read-RSResourceText -Path $File.FullName) -split "`r?`n"
    $entries = [System.Collections.Generic.List[object]]::new()
    $menuId = $null
    $depth = 0
    $parentStack = [System.Collections.Generic.Stack[string]]::new()
    $currentParent = $null
    $pendingPopupParent = $null
    $segments = @{}
    for ($index = 0; $index -lt $lines.Count; $index++) {
        $trimmed = $lines[$index].Trim()
        if ($null -eq $menuId) {
            $header = [regex]::Match($trimmed, '^(?<id>[A-Za-z_][A-Za-z0-9_]*)\s+MENU(?:EX)?\b')
            if ($header.Success) {
                $menuId = $header.Groups['id'].Value
                $currentParent = $menuId
                $segments[$currentParent] = 0
            }
            continue
        }

        if ($trimmed -match '^BEGIN\b') {
            $depth++
            if ($null -ne $pendingPopupParent) {
                $parentStack.Push($currentParent)
                $currentParent = $pendingPopupParent
                $segments[$currentParent] = 0
                $pendingPopupParent = $null
            }
            continue
        }
        if ($trimmed -match '^END\b') {
            $depth--
            if ($parentStack.Count -gt 0) {
                $currentParent = $parentStack.Pop()
            }
            if ($depth -eq 0) {
                $menuId = $null
                $currentParent = $null
                $pendingPopupParent = $null
                $parentStack.Clear()
                $segments = @{}
            }
            continue
        }

        $popup = [regex]::Match($trimmed, '^POPUP\s+(?<body>.+)$')
        if ($popup.Success) {
            $literal = Get-RSQuotedRcLiteral -Text $popup.Groups['body'].Value
            if ($null -ne $literal) {
                $entryId = "$menuId/popup/$($index + 1)"
                $entries.Add([pscustomobject]@{
                    MenuId = $menuId
                    Parent = "$currentParent#$($segments[$currentParent])"
                    Label = $literal.Text
                    Tail = $literal.Tail
                    Line = $index + 1
                    File = $File.FullName
                    IsPopup = $true
                })
                $pendingPopupParent = $entryId
            }
            continue
        }

        $menuItem = [regex]::Match($trimmed, '^MENUITEM\s+(?<body>.+)$')
        if (-not $menuItem.Success) {
            continue
        }
        if ($menuItem.Groups['body'].Value -match '^\s*(?:SEPARATOR|MFT_SEPARATOR)\b') {
            $segments[$currentParent] = 1 + $segments[$currentParent]
            continue
        }
        $literal = Get-RSQuotedRcLiteral -Text $menuItem.Groups['body'].Value
        if ($null -eq $literal) {
            continue
        }
        $entries.Add([pscustomobject]@{
            MenuId = $menuId
            Parent = "$currentParent#$($segments[$currentParent])"
            Label = $literal.Text
            Tail = $literal.Tail
            Line = $index + 1
            File = $File.FullName
            IsPopup = $false
        })
    }
    return @($entries)
}

function Get-RSMenuAccessKeys {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Label
    )

    $displayLabel = ($Label -split '\\t', 2)[0]
    $keys = [System.Collections.Generic.List[string]]::new()
    for ($index = 0; $index -lt $displayLabel.Length; $index++) {
        if ($displayLabel[$index] -ne '&') {
            continue
        }
        if (($index + 1) -lt $displayLabel.Length -and $displayLabel[$index + 1] -eq '&') {
            $index++
            continue
        }
        if (($index + 1) -lt $displayLabel.Length) {
            $keys.Add([string]::new($displayLabel[$index + 1], 1).ToUpperInvariant())
        }
    }
    return @($keys)
}

function Test-RSMenuAccessKeys {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $findings = [System.Collections.Generic.List[string]]::new()
    $entries = @(Get-RSMenuAccessKeyEntries -File $File | Where-Object {
        $_.MenuId -ne 'IDR_VIEWERTEXT_ENCODING_CATALOG' -and
        $_.Label -ne '' -and
        $_.Tail -notmatch '(?i)\b(?:GRAYED|INACTIVE|MFS_GRAYED|MFS_DISABLED)\b' -and
        $_.Tail -notmatch '^\s*,?\s*0(?:\s*,|$)' -and
        $_.Label -notmatch '^\s*\d+\s*$'
    })
    foreach ($entry in $entries) {
        $keys = @(Get-RSMenuAccessKeys -Label $entry.Label)
        if ($keys.Count -ne 1) {
            $findings.Add("$($entry.File):$($entry.Line) '$($entry.Label)' has $($keys.Count) access keys; expected exactly one")
        }
    }

    foreach ($group in ($entries | Group-Object Parent)) {
        $keyOwners = @{}
        foreach ($entry in $group.Group) {
            $keys = @(Get-RSMenuAccessKeys -Label $entry.Label)
            if ($keys.Count -ne 1) {
                continue
            }
            $key = $keys[0]
            if ($keyOwners.ContainsKey($key)) {
                $first = $keyOwners[$key]
                $findings.Add("$($entry.File) sibling access key '$key' duplicates '$($first.Label)' at line $($first.Line) and '$($entry.Label)' at line $($entry.Line)")
            }
            else {
                $keyOwners[$key] = $entry
            }
        }
    }
    return @($findings)
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
        $excludedRoots = @('.build', 'vcpkg_installed', 'packages', '.claude', '.git')
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

    It 'keeps file-operations count summaries count-neutral and removes the retired combined format' {
        $expectedByFile = [ordered]@{
            'RedSalamander.rc' = @('Running: {0:L}', 'Waiting: {0:L}', 'Needs attention: {0:L}')
            'RedSalamander-fr-FR.rc' = @('En cours : {0:L}', 'En attente : {0:L}', 'Attention requise : {0:L}')
            'RedSalamander-ja-JP.rc' = @('実行中: {0:L}', '待機中: {0:L}', '注意が必要: {0:L}')
            'RedSalamander-cs-CZ.rc' = @('Probíhá: {0:L}', 'Čeká: {0:L}', 'Vyžaduje pozornost: {0:L}')
            'RedSalamander-sk-SK.rc' = @('Prebieha: {0:L}', 'Čaká: {0:L}', 'Vyžaduje pozornosť: {0:L}')
        }
        $summaryIds = @(
            'IDS_FMT_FILEOPS_GLOBAL_RUNNING_COUNT',
            'IDS_FMT_FILEOPS_GLOBAL_WAITING_COUNT',
            'IDS_FMT_FILEOPS_GLOBAL_ATTENTION_COUNT'
        )
        $resourceFiles = @(Get-RSResourceFiles | Where-Object { $_.Name -in $expectedByFile.Keys })
        foreach ($resourceFile in $resourceFiles) {
            $entries = @{}
            foreach ($entry in @(Get-RSResourceStringEntries -File $resourceFile)) {
                $entries[$entry.Id] = $entry.Text
            }
            $expected = $expectedByFile[$resourceFile.Name]
            for ($index = 0; $index -lt $summaryIds.Count; $index++) {
                $actual = $entries[$summaryIds[$index]]
                if ($actual -cne $expected[$index]) {
                    throw "$($resourceFile.FullName) has '$actual' for $($summaryIds[$index]); expected exact text '$($expected[$index])'."
                }
            }
            if ($entries.ContainsKey('IDS_FMT_FILEOPS_GLOBAL_STATUS_SUMMARY')) {
                throw "$($resourceFile.FullName) still contains retired IDS_FMT_FILEOPS_GLOBAL_STATUS_SUMMARY."
            }
        }
        if ($resourceFiles.Count -ne 5) {
            throw "Expected exactly 5 File Operations resource files, found $($resourceFiles.Count)."
        }
    }

    It 'keeps base and satellite MENU structures identical' {
        $findings = [System.Collections.Generic.List[string]]::new()
        foreach ($contract in @(Get-RSMenuOwnerContracts)) {
            $basePath = Join-Path $repoRoot $contract.Base
            $baseFile = Get-Item -LiteralPath $basePath
            $satelliteRoot = Join-Path $repoRoot $contract.SatelliteDirectory
            $satelliteFiles = @(Get-ChildItem -LiteralPath $satelliteRoot -Recurse -Filter '*.rc')
            if ($satelliteFiles.Count -eq 0) {
                $findings.Add("$($contract.Owner) has no satellite menu resources below $satelliteRoot")
                continue
            }
            foreach ($satelliteFile in $satelliteFiles) {
                foreach ($finding in @(Compare-RSMenuStructures -Owner $contract.Owner -BaseFile $baseFile -SatelliteFile $satelliteFile)) {
                    $findings.Add($finding)
                }
            }
        }

        if ($findings.Count -ne 0) {
            throw "MENU/MENUEX structural parity violations:`r`n$($findings -join "`r`n")"
        }
    }

    It 'keeps the reviewed context, viewer, and monitor command placement exact' {
        function Get-MenuCommandIds {
            param(
                [Parameter(Mandatory = $true)] [string]$RelativePath,
                [Parameter(Mandatory = $true)] [string]$MenuId
            )

            $menus = Get-RSMenuStructures -File (Get-Item -LiteralPath (Join-Path $repoRoot $RelativePath))
            if (-not $menus.Contains($MenuId)) {
                throw "$RelativePath does not contain $MenuId."
            }
            $ids = [System.Collections.Generic.List[string]]::new()
            foreach ($token in $menus[$MenuId].Tokens) {
                $match = [regex]::Match($token.Text, '\b(IDM_[A-Z0-9_]+)\b')
                if ($match.Success) {
                    $ids.Add($match.Groups[1].Value)
                }
            }
            return @($ids)
        }

        function Assert-ExactMenuCommandIds {
            param(
                [Parameter(Mandatory = $true)] [string]$RelativePath,
                [Parameter(Mandatory = $true)] [string]$MenuId,
                [Parameter(Mandatory = $true)] [string[]]$Expected,
                [string]$ExcludePattern = ''
            )

            $actual = @(Get-MenuCommandIds -RelativePath $RelativePath -MenuId $MenuId)
            if ($ExcludePattern -ne '') {
                $actual = @($actual | Where-Object { $_ -notmatch $ExcludePattern })
            }
            if (($actual -join '|') -cne ($Expected -join '|')) {
                throw "$RelativePath $MenuId command order was '$($actual -join ', ')'; expected '$($Expected -join ', ')'."
            }
        }

        Assert-ExactMenuCommandIds -RelativePath 'RedSalamander\RedSalamander.rc' -MenuId 'IDR_FOLDERVIEW_ITEM_CONTEXT' -ExcludePattern 'OVERLAY_SAMPLE' -Expected @(
            'IDM_FOLDERVIEW_CONTEXT_OPEN', 'IDM_FOLDERVIEW_CONTEXT_OPEN_WITH', 'IDM_FOLDERVIEW_CONTEXT_VIEW_SPACE',
            'IDM_FOLDERVIEW_CONTEXT_CUT', 'IDM_FOLDERVIEW_CONTEXT_COPY', 'IDM_FOLDERVIEW_CONTEXT_PASTE',
            'IDM_FOLDERVIEW_CONTEXT_MOVE', 'IDM_FOLDERVIEW_CONTEXT_DELETE', 'IDM_FOLDERVIEW_CONTEXT_RENAME',
            'IDM_FOLDERVIEW_CONTEXT_PROPERTIES'
        )
        Assert-ExactMenuCommandIds -RelativePath 'RedSalamander\RedSalamander.rc' -MenuId 'IDR_FOLDERVIEW_BACKGROUND_CONTEXT' -ExcludePattern 'OVERLAY_SAMPLE' -Expected @(
            'IDM_FOLDERVIEW_CONTEXT_PASTE', 'IDM_PANE_CREATE_DIR', 'IDM_PANE_EDIT_NEW',
            'IDM_FOLDERVIEW_CONTEXT_REFRESH', 'IDM_FOLDERVIEW_CONTEXT_VIEW_SPACE'
        )
        Assert-ExactMenuCommandIds -RelativePath 'Plugins\ViewerText\ViewerTextResources.rc' -MenuId 'IDR_VIEWERTEXT_MENU' -Expected @(
            'IDM_VIEWER_FILE_OPEN', 'IDM_VIEWER_FILE_SAVE_AS', 'IDM_VIEWER_FILE_REFRESH',
            'IDM_VIEWER_OTHER_PREVIOUS', 'IDM_VIEWER_OTHER_NEXT', 'IDM_VIEWER_OTHER_FIRST', 'IDM_VIEWER_OTHER_LAST', 'IDM_VIEWER_FILE_EXIT',
            'IDM_VIEWER_SEARCH_FIND', 'IDM_VIEWER_SEARCH_FIND_NEXT', 'IDM_VIEWER_SEARCH_FIND_PREVIOUS',
            'IDM_VIEWER_VIEW_TEXT', 'IDM_VIEWER_VIEW_HEX', 'IDM_VIEWER_VIEW_HEX_BYTE_COLORS_LEADING_NIBBLE', 'IDM_VIEWER_VIEW_HEX_BYTE_COLORS_OFF',
            'IDM_VIEWER_VIEW_DIFF_SIDE_BY_SIDE', 'IDM_VIEWER_VIEW_DIFF_INLINE', 'IDM_VIEWER_VIEW_DIFF_SHOW_UNCHANGED',
            'IDM_VIEWER_VIEW_DIFF_NEXT_HUNK', 'IDM_VIEWER_VIEW_DIFF_PREVIOUS_HUNK', 'IDM_VIEWER_VIEW_LINE_NUMBERS', 'IDM_VIEWER_VIEW_WRAP',
            'IDM_VIEWER_VIEW_GOTO_TOP', 'IDM_VIEWER_VIEW_GOTO_BOTTOM', 'IDM_VIEWER_VIEW_GOTO_OFFSET',
            'IDM_VIEWER_ENCODING_NEXT', 'IDM_VIEWER_ENCODING_PREVIOUS', 'IDM_VIEWER_ENCODING_DISPLAY_UTF8',
            'IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM', 'IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM', 'IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM',
            'IDM_VIEWER_ENCODING_DISPLAY_ANSI', 'IDM_VIEWER_ENCODING_MORE', 'IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL',
            'IDM_VIEWER_ENCODING_SAVE_UTF8', 'IDM_VIEWER_ENCODING_SAVE_UTF8_BOM', 'IDM_VIEWER_ENCODING_SAVE_UTF16LE_BOM',
            'IDM_VIEWER_ENCODING_SAVE_UTF16BE_BOM'
        )
        Assert-ExactMenuCommandIds -RelativePath 'Plugins\ViewerImgRaw\ViewerImgRawResources.rc' -MenuId 'IDR_VIEWERRAW_MENU' -Expected @(
            'IDM_VIEWERRAW_FILE_REFRESH', 'IDM_VIEWERRAW_FILE_EXPORT', 'IDM_VIEWERRAW_OTHER_PREVIOUS', 'IDM_VIEWERRAW_OTHER_NEXT',
            'IDM_VIEWERRAW_OTHER_FIRST', 'IDM_VIEWERRAW_OTHER_LAST', 'IDM_VIEWERRAW_FILE_EXIT',
            'IDM_VIEWERRAW_VIEW_FIT', 'IDM_VIEWERRAW_VIEW_ACTUAL_SIZE', 'IDM_VIEWERRAW_VIEW_TOGGLE_FIT_100',
            'IDM_VIEWERRAW_VIEW_ZOOM_IN', 'IDM_VIEWERRAW_VIEW_ZOOM_OUT', 'IDM_VIEWERRAW_VIEW_ZOOM_RESET',
            'IDM_VIEWERRAW_VIEW_ROTATE_CW', 'IDM_VIEWERRAW_VIEW_ROTATE_CCW', 'IDM_VIEWERRAW_VIEW_FLIP_HORIZONTAL',
            'IDM_VIEWERRAW_VIEW_FLIP_VERTICAL', 'IDM_VIEWERRAW_VIEW_RESET_ORIENTATION',
            'IDM_VIEWERRAW_VIEW_BRIGHTNESS_INCREASE', 'IDM_VIEWERRAW_VIEW_BRIGHTNESS_DECREASE',
            'IDM_VIEWERRAW_VIEW_CONTRAST_INCREASE', 'IDM_VIEWERRAW_VIEW_CONTRAST_DECREASE',
            'IDM_VIEWERRAW_VIEW_GAMMA_INCREASE', 'IDM_VIEWERRAW_VIEW_GAMMA_DECREASE',
            'IDM_VIEWERRAW_VIEW_TOGGLE_GRAYSCALE', 'IDM_VIEWERRAW_VIEW_TOGGLE_NEGATIVE',
            'IDM_VIEWERRAW_VIEW_SOURCE_RAW', 'IDM_VIEWERRAW_VIEW_SOURCE_THUMBNAIL', 'IDM_VIEWERRAW_VIEW_SHOW_EXIF_OVERLAY'
        )
        Assert-ExactMenuCommandIds -RelativePath 'Plugins\ViewerPE\ViewerPEResources.rc' -MenuId 'IDR_VIEWERPE_MENU' -Expected @(
            'IDM_VIEWERPE_FILE_EXPORT_TEXT', 'IDM_VIEWERPE_FILE_EXPORT_MARKDOWN', 'IDM_VIEWERPE_FILE_REFRESH',
            'IDM_VIEWERPE_OTHER_PREVIOUS', 'IDM_VIEWERPE_OTHER_NEXT', 'IDM_VIEWERPE_OTHER_FIRST', 'IDM_VIEWERPE_OTHER_LAST',
            'IDM_VIEWERPE_FILE_EXIT', 'IDM_VIEWERPE_VIEW_GOTO_TOP', 'IDM_VIEWERPE_VIEW_GOTO_BOTTOM'
        )
        Assert-ExactMenuCommandIds -RelativePath 'Plugins\ViewerSpace\ViewerSpaceResources.rc' -MenuId 'IDR_VIEWERSPACE_MENU' -Expected @(
            'IDM_VIEWERSPACE_FILE_REFRESH', 'IDM_VIEWERSPACE_NAV_UP', 'IDM_VIEWERSPACE_FILE_EXIT'
        )
        Assert-ExactMenuCommandIds -RelativePath 'Plugins\ViewerWeb\ViewerWebResources.rc' -MenuId 'IDR_VIEWERWEB_MENU' -Expected @(
            'IDM_VIEWERWEB_FILE_SAVE_AS', 'IDM_VIEWERWEB_FILE_REFRESH', 'IDM_VIEWERWEB_OTHER_PREVIOUS', 'IDM_VIEWERWEB_OTHER_NEXT',
            'IDM_VIEWERWEB_OTHER_FIRST', 'IDM_VIEWERWEB_OTHER_LAST', 'IDM_VIEWERWEB_FILE_EXIT',
            'IDM_VIEWERWEB_SEARCH_FIND', 'IDM_VIEWERWEB_SEARCH_FIND_NEXT', 'IDM_VIEWERWEB_SEARCH_FIND_PREVIOUS',
            'IDM_VIEWERWEB_VIEW_ZOOM_IN', 'IDM_VIEWERWEB_VIEW_ZOOM_OUT', 'IDM_VIEWERWEB_VIEW_ZOOM_RESET',
            'IDM_VIEWERWEB_TOOLS_COPY_URL', 'IDM_VIEWERWEB_TOOLS_OPEN_EXTERNAL', 'IDM_VIEWERWEB_VIEW_DEVTOOLS',
            'IDM_VIEWERWEB_TOOLS_JSON_EXPAND_ALL', 'IDM_VIEWERWEB_TOOLS_JSON_COLLAPSE_ALL', 'IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE'
        )

        $monitorIds = @(Get-MenuCommandIds -RelativePath 'RedSalamanderMonitor\RedSalamanderMonitor.rc' -MenuId 'IDC_REDSALAMANDERMONITOR')
        $expectedMonitorTail = @(
            'IDM_OPTION_AUTO_SCROLL', 'IDM_OPTION_ID', 'IDM_OPTION_TOP',
            'IDM_FILTER_ERROR', 'IDM_FILTER_WARNING', 'IDM_FILTER_INFO', 'IDM_FILTER_PERF', 'IDM_FILTER_DEBUG', 'IDM_FILTER_TEXT',
            'IDM_FILTER_PRESET_ERRORS_ONLY', 'IDM_FILTER_PRESET_ERRORS_WARNINGS', 'IDM_FILTER_PRESET_ERRORS_DEBUG', 'IDM_FILTER_PRESET_ALL'
        )
        $monitorPlacement = @($monitorIds | Where-Object { $_ -in $expectedMonitorTail })
        if (($monitorPlacement -join '|') -cne ($expectedMonitorTail -join '|')) {
            throw "RedSalamanderMonitor Options command order was '$($monitorPlacement -join ', ')'."
        }
    }

    It 'keeps actionable static sibling menu access keys present and unique' {
        $findings = [System.Collections.Generic.List[string]]::new()
        foreach ($contract in @(Get-RSMenuOwnerContracts)) {
            $files = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            $files.Add((Get-Item -LiteralPath (Join-Path $repoRoot $contract.Base)))
            foreach ($satelliteFile in @(Get-ChildItem -LiteralPath (Join-Path $repoRoot $contract.SatelliteDirectory) -Recurse -Filter '*.rc')) {
                $files.Add($satelliteFile)
            }
            foreach ($file in $files) {
                foreach ($finding in @(Test-RSMenuAccessKeys -File $file)) {
                    $findings.Add($finding)
                }
            }
        }

        if ($findings.Count -ne 0) {
            throw "Static menu access-key violations:`r`n$($findings -join "`r`n")"
        }
    }
}
