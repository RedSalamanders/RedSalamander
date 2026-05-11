Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

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
    foreach ($line in ($text -split "`r?`n")) {
        $lineNo++
        $match = [regex]::Match($line, '^\s*(?<id>[A-Za-z_][A-Za-z0-9_]*)\s+(?<literal>"(?:[^"]|"")*")')
        if (-not $match.Success) {
            continue
        }

        $literal = $match.Groups['literal'].Value
        $entries += [pscustomobject]@{
            Path = $File.FullName
            Line = $lineNo
            Id = $match.Groups['id'].Value
            Text = $literal.Substring(1, $literal.Length - 2).Replace('""', '"')
            IsSatellite = $File.FullName -match '\\Lang\\'
        }
    }

    return $entries
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
    It 'keeps resource format placeholders positional and translation-safe' {
        $resourceFiles = @(Get-ChildItem -Path $repoRoot -Recurse -Filter '*.rc' | Where-Object { $_.FullName -notmatch '\\.build\\|\\packages\\' })
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
}
