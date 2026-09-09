Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot (
        'Tools\TerminalEngine\TerminalEngineArchive.psm1')) -Force

function New-RSArchiveTestZip {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [Collections.IDictionary]$File,

        [switch]$IncludeDirectory
    )

    $parent = [IO.Path]::GetDirectoryName($Path)
    [void][IO.Directory]::CreateDirectory($parent)
    $stream = [IO.FileStream]::new(
        $Path,
        [IO.FileMode]::CreateNew,
        [IO.FileAccess]::ReadWrite,
        [IO.FileShare]::None)
    try {
        $archive = [IO.Compression.ZipArchive]::new(
            $stream,
            [IO.Compression.ZipArchiveMode]::Create,
            $true)
        try {
            if ($IncludeDirectory) {
                $null = $archive.CreateEntry('root/')
            }
            foreach ($name in $File.Keys) {
                $entry = $archive.CreateEntry(
                    [string]$name,
                    [IO.Compression.CompressionLevel]::Optimal)
                $entryStream = $entry.Open()
                try {
                    $bytes = [Text.UTF8Encoding]::new($false).GetBytes(
                        [string]$File[$name])
                    $entryStream.Write($bytes, 0, $bytes.Length)
                }
                finally {
                    $entryStream.Dispose()
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
    return $Path
}

function Set-RSTarHeaderAscii {
    param(
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [Parameter(Mandatory = $true)][int]$Offset,
        [Parameter(Mandatory = $true)][int]$Length,
        [Parameter(Mandatory = $true)][string]$Value
    )

    $encoded = [Text.Encoding]::ASCII.GetBytes($Value)
    if ($encoded.Length -gt $Length) {
        throw "Tar test field '$Value' exceeds $Length bytes."
    }
    [Array]::Copy($encoded, 0, $Bytes, $Offset, $encoded.Length)
}

function New-RSArchiveTestTar {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object[]]$Entry
    )

    $parent = [IO.Path]::GetDirectoryName($Path)
    [void][IO.Directory]::CreateDirectory($parent)
    $stream = [IO.FileStream]::new(
        $Path,
        [IO.FileMode]::CreateNew,
        [IO.FileAccess]::Write,
        [IO.FileShare]::None)
    try {
        foreach ($record in $Entry) {
            [byte[]]$payload = @()
            if ($null -ne $record.Text) {
                $payload = [Text.UTF8Encoding]::new($false).GetBytes(
                    [string]$record.Text)
            }
            $header = [byte[]]::new(512)
            Set-RSTarHeaderAscii -Bytes $header -Offset 0 -Length 100 -Value ([string]$record.Path)
            Set-RSTarHeaderAscii -Bytes $header -Offset 100 -Length 8 -Value "0000777`0"
            Set-RSTarHeaderAscii -Bytes $header -Offset 108 -Length 8 -Value "0000000`0"
            Set-RSTarHeaderAscii -Bytes $header -Offset 116 -Length 8 -Value "0000000`0"
            $sizeField = [Convert]::ToString($payload.Length, 8).PadLeft(11, '0') + [char]0
            Set-RSTarHeaderAscii -Bytes $header -Offset 124 -Length 12 -Value $sizeField
            Set-RSTarHeaderAscii -Bytes $header -Offset 136 -Length 12 -Value "00000000000`0"
            [Array]::Fill[byte]($header, [byte][char]' ', 148, 8)
            $header[156] = [byte][char]([string]$record.Type)[0]
            if (-not [string]::IsNullOrEmpty([string]$record.LinkName)) {
                Set-RSTarHeaderAscii -Bytes $header -Offset 157 -Length 100 -Value ([string]$record.LinkName)
            }
            Set-RSTarHeaderAscii -Bytes $header -Offset 257 -Length 6 -Value "ustar`0"
            Set-RSTarHeaderAscii -Bytes $header -Offset 263 -Length 2 -Value '00'
            [uint64]$checksum = 0
            foreach ($value in $header) {
                $checksum += $value
            }
            $checksumField = [Convert]::ToString($checksum, 8).PadLeft(6, '0') + "`0 "
            Set-RSTarHeaderAscii -Bytes $header -Offset 148 -Length 8 -Value $checksumField
            $stream.Write($header, 0, $header.Length)
            if ($payload.Length -gt 0) {
                $stream.Write($payload, 0, $payload.Length)
                $padding = (512 - ($payload.Length % 512)) % 512
                if ($padding -gt 0) {
                    $stream.Write([byte[]]::new($padding), 0, $padding)
                }
            }
        }
        $stream.Write([byte[]]::new(1024), 0, 1024)
    }
    finally {
        $stream.Dispose()
    }
    return $Path
}

function Test-RSArchiveThrows {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    try {
        & $Action | Out-Null
        return $false
    }
    catch {
        return $true
    }
}

Describe 'Terminal-engine bounded archive preflight' {
    It 'preflights and extracts a valid ZIP byte-for-byte' {
        $archive = New-RSArchiveTestZip `
            -Path (Join-Path $TestDrive 'valid.zip') `
            -IncludeDirectory `
            -File @{
                'root/LICENSE.txt' = "license`n"
                'root/include/header.h' = "#pragma once`n"
            }
        $plan = Get-RSTerminalEngineArchivePlan `
            -LiteralPath $archive `
            -ArchiveKind zip
        $destination = Join-Path $TestDrive 'expanded'
        $expanded = Expand-RSTerminalEngineArchive `
            -LiteralPath $archive `
            -DestinationRoot $destination `
            -ArchiveKind zip `
            -PreflightPlan $plan

        $expanded.Sha256 | Should Be $plan.Sha256
        @($plan.Entries).Count | Should Be 3
        [IO.File]::ReadAllText(
            (Join-Path $destination 'root\LICENSE.txt')) |
            Should Be "license`n"
    }

    It 'rejects traversal, ADS, device, trailing-dot, and case-fold collisions' {
        $cases = @(
            @{ Name = 'traversal'; Files = @{ '../escape.txt' = 'x' } },
            @{ Name = 'ads'; Files = @{ 'root/file.txt:stream' = 'x' } },
            @{ Name = 'device'; Files = @{ 'root/NUL.txt' = 'x' } },
            @{ Name = 'trailing'; Files = @{ 'root/folder./file.txt' = 'x' } })
        foreach ($case in $cases) {
            $archive = New-RSArchiveTestZip `
                -Path (Join-Path $TestDrive "$($case.Name).zip") `
                -File $case.Files
            (Test-RSArchiveThrows -Action {
                    Get-RSTerminalEngineArchivePlan `
                        -LiteralPath $archive `
                        -ArchiveKind zip
                }) | Should Be $true
        }
        $caseFoldFiles =
            [Collections.Specialized.OrderedDictionary]::new(
                [StringComparer]::Ordinal)
        $caseFoldFiles.Add('root/File.txt', 'x')
        $caseFoldFiles.Add('root/file.txt', 'y')
        $caseFoldArchive = New-RSArchiveTestZip `
            -Path (Join-Path $TestDrive 'casefold.zip') `
            -File $caseFoldFiles
        (Test-RSArchiveThrows -Action {
                Get-RSTerminalEngineArchivePlan `
                    -LiteralPath $caseFoldArchive `
                    -ArchiveKind zip
            }) | Should Be $true
    }

    It 'rejects an encrypted central-directory flag before extraction' {
        $archive = New-RSArchiveTestZip `
            -Path (Join-Path $TestDrive 'encrypted.zip') `
            -File @{ 'root/file.txt' = 'x' }
        $bytes = [IO.File]::ReadAllBytes($archive)
        $signature = [byte[]](0x50, 0x4b, 0x01, 0x02)
        $central = -1
        for ($index = 0; $index -le $bytes.Length - 4; ++$index) {
            if ($bytes[$index] -eq $signature[0] -and
                $bytes[$index + 1] -eq $signature[1] -and
                $bytes[$index + 2] -eq $signature[2] -and
                $bytes[$index + 3] -eq $signature[3]) {
                $central = $index
                break
            }
        }
        $central | Should BeGreaterThan -1
        $bytes[$central + 8] = $bytes[$central + 8] -bor 1
        [IO.File]::WriteAllBytes($archive, $bytes)

        (Test-RSArchiveThrows -Action {
                Get-RSTerminalEngineArchivePlan `
                    -LiteralPath $archive `
                    -ArchiveKind zip
            }) | Should Be $true
    }

    It 'rejects a high-ratio compression bomb before extraction' {
        $archive = New-RSArchiveTestZip `
            -Path (Join-Path $TestDrive 'ratio.zip') `
            -File @{ 'root/zeros.bin' = ('0' * 1048576) }

        (Test-RSArchiveThrows -Action {
                Get-RSTerminalEngineArchivePlan `
                    -LiteralPath $archive `
                    -ArchiveKind zip
            }) | Should Be $true
    }

    It 'rejects archive mutation after a successful preflight' {
        $archive = New-RSArchiveTestZip `
            -Path (Join-Path $TestDrive 'changed.zip') `
            -File @{ 'root/file.txt' = 'first' }
        $plan = Get-RSTerminalEngineArchivePlan `
            -LiteralPath $archive `
            -ArchiveKind zip
        [IO.File]::WriteAllBytes(
            $archive,
            ([IO.File]::ReadAllBytes($archive) + [byte]0))

        (Test-RSArchiveThrows -Action {
                Expand-RSTerminalEngineArchive `
                    -LiteralPath $archive `
                    -DestinationRoot (Join-Path $TestDrive 'changed-output') `
                    -ArchiveKind zip `
                    -PreflightPlan $plan
            }) | Should Be $true
    }

    It 'preflights and extracts a valid tar archive before writing outputs' {
        $source = Join-Path $TestDrive 'tar-source'
        [void][IO.Directory]::CreateDirectory(
            (Join-Path $source 'root'))
        [IO.File]::WriteAllText(
            (Join-Path $source 'root\LICENSE.txt'),
            "tar license`n",
            [Text.UTF8Encoding]::new($false))
        $archive = Join-Path $TestDrive 'valid.tar'
        & tar.exe -cf $archive -C $source root
        $LASTEXITCODE | Should Be 0

        $plan = Get-RSTerminalEngineArchivePlan `
            -LiteralPath $archive `
            -ArchiveKind tar
        $destination = Join-Path $TestDrive 'tar-expanded'
        $expanded = Expand-RSTerminalEngineArchive `
            -LiteralPath $archive `
            -DestinationRoot $destination `
            -ArchiveKind tar `
            -PreflightPlan $plan

        $expanded.Sha256 | Should Be $plan.Sha256
        [IO.File]::ReadAllText(
            (Join-Path $destination 'root\LICENSE.txt')) |
            Should Be "tar license`n"
    }

    It 'rejects unsafe tar paths, duplicate names, links, and special files' {
        $cases = @(
            @{
                Name = 'tar-traversal'
                Entries = @([pscustomobject]@{
                        Path = '../escape.txt'; Type = '0'; Text = 'x'; LinkName = ''
                    })
            },
            @{
                Name = 'tar-duplicate'
                Entries = @(
                    [pscustomobject]@{
                        Path = 'root/file.txt'; Type = '0'; Text = 'x'; LinkName = ''
                    },
                    [pscustomobject]@{
                        Path = 'root/file.txt'; Type = '0'; Text = 'y'; LinkName = ''
                    })
            },
            @{
                Name = 'tar-symlink'
                Entries = @([pscustomobject]@{
                        Path = 'root/link'; Type = '2'; Text = $null; LinkName = '../outside'
                    })
            },
            @{
                Name = 'tar-device'
                Entries = @([pscustomobject]@{
                        Path = 'root/device'; Type = '3'; Text = $null; LinkName = ''
                    })
            })

        foreach ($case in $cases) {
            $archive = New-RSArchiveTestTar `
                -Path (Join-Path $TestDrive "$($case.Name).tar") `
                -Entry $case.Entries
            (Test-RSArchiveThrows -Action {
                    Get-RSTerminalEngineArchivePlan `
                        -LiteralPath $archive `
                        -ArchiveKind tar
                }) | Should Be $true
        }
    }

    It 'locks a tar archive hash, internal root, and required license path independently' {
        $archive = New-RSArchiveTestTar `
            -Path (Join-Path $TestDrive 'pinned-layout.tar') `
            -Entry @(
                [pscustomobject]@{
                    Path = 'pinned-root/LICENSE.md'; Type = '0'; Text = "license`n"; LinkName = ''
                },
                [pscustomobject]@{
                    Path = 'pinned-root/src/file.c'; Type = '0'; Text = 'code'; LinkName = ''
                })
        $sha256 = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).
            Hash.ToLowerInvariant()

        $plan = Assert-RSTerminalEngineArchiveLayout `
            -LiteralPath $archive `
            -ExpectedSha256 $sha256 `
            -InternalRoot 'pinned-root' `
            -RequiredFileRelativePath 'LICENSE.md' `
            -Label 'test archive'
        $plan.Sha256 | Should Be $sha256

        foreach ($invalid in @(
                @{ Hash = ('0' * 64); Root = 'pinned-root'; License = 'LICENSE.md' },
                @{ Hash = $sha256; Root = 'wrong-root'; License = 'LICENSE.md' },
                @{ Hash = $sha256; Root = 'pinned-root'; License = 'COPYING' })) {
            (Test-RSArchiveThrows -Action {
                    Assert-RSTerminalEngineArchiveLayout `
                        -LiteralPath $archive `
                        -ExpectedSha256 $invalid.Hash `
                        -InternalRoot $invalid.Root `
                        -RequiredFileRelativePath $invalid.License `
                        -Label 'test archive'
                }) | Should Be $true
        }
    }
}
