function Assert-RSVcpkgTripletLeafName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Triplet
    )

    if ([string]::IsNullOrWhiteSpace($Triplet)) {
        throw "Invalid vcpkg triplet: value must not be empty."
    }

    if ($Triplet -eq "." -or $Triplet -eq ".." -or $Triplet.Contains('\') -or $Triplet.Contains('/')) {
        throw "Invalid vcpkg triplet '$Triplet': pass a triplet leaf name, not a path."
    }

    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        if ($Triplet.Contains([string]$ch)) {
            throw "Invalid vcpkg triplet '$Triplet': contains a character that is not valid in a file name."
        }
    }

    return $Triplet
}

function Resolve-RSVcpkgSafeChildPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [string]$Child,

        [Parameter(Mandatory = $false)]
        [string]$Description = "path"
    )

    if ([string]::IsNullOrWhiteSpace($Root)) {
        throw "Cannot resolve ${Description}: root path is empty."
    }
    if ([string]::IsNullOrWhiteSpace($Child)) {
        throw "Cannot resolve ${Description}: child path is empty."
    }

    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $childFull = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($rootFull, $Child)).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar)

    if ($childFull -eq $rootFull -or -not $childFull.StartsWith("$rootFull\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to use $Description outside its root. root='$rootFull' path='$childFull'"
    }

    return $childFull
}

function Get-RSVcpkgFileSha256Hex {
    <#
    .SYNOPSIS
        Compute SHA-256 of a file using permissive FileShare so concurrent readers/indexers
        do not block us. Throws on any failure (including exclusive/unreadable locks).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $stream = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete)
    )
    try {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        try {
            $bytes = $sha.ComputeHash($stream)
        } finally {
            $sha.Dispose()
        }
        return [System.BitConverter]::ToString($bytes).Replace('-', '')
    } finally {
        $stream.Dispose()
    }
}

function Merge-RSVcpkgTripletSafe {
    <#
    .SYNOPSIS
        Honest, lock-aware merge of a staged vcpkg triplet directory into a destination
        triplet directory.

    .DESCRIPTION
        For every file under SourcePath:
          * Ensure the destination subdirectory exists.
          * If the destination file is absent  -> copy.
          * Else hash source and destination using permissive FileShare.
              - hashes equal     -> skip (no write/delete on destination,
                                          tolerates IDE/indexer read-share locks).
              - hashes differ    -> copy/replace. If the destination is locked and
                                    cannot be replaced, record a clear error.
              - destination hash fails (exclusive / unreadable) -> record a clear error;
                                    we DO NOT silently skip files we cannot verify.

        After processing all files, if any errors were recorded, throw with the full list.

        This function deliberately does NOT prune destination files that are absent in
        the staged output. Pruning is out of scope for this revision because deleting
        unrelated locked files would re-introduce the original failure mode. Stale
        destination extras are possible if a previous vcpkg version produced files the
        current version no longer emits; run `vcpkg-install.ps1 -Clean` to reset.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $true)]
        [string]$TripletName
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
        throw "Merge source not found or not a directory: $SourcePath"
    }
    if (-not (Test-Path -LiteralPath $DestinationPath -PathType Container)) {
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    }

    $sourceFull = [System.IO.Path]::GetFullPath($SourcePath).TrimEnd('\', '/')
    $destFull   = [System.IO.Path]::GetFullPath($DestinationPath).TrimEnd('\', '/')

    Write-Host "  Performing lock-aware merge ($TripletName)..." -ForegroundColor Cyan

    $filesCopied  = 0
    $filesSkipped = 0
    $dirsCreated  = 0
    $errors       = New-Object System.Collections.Generic.List[string]

    $sourceFiles = @(Get-ChildItem -LiteralPath $sourceFull -File -Recurse -Force)
    $totalFiles  = $sourceFiles.Count
    Write-Host "  Analyzing $totalFiles file(s)..." -ForegroundColor Gray

    foreach ($sourceFile in $sourceFiles) {
        $relativePath = $sourceFile.FullName.Substring($sourceFull.Length).TrimStart('\', '/')
        $destFile     = Join-Path $destFull $relativePath
        $destDir      = Split-Path -Parent $destFile

        if (-not (Test-Path -LiteralPath $destDir -PathType Container)) {
            try {
                New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null
                $dirsCreated++
            } catch {
                $errors.Add("Failed to create directory '$destDir': $($_.Exception.Message)")
                continue
            }
        }

        $destExists = Test-Path -LiteralPath $destFile -PathType Leaf

        if ($destExists) {
            $sourceHash = $null
            $destHash   = $null

            try {
                $sourceHash = Get-RSVcpkgFileSha256Hex -Path $sourceFile.FullName
            } catch {
                $errors.Add("Failed to hash staged source '$relativePath': $($_.Exception.Message)")
                continue
            }

            try {
                $destHash = Get-RSVcpkgFileSha256Hex -Path $destFile
            } catch {
                $errors.Add(
                    "Cannot read destination to verify equality: '$destFile'. " +
                    "The file appears to be held with an exclusive/non-shared lock " +
                    "(error: $($_.Exception.Message)). Close Visual Studio, any running " +
                    "build, antivirus real-time scan, and file indexers (Everything, " +
                    "Windows Search) for this path and retry. Refusing to skip silently " +
                    "because we cannot prove the destination matches the staged output."
                )
                continue
            }

            if ($sourceHash -eq $destHash) {
                $filesSkipped++
                Write-Verbose "  Skipped unchanged: $relativePath"
                continue
            }

            Write-Verbose "  Changed (will replace): $relativePath"
        } else {
            Write-Verbose "  New (will copy): $relativePath"
        }

        try {
            Copy-Item -LiteralPath $sourceFile.FullName -Destination $destFile -Force -ErrorAction Stop
            $filesCopied++
        } catch {
            $msg = $_.Exception.Message
            if ($msg -match 'being used by another process|because it is being used|cannot access the file|denied') {
                if ($destExists) {
                    $errors.Add(
                        "Destination file differs from staged output but is locked and " +
                        "cannot be replaced: '$destFile'. Close Visual Studio, builds, " +
                        "indexers, and antivirus that may hold this file, then re-run " +
                        "vcpkg-install.ps1. (Underlying error: $msg)"
                    )
                } else {
                    $errors.Add(
                        "Failed to copy new file to '$destFile' due to access/lock error: $msg"
                    )
                }
            } else {
                $errors.Add("Failed to copy '$relativePath' -> '$destFile': $msg")
            }
        }
    }

    Write-Host "  Merge statistics for ${TripletName}:" -ForegroundColor Cyan
    Write-Host "    Files analyzed: $totalFiles"  -ForegroundColor Gray
    Write-Host "    Files copied:   $filesCopied (absent or changed)" -ForegroundColor Gray
    Write-Host "    Files skipped:  $filesSkipped (unchanged, hash-verified)" -ForegroundColor Gray
    Write-Host "    Dirs created:   $dirsCreated" -ForegroundColor Gray
    Write-Host "    Errors:         $($errors.Count)" -ForegroundColor Gray

    if ($errors.Count -gt 0) {
        Write-Host ""
        Write-Host "  Merge of '$TripletName' failed with $($errors.Count) error(s):" -ForegroundColor Red
        $detail = New-Object System.Text.StringBuilder
        foreach ($e in $errors) {
            Write-Host "    - $e" -ForegroundColor Red
            [void]$detail.AppendLine("- $e")
        }
        throw "Failed to merge triplet '$TripletName' into '$DestinationPath'.`n$($detail.ToString())"
    }

    Write-Host "  Merge of '$TripletName' completed successfully." -ForegroundColor Green

    return [pscustomobject]@{
        Triplet      = $TripletName
        FilesTotal   = $totalFiles
        FilesCopied  = $filesCopied
        FilesSkipped = $filesSkipped
        DirsCreated  = $dirsCreated
    }
}
