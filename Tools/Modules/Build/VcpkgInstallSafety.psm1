# Private build module; import through its reviewed callers.
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

function Test-RSVcpkgVersionDatabaseEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryRoot,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Requested
    )

    $prefix = $Name.Substring(0, 1).ToLowerInvariant()
    $relativePath = "versions/$prefix-/$Name.json"
    $versionPath = Join-Path $RegistryRoot ($relativePath.Replace('/', '\'))
    if (-not (Test-Path -LiteralPath $versionPath -PathType Leaf)) {
        return [pscustomobject]@{
            Found  = $false
            Reason = "version database file missing: $relativePath"
        }
    }

    $database = Get-Content -LiteralPath $versionPath -Raw | ConvertFrom-Json
    $requestedParts = $Requested.Split('#', 2)
    $requestedVersion = $requestedParts[0]
    $requestedPortVersion = if ($requestedParts.Count -eq 2) { $requestedParts[1] } else { $null }
    $versionFields = @('version', 'version-semver', 'version-string', 'version-date', 'version-relaxed')

    foreach ($entry in @($database.versions)) {
        foreach ($field in $versionFields) {
            $versionProperty = $entry.PSObject.Properties[$field]
            if ($null -eq $versionProperty -or [string]$versionProperty.Value -cne $requestedVersion) {
                continue
            }

            if ($null -eq $requestedPortVersion) {
                return [pscustomobject]@{ Found = $true; Reason = $null }
            }

            $portVersionProperty = $entry.PSObject.Properties['port-version']
            $entryPortVersion = if ($null -eq $portVersionProperty) { '0' } else { [string]$portVersionProperty.Value }
            if ($entryPortVersion -ceq $requestedPortVersion) {
                return [pscustomobject]@{ Found = $true; Reason = $null }
            }
        }
    }

    return [pscustomobject]@{
        Found  = $false
        Reason = "no exact entry in $relativePath"
    }
}

function Assert-RSVcpkgManifestVersionConstraintsAvailable {
    <#
    .SYNOPSIS
        Fail before dependency mutation when a manifest version floor or exact
        override is absent from the selected vcpkg registry.

    .DESCRIPTION
        Every object-form dependency with a version>= constraint and every
        manifest override must have an exact version-database entry in the
        registry selected by the caller. This catches a dependency-floor bump
        or pin that was not paired with a compatible builtin-baseline update
        before vcpkg creates staging state.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [Parameter(Mandatory = $true)]
        [string]$RegistryRoot
    )

    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
        throw "vcpkg manifest not found: $ManifestPath"
    }
    if (-not (Test-Path -LiteralPath $RegistryRoot -PathType Container)) {
        throw "vcpkg registry root not found: $RegistryRoot"
    }

    $manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
    $missing = [System.Collections.Generic.List[string]]::new()
    foreach ($dependency in @($manifest.dependencies)) {
        if ($dependency -is [string]) {
            continue
        }

        $constraintProperty = $dependency.PSObject.Properties['version>=']
        if ($null -eq $constraintProperty) {
            continue
        }

        $name = [string]$dependency.name
        $requested = [string]$constraintProperty.Value
        if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($requested)) {
            [void]$missing.Add("malformed object-form dependency with version>= constraint")
            continue
        }

        $result = Test-RSVcpkgVersionDatabaseEntry -RegistryRoot $RegistryRoot -Name $name -Requested $requested
        if (-not $result.Found) {
            [void]$missing.Add("$name >= $requested ($($result.Reason))")
        }
    }

    $versionFields = @('version', 'version-semver', 'version-string', 'version-date', 'version-relaxed')
    $overrideProperty = $manifest.PSObject.Properties['overrides']
    $overrides = if ($null -eq $overrideProperty) { @() } else { @($overrideProperty.Value) }
    foreach ($override in $overrides) {
        $name = [string]$override.name
        $versionProperty = $null
        foreach ($field in $versionFields) {
            $candidate = $override.PSObject.Properties[$field]
            if ($null -ne $candidate) {
                $versionProperty = $candidate
                break
            }
        }

        if ([string]::IsNullOrWhiteSpace($name) -or $null -eq $versionProperty -or
            [string]::IsNullOrWhiteSpace([string]$versionProperty.Value)) {
            [void]$missing.Add('malformed manifest override without a name and exact version')
            continue
        }

        $requested = [string]$versionProperty.Value
        $portVersionProperty = $override.PSObject.Properties['port-version']
        if ($null -ne $portVersionProperty -and [int]$portVersionProperty.Value -gt 0) {
            $requested = "$requested#$([int]$portVersionProperty.Value)"
        }

        $result = Test-RSVcpkgVersionDatabaseEntry -RegistryRoot $RegistryRoot -Name $name -Requested $requested
        if (-not $result.Found) {
            [void]$missing.Add("$name override $requested ($($result.Reason))")
        }
    }

    if ($missing.Count -gt 0) {
        throw (
            "vcpkg manifest version constraints are absent from the selected registry. " +
            "Update vcpkg.json builtin-baseline, dependency floors, and exact overrides together before installing:`n- " +
            ($missing -join "`n- ")
        )
    }
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

Export-ModuleMember -Function @(
    'Assert-RSVcpkgTripletLeafName',
    'Assert-RSVcpkgManifestVersionConstraintsAvailable',
    'Resolve-RSVcpkgSafeChildPath',
    'Get-RSVcpkgFileSha256Hex',
    'Merge-RSVcpkgTripletSafe'
)
