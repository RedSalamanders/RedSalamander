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

# --- Default MSVC toolset pin -------------------------------------------------------------------------------------------
# vcpkg builds dependencies with the newest MSVC toolset of the Visual Studio installation it prefers, while MSBuild compiles
# and links this product with that installation's default toolset. The two differ when a newer toolset is installed beside
# the default. That newer toolset can lack a compiler for a target (a side-by-side toolset without the x64-hosted ARM64
# compiler failed every fresh ARM64 restore), and a compiled library (curl, 7-Zip, libjpeg-turbo) built by a toolset newer
# than the one that links it is a link hazard. Every triplet installed here is therefore pinned, through an overlay triplet, to
# the installation and default toolset the build discovers.
#
# These functions have the contract of DxUi's Tools/VisualStudio.psm1 and Tools/VcpkgTriplet.psm1, which the pinned library
# documents as part of its consumer interface. They are implemented here instead of imported because this installer runs before
# the pinned DxUi source is restored (the Ghostty ARM64 workflow installs first, and the README names it as the first step), so
# it cannot depend on that checkout's presence. VcpkgInstallSafety.Tests.ps1 compares both implementations whenever the pinned
# checkout is restored.

function Get-RSVcpkgVisualStudioInstallation {
    <#
    .SYNOPSIS
        The Visual Studio installation the build compiles with.

    .DESCRIPTION
        The one build.ps1 selects. In GitHub Actions that is the workflow-selected MSBuild (MSBUILD_EXE_PATH). Everywhere else it
        is the newest installation, prereleases included, that has MSBuild, as vswhere reports it (the query Tools\Restore-DxUi.ps1
        and DxUi use). Nothing here installs or writes anything, and no developer installation is named.
    #>
    [CmdletBinding()]
    param()

    if ($env:GITHUB_ACTIONS -eq 'true' -and -not [string]::IsNullOrWhiteSpace($env:MSBUILD_EXE_PATH)) {
        $selected = [System.IO.Path]::GetFullPath($env:MSBUILD_EXE_PATH)
        $match = [regex]::Match($selected, '^(?<root>.+?)\\MSBuild\\(?:Current|15\.0)\\Bin\\(?:(?:amd64|arm64)\\)?MSBuild\.exe$', 'IgnoreCase')
        if (-not $match.Success) {
            throw "The workflow-selected MSBuild is not inside a Visual Studio installation: $selected"
        }
        return $match.Groups['root'].Value
    }

    $vswhere = @(${env:ProgramFiles(x86)}, $env:ProgramFiles | Where-Object { $_ } |
            ForEach-Object { Join-Path $_ 'Microsoft Visual Studio\Installer\vswhere.exe' }) |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
    if (-not $vswhere) {
        throw 'Visual Studio Installer\vswhere.exe was not found; install Visual Studio 2026 with the C++ tools.'
    }
    $installation = @(& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath) |
        Where-Object { $_ } | Select-Object -First 1
    if (-not $installation) {
        throw 'Visual Studio with MSBuild was not found; install Visual Studio 2026 with the C++ tools.'
    }
    return [string]$installation
}

function Get-RSVcpkgDefaultToolset {
    <#
    .SYNOPSIS
        The MSVC toolset MSBuild compiles with when a project names none.

    .DESCRIPTION
        The version in the installation's VC\Auxiliary\Build\Microsoft.VCToolsVersion.default.txt, such as 14.51.36231, which
        its major.minor (14.51) selects. It can differ from the newest toolset installed side by side, which is what vcpkg would
        choose. A file that is missing or holds anything else fails here, naming the file and the repair, before anything is
        cloned, downloaded or installed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Installation
    )

    $path = [System.IO.Path]::GetFullPath(
        [System.IO.Path]::Combine($Installation, 'VC', 'Auxiliary', 'Build', 'Microsoft.VCToolsVersion.default.txt'))
    $cause = "Cannot pin vcpkg to the default MSVC toolset of the Visual Studio installation at '$Installation'"
    if (-not [System.IO.File]::Exists($path)) {
        throw "${cause}: '$path' does not exist. Install the Desktop development with C++ workload (MSVC build tools) in the Visual Studio Installer, or repair the installation, and try again."
    }

    $text = [System.IO.File]::ReadAllText($path).Trim()
    $match = [regex]::Match($text, '^(?<majorMinor>[0-9]+\.[0-9]+)\.[0-9]+\z')
    if (-not $match.Success) {
        $shown = $text -replace '\s+', ' '
        if ($shown.Length -gt 60) { $shown = $shown.Substring(0, 60) + '...' }
        $holds = if ($shown) { "holds '$shown'" } else { 'is empty' }
        throw "${cause}: '$path' $holds, not a toolset version such as 14.51.36231. Repair the installation in the Visual Studio Installer so the file names the default toolset, and try again."
    }

    return [pscustomobject]@{
        Installation = $Installation
        VersionFile  = $path
        Version      = $text
        MajorMinor   = $match.Groups['majorMinor'].Value
    }
}

function ConvertTo-RSCMakeQuotedArgument {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    # In a CMake quoted argument a backslash, a double quote and a dollar sign are special; every other character is literal.
    return '"' + [regex]::Replace($Value, '[\\"$]', '\$&') + '"'
}

function Get-RSVcpkgOverlayTripletText {
    <#
    .SYNOPSIS
        The text of an overlay triplet: the stock triplet unchanged, then the two toolset pins.

    .DESCRIPTION
        Keeps the stock text's line endings (arm64-windows.cmake is CRLF and x64-windows.cmake LF), so the copy stays one
        consistent file. vcpkg takes an overlay in place of the triplet of the same name. Pure: no I/O.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$StockText,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Installation,

        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[0-9]+\.[0-9]+$')]
        [string]$ToolsetVersion
    )

    $newline = if ($StockText.Contains("`r`n")) { "`r`n" } else { "`n" }
    $text = $StockText
    if ($text.Length -and -not $text.EndsWith("`n", [StringComparison]::Ordinal)) { $text += $newline }
    if ($text.Length) { $text += $newline }
    $pins = @(
        '# RedSalamander: build with the Visual Studio installation and default MSVC toolset the build uses (generated by vcpkg-install.ps1; do not edit).',
        "set(VCPKG_VISUAL_STUDIO_PATH $(ConvertTo-RSCMakeQuotedArgument $Installation))",
        "set(VCPKG_PLATFORM_TOOLSET_VERSION $(ConvertTo-RSCMakeQuotedArgument $ToolsetVersion))")
    return $text + (($pins -join $newline) + $newline)
}

function Resolve-RSVcpkgStockTripletPath {
    <#
    .SYNOPSIS
        The stock triplet file of the selected vcpkg root: triplets\<name>.cmake, else triplets\community\<name>.cmake.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryRoot,

        [Parameter(Mandatory = $true)]
        [string]$Triplet
    )

    [void](Assert-RSVcpkgTripletLeafName -Triplet $Triplet)
    foreach ($relative in @("triplets\$Triplet.cmake", "triplets\community\$Triplet.cmake")) {
        $candidate = Join-Path $RegistryRoot $relative
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    throw "vcpkg triplet '$Triplet' is not one of the triplets of the selected vcpkg root: neither triplets\$Triplet.cmake nor triplets\community\$Triplet.cmake exists beneath '$RegistryRoot'. Install a triplet that vcpkg ships, or check -VcpkgRoot."
}

function Update-RSVcpkgOverlayTriplet {
    <#
    .SYNOPSIS
        Writes <OutputDirectory>\<name>.cmake, the overlay of the stock triplet pinned to the toolset.

    .DESCRIPTION
        The file is rewritten only when its bytes would change, so an unchanged install leaves its timestamp alone. Changing the
        triplet changes vcpkg's package ABI hash, so the first install after a change rebuilds the packages; vcpkg's binary
        cache then serves them to any other root whose triplet text is identical. Returns Triplet, Directory (what
        --overlay-triplets takes), Path and Changed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$StockTripletPath,

        [Parameter(Mandatory = $true)]
        $Toolset,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory
    )

    # Relative paths mean the caller's PowerShell location, which .NET file calls do not know.
    $StockTripletPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($StockTripletPath)
    $OutputDirectory = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
    if (-not [System.IO.File]::Exists($StockTripletPath)) {
        throw "The triplet to pin does not exist: $StockTripletPath"
    }

    $name = [System.IO.Path]::GetFileName($StockTripletPath)
    $text = Get-RSVcpkgOverlayTripletText `
        -StockText ([System.IO.File]::ReadAllText($StockTripletPath)) `
        -Installation $Toolset.Installation `
        -ToolsetVersion $Toolset.MajorMinor
    $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($text)
    $path = Join-Path $OutputDirectory $name
    $changed = -not ([System.IO.File]::Exists($path) -and
        [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($path)) -ceq [Convert]::ToBase64String($bytes))
    if ($changed) {
        [void][System.IO.Directory]::CreateDirectory($OutputDirectory)
        [System.IO.File]::WriteAllBytes($path, $bytes)
    }

    return [pscustomobject]@{
        Triplet   = [System.IO.Path]::GetFileNameWithoutExtension($name)
        Directory = $OutputDirectory
        Path      = $path
        Changed   = $changed
    }
}

function Invoke-RSVcpkgBootstrapGit {
    param([string]$Root, [string[]]$Arguments)
    $output = @(& git -C $Root @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "Pinned vcpkg Git operation failed: git $($Arguments -join ' ') $($output -join [Environment]::NewLine)"
    }
    return ($output -join [Environment]::NewLine)
}

function Assert-RSVcpkgManagedBootstrapPath {
    param([string]$RepoRoot, [string]$Path)
    $relative = [IO.Path]::GetRelativePath($RepoRoot, $Path)
    $current = $RepoRoot
    foreach ($part in ($relative -split '[\\/]')) {
        $current = Join-Path $current $part
        if (Test-Path -LiteralPath $current) {
            if ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Pinned vcpkg bootstrap refuses a reparse point: $current"
            }
        }
    }
}

function Initialize-RSVcpkgToolContext {
    <#
    .SYNOPSIS
        Restore and bootstrap the independently pinned vcpkg tool and ports baseline.
    .DESCRIPTION
        Uses repository-local managed paths, validates both Git identities and tools
        schemas, and preserves an incomplete restored tool directory before cloning.
        A short repository tool-provisioning lock protects only this shared checkout;
        dependency installation keeps its separate platform-scoped locks.
        Existing dirty checkouts and reparse points fail without forced reset/deletion.
    .PARAMETER RepoRoot
        Repository containing vcpkg.json and vcpkg-tool.json.
    .OUTPUTS
        ToolRoot, RegistryRoot, Executable, ToolCommit and Baseline.
    .NOTES
        Requires Git and Windows. Downloads/clones/bootstrap output stay beneath .build.
        Does not install ports, change pins or modify user/global Git configuration.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RepoRoot)
    $RepoRoot = [IO.Path]::GetFullPath($RepoRoot)
    $manifest = Get-Content -LiteralPath (Join-Path $RepoRoot 'vcpkg.json') -Raw | ConvertFrom-Json
    $pin = Get-Content -LiteralPath (Join-Path $RepoRoot 'vcpkg-tool.json') -Raw | ConvertFrom-Json
    $baseline = ([string]$manifest.'builtin-baseline').Trim().ToLowerInvariant()
    $commit = ([string]$pin.commit).Trim().ToLowerInvariant()
    $repository = ([string]$pin.repository).Trim()
    if ($baseline -cnotmatch '^[0-9a-f]{40}$' -or $commit -cnotmatch '^[0-9a-f]{40}$' -or
        [string]::IsNullOrWhiteSpace($repository) -or $repository.StartsWith('-')) {
        throw 'Pinned vcpkg bootstrap requires a repository and full tool/baseline commit identities.'
    }
    $buildRoot = Resolve-RSVcpkgSafeChildPath -Root $RepoRoot -Child '.build'
    $toolRoot = Resolve-RSVcpkgSafeChildPath -Root $RepoRoot -Child '.build/vcpkg-tool'
    $registryRoot = Resolve-RSVcpkgSafeChildPath -Root $RepoRoot -Child '.build/vcpkg-registry'
    foreach ($path in @($buildRoot, $toolRoot, $registryRoot)) {
        Assert-RSVcpkgManagedBootstrapPath -RepoRoot $RepoRoot -Path $path
    }
    [void][IO.Directory]::CreateDirectory($buildRoot)
    $lock = $null
    $wait = [Diagnostics.Stopwatch]::StartNew()
    try {
        while ($null -eq $lock) {
            try {
                $lock = [IO.File]::Open((Join-Path $buildRoot 'vcpkg-bootstrap.lock'),
                    [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            } catch [IO.IOException] {
                if ($wait.Elapsed.TotalMinutes -ge 5) { throw 'Timed out waiting for pinned vcpkg tool provisioning.' }
                [Threading.Thread]::Sleep(100)
            }
        }
        $recovery = $null
        if (-not (Test-Path -LiteralPath (Join-Path $toolRoot '.git') -PathType Container)) {
            if (Test-Path -LiteralPath $toolRoot) {
                $recovery = Resolve-RSVcpkgSafeChildPath -Root $RepoRoot -Child ('.build/vcpkg-tool-recovery-' + [Guid]::NewGuid().ToString('N'))
                Move-Item -LiteralPath $toolRoot -Destination $recovery
                Write-Host "Preserved incomplete vcpkg checkout at $recovery"
            }
            Write-Host (Invoke-RSVcpkgBootstrapGit $RepoRoot @('clone', '--', $repository, $toolRoot))
        }
        $actualRoot = Invoke-RSVcpkgBootstrapGit $toolRoot @('rev-parse', '--show-toplevel')
        if ([IO.Path]::GetFullPath($actualRoot) -ine $toolRoot) { throw 'vcpkg tool is not its own managed Git checkout.' }
        $dirty = Invoke-RSVcpkgBootstrapGit $toolRoot @('status', '--porcelain', '--untracked-files=no')
        if ($dirty) { throw "Pinned vcpkg tool checkout is dirty; preserve or repair it explicitly: $toolRoot" }
        $shallow = Invoke-RSVcpkgBootstrapGit $toolRoot @('rev-parse', '--is-shallow-repository')
        if ($shallow -eq 'true') {
            Write-Host (Invoke-RSVcpkgBootstrapGit $toolRoot @('fetch', '--unshallow', 'origin'))
        }
        foreach ($required in @($commit, $baseline) | Select-Object -Unique) {
            & git -C $toolRoot cat-file -e ($required + '^{commit}') 2>$null
            if ($LASTEXITCODE -ne 0) {
                Write-Host (Invoke-RSVcpkgBootstrapGit $toolRoot @('fetch', 'origin', $required))
                $null = Invoke-RSVcpkgBootstrapGit $toolRoot @('cat-file', '-e', ($required + '^{commit}'))
            }
        }
        $null = Invoke-RSVcpkgBootstrapGit $toolRoot @('show', ($baseline + ':versions/baseline.json'))
        $toolMetadata = Invoke-RSVcpkgBootstrapGit $toolRoot @('show', ($commit + ':scripts/vcpkg-tools.json')) | ConvertFrom-Json
        $portsMetadata = Invoke-RSVcpkgBootstrapGit $toolRoot @('show', ($baseline + ':scripts/vcpkg-tools.json')) | ConvertFrom-Json
        $toolSchema = [int]$toolMetadata.'schema-version'
        $portsSchema = [int]$portsMetadata.'schema-version'
        if ($toolSchema -lt 1 -or $portsSchema -lt 1 -or $toolSchema -lt $portsSchema) {
            throw "vcpkg tools schema is incompatible: tool=$toolSchema baseline=$portsSchema. Update the pins together."
        }
        Write-Host (Invoke-RSVcpkgBootstrapGit $toolRoot @('checkout', '--detach', $commit))
        if ((Invoke-RSVcpkgBootstrapGit $toolRoot @('rev-parse', 'HEAD')) -cne $commit) {
            throw 'Pinned vcpkg tool identity mismatch.'
        }
        if (Test-Path -LiteralPath $registryRoot) {
            $actualRoot = Invoke-RSVcpkgBootstrapGit $registryRoot @('rev-parse', '--show-toplevel')
            if ([IO.Path]::GetFullPath($actualRoot) -ine $registryRoot) { throw 'vcpkg registry is not its own managed Git checkout.' }
            if (Invoke-RSVcpkgBootstrapGit $registryRoot @('status', '--porcelain', '--untracked-files=no')) {
                throw "Pinned vcpkg registry is dirty; preserve or repair it explicitly: $registryRoot"
            }
            Write-Host (Invoke-RSVcpkgBootstrapGit $registryRoot @('checkout', '--detach', $baseline))
        } else {
            Write-Host (Invoke-RSVcpkgBootstrapGit $toolRoot @('worktree', 'add', '--detach', $registryRoot, $baseline))
        }
        if ((Invoke-RSVcpkgBootstrapGit $registryRoot @('rev-parse', 'HEAD')) -cne $baseline) {
            throw 'Pinned vcpkg registry identity mismatch.'
        }
        if ($recovery) {
            $downloads = Join-Path $recovery 'downloads'
            Assert-RSVcpkgManagedBootstrapPath -RepoRoot $RepoRoot -Path $downloads
            if (Test-Path -LiteralPath $downloads -PathType Container) {
                Move-Item -LiteralPath $downloads -Destination (Join-Path $toolRoot 'downloads')
            }
        }
        $bootstrap = Join-Path $toolRoot 'bootstrap-vcpkg.bat'
        & $bootstrap -disableMetrics | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "Pinned vcpkg bootstrap failed with exit code $LASTEXITCODE." }
        $executable = Join-Path $toolRoot 'vcpkg.exe'
        if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) { throw 'Pinned bootstrap did not produce vcpkg.exe.' }
        Write-Host "Pinned vcpkg ready: tool=$commit baseline=$baseline schema=$toolSchema"
        return [pscustomobject]@{ ToolRoot=$toolRoot; RegistryRoot=$registryRoot; Executable=$executable; ToolCommit=$commit; Baseline=$baseline }
    } finally {
        if ($lock) { $lock.Dispose() }
    }
}

function Get-RSVcpkgBuildEnvironment {
    <#
    .SYNOPSIS
        Bind a CI build to its pinned vcpkg integration and installed triplet.
    .DESCRIPTION
        Returns process-local environment values without installing user-wide
        integration. Validates inputs before returning any changes. Callers apply
        the values only to the child build/test lifetime or hosted job.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][string]$ToolRoot,
        [Parameter(Mandatory)][ValidateSet('x64','ARM64')][string]$Platform,
        [Parameter(Mandatory)][ValidateSet('Debug','ASan Debug','Release')][string]$Configuration)
    $triplet = if ($Platform -eq 'ARM64') { 'arm64-windows' } else { 'x64-windows' }
    $installed = Resolve-RSVcpkgSafeChildPath -Root $RepoRoot -Child ('.build/vcpkg_installed/' + $Platform.ToLowerInvariant())
    $tripletRoot = Join-Path $installed $triplet
    $debug = $Configuration -like '*Debug*'
    $include = Join-Path $tripletRoot 'include'
    $lib = Join-Path $tripletRoot $(if ($debug) { 'debug/lib' } else { 'lib' })
    $bin = Join-Path $tripletRoot $(if ($debug) { 'debug/bin' } else { 'bin' })
    $props = Join-Path $ToolRoot 'scripts/buildsystems/msbuild/vcpkg.props'
    $targets = Join-Path $ToolRoot 'scripts/buildsystems/msbuild/vcpkg.targets'
    foreach ($path in @($include,$lib,$props,$targets)) {
        if (-not (Test-Path -LiteralPath $path)) { throw "Pinned vcpkg build input not found: $path" }
    }
    $linkPaths = @($lib)
    $manual = Join-Path $lib 'manual-link'
    if (Test-Path -LiteralPath $manual -PathType Container) { $linkPaths += $manual }
    $linkFlags = ($linkPaths | ForEach-Object { '/LIBPATH:"{0}"' -f $_ }) -join ' '
    return [ordered]@{
        ForceImportBeforeCppTargets = $props
        ForceImportAfterCppTargets = $targets
        VcpkgRoot = $ToolRoot
        VcpkgConfiguration = $(if ($debug) { 'Debug' } else { 'Release' })
        VcpkgManifestInstall = 'false'
        VcpkgXUseBuiltInApplocalDeps = 'false'
        VCPkgLocalAppDataDisabled = 'true'
        VCPKG_DEFAULT_TRIPLET = $triplet
        VCPKG_INSTALLED_DIR = $installed
        VcpkgInstalledDir = $installed.TrimEnd('\') + '\'
        CL = ('/I"{0}" {1}' -f $include,$env:CL).Trim()
        LINK = ($linkFlags + ' ' + $env:LINK).Trim()
        Path = $(if (Test-Path -LiteralPath $bin -PathType Container) { $bin + [IO.Path]::PathSeparator + $env:Path } else { $env:Path })
    }
}

Export-ModuleMember -Function @(
    'Get-RSVcpkgBuildEnvironment',
    'Initialize-RSVcpkgToolContext',
    'Assert-RSVcpkgTripletLeafName',
    'Assert-RSVcpkgManifestVersionConstraintsAvailable',
    'Resolve-RSVcpkgSafeChildPath',
    'Get-RSVcpkgFileSha256Hex',
    'Merge-RSVcpkgTripletSafe',
    'Get-RSVcpkgVisualStudioInstallation',
    'Get-RSVcpkgDefaultToolset',
    'Get-RSVcpkgOverlayTripletText',
    'Resolve-RSVcpkgStockTripletPath',
    'Update-RSVcpkgOverlayTriplet'
)
