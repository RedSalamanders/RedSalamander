# Private build module; import through its reviewed callers.
Set-StrictMode -Version Latest

$artifactLockStateVariable = Get-Variable -Name RSArtifactOperationLocks -Scope Global -ErrorAction SilentlyContinue
if ($null -eq $artifactLockStateVariable) {
    $global:RSArtifactOperationLocks = @{}
}

$artifactLockStackVariable = Get-Variable -Name RSArtifactOperationLockStacks -Scope Global -ErrorAction SilentlyContinue
if ($null -eq $artifactLockStackVariable) {
    $global:RSArtifactOperationLockStacks = @{}
}

function Resolve-RSGuardedRepositoryPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$GuardedRelativeRoot,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$CreateDirectory
    )

    if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
        throw 'A non-empty repository root is required for guarded path resolution.'
    }
    if ([string]::IsNullOrWhiteSpace($GuardedRelativeRoot) -or
        [System.IO.Path]::IsPathRooted($GuardedRelativeRoot)) {
        throw "GuardedRelativeRoot must be a non-empty repository-relative path: '$GuardedRelativeRoot'"
    }
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'A non-empty destination path is required for guarded path resolution.'
    }

    $guardSegments = @($GuardedRelativeRoot -split '[\\/]')
    if ($guardSegments.Count -eq 0 -or
        @($guardSegments | Where-Object { $_ -eq '.' -or $_ -eq '..' -or $_.Contains(':') }).Count -ne 0) {
        throw "GuardedRelativeRoot must name an exact repository-relative directory without traversal or stream segments: '$GuardedRelativeRoot'"
    }

    try {
        $normalizedRepoRoot = [System.IO.Path]::GetFullPath($RepoRoot)
        $repoPathRoot = [System.IO.Path]::GetPathRoot($normalizedRepoRoot)
        if (-not [string]::Equals(
                $normalizedRepoRoot,
                $repoPathRoot,
                [System.StringComparison]::OrdinalIgnoreCase)) {
            $normalizedRepoRoot = $normalizedRepoRoot.TrimEnd([char[]]@('\', '/'))
        }

        $guardedRoot = [System.IO.Path]::GetFullPath(
            (Join-Path $normalizedRepoRoot $GuardedRelativeRoot)).TrimEnd([char[]]@('\', '/'))
        $destination = if ([System.IO.Path]::IsPathRooted($Path)) {
            [System.IO.Path]::GetFullPath($Path)
        }
        else {
            [System.IO.Path]::GetFullPath((Join-Path $normalizedRepoRoot $Path))
        }
        if (-not [string]::Equals(
                $destination,
                [System.IO.Path]::GetPathRoot($destination),
                [System.StringComparison]::OrdinalIgnoreCase)) {
            $destination = $destination.TrimEnd([char[]]@('\', '/'))
        }
    }
    catch [System.ArgumentException] {
        throw "Guarded path resolution rejected an invalid filesystem path: $($_.Exception.Message)"
    }
    catch [System.NotSupportedException] {
        throw "Guarded path resolution rejected an unsupported filesystem path: $($_.Exception.Message)"
    }

    $repoBoundary = $normalizedRepoRoot + [System.IO.Path]::DirectorySeparatorChar
    if (-not [string]::Equals(
            $guardedRoot,
            $normalizedRepoRoot,
            [System.StringComparison]::OrdinalIgnoreCase) -and
        -not $guardedRoot.StartsWith($repoBoundary, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "GuardedRelativeRoot escapes the repository root '$normalizedRepoRoot': $guardedRoot"
    }

    $guardedBoundary = $guardedRoot + [System.IO.Path]::DirectorySeparatorChar
    if (-not [string]::Equals(
            $destination,
            $guardedRoot,
            [System.StringComparison]::OrdinalIgnoreCase) -and
        -not $destination.StartsWith($guardedBoundary, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Destination must remain beneath the guarded repository root '$guardedRoot': $destination"
    }

    $relativeDestination = $destination.Substring($normalizedRepoRoot.Length).TrimStart([char[]]@('\', '/'))
    $destinationSegments = if ([string]::IsNullOrEmpty($relativeDestination)) {
        @()
    }
    else {
        @($relativeDestination -split '[\\/]')
    }
    if (@($destinationSegments | Where-Object { $_.Contains(':') }).Count -ne 0) {
        throw "Destination stream syntax is not allowed beneath guarded repository outputs: $destination"
    }

    $components = [System.Collections.Generic.List[string]]::new()
    $components.Add($normalizedRepoRoot)
    $currentPath = $normalizedRepoRoot
    foreach ($segment in $destinationSegments) {
        $currentPath = Join-Path $currentPath $segment
        $components.Add($currentPath)
    }

    for ($index = 0; $index -lt $components.Count; $index++) {
        $component = $components[$index]
        $attributes = $null
        try {
            $attributes = [System.IO.File]::GetAttributes($component)
        }
        catch [System.IO.FileNotFoundException] {
            if ($index -eq 0) {
                throw "Repository root does not exist for guarded path resolution: '$normalizedRepoRoot'"
            }
            if (-not $CreateDirectory) {
                break
            }
        }
        catch [System.IO.DirectoryNotFoundException] {
            if ($index -eq 0) {
                throw "Repository root does not exist for guarded path resolution: '$normalizedRepoRoot'"
            }
            if (-not $CreateDirectory) {
                break
            }
        }
        catch [System.UnauthorizedAccessException] {
            throw "Unable to verify guarded path component '$component': $($_.Exception.Message)"
        }
        catch [System.IO.IOException] {
            throw "Unable to verify guarded path component '$component': $($_.Exception.Message)"
        }

        if ($null -eq $attributes) {
            [void][System.IO.Directory]::CreateDirectory($component)
            $attributes = [System.IO.File]::GetAttributes($component)
        }

        if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Guarded repository path contains a reparse point at '$component'."
        }

        $mustBeDirectory = $CreateDirectory -or $index -lt ($components.Count - 1)
        if ($mustBeDirectory -and
            ($attributes -band [System.IO.FileAttributes]::Directory) -eq 0) {
            throw "Guarded repository path component is not a directory: '$component'"
        }
    }

    # Revalidate the complete chain after creation so a newly created component
    # cannot silently resolve through a pre-existing or swapped junction.
    foreach ($component in $components) {
        $attributes = $null
        try {
            $attributes = [System.IO.File]::GetAttributes($component)
        }
        catch [System.IO.FileNotFoundException] {
            if ($CreateDirectory) {
                throw "Guarded output directory disappeared during creation: '$component'"
            }
            break
        }
        catch [System.IO.DirectoryNotFoundException] {
            if ($CreateDirectory) {
                throw "Guarded output directory disappeared during creation: '$component'"
            }
            break
        }
        catch [System.UnauthorizedAccessException] {
            throw "Unable to revalidate guarded path component '$component': $($_.Exception.Message)"
        }
        catch [System.IO.IOException] {
            throw "Unable to revalidate guarded path component '$component': $($_.Exception.Message)"
        }

        if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Guarded repository path contains a reparse point at '$component'."
        }
        if ($CreateDirectory -and
            ($attributes -band [System.IO.FileAttributes]::Directory) -eq 0) {
            throw "Guarded output path is not a directory: '$component'"
        }
    }

    return $destination
}

function Publish-RSGuardedRepositoryFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$GuardedRelativeRoot,

        [Parameter(Mandatory = $true)]
        [string]$StagedPath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $staged = Resolve-RSGuardedRepositoryPath `
        -RepoRoot $RepoRoot `
        -GuardedRelativeRoot $GuardedRelativeRoot `
        -Path $StagedPath
    $destination = Resolve-RSGuardedRepositoryPath `
        -RepoRoot $RepoRoot `
        -GuardedRelativeRoot $GuardedRelativeRoot `
        -Path $DestinationPath
    $stagedParent = [IO.Path]::GetFullPath((Split-Path $staged -Parent)).TrimEnd('\')
    $destinationParent = [IO.Path]::GetFullPath((Split-Path $destination -Parent)).TrimEnd('\')
    if (-not [string]::Equals(
            $stagedParent,
            $destinationParent,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw "Guarded atomic publication requires a same-directory staged file: $staged"
    }

    $stagedItem = Get-Item -LiteralPath $staged -Force -ErrorAction Stop
    if ($stagedItem.PSIsContainer -or
        ($stagedItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Guarded atomic publication requires a regular staged file: $staged"
    }
    if (Test-Path -LiteralPath $destination) {
        $destinationItem = Get-Item -LiteralPath $destination -Force -ErrorAction Stop
        if ($destinationItem.PSIsContainer -or
            ($destinationItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Guarded atomic publication refuses an unsafe destination: $destination"
        }
    }

    # Revalidate both complete chains immediately before the atomic filesystem
    # operation. The caller must already hold the lock for its shared output.
    $staged = Resolve-RSGuardedRepositoryPath `
        -RepoRoot $RepoRoot `
        -GuardedRelativeRoot $GuardedRelativeRoot `
        -Path $staged
    $destination = Resolve-RSGuardedRepositoryPath `
        -RepoRoot $RepoRoot `
        -GuardedRelativeRoot $GuardedRelativeRoot `
        -Path $destination
    if ([IO.File]::Exists($destination)) {
        $backup = Join-Path $destinationParent `
            ('.{0}.{1}.bak' -f [IO.Path]::GetFileName($destination), [Guid]::NewGuid().ToString('N'))
        $backup = Resolve-RSGuardedRepositoryPath `
            -RepoRoot $RepoRoot `
            -GuardedRelativeRoot $GuardedRelativeRoot `
            -Path $backup
        [IO.File]::Replace($staged, $destination, $backup)
        Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue
    }
    else {
        [IO.File]::Move($staged, $destination)
    }
    return $destination
}

function Add-RSArtifactOperationThreadState {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ManagedThreadId,

        [Parameter(Mandatory = $true)]
        [string]$StateKey
    )

    $threadKey = $ManagedThreadId.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    if (-not $global:RSArtifactOperationLockStacks.ContainsKey($threadKey)) {
        $global:RSArtifactOperationLockStacks[$threadKey] = [System.Collections.Generic.List[string]]::new()
    }

    $global:RSArtifactOperationLockStacks[$threadKey].Add($StateKey)
}

function Assert-RSArtifactOperationThreadStateIsCurrent {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ManagedThreadId,

        [Parameter(Mandatory = $true)]
        [string]$StateKey
    )

    $threadKey = $ManagedThreadId.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    if (-not $global:RSArtifactOperationLockStacks.ContainsKey($threadKey)) {
        throw "Artifact-operation state '$StateKey' is missing from its owning thread's lock stack."
    }

    $stack = $global:RSArtifactOperationLockStacks[$threadKey]
    if ($stack.Count -eq 0 -or
        -not [string]::Equals($stack[$stack.Count - 1], $StateKey, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Distinct artifact-operation locks must be released in same-thread LIFO order so inherited delegation markers remain scoped to the current lock."
    }
}

function Remove-RSArtifactOperationThreadState {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ManagedThreadId,

        [Parameter(Mandatory = $true)]
        [string]$StateKey
    )

    $threadKey = $ManagedThreadId.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    Assert-RSArtifactOperationThreadStateIsCurrent -ManagedThreadId $ManagedThreadId -StateKey $StateKey
    $stack = $global:RSArtifactOperationLockStacks[$threadKey]
    $stack.RemoveAt($stack.Count - 1)
    if ($stack.Count -eq 0) {
        [void]$global:RSArtifactOperationLockStacks.Remove($threadKey)
    }
}

function Get-RSArtifactOperationScopeInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $normalizedRoot = [System.IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $configuration = if ($Scope.ContainsKey('configuration')) {
        ([string]$Scope.configuration).Trim()
    } else { '' }
    $platform = if ($Scope.ContainsKey('platform')) {
        ([string]$Scope.platform).Trim()
    } else { '' }
    $safeSegmentPattern = '^[A-Za-z0-9][A-Za-z0-9 ._-]*$'
    $isProfileScoped = -not [string]::IsNullOrWhiteSpace($configuration) -and
        -not [string]::IsNullOrWhiteSpace($platform) -and
        $configuration -match $safeSegmentPattern -and
        $platform -match $safeSegmentPattern
    $coordination = if ($Scope.ContainsKey('coordination')) {
        ([string]$Scope.coordination).Trim()
    } else { 'artifact-profile' }
    if ($coordination -notin @('artifact-profile', 'shared-dependency', 'packaging')) {
        throw "Scope.coordination must be artifact-profile, shared-dependency, or packaging; received '$coordination'."
    }
    $isSharedDependencyScope = $isProfileScoped -and
        [string]::Equals($coordination, 'shared-dependency', [System.StringComparison]::OrdinalIgnoreCase)
    $isPackagingScope = $isProfileScoped -and
        [string]::Equals($coordination, 'packaging', [System.StringComparison]::OrdinalIgnoreCase)

    $normalizedScope = [ordered]@{}
    foreach ($scopeName in @('kind', 'target', 'configuration', 'platform', 'suite', 'coordination')) {
        if ($Scope.ContainsKey($scopeName) -and
            -not [string]::IsNullOrWhiteSpace([string]$Scope[$scopeName])) {
            $normalizedScope[$scopeName] = ([string]$Scope[$scopeName]).Trim()
        }
    }

    $scopeKey = 'repository'
    $fileStem = 'repository'
    $artifactRoots = @([System.IO.Path]::GetFullPath((Join-Path $normalizedRoot '.build')))
    if ($isProfileScoped) {
        $scopeKey = if ($isPackagingScope) {
            'packaging'
        } elseif ($isSharedDependencyScope) {
            # One platform-scoped vcpkg install root is shared by every configuration.
            'shared-dependency|platform={0}' -f $platform.ToUpperInvariant()
        } else {
            'profile|platform={0}|configuration={1}' -f
                $platform.ToUpperInvariant(),
                $configuration.ToUpperInvariant()
        }
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        try {
            $hashBytes = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($scopeKey))
        }
        finally {
            $sha256.Dispose()
        }
        $hash = ([System.BitConverter]::ToString($hashBytes) -replace '-', '').Substring(0, 12).ToLowerInvariant()
        $platformSlug = ($platform.ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-')
        $configurationSlug = ($configuration.ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-')
        $triplet = if ([string]::Equals($platform, 'ARM64', [System.StringComparison]::OrdinalIgnoreCase)) {
            'arm64-windows'
        } else {
            "$($platform.ToLowerInvariant())-windows"
        }
        if ($isPackagingScope) {
            $fileStem = "packaging-$hash"
            $artifactRoots = @(
                [System.IO.Path]::GetFullPath((Join-Path $normalizedRoot '.build\AppPackages')),
                [System.IO.Path]::GetFullPath((Join-Path $normalizedRoot 'Installer\msix')),
                [System.IO.Path]::GetFullPath((Join-Path $normalizedRoot 'Installer\winget'))
            )
        }
        elseif ($isSharedDependencyScope) {
            $fileStem = "shared-dependency-$platformSlug-$hash"
            $artifactRoots = @(
                [System.IO.Path]::GetFullPath((Join-Path $normalizedRoot ('.build\vcpkg_installed\{0}' -f $platformSlug))),
                [System.IO.Path]::GetFullPath((Join-Path $normalizedRoot ('.build\vcpkg_install_staging\{0}' -f $triplet)))
            )
        }
        else {
            $fileStem = "$platformSlug-$configurationSlug-$hash"
            $artifactRoots = @(
                [System.IO.Path]::GetFullPath((Join-Path $normalizedRoot ('.build\{0}\{1}' -f $platform, $configuration))),
                [System.IO.Path]::GetFullPath((Join-Path $normalizedRoot ('.build\Intermediate\{0}\{1}' -f $platform, $configuration)))
            )
        }
    }

    return [pscustomobject]@{
        RepoRoot = $normalizedRoot
        Key = $scopeKey
        FileStem = $fileStem
        IsProfileScoped = $isProfileScoped
        Configuration = $configuration
        Platform = $platform
        Coordination = $coordination
        Triplet = if ($isProfileScoped) { $triplet } else { '' }
        IsSharedDependencyScope = $isSharedDependencyScope
        IsPackagingScope = $isPackagingScope
        Scope = $normalizedScope
        ArtifactRoots = $artifactRoots
    }
}

function Get-RSArtifactOperationLockPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $scopeInfo = Get-RSArtifactOperationScopeInfo -RepoRoot $RepoRoot -Scope $Scope
    if (-not $scopeInfo.IsProfileScoped) {
        return Join-Path $scopeInfo.RepoRoot '.build\artifact-operation.lock'
    }

    return Join-Path $scopeInfo.RepoRoot ".build\artifact-operations\$($scopeInfo.FileStem).lock"
}

function Get-RSArtifactOperationStateKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LockPath
    )

    return [System.IO.Path]::GetFullPath($LockPath).ToUpperInvariant()
}

function Get-RSArtifactOperationOwnerMetadataPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $scopeInfo = Get-RSArtifactOperationScopeInfo -RepoRoot $RepoRoot -Scope $Scope
    if (-not $scopeInfo.IsProfileScoped) {
        return Join-Path $scopeInfo.RepoRoot '.build\artifact-operation-owner.json'
    }

    return Join-Path $scopeInfo.RepoRoot ".build\artifact-operations\$($scopeInfo.FileStem).owner.json"
}

function Read-RSArtifactOperationOwnerMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $RepoRoot -Scope $Scope
    if (-not (Test-Path -LiteralPath $metadataPath)) {
        return $null
    }

    try {
        $metadata = Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
        $requiredProperties = @(
            'schema',
            'lock_path',
            'repo_root',
            'owner_pid',
            'owner_process_start_utc_ticks',
            'operation',
            'started_utc',
            'token'
        )
        foreach ($propertyName in $requiredProperties) {
            if ($null -eq $metadata.PSObject.Properties[$propertyName]) {
                return $null
            }
        }

        $schema = [string]$metadata.schema
        if ($schema -notin @(
                'red-salamander.artifact-operation-owner.v2',
                'red-salamander.artifact-operation-owner.v3',
                'red-salamander.artifact-operation-owner.v4')) {
            return $null
        }

        if ($schema -in @('red-salamander.artifact-operation-owner.v3', 'red-salamander.artifact-operation-owner.v4') -and
            $null -eq $metadata.PSObject.Properties['scope']) {
            return $null
        }
        if ($schema -eq 'red-salamander.artifact-operation-owner.v4' -and
            ($null -eq $metadata.PSObject.Properties['scope_key'] -or
             $null -eq $metadata.PSObject.Properties['owner_process_path'] -or
             [string]::IsNullOrWhiteSpace([string]$metadata.scope_key) -or
             [string]::IsNullOrWhiteSpace([string]$metadata.owner_process_path))) {
            return $null
        }

        $ownerPid = 0
        $ownerProcessStartUtcTicks = 0L
        if (-not [int]::TryParse([string]$metadata.owner_pid, [ref]$ownerPid) -or $ownerPid -le 0 -or
            -not [long]::TryParse([string]$metadata.owner_process_start_utc_ticks, [ref]$ownerProcessStartUtcTicks) -or
            $ownerProcessStartUtcTicks -le 0 -or
            [string]::IsNullOrWhiteSpace([string]$metadata.lock_path) -or
            [string]::IsNullOrWhiteSpace([string]$metadata.operation) -or
            [string]::IsNullOrWhiteSpace([string]$metadata.started_utc) -or
            [string]::IsNullOrWhiteSpace([string]$metadata.token)) {
            return $null
        }

        return $metadata
    }
    catch {
        return $null
    }
}

function Write-RSArtifactOperationOwnerMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$LockPath,

        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$Token,

        [Parameter(Mandatory = $true)]
        [string]$StartedUtc,

        [Parameter(Mandatory = $true)]
        [long]$OwnerProcessStartUtcTicks,

        [Parameter(Mandatory = $true)]
        [string]$OwnerProcessPath,

        [hashtable]$Scope = @{}
    )

    $scopeInfo = Get-RSArtifactOperationScopeInfo -RepoRoot $RepoRoot -Scope $Scope
    $normalizedRoot = $scopeInfo.RepoRoot
    $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $normalizedRoot -Scope $Scope
    $metadataDirectory = Split-Path -Parent $metadataPath
    [void](New-Item -ItemType Directory -Path $metadataDirectory -Force)
    $temporaryPath = "$metadataPath.$PID.$([Guid]::NewGuid().ToString('N')).tmp"
    $backupPath = "$metadataPath.$PID.$([Guid]::NewGuid().ToString('N')).bak"
    $payload = [ordered]@{
        schema = 'red-salamander.artifact-operation-owner.v4'
        lock_path = [System.IO.Path]::GetFullPath($LockPath)
        repo_root = $normalizedRoot
        scope_key = $scopeInfo.Key
        owner_pid = $PID
        owner_process_start_utc_ticks = $OwnerProcessStartUtcTicks
        owner_process_path = [System.IO.Path]::GetFullPath($OwnerProcessPath)
        operation = $Operation
        started_utc = $StartedUtc
        token = $Token
        scope = $scopeInfo.Scope
    }

    try {
        $json = $payload | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($temporaryPath, $json, [System.Text.UTF8Encoding]::new($false))
        if (Test-Path -LiteralPath $metadataPath) {
            [System.IO.File]::Replace($temporaryPath, $metadataPath, $backupPath)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $metadataPath)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $backupPath) {
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        }
    }

    return $metadataPath
}

function Remove-RSArtifactOperationOwnerMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Token,

        [hashtable]$Scope = @{}
    )

    $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $RepoRoot -Scope $Scope
    if (-not (Test-Path -LiteralPath $metadataPath)) {
        return $false
    }

    $metadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $RepoRoot -Scope $Scope
    if ($null -eq $metadata -or
        -not [string]::Equals([string]$metadata.token, $Token, [System.StringComparison]::Ordinal)) {
        return $false
    }

    Remove-Item -LiteralPath $metadataPath -Force
    return $true
}

function Set-RSArtifactOperationLockMarker {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileStream]$LockStream,

        [Parameter(Mandatory = $true)]
        [string]$Token
    )

    $markerBytes = [System.Text.Encoding]::UTF8.GetBytes(
        "red-salamander.artifact-operation-lock.v1`ntoken=$Token`n")
    $LockStream.SetLength(0)
    $LockStream.Position = 0
    $LockStream.Write($markerBytes, 0, $markerBytes.Length)
    $LockStream.Flush($true)
}

function Clear-RSArtifactOperationLockMarker {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileStream]$LockStream
    )

    $LockStream.SetLength(0)
    $LockStream.Flush($true)
}

function Get-RSImmediateParentProcessId {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ProcessId
    )

    $containedProcessType = 'RedSalamander.Tooling.KillOnCloseJob' -as [type]
    if ($ProcessId -eq $PID -and $null -ne $containedProcessType) {
        try {
            return [RedSalamander.Tooling.KillOnCloseJob]::GetCurrentParentProcessId()
        }
        catch [System.ComponentModel.Win32Exception] {
            # Fall back to the existing read-only CIM query below.
        }
    }

    try {
        $process = @(Get-CimInstance Win32_Process -Filter "ProcessId=$ProcessId" -ErrorAction Stop |
                Select-Object -First 1)
        if ($process.Count -eq 0) {
            return 0
        }

        return [int]$process[0].ParentProcessId
    }
    catch {
        return 0
    }
}

function Test-RSArtifactOperationOwnerProcess {
    param(
        [Parameter(Mandatory = $true)]
        [int]$OwnerProcessId,

        [Parameter(Mandatory = $true)]
        [long]$OwnerProcessStartUtcTicks,

        [Parameter(Mandatory = $true)]
        [string]$OwnerProcessPath
    )

    $ownerProcess = $null
    try {
        $ownerProcess = Get-Process -Id $OwnerProcessId -ErrorAction Stop
        if ($ownerProcess.StartTime.ToUniversalTime().Ticks -ne $OwnerProcessStartUtcTicks -or
            [string]::IsNullOrWhiteSpace([string]$ownerProcess.Path)) {
            return $false
        }

        return [string]::Equals(
            [System.IO.Path]::GetFullPath([string]$ownerProcess.Path),
            [System.IO.Path]::GetFullPath($OwnerProcessPath),
            [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $ownerProcess) {
            $ownerProcess.Dispose()
        }
    }
}

function Get-RSDelegatedArtifactOperationOwner {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$LockPath,

        [hashtable]$Scope = @{}
    )

    $token = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', 'Process')
    $ownerPidText = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', 'Process')
    $inheritedLockPath = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', 'Process')
    $delegationToken = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_DELEGATION_TOKEN', 'Process')
    $delegationParentPidText = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_DELEGATION_PARENT_PID', 'Process')
    $delegationReadyEvent = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_DELEGATION_READY_EVENT', 'Process')
    if ([string]::IsNullOrWhiteSpace($token) -or
        [string]::IsNullOrWhiteSpace($ownerPidText) -or
        [string]::IsNullOrWhiteSpace($inheritedLockPath) -or
        [string]::IsNullOrWhiteSpace($delegationToken) -or
        [string]::IsNullOrWhiteSpace($delegationParentPidText) -or
        [string]::IsNullOrWhiteSpace($delegationReadyEvent) -or
        -not [string]::Equals(
            $delegationReadyEvent,
            "Local\RedSalamander.ArtifactDelegation.$delegationToken",
            [System.StringComparison]::Ordinal) -or
        -not [string]::Equals(
            [System.IO.Path]::GetFullPath($inheritedLockPath),
            [System.IO.Path]::GetFullPath($LockPath),
            [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    $ownerPid = 0
    $delegationParentPid = 0
    if (-not [int]::TryParse($ownerPidText, [ref]$ownerPid) -or
        -not [int]::TryParse($delegationParentPidText, [ref]$delegationParentPid) -or
        $ownerPid -le 0 -or
        $ownerPid -eq $PID -or
        $delegationParentPid -ne $ownerPid -or
        (Get-RSImmediateParentProcessId -ProcessId $PID) -ne $ownerPid) {
        return $null
    }

    $metadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $RepoRoot -Scope $Scope
    if ($null -eq $metadata -or
        $null -eq $metadata.PSObject.Properties['owner_process_path'] -or
        -not [string]::Equals([string]$metadata.token, $token, [System.StringComparison]::Ordinal) -or
        -not [string]::Equals(
            [System.IO.Path]::GetFullPath([string]$metadata.lock_path),
            [System.IO.Path]::GetFullPath($LockPath),
            [System.StringComparison]::OrdinalIgnoreCase) -or
        [int]$metadata.owner_pid -ne $ownerPid -or
        -not (Test-RSArtifactOperationOwnerProcess `
            -OwnerProcessId $ownerPid `
            -OwnerProcessStartUtcTicks ([long]$metadata.owner_process_start_utc_ticks) `
            -OwnerProcessPath ([string]$metadata.owner_process_path))) {
        return $null
    }

    $readyEvent = $null
    try {
        $readyEvent = [System.Threading.EventWaitHandle]::OpenExisting($delegationReadyEvent)
        if (-not $readyEvent.WaitOne(5000)) {
            return $null
        }
    }
    catch {
        return $null
    }
    finally {
        if ($null -ne $readyEvent) {
            $readyEvent.Dispose()
        }
    }

    $jobType = 'RedSalamander.Tooling.KillOnCloseJob' -as [type]
    if ($null -eq $jobType -or -not [RedSalamander.Tooling.KillOnCloseJob]::IsCurrentProcessInJob()) {
        return $null
    }

    if (-not (Test-RSArtifactOperationOwnerProcess `
            -OwnerProcessId $ownerPid `
            -OwnerProcessStartUtcTicks ([long]$metadata.owner_process_start_utc_ticks) `
            -OwnerProcessPath ([string]$metadata.owner_process_path))) {
        return $null
    }

    foreach ($environmentName in @(
            'REDSALAMANDER_ARTIFACT_DELEGATION_TOKEN',
            'REDSALAMANDER_ARTIFACT_DELEGATION_PARENT_PID',
            'REDSALAMANDER_ARTIFACT_DELEGATION_READY_EVENT')) {
        [Environment]::SetEnvironmentVariable($environmentName, $null, 'Process')
    }

    return $metadata
}

function Format-RSArtifactOperationOwnerDiagnostic {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $RepoRoot -Scope $Scope
    $metadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $RepoRoot -Scope $Scope
    if ($null -eq $metadata) {
        return "Owner metadata is unavailable or malformed at '$metadataPath'."
    }

    if ($null -eq $metadata.PSObject.Properties['owner_process_path'] -or
        -not (Test-RSArtifactOperationOwnerProcess `
            -OwnerProcessId ([int]$metadata.owner_pid) `
            -OwnerProcessStartUtcTicks ([long]$metadata.owner_process_start_utc_ticks) `
            -OwnerProcessPath ([string]$metadata.owner_process_path))) {
        return "Owner metadata at '$metadataPath' does not identify the current PID instance and executable path; it is stale or untrusted."
    }

    return "Owner PID=$($metadata.owner_pid); Path='$($metadata.owner_process_path)'; Operation='$($metadata.operation)'; StartedUtc='$($metadata.started_utc)'."
}

function Open-RSArtifactOperationLockStream {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LockPath,

        [ValidateRange(0, 60000)]
        [int]$TimeoutMilliseconds = 0
    )

    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
    do {
        try {
            return [System.IO.FileStream]::new(
                $LockPath,
                [System.IO.FileMode]::OpenOrCreate,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None)
        }
        catch [System.IO.IOException] {
            $nativeError = $_.Exception.HResult -band 0xFFFF
            if ($nativeError -notin @(32, 33)) {
                throw
            }

            if ($TimeoutMilliseconds -eq 0 -or [DateTime]::UtcNow -ge $deadline) {
                return $null
            }

            Start-Sleep -Milliseconds 25
        }
    } while ($true)
}

function Enter-RSArtifactOperationLock {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [hashtable]$Scope = @{},

        [switch]$AllowRepositoryScope,

        [ValidateRange(0, 60000)]
        [int]$TimeoutMilliseconds = 0
    )

    $scopeInfo = Get-RSArtifactOperationScopeInfo -RepoRoot $RepoRoot -Scope $Scope
    if (-not $scopeInfo.IsProfileScoped -and -not $AllowRepositoryScope) {
        throw "Artifact-mutating operations must provide non-empty, path-safe Scope.configuration and Scope.platform values. Repository-wide fallback requires the explicit -AllowRepositoryScope compatibility switch."
    }
    $normalizedRoot = $scopeInfo.RepoRoot
    $lockPath = Get-RSArtifactOperationLockPath -RepoRoot $normalizedRoot -Scope $Scope
    $lockDirectory = Split-Path -Parent $lockPath
    [void](New-Item -ItemType Directory -Path $lockDirectory -Force)
    $stateKey = Get-RSArtifactOperationStateKey -LockPath $lockPath
    $managedThreadId = [Environment]::CurrentManagedThreadId
    if ($global:RSArtifactOperationLocks.ContainsKey($stateKey)) {
        $state = $global:RSArtifactOperationLocks[$stateKey]
        if ([int]$state.OwnerManagedThreadId -eq $managedThreadId) {
            $leaseId = [Guid]::NewGuid().ToString('N')
            $state.Leases[$leaseId] = $true
            return [pscustomobject]@{
                Name = $lockPath
                StateKey = $stateKey
                LeaseId = $leaseId
                Operation = $Operation
                WasAbandoned = $false
                AbandonedOwner = $null
                IsDelegated = [bool]$state.IsDelegatedOwner
                IsRootOwner = $false
            }
        }
    }

    $delegatedOwner = Get-RSDelegatedArtifactOperationOwner `
        -RepoRoot $normalizedRoot `
        -LockPath $lockPath `
        -Scope $Scope
    if ($null -ne $delegatedOwner) {
        $leaseId = [Guid]::NewGuid().ToString('N')
        $leases = @{}
        $leases[$leaseId] = $true
        $global:RSArtifactOperationLocks[$stateKey] = [pscustomobject]@{
            LockStream = $null
            LockPath = $lockPath
            RepoRoot = $normalizedRoot
            Scope = $Scope.Clone()
            ScopeKey = $scopeInfo.Key
            Token = [string]$delegatedOwner.token
            OwnerManagedThreadId = $managedThreadId
            Leases = $leases
            ActiveDelegations = @{}
            IsDelegatedOwner = $true
            PreviousToken = $null
            PreviousOwnerPid = $null
            PreviousLockPath = $null
        }
        Add-RSArtifactOperationThreadState -ManagedThreadId $managedThreadId -StateKey $stateKey
        return [pscustomobject]@{
            Name = $lockPath
            StateKey = $stateKey
            LeaseId = $leaseId
            Operation = $Operation
            WasAbandoned = $false
            AbandonedOwner = $null
            IsDelegated = $true
            IsRootOwner = $false
        }
    }

    $lockStream = Open-RSArtifactOperationLockStream `
        -LockPath $lockPath `
        -TimeoutMilliseconds $TimeoutMilliseconds
    if ($null -eq $lockStream) {
        $ownerDiagnostic = Format-RSArtifactOperationOwnerDiagnostic -RepoRoot $normalizedRoot -Scope $Scope
        throw "Cannot start '$Operation' because another RedSalamander build or test operation owns artifact scope '$($scopeInfo.Key)' at '$lockPath'. $ownerDiagnostic Wait for it to finish and retry."
    }

    $abandonedOwner = Read-RSArtifactOperationOwnerMetadata -RepoRoot $normalizedRoot -Scope $Scope
    $wasAbandoned = $lockStream.Length -gt 0 -or $null -ne $abandonedOwner -or
        (Test-Path -LiteralPath (Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $normalizedRoot -Scope $Scope))
    $token = [Guid]::NewGuid().ToString('N')
    $leaseId = [Guid]::NewGuid().ToString('N')
    $startedUtc = [DateTime]::UtcNow.ToString('o')
    $currentProcess = [System.Diagnostics.Process]::GetCurrentProcess()
    try {
        $ownerProcessStartUtcTicks = $currentProcess.StartTime.ToUniversalTime().Ticks
        $ownerProcessPath = [System.IO.Path]::GetFullPath([string]$currentProcess.Path)
    }
    finally {
        $currentProcess.Dispose()
    }

    $previousToken = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', 'Process')
    $previousOwnerPid = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', 'Process')
    $previousLockPath = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', 'Process')
    try {
        Set-RSArtifactOperationLockMarker -LockStream $lockStream -Token $token
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', $token, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', $PID.ToString(), 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', $lockPath, 'Process')
        [void](Write-RSArtifactOperationOwnerMetadata `
                -RepoRoot $normalizedRoot `
                -LockPath $lockPath `
                -Operation $Operation `
                -Token $token `
                -StartedUtc $startedUtc `
                -OwnerProcessStartUtcTicks $ownerProcessStartUtcTicks `
                -OwnerProcessPath $ownerProcessPath `
                -Scope $Scope)

        $leases = @{}
        $leases[$leaseId] = $true
        $global:RSArtifactOperationLocks[$stateKey] = [pscustomobject]@{
            LockStream = $lockStream
            LockPath = $lockPath
            RepoRoot = $normalizedRoot
            Scope = $Scope.Clone()
            ScopeKey = $scopeInfo.Key
            Token = $token
            OwnerManagedThreadId = $managedThreadId
            Leases = $leases
            ActiveDelegations = @{}
            IsDelegatedOwner = $false
            PreviousToken = $previousToken
            PreviousOwnerPid = $previousOwnerPid
            PreviousLockPath = $previousLockPath
        }
        Add-RSArtifactOperationThreadState -ManagedThreadId $managedThreadId -StateKey $stateKey

        return [pscustomobject]@{
            Name = $lockPath
            StateKey = $stateKey
            LeaseId = $leaseId
            Operation = $Operation
            WasAbandoned = $wasAbandoned
            AbandonedOwner = $abandonedOwner
            IsDelegated = $false
            IsRootOwner = $true
        }
    }
    catch {
        [void]$global:RSArtifactOperationLocks.Remove($stateKey)
        [void](Remove-RSArtifactOperationOwnerMetadata -RepoRoot $normalizedRoot -Token $token -Scope $Scope)
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', $previousToken, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', $previousOwnerPid, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', $previousLockPath, 'Process')
        try {
            Clear-RSArtifactOperationLockMarker -LockStream $lockStream
        }
        finally {
            $lockStream.Dispose()
        }
        throw
    }
}

function New-RSArtifactOperationChildDelegation {
    $token = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', 'Process')
    $lockPath = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', 'Process')
    if ([string]::IsNullOrWhiteSpace($token) -or [string]::IsNullOrWhiteSpace($lockPath)) {
        return $null
    }

    $stateKey = Get-RSArtifactOperationStateKey -LockPath $lockPath
    if (-not $global:RSArtifactOperationLocks.ContainsKey($stateKey)) {
        return $null
    }

    $state = $global:RSArtifactOperationLocks[$stateKey]
    if (-not [string]::Equals([string]$state.Token, $token, [System.StringComparison]::Ordinal) -or
        [int]$state.OwnerManagedThreadId -ne [Environment]::CurrentManagedThreadId -or
        [bool]$state.IsDelegatedOwner) {
        # A delegated child owns only a local, reentrant lease. It cannot delegate
        # the parent's authority to a second process generation.
        return $null
    }

    $delegationId = [Guid]::NewGuid().ToString('N')
    $readyEventName = "Local\RedSalamander.ArtifactDelegation.$delegationId"
    $readyEvent = [System.Threading.EventWaitHandle]::new(
        $false,
        [System.Threading.EventResetMode]::ManualReset,
        $readyEventName)
    $delegation = [pscustomobject]@{
        Id = $delegationId
        StateKey = $stateKey
        Token = $token
        ReadyEvent = $readyEvent
        Environment = [ordered]@{
            REDSALAMANDER_ARTIFACT_DELEGATION_TOKEN = $delegationId
            REDSALAMANDER_ARTIFACT_DELEGATION_PARENT_PID = $PID.ToString()
            REDSALAMANDER_ARTIFACT_DELEGATION_READY_EVENT = $readyEventName
        }
    }
    $state.ActiveDelegations[$delegationId] = $delegation
    return $delegation
}

function Complete-RSArtifactOperationChildDelegation {
    param(
        [AllowNull()]
        [object]$Delegation
    )

    if ($null -eq $Delegation) {
        return
    }

    $stateKey = [string]$Delegation.StateKey
    if ($global:RSArtifactOperationLocks.ContainsKey($stateKey)) {
        $state = $global:RSArtifactOperationLocks[$stateKey]
        if ([string]::Equals([string]$state.Token, [string]$Delegation.Token, [System.StringComparison]::Ordinal)) {
            [void]$state.ActiveDelegations.Remove([string]$Delegation.Id)
        }
    }

    $Delegation.ReadyEvent.Dispose()
}

function Exit-RSArtifactOperationLock {
    param(
        [AllowNull()]
        [object]$Lock
    )

    if ($null -eq $Lock) {
        return
    }

    $stateKey = [string]$Lock.StateKey
    if (-not $global:RSArtifactOperationLocks.ContainsKey($stateKey)) {
        return
    }

    $state = $global:RSArtifactOperationLocks[$stateKey]
    if ([int]$state.OwnerManagedThreadId -ne [Environment]::CurrentManagedThreadId) {
        throw "Artifact-operation lease '$($Lock.LeaseId)' must be released on its owning managed thread."
    }

    $leaseId = [string]$Lock.LeaseId
    if (-not $state.Leases.ContainsKey($leaseId)) {
        return
    }

    if ($state.Leases.Count -eq 1 -and $state.ActiveDelegations.Count -gt 0) {
        throw "Cannot release the root artifact-operation lease while $($state.ActiveDelegations.Count) contained child delegation(s) remain active."
    }

    if ($state.Leases.Count -eq 1) {
        Assert-RSArtifactOperationThreadStateIsCurrent `
            -ManagedThreadId ([int]$state.OwnerManagedThreadId) `
            -StateKey $stateKey
    }

    [void]$state.Leases.Remove($leaseId)
    if ($state.Leases.Count -gt 0) {
        return
    }

    if ([bool]$state.IsDelegatedOwner) {
        foreach ($environmentName in @(
                'REDSALAMANDER_ARTIFACT_OPERATION_TOKEN',
                'REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID',
                'REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH')) {
            [Environment]::SetEnvironmentVariable($environmentName, $null, 'Process')
        }
        Remove-RSArtifactOperationThreadState `
            -ManagedThreadId ([int]$state.OwnerManagedThreadId) `
            -StateKey $stateKey
        [void]$global:RSArtifactOperationLocks.Remove($stateKey)
        return
    }

    try {
        [void](Remove-RSArtifactOperationOwnerMetadata `
                -RepoRoot $state.RepoRoot `
                -Token $state.Token `
                -Scope $state.Scope)
        Clear-RSArtifactOperationLockMarker -LockStream $state.LockStream
    }
    finally {
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', $state.PreviousToken, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', $state.PreviousOwnerPid, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', $state.PreviousLockPath, 'Process')
        $state.LockStream.Dispose()
        Remove-RSArtifactOperationThreadState `
            -ManagedThreadId ([int]$state.OwnerManagedThreadId) `
            -StateKey $stateKey
        [void]$global:RSArtifactOperationLocks.Remove($stateKey)
    }
}

function Get-RSArtifactContaminationMarkerPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $scopeInfo = Get-RSArtifactOperationScopeInfo -RepoRoot $RepoRoot -Scope $Scope
    if (-not $scopeInfo.IsProfileScoped) {
        return Join-Path $scopeInfo.RepoRoot '.build\artifact-operation-contaminated.json'
    }

    return Join-Path $scopeInfo.RepoRoot ".build\artifact-operations\$($scopeInfo.FileStem).contaminated.json"
}

function Set-RSArtifactOperationContaminated {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Reason,

        [AllowNull()]
        [object]$AbandonedOwner = $null,

        [hashtable]$Scope = @{}
    )

    $scopeInfo = Get-RSArtifactOperationScopeInfo -RepoRoot $RepoRoot -Scope $Scope
    $markerPath = Get-RSArtifactContaminationMarkerPath -RepoRoot $RepoRoot -Scope $Scope
    $markerDirectory = Split-Path -Parent $markerPath
    [void](New-Item -ItemType Directory -Path $markerDirectory -Force)

    $abandonedOperation = ''
    $abandonedStartedUtc = ''
    $abandonedScope = [ordered]@{}
    if ($null -ne $AbandonedOwner) {
        if ($null -ne $AbandonedOwner.PSObject.Properties['operation']) {
            $abandonedOperation = [string]$AbandonedOwner.operation
        }
        if ($null -ne $AbandonedOwner.PSObject.Properties['started_utc']) {
            $abandonedStartedUtc = [string]$AbandonedOwner.started_utc
        }
        if ($null -ne $AbandonedOwner.PSObject.Properties['scope'] -and
            $null -ne $AbandonedOwner.scope) {
            foreach ($scopeName in @('kind', 'target', 'configuration', 'platform', 'suite', 'coordination')) {
                if ($null -ne $AbandonedOwner.scope.PSObject.Properties[$scopeName] -and
                    -not [string]::IsNullOrWhiteSpace([string]$AbandonedOwner.scope.$scopeName)) {
                    $abandonedScope[$scopeName] = [string]$AbandonedOwner.scope.$scopeName
                }
            }
        }
    }

    $payload = [ordered]@{
        schema = 'red-salamander.artifact-operation-contamination.v3'
        recorded_utc = [DateTime]::UtcNow.ToString('o')
        process_id = $PID
        reason = $Reason
        scope_key = $scopeInfo.Key
        scope = $scopeInfo.Scope
        abandoned_operation = $abandonedOperation
        abandoned_started_utc = $abandonedStartedUtc
        abandoned_scope = $abandonedScope
    }
    $json = $payload | ConvertTo-Json -Depth 4
    [System.IO.File]::WriteAllText($markerPath, $json, [System.Text.UTF8Encoding]::new($false))
    return $markerPath
}

function Read-RSArtifactOperationContamination {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $markerPath = Get-RSArtifactContaminationMarkerPath -RepoRoot $RepoRoot -Scope $Scope
    if (-not (Test-Path -LiteralPath $markerPath)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $markerPath -Raw | ConvertFrom-Json
    }
    catch {
        return [pscustomobject]@{
            schema = 'red-salamander.artifact-operation-contamination.unknown'
            reason = 'The contamination marker is malformed.'
            abandoned_operation = ''
            abandoned_started_utc = ''
            abandoned_scope = [pscustomobject]@{}
        }
    }
}

function Test-RSArtifactOperationRepairAllowed {
    param(
        [AllowNull()]
        [object]$Contamination,

        [switch]$Rebuild,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$ProjectName = '',

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$Platform
    )

    if ($null -eq $Contamination -or -not $Rebuild -or
        -not [string]::IsNullOrWhiteSpace($ProjectName)) {
        return $false
    }

    $scopeProperty = $Contamination.PSObject.Properties['abandoned_scope']
    if ($null -eq $scopeProperty -or $null -eq $scopeProperty.Value) {
        # Legacy markers did not record scope. A full-solution rebuild is the only
        # safe automated recovery available for those markers.
        return $true
    }

    $scope = $scopeProperty.Value
    $recordedConfiguration = if ($null -ne $scope.PSObject.Properties['configuration']) {
        [string]$scope.configuration
    } else { '' }
    $recordedPlatform = if ($null -ne $scope.PSObject.Properties['platform']) {
        [string]$scope.platform
    } else { '' }
    $recordedCoordination = if ($null -ne $scope.PSObject.Properties['coordination']) {
        [string]$scope.coordination
    } else { '' }

    if (-not [string]::Equals($recordedCoordination, 'shared-dependency', [System.StringComparison]::OrdinalIgnoreCase) -and
        -not [string]::IsNullOrWhiteSpace($recordedConfiguration) -and
        -not [string]::Equals($recordedConfiguration, $Configuration, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }
    if (-not [string]::IsNullOrWhiteSpace($recordedPlatform) -and
        -not [string]::Equals($recordedPlatform, $Platform, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    return $true
}

function Test-RSArtifactOperationContaminated {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    return Test-Path -LiteralPath (Get-RSArtifactContaminationMarkerPath -RepoRoot $RepoRoot -Scope $Scope)
}

function Clear-RSArtifactOperationContaminated {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $markerPath = Get-RSArtifactContaminationMarkerPath -RepoRoot $RepoRoot -Scope $Scope
    if (Test-Path -LiteralPath $markerPath) {
        Remove-Item -LiteralPath $markerPath -Force
    }
}

function Test-RSArtifactCommandLineContainsPath {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$CommandLine,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($CommandLine)) {
        return $false
    }

    $normalizedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    $pathPattern = [regex]::Escape($normalizedPath).Replace('\\', '[\\/]')
    $leftBoundary = '(?<![A-Za-z0-9_:\.\\/-])'
    $rightBoundary = '(?=$|[\s"'';,\)\\/])'
    $attachedPathSwitch = '(?i:(?:/|-)(?:Fo|Fd|Fe|Fi|Fa|Fp|FR|OUT:|PDB:|IMPLIB:|LIBPATH:|external:I|I))"?'
    $exactPathPattern = '(?:' +
        $leftBoundary + $pathPattern + '|' +
        $leftBoundary + $attachedPathSwitch + $pathPattern + ')' +
        $rightBoundary
    return [regex]::IsMatch(
        $CommandLine,
        $exactPathPattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Split-RSArtifactCommandLineTokens {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$CommandLine
    )

    $tokens = [System.Collections.Generic.List[string]]::new()
    $builder = [System.Text.StringBuilder]::new()
    $insideQuotes = $false
    foreach ($character in $CommandLine.ToCharArray()) {
        if ($character -eq '"') {
            $insideQuotes = -not $insideQuotes
            continue
        }

        if ([char]::IsWhiteSpace($character) -and -not $insideQuotes) {
            if ($builder.Length -gt 0) {
                $tokens.Add($builder.ToString())
                [void]$builder.Clear()
            }
            continue
        }

        [void]$builder.Append($character)
    }

    if ($builder.Length -gt 0) {
        $tokens.Add($builder.ToString())
    }

    return @($tokens)
}

function Get-RSArtifactCommandLineOptionValue {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $foundValue = $null
    for ($index = 0; $index -lt $Arguments.Count; $index++) {
        $argument = ([string]$Arguments[$index]).Trim()
        if ([string]::Equals($argument, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            if ($index + 1 -lt $Arguments.Count) {
                $candidate = ([string]$Arguments[$index + 1]).Trim()
                if (-not $candidate.StartsWith('-', [System.StringComparison]::Ordinal)) {
                    $foundValue = $candidate
                }
            }
            continue
        }

        $attachedPrefix = "$Name="
        if ($argument.StartsWith($attachedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $foundValue = $argument.Substring($attachedPrefix.Length).Trim()
        }
    }

    return $foundValue
}

function Test-RSVcpkgCommandLineMatchesSharedDependencyScope {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandLine,

        [Parameter(Mandatory = $true)]
        [object]$ScopeInfo
    )

    $arguments = @(Split-RSArtifactCommandLineTokens -CommandLine $CommandLine)
    $triplet = Get-RSArtifactCommandLineOptionValue -Arguments $arguments -Name '--triplet'
    $installRoot = Get-RSArtifactCommandLineOptionValue -Arguments $arguments -Name '--x-install-root'
    if ([string]::IsNullOrWhiteSpace($triplet) -or
        [string]::IsNullOrWhiteSpace($installRoot) -or
        $triplet -notmatch '^(?i)(?<platform>x64|arm64)-[A-Za-z0-9][A-Za-z0-9._-]*$' -or
        -not [System.IO.Path]::IsPathFullyQualified($installRoot) -or
        -not [string]::Equals(
            $Matches.platform,
            $ScopeInfo.Platform,
            [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    try {
        $normalizedInstallRoot = [System.IO.Path]::GetFullPath($installRoot).TrimEnd('\')
        $platformSlug = $ScopeInfo.Platform.ToLowerInvariant()
        foreach ($relativePath in @(
                ('.build\vcpkg_installed\{0}' -f $platformSlug),
                ('.build\vcpkg_install_staging\{0}' -f $triplet))) {
            $expectedInstallRoot = [System.IO.Path]::GetFullPath(
                (Join-Path $ScopeInfo.RepoRoot $relativePath)).TrimEnd('\')
            if ([string]::Equals(
                    $normalizedInstallRoot,
                    $expectedInstallRoot,
                    [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
    }
    catch [System.ArgumentException] {
        return $false
    }
    catch [System.NotSupportedException] {
        return $false
    }

    return $false
}

function Test-RSArtifactCommandLineMatchesSharedDependencyPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandLine,

        [Parameter(Mandatory = $true)]
        [object]$ScopeInfo
    )

    $repoPattern = [regex]::Escape($ScopeInfo.RepoRoot.TrimEnd('\')).Replace('\\', '[\\/]')
    $leftBoundary = '(?<![A-Za-z0-9_:\.\\/-])'
    $attachedPathSwitch = '(?i:(?:/|-)(?:Fo|Fd|Fe|Fi|Fa|Fp|FR|OUT:|PDB:|IMPLIB:|LIBPATH:|external:I|I))"?'
    $sharedRootPattern = '[\\/]\.build[\\/]vcpkg_(?:installed|install_staging)[\\/]'
    $tripletPattern = '(?<platform>x64|arm64)-[A-Za-z0-9][A-Za-z0-9._-]*'
    $rightBoundary = '(?=$|[\s"'';,\)\\/])'
    $pathPattern = $repoPattern + $sharedRootPattern + $tripletPattern
    $evidencePattern = '(?:' +
        $leftBoundary + $pathPattern + '|' +
        $leftBoundary + $attachedPathSwitch + $pathPattern + ')' +
        $rightBoundary

    foreach ($match in [regex]::Matches(
            $CommandLine,
            $evidencePattern,
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
        if ([string]::Equals(
                $match.Groups['platform'].Value,
                $ScopeInfo.Platform,
                [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Get-RSMSBuildPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $propertyPrefixes = @('/p:', '-p:', '--property:', '/property:')
    $foundValue = $null
    foreach ($rawToken in $Arguments) {
        $token = ([string]$rawToken).Trim().Trim('"')
        $propertyList = $null
        foreach ($prefix in $propertyPrefixes) {
            if ($token.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $propertyList = $token.Substring($prefix.Length).Trim('"')
                break
            }
        }

        if ($null -eq $propertyList) {
            continue
        }

        foreach ($assignment in @($propertyList -split '[;,]')) {
            $equalsIndex = $assignment.IndexOf('=')
            if ($equalsIndex -le 0) {
                continue
            }

            $assignmentName = $assignment.Substring(0, $equalsIndex).Trim()
            $assignmentValue = $assignment.Substring($equalsIndex + 1).Trim().Trim('"', "'")
            if ([string]::Equals($assignmentName, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                # MSBuild applies the final occurrence when a property is supplied
                # more than once, so coordination resolves the same effective value.
                $foundValue = $assignmentValue
            }
        }
    }

    return $foundValue
}

function Test-RSMSBuildPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandLine,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $actualValue = Get-RSMSBuildPropertyValue `
        -Arguments @(Split-RSArtifactCommandLineTokens -CommandLine $CommandLine) `
        -Name $Name
    return [string]::Equals($actualValue, $Value, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-RSArtifactCommandLineMatchesScope {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$CommandLine,

        [Parameter(Mandatory = $true)]
        [object]$ScopeInfo
    )

    if ([string]::IsNullOrWhiteSpace($CommandLine) -or
        -not (Test-RSArtifactCommandLineContainsPath -CommandLine $CommandLine -Path $ScopeInfo.RepoRoot)) {
        return $false
    }

    if (-not $ScopeInfo.IsProfileScoped) {
        return $true
    }

    foreach ($artifactRoot in $ScopeInfo.ArtifactRoots) {
        if (Test-RSArtifactCommandLineContainsPath -CommandLine $CommandLine -Path $artifactRoot) {
            return $true
        }
    }

    if ($ScopeInfo.IsSharedDependencyScope) {
        if (Test-RSArtifactCommandLineMatchesSharedDependencyPath `
                -CommandLine $CommandLine `
                -ScopeInfo $ScopeInfo) {
            return $true
        }

        if (Test-RSVcpkgCommandLineMatchesSharedDependencyScope `
                -CommandLine $CommandLine `
                -ScopeInfo $ScopeInfo) {
            return $true
        }

        return Test-RSMSBuildPropertyValue `
            -CommandLine $CommandLine `
            -Name 'Platform' `
            -Value $ScopeInfo.Platform
    }

    return (Test-RSMSBuildPropertyValue `
            -CommandLine $CommandLine `
            -Name 'Configuration' `
            -Value $ScopeInfo.Configuration) -and
        (Test-RSMSBuildPropertyValue `
            -CommandLine $CommandLine `
            -Name 'Platform' `
            -Value $ScopeInfo.Platform)
}

function Get-RSResidualArtifactToolProcesses {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $scopeInfo = Get-RSArtifactOperationScopeInfo -RepoRoot $RepoRoot -Scope $Scope
    $processes = @()
    try {
        $processes = @(Get-CimInstance Win32_Process `
                -Filter "Name='MSBuild.exe' OR Name='cl.exe' OR Name='link.exe' OR Name='vcpkg.exe'" `
                -ErrorAction Stop)
    }
    catch {
        throw "Unable to verify that no residual RedSalamander build tools are active: $($_.Exception.Message)"
    }

    return @($processes | Where-Object {
            $commandLine = [string]$_.CommandLine
            if ([string]::IsNullOrWhiteSpace($commandLine)) {
                return $false
            }

            return Test-RSArtifactCommandLineMatchesScope `
                -CommandLine $commandLine `
                -ScopeInfo $scopeInfo
        })
}

function Assert-RSNoResidualArtifactToolProcesses {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [hashtable]$Scope = @{}
    )

    $scopeInfo = Get-RSArtifactOperationScopeInfo -RepoRoot $RepoRoot -Scope $Scope
    $processes = @(Get-RSResidualArtifactToolProcesses -RepoRoot $RepoRoot -Scope $Scope)
    if ($processes.Count -eq 0) {
        return
    }

    $diagnostics = @($processes | ForEach-Object {
            "  PID=$($_.ProcessId); ParentPID=$($_.ParentProcessId); Name='$($_.Name)'; CommandLine='$($_.CommandLine)'"
        })
    throw ("Cannot start while a build tool is still touching artifact scope '$($scopeInfo.Key)' in this repository. " +
           "Wait for it to exit; do not start another operation against the same outputs or intermediates.`n" +
           ($diagnostics -join "`n"))
}

Export-ModuleMember -Function @(
    'Resolve-RSGuardedRepositoryPath',
    'Publish-RSGuardedRepositoryFile',
    'Get-RSArtifactOperationLockPath',
    'Get-RSMSBuildPropertyValue',
    'Get-RSArtifactOperationOwnerMetadataPath',
    'Read-RSArtifactOperationOwnerMetadata',
    'Remove-RSArtifactOperationOwnerMetadata',
    'Enter-RSArtifactOperationLock',
    'New-RSArtifactOperationChildDelegation',
    'Complete-RSArtifactOperationChildDelegation',
    'Exit-RSArtifactOperationLock',
    'Get-RSArtifactContaminationMarkerPath',
    'Set-RSArtifactOperationContaminated',
    'Read-RSArtifactOperationContamination',
    'Test-RSArtifactOperationRepairAllowed',
    'Test-RSArtifactOperationContaminated',
    'Clear-RSArtifactOperationContaminated',
    'Assert-RSNoResidualArtifactToolProcesses'
)
