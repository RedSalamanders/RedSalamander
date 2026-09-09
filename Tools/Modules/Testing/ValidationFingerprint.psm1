Set-StrictMode -Version Latest

$script:CanonicalizationId = 'rs-canonical-json-v1'
$script:MaximumRecordBytes = 16MB
$script:MaximumDepth = 64
$script:MaximumNodes = 100000
$script:MaximumStringCharacters = 1MB
$script:MaximumRelativePathCharacters = 1024
$script:Utf8NoBomStrict = [System.Text.UTF8Encoding]::new($false, $true)
$script:ReasonCodes = @(
    'NO_PRIOR_EVIDENCE',
    'PREVIOUS_NOT_GREEN',
    'CORRUPT_EVIDENCE',
    'SCHEMA_MISMATCH',
    'SOURCE_INPUT_CHANGED',
    'BUILD_ARTIFACT_CHANGED',
    'LOGICAL_INPUT_CHANGED',
    'SHARED_CONTRACT_CHANGED',
    'RUNNER_CHANGED',
    'ENVIRONMENT_CHANGED',
    'ARGUMENTS_CHANGED',
    'COVERAGE_CHANGED',
    'OUTCOME_POLICY_CHANGED',
    'CAPABILITY_CHANGED',
    'DYNAMIC_SKIP_NOT_REUSABLE',
    'ARTIFACT_EPOCH_CHANGED',
    'ARTIFACT_MUTATOR_REQUIRES_REATTESTATION',
    'ORIGIN_NOT_PROMOTED',
    'POST_SNAPSHOT_MISSING',
    'PORTABLE_ATTESTATION_REJECTED',
    'NON_CACHEABLE',
    'SOURCE_CHANGED_DURING_RUN',
    'FORCED',
    'REUSED_EXACT',
    'REUSED_INDEPENDENT_LOGICAL_ARTIFACT',
    'MONOLITHIC_LOGICAL_CANDIDATE_NOT_FULL'
)

function Assert-RSStartrailPowerShellHost {
    [CmdletBinding()]
    param()

    if ($PSVersionTable.PSVersion.Major -lt 7 -or
        ($PSVersionTable.PSVersion.Major -eq 7 -and $PSVersionTable.PSVersion.Minor -lt 4)) {
        throw 'Operation Startrail evidence requires PowerShell 7.4 or newer.'
    }
}

function ConvertTo-RSCanonicalJsonStringLiteral {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    $normalized = $Value.Normalize([Text.NormalizationForm]::FormC)
    if ($normalized.Length -gt $script:MaximumStringCharacters) {
        throw "Canonical JSON string exceeds $($script:MaximumStringCharacters) characters."
    }

    $builder = [Text.StringBuilder]::new($normalized.Length + 2)
    [void]$builder.Append('"')
    for ($index = 0; $index -lt $normalized.Length; $index++) {
        $character = $normalized[$index]
        $code = [int]$character
        if ($code -eq 0x22) { [void]$builder.Append('\"'); continue }
        if ($code -eq 0x5c) { [void]$builder.Append('\\'); continue }
        if ($code -eq 0x08) { [void]$builder.Append('\b'); continue }
        if ($code -eq 0x0c) { [void]$builder.Append('\f'); continue }
        if ($code -eq 0x0a) { [void]$builder.Append('\n'); continue }
        if ($code -eq 0x0d) { [void]$builder.Append('\r'); continue }
        if ($code -eq 0x09) { [void]$builder.Append('\t'); continue }

        if ($code -lt 0x20) {
            [void]$builder.Append(('\u{0}' -f $code.ToString('x4', [Globalization.CultureInfo]::InvariantCulture)))
            continue
        }
        if ([char]::IsHighSurrogate($character)) {
            if ($index + 1 -ge $normalized.Length -or -not [char]::IsLowSurrogate($normalized[$index + 1])) {
                throw 'Canonical JSON rejects an unpaired high surrogate.'
            }
            [void]$builder.Append($character)
            $index++
            [void]$builder.Append($normalized[$index])
            continue
        }
        if ([char]::IsLowSurrogate($character)) {
            throw 'Canonical JSON rejects an unpaired low surrogate.'
        }
        [void]$builder.Append($character)
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function ConvertTo-RSCanonicalJsonValue {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][int]$Depth,
        [Parameter(Mandatory = $true)][ref]$NodeCount
    )

    if ($Depth -gt $script:MaximumDepth) {
        throw "Canonical JSON exceeds depth $($script:MaximumDepth)."
    }
    $NodeCount.Value++
    if ($NodeCount.Value -gt $script:MaximumNodes) {
        throw "Canonical JSON exceeds $($script:MaximumNodes) nodes."
    }

    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [string]) { return ConvertTo-RSCanonicalJsonStringLiteral -Value $Value }

    $integralTypes = @(
        [byte], [sbyte], [int16], [uint16], [int32], [uint32], [int64], [uint64])
    if ($Value.GetType() -in $integralTypes) {
        if ($Value -is [uint64] -and $Value -gt [int64]::MaxValue) {
            throw 'Canonical JSON integers are limited to signed 64-bit values.'
        }
        return ([int64]$Value).ToString([Globalization.CultureInfo]::InvariantCulture)
    }

    if ($Value -is [Collections.IDictionary]) {
        $properties = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
        foreach ($key in $Value.Keys) {
            if ($key -isnot [string]) { throw 'Canonical JSON object keys must be strings.' }
            $normalizedName = $key.Normalize([Text.NormalizationForm]::FormC)
            if ($properties.ContainsKey($normalizedName)) {
                throw "Canonical JSON object contains duplicate normalized key '$normalizedName'."
            }
            $properties.Add($normalizedName, $Value[$key])
        }
        $orderedNames = [string[]]@($properties.Keys)
        [Array]::Sort($orderedNames, [StringComparer]::Ordinal)
        $members = foreach ($propertyName in $orderedNames) {
            $nameJson = ConvertTo-RSCanonicalJsonStringLiteral -Value $propertyName
            $valueJson = ConvertTo-RSCanonicalJsonValue -Value $properties[$propertyName] -Depth ($Depth + 1) -NodeCount $NodeCount
            "$nameJson`:$valueJson"
        }
        return '{' + ($members -join ',') + '}'
    }

    if ($Value.GetType() -eq [Management.Automation.PSCustomObject]) {
        $dictionary = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
        foreach ($property in $Value.PSObject.Properties) {
            if ($property.MemberType -notin @('NoteProperty', 'Property', 'AliasProperty')) { continue }
            if ($dictionary.ContainsKey($property.Name)) {
                throw "Canonical JSON object contains duplicate key '$($property.Name)'."
            }
            $dictionary.Add($property.Name, $property.Value)
        }
        return ConvertTo-RSCanonicalJsonValue -Value $dictionary -Depth $Depth -NodeCount $NodeCount
    }

    if ($Value -is [Collections.IEnumerable]) {
        $items = foreach ($item in $Value) {
            ConvertTo-RSCanonicalJsonValue -Value $item -Depth ($Depth + 1) -NodeCount $NodeCount
        }
        return '[' + (@($items) -join ',') + ']'
    }

    throw "Canonical JSON does not support values of type '$($Value.GetType().FullName)'."
}

function ConvertTo-RSCanonicalJson {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowNull()][object]$InputObject)

    Assert-RSStartrailPowerShellHost
    $nodeCount = 0
    return ConvertTo-RSCanonicalJsonValue -Value $InputObject -Depth 1 -NodeCount ([ref]$nodeCount)
}

function ConvertTo-RSCanonicalJsonBytes {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowNull()][object]$InputObject)

    $json = ConvertTo-RSCanonicalJson -InputObject $InputObject
    return $script:Utf8NoBomStrict.GetBytes($json)
}

function Get-RSCanonicalJsonDigest {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowNull()][object]$InputObject)

    $bytes = ConvertTo-RSCanonicalJsonBytes -InputObject $InputObject
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        return ([Convert]::ToHexString($sha256.ComputeHash($bytes))).ToLowerInvariant()
    } finally {
        $sha256.Dispose()
    }
}

function Test-RSValidationRelativePath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Path)

    if ($Path.Length -lt 1 -or $Path.Length -gt $script:MaximumRelativePathCharacters) { return $false }
    if ($Path.Contains('\') -or $Path.Contains(':') -or $Path.StartsWith('/') -or $Path.Contains('//')) { return $false }
    foreach ($segment in $Path.Split('/')) {
        if ($segment.Length -eq 0 -or $segment -in @('.', '..')) { return $false }
    }
    return -not [regex]::IsMatch($Path, '[\u0000-\u001f]')
}

function Get-RSValidationReasonCodes {
    [CmdletBinding()]
    param()

    return @($script:ReasonCodes)
}

function Assert-RSValidationIdentifier {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Id)

    if ($Id.Length -gt 128 -or $Id -cnotmatch '^[a-z0-9](?:[a-z0-9._-]*[a-z0-9])?$') {
        throw "Invalid validation identifier '$Id'."
    }
}

function Assert-RSValidationPlanSemantics {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Plan)

    $ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in @($Plan.entries)) {
        Assert-RSValidationIdentifier -Id ([string]$entry.id)
        if (-not $ids.Add([string]$entry.id)) {
            throw "Duplicate validation entry id '$($entry.id)'."
        }
    }
}

function Assert-RSValidationPathSet {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string[]]$Path)

    $aliases = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $Path) {
        if (-not (Test-RSValidationRelativePath -Path $item)) {
            throw "Invalid validation relative path '$item'."
        }
        $normalized = $item.Normalize([Text.NormalizationForm]::FormC)
        if (-not $aliases.Add($normalized)) {
            throw "Validation path set contains a Windows alias for '$item'."
        }
    }
}

function Assert-RSValidationProvenance {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object[]]$Record)

    $origins = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    foreach ($item in $Record) {
        $runId = [string]$item.run_id
        if ($origins.ContainsKey($runId)) { throw "Duplicate provenance run '$runId'." }
        $origin = if ($null -eq $item.origin_run_id) { '' } else { [string]$item.origin_run_id }
        $origins.Add($runId, $origin)
    }
    foreach ($start in $origins.Keys) {
        $visited = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $current = $start
        while ($current.Length -gt 0) {
            if (-not $visited.Add($current)) { throw "Validation provenance cycle includes '$current'." }
            if (-not $origins.ContainsKey($current)) { throw "Validation provenance has dangling origin '$current'." }
            $current = $origins[$current]
        }
    }
}

function Assert-RSValidationRecordBounds {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $item = Get-Item -LiteralPath $LiteralPath -ErrorAction Stop
    if ($item.Length -gt $script:MaximumRecordBytes) {
        throw "Validation record exceeds $($script:MaximumRecordBytes) bytes: $LiteralPath"
    }
}

function Assert-RSValidationRecordPath {
    param(
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][string]$AllowedRoot
    )

    $root = [IO.Path]::GetFullPath($AllowedRoot).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $path = [IO.Path]::GetFullPath($LiteralPath)
    $prefix = $root + [IO.Path]::DirectorySeparatorChar
    if (-not $path.Equals($root, [StringComparison]::OrdinalIgnoreCase) -and
        -not $path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Validation record path must remain beneath '$root': $path"
    }
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        throw "Validation record root does not exist: $root"
    }

    $relativeParent = [IO.Path]::GetRelativePath($root, (Split-Path -Parent $path))
    $current = Get-Item -LiteralPath $root -Force
    if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Validation record root must not be a reparse point: $root"
    }
    if ($relativeParent -ne '.') {
        foreach ($segment in $relativeParent.Split([IO.Path]::DirectorySeparatorChar)) {
            $current = Get-Item -LiteralPath (Join-Path $current.FullName $segment) -Force
            if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Validation record path contains a reparse point: $($current.FullName)"
            }
        }
    }
    return $path
}

function Invoke-RSValidationGitBinary {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string[]]$Argument
    )

    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = 'git.exe'
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.CreateNoWindow = $true
    foreach ($value in @('-C', $RepoRoot) + $Argument) { [void]$start.ArgumentList.Add($value) }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    $memory = [IO.MemoryStream]::new()
    try {
        if (-not $process.Start()) { throw 'Unable to start Git for workspace snapshot.' }
        # Drain both redirected streams concurrently. Git can emit enough line-ending
        # diagnostics on a dirty workspace to fill stderr while stdout is still open.
        $copyTask = $process.StandardOutput.BaseStream.CopyToAsync($memory)
        $errorTask = $process.StandardError.ReadToEndAsync()
        $null = $copyTask.GetAwaiter().GetResult()
        $errorText = $errorTask.GetAwaiter().GetResult()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "Git workspace query failed with exit $($process.ExitCode): $errorText" }
        return ,$memory.ToArray()
    } finally {
        $memory.Dispose()
        $process.Dispose()
    }
}

function Get-RSValidationBytesSha256 {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)

    return ([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($Bytes))).ToLowerInvariant()
}

function ConvertFrom-RSValidationGitNullList {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)

    if ($Bytes.Count -eq 0) { return @() }
    return @($script:Utf8NoBomStrict.GetString($Bytes).Split([char]0, [StringSplitOptions]::RemoveEmptyEntries))
}

function Get-RSValidationGitObjectMap {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [AllowEmptyString()][string]$Tree = ''
    )

    $arguments = if ([string]::IsNullOrEmpty($Tree)) {
        @('ls-files', '--stage', '-z', '--', '.')
    } else {
        @('ls-tree', '-r', '-z', $Tree, '--', '.')
    }
    $records = ConvertFrom-RSValidationGitNullList -Bytes (
        Invoke-RSValidationGitBinary -RepoRoot $RepoRoot -Argument $arguments)
    $map = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    foreach ($record in $records) {
        $match = if ([string]::IsNullOrEmpty($Tree)) {
            [regex]::Match($record, '^(?<mode>[0-7]{6}) (?<oid>[0-9a-fA-F]{40,64}) (?<stage>[0-3])\t(?<path>[\s\S]+)$')
        } else {
            [regex]::Match($record, '^(?<mode>[0-7]{6}) (?<type>blob|commit) (?<oid>[0-9a-fA-F]{40,64})\t(?<path>[\s\S]+)$')
        }
        if (-not $match.Success) { throw "Unable to parse Git object record for workspace snapshot: $record" }
        if ([string]::IsNullOrEmpty($Tree) -and $match.Groups['stage'].Value -ne '0') {
            throw "Workspace snapshot refuses an unresolved Git index stage for '$($match.Groups['path'].Value)'."
        }
        $path = $match.Groups['path'].Value.Replace('\', '/')
        if (-not (Test-RSValidationWorkspacePathIncluded -Path $path)) { continue }
        if ($map.ContainsKey($path)) { throw "Duplicate Git object path '$path'." }
        $map.Add($path, [pscustomobject]@{
                Mode = $match.Groups['mode'].Value
                Oid = $match.Groups['oid'].Value.ToLowerInvariant()
                Type = if ([string]::IsNullOrEmpty($Tree)) { 'blob' } else { $match.Groups['type'].Value }
            })
    }
    return ,$map
}

function Test-RSValidationWorkspacePathIncluded {
    param([Parameter(Mandatory = $true)][string]$Path)

    $normalized = $Path.Replace('\', '/')
    if (-not (Test-RSValidationRelativePath -Path $normalized)) { throw "Git returned unsafe workspace path '$Path'." }
    foreach ($prefix in @('.git/', '.build/', 'Specs/TestRuns/')) {
        if ($normalized.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return $false }
    }
    if ($normalized.EndsWith('.log', [StringComparison]::OrdinalIgnoreCase)) { return $false }
    return $true
}

function Get-RSValidationGitNameStatus {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string[]]$Argument
    )

    $tokens = @(ConvertFrom-RSValidationGitNullList -Bytes (
            Invoke-RSValidationGitBinary -RepoRoot $RepoRoot -Argument $Argument))
    $records = @()
    $index = 0
    while ($index -lt $tokens.Count) {
        $status = $tokens[$index++]
        $inlinePath = $null
        if ($status.Contains("`t")) {
            $parts = $status.Split("`t", 2)
            $status = $parts[0]
            $inlinePath = $parts[1]
        }
        if ($status -notmatch '^[ACDMRTUXB][0-9]*$') { throw "Unexpected Git name-status token '$status'." }
        $pathCount = if ($status[0] -in @('R', 'C')) { 2 } else { 1 }
        $paths = @()
        if ($null -ne $inlinePath) { $paths += $inlinePath }
        while ($paths.Count -lt $pathCount) {
            if ($index -ge $tokens.Count) { throw "Incomplete Git name-status record '$status'." }
            $paths += $tokens[$index++]
        }
        $records += [pscustomobject]@{ Status = $status; Paths = @($paths | ForEach-Object { $_.Replace('\', '/') }) }
    }
    return $records
}

function Resolve-RSValidationWorkspaceFile {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$RelativePath
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $path = [IO.Path]::GetFullPath((Join-Path $root $RelativePath))
    if (-not $path.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Workspace path escaped the repository: $RelativePath"
    }
    $current = Get-Item -LiteralPath $root -Force -ErrorAction Stop
    if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Workspace root must not be a reparse point: $root"
    }
    foreach ($segment in [IO.Path]::GetRelativePath($root, $path).Split([IO.Path]::DirectorySeparatorChar)) {
        $next = Join-Path $current.FullName $segment
        if (-not (Test-Path -LiteralPath $next)) { return $path }
        $current = Get-Item -LiteralPath $next -Force -ErrorAction Stop
        if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Workspace path contains a reparse point: $($current.FullName)"
        }
    }
    return $path
}

function Get-RSValidationFileIdentity {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $item = Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
    if ($item.PSIsContainer) { throw "Workspace input must be a regular file: $LiteralPath" }
    $stream = [IO.File]::Open($item.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [pscustomobject]@{
            Sha256 = ([Convert]::ToHexString($sha.ComputeHash($stream))).ToLowerInvariant()
            SizeBytes = [int64]$stream.Length
        }
    } finally {
        $sha.Dispose()
        $stream.Dispose()
    }
}

function Get-RSValidationGitBlobSha256 {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][object]$Object,
        [Parameter(Mandatory = $true)][Collections.Generic.Dictionary[string, string]]$Cache
    )

    $key = "$($Object.Type):$($Object.Oid)"
    if ($Cache.ContainsKey($key)) { return $Cache[$key] }
    $digest = if ($Object.Type -eq 'blob') {
        Get-RSValidationBytesSha256 -Bytes (Invoke-RSValidationGitBinary -RepoRoot $RepoRoot -Argument @('cat-file', 'blob', $Object.Oid))
    } else {
        Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ gitlink = $Object.Oid })
    }
    $Cache.Add($key, $digest)
    return $digest
}

function Add-RSValidationGitBlobSha256Batch {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Object,
        [Parameter(Mandatory = $true)][Collections.Generic.Dictionary[string, string]]$Cache
    )

    $blobOids = [Collections.Generic.List[string]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($candidate in $Object) {
        if ($null -eq $candidate) { continue }
        $key = "$($candidate.Type):$($candidate.Oid)"
        if ($Cache.ContainsKey($key) -or -not $seen.Add($key)) { continue }
        if ($candidate.Type -ne 'blob') {
            $Cache.Add($key, (Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ gitlink = $candidate.Oid })))
            continue
        }
        $blobOids.Add([string]$candidate.Oid)
    }
    if ($blobOids.Count -eq 0) { return }

    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = 'git.exe'
    $start.UseShellExecute = $false
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.CreateNoWindow = $true
    foreach ($value in @('-C', $RepoRoot, 'cat-file', '--batch')) { [void]$start.ArgumentList.Add($value) }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    $memory = [IO.MemoryStream]::new()
    try {
        if (-not $process.Start()) { throw 'Unable to start Git blob batch for workspace snapshot.' }
        $copyTask = $process.StandardOutput.BaseStream.CopyToAsync($memory)
        $errorTask = $process.StandardError.ReadToEndAsync()
        foreach ($oid in $blobOids) { $process.StandardInput.WriteLine($oid) }
        $process.StandardInput.Close()
        $null = $copyTask.GetAwaiter().GetResult()
        $errorText = $errorTask.GetAwaiter().GetResult()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) {
            throw "Git blob batch failed with exit $($process.ExitCode): $errorText"
        }

        $bytes = $memory.ToArray()
        $offset = 0
        foreach ($requestedOid in $blobOids) {
            $lineEnd = [Array]::IndexOf($bytes, [byte]10, $offset)
            if ($lineEnd -lt 0) { throw "Git blob batch omitted the header for '$requestedOid'." }
            $header = $script:Utf8NoBomStrict.GetString($bytes, $offset, $lineEnd - $offset).TrimEnd("`r")
            $match = [regex]::Match($header, '^(?<oid>[0-9a-fA-F]{40,64}) blob (?<size>[0-9]+)$')
            if (-not $match.Success -or $match.Groups['oid'].Value -ne $requestedOid) {
                throw "Unexpected Git blob batch header '$header' for '$requestedOid'."
            }
            $size = [int64]::Parse($match.Groups['size'].Value, [Globalization.CultureInfo]::InvariantCulture)
            if ($size -gt [int]::MaxValue) { throw "Git blob '$requestedOid' is too large to fingerprint safely." }
            $contentOffset = $lineEnd + 1
            $nextOffset = $contentOffset + [int]$size
            if ($nextOffset -ge $bytes.Length -or $bytes[$nextOffset] -ne [byte]10) {
                throw "Git blob batch returned truncated content for '$requestedOid'."
            }
            $content = [byte[]]::new([int]$size)
            if ($size -gt 0) { [Array]::Copy($bytes, $contentOffset, $content, 0, [int]$size) }
            $Cache.Add("blob:$requestedOid", (Get-RSValidationBytesSha256 -Bytes $content))
            $offset = $nextOffset + 1
        }
        if ($offset -ne $bytes.Length) { throw 'Git blob batch returned unexpected trailing data.' }
    } finally {
        $memory.Dispose()
        $process.Dispose()
    }
}

function Resolve-RSValidationImpactBase {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [AllowEmptyString()][string]$ImpactBase
    )

    $head = $script:Utf8NoBomStrict.GetString((Invoke-RSValidationGitBinary -RepoRoot $RepoRoot -Argument @(
                'rev-parse', '--verify', '--end-of-options', 'HEAD^{commit}'))).Trim()
    if ($head -notmatch '^[0-9a-fA-F]{40,64}$') { throw 'Unable to resolve Git HEAD for workspace snapshot.' }
    if ([string]::IsNullOrWhiteSpace($ImpactBase)) {
        return [pscustomobject]@{ Head = $head.ToLowerInvariant(); Base = $head.ToLowerInvariant() }
    }
    $candidate = $script:Utf8NoBomStrict.GetString((Invoke-RSValidationGitBinary -RepoRoot $RepoRoot -Argument @(
                'rev-parse', '--verify', '--end-of-options', "$ImpactBase^{commit}"))).Trim()
    if ($candidate -notmatch '^[0-9a-fA-F]{40,64}$') { throw "Unable to resolve local impact base '$ImpactBase'." }
    $mergeBase = $script:Utf8NoBomStrict.GetString((Invoke-RSValidationGitBinary -RepoRoot $RepoRoot -Argument @(
                'merge-base', '--', $head, $candidate))).Trim()
    if ($mergeBase -notmatch '^[0-9a-fA-F]{40,64}$') { throw "Unable to resolve merge base for '$ImpactBase'." }
    return [pscustomobject]@{ Head = $head.ToLowerInvariant(); Base = $mergeBase.ToLowerInvariant() }
}

function Get-RSWorkspaceSnapshotOnce {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [AllowEmptyString()][string]$ImpactBase = ''
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $gitDiscoveryTimer = [Diagnostics.Stopwatch]::StartNew()
    $commits = Resolve-RSValidationImpactBase -RepoRoot $root -ImpactBase $ImpactBase
    $baseMap = Get-RSValidationGitObjectMap -RepoRoot $root -Tree $commits.Base
    $indexMap = Get-RSValidationGitObjectMap -RepoRoot $root
    $changedPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $renameByNewPath = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $cachedDiffs = @(Get-RSValidationGitNameStatus -RepoRoot $root -Argument @(
            'diff', '--cached', '--name-status', '-z', '-M', $commits.Base, '--', '.'))
    $worktreeDiffs = @(Get-RSValidationGitNameStatus -RepoRoot $root -Argument @(
            'diff', '--name-status', '-z', '-M', '--', '.'))
    foreach ($record in @($cachedDiffs) + @($worktreeDiffs)) {
        if ($record.Paths.Count -eq 2) {
            $oldPath = $record.Paths[0]
            $newPath = $record.Paths[1]
            $oldIncluded = Test-RSValidationWorkspacePathIncluded $oldPath
            $newIncluded = Test-RSValidationWorkspacePathIncluded $newPath
            if ($oldIncluded -and $newIncluded) {
                $renameByNewPath[$newPath] = $oldPath
                [void]$changedPaths.Add($newPath)
            } else {
                # Inclusion is a property of each side, not of the rename pair.
                # Crossing the evidence boundary must retain the included deletion
                # or addition even though the excluded side is intentionally absent.
                if ($oldIncluded) { [void]$changedPaths.Add($oldPath) }
                if ($newIncluded) { [void]$changedPaths.Add($newPath) }
            }
        } else {
            $path = $record.Paths[0]
            if (Test-RSValidationWorkspacePathIncluded $path) { [void]$changedPaths.Add($path) }
        }
    }
    $untracked = @(ConvertFrom-RSValidationGitNullList -Bytes (Invoke-RSValidationGitBinary -RepoRoot $root -Argument @(
                'ls-files', '--others', '--exclude-standard', '-z', '--', '.')))
    foreach ($path in $untracked) {
        $normalized = $path.Replace('\', '/')
        if (Test-RSValidationWorkspacePathIncluded $normalized) { [void]$changedPaths.Add($normalized) }
    }
    $paths = [string[]]@($changedPaths)
    [Array]::Sort($paths, [StringComparer]::Ordinal)
    if ($paths.Count -gt 0) { Assert-RSValidationPathSet -Path $paths }
    $gitDiscoveryTimer.Stop()
    $contentHashingTimer = [Diagnostics.Stopwatch]::StartNew()
    $blobCache = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $objectsToHash = foreach ($path in $paths) {
        $oldPath = if ($renameByNewPath.ContainsKey($path)) { $renameByNewPath[$path] } else { $null }
        $basePath = if ($null -eq $oldPath) { $path } else { $oldPath }
        if ($baseMap.ContainsKey($basePath)) { $baseMap[$basePath] }
        if ($indexMap.ContainsKey($path)) { $indexMap[$path] }
    }
    Add-RSValidationGitBlobSha256Batch -RepoRoot $root -Object @($objectsToHash) -Cache $blobCache
    $files = foreach ($path in $paths) {
        $oldPath = if ($renameByNewPath.ContainsKey($path)) { $renameByNewPath[$path] } else { $null }
        $basePath = if ($null -eq $oldPath) { $path } else { $oldPath }
        $baseObject = if ($baseMap.ContainsKey($basePath)) { $baseMap[$basePath] } else { $null }
        $indexObject = if ($indexMap.ContainsKey($path)) { $indexMap[$path] } else { $null }
        $fullPath = Resolve-RSValidationWorkspaceFile -RepoRoot $root -RelativePath $path
        $worktreeExists = Test-Path -LiteralPath $fullPath -PathType Leaf
        $worktree = if ($worktreeExists) { Get-RSValidationFileIdentity -LiteralPath $fullPath } else { $null }
        $status = if ($null -ne $oldPath) {
            'renamed'
        } elseif ($null -eq $baseObject -and $null -eq $indexObject -and $worktreeExists) {
            'untracked'
        } elseif ($null -eq $baseObject) {
            'added'
        } elseif (-not $worktreeExists) {
            'deleted'
        } elseif ($null -ne $indexObject -and $baseObject.Mode -ne $indexObject.Mode) {
            'type-changed'
        } else {
            'modified'
        }
        [pscustomobject][ordered]@{
            path = $path
            status = $status
            base_sha256 = if ($null -eq $baseObject) { $null } else { Get-RSValidationGitBlobSha256 -RepoRoot $root -Object $baseObject -Cache $blobCache }
            base_mode = if ($null -eq $baseObject) { $null } else { $baseObject.Mode }
            index_sha256 = if ($null -eq $indexObject) { $null } else { Get-RSValidationGitBlobSha256 -RepoRoot $root -Object $indexObject -Cache $blobCache }
            index_mode = if ($null -eq $indexObject) { $null } else { $indexObject.Mode }
            worktree_sha256 = if ($null -eq $worktree) { $null } else { $worktree.Sha256 }
            worktree_size_bytes = if ($null -eq $worktree) { $null } else { $worktree.SizeBytes }
            old_path = $oldPath
        }
    }
    $contentHashingTimer.Stop()
    $canonicalizationTimer = [Diagnostics.Stopwatch]::StartNew()
    $projection = [ordered]@{
        schema = 'red-salamander.workspace-snapshot.v1'
        canonicalization = 'rs-canonical-json-v1'
        repo_identity = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ physical_root = $root.ToLowerInvariant() })
        git_head = $commits.Head
        base = $commits.Base
        files = @($files)
    }
    $result = [pscustomobject][ordered]@{
        schema = $projection.schema
        canonicalization = $projection.canonicalization
        snapshot_id = Get-RSCanonicalJsonDigest -InputObject $projection
        repo_identity = $projection.repo_identity
        git_head = $projection.git_head
        base = $projection.base
        files = $projection.files
        acquisition_before = '0' * 64
        acquisition_after = '0' * 64
        acquisition_attempts = [int64]1
        git_discovery_ms = [int64][Math]::Max(0, [Math]::Round($gitDiscoveryTimer.Elapsed.TotalMilliseconds))
        content_hashing_ms = [int64][Math]::Max(0, [Math]::Round($contentHashingTimer.Elapsed.TotalMilliseconds))
        canonicalization_ms = [int64]0
        computation_ms = [int64]0
        created_utc = [datetime]::UtcNow.ToString('o')
    }
    $canonicalizationTimer.Stop()
    $result.canonicalization_ms = [int64][Math]::Max(0, [Math]::Round($canonicalizationTimer.Elapsed.TotalMilliseconds))
    return $result
}

function Get-RSWorkspaceSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [AllowEmptyString()][string]$ImpactBase = '',
        [ValidateRange(1, 10)][int]$MaximumAttempts = 3
    )

    Assert-RSStartrailPowerShellHost
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $before = Get-RSWorkspaceSnapshotOnce -RepoRoot $RepoRoot -ImpactBase $ImpactBase
    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        $after = Get-RSWorkspaceSnapshotOnce -RepoRoot $RepoRoot -ImpactBase $ImpactBase
        if ($before.snapshot_id -eq $after.snapshot_id) {
            $timer.Stop()
            $after.acquisition_before = $before.snapshot_id
            $after.acquisition_after = $after.snapshot_id
            $after.acquisition_attempts = [int64]$attempt
            $after.computation_ms = [int64][Math]::Max(0, [Math]::Round($timer.Elapsed.TotalMilliseconds))
            return $after
        }
        $before = $after
    }
    $timer.Stop()
    throw "Workspace changed during $MaximumAttempts consecutive snapshot acquisitions."
}

function Get-RSValidationBuildReceiptBindingDigest {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$BuildReceipt)

    $projection = [ordered]@{
        receipt_id = $BuildReceipt.receipt_id
        source_snapshot_id = $BuildReceipt.source_snapshot_id
        platform = $BuildReceipt.platform
        configuration = $BuildReceipt.configuration
        target = $BuildReceipt.target
        tests_enabled = [bool]$BuildReceipt.tests_enabled
        toolchain_identity_digest = $BuildReceipt.toolchain_identity_digest
        vcpkg_identity_digest = $BuildReceipt.vcpkg_identity_digest
        ghostty_runtime_identity = $BuildReceipt.ghostty_runtime_identity
        project_graph_digest = $BuildReceipt.project_graph_digest
        runtime_dependency_digest = $BuildReceipt.runtime_dependency_digest
        outputs = @($BuildReceipt.outputs | Sort-Object relative_path | ForEach-Object {
                [ordered]@{
                    id = $_.id; role = $_.role; relative_path = $_.relative_path
                    size_bytes = [int64]$_.size_bytes; sha256 = $_.sha256
                    runtime_closure = @($_.runtime_closure)
                }
            })
    }
    return Get-RSCanonicalJsonDigest -InputObject $projection
}

function Get-RSValidationEntryFingerprint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$EntryContract,
        [Parameter(Mandatory = $true)][string]$CoveragePlanDigest,
        [Parameter(Mandatory = $true)][object]$WorkspaceSnapshot,
        [AllowNull()][object]$BuildReceipt = $null,
        [AllowEmptyString()][string]$BuildReceiptBindingDigest = '',
        [hashtable]$InputSetDigests = @{},
        [hashtable]$RunnerDependencyDigests = @{},
        [hashtable]$CapabilityInputs = @{},
        [AllowEmptyString()][string]$QuarantineDigest = '',
        [AllowEmptyString()][string]$ClassificationMode = '',
        [ValidateRange(0, [int64]::MaxValue)][int64]$ArtifactEpoch = 0,
        [string]$CachePolicyVersion = 'cache-policy.v1'
    )

    Assert-RSValidationIdentifier -Id ([string]$EntryContract.id)
    if ($CoveragePlanDigest -notmatch '^[0-9a-f]{64}$' -or $WorkspaceSnapshot.snapshot_id -notmatch '^[0-9a-f]{64}$') {
        throw 'Entry fingerprint requires canonical coverage-plan and workspace digests.'
    }
    $timer = [Diagnostics.Stopwatch]::StartNew()
    if (-not [string]::IsNullOrEmpty($BuildReceiptBindingDigest) -and $BuildReceiptBindingDigest -notmatch '^[0-9a-f]{64}$') {
        throw 'Build receipt binding digest must be lowercase SHA-256.'
    }
    $receiptProjection = if ($null -eq $BuildReceipt) {
        $null
    } else {
        [ordered]@{
            receipt_id = $BuildReceipt.receipt_id
            binding_digest = if ([string]::IsNullOrEmpty($BuildReceiptBindingDigest)) {
                Get-RSValidationBuildReceiptBindingDigest -BuildReceipt $BuildReceipt
            } else {
                $BuildReceiptBindingDigest
            }
        }
    }
    $sharedContract = [ordered]@{}
    foreach ($property in $EntryContract.PSObject.Properties) {
        if ($property.Name -notin @(
                'arguments', 'coverage_contract', 'outcome_contract', 'cache_policy',
                'environment_inputs', 'input_sets')) {
            $sharedContract[$property.Name] = $property.Value
        }
    }
    $componentProjections = [ordered]@{
        build_artifact = $receiptProjection
        logical_input = [ordered]@{ input_sets = @($EntryContract.input_sets); digests = $InputSetDigests }
        shared_contract = $sharedContract
        runner = [ordered]@{ dependency_digests = $RunnerDependencyDigests; cache_policy_version = $CachePolicyVersion }
        environment = [ordered]@{
            classification_mode = $ClassificationMode
            quarantine_digest = if ([string]::IsNullOrEmpty($QuarantineDigest)) { $null } else { $QuarantineDigest }
        }
        arguments = @($EntryContract.arguments)
        coverage = [ordered]@{ plan_digest = $CoveragePlanDigest; contract = $EntryContract.coverage_contract }
        outcome_policy = [ordered]@{ cache_policy = $EntryContract.cache_policy; contract = $EntryContract.outcome_contract }
        capability = [ordered]@{ declarations = @($EntryContract.environment_inputs); inputs = $CapabilityInputs }
        artifact_epoch = [int64]$ArtifactEpoch
        source_input = [string]$WorkspaceSnapshot.snapshot_id
    }
    $componentDigests = [ordered]@{}
    foreach ($componentName in $componentProjections.Keys) {
        $componentDigests[$componentName] = Get-RSCanonicalJsonDigest -InputObject $componentProjections[$componentName]
    }
    $projection = [ordered]@{
        schema = 'red-salamander.entry-fingerprint.v2'
        canonicalization = 'rs-canonical-json-v1'
        component_digest_version = 'validation-entry-components.v1'
        component_digests = $componentDigests
    }
    $fingerprint = Get-RSCanonicalJsonDigest -InputObject $projection
    $timer.Stop()
    return [pscustomobject]@{
        EntryId = $EntryContract.id
        Fingerprint = $fingerprint
        ComponentDigestVersion = $projection.component_digest_version
        ComponentDigests = [pscustomobject]$componentDigests
        ComputationMs = [int64][Math]::Max(0, [Math]::Round($timer.Elapsed.TotalMilliseconds))
    }
}

function Assert-RSValidationJsonTokenBounds {
    param([Parameter(Mandatory = $true)][string]$Json)

    $stringReader = [IO.StringReader]::new($Json)
    $reader = [Newtonsoft.Json.JsonTextReader]::new($stringReader)
    $reader.DateParseHandling = [Newtonsoft.Json.DateParseHandling]::None
    $reader.MaxDepth = $script:MaximumDepth
    $objectKeys = [Collections.Generic.Stack[Collections.Generic.HashSet[string]]]::new()
    $nodeCount = 0
    try {
        while ($reader.Read()) {
            if ($reader.TokenType -eq [Newtonsoft.Json.JsonToken]::StartObject) {
                $nodeCount++
                $objectKeys.Push([Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal))
            } elseif ($reader.TokenType -eq [Newtonsoft.Json.JsonToken]::StartArray) {
                $nodeCount++
            } elseif ($reader.TokenType -eq [Newtonsoft.Json.JsonToken]::EndObject) {
                [void]$objectKeys.Pop()
            } elseif ($reader.TokenType -eq [Newtonsoft.Json.JsonToken]::PropertyName) {
                $rawName = [string]$reader.Value
                if ($rawName.Length -gt $script:MaximumStringCharacters) {
                    throw "Validation JSON property name exceeds $($script:MaximumStringCharacters) characters."
                }
                $name = $rawName.Normalize([Text.NormalizationForm]::FormC)
                if (-not $objectKeys.Peek().Add($name)) {
                    throw "Validation JSON contains duplicate normalized key '$name'."
                }
            } elseif ($reader.TokenType -eq [Newtonsoft.Json.JsonToken]::String) {
                $nodeCount++
                if (([string]$reader.Value).Length -gt $script:MaximumStringCharacters) {
                    throw "Validation JSON string exceeds $($script:MaximumStringCharacters) characters."
                }
            } elseif ($reader.TokenType -in @(
                    [Newtonsoft.Json.JsonToken]::Integer,
                    [Newtonsoft.Json.JsonToken]::Boolean,
                    [Newtonsoft.Json.JsonToken]::Null)) {
                $nodeCount++
            } elseif ($reader.TokenType -eq [Newtonsoft.Json.JsonToken]::Float) {
                throw 'Validation JSON rejects floating-point values.'
            } elseif ($reader.TokenType -notin @(
                    [Newtonsoft.Json.JsonToken]::EndArray,
                    [Newtonsoft.Json.JsonToken]::None)) {
                throw "Validation JSON contains unsupported token '$($reader.TokenType)'."
            }
            if ($nodeCount -gt $script:MaximumNodes) {
                throw "Validation JSON exceeds $($script:MaximumNodes) nodes."
            }
        }
    } finally {
        $reader.Dispose()
        $stringReader.Dispose()
    }
    return $true
}

function ConvertFrom-RSValidationJToken {
    param([Parameter(Mandatory = $true)][Newtonsoft.Json.Linq.JToken]$Token)

    if ($Token -is [Newtonsoft.Json.Linq.JObject]) {
        $result = [ordered]@{}
        foreach ($property in $Token.Properties()) {
            $result[$property.Name] = ConvertFrom-RSValidationJToken -Token $property.Value
        }
        return [pscustomobject]$result
    }
    if ($Token -is [Newtonsoft.Json.Linq.JArray]) {
        $items = @($Token.Children() | ForEach-Object { ConvertFrom-RSValidationJToken -Token $_ })
        return ,$items
    }
    if ($Token -is [Newtonsoft.Json.Linq.JValue]) {
        return $Token.Value
    }
    throw "Unsupported validation JSON token '$($Token.Type)'."
}

function Read-RSValidationJsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][string]$SchemaPath,
        [Parameter(Mandatory = $true)][string]$AllowedRoot
    )

    Assert-RSStartrailPowerShellHost
    $LiteralPath = Assert-RSValidationRecordPath -LiteralPath $LiteralPath -AllowedRoot $AllowedRoot
    Assert-RSValidationRecordBounds -LiteralPath $LiteralPath
    $bytes = [IO.File]::ReadAllBytes($LiteralPath)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf) {
        throw "Validation record must not contain a UTF-8 BOM: $LiteralPath"
    }
    $text = $script:Utf8NoBomStrict.GetString($bytes)
    if (-not $text.EndsWith("`n") -or $text.Contains("`r")) {
        throw "Validation record must use LF and end with exactly one LF: $LiteralPath"
    }
    $json = $text.Substring(0, $text.Length - 1)
    if ($json.EndsWith("`n")) { throw "Validation record has more than one trailing LF: $LiteralPath" }
    # Enforce complexity and string bounds with the streaming reader before
    # constructing a JToken graph from untrusted evidence.
    [void](Assert-RSValidationJsonTokenBounds -Json $json)
    if (-not (Test-Json -Json $json -SchemaFile $SchemaPath -ErrorAction SilentlyContinue)) {
        throw "Validation record does not satisfy schema '$SchemaPath': $LiteralPath"
    }
    $settings = [Newtonsoft.Json.JsonSerializerSettings]::new()
    $settings.DateParseHandling = [Newtonsoft.Json.DateParseHandling]::None
    $settings.FloatParseHandling = [Newtonsoft.Json.FloatParseHandling]::Decimal
    $settings.MaxDepth = $script:MaximumDepth
    $token = [Newtonsoft.Json.JsonConvert]::DeserializeObject($json, $settings)
    $record = ConvertFrom-RSValidationJToken -Token $token
    $canonical = ConvertTo-RSCanonicalJson -InputObject $record
    if (-not [string]::Equals($json, $canonical, [StringComparison]::Ordinal)) {
        throw "Validation record is not exact rs-canonical-json-v1: $LiteralPath"
    }
    return $record
}

function Get-RSValidationDependencyClosure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$RootRelativePath
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $tracked = @(& git -C $root ls-files --cached --others --exclude-standard)
    if ($LASTEXITCODE -ne 0) { throw 'Unable to enumerate validation dependency inputs with Git.' }
    $repositoryFiles = @($tracked | ForEach-Object { ([string]$_).Replace('\', '/') } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    $fileSet = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $filesByLeaf = [Collections.Generic.Dictionary[string, Collections.Generic.List[string]]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($repositoryFile in $repositoryFiles) {
        [void]$fileSet.Add($repositoryFile)
        $leaf = [IO.Path]::GetFileName($repositoryFile)
        if (-not $filesByLeaf.ContainsKey($leaf)) {
            $filesByLeaf.Add($leaf, [Collections.Generic.List[string]]::new())
        }
        $filesByLeaf[$leaf].Add($repositoryFile)
    }
    $pending = [Collections.Generic.Queue[string]]::new()
    $visited = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $normalizedRoot = $RootRelativePath.Replace('\', '/')
    if (-not $fileSet.Contains($normalizedRoot)) {
        throw "Validation dependency root is not a repository file: $RootRelativePath"
    }
    $pending.Enqueue($normalizedRoot)
    $quotedDependency = [regex]::new('(?<quote>[''"])(?<path>[^''"]+\.(?:ps1|psm1|json|props|targets))\k<quote>',
        [Text.RegularExpressions.RegexOptions]::IgnoreCase)

    while ($pending.Count -gt 0) {
        $relativePath = $pending.Dequeue()
        if (-not $visited.Add($relativePath)) { continue }
        if ($relativePath -notmatch '\.(?:ps1|psm1)$') { continue }
        $fullPath = Join-Path $root $relativePath
        $sourceDirectory = Split-Path $relativePath -Parent
        $source = Get-Content -LiteralPath $fullPath -Raw
        if ($null -eq $source) { $source = '' }
        foreach ($match in $quotedDependency.Matches($source)) {
            $literal = ([string]$match.Groups['path'].Value).Replace('\', '/').TrimStart('./')
            $candidates = [Collections.Generic.List[string]]::new()
            $candidates.Add($literal)
            if (-not [string]::IsNullOrWhiteSpace($sourceDirectory)) {
                $candidates.Add(($sourceDirectory.Replace('\', '/') + '/' + $literal))
            }
            $literalLeaf = [IO.Path]::GetFileName($literal)
            if ($filesByLeaf.ContainsKey($literalLeaf)) {
                foreach ($suffixMatch in $filesByLeaf[$literalLeaf]) {
                    if ($suffixMatch.EndsWith('/' + $literal, [StringComparison]::OrdinalIgnoreCase)) {
                        $candidates.Add($suffixMatch)
                    }
                }
            }
            $matches = @($candidates | ForEach-Object {
                    try {
                        [IO.Path]::GetRelativePath($root, [IO.Path]::GetFullPath((Join-Path $root $_))).Replace('\', '/')
                    } catch { $null }
                } | Where-Object { $null -ne $_ -and $fileSet.Contains($_) } | Sort-Object -Unique)
            if ($matches.Count -gt 1) {
                continue
            }
            if ($matches.Count -eq 1 -and $visited.Add($matches[0])) {
                $pending.Enqueue($matches[0])
                [void]$visited.Remove($matches[0])
            }
        }
    }
    return @($visited | Sort-Object)
}

function Write-RSAtomicCanonicalJsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][AllowNull()][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$SchemaPath,
        [Parameter(Mandatory = $true)][string]$AllowedRoot,
        [switch]$CreateNew
    )

    Assert-RSStartrailPowerShellHost
    $destination = Assert-RSValidationRecordPath -LiteralPath $LiteralPath -AllowedRoot $AllowedRoot
    $directory = Split-Path -Parent $destination

    $json = ConvertTo-RSCanonicalJson -InputObject $InputObject
    if (-not (Test-Json -Json $json -SchemaFile $SchemaPath -ErrorAction SilentlyContinue)) {
        throw "Validation record does not satisfy schema '$SchemaPath'."
    }
    $payload = $script:Utf8NoBomStrict.GetBytes($json + "`n")
    if ($payload.Length -gt $script:MaximumRecordBytes) { throw 'Validation record exceeds the byte limit.' }

    $temporary = Join-Path $directory ('.{0}.{1}.tmp' -f ([IO.Path]::GetFileName($destination)), [guid]::NewGuid().ToString('N'))
    try {
        $stream = [IO.FileStream]::new($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $stream.Write($payload, 0, $payload.Length)
            $stream.Flush($true)
        } finally {
            $stream.Dispose()
        }
        if ($CreateNew) {
            [IO.File]::Move($temporary, $destination)
        } elseif ([IO.File]::Exists($destination)) {
            $backup = Join-Path $directory ('.{0}.{1}.bak' -f ([IO.Path]::GetFileName($destination)), [guid]::NewGuid().ToString('N'))
            [IO.File]::Replace($temporary, $destination, $backup)
            Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue
        } else {
            [IO.File]::Move($temporary, $destination)
        }
    } finally {
        if ([IO.File]::Exists($temporary)) { Remove-Item -LiteralPath $temporary -Force }
    }
}

Export-ModuleMember -Function @(
    'Assert-RSStartrailPowerShellHost',
    'ConvertTo-RSCanonicalJson',
    'ConvertTo-RSCanonicalJsonBytes',
    'Get-RSCanonicalJsonDigest',
    'Test-RSValidationRelativePath',
    'Get-RSValidationReasonCodes',
    'Get-RSValidationDependencyClosure',
    'Assert-RSValidationIdentifier',
    'Assert-RSValidationPlanSemantics',
    'Assert-RSValidationPathSet',
    'Assert-RSValidationProvenance',
    'Get-RSWorkspaceSnapshot',
    'Get-RSValidationBuildReceiptBindingDigest',
    'Get-RSValidationEntryFingerprint',
    'Assert-RSValidationRecordBounds',
    'Read-RSValidationJsonFile',
    'Write-RSAtomicCanonicalJsonFile'
)
