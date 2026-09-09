Set-StrictMode -Version Latest

$script:TerminalLockFormat = 'red-salamander-terminal-engine-lock'
$script:TerminalLockVersion = 1

function Get-RSObjectProperty {
    param(
        [Parameter(Mandatory)]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Name,

        [object]$Default = $null,

        [switch]$Required
    )

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        if ($Required) {
            throw "Required property '$Name' is missing."
        }
        return $Default
    }
    return $property.Value
}

function Assert-RSObjectPropertySet {
    param(
        [Parameter(Mandatory)]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Context,

        [Parameter(Mandatory)]
        [string[]]$Allowed
    )

    foreach ($property in $Object.PSObject.Properties) {
        if ($Allowed -cnotcontains $property.Name) {
            throw "$Context contains unknown property '$($property.Name)'."
        }
    }
}

function Get-RSTerminalRepoRoot {
    return [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
}

function Test-RSPathIsWithin {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Root,

        [switch]$AllowRoot
    )

    $canonicalPath = [IO.Path]::GetFullPath($Path).TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar)
    $canonicalRoot = [IO.Path]::GetFullPath($Root).TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar)
    if ($AllowRoot -and
        [string]::Equals(
            $canonicalPath,
            $canonicalRoot,
            [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    return $canonicalPath.StartsWith(
        $canonicalRoot + [IO.Path]::DirectorySeparatorChar,
        [StringComparison]::OrdinalIgnoreCase)
}

function Assert-RSNoReparsePathComponents {
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $canonicalRoot = [IO.Path]::GetFullPath($Root)
    $canonicalPath = [IO.Path]::GetFullPath($Path)
    $current = $canonicalRoot
    $components = @([IO.Path]::GetRelativePath(
            $canonicalRoot,
            $canonicalPath).Split(
            [IO.Path]::DirectorySeparatorChar,
            [StringSplitOptions]::RemoveEmptyEntries))
    foreach ($component in @('.') + $components) {
        if ($component -ne '.') {
            $current = Join-Path $current $component
        }
        if (-not (Test-Path -LiteralPath $current)) {
            continue
        }
        $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Reparse points are forbidden in a locked or mutable lifecycle path: $current"
        }
    }
}

function Resolve-RSTerminalPathInRoot {
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [string]$Path,

        [switch]$MustExist,

        [switch]$AllowRoot
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'A relative path may not be empty.'
    }
    if ([IO.Path]::IsPathFullyQualified($Path)) {
        throw "An absolute path is not allowed in a terminal-engine lock: $Path"
    }

    $canonicalRoot = [IO.Path]::GetFullPath($Root)
    $candidate = [IO.Path]::GetFullPath((Join-Path $canonicalRoot $Path))
    if (-not (Test-RSPathIsWithin -Path $candidate -Root $canonicalRoot -AllowRoot:$AllowRoot)) {
        throw "Path escapes its allowed root '$canonicalRoot': $Path"
    }
    Assert-RSNoReparsePathComponents -Root $canonicalRoot -Path $candidate
    if ($MustExist -and -not (Test-Path -LiteralPath $candidate)) {
        throw "Required path does not exist: $candidate"
    }
    return $candidate
}

function Get-RSSha256 {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).
        Hash.ToLowerInvariant()
}

function Get-RSCanonicalObjectSha256 {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$InputObject
    )

    $evidenceModule = Join-Path (Get-RSTerminalRepoRoot) 'Tools\TerminalEvidence.psm1'
    Import-Module $evidenceModule -ErrorAction Stop
    $normalized = ConvertTo-RSJcsCompatibleObject -Value $InputObject
    return Get-RSJcsSha256 -InputObject $normalized
}

function ConvertTo-RSJcsCompatibleObject {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }
    if ($Value -is [Collections.IDictionary]) {
        $dictionary = [ordered]@{}
        foreach ($key in $Value.Keys) {
            $dictionary[[string]$key] = ConvertTo-RSJcsCompatibleObject -Value $Value[$key]
        }
        return $dictionary
    }
    if ($Value -is [Management.Automation.PSCustomObject]) {
        $dictionary = [ordered]@{}
        foreach ($property in $Value.PSObject.Properties) {
            $dictionary[$property.Name] = ConvertTo-RSJcsCompatibleObject `
                -Value $property.Value
        }
        return $dictionary
    }
    if ($Value -is [Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = [Collections.Generic.List[object]]::new()
        foreach ($item in $Value) {
            $items.Add((ConvertTo-RSJcsCompatibleObject -Value $item))
        }
        return ,$items.ToArray()
    }
    return $Value
}

function Get-RSDirectoryIdentity {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $root = [IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        throw "Directory input does not exist: $root"
    }

    $entries = [Collections.Generic.List[object]]::new()
    $caseFoldedPaths = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @(Get-ChildItem -LiteralPath $root -Recurse -Force -ErrorAction Stop)) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Directory input contains a reparse point: $($item.FullName)"
        }
        if ($item.PSIsContainer) {
            continue
        }
        $relative = [IO.Path]::GetRelativePath($root, $item.FullName).Replace('\', '/')
        if (-not $caseFoldedPaths.Add($relative)) {
            throw "Directory input contains a Windows case-folding collision: $relative"
        }
        $entries.Add([ordered]@{
                path = $relative
                sizeBytes = $item.Length.ToString([Globalization.CultureInfo]::InvariantCulture)
                sha256 = Get-RSSha256 -Path $item.FullName
            })
    }
    $orderedEntries = @($entries | Sort-Object -Stable -CaseSensitive -Property path)
    return [pscustomobject]@{
        sha256 = Get-RSCanonicalObjectSha256 -InputObject $orderedEntries
        fileCount = $orderedEntries.Count
        files = $orderedEntries
    }
}

function Resolve-RSTerminalGate0HostHeaderRoot {
    param(
        [Parameter(Mandatory)]
        [string]$HostHeaderIncludeRoot,

        [Parameter(Mandatory)]
        [ValidatePattern('^[0-9a-f]{64}$')]
        [string]$ExpectedTreeSha256
    )

    $root = [IO.Path]::GetFullPath($HostHeaderIncludeRoot)
    if (-not [IO.Directory]::Exists($root)) {
        throw [IO.DirectoryNotFoundException]::new(
            "Gate-0 host-header include root does not exist: $root")
    }
    $fileSystemRoot = [IO.Path]::GetPathRoot($root)
    if ([string]::IsNullOrEmpty($fileSystemRoot)) {
        throw [IO.InvalidDataException]::new(
            'Gate-0 host-header include root has no filesystem root.')
    }
    Assert-RSNoReparsePathComponents `
        -Root $fileSystemRoot `
        -Path $root

    $wilDirectory = Join-Path $root 'wil'
    $resourceHeader = Join-Path $wilDirectory 'resource.h'
    Assert-RSNoReparsePathComponents -Root $root -Path $resourceHeader
    if (-not [IO.Directory]::Exists($wilDirectory) -or
        (Get-Item -LiteralPath $wilDirectory -Force).Name -cne 'wil' -or
        -not [IO.File]::Exists($resourceHeader) -or
        (Get-Item -LiteralPath $resourceHeader -Force).Name -cne 'resource.h') {
        throw [IO.InvalidDataException]::new(
            "Gate-0 host-header include root must contain exact 'wil/resource.h'.")
    }

    $identity = Get-RSDirectoryIdentity -Path $root
    if ([string]$identity.sha256 -cne $ExpectedTreeSha256) {
        throw [IO.InvalidDataException]::new(
            'Gate-0 host-header tree differs from its frozen identity.')
    }
    return [pscustomobject]@{
        IncludeRoot = $root
        ResourceHeaderPath = [IO.Path]::GetFullPath($resourceHeader)
        TreeSha256 = [string]$identity.sha256
        FileCount = [int]$identity.fileCount
    }
}

function Assert-RSLowerSha256 {
    param(
        [Parameter(Mandatory)]
        [string]$Value,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($Value -cnotmatch '^[0-9a-f]{64}$') {
        throw "$Name must be exactly 64 lowercase hexadecimal characters."
    }
}

function Resolve-RSTerminalInputPath {
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [Parameter(Mandatory)]
        [object]$InputRecord
    )

    $source = Get-RSObjectProperty -Object $InputRecord -Name source -Required
    $sourceType = [string](Get-RSObjectProperty -Object $source -Name type -Required)
    if ($sourceType -ceq 'repository-file') {
        return Resolve-RSTerminalPathInRoot `
            -Root $Lock.repoRoot `
            -Path ([string](Get-RSObjectProperty -Object $source -Name path -Required)) `
            -MustExist
    }
    if ($sourceType -ceq 'https') {
        [void](Assert-RSTerminalHttpsInputUri `
                -Value ([string](Get-RSObjectProperty -Object $source -Name uri -Required)))
        $cachePath = Get-RSTerminalInputCachePath `
            -RepoRoot $Lock.repoRoot `
            -Sha256 ([string](Get-RSObjectProperty -Object $InputRecord -Name sha256 -Required))
        if (-not (Test-Path -LiteralPath $cachePath -PathType Leaf)) {
            throw "Cached input '$($InputRecord.id)' is missing: $cachePath"
        }
        return $cachePath
    }
    throw "Input '$($InputRecord.id)' has unsupported source type '$sourceType'."
}

function Assert-RSTerminalHttpsInputUri {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    $parsed = $null
    if (-not [Uri]::TryCreate($Value, [UriKind]::Absolute, [ref]$parsed) -or
        $parsed.Scheme -cne 'https' -or
        [string]::IsNullOrWhiteSpace($parsed.Host) -or
        -not [string]::IsNullOrEmpty($parsed.UserInfo) -or
        -not [string]::IsNullOrEmpty($parsed.Query) -or
        -not [string]::IsNullOrEmpty($parsed.Fragment) -or
        $parsed.AbsoluteUri -cne $Value) {
        throw 'HTTPS input URIs must be canonical, credential-free HTTPS URLs without query or fragment.'
    }
    return $parsed
}

function Get-RSTerminalInputCacheRoot {
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot
    )

    return [IO.Path]::GetFullPath(
        (Join-Path $RepoRoot '.build\TerminalEngineInputCache\v1\sha256'))
}

function Get-RSTerminalInputCachePath {
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot,

        [Parameter(Mandatory)]
        [string]$Sha256
    )

    Assert-RSLowerSha256 -Value $Sha256 -Name 'input cache SHA-256'
    $cacheRoot = Get-RSTerminalInputCacheRoot -RepoRoot $RepoRoot
    $path = [IO.Path]::GetFullPath((Join-Path $cacheRoot $Sha256))
    if (-not (Test-RSPathIsWithin -Path $path -Root $cacheRoot)) {
        throw 'The terminal-engine input cache path escaped its fixed root.'
    }
    Assert-RSNoReparsePathComponents -Root (Join-Path $RepoRoot '.build') -Path $path
    return $path
}

function Test-RSTerminalInputIdentity {
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [Parameter(Mandatory)]
        [object]$InputRecord
    )

    $id = [string](Get-RSObjectProperty -Object $InputRecord -Name id -Required)
    $kind = [string](Get-RSObjectProperty -Object $InputRecord -Name kind -Required)
    $expectedSha = [string](Get-RSObjectProperty -Object $InputRecord -Name sha256 -Required)
    Assert-RSLowerSha256 -Value $expectedSha -Name "inputs[$id].sha256"
    $path = Resolve-RSTerminalInputPath -Lock $Lock -InputRecord $InputRecord

    if ($kind -eq 'directory') {
        $identity = Get-RSDirectoryIdentity -Path $path
        $actualSha = $identity.sha256
        $size = $null
    }
    elseif ($kind -in @('file', 'zip')) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Input '$id' must resolve to a file."
        }
        $actualSha = Get-RSSha256 -Path $path
        $size = (Get-Item -LiteralPath $path -Force).Length
        $expectedSize = [string](Get-RSObjectProperty -Object $InputRecord -Name sizeBytes -Required)
        if ($expectedSize -cnotmatch '^(0|[1-9][0-9]*)$' -or
            [uint64]$expectedSize -ne [uint64]$size) {
            throw "Input '$id' size differs from the lock."
        }
    }
    else {
        throw "Input '$id' has unsupported kind '$kind'."
    }

    if (-not [string]::Equals($actualSha, $expectedSha, [StringComparison]::Ordinal)) {
        throw "Input '$id' SHA-256 differs from the lock."
    }

    return [pscustomobject]@{
        id = $id
        kind = $kind
        sourceType = [string]$InputRecord.source.type
        path = [IO.Path]::GetRelativePath($Lock.repoRoot, $path).Replace('\', '/')
        sha256 = $actualSha
        sizeBytes = if ($null -eq $size) { $null } else {
            ([uint64]$size).ToString([Globalization.CultureInfo]::InvariantCulture)
        }
    }
}

function Restore-RSTerminalEngineInputs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [switch]$Offline
    )

    if ($null -eq $Lock.PSObject.Properties['inputById']) {
        throw 'Input restore requires a lock returned by Read-RSTerminalEngineLock.'
    }

    $records = [Collections.Generic.List[object]]::new()
    foreach ($inputRecord in @($Lock.inputs)) {
        $id = [string]$inputRecord.id
        $sourceType = [string]$inputRecord.source.type
        if ($sourceType -ceq 'repository-file') {
            $verified = Test-RSTerminalInputIdentity -Lock $Lock -InputRecord $inputRecord
            $records.Add([ordered]@{
                    id = $id
                    sourceType = $sourceType
                    status = 'verified'
                    sha256 = [string]$verified.sha256
                    sizeBytes = $verified.sizeBytes
                })
            continue
        }

        $cachePath = Get-RSTerminalInputCachePath `
            -RepoRoot $Lock.repoRoot `
            -Sha256 ([string]$inputRecord.sha256)
        if (Test-Path -LiteralPath $cachePath -PathType Leaf) {
            $verified = Test-RSTerminalInputIdentity -Lock $Lock -InputRecord $inputRecord
            $records.Add([ordered]@{
                    id = $id
                    sourceType = $sourceType
                    status = 'cached'
                    sha256 = [string]$verified.sha256
                    sizeBytes = $verified.sizeBytes
                })
            continue
        }
        if ($Offline) {
            throw "Cached input '$id' is missing in offline mode: $cachePath"
        }

        $uri = Assert-RSTerminalHttpsInputUri -Value ([string]$inputRecord.source.uri)
        $expectedSizeText = [string](
            Get-RSObjectProperty -Object $inputRecord -Name sizeBytes -Required)
        if ($expectedSizeText -cnotmatch '^(0|[1-9][0-9]*)$') {
            throw "Input '$id' sizeBytes must be an unsigned canonical decimal string."
        }
        $expectedSize = [uint64]$expectedSizeText
        $cacheRoot = Get-RSTerminalInputCacheRoot -RepoRoot $Lock.repoRoot
        [void](New-Item -ItemType Directory -Path $cacheRoot -Force)
        Assert-RSNoReparsePathComponents `
            -Root (Join-Path $Lock.repoRoot '.build') `
            -Path $cacheRoot
        $temporary = Join-Path $cacheRoot (
            ".$($inputRecord.sha256).$([Guid]::NewGuid().ToString('N')).tmp")
        $handler = [Net.Http.HttpClientHandler]::new()
        $handler.AllowAutoRedirect = $false
        $handler.UseDefaultCredentials = $false
        $handler.Credentials = $null
        $handler.UseProxy = $false
        $handler.Proxy = $null
        $client = [Net.Http.HttpClient]::new($handler)
        # ResponseHeadersRead completes as soon as the headers arrive, so
        # HttpClient.Timeout alone would not bound a slow or stalled response
        # body. One cancellation token covers headers and every body read.
        $client.Timeout = [Threading.Timeout]::InfiniteTimeSpan
        $downloadCancellation = [Threading.CancellationTokenSource]::new(
            [TimeSpan]::FromMinutes(10))
        try {
            $request = [Net.Http.HttpRequestMessage]::new(
                [Net.Http.HttpMethod]::Get,
                $uri)
            try {
                $response = $client.Send(
                    $request,
                    [Net.Http.HttpCompletionOption]::ResponseHeadersRead,
                    $downloadCancellation.Token)
            }
            finally {
                $request.Dispose()
            }
            try {
                if ([int]$response.StatusCode -ne 200) {
                    throw "HTTPS input '$id' returned HTTP status $([int]$response.StatusCode); redirects are forbidden."
                }
                if ($null -ne $response.Content.Headers.ContentLength -and
                    [uint64]$response.Content.Headers.ContentLength -ne $expectedSize) {
                    throw "HTTPS input '$id' Content-Length differs from the lock."
                }
                $inputStream = $response.Content.ReadAsStream()
                $outputStream = [IO.FileStream]::new(
                    $temporary,
                    [IO.FileMode]::CreateNew,
                    [IO.FileAccess]::Write,
                    [IO.FileShare]::None)
                try {
                    $buffer = [byte[]]::new(1MB)
                    [uint64]$written = 0
                    while (($count = $inputStream.ReadAsync(
                                    $buffer,
                                    0,
                                    $buffer.Length,
                                    $downloadCancellation.Token).
                                GetAwaiter().GetResult()) -gt 0) {
                        $written += [uint64]$count
                        if ($written -gt $expectedSize) {
                            throw "HTTPS input '$id' exceeded its locked size."
                        }
                        $outputStream.Write($buffer, 0, $count)
                    }
                    if ($written -ne $expectedSize) {
                        throw "HTTPS input '$id' size differs from the lock."
                    }
                    $outputStream.Flush($true)
                }
                finally {
                    $outputStream.Dispose()
                    $inputStream.Dispose()
                }
            }
            finally {
                $response.Dispose()
            }

            if ((Get-RSSha256 -Path $temporary) -cne [string]$inputRecord.sha256) {
                throw "HTTPS input '$id' SHA-256 differs from the lock."
            }
            try {
                [IO.File]::Move($temporary, $cachePath, $false)
            }
            catch [IO.IOException] {
                if (-not (Test-Path -LiteralPath $cachePath -PathType Leaf)) {
                    throw
                }
            }
            [void](Test-RSTerminalInputIdentity -Lock $Lock -InputRecord $inputRecord)
            $records.Add([ordered]@{
                    id = $id
                    sourceType = $sourceType
                    status = 'restored'
                    sha256 = [string]$inputRecord.sha256
                    sizeBytes = $expectedSizeText
                })
        }
        catch [OperationCanceledException] {
            throw "HTTPS input '$id' timed out after 600 seconds."
        }
        finally {
            $downloadCancellation.Dispose()
            $client.Dispose()
            $handler.Dispose()
            if (Test-Path -LiteralPath $temporary -PathType Leaf) {
                Remove-Item -LiteralPath $temporary -Force
            }
        }
    }
    return $records.ToArray()
}

function Read-RSTerminalEngineLock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LockFile,

        [switch]$ValidateInputs
    )

    $repoRoot = Get-RSTerminalRepoRoot
    $candidate = if ([IO.Path]::IsPathFullyQualified($LockFile)) {
        [IO.Path]::GetFullPath($LockFile)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $repoRoot $LockFile))
    }
    if (-not (Test-RSPathIsWithin -Path $candidate -Root $repoRoot)) {
        throw "LockFile must be inside the repository, including its ignored .build tree: $candidate"
    }
    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        throw "Terminal engine lock does not exist: $candidate"
    }

    $lockBytes = [IO.File]::ReadAllBytes($candidate)
    if ($lockBytes.Length -eq 0 -or
        ($lockBytes.Length -ge 3 -and
            $lockBytes[0] -eq 0xef -and
            $lockBytes[1] -eq 0xbb -and
            $lockBytes[2] -eq 0xbf)) {
        throw 'Terminal engine lock must be nonempty UTF-8 without BOM.'
    }
    try {
        $lockJson = [Text.UTF8Encoding]::new($false, $true).GetString($lockBytes)
        $lockForCanonical = $lockJson |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String `
                -ErrorAction Stop
        $lock = $lockJson |
            ConvertFrom-Json -Depth 100 -NoEnumerate -DateKind String -ErrorAction Stop
    }
    catch {
        throw "Terminal engine lock is not duplicate-free UTF-8 JSON: $($_.Exception.Message)"
    }

    $evidenceModule = Join-Path $repoRoot 'Tools\TerminalEvidence.psm1'
    Import-Module $evidenceModule -ErrorAction Stop
    $canonicalLockBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $lockForCanonical
    if (-not [Linq.Enumerable]::SequenceEqual(
            [byte[]]$lockBytes,
            [byte[]]$canonicalLockBytes)) {
        throw 'Terminal engine lock must be exact RFC 8785 JCS bytes.'
    }
    if ([string](Get-RSObjectProperty -Object $lock -Name formatId -Required) -cne
        $script:TerminalLockFormat) {
        throw "Terminal engine lock formatId must be '$script:TerminalLockFormat'."
    }
    if ([int](Get-RSObjectProperty -Object $lock -Name schemaVersion -Required) -ne
        $script:TerminalLockVersion) {
        throw "Terminal engine lock schemaVersion must be $script:TerminalLockVersion."
    }
    Assert-RSObjectPropertySet `
        -Object $lock `
        -Context 'Terminal engine lock' `
        -Allowed @(
            'formatId',
            'schemaVersion',
            'engine',
            'inputs',
            'build',
            'sourceAdaptation',
            'derivedInputs',
            'installedToolchains',
            'lanes',
            'outputs',
            'dependencies',
            'abi',
            'runtimeClosure',
            'qualification',
            'upgrade')

    $engine = Get-RSObjectProperty -Object $lock -Name engine -Required
    Assert-RSObjectPropertySet `
        -Object $engine `
        -Context 'engine' `
        -Allowed @('candidateId', 'pin', 'pluginLinkage', 'privateRuntimePath')
    foreach ($name in @('candidateId', 'pin', 'pluginLinkage')) {
        if ([string]::IsNullOrWhiteSpace(
                [string](Get-RSObjectProperty -Object $engine -Name $name -Required))) {
            throw "engine.$name may not be empty."
        }
    }
    $privateRuntimePathProperty = $engine.PSObject.Properties['privateRuntimePath']
    if ($null -eq $privateRuntimePathProperty) {
        throw 'engine.privateRuntimePath must be present and explicitly string or null.'
    }
    $linkage = [string]$engine.pluginLinkage
    if ($linkage -ceq 'private-dynamic') {
        if ($privateRuntimePathProperty.Value -isnot [string] -or
            [string]::IsNullOrWhiteSpace([string]$privateRuntimePathProperty.Value)) {
            throw 'A private-dynamic engine requires engine.privateRuntimePath.'
        }
        $privateRuntimePath = [string]$privateRuntimePathProperty.Value
        if ($privateRuntimePath.Replace('\', '/') -cnotmatch
            '^Plugins/TerminalRuntime/[A-Za-z0-9._+-]+\.dll$') {
            throw 'engine.privateRuntimePath must be one DLL below Plugins/TerminalRuntime/.'
        }
    }
    elseif ($linkage -ceq 'source-static') {
        if ($null -ne $privateRuntimePathProperty.Value) {
            throw 'A source-static engine requires engine.privateRuntimePath=null.'
        }
    }
    else {
        throw "engine.pluginLinkage must be 'private-dynamic' or 'source-static'."
    }

    $inputs = @(Get-RSObjectProperty -Object $lock -Name inputs -Required)
    if ($inputs.Count -eq 0) {
        throw 'inputs[] must contain the complete bootstrap input set.'
    }
    $inputById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($inputRecord in $inputs) {
        $id = [string](Get-RSObjectProperty -Object $inputRecord -Name id -Required)
        if ($null -ne $inputRecord.PSObject.Properties['path']) {
            throw "Input '$id' must use the closed source object; legacy top-level path is forbidden."
        }
        Assert-RSObjectPropertySet `
            -Object $inputRecord `
            -Context "inputs[$id]" `
            -Allowed @('id', 'category', 'kind', 'source', 'sha256', 'sizeBytes')
        if ($id -cnotmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$') {
            throw "Invalid input id '$id'."
        }
        if ($inputById.ContainsKey($id)) {
            throw "Duplicate input id '$id'."
        }
        $inputKind = [string](Get-RSObjectProperty `
                -Object $inputRecord `
                -Name kind `
                -Required)
        if ($inputKind -cnotin @('file', 'zip')) {
            throw "Input '$id' kind must be file or zip; derived directories belong in derivedInputs."
        }
        $inputById.Add($id, $inputRecord)
        $source = Get-RSObjectProperty -Object $inputRecord -Name source -Required
        $sourceType = [string](Get-RSObjectProperty -Object $source -Name type -Required)
        if ($sourceType -ceq 'repository-file') {
            if (@($source.PSObject.Properties).Count -ne 2 -or
                $null -eq $source.PSObject.Properties['path']) {
                throw "Repository input '$id' source must contain exactly type and path."
            }
            [void](Resolve-RSTerminalPathInRoot `
                    -Root $repoRoot `
                    -Path ([string]$source.path))
        }
        elseif ($sourceType -ceq 'https') {
            if (@($source.PSObject.Properties).Count -ne 2 -or
                $null -eq $source.PSObject.Properties['uri']) {
                throw "HTTPS input '$id' source must contain exactly type and uri."
            }
            [void](Assert-RSTerminalHttpsInputUri -Value ([string]$source.uri))
        }
        else {
            throw "Input '$id' has unsupported source type '$sourceType'."
        }
        Assert-RSLowerSha256 `
            -Value ([string](Get-RSObjectProperty -Object $inputRecord -Name sha256 -Required)) `
            -Name "inputs[$id].sha256"
        $sizeText = [string](Get-RSObjectProperty `
                -Object $inputRecord `
                -Name sizeBytes `
                -Required)
        if ($sizeText -cnotmatch '^(0|[1-9][0-9]*)$') {
            throw "Input '$id' sizeBytes must be an unsigned canonical decimal string."
        }
    }

    $build = Get-RSObjectProperty -Object $lock -Name build -Required
    Assert-RSObjectPropertySet `
        -Object $build `
        -Context 'build' `
        -Allowed @(
            'sourceInputId',
            'materializations',
            'executableRelativePath',
            'workingDirectoryRelativePath',
            'environment')
    $commonMaterializations = @(Get-RSObjectProperty `
            -Object $build `
            -Name materializations `
            -Default @())
    $lanes = @(Get-RSObjectProperty -Object $lock -Name lanes -Required)
    $outputs = @(Get-RSObjectProperty -Object $lock -Name outputs -Required)
    if ($lanes.Count -eq 0 -or $outputs.Count -eq 0) {
        throw 'lanes[] and outputs[] must both be nonempty.'
    }

    $laneByKey = @{}
    foreach ($lane in $lanes) {
        $platform = [string](Get-RSObjectProperty -Object $lane -Name platform -Required)
        $configuration = [string](Get-RSObjectProperty -Object $lane -Name configuration -Required)
        Assert-RSObjectPropertySet `
            -Object $lane `
            -Context "lane[$platform/$configuration]" `
            -Allowed @(
                'platform',
                'configuration',
                'reuseFrom',
                'materializations',
                'executableRelativePath',
                'workingDirectoryRelativePath',
                'arguments',
                'timeoutSeconds',
                'environment',
                'outputIds')
        if ($platform -cnotin @('x64', 'ARM64') -or
            $configuration -cnotin @('Debug', 'Release', 'ASan Debug')) {
            throw "Unsupported lane '$platform/$configuration'."
        }
        $key = "$platform`0$configuration"
        if ($laneByKey.ContainsKey($key)) {
            throw "Duplicate lane '$platform/$configuration'."
        }
        $laneByKey[$key] = $lane
        $reuseProperty = $lane.PSObject.Properties['reuseFrom']
        if ($null -eq $reuseProperty) {
            throw "Lane '$platform/$configuration' must explicitly declare reuseFrom (object or null)."
        }
        $reuseFrom = $reuseProperty.Value
        if ($null -eq $reuseFrom) {
            $executable = [string](Get-RSObjectProperty `
                    -Object $lane `
                    -Name executableRelativePath `
                    -Default (Get-RSObjectProperty `
                        -Object $build `
                        -Name executableRelativePath `
                        -Default ''))
            if ([string]::IsNullOrWhiteSpace($executable)) {
                throw "Lane '$platform/$configuration' has no explicit build executable."
            }
            [void](Resolve-RSTerminalPathInRoot -Root $repoRoot -Path $executable)
            $arguments = @(Get-RSObjectProperty -Object $lane -Name arguments -Required)
            if ($arguments.Count -eq 0) {
                throw "Lane '$platform/$configuration' build arguments may not be empty."
            }
        }
        elseif ($configuration -cne 'ASan Debug') {
            throw "Only an ASan Debug lane may reuse another engine artifact."
        }
        else {
            Assert-RSObjectPropertySet `
                -Object $reuseFrom `
                -Context "lane[$platform/$configuration].reuseFrom" `
                -Allowed @('platform', 'configuration')
        }

        foreach ($materialization in @(
                $commonMaterializations +
                @(Get-RSObjectProperty -Object $lane -Name materializations -Default @()))) {
            $inputId = [string](Get-RSObjectProperty `
                    -Object $materialization `
                    -Name inputId `
                    -Required)
            Assert-RSObjectPropertySet `
                -Object $materialization `
                -Context "materialization[$inputId]" `
                -Allowed @(
                    'inputId',
                    'mode',
                    'destination',
                    'derivedInputId',
                    'contentRoot',
                    'contentSha256')
            if (-not $inputById.ContainsKey($inputId)) {
                throw "Materialization refers to unknown input '$inputId'."
            }
            $mode = [string](Get-RSObjectProperty `
                    -Object $materialization `
                    -Name mode `
                    -Required)
            if ($mode -ceq 'extract-zip') {
                $contentSha256 = [string](Get-RSObjectProperty `
                        -Object $materialization `
                        -Name contentSha256 `
                        -Required)
                Assert-RSLowerSha256 `
                    -Value $contentSha256 `
                    -Name "materialization[$inputId].contentSha256"
            }
            [void](Resolve-RSTerminalPathInRoot `
                    -Root $repoRoot `
                    -Path ([string](Get-RSObjectProperty `
                        -Object $materialization `
                        -Name destination `
                        -Required)))
        }
    }

    $outputById = @{}
    foreach ($output in $outputs) {
        $id = [string](Get-RSObjectProperty -Object $output -Name id -Required)
        Assert-RSObjectPropertySet `
            -Object $output `
            -Context "outputs[$id]" `
            -Allowed @(
                'id',
                'platform',
                'configuration',
                'kind',
                'builtRelativePath',
                'sourceOutputId',
                'relativePath',
                'rawSha256Policy',
                'sha256',
                'semanticSha256',
                'machine',
                'requiredExports',
                'runtimeIdentity',
                'smoke')
        if ($outputById.ContainsKey($id)) {
            throw "Duplicate output id '$id'."
        }
        $outputById[$id] = $output
        $platform = [string](Get-RSObjectProperty -Object $output -Name platform -Required)
        $configuration = [string](Get-RSObjectProperty `
                -Object $output `
                -Name configuration `
                -Required)
        if (-not $laneByKey.ContainsKey("$platform`0$configuration")) {
            throw "Output '$id' refers to an unknown lane."
        }
        [void](Resolve-RSTerminalPathInRoot `
                -Root $repoRoot `
                -Path ([string](Get-RSObjectProperty -Object $output -Name relativePath -Required)))
        $kind = [string](Get-RSObjectProperty -Object $output -Name kind -Required)
        if ($kind -cnotin @('pe', 'file')) {
            throw "Output '$id' has unsupported kind '$kind'."
        }
        $policy = [string](Get-RSObjectProperty -Object $output -Name rawSha256Policy -Required)
        if ($policy -cne 'exact') {
            throw "Output '$id' must use exact raw SHA-256 verification."
        }
        Assert-RSLowerSha256 `
            -Value ([string](Get-RSObjectProperty -Object $output -Name sha256 -Required)) `
            -Name "outputs[$id].sha256"
        if ($kind -eq 'pe') {
            Assert-RSLowerSha256 `
                -Value ([string](Get-RSObjectProperty `
                    -Object $output `
                    -Name semanticSha256 `
                    -Required)) `
                -Name "outputs[$id].semanticSha256"
        }
        $runtimeIdentity = Get-RSObjectProperty `
            -Object $output `
            -Name runtimeIdentity `
            -Default $null
        if ($null -ne $runtimeIdentity) {
            Assert-RSObjectPropertySet `
                -Object $runtimeIdentity `
                -Context "outputs[$id].runtimeIdentity" `
                -Allowed @(
                    'upstreamCommit',
                    'patchAbi',
                    'buildIdentity',
                    'architecture',
                    'callingConvention',
                    'capabilitiesHex',
                    'pointerSize',
                    'terminalResourceLimitsSize')
        }
        $smoke = Get-RSObjectProperty -Object $output -Name smoke -Default $null
        if ($null -ne $smoke) {
            Assert-RSObjectPropertySet `
                -Object $smoke `
                -Context "outputs[$id].smoke" `
                -Allowed @(
                    'executableInputId',
                    'arguments',
                    'timeoutSeconds',
                    'expectedStdoutSha256',
                    'expectedStderrSha256')
        }
    }
    foreach ($lane in $lanes) {
        foreach ($outputId in @(Get-RSObjectProperty -Object $lane -Name outputIds -Required)) {
            if (-not $outputById.ContainsKey([string]$outputId)) {
                throw "Lane refers to unknown output '$outputId'."
            }
            $output = $outputById[[string]$outputId]
            if ([string]$output.platform -cne [string]$lane.platform -or
                [string]$output.configuration -cne [string]$lane.configuration) {
                throw "Output '$outputId' does not belong to its declaring lane."
            }
        }
    }

    foreach ($requiredTopLevel in @('dependencies', 'abi', 'runtimeClosure')) {
        if ($null -eq $lock.PSObject.Properties[$requiredTopLevel]) {
            throw "Terminal engine lock must declare top-level $requiredTopLevel."
        }
    }
    foreach ($requiredTopLevel in @(
            'sourceAdaptation',
            'derivedInputs',
            'installedToolchains')) {
        if ($null -eq $lock.PSObject.Properties[$requiredTopLevel]) {
            throw "Terminal engine lock must declare top-level $requiredTopLevel."
        }
    }
    $sourceAdaptation = $lock.sourceAdaptation
    Assert-RSObjectPropertySet `
        -Object $sourceAdaptation `
        -Context 'sourceAdaptation' `
        -Allowed @(
            'upstreamTree',
            'patchSha256',
            'resultTree',
            'adaptedLineCount',
            'adaptedFileCount',
            'concernMappingSha256')
    foreach ($name in @('upstreamTree', 'resultTree')) {
        $tree = Get-RSObjectProperty `
            -Object $sourceAdaptation `
            -Name $name `
            -Required
        Assert-RSObjectPropertySet `
            -Object $tree `
            -Context "sourceAdaptation.$name" `
            -Allowed @('algorithm', 'digest')
        $algorithm = [string](Get-RSObjectProperty `
                -Object $tree `
                -Name algorithm `
                -Required)
        $digest = [string](Get-RSObjectProperty `
                -Object $tree `
                -Name digest `
                -Required)
        $pattern = if ($algorithm -ceq 'git-sha1') {
            '^[0-9a-f]{40}$'
        }
        elseif ($algorithm -ceq 'git-sha256') {
            '^[0-9a-f]{64}$'
        }
        else {
            throw "sourceAdaptation.$name has unsupported Git object algorithm '$algorithm'."
        }
        if ($digest -cnotmatch $pattern) {
            throw "sourceAdaptation.$name digest does not match $algorithm."
        }
    }
    foreach ($name in @('patchSha256', 'concernMappingSha256')) {
        Assert-RSLowerSha256 `
            -Value ([string](Get-RSObjectProperty `
                    -Object $sourceAdaptation `
                    -Name $name `
                    -Required)) `
            -Name "sourceAdaptation.$name"
    }
    foreach ($name in @('adaptedLineCount', 'adaptedFileCount')) {
        $countText = [string](Get-RSObjectProperty `
                -Object $sourceAdaptation `
                -Name $name `
                -Required)
        if ($countText -cnotmatch '^[1-9][0-9]*$') {
            throw "sourceAdaptation.$name must be a positive canonical decimal string."
        }
        $maximum = if ($name -ceq 'adaptedLineCount') { [uint64]2500 } else { [uint64]16 }
        if ([uint64]$countText -gt $maximum) {
            throw "sourceAdaptation.$name exceeds the frozen adaptation cap of $maximum."
        }
    }

    $derivedInputIds = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::Ordinal)
    $derivedInputs = @(Get-RSObjectProperty `
            -Object $lock `
            -Name derivedInputs `
            -Required)
    if ($derivedInputs.Count -eq 0) {
        throw 'derivedInputs[] must bind at least one exact derived file set.'
    }
    foreach ($derived in $derivedInputs) {
        $derivedId = [string](Get-RSObjectProperty -Object $derived -Name id -Required)
        Assert-RSObjectPropertySet `
            -Object $derived `
            -Context "derivedInputs[$derivedId]" `
            -Allowed @('id', 'kind', 'inputIds', 'sha256', 'fileCount')
        if ($derivedId -cnotmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$' -or
            -not $derivedInputIds.Add($derivedId)) {
            throw "Invalid or duplicate derived input id '$derivedId'."
        }
        if ([string](Get-RSObjectProperty -Object $derived -Name kind -Required) -cne
            'extracted-file-set') {
            throw "Derived input '$derivedId' must use kind extracted-file-set."
        }
        Assert-RSLowerSha256 `
            -Value ([string](Get-RSObjectProperty -Object $derived -Name sha256 -Required)) `
            -Name "derivedInputs[$derivedId].sha256"
        if ([string](Get-RSObjectProperty `
                -Object $derived `
                -Name fileCount `
                -Required) -cnotmatch '^[1-9][0-9]*$') {
            throw "Derived input '$derivedId' fileCount must be a positive canonical decimal string."
        }
        $derivedSources = @(Get-RSObjectProperty `
                -Object $derived `
                -Name inputIds `
                -Required)
        if ($derivedSources.Count -eq 0) {
            throw "Derived input '$derivedId' inputIds may not be empty."
        }
        $derivedSourceIds = [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
        foreach ($inputId in $derivedSources) {
            if (-not $derivedSourceIds.Add([string]$inputId)) {
                throw "Derived input '$derivedId' has duplicate inputId '$inputId'."
            }
            if (-not $inputById.ContainsKey([string]$inputId)) {
                throw "Derived input '$derivedId' refers to unknown input '$inputId'."
            }
        }
    }
    foreach ($materialization in @(
            $commonMaterializations +
            @($lanes | ForEach-Object {
                    @(Get-RSObjectProperty `
                        -Object $_ `
                        -Name materializations `
                        -Default @())
                }))) {
        if ([string]$materialization.mode -cne 'extract-zip') {
            if ($null -ne $materialization.PSObject.Properties['derivedInputId']) {
                throw 'Only extract-zip materializations may bind derivedInputId.'
            }
            continue
        }
        $derivedId = [string](Get-RSObjectProperty `
                -Object $materialization `
                -Name derivedInputId `
                -Required)
        $derived = @($derivedInputs | Where-Object {
                [string]$_.id -ceq $derivedId
            })
        if ($derived.Count -ne 1) {
            throw "Extracted materialization refers to unknown derived input '$derivedId'."
        }
        if ([string]$derived[0].sha256 -cne [string]$materialization.contentSha256 -or
            @($derived[0].inputIds) -cnotcontains [string]$materialization.inputId) {
            throw "Extracted materialization '$derivedId' differs from its derived-input identity."
        }
    }

    $toolchainIds = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::Ordinal)
    $installedToolchains = @(Get-RSObjectProperty `
            -Object $lock `
            -Name installedToolchains `
            -Required)
    if ($installedToolchains.Count -eq 0) {
        throw 'installedToolchains[] must bind the exact installed build toolchain.'
    }
    foreach ($toolchain in $installedToolchains) {
        $toolchainId = [string](Get-RSObjectProperty `
                -Object $toolchain `
                -Name id `
                -Required)
        Assert-RSObjectPropertySet `
            -Object $toolchain `
            -Context "installedToolchains[$toolchainId]" `
            -Allowed @('id', 'kind', 'version', 'identitySha256', 'platforms')
        if ($toolchainId -cnotmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$' -or
            -not $toolchainIds.Add($toolchainId)) {
            throw "Invalid or duplicate installed toolchain id '$toolchainId'."
        }
        foreach ($name in @('kind', 'version')) {
            if ([string]::IsNullOrWhiteSpace(
                    [string](Get-RSObjectProperty `
                        -Object $toolchain `
                        -Name $name `
                        -Required))) {
                throw "Installed toolchain '$toolchainId' $name may not be empty."
            }
        }
        Assert-RSLowerSha256 `
            -Value ([string](Get-RSObjectProperty `
                    -Object $toolchain `
                    -Name identitySha256 `
                    -Required)) `
            -Name "installedToolchains[$toolchainId].identitySha256"
        $platforms = @(Get-RSObjectProperty `
                -Object $toolchain `
                -Name platforms `
                -Required)
        if ($platforms.Count -eq 0 -or
            @($platforms | Where-Object { [string]$_ -cnotin @('x64', 'ARM64') }).Count -gt 0) {
            throw "Installed toolchain '$toolchainId' platforms are invalid."
        }
        if (@($platforms | Select-Object -Unique).Count -ne $platforms.Count) {
            throw "Installed toolchain '$toolchainId' platforms contain duplicates."
        }
    }
    $dependencies = @(Get-RSObjectProperty `
            -Object $lock `
            -Name dependencies `
            -Required)
    if ($dependencies.Count -eq 0) {
        throw 'dependencies[] must contain the complete engine dependency set.'
    }
    $dependencyById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    foreach ($dependency in $dependencies) {
        $dependencyId = [string](Get-RSObjectProperty `
                -Object $dependency `
                -Name id `
                -Required)
        Assert-RSObjectPropertySet `
            -Object $dependency `
            -Context "dependencies[$dependencyId]" `
            -Allowed @('id', 'pin', 'inputIds', 'license', 'updateDiscovery')
        if ($dependencyId -cnotmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$') {
            throw "Invalid dependency id '$dependencyId'."
        }
        if ($dependencyById.ContainsKey($dependencyId)) {
            throw "Duplicate dependency id '$dependencyId'."
        }
        $pin = [string](Get-RSObjectProperty `
                -Object $dependency `
                -Name pin `
                -Required)
        if ([string]::IsNullOrWhiteSpace($pin)) {
            throw "Dependency '$dependencyId' pin may not be empty."
        }
        $dependencyInputIds = @(Get-RSObjectProperty `
                -Object $dependency `
                -Name inputIds `
                -Required)
        if ($dependencyInputIds.Count -eq 0) {
            throw "Dependency '$dependencyId' inputIds may not be empty."
        }
        $dependencyInputSet = [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
        foreach ($inputIdValue in $dependencyInputIds) {
            $inputId = [string]$inputIdValue
            if ([string]::IsNullOrWhiteSpace($inputId) -or
                -not $dependencyInputSet.Add($inputId)) {
                throw "Dependency '$dependencyId' has an empty or duplicate inputId '$inputId'."
            }
            if (-not $inputById.ContainsKey($inputId)) {
                throw "Dependency '$dependencyId' refers to unknown input '$inputId'."
            }
        }

        $license = Get-RSObjectProperty `
            -Object $dependency `
            -Name license `
            -Required
        Assert-RSObjectPropertySet `
            -Object $license `
            -Context "dependencies[$dependencyId].license" `
            -Allowed @('spdx', 'licenseInputIds', 'noticeInputIds')
        $spdx = [string](Get-RSObjectProperty `
                -Object $license `
                -Name spdx `
                -Required)
        if ([string]::IsNullOrWhiteSpace($spdx)) {
            throw "Dependency '$dependencyId' has no SPDX identity."
        }
        foreach ($field in @('licenseInputIds', 'noticeInputIds')) {
            if ($null -eq $license.PSObject.Properties[$field]) {
                throw "Dependency '$dependencyId' license.$field must be explicitly declared."
            }
            $references = @(Get-RSObjectProperty `
                    -Object $license `
                    -Name $field `
                    -Required)
            if ($field -ceq 'licenseInputIds' -and $references.Count -eq 0) {
                throw "Dependency '$dependencyId' licenseInputIds may not be empty."
            }
            $referenceSet = [Collections.Generic.HashSet[string]]::new(
                [StringComparer]::Ordinal)
            foreach ($inputIdValue in $references) {
                $inputId = [string]$inputIdValue
                if ([string]::IsNullOrWhiteSpace($inputId) -or
                    -not $referenceSet.Add($inputId)) {
                    throw "Dependency '$dependencyId' $field has an empty or duplicate inputId '$inputId'."
                }
                if (-not $dependencyInputSet.Contains($inputId) -or
                    -not $inputById.ContainsKey($inputId)) {
                    throw "Dependency '$dependencyId' $field must refer exactly to one of its locked inputIds ('$inputId')."
                }
            }
        }
        $dependencyById.Add($dependencyId, $dependency)
        foreach ($discovery in @(Get-RSObjectProperty `
                -Object $dependency `
                -Name updateDiscovery `
                -Required)) {
            Assert-RSObjectPropertySet `
                -Object $discovery `
                -Context "dependencies[$dependencyId].updateDiscovery" `
                -Allowed @(
                    'id',
                    'kind',
                    'uri',
                    'currentValue',
                    'jsonProperty',
                    'securityAdvisory')
        }
    }
    $requiredLaneKeys = @(
        "ARM64`0Debug",
        "ARM64`0Release",
        "x64`0ASan Debug",
        "x64`0Debug",
        "x64`0Release")
    foreach ($requiredLaneKey in $requiredLaneKeys) {
        if (-not $laneByKey.ContainsKey($requiredLaneKey)) {
            $display = $requiredLaneKey.Replace("`0", '/')
            throw "Terminal engine lock is missing required lane '$display'."
        }
    }

    $runtimeClosure = @(Get-RSObjectProperty `
            -Object $lock `
            -Name runtimeClosure `
            -Required)
    $runtimeKeys = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    $runtimeOutputIds = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::Ordinal)
    $runtimeEntriesByLane = @{}
    foreach ($entry in $runtimeClosure) {
        $platform = [string](Get-RSObjectProperty `
                -Object $entry `
                -Name platform `
                -Required)
        $configuration = [string](Get-RSObjectProperty `
                -Object $entry `
                -Name configuration `
                -Required)
        Assert-RSObjectPropertySet `
            -Object $entry `
            -Context "runtimeClosure[$platform/$configuration]" `
            -Allowed @(
                'platform',
                'configuration',
                'module',
                'system',
                'loadSource',
                'loadedBy',
                'outputId',
                'inputId')
        if (-not $laneByKey.ContainsKey("$platform`0$configuration")) {
            throw "Runtime closure entry refers to unknown lane '$platform/$configuration'."
        }
        $module = [string](Get-RSObjectProperty `
                -Object $entry `
                -Name module `
                -Required)
        if ($module -cnotmatch '^[A-Za-z0-9_.+-]+\.dll$') {
            throw "Runtime closure module '$module' is not one canonical DLL basename."
        }
        $key = "$platform`0$configuration`0$module"
        if (-not $runtimeKeys.Add($key)) {
            throw "Duplicate runtime closure entry '$platform/$configuration/$module'."
        }
        $loadSource = [string](Get-RSObjectProperty `
                -Object $entry `
                -Name loadSource `
                -Required)
        if ($loadSource -cnotin @(
                'normal-import',
                'delay-import',
                'explicit-load')) {
            throw "Runtime closure '$module' has unsupported loadSource '$loadSource'."
        }
        $loadedBy = [string](Get-RSObjectProperty `
                -Object $entry `
                -Name loadedBy `
                -Required)
        if ($loadedBy -cne '$product' -and
            $loadedBy -cnotmatch '^[A-Za-z0-9_.+-]+\.dll$') {
            throw "Runtime closure '$module' loadedBy must be '`$product' or one DLL basename."
        }
        $laneKey = "$platform`0$configuration"
        if (-not $runtimeEntriesByLane.ContainsKey($laneKey)) {
            $runtimeEntriesByLane[$laneKey] =
                [Collections.Generic.Dictionary[string, object]]::new(
                    [StringComparer]::OrdinalIgnoreCase)
        }
        $runtimeEntriesByLane[$laneKey].Add($module, $entry)
        $system = Get-RSObjectProperty -Object $entry -Name system -Required
        if ($system -isnot [bool]) {
            throw "Runtime closure '$module' system must be Boolean."
        }
        $outputId = Get-RSObjectProperty -Object $entry -Name outputId -Default $null
        $inputId = Get-RSObjectProperty -Object $entry -Name inputId -Default $null
        if ([bool]$system) {
            if ($null -ne $outputId -or $null -ne $inputId) {
                throw "System runtime '$module' cannot bind an app-local output or input."
            }
        }
        else {
            if ($linkage -ceq 'source-static') {
                throw "A source-static engine cannot declare app-local runtime leaf '$module'."
            }
            if (($null -eq $outputId) -eq ($null -eq $inputId)) {
                throw "Non-system runtime '$module' must bind exactly one outputId or inputId."
            }
            if ($null -ne $outputId) {
                if (-not $outputById.ContainsKey([string]$outputId)) {
                    throw "Runtime closure '$module' refers to unknown output '$outputId'."
                }
                $output = $outputById[[string]$outputId]
                if ([string]$output.platform -cne $platform -or
                    [string]$output.configuration -cne $configuration) {
                    throw "Runtime closure '$module' output belongs to another lane."
                }
                if (-not $runtimeOutputIds.Add([string]$outputId)) {
                    throw "Output '$outputId' is bound by more than one runtime closure entry."
                }
            }
            if ($null -ne $inputId -and
                -not $inputById.ContainsKey([string]$inputId)) {
                throw "Runtime closure '$module' refers to unknown input '$inputId'."
            }
        }
    }
    foreach ($laneKey in $runtimeEntriesByLane.Keys) {
        $entries = $runtimeEntriesByLane[$laneKey]
        foreach ($entry in $entries.Values) {
            $visited = [Collections.Generic.HashSet[string]]::new(
                [StringComparer]::OrdinalIgnoreCase)
            $current = $entry
            while ($true) {
                $module = [string]$current.module
                if (-not $visited.Add($module)) {
                    $displayLane = $laneKey.Replace("`0", '/')
                    throw "Runtime closure lane '$displayLane' contains a load cycle at '$module'."
                }
                $loadedBy = [string]$current.loadedBy
                if ($loadedBy -ceq '$product') {
                    break
                }
                if (-not $entries.ContainsKey($loadedBy)) {
                    $displayLane = $laneKey.Replace("`0", '/')
                    throw "Runtime closure '$module' in '$displayLane' is loaded by missing parent '$loadedBy'."
                }
                $current = $entries[$loadedBy]
            }
        }
    }
    if ($linkage -ceq 'private-dynamic') {
        foreach ($output in $outputs) {
            if ([string]$output.kind -ceq 'pe' -and
                -not $runtimeOutputIds.Contains([string]$output.id)) {
                throw "Private-dynamic PE output '$($output.id)' is absent from runtimeClosure."
            }
        }
    }

    $abi = $lock.abi
    Assert-RSObjectPropertySet `
        -Object $abi `
        -Context 'abi' `
        -Allowed @(
            'headerSetSha256',
            'enumIdentity',
            'typeLayoutIdentity',
            'callingConvention',
            'capabilities')
    foreach ($name in @(
            'headerSetSha256',
            'enumIdentity',
            'typeLayoutIdentity',
            'callingConvention',
            'capabilities')) {
        if ($null -eq $abi.PSObject.Properties[$name]) {
            throw "abi.$name is required."
        }
    }
    Assert-RSLowerSha256 -Value ([string]$abi.headerSetSha256) -Name 'abi.headerSetSha256'

    $qualification = Get-RSObjectProperty -Object $lock -Name qualification -Required
    Assert-RSObjectPropertySet `
        -Object $qualification `
        -Context 'qualification' `
        -Allowed @(
            'security',
            'conformance',
            'fuzz',
            'performance',
            'memory',
            'packageSize')
    foreach ($name in @(
            'security',
            'conformance',
            'fuzz',
            'performance',
            'memory',
            'packageSize')) {
        $record = Get-RSObjectProperty -Object $qualification -Name $name -Required
        Assert-RSObjectPropertySet `
            -Object $record `
            -Context "qualification.$name" `
            -Allowed @('status', 'evidenceSha256')
    }

    $upgrade = Get-RSObjectProperty -Object $lock -Name upgrade -Required
    Assert-RSObjectPropertySet `
        -Object $upgrade `
        -Context 'upgrade' `
        -Allowed @(
            'basedOnLockSha256',
            'targetSourceSha256',
            'persistedStateSchemaChanged',
            'requiredChanges')

    $lockSchema = Join-Path $repoRoot 'Specs\Terminal\TerminalEngineLock.schema.json'
    if (-not (Test-Json `
            -Json $lockJson `
            -SchemaFile $lockSchema `
            -ErrorAction Stop)) {
        throw 'Terminal engine lock does not satisfy its closed schema.'
    }

    $lock | Add-Member -NotePropertyName repoRoot -NotePropertyValue $repoRoot -Force
    $lock | Add-Member -NotePropertyName lockPath -NotePropertyValue $candidate -Force
    $lock | Add-Member -NotePropertyName lockSha256 -NotePropertyValue (
        Get-RSSha256 -Path $candidate) -Force
    $lock | Add-Member -NotePropertyName inputById -NotePropertyValue $inputById -Force
    $lock | Add-Member -NotePropertyName dependencyById -NotePropertyValue $dependencyById -Force
    $lock | Add-Member -NotePropertyName laneByKey -NotePropertyValue $laneByKey -Force
    $lock | Add-Member -NotePropertyName outputById -NotePropertyValue $outputById -Force

    if ($ValidateInputs) {
        $verified = @($inputs | ForEach-Object {
                Test-RSTerminalInputIdentity -Lock $lock -InputRecord $_
            })
        $lock | Add-Member -NotePropertyName verifiedInputs -NotePropertyValue $verified -Force
    }
    return $lock
}

function Get-RSTerminalLicenseClosure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Lock
    )

    if ($null -eq $Lock.PSObject.Properties['inputById']) {
        throw 'License closure requires a lock returned by Read-RSTerminalEngineLock.'
    }
    return @(
        foreach ($dependency in @($Lock.dependencies |
                Sort-Object -Stable -CaseSensitive -Property id)) {
            $license = Get-RSObjectProperty `
                -Object $dependency `
                -Name license `
                -Required
            [ordered]@{
                dependencyId = [string]$dependency.id
                pin = [string]$dependency.pin
                spdx = [string]$license.spdx
                inputs = @(
                    foreach ($inputId in @($license.licenseInputIds |
                            Sort-Object -CaseSensitive)) {
                        $record = $Lock.inputById[[string]$inputId]
                        [ordered]@{
                            id = [string]$inputId
                            sha256 = [string]$record.sha256
                        }
                    }
                )
            }
        }
    )
}

function Get-RSTerminalNoticeClosure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Lock
    )

    if ($null -eq $Lock.PSObject.Properties['inputById']) {
        throw 'Notice closure requires a lock returned by Read-RSTerminalEngineLock.'
    }
    return @(
        foreach ($dependency in @($Lock.dependencies |
                Sort-Object -Stable -CaseSensitive -Property id)) {
            $license = Get-RSObjectProperty `
                -Object $dependency `
                -Name license `
                -Required
            [ordered]@{
                dependencyId = [string]$dependency.id
                pin = [string]$dependency.pin
                inputs = @(
                    foreach ($inputId in @($license.noticeInputIds |
                            Sort-Object -CaseSensitive)) {
                        $record = $Lock.inputById[[string]$inputId]
                        [ordered]@{
                            id = [string]$inputId
                            sha256 = [string]$record.sha256
                        }
                    }
                )
            }
        }
    )
}

function Get-RSTerminalLane {
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [Parameter(Mandatory)]
        [string]$Platform,

        [Parameter(Mandatory)]
        [string]$Configuration
    )

    $key = "$Platform`0$Configuration"
    if (-not $Lock.laneByKey.ContainsKey($key)) {
        throw "Lock does not declare lane '$Platform/$Configuration'."
    }
    return $Lock.laneByKey[$key]
}

function Get-RSTerminalLaneOutputs {
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [Parameter(Mandatory)]
        [object]$Lane
    )

    return @(@($Lane.outputIds) | ForEach-Object {
            $Lock.outputById[[string]$_]
        })
}

function Assert-RSTerminalSafeBuildRoot {
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot,

        [Parameter(Mandatory)]
        [string]$BuildRoot
    )

    $canonical = if ([IO.Path]::IsPathFullyQualified($BuildRoot)) {
        [IO.Path]::GetFullPath($BuildRoot)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $RepoRoot $BuildRoot))
    }
    $dotBuild = Join-Path $RepoRoot '.build'
    if (-not (Test-RSPathIsWithin -Path $canonical -Root $dotBuild)) {
        throw "Build root must be a child of the repository .build directory: $canonical"
    }
    Assert-RSNoReparsePathComponents -Root $dotBuild -Path $canonical
    return $canonical
}

function Remove-RSTerminalSafeDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $canonical = Assert-RSTerminalSafeBuildRoot -RepoRoot $RepoRoot -BuildRoot $Path
    if (Test-Path -LiteralPath $canonical) {
        Remove-Item -LiteralPath $canonical -Recurse -Force -ErrorAction Stop
    }
    if (Test-Path -LiteralPath $canonical) {
        throw "Clean failed to remove exact build directory: $canonical"
    }
}

function Expand-RSSafeZip {
    param(
        [Parameter(Mandatory)]
        [string]$Archive,

        [Parameter(Mandatory)]
        [string]$Destination
    )

    Add-Type -AssemblyName System.IO.Compression
    $destinationRoot = [IO.Path]::GetFullPath($Destination)
    [void](New-Item -ItemType Directory -Path $destinationRoot -Force)
    $seen = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    $zip = [IO.Compression.ZipFile]::OpenRead($Archive)
    try {
        foreach ($entry in $zip.Entries) {
            $name = $entry.FullName.Replace('\', '/')
            if ([string]::IsNullOrEmpty($name)) {
                continue
            }
            $parts = @($name.Split('/', [StringSplitOptions]::RemoveEmptyEntries))
            if ($name.StartsWith('/') -or
                $name.Contains(':') -or
                @($parts | Where-Object { $_ -in @('.', '..') }).Count -gt 0) {
                throw "ZIP entry is unsafe: $name"
            }
            $unixMode = ([uint32]$entry.ExternalAttributes -shr 16) -band 0xf000
            if ($unixMode -eq 0xa000) {
                throw "ZIP entry is a symbolic link: $name"
            }
            $target = [IO.Path]::GetFullPath((Join-Path $destinationRoot ($parts -join '\')))
            if (-not (Test-RSPathIsWithin -Path $target -Root $destinationRoot)) {
                throw "ZIP entry escapes its destination: $name"
            }
            if ($name.EndsWith('/')) {
                [void](New-Item -ItemType Directory -Path $target -Force)
                continue
            }
            if (-not $seen.Add($target)) {
                throw "ZIP contains duplicate Windows path '$name'."
            }
            [void](New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force)
            $inputStream = $entry.Open()
            $outputStream = [IO.FileStream]::new(
                $target,
                [IO.FileMode]::CreateNew,
                [IO.FileAccess]::Write,
                [IO.FileShare]::None)
            try {
                $inputStream.CopyTo($outputStream)
            }
            finally {
                $outputStream.Dispose()
                $inputStream.Dispose()
            }
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Copy-RSSafeDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$Source,

        [Parameter(Mandatory)]
        [string]$Destination
    )

    $identity = Get-RSDirectoryIdentity -Path $Source
    [void](New-Item -ItemType Directory -Path $Destination -Force)
    foreach ($file in $identity.files) {
        $sourceFile = Resolve-RSTerminalPathInRoot `
            -Root $Source `
            -Path ([string]$file.path) `
            -MustExist
        $destinationFile = Resolve-RSTerminalPathInRoot `
            -Root $Destination `
            -Path ([string]$file.path)
        [void](New-Item -ItemType Directory -Path (Split-Path -Parent $destinationFile) -Force)
        [IO.File]::Copy($sourceFile, $destinationFile, $false)
    }
}

function Invoke-RSTerminalMaterialization {
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [Parameter(Mandatory)]
        [object]$Materialization,

        [Parameter(Mandatory)]
        [string]$WorkRoot
    )

    $inputId = [string]$Materialization.inputId
    $inputRecord = $Lock.inputById[$inputId]
    [void](Test-RSTerminalInputIdentity `
            -Lock $Lock `
            -InputRecord $inputRecord)
    $source = Resolve-RSTerminalInputPath -Lock $Lock -InputRecord $inputRecord
    $destination = Resolve-RSTerminalPathInRoot `
        -Root $WorkRoot `
        -Path ([string]$Materialization.destination)
    $mode = [string](Get-RSObjectProperty -Object $Materialization -Name mode -Required)

    switch ($mode) {
        'copy-file' {
            if ([string]$inputRecord.kind -cne 'file') {
                throw "Materialization '$inputId' requires a file input."
            }
            [void](New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force)
            [IO.File]::Copy($source, $destination, $false)
            if ((Get-RSSha256 -Path $destination) -cne [string]$inputRecord.sha256) {
                throw "Materialized file input '$inputId' changed while it was copied."
            }
        }
        'copy-directory' {
            if ([string]$inputRecord.kind -cne 'directory') {
                throw "Materialization '$inputId' requires a directory input."
            }
            Copy-RSSafeDirectory -Source $source -Destination $destination
            if ((Get-RSDirectoryIdentity -Path $destination).sha256 -cne
                [string]$inputRecord.sha256) {
                throw "Materialized directory input '$inputId' changed while it was copied."
            }
        }
        'extract-zip' {
            if ([string]$inputRecord.kind -cne 'zip') {
                throw "Materialization '$inputId' requires a ZIP input."
            }
            $temporary = "$destination.__extract"
            Expand-RSSafeZip -Archive $source -Destination $temporary
            $contentRoot = [string](Get-RSObjectProperty `
                    -Object $Materialization `
                    -Name contentRoot `
                    -Default '')
            $selected = if ([string]::IsNullOrWhiteSpace($contentRoot)) {
                $temporary
            }
            else {
                Resolve-RSTerminalPathInRoot `
                    -Root $temporary `
                    -Path $contentRoot `
                    -MustExist
            }
            Copy-RSSafeDirectory -Source $selected -Destination $destination
            $materializedSha256 = (Get-RSDirectoryIdentity -Path $destination).sha256
            if ($materializedSha256 -cne [string]$Materialization.contentSha256) {
                throw "Extracted ZIP input '$inputId' content differs from the lock."
            }
            Remove-RSTerminalSafeDirectory -RepoRoot $Lock.repoRoot -Path $temporary
        }
        default {
            throw "Unsupported materialization mode '$mode'."
        }
    }
}

function Convert-RSTerminalBuildToken {
    param(
        [Parameter(Mandatory)]
        [string]$Value,

        [Parameter(Mandatory)]
        [hashtable]$Tokens
    )

    $result = $Value
    foreach ($key in $Tokens.Keys) {
        $result = $result.Replace("{$key}", [string]$Tokens[$key])
    }
    if ($result -match '\{(work|output|cache|globalCache|platform|configuration|lockSha256)\}') {
        throw "Build argument contains an unresolved lifecycle token: $result"
    }
    return $result
}

function Invoke-RSCapturedTerminalProcess {
    param(
        [Parameter(Mandatory)]
        [string]$Executable,

        [Parameter(Mandatory)]
        [string[]]$Arguments,

        [Parameter(Mandatory)]
        [string]$WorkingDirectory,

        [hashtable]$Environment = @{},

        [ValidateRange(1, 7200)]
        [int]$TimeoutSeconds = 3600
    )

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $Executable
    $startInfo.WorkingDirectory = $WorkingDirectory
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.Environment.Clear()
    foreach ($argument in $Arguments) {
        [void]$startInfo.ArgumentList.Add($argument)
    }
    foreach ($pair in $Environment.GetEnumerator()) {
        $startInfo.Environment[[string]$pair.Key] = [string]$pair.Value
    }
    $environmentIdentity = [object[]]@(
        $Environment.GetEnumerator() |
            Sort-Object -Property Key -CaseSensitive |
            ForEach-Object {
                [ordered]@{
                    name = [string]$_.Key
                    value = [string]$_.Value
                }
            })
    $environmentSha256 = Get-RSCanonicalObjectSha256 -InputObject $environmentIdentity

    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) {
            throw "Failed to start build executable '$Executable'."
        }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $process.Kill($true)
            $process.WaitForExit()
            throw "Build process timed out after $TimeoutSeconds seconds."
        }
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        $encoding = [Text.UTF8Encoding]::new($false)
        $sha = [Security.Cryptography.SHA256]::Create()
        try {
            $stdoutSha = [Convert]::ToHexString(
                $sha.ComputeHash($encoding.GetBytes($stdout))).ToLowerInvariant()
            $stderrSha = [Convert]::ToHexString(
                $sha.ComputeHash($encoding.GetBytes($stderr))).ToLowerInvariant()
        }
        finally {
            $sha.Dispose()
        }
        if ($process.ExitCode -ne 0) {
            $summary = if ([string]::IsNullOrWhiteSpace($stderr)) {
                'no stderr'
            }
            else {
                $stderr.Trim().Substring(0, [Math]::Min(500, $stderr.Trim().Length))
            }
            throw "Build command failed with exit code $($process.ExitCode): $summary"
        }
        return [pscustomobject]@{
            exitCode = $process.ExitCode
            stdoutSha256 = $stdoutSha
            stderrSha256 = $stderrSha
            environmentSha256 = $environmentSha256
        }
    }
    finally {
        $process.Dispose()
    }
}

function Get-RSTerminalNativeArchitecture {
    return [pscustomobject]@{
        process = [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString()
        os = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
    }
}

function Test-RSTerminalOutputFiles {
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [Parameter(Mandatory)]
        [object]$Lane,

        [Parameter(Mandatory)]
        [string]$LaneOutputRoot
    )

    $peModule = Join-Path $Lock.repoRoot 'Tools\TerminalEngine\TerminalEnginePeIdentity.psm1'
    Import-Module $peModule -ErrorAction Stop
    $records = [Collections.Generic.List[object]]::new()
    foreach ($output in @(Get-RSTerminalLaneOutputs -Lock $Lock -Lane $Lane)) {
        $path = Resolve-RSTerminalPathInRoot `
            -Root $LaneOutputRoot `
            -Path ([string]$output.relativePath) `
            -MustExist
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Expected output is not a file: $path"
        }
        $rawSha = Get-RSSha256 -Path $path
        $rawMatch = [string]::Equals(
            $rawSha,
            [string]$output.sha256,
            [StringComparison]::Ordinal)
        if (-not $rawMatch) {
            throw "Output '$($output.id)' raw SHA-256 differs from the exact lock."
        }

        $semanticSha = $null
        $machine = $null
        $exports = @()
        if ([string]$output.kind -ceq 'pe') {
            $identity = Get-RSPeIdentity -Path $path
            $semanticSha = Get-RSPeSemanticSha256 -Path $path
            $machine = $identity.machine
            if ($semanticSha -cne [string]$output.semanticSha256) {
                throw "Output '$($output.id)' semantic PE identity differs from the lock."
            }
            if ($machine -cne [string](Get-RSObjectProperty `
                    -Object $output `
                    -Name machine `
                    -Required)) {
                throw "Output '$($output.id)' PE machine differs from the lock."
            }
            $exports = @($identity.semantic.exports | ForEach-Object { $_.name } |
                    Where-Object { $null -ne $_ })
            foreach ($requiredExport in @(Get-RSObjectProperty `
                    -Object $output `
                    -Name requiredExports `
                    -Default @())) {
                if ($exports -cnotcontains [string]$requiredExport) {
                    throw "Output '$($output.id)' lacks required export '$requiredExport'."
                }
            }
        }
        $records.Add([pscustomobject]@{
                id = [string]$output.id
                relativePath = ([string]$output.relativePath).Replace('\', '/')
                sizeBytes = (Get-Item -LiteralPath $path).Length.ToString(
                    [Globalization.CultureInfo]::InvariantCulture)
                sha256 = $rawSha
                canonicalReferenceSha256 = [string]$output.sha256
                rawSha256Policy = [string]$output.rawSha256Policy
                rawMatchesCanonicalReference = $rawMatch
                semanticSha256 = $semanticSha
                machine = $machine
                exportSetSha256 = if ($exports.Count -eq 0) { $null } else {
                    Get-RSCanonicalObjectSha256 -InputObject @(
                        $exports | Sort-Object -CaseSensitive)
                }
            })
    }
    return $records.ToArray()
}

function Invoke-RSTerminalBuildLane {
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [Parameter(Mandatory)]
        [string]$Platform,

        [Parameter(Mandatory)]
        [string]$Configuration,

        [Parameter(Mandatory)]
        [string]$OutputRoot,

        [switch]$Clean,

        [hashtable]$Visited = @{}
    )

    $lane = Get-RSTerminalLane -Lock $Lock -Platform $Platform -Configuration $Configuration
    $key = "$Platform/$Configuration"
    if ($Visited.ContainsKey($key)) {
        throw "Artifact reuse cycle detected at '$key'."
    }
    $Visited[$key] = $true
    try {
        $laneOutputRoot = Join-Path (Join-Path $OutputRoot $Platform) $Configuration
        $workRoot = Join-Path (
            Join-Path (
                Join-Path (
                    Join-Path $OutputRoot '.work') $Lock.lockSha256) $Platform) $Configuration
        if ($Clean) {
            Remove-RSTerminalSafeDirectory -RepoRoot $Lock.repoRoot -Path $laneOutputRoot
            Remove-RSTerminalSafeDirectory -RepoRoot $Lock.repoRoot -Path $workRoot
        }
        [void](New-Item -ItemType Directory -Path $laneOutputRoot -Force)

        if ($null -ne $lane.reuseFrom) {
            $originPlatform = [string](Get-RSObjectProperty `
                    -Object $lane.reuseFrom `
                    -Name platform `
                    -Required)
            $originConfiguration = [string](Get-RSObjectProperty `
                    -Object $lane.reuseFrom `
                    -Name configuration `
                    -Required)
            $originResult = Invoke-RSTerminalBuildLane `
                -Lock $Lock `
                -Platform $originPlatform `
                -Configuration $originConfiguration `
                -OutputRoot $OutputRoot `
                -Clean:$Clean `
                -Visited $Visited
            foreach ($output in @(Get-RSTerminalLaneOutputs -Lock $Lock -Lane $lane)) {
                $sourceOutputId = [string](Get-RSObjectProperty `
                        -Object $output `
                        -Name sourceOutputId `
                        -Required)
                $sourceOutput = $Lock.outputById[$sourceOutputId]
                if ($null -eq $sourceOutput -or
                    [string]$sourceOutput.platform -cne $originPlatform -or
                    [string]$sourceOutput.configuration -cne $originConfiguration) {
                    throw "Reused output '$($output.id)' has an invalid sourceOutputId."
                }
                $source = Resolve-RSTerminalPathInRoot `
                    -Root $originResult.outputRoot `
                    -Path ([string]$sourceOutput.relativePath) `
                    -MustExist
                $destination = Resolve-RSTerminalPathInRoot `
                    -Root $laneOutputRoot `
                    -Path ([string]$output.relativePath)
                [void](New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force)
                [IO.File]::Copy($source, $destination, $true)
            }
            $commandResult = [pscustomobject]@{
                reused = $true
                fromPlatform = $originPlatform
                fromConfiguration = $originConfiguration
                stdoutSha256 = $null
                stderrSha256 = $null
            }
        }
        else {
            if (Test-Path -LiteralPath $workRoot) {
                Remove-RSTerminalSafeDirectory -RepoRoot $Lock.repoRoot -Path $workRoot
            }
            [void](New-Item -ItemType Directory -Path $workRoot -Force)
            $materializations = @(
                @(Get-RSObjectProperty `
                    -Object $Lock.build `
                    -Name materializations `
                    -Default @()) +
                @(Get-RSObjectProperty -Object $lane -Name materializations -Default @()))
            foreach ($materialization in $materializations) {
                Invoke-RSTerminalMaterialization `
                    -Lock $Lock `
                    -Materialization $materialization `
                    -WorkRoot $workRoot
            }

            $executableRelative = [string](Get-RSObjectProperty `
                    -Object $lane `
                    -Name executableRelativePath `
                    -Default (Get-RSObjectProperty `
                        -Object $Lock.build `
                        -Name executableRelativePath `
                        -Default ''))
            $executable = Resolve-RSTerminalPathInRoot `
                -Root $workRoot `
                -Path $executableRelative `
                -MustExist
            if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
                throw "Build executable is not a file: $executable"
            }
            $workingRelative = [string](Get-RSObjectProperty `
                    -Object $lane `
                    -Name workingDirectoryRelativePath `
                    -Default (Get-RSObjectProperty `
                        -Object $Lock.build `
                        -Name workingDirectoryRelativePath `
                        -Default '.'))
            $workingDirectory = if ($workingRelative -eq '.') {
                $workRoot
            }
            else {
                Resolve-RSTerminalPathInRoot `
                    -Root $workRoot `
                    -Path $workingRelative `
                    -MustExist
            }
            $tokens = @{
                work = $workRoot
                output = $laneOutputRoot
                cache = (Join-Path $workRoot 'cache')
                globalCache = (Join-Path $workRoot 'global-cache')
                platform = $Platform
                configuration = $Configuration
                lockSha256 = $Lock.lockSha256
            }
            $arguments = @($lane.arguments | ForEach-Object {
                    Convert-RSTerminalBuildToken -Value ([string]$_) -Tokens $tokens
                })
            $environmentRoot = Join-Path $workRoot 'sealed-environment'
            $temporaryRoot = Join-Path $environmentRoot 'temp'
            $profileRoot = Join-Path $environmentRoot 'profile'
            $appDataRoot = Join-Path $profileRoot 'AppData\Roaming'
            $localAppDataRoot = Join-Path $profileRoot 'AppData\Local'
            foreach ($directory in @(
                    $temporaryRoot,
                    $profileRoot,
                    $appDataRoot,
                    $localAppDataRoot,
                    $tokens.cache,
                    $tokens.globalCache)) {
                [void](New-Item -ItemType Directory -Path $directory -Force)
            }
            $windowsRoot = [Environment]::GetFolderPath(
                [Environment+SpecialFolder]::Windows)
            if ([string]::IsNullOrWhiteSpace($windowsRoot) -or
                -not [IO.Path]::IsPathFullyQualified($windowsRoot) -or
                -not [IO.Directory]::Exists($windowsRoot)) {
                throw 'The Windows directory required for the sealed build environment is unavailable.'
            }
            $environment = @{
                APPDATA = $appDataRoot
                GIT_CONFIG_GLOBAL = 'NUL'
                GIT_CONFIG_NOSYSTEM = '1'
                HOME = $profileRoot
                LOCALAPPDATA = $localAppDataRoot
                PATH = (Split-Path -Parent $executable)
                TEMP = $temporaryRoot
                TMP = $temporaryRoot
                USERPROFILE = $profileRoot
                WINDIR = $windowsRoot
                SystemRoot = $windowsRoot
            }
            $reservedEnvironmentNames = [string[]]@($environment.Keys)
            $environmentObject = Get-RSObjectProperty `
                -Object $lane `
                -Name environment `
                -Default (Get-RSObjectProperty `
                    -Object $Lock.build `
                    -Name environment `
                    -Default ([pscustomobject]@{}))
            foreach ($property in $environmentObject.PSObject.Properties) {
                if ($property.Name -cnotmatch '^[A-Z][A-Z0-9_]*$') {
                    throw "Build environment name '$($property.Name)' is not allowed."
                }
                if ($reservedEnvironmentNames -ccontains $property.Name) {
                    throw "Build environment name '$($property.Name)' is lifecycle-reserved."
                }
                $environment[$property.Name] = Convert-RSTerminalBuildToken `
                    -Value ([string]$property.Value) `
                    -Tokens $tokens
            }
            $commandResult = Invoke-RSCapturedTerminalProcess `
                -Executable $executable `
                -Arguments $arguments `
                -WorkingDirectory $workingDirectory `
                -Environment $environment `
                -TimeoutSeconds ([int](Get-RSObjectProperty `
                    -Object $lane `
                    -Name timeoutSeconds `
                    -Default 3600))

            foreach ($output in @(Get-RSTerminalLaneOutputs -Lock $Lock -Lane $lane)) {
                $builtRelative = [string](Get-RSObjectProperty `
                        -Object $output `
                        -Name builtRelativePath `
                        -Required)
                $source = Resolve-RSTerminalPathInRoot `
                    -Root $workRoot `
                    -Path $builtRelative `
                    -MustExist
                $destination = Resolve-RSTerminalPathInRoot `
                    -Root $laneOutputRoot `
                    -Path ([string]$output.relativePath)
                [void](New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force)
                [IO.File]::Copy($source, $destination, $true)
            }
        }

        $verifiedOutputs = Test-RSTerminalOutputFiles `
            -Lock $Lock `
            -Lane $lane `
            -LaneOutputRoot $laneOutputRoot
        return [pscustomobject]@{
            platform = $Platform
            configuration = $Configuration
            outputRoot = $laneOutputRoot
            command = $commandResult
            outputs = $verifiedOutputs
        }
    }
    finally {
        [void]$Visited.Remove($key)
    }
}

function Write-RSTerminalJsonFile {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [object]$Value
    )

    $parent = Split-Path -Parent $Path
    [void](New-Item -ItemType Directory -Path $parent -Force)
    $temporary = Join-Path $parent (
        ".$([IO.Path]::GetFileName($Path)).$([Guid]::NewGuid().ToString('N')).tmp")
    $json = $Value | ConvertTo-Json -Depth 100
    [IO.File]::WriteAllText($temporary, $json + "`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::Move($temporary, $Path, $true)
}

function Initialize-RSTerminalRuntimeIdentityInterop {
    if ('RedSalamander.TerminalEngine.NativeIdentity' -as [type]) {
        return
    }

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace RedSalamander.TerminalEngine {
    public static class NativeIdentity {
        [StructLayout(LayoutKind.Sequential)]
        private struct NativeString {
            public IntPtr Pointer;
            public UIntPtr Length;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Identity {
            public UIntPtr Size;
            public NativeString UpstreamCommit;
            public UInt32 PatchAbi;
            public NativeString PublicHeaderSha256;
            public NativeString BuildIdentity;
            public Int32 Architecture;
            public Int32 CallingConvention;
            public UInt64 Capabilities;
            public UInt64 EnumIdentity;
            public UInt64 TypeLayoutIdentity;
            public UInt32 PointerSize;
            public UInt32 TerminalResourceLimitsSize;
        }

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate Int32 IdentityFunction(ref Identity identity);

        private static string ReadString(NativeString value) {
            ulong length = value.Length.ToUInt64();
            if (value.Pointer == IntPtr.Zero || length > 4096)
                throw new InvalidOperationException("Runtime identity string is invalid.");
            byte[] bytes = new byte[(int)length];
            Marshal.Copy(value.Pointer, bytes, 0, bytes.Length);
            return System.Text.Encoding.UTF8.GetString(bytes);
        }

        public static string ReadJson(string path) {
            IntPtr module = NativeLibrary.Load(path);
            try {
                IntPtr address = NativeLibrary.GetExport(
                    module,
                    "rs_terminal_runtime_identity");
                Identity value = new Identity();
                value.Size = (UIntPtr)(uint)Marshal.SizeOf<Identity>();
                IdentityFunction function =
                    Marshal.GetDelegateForFunctionPointer<IdentityFunction>(address);
                int result = function(ref value);
                if (result != 0)
                    throw new InvalidOperationException(
                        "rs_terminal_runtime_identity returned " +
                        result.ToString());
                return System.Text.Json.JsonSerializer.Serialize(new {
                    size = value.Size.ToUInt64().ToString(),
                    upstreamCommit = ReadString(value.UpstreamCommit),
                    patchAbi = value.PatchAbi.ToString(),
                    publicHeaderSha256 = ReadString(value.PublicHeaderSha256),
                    buildIdentity = ReadString(value.BuildIdentity),
                    architecture = value.Architecture.ToString(),
                    callingConvention = value.CallingConvention.ToString(),
                    capabilitiesHex = value.Capabilities.ToString("x16"),
                    enumIdentityHex = value.EnumIdentity.ToString("x16"),
                    typeLayoutIdentityHex = value.TypeLayoutIdentity.ToString("x16"),
                    pointerSize = value.PointerSize.ToString(),
                    terminalResourceLimitsSize =
                        value.TerminalResourceLimitsSize.ToString()
                });
            }
            finally {
                NativeLibrary.Free(module);
            }
        }
    }
}
'@
}

function Invoke-RSTerminalRuntimeIdentity {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [object]$Expected,

        [Parameter(Mandatory)]
        [object]$Abi
    )

    Initialize-RSTerminalRuntimeIdentityInterop
    $actual = [RedSalamander.TerminalEngine.NativeIdentity]::ReadJson($Path) |
        ConvertFrom-Json -Depth 20
    $checks = [ordered]@{
        upstreamCommit = [string](Get-RSObjectProperty `
                -Object $Expected `
                -Name upstreamCommit `
                -Required)
        patchAbi = [string](Get-RSObjectProperty -Object $Expected -Name patchAbi -Required)
        publicHeaderSha256 = [string]$Abi.headerSetSha256
        architecture = [string](Get-RSObjectProperty `
                -Object $Expected `
                -Name architecture `
                -Required)
        callingConvention = [string](Get-RSObjectProperty `
                -Object $Expected `
                -Name callingConvention `
                -Required)
        capabilitiesHex = [string](Get-RSObjectProperty `
                -Object $Expected `
                -Name capabilitiesHex `
                -Required)
        enumIdentityHex = [string]$Abi.enumIdentity
        typeLayoutIdentityHex = [string]$Abi.typeLayoutIdentity
        pointerSize = [string](Get-RSObjectProperty `
                -Object $Expected `
                -Name pointerSize `
                -Required)
        terminalResourceLimitsSize = [string](Get-RSObjectProperty `
                -Object $Expected `
                -Name terminalResourceLimitsSize `
                -Required)
    }
    $expectedBuildIdentity = Get-RSObjectProperty `
        -Object $Expected `
        -Name buildIdentity `
        -Default $null
    if ($null -ne $expectedBuildIdentity) {
        $checks.buildIdentity = [string]$expectedBuildIdentity
    }
    foreach ($pair in $checks.GetEnumerator()) {
        if ([string]$actual.($pair.Key) -cne [string]$pair.Value) {
            throw "Runtime identity field '$($pair.Key)' differs from the lock."
        }
    }
    return $actual
}

function Invoke-RSTerminalSmokeHook {
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [Parameter(Mandatory)]
        [object]$Smoke,

        [Parameter(Mandatory)]
        [string]$ArtifactPath
    )

    $inputId = [string](Get-RSObjectProperty `
            -Object $Smoke `
            -Name executableInputId `
            -Required)
    if (-not $Lock.inputById.ContainsKey($inputId)) {
        throw "Smoke hook refers to unknown executable input '$inputId'."
    }
    $executable = Resolve-RSTerminalInputPath `
        -Lock $Lock `
        -InputRecord $Lock.inputById[$inputId]
    $arguments = @(@(Get-RSObjectProperty -Object $Smoke -Name arguments -Default @()) |
            ForEach-Object { ([string]$_).Replace('{artifact}', $ArtifactPath) })
    return Invoke-RSCapturedTerminalProcess `
        -Executable $executable `
        -Arguments $arguments `
        -WorkingDirectory $Lock.repoRoot `
        -TimeoutSeconds ([int](Get-RSObjectProperty `
            -Object $Smoke `
            -Name timeoutSeconds `
            -Default 120))
}

Export-ModuleMember -Function `
    Get-RSObjectProperty, `
    Get-RSTerminalRepoRoot, `
    Test-RSPathIsWithin, `
    Resolve-RSTerminalPathInRoot, `
    Get-RSSha256, `
    Get-RSCanonicalObjectSha256, `
    Get-RSDirectoryIdentity, `
    Resolve-RSTerminalGate0HostHeaderRoot, `
    Assert-RSTerminalHttpsInputUri, `
    Get-RSTerminalInputCacheRoot, `
    Get-RSTerminalInputCachePath, `
    Read-RSTerminalEngineLock, `
    Restore-RSTerminalEngineInputs, `
    Get-RSTerminalLicenseClosure, `
    Get-RSTerminalNoticeClosure, `
    Resolve-RSTerminalInputPath, `
    Test-RSTerminalInputIdentity, `
    Get-RSTerminalLane, `
    Get-RSTerminalLaneOutputs, `
    Assert-RSTerminalSafeBuildRoot, `
    Remove-RSTerminalSafeDirectory, `
    Invoke-RSTerminalBuildLane, `
    Test-RSTerminalOutputFiles, `
    Write-RSTerminalJsonFile, `
    Get-RSTerminalNativeArchitecture, `
    Invoke-RSTerminalRuntimeIdentity, `
    Invoke-RSTerminalSmokeHook
