Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) (
        'TerminalEvidence.psm1'))

$script:RSCollectionDescriptorLeaf =
    'red-salamander/collection-descriptor.v1.json'
$script:RSCollectionDescriptorMaxBytes = [uint64]65536
$script:RSCollectionDescriptorMaxLegalFileBytes = [uint64]16777216
$script:RSCollectionDescriptorMaxLegalTotalBytes = [uint64]67108864

function Get-RSTerminalEngineCollectionDescriptorLeaf {
    return $script:RSCollectionDescriptorLeaf
}

function Get-RSCollectionDescriptorSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath
    )

    return (Get-FileHash -LiteralPath $LiteralPath -Algorithm SHA256).
        Hash.ToLowerInvariant()
}

function Assert-RSCollectionDescriptorSafeLeaf {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $baseName = $Value.Split('.')[0]
    if ([string]::IsNullOrEmpty($Value) -or
        $Value.Length -gt 128 -or
        $Value.EndsWith('.', [StringComparison]::Ordinal) -or
        $Value.EndsWith(' ', [StringComparison]::Ordinal) -or
        $Value -match '[<>:"/\\|?*\x00-\x1f]' -or
        [IO.Path]::GetFileName($Value) -cne $Value -or
        $baseName -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') {
        throw [IO.InvalidDataException]::new(
            "$Label is not one safe filesystem leaf: '$Value'.")
    }
}

function Assert-RSCollectionDescriptorRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [ValidateRange(1, 4096)]
        [int]$MaximumLength = 512,

        [switch]$AllowSpace
    )

    if ([string]::IsNullOrEmpty($Value) -or
        $Value.Length -gt $MaximumLength -or
        $Value.Contains('\', [StringComparison]::Ordinal) -or
        $Value.StartsWith('/', [StringComparison]::Ordinal) -or
        $Value.EndsWith('/', [StringComparison]::Ordinal) -or
        $Value.Contains('//', [StringComparison]::Ordinal) -or
        $Value -match '[:\x00-\x1f]') {
        throw [IO.InvalidDataException]::new(
            "$Label is not a bounded slash-separated relative path.")
    }
    foreach ($segment in $Value.Split('/')) {
        if ($segment -ceq '.' -or
            $segment -ceq '..' -or
            $segment.Length -gt 255 -or
            $segment.EndsWith('.', [StringComparison]::Ordinal) -or
            $segment.EndsWith(' ', [StringComparison]::Ordinal) -or
            (-not $AllowSpace -and
                $segment.Contains(' ', [StringComparison]::Ordinal))) {
            throw [IO.InvalidDataException]::new(
                "$Label contains an unsafe path segment.")
        }
        Assert-RSCollectionDescriptorSafeLeaf `
            -Value $segment `
            -Label $Label
    }
}

function Assert-RSCollectionDescriptorOrdinalRecords {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [object[]]$Record,

        [Parameter(Mandatory = $true)]
        [string]$Property,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if ($null -eq $Record) {
        $Record = [object[]]@()
    }
    $values = [string[]]@(
        $Record | ForEach-Object { [string]$_[$Property] })
    $ordered = [string[]]@($values)
    [Array]::Sort($ordered, [StringComparer]::Ordinal)
    if (-not [Linq.Enumerable]::SequenceEqual($values, $ordered)) {
        throw [IO.InvalidDataException]::new(
            "$Label must be sorted ordinally by $Property.")
    }
    $caseFolded =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($value in $values) {
        if (-not $caseFolded.Add($value)) {
            throw [IO.InvalidDataException]::new(
                "$Label repeats $Property '$value' under Windows case folding.")
        }
    }
}

function Assert-RSCollectionDescriptorOrdinalStrings {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [string[]]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if ($null -eq $Value) {
        $Value = [string[]]@()
    }
    $ordered = [string[]]@($Value)
    [Array]::Sort($ordered, [StringComparer]::Ordinal)
    if (-not [Linq.Enumerable]::SequenceEqual($Value, $ordered)) {
        throw [IO.InvalidDataException]::new("$Label must be sorted ordinally.")
    }
    $caseFolded =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $Value) {
        if (-not $caseFolded.Add($entry)) {
            throw [IO.InvalidDataException]::new(
                "$Label repeats '$entry' under Windows case folding.")
        }
    }
}

function Assert-RSCollectionDescriptorHttpsUri {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $uri = $null
    if (-not [Uri]::TryCreate($Value, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -cne 'https' -or
        -not [string]::IsNullOrEmpty($uri.UserInfo) -or
        -not [string]::IsNullOrEmpty($uri.Fragment)) {
        throw [IO.InvalidDataException]::new(
            "$Label must be an absolute HTTPS URI without credentials or a fragment.")
    }
}

function Assert-RSCollectionDescriptorContainedNoReparsePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $canonicalRoot = [IO.Path]::GetFullPath($Root).
        TrimEnd([IO.Path]::DirectorySeparatorChar)
    $current = [IO.Path]::GetFullPath($LiteralPath)
    if (-not $current.StartsWith(
            $canonicalRoot + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw [IO.InvalidDataException]::new(
            "$Label escaped its allowed root.")
    }
    while (-not $current.Equals(
            $canonicalRoot,
            [StringComparison]::OrdinalIgnoreCase)) {
        if ([IO.File]::Exists($current) -or
            [IO.Directory]::Exists($current)) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw [IO.InvalidDataException]::new(
                    "$Label contains a reparse point.")
            }
        }
        $parent = [IO.Path]::GetDirectoryName($current)
        if ([string]::IsNullOrEmpty($parent) -or
            $parent -ceq $current) {
            throw [IO.InvalidDataException]::new(
                "$Label could not be proven inside its allowed root.")
        }
        $current = $parent
    }
}

function Assert-RSCollectionDescriptorSemantics {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId
    )

    if ([string]$Descriptor['formatId'] -cne
            'red-salamander-terminal-engine-collection-descriptor' -or
        [int]$Descriptor['version'] -ne 1 -or
        [string]$Descriptor['candidateId'] -cne $CandidateId) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' collection descriptor identity is invalid.")
    }

    $legal = $Descriptor['legal']
    $outputPath = [string]$legal['outputRelativePath']
    Assert-RSCollectionDescriptorRelativePath `
        -Value $outputPath `
        -Label "Candidate '$CandidateId' generated legal output"
    if ([StringComparer]::OrdinalIgnoreCase.Equals(
            $script:RSCollectionDescriptorLeaf,
            $outputPath)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' descriptor and generated output paths collide.")
    }

    $headerPackages = [object[]]@($Descriptor['headerPackages'])
    Assert-RSCollectionDescriptorOrdinalRecords `
        -Record $headerPackages `
        -Property inputId `
        -Label "Candidate '$CandidateId' header packages"
    foreach ($record in $headerPackages) {
        Assert-RSCollectionDescriptorRelativePath `
            -Value ([string]$record['includeRelativePath']) `
            -Label "Candidate '$CandidateId' header-package include path"
    }

    $dependencies = [object[]]@($legal['dependencies'])
    Assert-RSCollectionDescriptorOrdinalRecords `
        -Record $dependencies `
        -Property dependencyId `
        -Label "Candidate '$CandidateId' legal dependencies"
    $destinationLeaves =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    $candidateSourcePaths =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    $legalFrozenInputIds =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    [uint64]$legalBytes = 0
    foreach ($dependency in $dependencies) {
        $inputs = [object[]]@($dependency['inputs'])
        Assert-RSCollectionDescriptorOrdinalRecords `
            -Record $inputs `
            -Property inputId `
            -Label "Candidate '$CandidateId' legal dependency inputs"
        [int]$licenseCount = 0
        foreach ($record in $inputs) {
            if ([string]$record['kind'] -ceq 'license') {
                ++$licenseCount
            }
            $leaf = [string]$record['destinationLeaf']
            Assert-RSCollectionDescriptorSafeLeaf `
                -Value $leaf `
                -Label "Candidate '$CandidateId' legal destination"
            if (-not $destinationLeaves.Add($leaf)) {
                throw [IO.InvalidDataException]::new(
                    "Candidate '$CandidateId' repeats a legal destination under Windows case folding.")
            }
            $size = [uint64]::Parse(
                [string]$record['sizeBytes'],
                [Globalization.CultureInfo]::InvariantCulture)
            if ($size -gt $script:RSCollectionDescriptorMaxLegalFileBytes -or
                $legalBytes -gt
                    ($script:RSCollectionDescriptorMaxLegalTotalBytes - $size)) {
                throw [IO.InvalidDataException]::new(
                    "Candidate '$CandidateId' legal inputs exceed a fixed byte bound.")
            }
            $legalBytes += $size
            $source = $record['source']
            switch ([string]$source['sourceKind']) {
                'candidate-source' {
                    $path = [string]$source['sourceRelativePath']
                    Assert-RSCollectionDescriptorRelativePath `
                        -Value $path `
                        -Label "Candidate '$CandidateId' legal candidate-source path"
                    if (-not $candidateSourcePaths.Add($path)) {
                        throw [IO.InvalidDataException]::new(
                            "Candidate '$CandidateId' repeats a legal candidate-source path under Windows case folding.")
                    }
                }
                'frozen-input-member' {
                    [void]$legalFrozenInputIds.Add(
                        [string]$source['sourceInputId'])
                    Assert-RSCollectionDescriptorRelativePath `
                        -Value ([string]$source['archivePath']) `
                        -Label "Candidate '$CandidateId' legal archive member" `
                        -MaximumLength 1024 `
                        -AllowSpace
                }
                'frozen-input-whole' {
                    [void]$legalFrozenInputIds.Add(
                        [string]$source['sourceInputId'])
                }
                default {
                    throw [IO.InvalidDataException]::new(
                        "Candidate '$CandidateId' legal source kind is invalid.")
                }
            }
        }
        if ($licenseCount -eq 0) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' legal dependency has no license input.")
        }
    }

    foreach ($record in $headerPackages) {
        if (-not $legalFrozenInputIds.Contains(
                [string]$record['inputId'])) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' header package is absent from inline legal provenance.")
        }
    }

    $build = $Descriptor['build']
    $buildKind = [string]$build['kind']
    $genericKinds = [string[]]@(
        'provider-generic-cargo-msvc-v1',
        'provider-generic-cmake-msvc-v1',
        'provider-generic-zig-msvc-v1')
    if ($buildKind -cnotin ($genericKinds + @('candidate-wrapper-v1'))) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' build kind is invalid.")
    }
    if ($buildKind -ceq 'provider-generic-cmake-msvc-v1') {
        Assert-RSCollectionDescriptorOrdinalStrings `
            -Value ([string[]]$build['buildTargets']) `
            -Label "Candidate '$CandidateId' CMake build targets"
    }
    elseif ($buildKind -ceq 'provider-generic-zig-msvc-v1') {
        Assert-RSCollectionDescriptorOrdinalStrings `
            -Value ([string[]]$build['buildSteps']) `
            -Label "Candidate '$CandidateId' Zig build steps"
    }
    elseif ($buildKind -ceq 'provider-generic-cargo-msvc-v1') {
        Assert-RSCollectionDescriptorOrdinalStrings `
            -Value ([string[]]$build['features']) `
            -Label "Candidate '$CandidateId' Cargo features"
    }

    $bindings = [object[]]@($build['pathBindings'])
    Assert-RSCollectionDescriptorOrdinalRecords `
        -Record $bindings `
        -Property bindingKey `
        -Label "Candidate '$CandidateId' build path bindings"
    $headerIds =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($record in $headerPackages) {
        [void]$headerIds.Add([string]$record['inputId'])
    }
    foreach ($binding in $bindings) {
        $kind = [string]$binding['valueKind']
        if ($kind -ceq 'header-package-root' -and
            -not $headerIds.Contains([string]$binding['inputId'])) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' path binding references an undeclared header package.")
        }
    }

    $artifactLeaves =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    $adapter = $build['adapter']
    $adapterDll = [string]$adapter['dllLeaf']
    Assert-RSCollectionDescriptorSafeLeaf `
        -Value $adapterDll `
        -Label "Candidate '$CandidateId' adapter DLL"
    if (-not $adapterDll.EndsWith(
            '.dll',
            [StringComparison]::OrdinalIgnoreCase) -or
        -not $artifactLeaves.Add($adapterDll)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' adapter artifact leaf is invalid.")
    }

    $staticLibraries = if ($buildKind -ceq
            'provider-generic-cmake-msvc-v1') {
        [object[]]@($build['staticLibraries'])
    }
    else {
        [object[]]@()
    }
    if ($buildKind -ceq 'provider-generic-cmake-msvc-v1') {
        $projectLeaf = [string]$adapter['projectLeaf']
        Assert-RSCollectionDescriptorSafeLeaf `
            -Value $projectLeaf `
            -Label "Candidate '$CandidateId' adapter project"
        if (-not $projectLeaf.EndsWith(
                '.vcxproj',
                [StringComparison]::OrdinalIgnoreCase) -or
            -not $artifactLeaves.Add($projectLeaf)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' adapter project leaf is invalid or colliding.")
        }
        Assert-RSCollectionDescriptorOrdinalRecords `
            -Record $staticLibraries `
            -Property libraryId `
            -Label "Candidate '$CandidateId' static libraries"
    }
    foreach ($record in $staticLibraries) {
        if (-not $legalFrozenInputIds.Contains(
                [string]$record['sourceInputId'])) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' static library has no inline legal provenance.")
        }
        foreach ($artifact in @(
                @([string]$record['projectLeaf'], '.vcxproj'),
                @([string]$record['libraryLeaf'], '.lib'))) {
            Assert-RSCollectionDescriptorSafeLeaf `
                -Value $artifact[0] `
                -Label "Candidate '$CandidateId' static-library artifact"
            if (-not $artifact[0].EndsWith(
                    $artifact[1],
                    [StringComparison]::OrdinalIgnoreCase) -or
                -not $artifactLeaves.Add($artifact[0])) {
                throw [IO.InvalidDataException]::new(
                    "Candidate '$CandidateId' static-library artifact leaf is invalid or colliding.")
            }
        }
    }

    $privateLibraries =
        [object[]]@($build['privateRuntimeLibraries'])
    Assert-RSCollectionDescriptorOrdinalRecords `
        -Record $privateLibraries `
        -Property fileName `
        -Label "Candidate '$CandidateId' private runtime libraries"
    $privatePaths =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($record in $privateLibraries) {
        $fileName = [string]$record['fileName']
        $relative = [string]$record['relativeBuildPath']
        Assert-RSCollectionDescriptorSafeLeaf `
            -Value $fileName `
            -Label "Candidate '$CandidateId' private runtime leaf"
        Assert-RSCollectionDescriptorRelativePath `
            -Value $relative `
            -Label "Candidate '$CandidateId' private runtime path"
        if (-not $fileName.EndsWith(
                '.dll',
                [StringComparison]::OrdinalIgnoreCase) -or
            [IO.Path]::GetFileName($relative) -cne $fileName -or
            -not $privatePaths.Add($relative) -or
            -not $artifactLeaves.Add($fileName)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' private runtime path is invalid or colliding.")
        }
    }

    if ($buildKind -ceq 'candidate-wrapper-v1') {
        $driverPath = [string]$build['driverRelativePath']
        Assert-RSCollectionDescriptorRelativePath `
            -Value $driverPath `
            -Label "Candidate '$CandidateId' wrapper driver"
        if (-not $driverPath.EndsWith(
                '.ps1',
                [StringComparison]::Ordinal) -or
            [StringComparer]::OrdinalIgnoreCase.Equals(
                $driverPath,
                $script:RSCollectionDescriptorLeaf) -or
            [StringComparer]::OrdinalIgnoreCase.Equals(
                $driverPath,
                $outputPath)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' wrapper driver path collides with a reserved descriptor path.")
        }
    }
}

function Read-RSTerminalEngineCollectionDescriptor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId,

        [Parameter(Mandatory = $true)]
        [string]$SchemaFile,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    if (-not [IO.File]::Exists($LiteralPath)) {
        throw [IO.FileNotFoundException]::new(
            "Candidate '$CandidateId' collection descriptor is missing.",
            $LiteralPath)
    }
    $item = Get-Item -LiteralPath $LiteralPath -Force
    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
        $item.Length -le 0 -or
        [uint64]$item.Length -gt $script:RSCollectionDescriptorMaxBytes) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' collection descriptor violates the fixed byte or reparse bound.")
    }
    $bytes = [IO.File]::ReadAllBytes($LiteralPath)
    try {
        $json = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
        $descriptor = $json |
            ConvertFrom-Json `
                -AsHashtable `
                -Depth 64 `
                -NoEnumerate `
                -DateKind String
    }
    catch {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' collection descriptor is not strict UTF-8 JSON.",
            $_.Exception)
    }

    $canonicalBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $descriptor
    if (-not [Linq.Enumerable]::SequenceEqual(
            [byte[]]$bytes,
            [byte[]]$canonicalBytes)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' collection descriptor is not exact RFC 8785 JCS.")
    }
    if (-not [IO.File]::Exists($SchemaFile) -or
        -not (Test-Json `
            -Json $json `
            -SchemaFile $SchemaFile `
            -ErrorAction Stop)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' collection descriptor violates its closed schema.")
    }
    Assert-RSCollectionDescriptorSemantics `
        -Descriptor $descriptor `
        -CandidateId $CandidateId
    return [pscustomobject][ordered]@{
        Path = [IO.Path]::GetFullPath($LiteralPath)
        Sha256 = Get-RSCollectionDescriptorSha256 -LiteralPath $LiteralPath
        SizeBytes = $item.Length.ToString(
            [Globalization.CultureInfo]::InvariantCulture)
        BuildKind = [string]$descriptor['build']['kind']
        MaintainedBuildWrapper = (
            [string]$descriptor['build']['kind'] -ceq
                'candidate-wrapper-v1')
        Value = $descriptor
    }
}

function Assert-RSTerminalEngineCollectionDescriptorPathClosure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Descriptor,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$CandidateAuthoredPath,

        [Parameter(Mandatory = $true)]
        [string]$CandidateId
    )

    $root = [IO.Path]::GetFullPath($SourceRoot).
        TrimEnd([IO.Path]::DirectorySeparatorChar)
    $changed =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::Ordinal)
    $changedCaseFolded =
        [Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    foreach ($raw in $CandidateAuthoredPath) {
        $path = $raw.Replace('\', '/')
        Assert-RSCollectionDescriptorRelativePath `
            -Value $path `
            -Label "Candidate '$CandidateId' authored patch path" `
            -MaximumLength 1024 `
            -AllowSpace
        if (-not $changed.Add($path) -or
            -not $changedCaseFolded.Add($path)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' authored paths collide under Windows case folding.")
        }
    }

    $value = if ($Descriptor -is [Collections.IDictionary]) {
        $Descriptor
    }
    elseif ($Descriptor.PSObject.Properties.Name -contains 'Value') {
        $Descriptor.Value
    }
    else {
        $Descriptor
    }
    $requiredAuthored = [Collections.Generic.List[string]]::new()
    $requiredAuthored.Add($script:RSCollectionDescriptorLeaf)
    if ([string]$value['build']['kind'] -ceq 'candidate-wrapper-v1') {
        $requiredAuthored.Add(
            [string]$value['build']['driverRelativePath'])
    }
    $candidateSourceInputs = [Collections.Generic.List[object]]::new()
    foreach ($dependency in [object[]]@(
            $value['legal']['dependencies'])) {
        foreach ($record in [object[]]@($dependency['inputs'])) {
            $source = $record['source']
            if ([string]$source['sourceKind'] -cne
                'candidate-source') {
                continue
            }
            $candidateSourceInputs.Add($record)
            if ([string]$source['origin'] -ceq 'candidate-patch') {
                $requiredAuthored.Add(
                    [string]$source['sourceRelativePath'])
            }
        }
    }
    foreach ($path in $requiredAuthored) {
        if (-not $changed.Contains($path)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' candidate-authored descriptor path '$path' is absent from the exact patch.")
        }
        $absolute = [IO.Path]::GetFullPath((
                Join-Path $root ($path.Replace('/', '\'))))
        if (-not $absolute.StartsWith(
                $root + [IO.Path]::DirectorySeparatorChar,
                [StringComparison]::OrdinalIgnoreCase) -or
            -not [IO.File]::Exists($absolute)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' candidate-authored descriptor path '$path' is absent from the result tree.")
        }
        Assert-RSCollectionDescriptorContainedNoReparsePath `
            -LiteralPath $absolute `
            -Root $root `
            -Label "Candidate '$CandidateId' candidate-authored descriptor path '$path'"
    }

    foreach ($record in $candidateSourceInputs) {
        $source = $record['source']
        $path = [string]$source['sourceRelativePath']
        $origin = [string]$source['origin']
        if ($origin -ceq 'upstream' -and
            $changedCaseFolded.Contains($path)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' upstream legal source input '$path' is unexpectedly patch-authored.")
        }
        $absolute = [IO.Path]::GetFullPath((
                Join-Path $root ($path.Replace('/', '\'))))
        if (-not $absolute.StartsWith(
                $root + [IO.Path]::DirectorySeparatorChar,
                [StringComparison]::OrdinalIgnoreCase) -or
            -not [IO.File]::Exists($absolute)) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' legal source input '$path' is absent from the result tree.")
        }
        Assert-RSCollectionDescriptorContainedNoReparsePath `
            -LiteralPath $absolute `
            -Root $root `
            -Label "Candidate '$CandidateId' legal source input '$path'"
        $item = Get-Item -LiteralPath $absolute -Force
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
            $item.Length -ne [int64]$record['sizeBytes'] -or
            (Get-RSCollectionDescriptorSha256 -LiteralPath $absolute) -cne
                [string]$record['sha256']) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$CandidateId' legal source input '$path' identity is invalid.")
        }
    }

    $outputPath = [string]$value['legal']['outputRelativePath']
    $outputAbsolute = [IO.Path]::GetFullPath((
            Join-Path $root ($outputPath.Replace('/', '\'))))
    Assert-RSCollectionDescriptorContainedNoReparsePath `
        -LiteralPath $outputAbsolute `
        -Root $root `
        -Label "Candidate '$CandidateId' generated legal output"
    if (-not $outputAbsolute.StartsWith(
            $root + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase) -or
        $changedCaseFolded.Contains($outputPath) -or
        [IO.File]::Exists($outputAbsolute) -or
        [IO.Directory]::Exists($outputAbsolute)) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$CandidateId' generated legal output must be absent and untracked before materialization.")
    }
}

Export-ModuleMember -Function @(
    'Get-RSTerminalEngineCollectionDescriptorLeaf',
    'Read-RSTerminalEngineCollectionDescriptor',
    'Assert-RSTerminalEngineCollectionDescriptorPathClosure')
